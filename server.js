require('dotenv').config();

const express = require('express');
const session = require('express-session');
const multer = require('multer');
const { google } = require('googleapis');
const fs = require('fs');
const xlsx = require('xlsx');
const pdfParse = require('pdf-parse');
const archiver = require('archiver');
const mongoose = require('mongoose');

const app = express();
const upload = multer({ dest: 'uploads/' });

// ================= PORT =================
const PORT = process.env.PORT || 3000;

// ================= ENV =================
const FOLDER_ID = process.env.FOLDER_ID;
const EXCEL_FILE_ID = process.env.EXCEL_FILE_ID;

// ================= MONGODB =================
mongoose.connect(process.env.MONGO_URI)
  .then(() => console.log("✅ MongoDB Connected"))
  .catch(err => console.log("❌ DB Error:", err));

// ================= SCHEMA =================
const UploadSchema = new mongoose.Schema({
  code: String,
  type: String,
  files: Array,
  value: String,
  uploadedBy: String,
  uploadedAt: String
});

const Upload = mongoose.model('Upload', UploadSchema);

// ================= GOOGLE DRIVE =================
if (!process.env.GOOGLE_CREDENTIALS) {
  console.error("❌ GOOGLE_CREDENTIALS missing");
  process.exit(1);
}

const auth = new google.auth.GoogleAuth({
  credentials: JSON.parse(process.env.GOOGLE_CREDENTIALS),
  scopes: ['https://www.googleapis.com/auth/drive']
});

const drive = google.drive({ version: 'v3', auth });

// ================= MIDDLEWARE =================
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

app.use(session({
  secret: process.env.SESSION_SECRET || 'secret',
  resave: false,
  saveUninitialized: false,
  cookie: { secure: false }
}));

app.use(express.static('public'));

// ================= LOGIN =================
app.post('/adminLogin', (req, res) => {
  if (req.body.username === 'admin' && req.body.password === 'admin123') {
    req.session.user = 'admin';
    req.session.role = 'admin';
    return res.json({ status: 'success' });
  }
  res.json({ status: 'error' });
});

app.post('/userLogin', (req, res) => {
  if (!req.body.email) return res.json({ status: 'error' });

  req.session.user = req.body.email.toLowerCase();
  req.session.role = 'user';

  res.json({ status: 'success' });
});

// ================= DOWNLOAD EXCEL =================
async function downloadExcel() {
  try {
    const dest = fs.createWriteStream('temp.xlsx');

    const res = await drive.files.get(
      { fileId: EXCEL_FILE_ID, alt: 'media' },
      { responseType: 'stream' }
    );

    await new Promise((resolve, reject) => {
      res.data.pipe(dest).on('finish', resolve).on('error', reject);
    });

    const wb = xlsx.readFile('temp.xlsx');
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const data = xlsx.utils.sheet_to_json(sheet);

    fs.unlinkSync('temp.xlsx');
    return data;

  } catch (err) {
    console.error("❌ Excel Error:", err.message);
    return [];
  }
}

// ================= MERGE =================
function convertToMap(dbData) {
  let map = {};
  dbData.forEach(d => {
    const key = `${d.code}_${d.type}`;
    map[key] = d;
  });
  return map;
}

function mergeData(master, updates) {
  return master.map(row => {

    const code = row.Code;

    const sss = updates[`${code}_SSS`];
    const aws = updates[`${code}_AWS`];

    return {
      ...row,
      SSS_Status: sss ? "Done" : "Pending",
      AWS_Status: aws ? "Done" : "Pending",
      SSS_Files: sss?.files || [],
      AWS_Files: aws?.files || [],
      SSS_Value: sss?.value || '',
      AWS_Value: aws?.value || ''
    };
  });
}

// ================= FOLDER =================
const folderCache = {};

async function getOrCreateFolder(name, parentId) {

  const key = parentId + "_" + name;
  if (folderCache[key]) return folderCache[key];

  const res = await drive.files.list({
    q: `name='${name}' and mimeType='application/vnd.google-apps.folder' and '${parentId}' in parents and trashed=false`,
    fields: 'files(id)'
  });

  let id;

  if (res.data.files.length > 0) {
    id = res.data.files[0].id;
  } else {
    const folder = await drive.files.create({
      resource: {
        name,
        mimeType: 'application/vnd.google-apps.folder',
        parents: [parentId]
      },
      fields: 'id'
    });
    id = folder.data.id;
  }

  folderCache[key] = id;
  return id;
}

// ================= GET DATA =================
app.get('/getData', async (req, res) => {

  if (!req.session.user) return res.status(401).send("Login required");

  const master = await downloadExcel();
  const dbData = await Upload.find();
  const updates = convertToMap(dbData);

  let rows = mergeData(master, updates);

  if (req.session.role !== 'admin') {
    const email = req.session.user;

    rows = rows.filter(row =>
      [row.BH_Email, row.SM_Email, row.ZBM_Email, row.RBM_Email, row.ABM_Email]
        .some(e => e && e.toLowerCase().includes(email))
    );
  }

  res.json({ rows, uploads: updates, role: req.session.role });
});

// ================= FILE VALIDATION =================
function isValidFile(file) {
  const allowed = ['pdf','doc','docx','xls','xlsx','txt','html'];
  return allowed.includes(file.originalname.split('.').pop().toLowerCase());
}

async function isReadablePDF(path) {
  try {
    const data = await pdfParse(fs.readFileSync(path));
    return data.text.trim().length > 10;
  } catch {
    return false;
  }
}

// ================= UPLOAD =================
app.post('/uploadFile', upload.array('files', 10), async (req, res) => {

  if (!req.session.user) return res.status(401).send("Login required");

  const { code, stockistName, value, type } = req.body;
  const files = req.files;

  if (!value) return res.json({ status: 'error', msg: 'Value required' });

  const exists = await Upload.findOne({ code, type });
  if (exists) return res.json({ status: 'error', msg: `${type} already uploaded` });

  const master = await downloadExcel();
  const row = master.find(r => r.Code === code);
  const state = row?.State || "Others";

  const stateFolderId = await getOrCreateFolder(state, FOLDER_ID);
  const typeFolderId = await getOrCreateFolder(type, stateFolderId);

  let uploadedFiles = [];

  try {
    for (let file of files) {

      if (!isValidFile(file)) {
        fs.unlinkSync(file.path);
        return res.json({ status: 'error', msg: 'Invalid format' });
      }

      if (file.mimetype === 'application/pdf') {
        const ok = await isReadablePDF(file.path);
        if (!ok) {
          fs.unlinkSync(file.path);
          return res.json({ status: 'error', msg: 'Unreadable PDF' });
        }
      }

      const fileName = `${stockistName}_${code}_${type}_${Date.now()}`;

      const response = await drive.files.create({
        resource: { name: fileName, parents: [typeFolderId] },
        media: { mimeType: file.mimetype, body: fs.createReadStream(file.path) },
        fields: 'id'
      });

      await drive.permissions.create({
        fileId: response.data.id,
        requestBody: { role: 'reader', type: 'anyone' }
      });

      fs.unlinkSync(file.path);

      uploadedFiles.push({
        fileName,
        fileUrl: `https://drive.google.com/file/d/${response.data.id}/view`
      });
    }

    await Upload.create({
      code,
      type,
      files: uploadedFiles,
      value,
      uploadedBy: req.session.user,
      uploadedAt: new Date().toLocaleString()
    });

    res.json({ status: 'success' });

  } catch (err) {
    console.error("❌ Upload Error:", err);
    res.json({ status: 'error', msg: 'Upload failed' });
  }
});

// ================= DOWNLOAD REPORT =================
app.get('/downloadReport', async (req, res) => {

  if (req.session.role !== 'admin') return res.send("Admin only");

  const master = await downloadExcel();
  const dbData = await Upload.find();
  const updates = convertToMap(dbData);

  const merged = mergeData(master, updates);

  const wb = xlsx.utils.book_new();
  const ws = xlsx.utils.json_to_sheet(merged);

  xlsx.utils.book_append_sheet(wb, ws, "Report");
  xlsx.writeFile(wb, "report.xlsx");

  res.download("report.xlsx");
});

// ================= DOWNLOAD ZIP =================
app.get('/downloadZip', async (req, res) => {

  if (req.session.role !== 'admin') return res.send("Admin only");

  const dbData = await Upload.find();
  const master = await downloadExcel();

  res.attachment('files.zip');
  const archive = archiver('zip');

  archive.pipe(res);

  for (let d of dbData) {

    const row = master.find(r => r.Code === d.code);
    const state = row?.State || "Others";

    for (let file of d.files) {

      const fileId = file.fileUrl.split('/d/')[1].split('/')[0];

      const response = await drive.files.get(
        { fileId, alt: 'media' },
        { responseType: 'stream' }
      );

      archive.append(response.data, {
        name: `${state}/${d.type}/${file.fileName}`
      });
    }
  }

  await archive.finalize();
});

// ================= START =================
app.listen(PORT, () => console.log(`🚀 Server running on port ${PORT}`));