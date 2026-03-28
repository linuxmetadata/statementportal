require('dotenv').config();

const express = require('express');
const session = require('express-session');
const multer = require('multer');
const { google } = require('googleapis');
const fs = require('fs');
const xlsx = require('xlsx');
const pdfParse = require('pdf-parse');
const archiver = require('archiver');

const app = express();
const upload = multer({ dest: 'uploads/' });

const PORT = process.env.PORT || 3000;

// ================= ENV =================
const FOLDER_ID = process.env.FOLDER_ID;
const EXCEL_FILE_ID = process.env.EXCEL_FILE_ID;
const JSON_FILE_NAME = "updates.json";

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

app.get('/', (req, res) => {
  res.sendFile(__dirname + '/public/login.html');
});

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

// ================= EXCEL =================
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

// ================= JSON FILE =================
async function getUpdatesFile() {

  const res = await drive.files.list({
    q: `name='${JSON_FILE_NAME}' and '${FOLDER_ID}' in parents and trashed=false`,
    fields: 'files(id)'
  });

  if (res.data.files.length === 0) {

    const file = await drive.files.create({
      resource: {
        name: JSON_FILE_NAME,
        parents: [FOLDER_ID]
      },
      media: {
        mimeType: 'application/json',
        body: Buffer.from(JSON.stringify({}))
      },
      fields: 'id'
    });

    return file.data.id;
  }

  return res.data.files[0].id;
}

async function readUpdates() {

  const fileId = await getUpdatesFile();

  const res = await drive.files.get(
    { fileId, alt: 'media' },
    { responseType: 'stream' }
  );

  let data = '';

  return new Promise((resolve, reject) => {
    res.data.on('data', chunk => data += chunk);
    res.data.on('end', () => resolve(JSON.parse(data || "{}")));
    res.data.on('error', reject);
  });
}

async function writeUpdates(content) {

  const fileId = await getUpdatesFile();

  await drive.files.update({
    fileId,
    media: {
      mimeType: 'application/json',
      body: Buffer.from(JSON.stringify(content))
    }
  });
}

// ================= MERGE =================
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

  const updates = await readUpdates();
  const master = await downloadExcel();

  let rows = mergeData(master, updates);

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

  let updates = await readUpdates();
  const key = `${code}_${type}`;

  if (updates[key]) {
    return res.json({ status: 'error', msg: `${type} already uploaded` });
  }

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

    updates[key] = {
      files: uploadedFiles,
      value,
      uploadedBy: req.session.user,
      uploadedAt: new Date().toLocaleString()
    };

    await writeUpdates(updates);

    res.json({ status: 'success' });

  } catch (err) {
    console.error("❌ Upload Error:", err);
    res.json({ status: 'error', msg: 'Upload failed' });
  }
});

// ================= DOWNLOAD REPORT =================
app.get('/downloadReport', async (req, res) => {

  if (req.session.role !== 'admin') return res.send("Admin only");

  const updates = await readUpdates();
  const master = await downloadExcel();

  const merged = mergeData(master, updates);

  const wb = xlsx.utils.book_new();
  const ws = xlsx.utils.json_to_sheet(merged);

  xlsx.utils.book_append_sheet(wb, ws, "Report");
  xlsx.writeFile(wb, "report.xlsx");

  res.download("report.xlsx");
});

// ================= START =================
app.listen(PORT, () => console.log(`🚀 Server running on port ${PORT}`));