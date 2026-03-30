require('dotenv').config();

const express = require('express');
const session = require('express-session');
const multer = require('multer');
const { google } = require('googleapis');
const fs = require('fs');
const xlsx = require('xlsx');
const pdfParse = require('pdf-parse');
const archiver = require('archiver');
const { Readable } = require('stream');

const app = express();
const upload = multer({ dest: 'uploads/' });

const PORT = process.env.PORT || 3000;

// ================= ENV =================
const FOLDER_ID = process.env.FOLDER_ID;
const EXCEL_FILE_ID = process.env.EXCEL_FILE_ID;
const JSON_FILE_NAME = "updates.json";

// ================= GOOGLE DRIVE =================
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
  saveUninitialized: false
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
  req.session.user = req.body.email.toLowerCase();
  req.session.role = 'user';
  res.json({ status: 'success' });
});

// ================= EXCEL =================
async function downloadExcel() {
  try {
    console.log("📥 Downloading Excel...");

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

    console.log("✅ Excel rows:", data.length);

    fs.unlinkSync('temp.xlsx');

    return data;

  } catch (err) {
    console.error("❌ Excel Error:", err);
    return [];
  }
}

// ================= JSON =================
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
        body: Readable.from([JSON.stringify({})])
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
      body: Readable.from([JSON.stringify(content)])
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
async function getOrCreateFolder(name, parentId) {

  const res = await drive.files.list({
    q: `name='${name}' and mimeType='application/vnd.google-apps.folder' and '${parentId}' in parents and trashed=false`,
    fields: 'files(id)'
  });

  if (res.data.files.length > 0) {
    return res.data.files[0].id;
  }

  const folder = await drive.files.create({
    resource: {
      name,
      mimeType: 'application/vnd.google-apps.folder',
      parents: [parentId]
    },
    fields: 'id'
  });

  return folder.data.id;
}

// ================= GET DATA =================
app.get('/getData', async (req, res) => {

  if (!req.session.user) return res.status(401).send("Login required");

  const updates = await readUpdates();
  const master = await downloadExcel();

  console.log("MASTER:", master.length);

  const rows = mergeData(master, updates);

  res.json({ rows, uploads: updates, role: req.session.role });
});

// ================= VALIDATION =================
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

  const stateFolder = await getOrCreateFolder(state, FOLDER_ID);
  const typeFolder = await getOrCreateFolder(type, stateFolder);

  let uploadedFiles = [];

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
      resource: { name: fileName, parents: [typeFolder] },
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
});

// ================= START =================
app.listen(PORT, () => console.log(`🚀 Server running on port ${PORT}`));