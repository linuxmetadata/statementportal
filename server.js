require('dotenv').config();

const express = require('express');
const { google } = require('googleapis');
const multer = require('multer');
const xlsx = require('xlsx');
const path = require('path');
const { PassThrough } = require('stream');

const app = express();
const PORT = process.env.PORT || 10000;

const upload = multer({ storage: multer.memoryStorage() });

// ================= GOOGLE AUTH =================
const auth = new google.auth.GoogleAuth({
  credentials: JSON.parse(process.env.GOOGLE_CREDENTIALS),
  scopes: ['https://www.googleapis.com/auth/drive']
});

const drive = google.drive({ version: 'v3', auth });

// ================= MIDDLEWARE =================
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.json());

// ================= ROUTES =================
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'login.html'));
});

app.get('/dashboard', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'dashboard.html'));
});

// ================= GET DATA =================
async function getSheetData() {
  try {
    const response = await drive.files.export(
      {
        fileId: process.env.EXCEL_FILE_ID,
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      },
      { responseType: 'arraybuffer' }
    );

    const workbook = xlsx.read(response.data, { type: 'buffer' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    return xlsx.utils.sheet_to_json(sheet);

  } catch (err) {
    console.error('❌ Excel Read Error:', err.message);
    return [];
  }
}

app.get('/getData', async (req, res) => {
  const data = await getSheetData();
  res.json({ rows: data });
});

// ================= CREATE FOLDER =================
async function createFolder(name, parentId) {
  const res = await drive.files.create({
    requestBody: {
      name: name,
      mimeType: 'application/vnd.google-apps.folder',
      parents: [parentId]
    }
  });
  return res.data.id;
}

// ================= UPLOAD =================
app.post('/upload', upload.array('files'), async (req, res) => {
  try {
    console.log("BODY:", req.body);

    const { code, type, state, name, value } = req.body;

    if (!req.files || req.files.length === 0) {
      return res.json({ message: "No file selected ❌" });
    }

    if (!code || !type || !state || !name || !value) {
      return res.json({ message: "Missing fields ❌" });
    }

    const rootFolder = process.env.FOLDER_ID;

    const stateFolder = await createFolder(state, rootFolder);
    const typeFolder = await createFolder(type, stateFolder);

    for (let file of req.files) {
      const stream = new PassThrough();
      stream.end(file.buffer);

      const fileName = `${name}_${code}_${type}_${value}_${file.originalname}`;

      await drive.files.create({
        requestBody: {
          name: fileName,
          parents: [typeFolder]
        },
        media: {
          mimeType: file.mimetype,
          body: stream
        }
      });
    }

    res.json({ message: "Upload Success ✅" });

  } catch (err) {
    console.error("❌ UPLOAD ERROR:", err);
    res.json({ message: "Upload Failed ❌" });
  }
});

// ================= START =================
app.listen(PORT, () => {
  console.log(`🚀 Running on ${PORT}`);
});