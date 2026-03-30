require('dotenv').config();

const express = require('express');
const { google } = require('googleapis');
const multer = require('multer');
const xlsx = require('xlsx');
const path = require('path');
const { PassThrough } = require('stream');

const app = express();
const PORT = process.env.PORT || 10000;

// ================= MULTER =================
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

// Login page
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'login.html'));
});

// Dashboard page
app.get('/dashboard', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'dashboard.html'));
});

// ================= GET GOOGLE SHEET DATA =================
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

// ================= CREATE FOLDER (IF NOT EXISTS) =================
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

// ================= UPLOAD FILES =================
app.post('/upload', upload.array('files'), async (req, res) => {
  try {
    const { code, type, state, name } = req.body;

    if (!req.files || req.files.length === 0) {
      return res.json({ message: "No files selected ❌" });
    }

    if (!code || !type || !state || !name) {
      return res.json({ message: "Missing required data ❌" });
    }

    const rootFolder = process.env.FOLDER_ID;

    // 🔹 Create State Folder
    const stateFolder = await createFolder(state, rootFolder);

    // 🔹 Create Type Folder (SSS / AWS)
    const typeFolder = await createFolder(type, stateFolder);

    for (let file of req.files) {

      const bufferStream = new PassThrough();
      bufferStream.end(file.buffer);

      const fileName = `${name}_${code}_${type}_${file.originalname}`;

      await drive.files.create({
        requestBody: {
          name: fileName,
          parents: [typeFolder]
        },
        media: {
          mimeType: file.mimetype,
          body: bufferStream
        }
      });
    }

    res.json({ message: "Upload Success ✅" });

  } catch (err) {
    console.error("❌ UPLOAD ERROR FULL:", err);
    res.json({ message: "Upload Failed ❌" });
  }
});

// ================= START SERVER =================
app.listen(PORT, () => {
  console.log(`🚀 Server running on port ${PORT}`);
});