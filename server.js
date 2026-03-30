require('dotenv').config();

const express = require('express');
const { google } = require('googleapis');
const xlsx = require('xlsx');
const multer = require('multer');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

const upload = multer({ storage: multer.memoryStorage() });

// ================= GOOGLE AUTH =================
const auth = new google.auth.GoogleAuth({
  credentials: JSON.parse(process.env.GOOGLE_CREDENTIALS),
  scopes: ['https://www.googleapis.com/auth/drive']
});

const drive = google.drive({ version: 'v3', auth });

// ================= STATIC =================
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.json());

// ================= ROUTES =================
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'login.html'));
});

app.get('/dashboard', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'dashboard.html'));
});

// ================= EXCEL =================
async function downloadExcel() {
  const res = await drive.files.export(
    {
      fileId: process.env.EXCEL_FILE_ID,
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    },
    { responseType: 'arraybuffer' }
  );

  const wb = xlsx.read(res.data, { type: 'buffer' });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  return xlsx.utils.sheet_to_json(sheet);
}

app.get('/getData', async (req, res) => {
  const data = await downloadExcel();
  res.json({ rows: data });
});

// ================= CREATE FOLDER =================
async function createFolder(name, parentId) {
  const file = await drive.files.create({
    requestBody: {
      name,
      mimeType: 'application/vnd.google-apps.folder',
      parents: [parentId]
    }
  });
  return file.data.id;
}

// ================= UPLOAD =================
app.post('/upload', upload.array('files'), async (req, res) => {
  try {
    const { code, type, state, name } = req.body;

    const baseFolder = process.env.FOLDER_ID;

    // Create State folder
    const stateFolder = await createFolder(state, baseFolder);

    // Create SSS/AWS folder
    const typeFolder = await createFolder(type, stateFolder);

    for (let file of req.files) {

      const fileName = `${name}_${code}_${type}_${file.originalname}`;

      await drive.files.create({
        requestBody: {
          name: fileName,
          parents: [typeFolder]
        },
        media: {
          mimeType: file.mimetype,
          body: Buffer.from(file.buffer)
        }
      });
    }

    res.json({ message: "Uploaded Successfully ✅" });

  } catch (err) {
    console.error(err);
    res.json({ message: "Upload Failed ❌" });
  }
});

// ================= START =================
app.listen(PORT, () => console.log(`🚀 Running on ${PORT}`));