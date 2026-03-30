require('dotenv').config();

const express = require('express');
const { google } = require('googleapis');
const multer = require('multer');
const xlsx = require('xlsx');
const path = require('path');
const { Readable } = require('stream');

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

// ================= GET DATA FROM GOOGLE SHEET =================
async function getSheetData() {
  try {
    const response = await drive.files.export(
      {
        fileId: process.env.EXCEL_FILE_ID,
        mimeType:
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      },
      { responseType: 'arraybuffer' }
    );

    const workbook = xlsx.read(response.data, { type: 'buffer' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = xlsx.utils.sheet_to_json(sheet);

    return data;
  } catch (err) {
    console.error('Excel Read Error:', err.message);
    return [];
  }
}

app.get('/getData', async (req, res) => {
  const data = await getSheetData();
  res.json({ rows: data });
});

// ================= UPLOAD FILES =================
app.post('/upload', upload.array('files'), async (req, res) => {
  try {
    const { code, type, state, name } = req.body;

    if (!req.files || req.files.length === 0) {
      return res.json({ message: 'No files selected ❌' });
    }

    if (!code || !type) {
      return res.json({ message: 'Missing data ❌' });
    }

    const parentFolder = process.env.FOLDER_ID;

    for (let file of req.files) {
      const fileName = `${name}_${code}_${type}_${file.originalname}`;

      await drive.files.create({
        requestBody: {
          name: fileName,
          parents: [parentFolder]
        },
        media: {
          mimeType: file.mimetype,
          body: Readable.from(file.buffer) // ✅ FIXED
        }
      });
    }

    res.json({ message: 'Upload Success ✅' });
  } catch (err) {
    console.error('UPLOAD ERROR:', err);
    res.json({ message: 'Upload Failed ❌' });
  }
});

// ================= START SERVER =================
app.listen(PORT, () => {
  console.log(`🚀 Server running on port ${PORT}`);
});