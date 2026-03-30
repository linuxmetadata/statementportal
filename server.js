require('dotenv').config();

const express = require('express');
const { google } = require('googleapis');
const multer = require('multer');
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
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// ================= SIMPLE UPLOAD =================
app.post('/upload', upload.single('file'), async (req, res) => {
  try {
    console.log("FILE:", req.file);

    if (!req.file) {
      return res.json({ message: "No file ❌" });
    }

    const stream = new PassThrough();
    stream.end(req.file.buffer);

    const response = await drive.files.create({
      requestBody: {
        name: req.file.originalname,
        parents: [process.env.FOLDER_ID]
      },
      media: {
        mimeType: req.file.mimetype,
        body: stream
      }
    });

    console.log("UPLOAD SUCCESS:", response.data.id);

    res.json({ message: "Upload Success ✅" });

  } catch (err) {
    console.error("UPLOAD ERROR:", err);
    res.json({ message: "Upload Failed ❌" });
  }
});

// ================= START =================
app.listen(PORT, () => {
  console.log(`🚀 Server running on port ${PORT}`);
});