require('dotenv').config();

const express = require('express');
const session = require('express-session');
const multer = require('multer');
const { google } = require('googleapis');
const fs = require('fs');
const xlsx = require('xlsx');
const { Readable } = require('stream');

const app = express();
const upload = multer({ dest: 'uploads/' });

const PORT = process.env.PORT || 3000;

// ================= ENV =================
const FOLDER_ID = process.env.FOLDER_ID;
const EXCEL_FILE_ID = process.env.EXCEL_FILE_ID;

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
  secret: 'secret',
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
    console.error("❌ Excel Error FULL:", err);
    return [];
  }
}

// ================= GET DATA (DEBUG MODE) =================
app.get('/getData', async (req, res) => {

  console.log("🔥 API HIT /getData");

  const master = await downloadExcel();

  console.log("📊 MASTER DATA:", master);

  // RETURN DIRECTLY (NO MERGE)
  res.json({ rows: master });

});

// ================= START =================
app.listen(PORT, () => console.log(`🚀 Server running on port ${PORT}`));