const express = require('express');
const session = require('express-session');
const multer = require('multer');
const { google } = require('googleapis');
const fs = require('fs');
const xlsx = require('xlsx');

const app = express();
const upload = multer({ dest: 'uploads/' });

const PORT = process.env.PORT || 3000;

// ================= ENV =================
const CLIENT_ID = process.env.GOOGLE_CLIENT_ID;
const CLIENT_SECRET = process.env.GOOGLE_CLIENT_SECRET;
const REDIRECT_URI = process.env.GOOGLE_REDIRECT_URI;

// 👉 Your Drive Folder
const FOLDER_ID = '168KzEusKlXsHQ-votUNTNA9g0VIai4X-';

// ================= MIDDLEWARE =================
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

app.use(session({
  secret: 'statement-secret',
  resave: false,
  saveUninitialized: true
}));

app.use(express.static('public'));

// ================= GOOGLE AUTH =================
const oauth2Client = new google.auth.OAuth2(
  CLIENT_ID,
  CLIENT_SECRET,
  REDIRECT_URI
);

// ================= STORAGE =================
let excelData = [];
let uploadsData = {};

// ================= ROOT FIX =================
app.get('/', (req, res) => {
  if (!req.session.user) {
    res.redirect('/login.html');
  } else {
    res.redirect('/dashboard.html');
  }
});

// ================= ADMIN LOGIN =================
app.post('/login', (req, res) => {
  const { username, password } = req.body;

  if (username === 'admin' && password === 'admin123') {
    req.session.user = 'Admin';
    req.session.role = 'admin';
    return res.json({ status: 'success' });
  }

  res.json({ status: 'error', msg: 'Invalid Login' });
});

// ================= GOOGLE LOGIN =================
app.get('/auth', (req, res) => {
  const url = oauth2Client.generateAuthUrl({
    access_type: 'offline',
    prompt: 'consent',
    scope: [
      'https://www.googleapis.com/auth/userinfo.email',
      'https://www.googleapis.com/auth/drive'
    ]
  });

  res.redirect(url);
});

// ================= CALLBACK =================
app.get('/oauth2callback', async (req, res) => {
  try {
    const code = req.query.code;

    const { tokens } = await oauth2Client.getToken(code);
    oauth2Client.setCredentials(tokens);

    const oauth2 = google.oauth2({ auth: oauth2Client, version: 'v2' });
    const user = await oauth2.userinfo.get();

    req.session.tokens = tokens;
    req.session.user = user.data.email.toLowerCase();
    req.session.role = 'user';

    res.redirect('/dashboard.html');

  } catch (err) {
    res.send("Login failed: " + err.message);
  }
});

// ================= AUTH CHECK =================
function checkAuth(req, res, next) {
  if (!req.session.user) {
    return res.status(401).json({ msg: 'Not logged in' });
  }

  if (req.session.tokens) {
    oauth2Client.setCredentials(req.session.tokens);
  }

  next();
}

// ================= LOAD EXCEL =================
app.post('/uploadExcel', upload.single('file'), (req, res) => {
  try {
    const wb = xlsx.readFile(req.file.path);
    const sheet = wb.Sheets[wb.SheetNames[0]];
    excelData = xlsx.utils.sheet_to_json(sheet);

    fs.unlinkSync(req.file.path);

    res.json({ status: 'success' });

  } catch (err) {
    res.json({ status: 'error', msg: err.message });
  }
});

// ================= GET DATA =================
app.get('/getData', checkAuth, (req, res) => {

  // ADMIN → ALL DATA
  if (req.session.role === 'admin') {
    return res.json({ rows: excelData, uploads: uploadsData });
  }

  const email = req.session.user;

  const filtered = excelData.filter(row => {
    return [
      row.BH_Email,
      row.SM_Email,
      row.ZBM_Email,
      row.RBM_Email,
      row.ABM_Email
    ].some(e => e && e.toLowerCase().includes(email));
  });

  res.json({ rows: filtered, uploads: uploadsData });
});

// ================= UPLOAD FILE =================
app.post('/uploadFile', checkAuth, upload.single('file'), async (req, res) => {
  try {
    const file = req.file;
    const { code, type, stockistName } = req.body;

    const key = `${code}_${type}`;

    // 🔒 Prevent duplicate upload
    if (uploadsData[key]) {
      return res.json({ status: 'error', msg: 'Already uploaded' });
    }

    // 🔥 FILE NAME FORMAT
    const cleanName = stockistName.replace(/[^a-zA-Z0-9 ]/g, '').trim();

    const fileName = `${cleanName}_${code}_${type}${file.originalname.substring(file.originalname.lastIndexOf('.'))}`;

    const drive = google.drive({ version: 'v3', auth: oauth2Client });

    const response = await drive.files.create({
      resource: {
        name: fileName,
        parents: [FOLDER_ID]
      },
      media: {
        mimeType: file.mimetype,
        body: fs.createReadStream(file.path)
      },
      fields: 'id'
    });

    fs.unlinkSync(file.path);

    // 📊 STORE META
    uploadsData[key] = {
      fileName,
      fileId: response.data.id,
      uploadedBy: req.session.user,
      uploadedAt: new Date().toLocaleString()
    };

    res.json({ status: 'success' });

  } catch (err) {
    res.json({ status: 'error', msg: err.message });
  }
});

// ================= LOGOUT =================
app.get('/logout', (req, res) => {
  req.session.destroy(() => {
    res.redirect('/login.html');
  });
});

// ================= START =================
app.listen(PORT, () => {
  console.log("🚀 Server running...");
});