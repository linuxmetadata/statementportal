require('dotenv').config();

const express = require('express');
const session = require('express-session');
const multer = require('multer');
const { google } = require('googleapis');
const fs = require('fs');

const app = express();
const upload = multer({ dest: 'uploads/' });

const PORT = process.env.PORT || 3000;

const CLIENT_ID = process.env.GOOGLE_CLIENT_ID;
const CLIENT_SECRET = process.env.GOOGLE_CLIENT_SECRET;
const REDIRECT_URI = process.env.GOOGLE_REDIRECT_URI;

const FOLDER_ID = '168KzEusKlXsHQ-votUNTNA9g0VIai4X-';

// ================= SESSION =================
app.use(session({
  secret: 'portal-secret',
  resave: false,
  saveUninitialized: true
}));

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static('public'));

// ================= GOOGLE AUTH =================
const oauth2Client = new google.auth.OAuth2(
  CLIENT_ID,
  CLIENT_SECRET,
  REDIRECT_URI
);

// ================= STORAGE (IMPORTANT) =================
let uploadsData = {}; 
// format:
// uploadsData["AP183_SSS"] = {
//   fileName,
//   fileId,
//   uploadedBy,
//   uploadedAt
// }

// ================= ADMIN LOGIN =================
app.post('/login', (req, res) => {
  const { username, password } = req.body;

  if (username === 'admin' && password === 'admin') {
    req.session.user = username;
    return res.json({ status: 'success' });
  }

  res.json({ status: 'error', msg: 'Invalid login' });
});

// ================= GOOGLE LOGIN =================
app.get('/auth', (req, res) => {
  const url = oauth2Client.generateAuthUrl({
    access_type: 'offline',
    prompt: 'consent',
    scope: ['https://www.googleapis.com/auth/drive']
  });

  res.redirect(url);
});

// ================= CALLBACK =================
app.get('/oauth2callback', async (req, res) => {
  try {
    const code = req.query.code;

    const { tokens } = await oauth2Client.getToken(code);

    oauth2Client.setCredentials(tokens);
    req.session.tokens = tokens;

    res.redirect('/dashboard.html');

  } catch (err) {
    console.log("❌ OAuth Error:", err.message);
    res.send("Login failed: " + err.message);
  }
});

// ================= AUTH =================
function checkAuth(req, res, next) {
  if (!req.session.tokens) {
    return res.status(401).json({ status: 'error', msg: 'Not authenticated' });
  }

  oauth2Client.setCredentials(req.session.tokens);
  next();
}

// ================= GET DATA =================
app.get('/getData', (req, res) => {
  res.json({ status: 'success', uploads: uploadsData });
});

// ================= UPLOAD =================
app.post('/uploadFile', checkAuth, upload.single('file'), async (req, res) => {
  try {
    const file = req.file;
    const { code, type, stockistName } = req.body;

    if (!file) {
      return res.json({ status: 'error', msg: 'No file uploaded' });
    }

    const key = `${code}_${type}`;

    // prevent re-upload
    if (uploadsData[key]) {
      return res.json({ status: 'error', msg: 'Already uploaded' });
    }

    const cleanName = stockistName
      .replace(/[^a-zA-Z0-9 ]/g, '')
      .replace(/\s+/g, ' ')
      .trim();

    const finalName = `${cleanName}_${code}_${type}${file.originalname.substring(file.originalname.lastIndexOf('.'))}`;

    const drive = google.drive({ version: 'v3', auth: oauth2Client });

    const response = await drive.files.create({
      resource: {
        name: finalName,
        parents: [FOLDER_ID]
      },
      media: {
        mimeType: file.mimetype,
        body: fs.createReadStream(file.path)
      },
      fields: 'id'
    });

    fs.unlinkSync(file.path);

    // SAVE DATA
    uploadsData[key] = {
      fileName: finalName,
      fileId: response.data.id,
      uploadedBy: req.session.user || "Admin",
      uploadedAt: new Date().toLocaleString()
    };

    res.json({ status: 'success' });

  } catch (err) {
    console.log("❌ Upload Error:", err.message);
    res.json({ status: 'error', msg: err.message });
  }
});

// ================= DOWNLOAD =================
app.get('/download/:id', checkAuth, async (req, res) => {
  try {
    const drive = google.drive({ version: 'v3', auth: oauth2Client });

    const file = await drive.files.get({
      fileId: req.params.id,
      alt: 'media'
    }, { responseType: 'stream' });

    file.data.pipe(res);

  } catch (err) {
    res.send(err.message);
  }
});

// ================= LOGOUT =================
app.get('/logout', (req, res) => {
  req.session.destroy(() => {
    res.redirect('/login.html');
  });
});

// ================= DEFAULT =================
app.get('/', (req, res) => {
  res.redirect('/login.html');
});

// ================= START =================
app.listen(PORT, () => {
  console.log("🚀 Server running...");
});