const express = require('express');
const session = require('express-session');
const multer = require('multer');
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const { google } = require('googleapis');

const app = express();
const PORT = process.env.PORT || 3000;

// ================= BASIC =================
app.use(express.static('public'));
app.use(express.json());

app.use(session({
  secret: 'statementportal_secret',
  resave: false,
  saveUninitialized: true
}));

// ================= FILE UPLOAD =================
const upload = multer({ dest: 'uploads/' });

// ================= OAUTH =================
const oauth2Client = new google.auth.OAuth2(
  process.env.GOOGLE_CLIENT_ID,
  process.env.GOOGLE_CLIENT_SECRET,
  process.env.GOOGLE_REDIRECT_URI
);

const SCOPES = ['https://www.googleapis.com/auth/drive'];

// ================= LOGIN =================
app.get('/auth', (req, res) => {
  const url = oauth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES,
    prompt: 'consent'
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
    console.log("❌ OAuth Error:", err);
    res.send("Login failed: " + err.message);
  }
});

// ================= DRIVE =================
function getDrive(req) {
  oauth2Client.setCredentials(req.session.tokens);
  return google.drive({ version: 'v3', auth: oauth2Client });
}

// ================= DATA STORAGE =================
let DATA = [];

// ================= LOAD SOURCE EXCEL =================
async function loadExcelFromDrive(req) {
  try {

    const drive = getDrive(req);

    const list = await drive.files.list({
      pageSize: 1,
      fields: 'files(id, name)'
    });

    if (!list.data.files.length) {
      console.log("❌ No Excel found in Drive");
      return;
    }

    const file = list.data.files[0];

    const tempPath = path.join(__dirname, 'temp.xlsx');
    const dest = fs.createWriteStream(tempPath);

    const response = await drive.files.get(
      { fileId: file.id, alt: 'media' },
      { responseType: 'stream' }
    );

    await new Promise((resolve, reject) => {
      response.data
        .pipe(dest)
        .on('finish', resolve)
        .on('error', reject);
    });

    const workbook = XLSX.readFile(tempPath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const raw = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    DATA = raw.map(r => ({
      STATE: r["STATE"] || "",
      BM_HQ: r["BM_HQ"] || "",
      Code: r["Code"] || "",
      Name: r["Stockist Name"] || "",
      Value: "",
      SSS_File: "",
      AWS_File: "",
      SSS_Status: "Pending",
      AWS_Status: "Pending",
      Submitted_By: "",
      Submitted_On: ""
    }));

    fs.unlinkSync(tempPath);

    console.log("✅ Excel Loaded:", DATA.length);

  } catch (err) {
    console.log("❌ Excel Load Error:", err.message);
  }
}

// ================= GET DATA =================
app.get('/getData', async (req, res) => {
  try {
    if (!req.session.tokens) {
      return res.json({ data: [] });
    }

    await loadExcelFromDrive(req);

    res.json({ data: DATA });

  } catch (err) {
    res.json({ data: [] });
  }
});

// ================= UPLOAD FILE =================
app.post('/uploadFile', upload.single('file'), async (req, res) => {
  try {

    if (!req.session.tokens) {
      return res.json({ status: "error", msg: "Login required" });
    }

    const drive = getDrive(req);

    const file = req.file;
    const code = req.body.code;
    const type = req.body.type; // SSS or AWS

    if (!file) {
      return res.json({ status: "error", msg: "No file selected" });
    }

    // Upload to Google Drive
    const uploadRes = await drive.files.create({
      requestBody: {
        name: `${type}_${code}_${file.originalname}`
      },
      media: {
        body: fs.createReadStream(file.path)
      }
    });

    const fileId = uploadRes.data.id;

    // Make public
    await drive.permissions.create({
      fileId,
      requestBody: {
        role: 'reader',
        type: 'anyone'
      }
    });

    const fileLink = `https://drive.google.com/file/d/${fileId}/view`;

    // Update DATA
    const row = DATA.find(r => r.Code === code);

    if (row) {
      if (type === "SSS") {
        row.SSS_File = fileLink;
        row.SSS_Status = "Uploaded";
      } else {
        row.AWS_File = fileLink;
        row.AWS_Status = "Uploaded";
      }

      row.Submitted_By = "Admin";
      row.Submitted_On = new Date().toLocaleString();
    }

    fs.unlinkSync(file.path);

    console.log("✅ Upload Success:", file.originalname);

    res.json({ status: "success" });

  } catch (err) {
    console.log("❌ Upload Error:", err.message);

    res.json({
      status: "error",
      msg: err.message || "Upload failed"
    });
  }
});

// ================= DOWNLOAD REPORT =================
app.get('/downloadReport', (req, res) => {
  try {
    const ws = XLSX.utils.json_to_sheet(DATA);
    const wb = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(wb, ws, "Report");

    const filePath = "report.xlsx";

    XLSX.writeFile(wb, filePath);

    res.download(filePath);

  } catch (err) {
    res.send("Error generating report");
  }
});

// ================= ROOT FIX =================
app.get('/', (req, res) => {
  res.redirect('/login.html');
});

// ================= START =================
app.listen(PORT, () => {
  console.log("🚀 OAuth Server Running...");
});