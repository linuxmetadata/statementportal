const express = require('express');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const session = require('express-session');
const multer = require('multer');
const pdfParse = require('pdf-parse');
const axios = require('axios');
const archiver = require('archiver');
const { google } = require('googleapis');

const app = express();
const PORT = process.env.PORT || 3000;

// ================= BASIC =================
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.json());

app.use(session({
  secret: 'secret123',
  resave: false,
  saveUninitialized: true
}));

const upload = multer({ dest: 'uploads/' });

// ================= GOOGLE DRIVE =================
let drive;

try {
  const creds = JSON.parse(process.env.GOOGLE_CREDENTIALS);
  creds.private_key = creds.private_key.replace(/\\n/g, '\n');

  const auth = new google.auth.GoogleAuth({
    credentials: creds,
    scopes: ['https://www.googleapis.com/auth/drive']
  });

  drive = google.drive({ version: 'v3', auth });

} catch (err) {
  console.log("Google Auth Error:", err.message);
}

const FOLDER_ID = '168KzEusKlXsHQ-votUNTNA9g0VIai4X-';

// ================= DATA =================
let DATA = [];

// ================= HOME =================
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'login.html'));
});

// ================= AUTH =================
function checkAuth(req, res) {
  if (!req.session.user) {
    res.send("Login required");
    return false;
  }
  return true;
}

// ================= LOAD EXCEL FROM DRIVE =================
async function loadExcel() {
  try {

    const list = await drive.files.list({
      q: `'${FOLDER_ID}' in parents`,
      orderBy: 'createdTime desc',
      pageSize: 1,
      fields: 'files(id, name)',

      supportsAllDrives: true,
      includeItemsFromAllDrives: true
    });

    if (list.data.files.length === 0) {
      console.log("❌ No Excel found in Drive");
      return;
    }

    const file = list.data.files[0];
    const fileId = file.id;

    console.log("✅ Using file:", file.name);

    const dest = fs.createWriteStream("temp.xlsx");

    const res = await drive.files.get(
      {
        fileId,
        alt: 'media',
        supportsAllDrives: true
      },
      { responseType: 'stream' }
    );

    await new Promise((resolve, reject) => {
      res.data.pipe(dest).on('finish', resolve).on('error', reject);
    });

    const wb = XLSX.readFile("temp.xlsx");
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const raw = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    console.log("🔥 TOTAL ROWS:", raw.length);

    DATA = raw.map(row => {

      let r = {};
      Object.keys(row).forEach(k => {
        r[k.trim()] = row[k];
      });

      return {
        STATE: r["STATE"] || "",
        BM_HQ: r["BM_HQ"] || "",
        Code: r["Code"] || "",
        Name: r["Stockist Name"] || "",

        BH_Email: r["BH_Email"] || "",
        SM_Email: r["SM_Email"] || "",
        ZBM_Email: r["ZBM_Email"] || "",
        RBM_Email: r["RBM_Email"] || "",
        ABM_Email: r["ABM_Email"] || "",

        Value: "",
        SSS: false,
        AWS: false,

        SSS_File: "",
        AWS_File: "",

        SSS_Submitted_By: "",
        AWS_Submitted_By: "",
        SSS_Date: "",
        AWS_Date: ""
      };

    }).filter(r => r.Code && r.Name);

    fs.unlinkSync("temp.xlsx");

    console.log("✅ DATA LOADED:", DATA.length);

  } catch (err) {
    console.log("❌ Excel load error:", err.message);
  }
}

// ================= LOGIN =================
app.post('/login', (req, res) => {
  const { type, email, username, password } = req.body;

  if (type === "admin") {
    if (username === "admin" && password === "admin123") {
      req.session.user = { role: "admin" };
      return res.send("success");
    }
    return res.send("fail");
  }

  if (type === "user") {
    req.session.user = { role: "user", email };
    return res.send("success");
  }
});

// ================= GET DATA =================
app.get('/getData', (req, res) => {

  if (!checkAuth(req, res)) return;

  let result = DATA;

  if (req.session.user.role === "user") {
    const email = req.session.user.email;

    result = DATA.filter(r =>
      r.BH_Email === email ||
      r.SM_Email === email ||
      r.ZBM_Email === email ||
      r.RBM_Email === email ||
      r.ABM_Email === email
    );
  }

  res.json({ role: req.session.user.role, data: result });
});

// ================= REFRESH DATA =================
app.get('/refreshData', async (req, res) => {
  await loadExcel();
  res.send("Refreshed");
});

// ================= SAVE VALUE =================
app.post('/saveValue', (req, res) => {
  const { code, value } = req.body;

  DATA.forEach(r => {
    if (r.Code == code) r.Value = value;
  });

  res.send("saved");
});

// ================= UPLOAD FILE =================
app.post('/uploadFile', upload.single('file'), async (req, res) => {

  const { code, type } = req.body;
  const file = req.file;

  const ext = path.extname(file.originalname).toLowerCase();
  const allowed = ['.pdf', '.doc', '.docx', '.xls', '.xlsx', '.txt', '.html'];

  if (!allowed.includes(ext)) {
    fs.unlinkSync(file.path);
    return res.send("Invalid format");
  }

  if (ext === '.pdf') {
    const buffer = fs.readFileSync(file.path);
    const parsed = await pdfParse(buffer);

    if (!parsed.text || parsed.text.trim().length < 10) {
      fs.unlinkSync(file.path);
      return res.send("Invalid PDF");
    }
  }

  let row = DATA.find(r => r.Code == code);

  if (!row || !row.Value) {
    fs.unlinkSync(file.path);
    return res.send("Enter Value first");
  }

  const safe = row.Name.replace(/[^a-zA-Z0-9]/g, "_");
  const name = `${safe}_${code}_${type}${ext}`;

  const uploadRes = await drive.files.create({
    requestBody: {
      name,
      parents: [FOLDER_ID]
    },
    media: {
      body: fs.createReadStream(file.path)
    },
    supportsAllDrives: true
  });

  const fileId = uploadRes.data.id;

  await drive.permissions.create({
    fileId,
    requestBody: { role: 'reader', type: 'anyone' },
    supportsAllDrives: true
  });

  const link = `https://drive.google.com/uc?id=${fileId}`;

  fs.unlinkSync(file.path);

  row[type] = true;
  row[`${type}_File`] = link;
  row[`${type}_Submitted_By`] = req.session.user.email || "admin";
  row[`${type}_Date`] = new Date().toLocaleString();

  res.send("uploaded");
});

// ================= VIEW =================
app.get('/viewFile', (req, res) => {
  const { code, type } = req.query;
  let row = DATA.find(r => r.Code == code);
  res.redirect(row[`${type}_File`]);
});

// ================= START =================
app.listen(PORT, async () => {
  console.log("Server running...");
  await loadExcel();
});