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
const creds = JSON.parse(process.env.GOOGLE_CREDENTIALS);
creds.private_key = creds.private_key.replace(/\\n/g, '\n');

const auth = new google.auth.GoogleAuth({
  credentials: creds,
  scopes: ['https://www.googleapis.com/auth/drive']
});

const drive = google.drive({ version: 'v3', auth });
const FOLDER_ID = '168KzEusKlXsHQ-votUNTNA9g0VIai4X-';

// ================= DATA =================
let DATA = [];
let saved = [];

if (fs.existsSync('data.json')) {
  saved = JSON.parse(fs.readFileSync('data.json'));
}

// ================= SAVE =================
function saveToFile() {
  fs.writeFileSync('data.json', JSON.stringify(DATA, null, 2));
}

// ================= LOAD EXCEL =================
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

    if (!list.data.files.length) {
      console.log("❌ No Excel found");
      return;
    }

    const file = list.data.files[0];
    console.log("✅ Using Excel:", file.name);

    const dest = fs.createWriteStream("temp.xlsx");

    const res = await drive.files.get(
      { fileId: file.id, alt: 'media', supportsAllDrives: true },
      { responseType: 'stream' }
    );

    await new Promise((resolve, reject) => {
      res.data.pipe(dest).on('finish', resolve).on('error', reject);
    });

    const wb = XLSX.readFile("temp.xlsx");
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const raw = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    const savedMap = {};
    saved.forEach(s => savedMap[s.Code] = s);

    DATA = raw.map(row => {

      let r = {};
      Object.keys(row).forEach(k => r[k.trim()] = row[k]);

      const old = savedMap[r["Code"]] || {};

      return {
        STATE: r["STATE"] || "",
        BM_HQ: r["BM_HQ"] || "",
        Code: r["Code"] || "",
        Name: r["Stockist Name"] || "",

        // ✅ EMAILS (FIXED)
        BH_Email: r["BH_Email"] || "",
        SM_Email: r["SM_Email"] || "",
        ZBM_Email: r["ZBM_Email"] || "",
        RBM_Email: r["RBM_Email"] || "",
        ABM_Email: r["ABM_Email"] || "",

        // ✅ MERGED DATA
        Value: old.Value || "",

        SSS: old.SSS || false,
        AWS: old.AWS || false,

        SSS_File: old.SSS_File || "",
        AWS_File: old.AWS_File || "",

        SSS_Submitted_By: old.SSS_Submitted_By || "",
        AWS_Submitted_By: old.AWS_Submitted_By || "",

        SSS_Date: old.SSS_Date || "",
        AWS_Date: old.AWS_Date || ""
      };

    }).filter(r => r.Code && r.Name);

    fs.unlinkSync("temp.xlsx");

    console.log("✅ Loaded rows:", DATA.length);

  } catch (err) {
    console.log("❌ Load error:", err.message);
  }
}

// ================= AUTH =================
app.post('/login', (req, res) => {

  const { type, email, username, password } = req.body;

  if (type === "admin") {
    if (username === "admin" && password === "admin123") {
      req.session.user = { role: "admin" };
      return res.send("success");
    }
    return res.send("fail");
  }

  req.session.user = { role: "user", email };
  res.send("success");
});

// ================= GET DATA =================
app.get('/getData', (req, res) => {

  if (!req.session.user) return res.json({ role: "", data: [] });

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

// ================= SAVE VALUE =================
app.post('/saveValue', (req, res) => {

  const { code, value } = req.body;

  DATA.forEach(r => {
    if (r.Code == code) r.Value = value;
  });

  saveToFile();

  res.json({ status: "saved" });
});

// ================= UPLOAD =================
app.post('/uploadFile', upload.single('file'), async (req, res) => {

  try {

    const { code, type } = req.body;
    const file = req.file;

    let row = DATA.find(r => r.Code == code);

    // ✅ VALUE CHECK
    if (!row.Value || row.Value.trim() === "") {
      fs.unlinkSync(file.path);
      return res.json({ status: "error", msg: "Enter Value first" });
    }

    const ext = path.extname(file.originalname).toLowerCase();
    const allowed = ['.pdf', '.doc', '.docx', '.xls', '.xlsx', '.txt', '.html'];

    if (!allowed.includes(ext)) {
      fs.unlinkSync(file.path);
      return res.json({ status: "error", msg: "Invalid format" });
    }

    // ✅ PDF VALIDATION
    if (ext === '.pdf') {
      const parsed = await pdfParse(fs.readFileSync(file.path));
      if (!parsed.text || parsed.text.trim().length < 10) {
        fs.unlinkSync(file.path);
        return res.json({ status: "error", msg: "Invalid PDF" });
      }
    }

    // ✅ FILE NAME FIX
    const safe = row.Name.replace(/[^a-zA-Z0-9]/g, "_");
    const newName = `${safe}_${row.Code}_${type}${ext}`;

    const uploadRes = await drive.files.create({
      requestBody: { name: newName, parents: [FOLDER_ID] },
      media: { body: fs.createReadStream(file.path) },
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

    // ✅ UPDATE DATA
    row[type] = true;
    row[`${type}_File`] = link;
    row[`${type}_Submitted_By`] = req.session.user.email || "admin";
    row[`${type}_Date`] = new Date().toLocaleString();

    saveToFile();

    res.json({ status: "success" });

  } catch (err) {
    console.log(err);
    res.json({ status: "error", msg: "Upload failed" });
  }
});

// ================= DELETE =================
app.post('/deleteFile', (req, res) => {

  if (req.session.user.role !== "admin") {
    return res.json({ status: "error" });
  }

  const { code, type } = req.body;

  let row = DATA.find(r => r.Code == code);

  row[type] = false;
  row[`${type}_File`] = "";
  row[`${type}_Submitted_By`] = "";
  row[`${type}_Date`] = "";

  saveToFile();

  res.json({ status: "deleted" });
});

// ================= VIEW =================
app.get('/viewFile', (req, res) => {
  const { code, type } = req.query;
  let row = DATA.find(r => r.Code == code);
  res.redirect(row[`${type}_File`]);
});

// ================= REFRESH =================
app.get('/refreshData', async (req, res) => {
  await loadExcel();
  res.send("refreshed");
});

// ================= START =================
app.listen(PORT, async () => {
  console.log("🚀 Server running...");
  await loadExcel();
});