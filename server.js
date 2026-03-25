const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const session = require('express-session');
const pdfParse = require('pdf-parse');
const archiver = require('archiver');

const app = express();

// ===== MIDDLEWARE =====
app.use(express.static('public'));
app.use(express.json());

app.use(session({
  secret: 'secret123',
  resave: false,
  saveUninitialized: true
}));

// ===== ROOT FIX (IMPORTANT) =====
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'login.html'));
});

// ===== PORT =====
const PORT = process.env.PORT || 3000;

// ===== STORAGE =====
const upload = multer({ dest: 'uploads/' });

// ===== DATA FUNCTIONS =====
function readData() {
  try {
    if (!fs.existsSync('data.json')) return [];
    return JSON.parse(fs.readFileSync('data.json'));
  } catch {
    return [];
  }
}

function writeData(data) {
  fs.writeFileSync('data.json', JSON.stringify(data, null, 2));
}

// ===== EMAIL MATCH =====
function emailMatch(row, email) {
  email = (email || "").toLowerCase();

  return [
    row.BH_Email,
    row.SM_Email,
    row.ZBM_Email,
    row.RBM_Email,
    row.ABM_Email
  ].some(e => (e || "").toLowerCase().includes(email));
}

// ===== LOGIN =====
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
    let data = readData();

    const exists = data.some(r => emailMatch(r, email));

    if (!exists) return res.send("fail");

    req.session.user = { role: "user", email };
    return res.send("success");
  }

  res.send("fail");
});

// ===== UPLOAD EXCEL =====
app.post('/uploadExcel', upload.single('file'), (req, res) => {

  if (!req.session.user || req.session.user.role !== "admin") {
    return res.send("Access denied");
  }

  try {

    const wb = XLSX.readFile(req.file.path);
    const sheet = wb.Sheets[wb.SheetNames[0]];
    let data = XLSX.utils.sheet_to_json(sheet);

    data = data.map(row => ({
      BH_Email: row["BH_Email"] || "",
      SM_Email: row["SM_Email"] || "",
      ZBM_Email: row["ZBM_Email"] || "",
      RBM_Email: row["RBM_Email"] || "",
      ABM_Email: row["ABM_Email"] || "",

      STATE: row["STATE"] || "",
      BM_HQ: row["BM HQ"] || row["BM_HQ"] || "",
      Code: row["Stockist Code"] || row["CODE"] || "",
      Name: row["Stockist Name"] || row["STOCKIST NAME"] || "",

      Value: "",

      SSS: false,
      AWS: false,

      SSS_File: "",
      AWS_File: "",

      SSS_Submitted_By: "",
      AWS_Submitted_By: "",

      SSS_Date: "",
      AWS_Date: ""
    }));

    writeData(data);
    fs.unlinkSync(req.file.path);

    res.send("uploaded");

  } catch (err) {
    console.log(err);
    res.send("error uploading excel");
  }
});

// ===== GET DATA =====
app.get('/getData', (req, res) => {

  if (!req.session.user) return res.json([]);

  let data = readData();

  if (req.session.user.role === "admin") {
    return res.json({ role: "admin", data });
  }

  let email = req.session.user.email;
  let filtered = data.filter(r => emailMatch(r, email));

  res.json({ role: "user", data: filtered });
});

// ===== SAVE VALUE =====
app.post('/saveValue', (req, res) => {

  const { code, value } = req.body;

  let data = readData();

  data.forEach(r => {
    if (r.Code == code) {
      r.Value = value;
    }
  });

  writeData(data);

  res.send("saved");
});

// ===== FILE UPLOAD =====
app.post('/uploadFile', upload.single('file'), async (req, res) => {

  try {

    const { code, type } = req.body;
    const file = req.file;

    if (!file) return res.send("No file");

    const allowed = ['.pdf', '.doc', '.docx', '.xls', '.xlsx', '.txt', '.html'];
    const ext = path.extname(file.originalname).toLowerCase();

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

    let data = readData();
    let row = data.find(r => r.Code == code);

    if (!row || !row.Value) {
      fs.unlinkSync(file.path);
      return res.send("Enter Value first");
    }

    const state = row.STATE || "UNKNOWN";
    const folder = path.join('uploads', state, type);

    fs.mkdirSync(folder, { recursive: true });

    const cleanName = (row.Name || "FILE").replace(/[^a-zA-Z0-9]/g, "_");
    const newName = `${cleanName}_${code}_${type}${ext}`;
    const newPath = path.join(folder, newName);

    fs.renameSync(file.path, newPath);

    row[type] = true;
    row[`${type}_File`] = newPath;
    row[`${type}_Submitted_By`] = req.session.user.email || "admin";
    row[`${type}_Date`] = new Date().toLocaleString();

    writeData(data);

    res.send("uploaded");

  } catch (err) {
    console.log(err);
    res.send("error");
  }
});

// ===== VIEW FILE =====
app.get('/viewFile', (req, res) => {

  const { code, type } = req.query;

  let data = readData();
  let row = data.find(r => r.Code == code);

  if (!row) return res.send("Not found");

  let filePath = row[`${type}_File`];

  if (!filePath || !fs.existsSync(filePath)) {
    return res.send("File not found");
  }

  res.sendFile(path.resolve(filePath));
});

// ===== DELETE FILE =====
app.post('/deleteFile', (req, res) => {

  if (!req.session.user || req.session.user.role !== "admin") {
    return res.send("Access denied");
  }

  const { code, type } = req.body;

  let data = readData();
  let row = data.find(r => r.Code == code);

  if (!row) return res.send("Not found");

  let filePath = row[`${type}_File`];

  if (filePath && fs.existsSync(filePath)) {
    fs.unlinkSync(filePath);
  }

  row[type] = false;
  row[`${type}_File`] = "";

  writeData(data);

  res.send("deleted");
});

// ===== DASHBOARD =====
app.get('/dashboard', (req, res) => {

  if (!req.session.user) return res.json({});

  let data = readData();

  if (req.session.user.role === "user") {
    let email = req.session.user.email;
    data = data.filter(r => emailMatch(r, email));
  }

  const total = data.length;
  const sssReceived = data.filter(r => r.SSS).length;
  const awsReceived = data.filter(r => r.AWS).length;

  const sssPending = total - sssReceived;
  const awsPending = total - awsReceived;

  function percent(val) {
    return total === 0 ? 0 : ((val / total) * 100).toFixed(1);
  }

  res.json({
    total,
    sssReceived,
    sssPending,
    awsReceived,
    awsPending,
    sssReceivedPercent: percent(sssReceived),
    sssPendingPercent: percent(sssPending),
    awsReceivedPercent: percent(awsReceived),
    awsPendingPercent: percent(awsPending)
  });
});

// ===== DOWNLOAD REPORT =====
app.get('/downloadReport', (req, res) => {

  let data = readData();

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(data);

  XLSX.utils.book_append_sheet(wb, ws, "Report");

  const file = "report.xlsx";
  XLSX.writeFile(wb, file);

  res.download(file);
});

// ===== DOWNLOAD ZIP =====
app.get('/downloadAllFiles', (req, res) => {

  if (!req.session.user || req.session.user.role !== "admin") {
    return res.send("Access denied");
  }

  const archive = archiver('zip');

  res.attachment('All_Files.zip');
  archive.pipe(res);

  if (fs.existsSync('uploads')) {
    archive.directory('uploads', false);
  }

  archive.finalize();
});

// ===== START SERVER =====
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});