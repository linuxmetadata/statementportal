const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const session = require('express-session');
const pdfParse = require('pdf-parse');
const { google } = require('googleapis');

const app = express();

// ===== SAFETY =====
process.on('uncaughtException', err => console.error('Uncaught:', err));
process.on('unhandledRejection', err => console.error('Unhandled:', err));

// ===== MIDDLEWARE =====
app.use(express.static('public'));
app.use(express.json());

app.use(session({
  secret: 'secret123',
  resave: false,
  saveUninitialized: true
}));

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'login.html'));
});

const PORT = process.env.PORT || 3000;
const upload = multer({ dest: 'uploads/' });

// ===== DEBUG ENV =====
console.log("ENV CHECK:", !!process.env.GOOGLE_CREDENTIALS);

// ===== GOOGLE DRIVE AUTH (SAFE VERSION) =====
let auth;

try {
  if (!process.env.GOOGLE_CREDENTIALS) {
    throw new Error("Missing GOOGLE_CREDENTIALS");
  }

  const parsed = JSON.parse(process.env.GOOGLE_CREDENTIALS);

  auth = new google.auth.GoogleAuth({
    credentials: parsed,
    scopes: ['https://www.googleapis.com/auth/drive']
  });

} catch (err) {
  console.error("❌ GOOGLE DRIVE INIT FAILED:", err.message);

  // fallback (prevents crash)
  auth = new google.auth.GoogleAuth({
    scopes: ['https://www.googleapis.com/auth/drive']
  });
}

const drive = google.drive({ version: 'v3', auth });

const FOLDER_ID = '168KzEusKlXsHQ-votUNTNA9g0VIai4X-';

// ===== GLOBAL DATA =====
let DATA = [];

// ===== HELPERS =====
function checkAuth(req, res){
  if(!req.session.user){
    res.send("Login required");
    return false;
  }
  return true;
}

// ===== DRIVE FUNCTIONS =====
async function uploadToDrive(filePath, fileName){
  const res = await drive.files.create({
    requestBody: {
      name: fileName,
      parents: [FOLDER_ID]
    },
    media: {
      mimeType: 'application/octet-stream',
      body: fs.createReadStream(filePath)
    }
  });
  return res.data.id;
}

async function getFileLink(fileId){
  await drive.permissions.create({
    fileId,
    requestBody: { role:'reader', type:'anyone' }
  });
  return `https://drive.google.com/uc?id=${fileId}`;
}

// ===== LOAD EXCEL =====
async function loadExcelFromDrive(){
  try{
    const res = await drive.files.list({
      q: `'${FOLDER_ID}' in parents and name='MASTER_EXCEL.xlsx'`,
      fields: 'files(id)'
    });

    if(res.data.files.length === 0){
      console.log("⚠ No Excel found");
      return;
    }

    const fileId = res.data.files[0].id;

    const dest = fs.createWriteStream("temp.xlsx");

    const response = await drive.files.get(
      { fileId, alt:'media' },
      { responseType:'stream' }
    );

    await new Promise((resolve,reject)=>{
      response.data.pipe(dest)
        .on('finish', resolve)
        .on('error', reject);
    });

    const wb = XLSX.readFile("temp.xlsx");
    const sheet = wb.Sheets[wb.SheetNames[0]];
    let data = XLSX.utils.sheet_to_json(sheet);

    DATA = data.map(row => ({
      STATE: row["STATE"] || row["State"] || "",
      BM_HQ: row["BM HQ"] || row["BM_HQ"] || "",
      Code: row["Stockist Code"] || row["Code"] || "",
      Name: row["Stockist Name"] || row["Name"] || "",

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

    fs.unlinkSync("temp.xlsx");

    console.log("✅ Excel loaded");

  }catch(err){
    console.error("❌ Excel load error:", err);
  }
}

// ===== LOGIN =====
app.post('/login', (req,res)=>{

  const { type, email, username, password } = req.body;

  if(type==="admin"){
    if(username==="admin" && password==="admin123"){
      req.session.user = { role:"admin" };
      return res.send("success");
    }
    return res.send("fail");
  }

  if(type==="user"){
    req.session.user = { role:"user", email };
    return res.send("success");
  }

});

// ===== UPLOAD EXCEL =====
app.post('/uploadExcel', upload.single('file'), async (req,res)=>{

  if(!checkAuth(req,res)) return;
  if(req.session.user.role !== "admin") return res.send("Access denied");

  try{

    // delete old excel
    const existing = await drive.files.list({
      q: `'${FOLDER_ID}' in parents and name='MASTER_EXCEL.xlsx'`,
      fields: 'files(id)'
    });

    if(existing.data.files.length > 0){
      await drive.files.delete({ fileId: existing.data.files[0].id });
    }

    // upload new
    await uploadToDrive(req.file.path, "MASTER_EXCEL.xlsx");

    fs.unlinkSync(req.file.path);

    await loadExcelFromDrive();

    res.send("Excel uploaded");

  }catch(err){
    console.error(err);
    res.send("Upload error");
  }
});

// ===== GET DATA =====
app.get('/getData',(req,res)=>{
  if(!checkAuth(req,res)) return;
  res.json({ role:req.session.user.role, data:DATA });
});

// ===== SAVE VALUE =====
app.post('/saveValue',(req,res)=>{
  if(!checkAuth(req,res)) return;

  const { code, value } = req.body;

  DATA.forEach(r=>{
    if(r.Code == code){
      r.Value = value;
    }
  });

  res.send("saved");
});

// ===== FILE UPLOAD =====
app.post('/uploadFile', upload.single('file'), async (req,res)=>{

  try{

    if(!checkAuth(req,res)) return;

    const { code, type } = req.body;
    const file = req.file;

    if(!file) return res.send("No file");

    const ext = path.extname(file.originalname).toLowerCase();

    const allowed = ['.pdf','.doc','.docx','.xls','.xlsx','.txt','.html'];

    if(!allowed.includes(ext)){
      fs.unlinkSync(file.path);
      return res.send("Invalid format");
    }

    if(ext === '.pdf'){
      const buffer = fs.readFileSync(file.path);
      const parsed = await pdfParse(buffer);

      if(!parsed.text || parsed.text.trim().length < 10){
        fs.unlinkSync(file.path);
        return res.send("Invalid PDF");
      }
    }

    let row = DATA.find(r => r.Code == code);

    if(!row){
      fs.unlinkSync(file.path);
      return res.send("Invalid code");
    }

    if(!row.Value){
      fs.unlinkSync(file.path);
      return res.send("Enter Value first");
    }

    const cleanName = row.Name.replace(/[^a-zA-Z0-9]/g,"_");
    const newName = `${cleanName}_${code}_${type}${ext}`;

    const fileId = await uploadToDrive(file.path, newName);
    const link = await getFileLink(fileId);

    fs.unlinkSync(file.path);

    row[type] = true;
    row[`${type}_File`] = link;
    row[`${type}_Submitted_By`] = req.session.user.email || "admin";
    row[`${type}_Date`] = new Date().toLocaleString();

    res.send("uploaded");

  }catch(err){
    console.error(err);
    res.send("Upload error");
  }
});

// ===== VIEW FILE =====
app.get('/viewFile',(req,res)=>{

  if(!checkAuth(req,res)) return;

  const { code, type } = req.query;

  let row = DATA.find(r => r.Code == code);

  if(!row) return res.send("Not found");

  let filePath = row[`${type}_File`];

  if(!filePath) return res.send("File not found");

  res.redirect(filePath);
});

// ===== START =====
app.listen(PORT, async ()=>{
  console.log("🚀 Server running on port " + PORT);
  await loadExcelFromDrive();
});