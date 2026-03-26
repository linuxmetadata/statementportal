const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const session = require('express-session');
const pdfParse = require('pdf-parse');
const archiver = require('archiver');
const axios = require('axios');
const { google } = require('googleapis');

const app = express();

app.use(express.static('public'));
app.use(express.json());

app.use(session({
  secret: 'secret123',
  resave: false,
  saveUninitialized: true
}));

const upload = multer({ dest: 'uploads/' });
const PORT = process.env.PORT || 3000;

// ================= GOOGLE AUTH =================
let auth;

try {
  const creds = JSON.parse(process.env.GOOGLE_CREDENTIALS);

  // fix newline issue
  creds.private_key = creds.private_key.replace(/\\n/g, '\n');

  auth = new google.auth.GoogleAuth({
    credentials: creds,
    scopes: ['https://www.googleapis.com/auth/drive']
  });

} catch (err) {
  console.error("Google Auth Error:", err.message);
}

const drive = google.drive({ version: 'v3', auth });
const FOLDER_ID = '168KzEusKlXsHQ-votUNTNA9g0VIai4X-';

let DATA = [];

// ================= HOME ROUTE =================
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'login.html'));
});

// ================= HELPERS =================
function checkAuth(req,res){
  if(!req.session.user){
    res.send("Login required");
    return false;
  }
  return true;
}

function cleanKey(obj){
  let o = {};
  Object.keys(obj).forEach(k=>{
    o[k.trim()] = obj[k];
  });
  return o;
}

// ================= DRIVE =================
async function uploadToDrive(filePath,name){
  const res = await drive.files.create({
    requestBody:{ name, parents:[FOLDER_ID] },
    media:{ body: fs.createReadStream(filePath) }
  });
  return res.data.id;
}

async function getLink(id){
  await drive.permissions.create({
    fileId:id,
    requestBody:{ role:'reader', type:'anyone' }
  });
  return `https://drive.google.com/uc?id=${id}`;
}

// ================= LOAD EXCEL =================
async function loadExcel(){
  try {
    const list = await drive.files.list({
      q:`'${FOLDER_ID}' in parents and name='MASTER_EXCEL.xlsx'`,
      fields:'files(id)'
    });

    if(list.data.files.length===0) return;

    const fileId = list.data.files[0].id;

    const dest = fs.createWriteStream("temp.xlsx");

    const res = await drive.files.get(
      { fileId, alt:'media' },
      { responseType:'stream' }
    );

    await new Promise((resolve,reject)=>{
      res.data.pipe(dest).on('finish',resolve).on('error',reject);
    });

    const wb = XLSX.readFile("temp.xlsx");
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const raw = XLSX.utils.sheet_to_json(sheet);

    DATA = raw.map(r=>{
      let row = cleanKey(r);
      return {
        STATE: row["State"] || "",
        BM_HQ: row["BM HQ"] || "",
        Code: row["Code"] || "",
        Name: row["Stockist Name"] || "",

        BH_Email: row["BH_Email"] || "",
        SM_Email: row["SM_Email"] || "",
        ZBM_Email: row["ZBM_Email"] || "",
        RBM_Email: row["RBM_Email"] || "",
        ABM_Email: row["ABM_Email"] || "",

        Value:"",
        SSS:false,
        AWS:false,

        SSS_File:"",
        AWS_File:"",

        SSS_Submitted_By:"",
        AWS_Submitted_By:"",
        SSS_Date:"",
        AWS_Date:""
      }
    });

    fs.unlinkSync("temp.xlsx");
    console.log("Loaded rows:", DATA.length);

  } catch (err) {
    console.log("Excel load error:", err.message);
  }
}

// ================= LOGIN =================
app.post('/login',(req,res)=>{
  const {type,email,username,password} = req.body;

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

// ================= UPLOAD EXCEL =================
app.post('/uploadExcel', upload.single('file'), async (req,res)=>{

  if(!checkAuth(req,res)) return;
  if(req.session.user.role!=="admin") return res.send("Access denied");

  const existing = await drive.files.list({
    q:`'${FOLDER_ID}' in parents and name='MASTER_EXCEL.xlsx'`,
    fields:'files(id)'
  });

  if(existing.data.files.length>0){
    await drive.files.delete({ fileId: existing.data.files[0].id });
  }

  await uploadToDrive(req.file.path,"MASTER_EXCEL.xlsx");
  fs.unlinkSync(req.file.path);

  await loadExcel();

  res.send("Excel uploaded");
});

// ================= GET DATA =================
app.get('/getData',(req,res)=>{
  if(!checkAuth(req,res)) return;

  let result = DATA;

  if(req.session.user.role==="user"){
    const email = req.session.user.email;

    result = DATA.filter(r =>
      r.BH_Email===email ||
      r.SM_Email===email ||
      r.ZBM_Email===email ||
      r.RBM_Email===email ||
      r.ABM_Email===email
    );
  }

  res.json({ role:req.session.user.role, data:result });
});

// ================= SAVE VALUE =================
app.post('/saveValue',(req,res)=>{
  const {code,value} = req.body;

  DATA.forEach(r=>{
    if(r.Code==code) r.Value=value;
  });

  res.send("saved");
});

// ================= UPLOAD FILE =================
app.post('/uploadFile', upload.single('file'), async (req,res)=>{

  const {code,type} = req.body;
  const file = req.file;

  const ext = path.extname(file.originalname).toLowerCase();
  const allowed = ['.pdf','.doc','.docx','.xls','.xlsx','.txt','.html'];

  if(!allowed.includes(ext)){
    fs.unlinkSync(file.path);
    return res.send("Invalid format");
  }

  if(ext==='.pdf'){
    const buffer = fs.readFileSync(file.path);
    const parsed = await pdfParse(buffer);

    if(!parsed.text || parsed.text.trim().length<10){
      fs.unlinkSync(file.path);
      return res.send("Invalid PDF");
    }
  }

  let row = DATA.find(r=>r.Code==code);

  if(!row.Value){
    fs.unlinkSync(file.path);
    return res.send("Enter Value first");
  }

  const safe = row.Name.replace(/[^a-zA-Z0-9]/g,"_");
  const name = `${safe}_${code}_${type}${ext}`;

  const id = await uploadToDrive(file.path,name);
  const link = await getLink(id);

  fs.unlinkSync(file.path);

  row[type]=true;
  row[`${type}_File`]=link;
  row[`${type}_Submitted_By`]=req.session.user.email || "admin";
  row[`${type}_Date`]=new Date().toLocaleString();

  res.send("uploaded");
});

// ================= VIEW =================
app.get('/viewFile',(req,res)=>{
  const {code,type} = req.query;
  let row = DATA.find(r=>r.Code==code);
  res.redirect(row[`${type}_File`]);
});

// ================= DELETE =================
app.post('/deleteFile',(req,res)=>{
  const {code,type} = req.body;
  let row = DATA.find(r=>r.Code==code);

  row[type]=false;
  row[`${type}_File`]="";

  res.send("deleted");
});

// ================= DOWNLOAD REPORT =================
app.get('/downloadReport',(req,res)=>{

  let report = DATA.map(r=>({
    State:r.STATE,
    BM_HQ:r.BM_HQ,
    Code:r.Code,
    Name:r.Name,
    Value:r.Value,
    SSS_Status:r.SSS?"Done":"Pending",
    AWS_Status:r.AWS?"Done":"Pending",
    SSS_By:r.SSS_Submitted_By,
    SSS_Date:r.SSS_Date,
    AWS_By:r.AWS_Submitted_By,
    AWS_Date:r.AWS_Date
  }));

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(report);
  XLSX.utils.book_append_sheet(wb,ws,"Report");

  XLSX.writeFile(wb,"report.xlsx");
  res.download("report.xlsx");
});

// ================= DOWNLOAD ALL FILES =================
app.get('/downloadAllFiles', async (req,res)=>{

  const archive = archiver('zip');
  res.attachment('All.zip');
  archive.pipe(res);

  for(let r of DATA){

    if(r.SSS_File){
      const stream = await axios.get(r.SSS_File,{responseType:'stream'});
      archive.append(stream.data,{ name:`${r.STATE}/SSS/${r.Code}.pdf` });
    }

    if(r.AWS_File){
      const stream = await axios.get(r.AWS_File,{responseType:'stream'});
      archive.append(stream.data,{ name:`${r.STATE}/AWS/${r.Code}.pdf` });
    }

  }

  archive.finalize();
});

// ================= DOWNLOAD PENDING =================
app.get('/downloadPending',(req,res)=>{

  const wb = XLSX.utils.book_new();
  const pending = DATA.filter(r=>!r.SSS || !r.AWS);

  const ws = XLSX.utils.json_to_sheet(pending);
  XLSX.utils.book_append_sheet(wb,ws,"Pending");

  XLSX.writeFile(wb,"pending.xlsx");
  res.download("pending.xlsx");
});

// ================= START =================
app.listen(PORT, async ()=>{
  console.log("Server running...");
  await loadExcel();
});