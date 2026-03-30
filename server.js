require('dotenv').config();

const express = require('express');
const session = require('express-session');
const multer = require('multer');
const { google } = require('googleapis');
const fs = require('fs');
const xlsx = require('xlsx');
const pdfParse = require('pdf-parse');
const { Readable } = require('stream');

const app = express();
const upload = multer({ dest: 'uploads/' });

const PORT = process.env.PORT || 3000;

// ================= ENV =================
const FOLDER_ID = process.env.FOLDER_ID;
const EXCEL_FILE_ID = process.env.EXCEL_FILE_ID;
const JSON_FILE_NAME = "updates.json";

// ================= GOOGLE AUTH =================
const auth = new google.auth.GoogleAuth({
  credentials: JSON.parse(process.env.GOOGLE_CREDENTIALS),
  scopes: ['https://www.googleapis.com/auth/drive']
});

const drive = google.drive({ version: 'v3', auth });

// ================= MIDDLEWARE =================
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

app.use(session({
  secret: process.env.SESSION_SECRET || 'secret',
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

// ================= EXCEL (FIXED) =================
async function downloadExcel() {
  try {
    console.log("📥 Downloading Excel...");

    // Try Google Sheets export first
    try {
      const res = await drive.files.export(
        {
          fileId: EXCEL_FILE_ID,
          mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        },
        { responseType: 'arraybuffer' }
      );

      const wb = xlsx.read(res.data, { type: 'buffer' });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const data = xlsx.utils.sheet_to_json(sheet);

      console.log("✅ Excel rows (Sheets):", data.length);
      return data;

    } catch (err) {
      console.log("⚠️ Not Google Sheet, trying direct file...");
    }

    // Fallback for Excel file
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

    fs.unlinkSync('temp.xlsx');

    console.log("✅ Excel rows (File):", data.length);
    return data;

  } catch (err) {
    console.error("❌ Excel Error:", err.message);
    return [];
  }
}

// ================= JSON =================
async function getUpdatesFile() {
  const res = await drive.files.list({
    q: `name='${JSON_FILE_NAME}' and '${FOLDER_ID}' in parents and trashed=false`,
    fields: 'files(id)'
  });

  if (res.data.files.length === 0) {
    const file = await drive.files.create({
      resource: { name: JSON_FILE_NAME, parents: [FOLDER_ID] },
      media: {
        mimeType: 'application/json',
        body: Readable.from([JSON.stringify({})])
      },
      fields: 'id'
    });

    console.log("🆕 Created updates.json");
    return file.data.id;
  }

  return res.data.files[0].id;
}

async function readUpdates() {
  try {
    const fileId = await getUpdatesFile();

    const res = await drive.files.get(
      { fileId, alt: 'media' },
      { responseType: 'stream' }
    );

    let data = '';

    return new Promise((resolve, reject) => {
      res.data.on('data', chunk => data += chunk);
      res.data.on('end', () => resolve(JSON.parse(data || "{}")));
      res.data.on('error', reject);
    });

  } catch (err) {
    console.error("❌ JSON Read Error:", err.message);
    return {};
  }
}

async function writeUpdates(content) {
  const fileId = await getUpdatesFile();

  await drive.files.update({
    fileId,
    media: {
      mimeType: 'application/json',
      body: Readable.from([JSON.stringify(content)])
    }
  });
}

// ================= MERGE =================
function mergeData(master, updates) {
  return master.map(row => {

    const code = row.Code;

    const sss = updates[`${code}_SSS`];
    const aws = updates[`${code}_AWS`];

    return {
      ...row,
      SSS_Status: sss ? "Done" : "Pending",
      AWS_Status: aws ? "Done" : "Pending",
      SSS_Files: sss?.files || [],
      AWS_Files: aws?.files || [],
      SSS_Value: sss?.value || '',
      AWS_Value: aws?.value || ''
    };
  });
}

// ================= GET DATA =================
app.get('/getData', async (req, res) => {

  console.log("🔥 API HIT /getData");

  const master = await downloadExcel();
  const updates = await readUpdates();

  console.log("📊 MASTER:", master.length);

  const rows = mergeData(master, updates);

  res.json({ rows, uploads: updates, role: req.session.role });
});

// ================= START =================
app.listen(PORT, () => console.log(`🚀 Server running on port ${PORT}`));