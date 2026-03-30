require('dotenv').config();

const express = require('express');
const { google } = require('googleapis');
const xlsx = require('xlsx');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

// ================= GOOGLE AUTH =================
const auth = new google.auth.GoogleAuth({
  credentials: JSON.parse(process.env.GOOGLE_CREDENTIALS),
  scopes: ['https://www.googleapis.com/auth/drive.readonly']
});

const drive = google.drive({ version: 'v3', auth });

// ================= SERVE STATIC =================
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.json());

// ================= ROUTES =================

// Login Page
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'login.html'));
});

// Dashboard Page
app.get('/dashboard', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'dashboard.html'));
});

// ================= EXCEL FETCH =================
async function downloadExcel() {
  try {
    console.log("📥 Fetching Google Sheet...");

    const res = await drive.files.export(
      {
        fileId: process.env.EXCEL_FILE_ID,
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      },
      { responseType: 'arraybuffer' }
    );

    const wb = xlsx.read(res.data, { type: 'buffer' });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const data = xlsx.utils.sheet_to_json(sheet);

    console.log("✅ Rows:", data.length);

    return data;

  } catch (err) {
    console.error("❌ ERROR:", err.message);
    return [];
  }
}

// ================= API =================
app.get('/getData', async (req, res) => {
  const data = await downloadExcel();
  res.json({ rows: data });
});

// ================= START =================
app.listen(PORT, () => console.log(`🚀 Running on ${PORT}`));