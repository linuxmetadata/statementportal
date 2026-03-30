require('dotenv').config();

const express = require('express');
const { google } = require('googleapis');
const xlsx = require('xlsx');

const app = express();
const PORT = process.env.PORT || 3000;

// ================= GOOGLE AUTH =================
const auth = new google.auth.GoogleAuth({
  credentials: JSON.parse(process.env.GOOGLE_CREDENTIALS),
  scopes: ['https://www.googleapis.com/auth/drive']
});

const drive = google.drive({ version: 'v3', auth });

// ================= TEST EXCEL =================
async function downloadExcel() {
  try {
    console.log("📥 Trying Google Sheets export...");

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

    console.log("✅ Excel rows:", data.length);

    return data;

  } catch (err) {
    console.error("❌ Excel ERROR:", err.message);
    return [];
  }
}

// ================= API =================
app.get('/getData', async (req, res) => {

  console.log("🔥 API HIT");

  const data = await downloadExcel();

  res.json({ rows: data });

});

// ================= START =================
app.listen(PORT, () => console.log(`🚀 Running on ${PORT}`));