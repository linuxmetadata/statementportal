require("dotenv").config();
const express = require("express");
const multer = require("multer");
const { google } = require("googleapis");
const fs = require("fs");
const path = require("path");
const xlsx = require("xlsx");
const archiver = require("archiver");
const pdfParse = require("pdf-parse");

const app = express();
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// ================= GOOGLE AUTH =================
const oauth2Client = new google.auth.OAuth2(
  process.env.CLIENT_ID,
  process.env.CLIENT_SECRET,
  process.env.REDIRECT_URI
);

if (fs.existsSync("token.json")) {
  oauth2Client.setCredentials(JSON.parse(fs.readFileSync("token.json")));
}

const drive = google.drive({ version: "v3", auth: oauth2Client });

// ================= MULTER =================
const upload = multer({ dest: "uploads/" });

// ================= CONFIG =================
const DATA_FILE = "portal_data.json";

// ================= EXCEL =================
async function getExcelData() {
  const res = await drive.files.export(
    {
      fileId: process.env.EXCEL_FILE_ID,
      mimeType:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    },
    { responseType: "arraybuffer" }
  );

  const wb = xlsx.read(res.data, { type: "buffer" });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  return xlsx.utils.sheet_to_json(sheet);
}

// ================= DATA FILE =================
async function loadData() {
  try {
    const list = await drive.files.list({
      q: `name='${DATA_FILE}'`,
      fields: "files(id)",
    });

    if (!list.data.files.length) return {};

    const file = await drive.files.get(
      { fileId: list.data.files[0].id, alt: "media" },
      { responseType: "stream" }
    );

    let data = "";
    for await (const chunk of file.data) data += chunk;

    return JSON.parse(data || "{}");
  } catch {
    return {};
  }
}

async function saveData(data) {
  const content = Buffer.from(JSON.stringify(data, null, 2));

  const list = await drive.files.list({
    q: `name='${DATA_FILE}'`,
    fields: "files(id)",
  });

  if (!list.data.files.length) {
    await drive.files.create({
      requestBody: { name: DATA_FILE },
      media: { mimeType: "application/json", body: content },
    });
  } else {
    await drive.files.update({
      fileId: list.data.files[0].id,
      media: { mimeType: "application/json", body: content },
    });
  }
}

// ================= LOGIN =================
app.post("/login", async (req, res) => {
  const { empId } = req.body;
  const excel = await getExcelData();

  const user = excel.find(
    (r) => String(r["Emp ID"]).trim() === empId.trim()
  );

  if (!user) return res.json({ success: false });

  res.json({
    success: true,
    role: empId === "admin" ? "admin" : "user",
  });
});

// ================= GET DATA =================
app.get("/getData", async (req, res) => {
  const excel = await getExcelData();
  const uploads = await loadData();

  res.json({ rows: excel, uploads });
});

// ================= FILE VALIDATION =================
async function validateFile(file) {
  const allowed = [
    "application/pdf",
    "application/msword",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "application/vnd.ms-excel",
    "text/plain",
    "text/html",
  ];

  if (!allowed.includes(file.mimetype)) return false;

  if (file.mimetype === "application/pdf") {
    const buffer = fs.readFileSync(file.path);
    const data = await pdfParse(buffer);
    return data.text && data.text.length > 10;
  }

  return true;
}

// ================= UPLOAD =================
app.post("/upload", upload.array("files"), async (req, res) => {
  try {
    const { code, state, name, type, value, empId } = req.body;

    if (!value) return res.json({ message: "Value mandatory ❗" });

    const data = await loadData();
    const key = `${code}_${type}`;

    if (data[key]) return res.json({ message: "Already uploaded ❌" });

    const folder = await drive.files.create({
      requestBody: {
        name: `${state}/${type}`,
        mimeType: "application/vnd.google-apps.folder",
        parents: [process.env.FOLDER_ID],
      },
    });

    const links = [];

    for (let f of req.files) {
      if (!(await validateFile(f))) {
        fs.unlinkSync(f.path);
        return res.json({ message: "Invalid format ❌" });
      }

      const uploadRes = await drive.files.create({
        requestBody: {
          name: `${name}_${code}_${type}_${Date.now()}`,
          parents: [folder.data.id],
        },
        media: {
          mimeType: f.mimetype,
          body: fs.createReadStream(f.path),
        },
      });

      links.push(`https://drive.google.com/file/d/${uploadRes.data.id}/view`);
      fs.unlinkSync(f.path);
    }

    data[key] = {
      value,
      links,
      submittedBy: empId,
      date: new Date().toLocaleString(),
      type,
      state,
    };

    await saveData(data);

    res.json({ message: "Upload Success ✅" });
  } catch (e) {
    console.error(e);
    res.json({ message: "Upload Failed ❌" });
  }
});

// ================= DELETE =================
app.post("/delete", async (req, res) => {
  const data = await loadData();
  delete data[req.body.key];
  await saveData(data);
  res.json({ message: "Deleted ✅" });
});

// ================= REPORT DOWNLOAD =================
app.get("/download-report", async (req, res) => {
  const excel = await getExcelData();
  const uploads = await loadData();

  const report = excel.map((r) => {
    const sss = uploads[r.Code + "_SSS"];
    const aws = uploads[r.Code + "_AWS"];

    return {
      ...r,
      SSS_Status: sss ? "Done" : "Pending",
      AWS_Status: aws ? "Done" : "Pending",
      Submitted_By: sss?.submittedBy || aws?.submittedBy || "",
      Date: sss?.date || aws?.date || "",
    };
  });

  const wb = xlsx.utils.book_new();
  const ws = xlsx.utils.json_to_sheet(report);
  xlsx.utils.book_append_sheet(wb, ws, "Report");

  const file = "report.xlsx";
  xlsx.writeFile(wb, file);

  res.download(file);
});

// ================= STATIC =================
app.use(express.static(path.join(__dirname, "public")));

app.get("/", (req, res) =>
  res.sendFile(path.join(__dirname, "public/login.html"))
);

app.get("/dashboard", (req, res) =>
  res.sendFile(path.join(__dirname, "public/dashboard.html"))
);

app.listen(process.env.PORT || 10000, () =>
  console.log("🚀 Running")
);