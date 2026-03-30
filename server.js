require("dotenv").config();
const express = require("express");
const multer = require("multer");
const { google } = require("googleapis");
const fs = require("fs");
const path = require("path");
const xlsx = require("xlsx");

const app = express();
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// ================= GOOGLE OAUTH =================

const oauth2Client = new google.auth.OAuth2(
  process.env.CLIENT_ID,
  process.env.CLIENT_SECRET,
  process.env.REDIRECT_URI
);

// LOGIN URL
app.get("/auth", (req, res) => {
  const url = oauth2Client.generateAuthUrl({
    access_type: "offline",
    scope: ["https://www.googleapis.com/auth/drive"],
  });
  res.redirect(url);
});

// CALLBACK
app.get("/oauth2callback", async (req, res) => {
  const code = req.query.code;

  const { tokens } = await oauth2Client.getToken(code);
  oauth2Client.setCredentials(tokens);

  // SAVE TOKEN
  fs.writeFileSync("token.json", JSON.stringify(tokens));

  res.send("✅ Google Drive Connected! You can close this.");
});

// LOAD TOKEN
if (fs.existsSync("token.json")) {
  oauth2Client.setCredentials(JSON.parse(fs.readFileSync("token.json")));
}

const drive = google.drive({ version: "v3", auth: oauth2Client });

// ================= MULTER =================
const upload = multer({ dest: "uploads/" });

// ================= EXCEL =================
async function getExcelData() {
  const response = await drive.files.export(
    {
      fileId: process.env.EXCEL_FILE_ID,
      mimeType:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    },
    { responseType: "arraybuffer" }
  );

  const workbook = xlsx.read(response.data, { type: "buffer" });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  return xlsx.utils.sheet_to_json(sheet);
}

// ================= LOGIN =================
app.post("/login", async (req, res) => {
  try {
    const { empId } = req.body;

    const data = await getExcelData();

    const user = data.find(
      (row) => String(row["Emp ID"]).trim() === String(empId).trim()
    );

    if (!user) {
      return res.json({ success: false, message: "Invalid Emp ID ❌" });
    }

    let role = empId === "admin" ? "admin" : "user";

    res.json({ success: true, role });

  } catch (err) {
    console.error(err);
    res.json({ success: false, message: "Login Error ❌" });
  }
});

// ================= UPLOAD =================
app.post("/upload", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) return res.json({ message: "No file ❌" });

    const response = await drive.files.create({
      requestBody: {
        name: req.file.originalname,
        parents: [process.env.FOLDER_ID],
      },
      media: {
        mimeType: req.file.mimetype,
        body: fs.createReadStream(req.file.path),
      },
    });

    fs.unlinkSync(req.file.path);

    res.json({ message: "Upload Success ✅" });

  } catch (err) {
    console.error(err);
    res.json({ message: "Upload Failed ❌" });
  }
});

// ================= STATIC =================
app.use(express.static(path.join(__dirname, "public")));

app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "login.html"));
});

app.get("/dashboard", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "dashboard.html"));
});

// ================= START =================
const PORT = process.env.PORT || 10000;
app.listen(PORT, () => console.log("🚀 Running on", PORT));