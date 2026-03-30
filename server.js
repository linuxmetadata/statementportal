require("dotenv").config();
const express = require("express");
const session = require("express-session");
const multer = require("multer");
const { google } = require("googleapis");
const fs = require("fs");
const path = require("path");

const app = express();
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// SESSION
app.use(
  session({
    secret: "secret123",
    resave: false,
    saveUninitialized: true,
  })
);

// ================= GOOGLE AUTH =================

const auth = new google.auth.GoogleAuth({
  credentials: JSON.parse(process.env.GOOGLE_CREDENTIALS),
  scopes: ["https://www.googleapis.com/auth/drive"],
});

const drive = google.drive({ version: "v3", auth });

// ================= MULTER =================

const upload = multer({ dest: "uploads/" });

// ================= TEST ROUTE =================

app.get("/", (req, res) => {
  res.send(`
    <h2>Simple Upload Test</h2>
    <form action="/upload" method="post" enctype="multipart/form-data">
      <input type="file" name="file" required />
      <button type="submit">Upload</button>
    </form>
  `);
});

// ================= UPLOAD ROUTE =================

app.post("/upload", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) {
      return res.send("No file selected");
    }

    console.log("📂 File received:", req.file.originalname);

    const folderId = process.env.GOOGLE_FOLDER_ID;

    // FILE METADATA
    const fileMetadata = {
      name: req.file.originalname,
      parents: [folderId],
    };

    // MEDIA
    const media = {
      mimeType: req.file.mimetype,
      body: fs.createReadStream(req.file.path),
    };

    // UPLOAD TO DRIVE
    const response = await drive.files.create({
      resource: fileMetadata,
      media: media,
      fields: "id",
    });

    console.log("✅ Uploaded File ID:", response.data.id);

    // DELETE LOCAL FILE
    fs.unlinkSync(req.file.path);

    res.send("✅ Upload Success");
  } catch (err) {
    console.error("❌ Upload Error:", err.message);
    res.send("❌ Upload Failed");
  }
});

// ================= START =================

const PORT = process.env.PORT || 10000;
app.listen(PORT, () => console.log("🚀 Running on", PORT));