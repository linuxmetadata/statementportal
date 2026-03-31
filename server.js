require("dotenv").config();
const express = require("express");
const multer = require("multer");
const session = require("express-session");
const { google } = require("googleapis");
const fs = require("fs");
const path = require("path");
const xlsx = require("xlsx");

const app = express();

// ================= BASIC =================
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

app.use(session({
  secret: "portal_secret",
  resave: false,
  saveUninitialized: true,
}));

// ================= GOOGLE SERVICE ACCOUNT =================
const serviceKey = JSON.parse(process.env.GOOGLE_SERVICE_KEY);

const auth = new google.auth.GoogleAuth({
  credentials: serviceKey,
  scopes: ["https://www.googleapis.com/auth/drive"],
});

const drive = google.drive({ version: "v3", auth });

// ================= GOOGLE OAUTH =================
const oauth2Client = new google.auth.OAuth2(
  process.env.CLIENT_ID,
  process.env.CLIENT_SECRET,
  process.env.REDIRECT_URI
);

// ================= MULTER =================
const upload = multer({ dest: "uploads/" });

// ================= AUTH =================
function isAuth(req, res, next) {
  if (req.session.user) return next();
  res.redirect("/");
}

// ================= GOOGLE LOGIN =================
app.get("/auth/google", (req, res) => {
  const url = oauth2Client.generateAuthUrl({
    access_type: "offline",
    scope: ["https://www.googleapis.com/auth/userinfo.email"],
  });
  res.redirect(url);
});

app.get("/auth/google/callback", async (req, res) => {
  try {
    const { code } = req.query;

    const { tokens } = await oauth2Client.getToken(code);
    oauth2Client.setCredentials(tokens);

    const oauth2 = google.oauth2({
      auth: oauth2Client,
      version: "v2",
    });

    const userInfo = await oauth2.userinfo.get();
    const email = userInfo.data.email;

    if (email !== process.env.ADMIN_EMAIL) {
      return res.send("Not Authorized ❌");
    }

    req.session.user = { role: "admin", email };
    res.redirect("/dashboard");

  } catch (err) {
    console.error("OAuth Error:", err);
    res.send("Google Login Failed ❌");
  }
});

// ================= GET EXCEL DATA =================
async function getExcelData() {
  try {
    console.log("📥 Fetching Google Sheet...");

    const res = await drive.files.export({
      fileId: process.env.EXCEL_FILE_ID,
      mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }, { responseType: "arraybuffer" });

    const wb = xlsx.read(res.data, { type: "buffer" });

    // 👉 Change index if your data is in another sheet
    const sheetName = wb.SheetNames[0];
    const sheet = wb.Sheets[sheetName];

    const data = xlsx.utils.sheet_to_json(sheet);

    console.log("📊 Sheet:", sheetName);
    console.log("📊 Rows Loaded:", data.length);

    return data;

  } catch (e) {
    console.error("❌ Excel Error:", e.message);
    return [];
  }
}

// ================= USER LOGIN (ID BASED) =================
app.post("/user-login", async (req, res) => {
  try {
    const { empId } = req.body;

    const data = await getExcelData();

    const matched = data.filter(r =>
      r.BH_ID?.toString().trim() === empId.trim() ||
      r.SM_ID?.toString().trim() === empId.trim() ||
      r.ZBM_ID?.toString().trim() === empId.trim() ||
      r.RBM_ID?.toString().trim() === empId.trim() ||
      r.ABM_ID?.toString().trim() === empId.trim()
    );

    if (matched.length === 0) {
      return res.json({ success: false });
    }

    req.session.user = { empId, role: "user" };
    res.json({ success: true });

  } catch (e) {
    console.error(e);
    res.json({ success: false });
  }
});

// ================= JSON STORAGE =================
const DATA_FILE = "portal_data.json";

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

// ================= GET DATA =================
app.get("/getData", isAuth, async (req, res) => {

  const excel = await getExcelData();
  const uploads = await loadData();

  if (req.session.user.role === "admin") {
    return res.json({ rows: excel, uploads });
  }

  const empId = req.session.user.empId;

  const filtered = excel.filter(r =>
    r.BH_ID?.toString().trim() === empId.trim() ||
    r.SM_ID?.toString().trim() === empId.trim() ||
    r.ZBM_ID?.toString().trim() === empId.trim() ||
    r.RBM_ID?.toString().trim() === empId.trim() ||
    r.ABM_ID?.toString().trim() === empId.trim()
  );

  res.json({ rows: filtered, uploads });
});

// ================= UPLOAD =================
app.post("/upload", isAuth, upload.array("files"), async (req, res) => {
  try {
    const { code, state, name, type, value } = req.body;

    if (!value) return res.json({ message: "Value required ❗" });

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
      const up = await drive.files.create({
        requestBody: {
          name: `${name}_${code}_${type}_${Date.now()}`,
          parents: [folder.data.id],
        },
        media: {
          mimeType: f.mimetype,
          body: fs.createReadStream(f.path),
        },
      });

      links.push(`https://drive.google.com/file/d/${up.data.id}/view`);
      fs.unlinkSync(f.path);
    }

    data[key] = { value, links };

    await saveData(data);

    res.json({ message: "Upload Success ✅" });

  } catch (err) {
    console.error(err);
    res.json({ message: "Upload Failed ❌" });
  }
});

// ================= STATIC =================
app.use(express.static(path.join(__dirname, "public")));

app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "public/login.html"));
});

app.get("/dashboard", isAuth, (req, res) => {
  res.sendFile(path.join(__dirname, "public/dashboard.html"));
});

// ================= START =================
app.listen(process.env.PORT || 10000, () => {
  console.log("🚀 FINAL SERVER RUNNING");
});