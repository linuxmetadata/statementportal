require("dotenv").config();
const express = require("express");
const multer = require("multer");
const session = require("express-session");
const { google } = require("googleapis");
const fs = require("fs");
const path = require("path");
const xlsx = require("xlsx");
const pdfParse = require("pdf-parse");

const app = express();

// ================= BASIC =================
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

app.use(session({
  secret: "portal_secret",
  resave: false,
  saveUninitialized: true,
}));

// ================= SAFE SERVICE ACCOUNT =================
let serviceKey = null;

try {
  if (!process.env.GOOGLE_SERVICE_KEY) {
    console.log("❌ GOOGLE_SERVICE_KEY missing");
  } else {
    serviceKey = JSON.parse(process.env.GOOGLE_SERVICE_KEY);
    console.log("✅ Service Key Loaded");
  }
} catch (e) {
  console.error("❌ SERVICE KEY ERROR:", e.message);
}

const auth = new google.auth.GoogleAuth({
  credentials: serviceKey,
  scopes: ["https://www.googleapis.com/auth/drive"],
});

const drive = google.drive({
  version: "v3",
  auth,
});

// ================= SAFE OAUTH =================
let oauth2Client;

try {
  if (!process.env.CLIENT_ID) console.log("❌ CLIENT_ID missing");
  if (!process.env.CLIENT_SECRET) console.log("❌ CLIENT_SECRET missing");
  if (!process.env.REDIRECT_URI) console.log("❌ REDIRECT_URI missing");

  oauth2Client = new google.auth.OAuth2(
    process.env.CLIENT_ID,
    process.env.CLIENT_SECRET,
    process.env.REDIRECT_URI
  );

  console.log("✅ OAuth Ready");

} catch (e) {
  console.error("❌ OAUTH ERROR:", e.message);
}

// ================= MULTER =================
const upload = multer({ dest: "uploads/" });

// ================= CONFIG =================
const DATA_FILE = "portal_data.json";

// ================= AUTH =================
function isAuth(req, res, next) {
  if (req.session.user) return next();
  res.redirect("/");
}

function isAdmin(req, res, next) {
  if (req.session.user?.role === "admin") return next();
  res.json({ message: "Admin only ❌" });
}

// ================= DEBUG ROUTE =================
app.get("/test", (req, res) => {
  res.json({
    CLIENT_ID: process.env.CLIENT_ID ? "OK" : "MISSING",
    CLIENT_SECRET: process.env.CLIENT_SECRET ? "OK" : "MISSING",
    REDIRECT_URI: process.env.REDIRECT_URI,
    SERVICE_KEY: process.env.GOOGLE_SERVICE_KEY ? "OK" : "MISSING"
  });
});

// ================= GOOGLE LOGIN =================
app.get("/auth/google", (req, res) => {
  try {
    if (!oauth2Client) return res.send("OAuth not initialized ❌");

    const url = oauth2Client.generateAuthUrl({
      access_type: "offline",
      scope: ["https://www.googleapis.com/auth/userinfo.email"],
    });

    res.redirect(url);

  } catch (err) {
    console.error("AUTH ERROR:", err);
    res.send("OAuth Error ❌");
  }
});

// ================= CALLBACK =================
app.get("/auth/google/callback", async (req, res) => {
  try {

    const { code } = req.query;

    if (!code) return res.send("No code ❌");

    const { tokens } = await oauth2Client.getToken(code);
    oauth2Client.setCredentials(tokens);

    const oauth2 = google.oauth2({
      auth: oauth2Client,
      version: "v2",
    });

    const userInfo = await oauth2.userinfo.get();
    const email = userInfo.data.email;

    console.log("Admin Login:", email);

    if (email !== process.env.ADMIN_EMAIL) {
      return res.send("Not Authorized ❌");
    }

    req.session.user = { role: "admin", email };

    res.redirect("/dashboard");

  } catch (err) {
    console.error("CALLBACK ERROR:", err);
    res.send("Google Login Failed ❌");
  }
});

// ================= USER LOGIN =================
app.post("/user-login", async (req, res) => {

  try {

    const { empId } = req.body;

    const data = await getExcelData();

    const matched = data.filter(r =>
      r.BH_ID === empId ||
      r.SM_ID === empId ||
      r.ZBM_ID === empId ||
      r.RBM_ID === empId ||
      r.ABM_ID === empId
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

// ================= EXCEL =================
async function getExcelData() {

  try {

    const res = await drive.files.export({
      fileId: process.env.EXCEL_FILE_ID,
      mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }, { responseType: "arraybuffer" });

    const wb = xlsx.read(res.data, { type: "buffer" });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    return xlsx.utils.sheet_to_json(sheet);

  } catch (e) {
    console.error("Excel Error:", e.message);
    return [];
  }
}

// ================= JSON =================
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
    r.BH_ID === empId ||
    r.SM_ID === empId ||
    r.ZBM_ID === empId ||
    r.RBM_ID === empId ||
    r.ABM_ID === empId
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
        name: `${state}_${type}`,
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

// ================= ROUTES =================
app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "public/login.html"));
});

app.get("/dashboard", isAuth, (req, res) => {
  res.sendFile(path.join(__dirname, "public/dashboard.html"));
});

// ================= START =================
app.listen(process.env.PORT || 10000, () => {
  console.log("🚀 FINAL SAFE SERVER RUNNING");
});