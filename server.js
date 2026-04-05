const express = require("express");
const cors = require("cors");
const { v4: uuidv4 } = require("uuid");
const Database = require("better-sqlite3");
const ExcelJS = require("exceljs");
const PDFDocument = require("pdfkit");
const path = require("path");

const app = express();
const PORT = 3000;
const ADMIN_PASSWORD = "admin1234"; // Change this

app.use(cors());
app.use(express.json());

// ── DB Setup ──────────────────────────────────────────────────────────────────
const db = new Database(path.join(__dirname, "attendance.db"));

db.exec(`
  CREATE TABLE IF NOT EXISTS sessions (
    id TEXT PRIMARY KEY,
    title TEXT NOT NULL,
    course_code TEXT NOT NULL,
    token TEXT UNIQUE NOT NULL,
    expires_at INTEGER NOT NULL,
    created_at INTEGER NOT NULL
  );

  CREATE TABLE IF NOT EXISTS submissions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    session_id TEXT NOT NULL,
    full_name TEXT NOT NULL,
    department TEXT NOT NULL,
    matric_number TEXT NOT NULL,
    submitted_at INTEGER NOT NULL,
    FOREIGN KEY (session_id) REFERENCES sessions(id),
    UNIQUE(session_id, matric_number)
  );
`);

// ── Auth ──────────────────────────────────────────────────────────────────────
const tokens = new Set();

function requireAuth(req, res, next) {
  const auth = req.headers["authorization"] || req.query.auth;
  if (!auth || !tokens.has(auth)) {
    return res.status(401).json({ error: "Unauthorized" });
  }
  next();
}

app.post("/api/login", (req, res) => {
  const { password } = req.body;
  if (password !== ADMIN_PASSWORD) {
    return res.status(401).json({ error: "Wrong password" });
  }
  const token = uuidv4();
  tokens.add(token);
  res.json({ token });
});

app.post("/api/logout", requireAuth, (req, res) => {
  tokens.delete(req.headers["authorization"]);
  res.json({ ok: true });
});

// ── Sessions ──────────────────────────────────────────────────────────────────
app.post("/api/sessions", requireAuth, (req, res) => {
  const { title, course_code, duration_minutes } = req.body;
  if (!title || !course_code || !duration_minutes) {
    return res.status(400).json({ error: "title, course_code and duration_minutes required" });
  }
  const id = uuidv4();
  const token = uuidv4().replace(/-/g, "").slice(0, 12);
  const now = Date.now();
  const expires_at = now + duration_minutes * 60 * 1000;

  db.prepare(
    "INSERT INTO sessions (id, title, course_code, token, expires_at, created_at) VALUES (?, ?, ?, ?, ?, ?)"
  ).run(id, title, course_code, token, expires_at, now);

  res.json({ id, token, expires_at });
});

app.get("/api/sessions", requireAuth, (req, res) => {
  const sessions = db.prepare("SELECT * FROM sessions ORDER BY created_at DESC").all();
  const result = sessions.map((s) => {
    const subs = db.prepare("SELECT * FROM submissions WHERE session_id = ?").all(s.id);
    const deptBreakdown = {};
    subs.forEach((sub) => {
      deptBreakdown[sub.department] = (deptBreakdown[sub.department] || 0) + 1;
    });
    return {
      ...s,
      total: subs.length,
      is_active: Date.now() < s.expires_at,
      department_breakdown: deptBreakdown,
    };
  });
  res.json(result);
});

app.get("/api/sessions/:id", requireAuth, (req, res) => {
  const session = db.prepare("SELECT * FROM sessions WHERE id = ?").get(req.params.id);
  if (!session) return res.status(404).json({ error: "Session not found" });

  const subs = db
    .prepare("SELECT * FROM submissions WHERE session_id = ? ORDER BY submitted_at ASC")
    .all(req.params.id);

  const deptBreakdown = {};
  subs.forEach((sub) => {
    deptBreakdown[sub.department] = (deptBreakdown[sub.department] || 0) + 1;
  });

  res.json({
    ...session,
    submissions: subs,
    total: subs.length,
    is_active: Date.now() < session.expires_at,
    department_breakdown: deptBreakdown,
  });
});

app.delete("/api/sessions/:id", requireAuth, (req, res) => {
  db.prepare("DELETE FROM submissions WHERE session_id = ?").run(req.params.id);
  db.prepare("DELETE FROM sessions WHERE id = ?").run(req.params.id);
  res.json({ ok: true });
});

// ── Attendance Form ───────────────────────────────────────────────────────────
app.get("/api/attend/:token", (req, res) => {
  const session = db.prepare("SELECT * FROM sessions WHERE token = ?").get(req.params.token);
  if (!session) return res.status(404).json({ error: "Invalid link" });

  const expired = Date.now() >= session.expires_at;
  res.json({
    id: session.id,
    title: session.title,
    course_code: session.course_code,
    expired,
    expires_at: session.expires_at,
  });
});

app.post("/api/attend/:token", (req, res) => {
  const session = db.prepare("SELECT * FROM sessions WHERE token = ?").get(req.params.token);
  if (!session) return res.status(404).json({ error: "Invalid link" });
  if (Date.now() >= session.expires_at) {
    return res.status(410).json({ error: "expired" });
  }

  const { full_name, department, matric_number } = req.body;
  if (!full_name || !department || !matric_number) {
    return res.status(400).json({ error: "All fields required" });
  }

  const existing = db
    .prepare("SELECT id FROM submissions WHERE session_id = ? AND matric_number = ?")
    .get(session.id, matric_number.trim().toUpperCase());

  if (existing) {
    return res.status(409).json({ error: "Matric number already submitted for this session" });
  }

  db.prepare(
    "INSERT INTO submissions (session_id, full_name, department, matric_number, submitted_at) VALUES (?, ?, ?, ?, ?)"
  ).run(session.id, full_name.trim(), department.trim(), matric_number.trim().toUpperCase(), Date.now());

  res.json({ ok: true, message: "Attendance recorded!" });
});

// ── Export ────────────────────────────────────────────────────────────────────
app.get("/api/sessions/:id/export/excel", requireAuth, async (req, res) => {
  const session = db.prepare("SELECT * FROM sessions WHERE id = ?").get(req.params.id);
  if (!session) return res.status(404).json({ error: "Not found" });

  const subs = db
    .prepare("SELECT * FROM submissions WHERE session_id = ? ORDER BY submitted_at ASC")
    .all(req.params.id);

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Attendance");

  // Header info
  sheet.mergeCells("A1:E1");
  sheet.getCell("A1").value = `${session.title} — ${session.course_code}`;
  sheet.getCell("A1").font = { bold: true, size: 14 };
  sheet.getCell("A1").alignment = { horizontal: "center" };

  sheet.mergeCells("A2:E2");
  sheet.getCell("A2").value = `Generated: ${new Date().toLocaleString()}  |  Total: ${subs.length}`;
  sheet.getCell("A2").alignment = { horizontal: "center" };
  sheet.getCell("A2").font = { color: { argb: "FF666666" }, size: 10 };

  sheet.addRow([]);

  // Column headers
  const headerRow = sheet.addRow(["#", "Full Name", "Department", "Matric Number", "Submitted At"]);
  headerRow.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1a1a2e" } };
    cell.alignment = { horizontal: "center" };
  });

  subs.forEach((sub, i) => {
    const row = sheet.addRow([
      i + 1,
      sub.full_name,
      sub.department,
      sub.matric_number,
      new Date(sub.submitted_at).toLocaleString(),
    ]);
    if (i % 2 === 0) {
      row.eachCell((cell) => {
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF5F5F5" } };
      });
    }
  });

  sheet.columns = [
    { width: 5 },
    { width: 30 },
    { width: 25 },
    { width: 18 },
    { width: 22 },
  ];

  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", `attachment; filename="${session.course_code}-attendance.xlsx"`);
  await workbook.xlsx.write(res);
  res.end();
});

app.get("/api/sessions/:id/export/pdf", requireAuth, async (req, res) => {
  const session = db.prepare("SELECT * FROM sessions WHERE id = ?").get(req.params.id);
  if (!session) return res.status(404).json({ error: "Not found" });

  const subs = db
    .prepare("SELECT * FROM submissions WHERE session_id = ? ORDER BY submitted_at ASC")
    .all(req.params.id);

  const doc = new PDFDocument({ margin: 40, size: "A4" });
  res.setHeader("Content-Type", "application/pdf");
  res.setHeader("Content-Disposition", `attachment; filename="${session.course_code}-attendance.pdf"`);
  doc.pipe(res);

  // Title block
  doc.rect(0, 0, doc.page.width, 80).fill("#1a1a2e");
  doc.fillColor("#ffffff").fontSize(18).font("Helvetica-Bold").text(session.title, 40, 20);
  doc.fontSize(11).font("Helvetica").text(`Course: ${session.course_code}  |  Total Present: ${subs.length}`, 40, 48);
  doc.moveDown(3);

  // Table header
  const tableTop = 100;
  const col = { num: 40, name: 70, dept: 240, matric: 370, time: 450 };

  doc.fillColor("#1a1a2e").rect(40, tableTop, doc.page.width - 80, 22).fill();
  doc.fillColor("#ffffff").fontSize(9).font("Helvetica-Bold");
  doc.text("#", col.num, tableTop + 6);
  doc.text("Full Name", col.name, tableTop + 6);
  doc.text("Department", col.dept, tableTop + 6);
  doc.text("Matric No.", col.matric, tableTop + 6);
  doc.text("Time", col.time, tableTop + 6);

  let y = tableTop + 26;
  doc.fontSize(8).font("Helvetica");

  subs.forEach((sub, i) => {
    if (y > doc.page.height - 60) {
      doc.addPage();
      y = 40;
    }
    if (i % 2 === 0) {
      doc.fillColor("#f5f5f5").rect(40, y - 4, doc.page.width - 80, 18).fill();
    }
    doc.fillColor("#222222");
    doc.text(String(i + 1), col.num, y);
    doc.text(sub.full_name.slice(0, 22), col.name, y);
    doc.text(sub.department.slice(0, 18), col.dept, y);
    doc.text(sub.matric_number, col.matric, y);
    doc.text(new Date(sub.submitted_at).toLocaleTimeString(), col.time, y);
    y += 18;
  });

  // Footer
  doc
    .fillColor("#999999")
    .fontSize(8)
    .text(`Generated ${new Date().toLocaleString()}`, 40, doc.page.height - 40);

  doc.end();
});

app.listen(PORT, () => console.log(`Backend running on http://localhost:${PORT}`));
