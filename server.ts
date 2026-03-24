import express from 'express';
import { createServer as createViteServer } from 'vite';
import path from 'path';
import cookieParser from 'cookie-parser';
import jwt from 'jsonwebtoken';
import bcrypt from 'bcryptjs';
import Database from 'better-sqlite3';
import * as XLSX from 'xlsx';
import { google } from 'googleapis';
import axios from 'axios';
import { GoogleGenAI } from '@google/genai';

const PORT = 3000;
const JWT_SECRET = process.env.JWT_SECRET || 'academic-kpi-secret-2024';

// Google OAuth Config
const GOOGLE_CLIENT_ID = process.env.GOOGLE_CLIENT_ID;
const GOOGLE_CLIENT_SECRET = process.env.GOOGLE_CLIENT_SECRET;
const oauth2Client = new google.auth.OAuth2(
  GOOGLE_CLIENT_ID,
  GOOGLE_CLIENT_SECRET,
  `${process.env.APP_URL}/api/auth/google/callback`
);

// Microsoft OAuth Config
const MS_CLIENT_ID = process.env.MS_CLIENT_ID;
const MS_CLIENT_SECRET = process.env.MS_CLIENT_SECRET;
const MS_REDIRECT_URI = `${process.env.APP_URL}/api/auth/ms/callback`;

// Initialize Database
const db = new Database('academic.db');
db.pragma('journal_mode = WAL');

// Create Tables
db.exec(`
  CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    email TEXT UNIQUE NOT NULL,
    password TEXT NOT NULL,
    role TEXT NOT NULL, -- ADMIN, MANAGER, LECTURER
    lecturerId INTEGER,
    name TEXT NOT NULL
  );

  CREATE TABLE IF NOT EXISTS lecturers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT UNIQUE NOT NULL,
    department TEXT NOT NULL
  );

  CREATE TABLE IF NOT EXISTS programmes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT UNIQUE NOT NULL
  );

  CREATE TABLE IF NOT EXISTS modules (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    programmeId INTEGER NOT NULL,
    UNIQUE(name, programmeId),
    FOREIGN KEY (programmeId) REFERENCES programmes(id)
  );

  CREATE TABLE IF NOT EXISTS kpi_records (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    lecturerId INTEGER NOT NULL,
    moduleId INTEGER NOT NULL,
    term TEXT NOT NULL,
    intake TEXT NOT NULL,
    level TEXT NOT NULL,
    deliveryGroup TEXT NOT NULL,
    course TEXT,
    totalEnrolled INTEGER NOT NULL,
    totalSubmissions INTEGER NOT NULL,
    nonSubmissions INTEGER NOT NULL,
    passes INTEGER NOT NULL,
    fails INTEGER NOT NULL,
    attendanceRate REAL NOT NULL,
    meqSatisfaction REAL NOT NULL,
    meqResponseRate REAL NOT NULL,
    FOREIGN KEY (lecturerId) REFERENCES lecturers(id),
    FOREIGN KEY (moduleId) REFERENCES modules(id)
  );

  CREATE TABLE IF NOT EXISTS kpi_weights (
    id INTEGER PRIMARY KEY CHECK (id = 1),
    submissionWeight REAL DEFAULT 0.25,
    passWeight REAL DEFAULT 0.25,
    attendanceWeight REAL DEFAULT 0.25,
    meqWeight REAL DEFAULT 0.25
  );

  -- Seed initial weights if not exists
  INSERT OR IGNORE INTO kpi_weights (id, submissionWeight, passWeight, attendanceWeight, meqWeight)
  VALUES (1, 0.25, 0.25, 0.25, 0.25);
`);

// Migration for kpi_records table to add missing columns if they don't exist
const tableInfo = db.prepare("PRAGMA table_info(kpi_records)").all() as any[];
const columns = tableInfo.map(c => c.name);

const migrations = [
  { name: 'level', type: 'TEXT NOT NULL DEFAULT "4"' },
  { name: 'deliveryGroup', type: 'TEXT NOT NULL DEFAULT "Weekday"' },
  { name: 'course', type: 'TEXT' },
  { name: 'nonSubmissions', type: 'INTEGER NOT NULL DEFAULT 0' },
  { name: 'meqResponseRate', type: 'REAL NOT NULL DEFAULT 0' }
];

for (const m of migrations) {
  if (!columns.includes(m.name)) {
    try {
      db.exec(`ALTER TABLE kpi_records ADD COLUMN ${m.name} ${m.type}`);
      console.log(`Added column ${m.name} to kpi_records`);
    } catch (e) {
      console.error(`Failed to add column ${m.name}:`, e);
    }
  }
}

// Seed Admin User if not exists
const adminExists = db.prepare('SELECT * FROM users WHERE email = ?').get('admin@institution.edu');
if (!adminExists) {
  const hashedPassword = bcrypt.hashSync('admin123', 10);
  db.prepare('INSERT INTO users (email, password, role, name) VALUES (?, ?, ?, ?)').run(
    'admin@institution.edu',
    hashedPassword,
    'ADMIN',
    'System Administrator'
  );
}

async function startServer() {
  const app = express();
  app.use(express.json({ limit: '50mb' }));
  app.use(cookieParser());

  // Auth Middleware
  const authenticate = (req: any, res: any, next: any) => {
    const authHeader = req.headers.authorization;
    const token = (authHeader && authHeader.split(' ')[1]) || req.cookies.token;
    if (!token) return res.status(401).json({ error: 'Unauthorized' });
    try {
      const decoded = jwt.verify(token, JWT_SECRET);
      (req as any).user = decoded;
      next();
    } catch (err) {
      res.status(401).json({ error: 'Invalid token' });
    }
  };

  // API Routes
  app.post('/api/auth/login', (req, res) => {
    const { email, password } = req.body;
    const user = db.prepare('SELECT * FROM users WHERE email = ?').get(email) as any;
    if (!user || !bcrypt.compareSync(password, user.password)) {
      return res.status(401).json({ error: 'Invalid credentials' });
    }
    const token = jwt.sign({ id: user.id, email: user.email, role: user.role, name: user.name, lecturerId: user.lecturerId }, JWT_SECRET, { expiresIn: '24h' });
    res.cookie('token', token, { httpOnly: true, secure: true, sameSite: 'none' });
    res.json({ token, user: { id: user.id, email: user.email, role: user.role, name: user.name, lecturerId: user.lecturerId } });
  });

  app.post('/api/auth/logout', (req, res) => {
    res.clearCookie('token');
    res.json({ success: true });
  });

  app.get('/api/auth/me', authenticate, (req: any, res) => {
    res.json({ user: req.user });
  });

  const KPI_ALIASES: Record<string, string[]> = {
    programme: ['programme', 'program', 'course', 'programme name', 'dept', 'department', 'school', 'faculty', 'group', 'class'],
    module: ['module', 'subject', 'module name', 'unit', 'code', 'module code', 'row labels', 'item', 'title'],
    lecturer: ['lecturer', 'teacher', 'instructor', 'lecturer name', 'staff', 'academic', 'tutor', 'professor', 'name', 'group', 'assigned to'],
    term: ['term', 'semester', 'period', 'academic year', 'year', 'session', 'trimester'],
    intake: ['intake', 'batch', 'cohort', 'month', 'start date', 'group', 'period'],
    level: ['level', 'lvl', 'stage', 'year of study', 'level of study'],
    group: ['group', 'delivery mode', 'mode', 'session', 'class', 'delivery group'],
    course: ['course', 'programme', 'degree', 'course name'],
    totalenrolled: ['totalenrolled', 'enrolled', 'students', 'total students', 'total enrolled', 'count', 'size', 'enrolment', 'no. of students'],
    totalsubmissions: ['totalsubmissions', 'submissions', 'total submissions', 'total submission', 'submitted', 'submission count', 'average of submission %', 'total'],
    nonsubmissions: ['nonsubmissions', 'non-submissions', 'not submitted', 'absent', 'non-submission count'],
    passes: ['passes', 'pass', 'passed', 'success', 'passed count', 'achieved'],
    fails: ['fails', 'fail', 'failed', 'failure', 'failed count', 'not achieved'],
    attendancerate: ['attendancerate', 'attendance', 'avg attendance', 'attendance rate', 'presence', 'attendance %', 'at %'],
    meqsatisfaction: ['meqsatisfaction', 'meq', 'satisfaction', 'student satisfaction', 'meq satisfaction', 'feedback', 'rating', 'score', 'satisfaction %'],
    meqresponserate: ['meq response rate', 'response rate', 'meq %', 'feedback rate', 'meq response %']
  };

  const mapRowToKPI = (rawRow: any) => {
    const row: any = {};
    const required = Object.keys(KPI_ALIASES);
    
    required.forEach(targetCol => {
      const aliases = KPI_ALIASES[targetCol];
      const foundKey = Object.keys(rawRow).find(actualKey => 
        aliases.includes(actualKey.trim().toLowerCase()) || actualKey.trim().toLowerCase() === targetCol
      );
      if (foundKey) {
        row[targetCol] = rawRow[foundKey];
      }
    });

    const deptKey = Object.keys(rawRow).find(k => k.trim().toLowerCase() === 'department' || k.trim().toLowerCase() === 'dept');
    if (deptKey) row.department = rawRow[deptKey];

    return row;
  };

  const validateKPIRows = (rows: any[]) => {
    if (!Array.isArray(rows) || rows.length === 0) {
      throw new Error('No data provided or invalid format.');
    }

    // Find the first row that has some data
    const firstDataRow = rows.find(r => Object.keys(r).length > 0);
    if (!firstDataRow) throw new Error('No data found in file');

    const mapped = mapRowToKPI(firstDataRow);
    // Only require lecturer and module as a bare minimum
    const essential = ['lecturer', 'module'];
    const missing = essential.filter(k => mapped[k] === undefined || mapped[k] === null);
    
    if (missing.length > 0) {
      const found = Object.keys(firstDataRow).join(', ');
      throw new Error(`Missing essential columns: ${missing.join(', ')}. Please ensure your file has columns for Lecturer and Module. Found columns: ${found}`);
    }
    return true;
  };

  // Google OAuth Routes
  app.get('/api/auth/google/url', authenticate, (req, res) => {
    if (!GOOGLE_CLIENT_ID || !GOOGLE_CLIENT_SECRET) {
      return res.status(400).json({ error: 'Google OAuth is not configured. Please set GOOGLE_CLIENT_ID and GOOGLE_CLIENT_SECRET.' });
    }
    const url = oauth2Client.generateAuthUrl({
      access_type: 'offline',
      scope: ['https://www.googleapis.com/auth/drive.readonly'],
      prompt: 'consent'
    });
    res.json({ url });
  });

  app.get('/api/auth/google/callback', async (req, res) => {
    const { code } = req.query;
    try {
      const { tokens } = await oauth2Client.getToken(code as string);
      // In a real app, you'd store this in the database for the specific user
      // For this demo, we'll send it back to the client to store in local storage (not secure for production)
      res.send(`
        <html>
          <body>
            <script>
              window.opener.postMessage({ type: 'GOOGLE_AUTH_SUCCESS', tokens: ${JSON.stringify(tokens)} }, '*');
              window.close();
            </script>
          </body>
        </html>
      `);
    } catch (err) {
      res.status(500).send('Authentication failed');
    }
  });

  app.get('/api/drive/files', authenticate, async (req: any, res) => {
    try {
      const tokensStr = req.headers['x-google-tokens'] as string;
      if (!tokensStr) return res.status(400).json({ error: 'Missing tokens' });
      const tokens = JSON.parse(tokensStr);
      if (!tokens) return res.status(400).json({ error: 'Invalid tokens' });
      oauth2Client.setCredentials(tokens);
      
      const drive = google.drive({ version: 'v3', auth: oauth2Client });
      const response = await drive.files.list({
        q: "mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' or mimeType='application/vnd.ms-excel' or mimeType='text/csv' or mimeType='application/json' or name contains '.xlsx' or name contains '.xls' or name contains '.csv' or name contains '.json' or name contains '.ods'",
        fields: 'files(id, name, modifiedTime)',
      });
      res.json(response.data.files);
    } catch (err) {
      console.error(err);
      res.status(500).json({ error: 'Failed to fetch files' });
    }
  });

  app.post('/api/drive/import', authenticate, async (req: any, res) => {
    try {
      const { fileId, tokens: tokenStr } = req.body;
      if (!tokenStr) return res.status(400).json({ error: 'Missing tokens' });
      const tokens = JSON.parse(tokenStr);
      if (!tokens) return res.status(400).json({ error: 'Invalid tokens' });
      oauth2Client.setCredentials(tokens);
      const drive = google.drive({ version: 'v3', auth: oauth2Client });

      const response = await drive.files.get(
        { fileId, alt: 'media' },
        { responseType: 'arraybuffer' }
      );
      
      let json: any[];
      const fileMeta = await drive.files.get({ fileId, fields: 'name' });
      const fileName = fileMeta.data.name?.toLowerCase() || '';

      if (fileName.endsWith('.json')) {
        const text = Buffer.from(response.data as any).toString('utf-8');
        json = JSON.parse(text);
      } else {
        const workbook = XLSX.read(response.data as any, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        json = XLSX.utils.sheet_to_json(worksheet);
      }

      validateKPIRows(json);

      // Reuse the import logic
      const insertProgramme = db.prepare('INSERT OR IGNORE INTO programmes (name) VALUES (?)');
      const getProgramme = db.prepare('SELECT id FROM programmes WHERE name = ?');
      const insertModule = db.prepare('INSERT OR IGNORE INTO modules (name, programmeId) VALUES (?, ?)');
      const getModule = db.prepare('SELECT id FROM modules WHERE name = ? AND programmeId = ?');
      const insertLecturer = db.prepare('INSERT OR IGNORE INTO lecturers (name, department) VALUES (?, ?)');
      const getLecturer = db.prepare('SELECT id FROM lecturers WHERE name = ?');
      const deleteExistingKPI = db.prepare(`
        DELETE FROM kpi_records 
        WHERE lecturerId = ? AND moduleId = ? AND term = ? AND intake = ? AND level = ? AND deliveryGroup = ?
      `);
      const insertKPI = db.prepare(`
        INSERT INTO kpi_records (
          lecturerId, moduleId, term, intake, level, deliveryGroup, course,
          totalEnrolled, totalSubmissions, nonSubmissions, passes, fails, 
          attendanceRate, meqSatisfaction, meqResponseRate
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `);

      const transaction = db.transaction((rows: any[]) => {
        for (const rawRow of rows) {
          const row = mapRowToKPI(rawRow);
          
          // Skip rows that don't have the bare essentials
          if (!row.lecturer || !row.module) continue;

          // Provide defaults for missing data
          const programmeName = row.programme || 'General';
          const term = row.term || 'Unknown';
          const intake = row.intake || 'Unknown';
          const level = String(row.level || '4');
          const deliveryGroup = row.group || 'Weekday';
          const course = row.course || programmeName;
          
          const totalSubmissions = Number(row.totalsubmissions) || 0;
          const passes = Number(row.passes) || 0;
          const fails = row.fails !== undefined ? Number(row.fails) : Math.max(0, totalSubmissions - passes);
          const nonSubmissions = row.nonsubmissions !== undefined ? Number(row.nonsubmissions) : Math.max(0, (Number(row.totalenrolled) || 0) - totalSubmissions);
          const totalEnrolled = Number(row.totalenrolled) || (totalSubmissions + nonSubmissions);
          
          const attendanceRate = Number(row.attendancerate) || 0;
          const meqSatisfaction = Number(row.meqsatisfaction) || 0;
          const meqResponseRate = Number(row.meqresponserate) || 0;

          insertProgramme.run(programmeName);
          const pResult = getProgramme.get(programmeName) as any;
          if (!pResult) continue;
          const pId = pResult.id;

          insertModule.run(row.module, pId);
          const mResult = getModule.get(row.module, pId) as any;
          if (!mResult) continue;
          const mId = mResult.id;

          insertLecturer.run(row.lecturer, row.department || 'General');
          const lResult = getLecturer.get(row.lecturer) as any;
          if (!lResult) continue;
          const lId = lResult.id;

          deleteExistingKPI.run(lId, mId, term, intake, level, deliveryGroup);
          insertKPI.run(
            lId, mId, term, intake, level, deliveryGroup, course,
            totalEnrolled, totalSubmissions, nonSubmissions, passes, fails, 
            attendanceRate, meqSatisfaction, meqResponseRate
          );
        }
      });

      transaction(json);
      res.json({ success: true, count: json.length });
    } catch (err) {
      console.error(err);
      res.status(500).json({ error: 'Failed to import from Drive' });
    }
  });

  // Microsoft OneDrive Routes
  app.get('/api/auth/ms/url', authenticate, (req, res) => {
    if (!MS_CLIENT_ID || !MS_CLIENT_SECRET) {
      return res.status(400).json({ error: 'Microsoft OAuth is not configured. Please set MS_CLIENT_ID and MS_CLIENT_SECRET.' });
    }
    const url = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=${MS_CLIENT_ID}&response_type=code&redirect_uri=${encodeURIComponent(MS_REDIRECT_URI)}&response_mode=query&scope=Files.Read.All offline_access`;
    res.json({ url });
  });

  app.post('/api/auth/ms/refresh', authenticate, async (req, res) => {
    const { refresh_token } = req.body;
    try {
      const response = await axios.post('https://login.microsoftonline.com/common/oauth2/v2.0/token', 
        new URLSearchParams({
          client_id: MS_CLIENT_ID!,
          client_secret: MS_CLIENT_SECRET!,
          refresh_token: refresh_token,
          grant_type: 'refresh_token',
          scope: 'Files.Read.All offline_access'
        }).toString(),
        { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
      );
      res.json(response.data);
    } catch (err) {
      res.status(500).json({ error: 'Failed to refresh Microsoft token' });
    }
  });

  app.get('/api/auth/ms/callback', async (req, res) => {
    const { code } = req.query;
    try {
      const response = await axios.post('https://login.microsoftonline.com/common/oauth2/v2.0/token', 
        new URLSearchParams({
          client_id: MS_CLIENT_ID!,
          client_secret: MS_CLIENT_SECRET!,
          code: code as string,
          redirect_uri: MS_REDIRECT_URI,
          grant_type: 'authorization_code'
        }).toString(),
        { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
      );
      
      res.send(`
        <html>
          <body>
            <script>
              window.opener.postMessage({ type: 'MS_AUTH_SUCCESS', tokens: ${JSON.stringify(response.data)} }, '*');
              window.close();
            </script>
          </body>
        </html>
      `);
    } catch (err) {
      res.status(500).send('Microsoft Authentication failed');
    }
  });

  app.get('/api/onedrive/files', authenticate, async (req: any, res) => {
    try {
      const tokensStr = req.headers['x-ms-tokens'] as string;
      if (!tokensStr) return res.status(400).json({ error: 'Missing tokens' });
      const tokens = JSON.parse(tokensStr);
      if (!tokens || !tokens.access_token) return res.status(400).json({ error: 'Invalid tokens' });

      // Search for any data files
      const response = await axios.get("https://graph.microsoft.com/v1.0/me/drive/root/search(q='.')?$expand=thumbnails", {
        headers: { Authorization: `Bearer ${tokens.access_token}` }
      });
      
      // Graph API search is a bit limited in one go for multiple extensions easily without complex filters
      // We'll just return what we found. Users can also search for .csv manually if we added a search box.
      res.json(response.data.value);
    } catch (err: any) {
      if (err.response?.status === 401) {
        return res.status(401).json({ error: 'Token expired', code: 'TOKEN_EXPIRED' });
      }
      console.error(err);
      res.status(500).json({ error: 'Failed to fetch OneDrive files' });
    }
  });

  app.post('/api/onedrive/import', authenticate, async (req: any, res) => {
    try {
      const { fileId, tokens: tokenStr } = req.body;
      if (!tokenStr) return res.status(400).json({ error: 'Missing tokens' });
      const tokens = JSON.parse(tokenStr);
      if (!tokens || !tokens.access_token) return res.status(400).json({ error: 'Invalid tokens' });

      const response = await axios.get(`https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/content`, {
        headers: { Authorization: `Bearer ${tokens.access_token}` },
        responseType: 'arraybuffer'
      });
      
      let json: any[];
      const itemMeta = await axios.get(`https://graph.microsoft.com/v1.0/me/drive/items/${fileId}`, {
        headers: { Authorization: `Bearer ${tokens.access_token}` }
      });
      const fileName = itemMeta.data.name?.toLowerCase() || '';

      if (fileName.endsWith('.json')) {
        const text = Buffer.from(response.data as any).toString('utf-8');
        json = JSON.parse(text);
      } else {
        const workbook = XLSX.read(response.data as any, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        json = XLSX.utils.sheet_to_json(worksheet);
      }

      validateKPIRows(json);

      const insertProgramme = db.prepare('INSERT OR IGNORE INTO programmes (name) VALUES (?)');
      const getProgramme = db.prepare('SELECT id FROM programmes WHERE name = ?');
      const insertModule = db.prepare('INSERT OR IGNORE INTO modules (name, programmeId) VALUES (?, ?)');
      const getModule = db.prepare('SELECT id FROM modules WHERE name = ? AND programmeId = ?');
      const insertLecturer = db.prepare('INSERT OR IGNORE INTO lecturers (name, department) VALUES (?, ?)');
      const getLecturer = db.prepare('SELECT id FROM lecturers WHERE name = ?');
      const deleteExistingKPI = db.prepare(`
        DELETE FROM kpi_records 
        WHERE lecturerId = ? AND moduleId = ? AND term = ? AND intake = ? AND level = ? AND deliveryGroup = ?
      `);
      const insertKPI = db.prepare(`
        INSERT INTO kpi_records (
          lecturerId, moduleId, term, intake, level, deliveryGroup, course,
          totalEnrolled, totalSubmissions, nonSubmissions, passes, fails, 
          attendanceRate, meqSatisfaction, meqResponseRate
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `);

      const transaction = db.transaction((rows: any[]) => {
        for (const rawRow of rows) {
          const row = mapRowToKPI(rawRow);
          
          // Skip rows that don't have the bare essentials
          if (!row.lecturer || !row.module) continue;

          // Provide defaults for missing data
          const programmeName = row.programme || 'General';
          const term = row.term || 'Unknown';
          const intake = row.intake || 'Unknown';
          const level = String(row.level || '4');
          const deliveryGroup = row.group || 'Weekday';
          const course = row.course || programmeName;
          
          const totalSubmissions = Number(row.totalsubmissions) || 0;
          const passes = Number(row.passes) || 0;
          const fails = row.fails !== undefined ? Number(row.fails) : Math.max(0, totalSubmissions - passes);
          const nonSubmissions = row.nonsubmissions !== undefined ? Number(row.nonsubmissions) : Math.max(0, (Number(row.totalenrolled) || 0) - totalSubmissions);
          const totalEnrolled = Number(row.totalenrolled) || (totalSubmissions + nonSubmissions);
          
          const attendanceRate = Number(row.attendancerate) || 0;
          const meqSatisfaction = Number(row.meqsatisfaction) || 0;
          const meqResponseRate = Number(row.meqresponserate) || 0;

          insertProgramme.run(programmeName);
          const pResult = getProgramme.get(programmeName) as any;
          if (!pResult) continue;
          const pId = pResult.id;

          insertModule.run(row.module, pId);
          const mResult = getModule.get(row.module, pId) as any;
          if (!mResult) continue;
          const mId = mResult.id;

          insertLecturer.run(row.lecturer, row.department || 'General');
          const lResult = getLecturer.get(row.lecturer) as any;
          if (!lResult) continue;
          const lId = lResult.id;

          deleteExistingKPI.run(lId, mId, term, intake, level, deliveryGroup);
          insertKPI.run(
            lId, mId, term, intake, level, deliveryGroup, course,
            totalEnrolled, totalSubmissions, nonSubmissions, passes, fails, 
            attendanceRate, meqSatisfaction, meqResponseRate
          );
        }
      });

      transaction(json);
      res.json({ success: true, count: json.length });
    } catch (err) {
      console.error(err);
      res.status(500).json({ error: 'Failed to import from OneDrive' });
    }
  });

  app.get('/api/kpi/weights', authenticate, (req, res) => {
    const weights = db.prepare('SELECT * FROM kpi_weights WHERE id = 1').get();
    res.json(weights);
  });

  app.put('/api/kpi/weights', authenticate, (req: any, res) => {
    if (req.user.role !== 'ADMIN') return res.status(403).json({ error: 'Forbidden' });
    const { submissionWeight, passWeight, attendanceWeight, meqWeight } = req.body;
    db.prepare(`
      UPDATE kpi_weights 
      SET submissionWeight = ?, passWeight = ?, attendanceWeight = ?, meqWeight = ?
      WHERE id = 1
    `).run(submissionWeight, passWeight, attendanceWeight, meqWeight);
    res.json({ success: true });
  });

  app.get('/api/kpi/institution', authenticate, (req, res) => {
    const records = db.prepare(`
      SELECT 
        k.id, k.lecturerId, k.moduleId, k.term, k.intake, k.level, k.deliveryGroup as "group", k.course,
        k.totalEnrolled, k.totalSubmissions, k.nonSubmissions, k.passes, k.fails,
        k.attendanceRate, k.meqSatisfaction, k.meqResponseRate,
        l.name as lecturerName, l.department,
        m.name as moduleName,
        p.name as programmeName
      FROM kpi_records k
      JOIN lecturers l ON k.lecturerId = l.id
      JOIN modules m ON k.moduleId = m.id
      JOIN programmes p ON m.programmeId = p.id
    `).all();
    res.json(records);
  });

  app.get('/api/kpi/raw-data', authenticate, (req, res) => {
    try {
      const records = db.prepare(`
        SELECT 
          k.*, 
          l.name as lecturerName, 
          m.name as moduleName,
          p.name as programmeName
        FROM kpi_records k
        JOIN lecturers l ON k.lecturerId = l.id
        JOIN modules m ON k.moduleId = m.id
        JOIN programmes p ON m.programmeId = p.id
      `).all();
      res.json(records);
    } catch (err) {
      res.status(500).json({ error: 'Failed to fetch raw data' });
    }
  });

  app.get('/api/lecturers', authenticate, (req, res) => {
    const lecturers = db.prepare('SELECT * FROM lecturers').all();
    res.json(lecturers);
  });

  app.post('/api/upload/performance', authenticate, (req: any, res) => {
    if (req.user.role === 'LECTURER') return res.status(403).json({ error: 'Forbidden' });
    const { data } = req.body; // Expecting array of objects from Excel
    
    try {
      validateKPIRows(data);
    } catch (err: any) {
      return res.status(400).json({ error: err.message });
    }

    const insertProgramme = db.prepare('INSERT OR IGNORE INTO programmes (name) VALUES (?)');
    const getProgramme = db.prepare('SELECT id FROM programmes WHERE name = ?');
    const insertModule = db.prepare('INSERT OR IGNORE INTO modules (name, programmeId) VALUES (?, ?)');
    const getModule = db.prepare('SELECT id FROM modules WHERE name = ? AND programmeId = ?');
    const insertLecturer = db.prepare('INSERT OR IGNORE INTO lecturers (name, department) VALUES (?, ?)');
    const getLecturer = db.prepare('SELECT id FROM lecturers WHERE name = ?');
    const deleteExistingKPI = db.prepare(`
      DELETE FROM kpi_records 
      WHERE lecturerId = ? AND moduleId = ? AND term = ? AND intake = ? AND level = ? AND deliveryGroup = ?
    `);
    const insertKPI = db.prepare(`
      INSERT INTO kpi_records (
        lecturerId, moduleId, term, intake, level, deliveryGroup, course,
        totalEnrolled, totalSubmissions, nonSubmissions, passes, fails, 
        attendanceRate, meqSatisfaction, meqResponseRate
      )
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `);

    const transaction = db.transaction((rows: any[]) => {
      for (const rawRow of rows) {
        const row = mapRowToKPI(rawRow);
        
        // Skip rows that don't have the bare essentials
        if (!row.lecturer || !row.module) continue;

        // Provide defaults for missing data
        const programmeName = row.programme || 'General';
        const term = row.term || 'Unknown';
        const intake = row.intake || 'Unknown';
        const level = String(row.level || '4');
        const deliveryGroup = row.group || 'Weekday';
        const course = row.course || programmeName;
        
        const totalSubmissions = Number(row.totalsubmissions) || 0;
        const passes = Number(row.passes) || 0;
        const fails = row.fails !== undefined ? Number(row.fails) : Math.max(0, totalSubmissions - passes);
        const nonSubmissions = row.nonsubmissions !== undefined ? Number(row.nonsubmissions) : Math.max(0, (Number(row.totalenrolled) || 0) - totalSubmissions);
        const totalEnrolled = Number(row.totalenrolled) || (totalSubmissions + nonSubmissions);
        
        const attendanceRate = Number(row.attendancerate) || 0;
        const meqSatisfaction = Number(row.meqsatisfaction) || 0;
        const meqResponseRate = Number(row.meqresponserate) || 0;

        insertProgramme.run(programmeName);
        const pResult = getProgramme.get(programmeName) as any;
        if (!pResult) continue;
        const pId = pResult.id;

        insertModule.run(row.module, pId);
        const mResult = getModule.get(row.module, pId) as any;
        if (!mResult) continue;
        const mId = mResult.id;

        insertLecturer.run(row.lecturer, row.department || 'General');
        const lResult = getLecturer.get(row.lecturer) as any;
        if (!lResult) continue;
        const lId = lResult.id;

        deleteExistingKPI.run(lId, mId, term, intake, level, deliveryGroup);
        insertKPI.run(
          lId, mId, term, intake, level, deliveryGroup, course,
          totalEnrolled, totalSubmissions, nonSubmissions, passes, fails, 
          attendanceRate, meqSatisfaction, meqResponseRate
        );
      }
    });

    try {
      transaction(data);
      res.json({ success: true, count: data.length });
    } catch (err) {
      console.error(err);
      res.status(500).json({ error: 'Failed to import data' });
    }
  });

  // Vite middleware for development
  if (process.env.NODE_ENV !== 'production') {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    app.use(express.static('dist'));
    app.get('*', (req, res) => res.sendFile(path.resolve('dist/index.html')));
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
