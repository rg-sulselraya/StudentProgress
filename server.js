const http = require('http');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
let google = null;
try { ({ google } = require('googleapis')); } catch (_) { /* npm install required before Google Sheets access */ }
try { const envPath = fs.existsSync(path.join(process.cwd(), '.env')) ? path.join(process.cwd(), '.env') : path.join(__dirname, '.env'); require('dotenv').config({ path: envPath }); } catch (_) { /* environment variables may be supplied by the process */ }

const PORT = Number(process.env.PORT || 4173);
const HOST = process.env.HOST || '127.0.0.1';
const SPREADSHEET_ID = '1FcnkMsR24hU6A-bF5BYil2AX5iq-Ib38-Ta0dN_luBM';
const SHEET_NAME = process.env.GOOGLE_SHEETS_STUDENT_SHEET_NAME || 'siswa';
const CREDENTIALS = process.env.GOOGLE_APPLICATION_CREDENTIALS;
const APPS_SCRIPT_URL = process.env.GOOGLE_APPS_SCRIPT_URL;
const DATA_DIR = path.join(__dirname, 'data');
const SUBMISSIONS_FILE = path.join(DATA_DIR, 'submissions.json');

function apiError(message, details) { const e = new Error(message); e.details = details; return e; }
function normalize(value) { return String(value || '').trim().toLowerCase(); }
function headerIndex(headers, names) {
  return headers.findIndex(header => names.some(name => normalize(header).replace(/[^a-z0-9]/g, '').includes(name)));
}
// rowCount/totalRows describe raw non-empty rows read from the sheet.
// studentsLoaded describes valid rows after User Serial/name/class validation.
let lastSheetMeta = { headers: [], rowCount: 0, totalRows: 0, rowsRead: 0, rangeRowsRead: 0, studentsLoaded: 0, classCounts: {}, classNames: [], excludedRows: 0, warnings: { emptyUserSerial: 0, duplicateUserSerial: 0, emptyName: 0, emptyClass: 0, invalidRows: [] } };
function normalizeClassName(value) { return String(value || '').trim(); }
function normalizeStudents(students, totalRows) {
  const classCounts = {};
  const serialCounts = {};
  (students || []).forEach(student => { const code = String(student.studentCode || student.userSerial || '').trim(); if (code) serialCounts[code] = (serialCounts[code] || 0) + 1; });
  const warnings = {emptyUserSerial:0,duplicateUserSerial:0,emptyName:0,emptyClass:0,invalidRows:[]};
  const normalized = (students || []).map(student => {
    const name = String(student.name || '').trim();
    const className = normalizeClassName(student.className);
    const studentCode = String(student.studentCode || student.userSerial || '').trim();
    if (!studentCode) warnings.emptyUserSerial++;
    if (studentCode && serialCounts[studentCode] > 1) warnings.duplicateUserSerial++;
    if (!name) warnings.emptyName++;
    if (!className) warnings.emptyClass++;
    const invalid = !studentCode || serialCounts[studentCode] > 1 || !name || !className;
    if (invalid) { warnings.invalidRows.push(student.rowNumber || null); return null; }
    const source = { ...student, studentCode, name, studentName: name, className, schoolName: String(student.schoolName || '').trim(), grade: String(student.grade || '').trim(), active: true };
    delete source.email; delete source.paymentDate; delete source.days; delete source.studyDays;
    classCounts[className] = (classCounts[className] || 0) + 1;
    return source;
  }).filter(Boolean);
  lastSheetMeta.classCounts = classCounts;
  lastSheetMeta.classNames = Object.keys(classCounts).sort((a, b) => a.localeCompare(b));
  lastSheetMeta.excludedRows = Math.max(0, (totalRows ?? students?.length ?? 0) - normalized.length);
  lastSheetMeta.warnings = warnings;
  return normalized;
}

async function readStudents() {
  console.log('[STUDENT DATA] Connecting to Google Sheets...');
  console.log('[STUDENT DATA] Spreadsheet ID:', SPREADSHEET_ID);
  console.log('[STUDENT DATA] Sheet:', SHEET_NAME);
  if (!APPS_SCRIPT_URL) throw apiError('Google Apps Script URL belum dikonfigurasi.', 'Isi GOOGLE_APPS_SCRIPT_URL pada .env dengan URL Web App Apps Script.');
  if (APPS_SCRIPT_URL) {
    console.log('[STUDENT DATA] Provider: Google Apps Script');
    try {
      const endpoint = APPS_SCRIPT_URL + (APPS_SCRIPT_URL.includes('?') ? '&' : '?') + 'action=students';
      let response;
      let lastError;
      for (let attempt = 1; attempt <= 2; attempt += 1) {
        const controller = new AbortController();
        const timeout = setTimeout(() => controller.abort(), 30_000);
        try {
          console.log(`[STUDENT DATA] Apps Script request attempt ${attempt}/2...`);
          response = await fetch(endpoint, { redirect: 'follow', signal: controller.signal, headers: { 'Accept': 'application/json,text/plain,*/*', 'User-Agent': 'StudentProgressBackend/1.0' } });
          if (response.ok) break;
          lastError = new Error(`Apps Script HTTP ${response.status}`);
        } catch (err) { lastError = err; }
        finally { clearTimeout(timeout); }
        if (attempt < 2) await new Promise(resolve => setTimeout(resolve, 1000));
      }
      if (!response) throw lastError || new Error('Apps Script tidak merespons.');
      const rawBody = await response.text();
      let payload;
      try { payload = JSON.parse(rawBody); }
      catch (_) {
        // Some Apps Script deployments prepend a short HTML wrapper before
        // the JSON body. Recover the JSON object instead of treating it as a
        // successful HTML page with no student data.
        const starts = [rawBody.indexOf('{"success"'), rawBody.indexOf('{"connected"'), rawBody.indexOf('{"students"')].filter(i => i >= 0);
        const start = starts.length ? Math.min(...starts) : -1;
        const end = rawBody.lastIndexOf('}');
        if (start < 0 || end <= start) throw new Error('Respons Apps Script bukan JSON yang valid.');
        payload = JSON.parse(rawBody.slice(start, end + 1));
      }
      const upstreamStudents = payload.students ?? payload.data;
      if (!response.ok || payload.success !== true || payload.connected === false || !Array.isArray(upstreamStudents)) throw apiError(payload.message || payload.error || 'Apps Script tidak mengembalikan data siswa.', payload.details || 'Periksa deployment Web App Apps Script versi terbaru.');
      // Apps Script returns normalized rows from the latest five-column schema.
      const totalRows = payload.rowCount ?? upstreamStudents.length;
      const students = normalizeStudents(upstreamStudents, totalRows);
      lastSheetMeta = { headers: payload.headers || [], rowCount: totalRows, totalRows, rowsRead: totalRows, rangeRowsRead: payload.rangeRowsRead ?? payload.rowsRead ?? totalRows, studentsLoaded: students.length, skippedRows: payload.skippedRows || [], classCounts: lastSheetMeta.classCounts, classNames: lastSheetMeta.classNames, warnings: payload.warnings || lastSheetMeta.warnings, excludedRows: lastSheetMeta.excludedRows };
      console.log('[STUDENT DATA] Headers:', (lastSheetMeta.headers || []).join(' | '));
      console.log('[STUDENT DATA] Rows found:', lastSheetMeta.totalRows);
      console.log('[STUDENT DATA] Classes included:', JSON.stringify(lastSheetMeta.classCounts));
      console.log('[STUDENT DATA] Validation warnings:', JSON.stringify(lastSheetMeta.warnings));
      console.log('[STUDENT DATA] Excluded:', lastSheetMeta.excludedRows);
      console.log('[STUDENT DATA] Students loaded:', students.length);
      return students;
    } catch (err) {
      const detail = err?.name === 'AbortError' ? 'Apps Script tidak merespons dalam 30 detik.' : (err.details || err.message);
      console.error('[STUDENT DATA ERROR]', err.message, detail || '');
      throw apiError('Data siswa belum dapat dimuat dari Google Apps Script.', detail);
    }
  }
  if (!google) throw apiError('Google Sheets API dependency belum terpasang.', 'Jalankan npm install dari folder project.');
  if (!CREDENTIALS) throw apiError('Google credentials belum dikonfigurasi di backend.', 'Isi GOOGLE_APPLICATION_CREDENTIALS pada .env backend.');
  if (!fs.existsSync(CREDENTIALS)) throw apiError('Google credentials tidak ditemukan.', `File credential tidak ditemukan: ${CREDENTIALS}`);
  let auth;
  try { auth = new google.auth.GoogleAuth({ keyFile: CREDENTIALS, scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'] }); }
  catch (err) { throw apiError('Google Sheets API belum dapat diinisialisasi.', err.message); }
  const sheets = google.sheets({ version: 'v4', auth });
  let values;
  try {
    const response = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: `${SHEET_NAME}!A:E`, majorDimension: 'ROWS' });
    values = response.data.values || [];
  } catch (err) {
    const reason = err?.code === 403 ? 'Service Account tidak memiliki akses. Bagikan spreadsheet kepada email Service Account dengan akses Viewer.' : err?.code === 404 ? `Sheet ${SHEET_NAME} tidak ditemukan atau Spreadsheet ID salah.` : 'Pastikan Google Sheets API aktif dan Service Account memiliki akses Viewer.';
    throw apiError('Database siswa tidak dapat diakses.', `${reason} Detail Google API: ${err.message}`);
  }
  if (values.length < 1) { console.log('[STUDENT DATA] Rows found: 0'); return []; }
  const headers = values[0];
  const totalRows = Math.max(0, values.length - 1);
  lastSheetMeta = { headers, rowCount: totalRows, totalRows, rowsRead: totalRows, rangeRowsRead: totalRows, studentsLoaded: 0, classCounts: {}, classNames: [], excludedRows: 0, warnings: { emptyUserSerial: 0, duplicateUserSerial: 0, emptyName: 0, emptyClass: 0, invalidRows: [] } };
  console.log('[STUDENT DATA] Headers:', headers.join(' | '));
  const indexes = {
    studentCode: headerIndex(headers, ['userserial']), name: headerIndex(headers, ['namasiswa']),
    schoolName: headerIndex(headers, ['namasekolah']), grade: headerIndex(headers, ['grade']), className: headerIndex(headers, ['kelas'])
  };
  if (Object.values(indexes).some(index => index < 0)) {
    throw apiError('Header sheet siswa tidak sesuai struktur database terbaru.', `Header terbaca: ${headers.join(' | ')}. Wajib memiliki: User Serial, Nama Siswa, Nama Sekolah, Grade, Kelas.`);
  }
  const rawStudents = values.slice(1).map((row, offset) => ({ rowNumber: offset + 2, studentCode: String(row[indexes.studentCode] || '').trim(), name: String(row[indexes.name] || '').trim(), schoolName: String(row[indexes.schoolName] || '').trim(), grade: String(row[indexes.grade] || '').trim(), className: normalizeClassName(row[indexes.className]) })).filter(student => Object.values(student).some(value => String(value || '').trim() !== ''));
  const students = normalizeStudents(rawStudents, totalRows);
  lastSheetMeta.studentsLoaded = students.length;
  console.log('[STUDENT DATA] Total rows from Google Sheets:', totalRows);
  console.log('[STUDENT DATA] Classes included:', JSON.stringify(lastSheetMeta.classCounts));
  console.log('[STUDENT DATA] Validation warnings:', JSON.stringify(lastSheetMeta.warnings));
  console.log('[STUDENT DATA] Excluded:', lastSheetMeta.excludedRows);
  console.log('[STUDENT DATA] Rows found:', totalRows);
  console.log('[STUDENT DATA] Students loaded:', students.length);
  return students;
}
function loadSubmissions() { try { return JSON.parse(fs.readFileSync(SUBMISSIONS_FILE, 'utf8')); } catch (_) { return []; } }
function saveSubmissions(rows) { fs.mkdirSync(DATA_DIR, { recursive: true }); fs.writeFileSync(SUBMISSIONS_FILE, JSON.stringify(rows, null, 2)); }
function isoDate(date) { return date.toISOString().slice(0, 10); }
const APP_TIMEZONE = 'Asia/Makassar';
const WEEK_ANCHOR = '2026-08-03';
function zonedParts(date) {
  const parts = new Intl.DateTimeFormat('en-CA', { timeZone: APP_TIMEZONE, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', weekday: 'long', hourCycle: 'h23' }).formatToParts(new Date(date));
  return Object.fromEntries(parts.filter(part => part.type !== 'literal').map(part => [part.type, part.value]));
}
function weekRange(date) {
  const p = zonedParts(date);
  const dayKey = `${p.year}-${p.month}-${p.day}`;
  const current = Date.parse(`${dayKey}T00:00:00Z`);
  const anchor = Date.parse(`${WEEK_ANCHOR}T00:00:00Z`);
  if (!Number.isFinite(current) || current < anchor) return null;
  const weekNumber = Math.floor((current - anchor) / 86400000 / 7) + 1;
  const startDate = new Date(anchor + (weekNumber - 1) * 7 * 86400000);
  const endDate = new Date(startDate.getTime() + 6 * 86400000);
  return { weekNumber, start: startDate.toISOString().slice(0, 10), end: endDate.toISOString().slice(0, 10), date: dayKey, time: `${p.hour}:${p.minute}` };
}
function weekFromNumber(number) {
  const weekNumber = Number(number);
  if (!Number.isInteger(weekNumber) || weekNumber < 1) return null;
  const anchor = Date.parse(`${WEEK_ANCHOR}T00:00:00Z`);
  const startDate = new Date(anchor + (weekNumber - 1) * 7 * 86400000);
  const endDate = new Date(startDate.getTime() + 6 * 86400000);
  return { weekNumber, start: startDate.toISOString().slice(0, 10), end: endDate.toISOString().slice(0, 10) };
}
function dayName(date) { return { Sunday:'Minggu', Monday:'Senin', Tuesday:'Selasa', Wednesday:'Rabu', Thursday:'Kamis', Friday:'Jumat', Saturday:'Sabtu' }[zonedParts(date).weekday]; }
function enrichSubmission(row) {
  if (!row || typeof row !== 'object') return row;
  let week = Number(row.weekNumber);
  if (!Number.isInteger(week) && row.weekStart) {
    const start = Date.parse(`${String(row.weekStart).slice(0, 10)}T00:00:00Z`);
    const anchor = Date.parse(`${WEEK_ANCHOR}T00:00:00Z`);
    if (Number.isFinite(start)) week = Math.floor((start - anchor) / 86400000 / 7) + 1;
  }
  const weekStartDate = row.weekStartDate || row.weekStart || null;
  const weekEndDate = row.weekEndDate || row.weekEnd || null;
  const inputDate = row.date || (row.scannedAt ? zonedParts(row.scannedAt) : null);
  const inputDay = typeof inputDate === 'string' ? inputDate.slice(0, 10) : (inputDate ? `${inputDate.year}-${inputDate.month}-${inputDate.day}` : null);
  const inputStatus = row.inputStatus || (inputDay && weekEndDate ? (inputDay > String(weekEndDate).slice(0, 10) ? 'late' : 'on_time') : null);
  return { ...row, recordId: row.recordId || row.id, recordStatus: row.recordStatus === 'cancelled' ? 'cancelled' : 'active', weekNumber: Number.isInteger(week) && week > 0 ? week : null, weekStartDate, weekEndDate, inputStatus, inputStatusLabel: inputStatus === 'late' ? 'Input Terlambat' : inputStatus === 'on_time' ? 'Tepat Waktu' : null };
}

const cache = { students: null, at: 0, inflight: null };
async function getStudents(force = false) {
  if (!force && cache.students && Date.now() - cache.at < 60_000) return cache.students;
  // Share one upstream request between the dashboard, diagnostic page, and any
  // other tabs opened at the same time. Apps Script is slow and can otherwise
  // throttle concurrent requests, leaving the dashboard stuck on “Memuat…”.
  if (cache.inflight) return cache.inflight;
  cache.inflight = readStudents().then(students => { cache.students = students; cache.at = Date.now(); return students; }).finally(() => { cache.inflight = null; });
  return cache.inflight;
}
async function json(res, status, body) { const payload = JSON.stringify(body); res.writeHead(status, { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store', 'Access-Control-Allow-Origin': '*' }); res.end(payload); }
function readBody(req) { return new Promise((resolve, reject) => { let data = ''; req.on('data', chunk => { data += chunk; if (data.length > 1e6) reject(Error('Payload terlalu besar')); }); req.on('end', () => { try { resolve(data ? JSON.parse(data) : {}); } catch (e) { reject(e); } }); req.on('error', reject); }); }
function serveStatic(req, res) { const urlPath = req.url === '/' ? '/index.html' : req.url.split('?')[0]; const file = path.resolve(__dirname, '.' + urlPath); if (!file.startsWith(path.resolve(__dirname)) || !fs.existsSync(file) || fs.statSync(file).isDirectory()) return json(res, 404, { error: 'Not found' }); const ext = path.extname(file); const types = { '.html':'text/html; charset=utf-8', '.js':'text/javascript; charset=utf-8', '.css':'text/css; charset=utf-8', '.json':'application/json; charset=utf-8' }; res.writeHead(200, { 'Content-Type': types[ext] || 'application/octet-stream', 'Cache-Control': 'no-store, no-cache, must-revalidate' }); fs.createReadStream(file).pipe(res); }

const server = http.createServer(async (req, res) => {
  try {
    if (req.method === 'OPTIONS') return json(res, 204, {});
    if (req.url === '/api/health') return json(res, 200, { ok: true, spreadsheetId: SPREADSHEET_ID, sheetNameConfigured: Boolean(SHEET_NAME), credentialsConfigured: Boolean(CREDENTIALS) });
    if (req.method === 'GET' && req.url === '/api/students/test') {
      const diagnostic = { success: false, connected: false, spreadsheetId: SPREADSHEET_ID, sheetName: SHEET_NAME, range: `${SHEET_NAME}!A:E`, rowCount: 0, totalRows: 0, rowsRead: 0, rangeRowsRead: 0, classCounts: {}, classNames: [], excludedRows: 0, studentsLoaded: 0, skippedRows: [], warnings: {}, headers: [], sampleStudents: [] };
      try { const students = await getStudents(true); diagnostic.success = true; diagnostic.connected = true; diagnostic.headers = lastSheetMeta.headers; diagnostic.rowCount = lastSheetMeta.rowCount; diagnostic.totalRows = lastSheetMeta.totalRows; diagnostic.rowsRead = lastSheetMeta.rowsRead ?? lastSheetMeta.totalRows; diagnostic.rangeRowsRead = lastSheetMeta.rangeRowsRead ?? diagnostic.rowsRead; diagnostic.classCounts = lastSheetMeta.classCounts; diagnostic.classNames = lastSheetMeta.classNames; diagnostic.warnings = lastSheetMeta.warnings || {}; diagnostic.excludedRows = lastSheetMeta.excludedRows; diagnostic.skippedRows = lastSheetMeta.skippedRows || []; diagnostic.studentsLoaded = lastSheetMeta.studentsLoaded ?? students.length; diagnostic.sampleStudents = students.slice(0, 3); return json(res, 200, diagnostic); }
      catch (err) { diagnostic.message = err.message; diagnostic.error = err.message; diagnostic.errorCode = 'STUDENT_DATA_ERROR'; diagnostic.details = err.details || 'Tidak ada detail tambahan.'; console.error('[STUDENT DATA ERROR]', err.message, err.details || ''); return json(res, 200, diagnostic); }
    }
    if (req.method === 'GET' && req.url.startsWith('/api/students')) {
      const students = await getStudents(req.url.includes('refresh=1'));
      const code = req.url.split('/api/students/')[1]?.split('?')[0];
      if (code) { const student = students.find(s => s.studentCode === decodeURIComponent(code)); return student ? json(res, 200, student) : json(res, 404, { error: 'Siswa tidak ditemukan.' }); }
      return json(res, 200, { success: true, data: students, students, source: 'google_sheets', sheetName: SHEET_NAME, meta: { ...lastSheetMeta, filteredRows: students.length } });
    }
    if (req.method === 'GET' && req.url.startsWith('/api/submissions')) { const submissions = loadSubmissions().map(enrichSubmission); return json(res, 200, { success: true, data: submissions, submissions }); }
    if (req.method === 'POST' && req.url === '/api/submissions') {
      const body = await readBody(req); const students = await getStudents(); const student = students.find(s => s.studentCode === body.studentCode); if (!student) return json(res, 404, { error: 'Student code tidak ditemukan.' });
      const scannedAt = body.scannedAt ? new Date(body.scannedAt) : new Date();
      if (Number.isNaN(scannedAt.getTime())) return json(res, 400, { error: 'Tanggal scan tidak valid.' });
      const week = weekFromNumber(body.weekNumber) || weekRange(scannedAt);
      if (!week) return json(res, 400, { error: `Week tugas tidak valid. Week dimulai ${WEEK_ANCHOR}.` });
      const taskCount = Math.max(1, Math.min(2, Number(body.taskCount) || 1));
      const rows = loadSubmissions().map(enrichSubmission); const current = rows.filter(x => x.studentCode === student.studentCode && x.weekNumber === week.weekNumber && x.recordStatus !== 'cancelled'); const duplicate = rows.some(x => x.studentCode === student.studentCode && x.recordStatus !== 'cancelled' && Math.abs(new Date(x.scannedAt) - scannedAt) < 10_000); const scheduleStatus = Array.isArray(student.studyDays) && student.studyDays.length && student.studyDays.includes(dayName(scannedAt)) ? 'on_schedule' : 'outside_schedule'; const inputParts = zonedParts(scannedAt); const inputDate = `${inputParts.year}-${inputParts.month}-${inputParts.day}`; const inputStatus = inputDate > week.end ? 'late' : 'on_time';
      if (body.confirmDuplicate === false && duplicate) return json(res, 409, { duplicateWarning: true, student, existing: current[current.length - 1] });
      const scanGroupId = crypto.randomUUID(); const submissions = Array.from({ length: taskCount }, (_, offset) => { const id = crypto.randomUUID(); const submissionNumber = current.length + offset + 1; return { id, recordId: id, scanGroupId, studentCode: student.studentCode, scannedAt: scannedAt.toISOString(), date: inputDate, time: `${inputParts.hour}:${inputParts.minute}`, dayOfWeek: dayName(scannedAt), weekNumber: week.weekNumber, weekStart: week.start, weekEnd: week.end, weekStartDate: week.start, weekEndDate: week.end, submissionNumber, taskNumber: offset + 1, taskCount, scheduleStatus, inputStatus, inputStatusLabel: inputStatus === 'late' ? 'Input Terlambat' : 'Tepat Waktu', recordStatus: 'active', status: submissionNumber > 2 ? 'additional' : scheduleStatus }; });
      rows.push(...submissions); saveSubmissions(rows); return json(res, 201, { submission: submissions[0], submissions, student, duplicateWarning: duplicate });
    }
    if (req.method === 'PATCH' && req.url.startsWith('/api/submissions/')) {
      const id = decodeURIComponent(req.url.split('/api/submissions/')[1].split('?')[0]);
      const body = await readBody(req); const rows = loadSubmissions().map(enrichSubmission); const index = rows.findIndex(row => row.id === id);
      if (index < 0) return json(res, 404, { error: 'Record scan tidak ditemukan.' });
      const record = rows[index]; if (record.recordStatus === 'cancelled') return json(res, 409, { error: 'Record scan sudah dibatalkan dan tidak dapat diubah.' });
      const now = new Date();
      if (body.action === 'cancel' || body.action === 'cancelScan') {
        const reason = String(body.cancelReason || '').trim(); if (!reason) return json(res, 400, { error: 'Alasan pembatalan wajib dipilih.' });
        rows[index] = { ...record, recordStatus: 'cancelled', cancelledAt: now.toISOString(), cancelledBy: String(body.cancelledBy || 'Wali Kelas'), cancelReason: reason, updatedAt: now.toISOString(), updatedBy: String(body.updatedBy || body.cancelledBy || 'Wali Kelas') };
        saveSubmissions(rows); return json(res, 200, { success: true, message: 'Scan berhasil dibatalkan', submission: rows[index] });
      }
      const week = body.weekNumber == null ? weekFromNumber(record.weekNumber) : weekFromNumber(body.weekNumber);
      if (!week) return json(res, 400, { error: 'Week tugas tidak valid.' });
      const inputParts = zonedParts(record.scannedAt || now); const inputDate = `${inputParts.year}-${inputParts.month}-${inputParts.day}`; const inputStatus = inputDate > week.end ? 'late' : 'on_time';
      const requestedTaskCount = body.taskCount == null && body.taskQuantity == null
        ? Number(record.taskCount || 1)
        : Math.max(1, Math.min(2, Number(body.taskCount ?? body.taskQuantity) || 1));
      const updated = { ...record, weekNumber: week.weekNumber, weekStart: week.start, weekEnd: week.end, weekStartDate: week.start, weekEndDate: week.end, taskCount: requestedTaskCount, inputStatus, inputStatusLabel: inputStatus === 'late' ? 'Input Terlambat' : 'Tepat Waktu', updatedAt: now.toISOString(), updatedBy: String(body.updatedBy || 'Wali Kelas') };
      rows[index] = updated;
      if (requestedTaskCount === 1 && Number(record.taskCount || 1) > 1 && record.scanGroupId) {
        const primaryIndex = Number(record.taskNumber) === 1 ? index : rows.findIndex((row, i) => i !== index && row.scanGroupId === record.scanGroupId && row.recordStatus !== 'cancelled' && Number(row.taskNumber) === 1);
        if (primaryIndex >= 0 && primaryIndex !== index) {
          rows[index] = { ...rows[index], recordStatus: 'cancelled', cancelledAt: now.toISOString(), cancelledBy: String(body.updatedBy || 'Wali Kelas'), cancelReason: 'Jumlah tugas diubah menjadi 1', updatedAt: now.toISOString(), updatedBy: String(body.updatedBy || 'Wali Kelas') };
          rows[primaryIndex] = { ...rows[primaryIndex], weekNumber: week.weekNumber, weekStart: week.start, weekEnd: week.end, weekStartDate: week.start, weekEndDate: week.end, taskCount: 1, inputStatus, inputStatusLabel: inputStatus === 'late' ? 'Input Terlambat' : 'Tepat Waktu', updatedAt: now.toISOString(), updatedBy: String(body.updatedBy || 'Wali Kelas') };
        }
        rows.forEach((row, i) => { const keep = primaryIndex >= 0 ? primaryIndex : index; if (i !== keep && row.scanGroupId === record.scanGroupId && row.recordStatus !== 'cancelled' && Number(row.taskNumber) > 1) rows[i] = { ...row, recordStatus: 'cancelled', cancelledAt: now.toISOString(), cancelledBy: String(body.updatedBy || 'Wali Kelas'), cancelReason: 'Jumlah tugas diubah menjadi 1', updatedAt: now.toISOString(), updatedBy: String(body.updatedBy || 'Wali Kelas') }; });
      }
      if (requestedTaskCount === 2) {
        const groupId = record.scanGroupId || crypto.randomUUID(); rows[index] = { ...rows[index], scanGroupId: groupId, taskCount: 2 };
        const sibling = rows.find(row => row.id !== id && row.scanGroupId === groupId && row.recordStatus !== 'cancelled' && Number(row.taskNumber) !== Number(record.taskNumber));
        if (!sibling) {
          const taskNumber = Number(record.taskNumber) === 1 ? 2 : 1;
          rows.push({ ...rows[index], id: crypto.randomUUID(), scanGroupId: groupId, taskNumber, submissionNumber: Number(record.submissionNumber || 1) + (taskNumber === 2 ? 1 : -1), taskCount: 2, status: Number(record.submissionNumber || 1) + (taskNumber === 2 ? 1 : -1) > 2 ? 'additional' : record.status });
        }
      }
      saveSubmissions(rows); return json(res, 200, { success: true, message: 'Scan berhasil diedit', submission: rows[index], submissions: rows.filter(row => row.scanGroupId === record.scanGroupId || row.id === id) });
    }
    return serveStatic(req, res);
  } catch (err) { return json(res, err.message === 'Payload terlalu besar' ? 413 : 500, { success: false, message: err.message || 'Server error', error: err.message || 'Server error', errorCode: 'STUDENT_DATA_ERROR', details: err.details || undefined }); }
});
server.listen(PORT, HOST, () => console.log(`Student Task Tracker backend: http://${HOST}:${PORT}`));
