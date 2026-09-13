const SPREADSHEET_ID = '1FcnkMsR24hU6A-bF5BYil2AX5iq-Ib38-Ta0dN_luBM';
const SHEET_NAME = 'siswa';
const SCAN_SHEET_NAME = 'Scan';
const APP_TIMEZONE = 'Asia/Makassar';
const WEEK_ANCHOR = '2026-08-03';

function normalize_(value) { return String(value || '').trim().toLowerCase(); }
function headerIndex_(headers, names) {
  return headers.findIndex(h => names.some(n => normalize_(h).replace(/[^a-z0-9]/g, '').indexOf(n) !== -1));
}
function days_(value) {
  const aliases = {sen:'Senin',senin:'Senin',sel:'Selasa',selasa:'Selasa',rab:'Rabu',rabu:'Rabu',kam:'Kamis',kamis:'Kamis',jum:'Jumat',jumat:'Jumat',sab:'Sabtu',sabtu:'Sabtu',min:'Minggu',minggu:'Minggu'};
  return String(value || '').split(/[,;|/]/).map(x => aliases[normalize_(x)]).filter(Boolean).filter((x,i,a) => a.indexOf(x) === i);
}
function code_(row) {
  const identity = normalize_(row.email) || [row.name,row.className,row.schoolName].map(normalize_).join('|');
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, identity || JSON.stringify(row));
  return 'STD-' + bytes.map(b => ('0' + (b < 0 ? b + 256 : b).toString(16)).slice(-2)).join('').slice(0, 8).toUpperCase();
}
function students_() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet siswa tidak ditemukan.');
  const lastRow = sheet.getLastRow();
  const values = lastRow ? sheet.getRange(1, 1, lastRow, 6).getDisplayValues() : [];
  if (!values.length) return {headers:[], students:[]};
  const headers = values[0];
  const ix = {name:headerIndex_(headers,['namasiswa','nama']), email:headerIndex_(headers,['emailsiswa','email']), className:headerIndex_(headers,['kelas']), paymentDate:headerIndex_(headers,['tanggalbayar','bayar']), days:headerIndex_(headers,['haribelajar','jadwal']), schoolName:headerIndex_(headers,['namasekolah','sekolah'])};
  const required = ['name','email','className','paymentDate','days','schoolName'];
  const missing = required.filter(k => ix[k] < 0);
  if (missing.length) throw new Error('Header Google Sheet tidak cocok. Header terbaca: ' + headers.join(' | ') + '. Kolom hilang: ' + missing.join(', '));
  const used = {};
  const rowsWithContent = values.slice(1).filter(r => r.some(cell => String(cell || '').trim() !== ''));
  const skippedRows = [];
  const students = values.slice(1).map((r, i) => {
    const raw = {name:r[ix.name],email:r[ix.email],className:r[ix.className],paymentDate:r[ix.paymentDate],days:r[ix.days],schoolName:r[ix.schoolName]};
    if (!raw.name) { if (r.some(cell => String(cell || '').trim() !== '')) skippedRows.push(i + 2); return null; }
    const base = code_(raw); let code = base; let n = 0;
    while (used[code]) { n++; code = base + n; } used[code] = true;
    return {studentCode:code,name:String(raw.name).trim(),email:String(raw.email || '').trim(),className:String(raw.className).trim(),paymentDate:String(raw.paymentDate || '').trim(),studyDays:days_(raw.days),schoolName:String(raw.schoolName || '').trim(),active:true};
  }).filter(Boolean);
  // rowCount/rowsRead count non-empty student rows. rangeRowsRead is retained
  // only as a diagnostic because Google Sheets may include formatted empty rows
  // in getLastRow()/the requested range.
  const classCounts = {};
  students.forEach(s => { const className = String(s.className || '').trim(); if (className) classCounts[className] = (classCounts[className] || 0) + 1; });
  return {headers:headers, rowCount:rowsWithContent.length, rowsRead:rowsWithContent.length, rangeRowsRead:Math.max(0, values.length - 1), skippedRows:skippedRows, classCounts:classCounts, classNames:Object.keys(classCounts).sort(), students:students};
}
function json_(body) { return ContentService.createTextOutput(JSON.stringify(body)).setMimeType(ContentService.MimeType.JSON); }
function allowedClass_(value) { const v = String(value || '').trim().replace(/\s+/g, ' ').toUpperCase(); return /^12 KURMER(?:\s|$)/.test(v) || /^SIAP SNBT(?:\s|$)/.test(v) || /^SNBT KEDINASAN(?:\s|$)/.test(v); }
function weekFromDate_(date) { const key = Utilities.formatDate(new Date(date), APP_TIMEZONE, 'yyyy-MM-dd'); const current = Date.parse(key + 'T00:00:00Z'); const anchor = Date.parse(WEEK_ANCHOR + 'T00:00:00Z'); if (!Number.isFinite(current) || current < anchor) return null; const weekNumber = Math.floor((current - anchor) / 86400000 / 7) + 1; const start = new Date(anchor + (weekNumber - 1) * 7 * 86400000); const end = new Date(start.getTime() + 6 * 86400000); return {weekNumber, start:Utilities.formatDate(start, 'UTC', 'yyyy-MM-dd'), end:Utilities.formatDate(end, 'UTC', 'yyyy-MM-dd')}; }
function dayName_(date) { return {Sunday:'Minggu',Monday:'Senin',Tuesday:'Selasa',Wednesday:'Rabu',Thursday:'Kamis',Friday:'Jumat',Saturday:'Sabtu'}[Utilities.formatDate(new Date(date), APP_TIMEZONE, 'EEEE')]; }
function scanSheet_() { const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID); let sheet = spreadsheet.getSheetByName(SCAN_SHEET_NAME); if (!sheet) sheet = spreadsheet.insertSheet(SCAN_SHEET_NAME); if (sheet.getLastRow() === 0) sheet.appendRow(['ID','Scan Group ID','Student Code','Tanggal Scan','Jam Scan','Hari','Week Number','Week Start','Week End','Task Number','Task Count','Status']); return sheet; }
function existingScans_(sheet) { const values = sheet.getDataRange().getDisplayValues(); if (values.length < 2) return []; const headers = values[0].map(String); const ix = {}; headers.forEach((h,i) => ix[h] = i); return values.slice(1).map(r => { const date = r[ix['Tanggal Scan']] || ''; const time = r[ix['Jam Scan']] || '00:00'; return {id:r[ix.ID],scanGroupId:r[ix['Scan Group ID']],studentCode:r[ix['Student Code']],scannedAt:date ? `${date}T${time}:00+08:00` : null,date,time,dayOfWeek:r[ix.Hari],weekNumber:Number(r[ix['Week Number']]),weekStart:r[ix['Week Start']],weekEnd:r[ix['Week End']],taskNumber:Number(r[ix['Task Number']]),taskCount:Number(r[ix['Task Count']]),status:r[ix.Status]||'Tercatat',recordStatus:'active'}; }); }
 function doPost(e) { try { const body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); const code = String(body.studentCode || '').trim(); if (!code) return json_({error:'studentCode wajib diisi.'}); const data = students_(); const student = data.students.find(s => s.studentCode === code && allowedClass_(s.className)); if (!student) return json_({error:'Student code tidak ditemukan atau siswa tidak termasuk dalam cakupan program.'}); const scannedAt = body.scannedAt ? new Date(body.scannedAt) : new Date(); if (isNaN(scannedAt.getTime())) return json_({error:'Tanggal scan tidak valid.'}); const week = weekFromDate_(scannedAt); if (!week) return json_({error:'Tanggal scan berada sebelum Week 1 (3 Agustus 2026).'}); const selectedWeek = Number(body.weekNumber); if (Number.isInteger(selectedWeek) && selectedWeek !== week.weekNumber) return json_({error:'Week scan tidak sesuai dengan tanggal scan.',expectedWeek:week.weekNumber}); const taskCount = Math.max(1, Math.min(2, Number(body.taskCount) || 1)); const sheet = scanSheet_(); const previous = existingScans_(sheet).filter(row => row.studentCode === code && row.weekNumber === week.weekNumber && row.recordStatus !== 'cancelled'); const groupId = Utilities.getUuid(); const rows = []; for (let i = 1; i <= taskCount; i++) { const id = Utilities.getUuid(); const status = previous.length + i > 2 ? 'Tambahan' : 'Tercatat'; rows.push({id,scanGroupId:groupId,studentCode:code,scannedAt:Utilities.formatDate(scannedAt,APP_TIMEZONE,"yyyy-MM-dd'T'HH:mm:ssXXX"),date:Utilities.formatDate(scannedAt,APP_TIMEZONE,'yyyy-MM-dd'),time:Utilities.formatDate(scannedAt,APP_TIMEZONE,'HH:mm'),dayOfWeek:dayName_(scannedAt),weekNumber:week.weekNumber,weekStart:week.start,weekEnd:week.end,taskNumber:i,taskCount,status,recordStatus:'active'}); sheet.appendRow([id,groupId,code,rows[i-1].date,rows[i-1].time,rows[i-1].dayOfWeek,week.weekNumber,week.start,week.end,i,taskCount,status]); } return json_({ok:true,student,submission:rows[0],submissions:rows}); } catch (err) { return json_({ok:false,error:err.message}); } }
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'students';
    if (action === 'submissions') { const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SCAN_SHEET_NAME); return json_({submissions:sheet ? existingScans_(sheet) : []}); }
    const data = students_();
    if (action === 'test') return json_({connected:true,spreadsheetId:SPREADSHEET_ID,sheetName:SHEET_NAME,range:SHEET_NAME+'!A:F',rowCount:data.rowCount,totalRows:data.rowCount,rowsRead:data.rowsRead,rangeRowsRead:data.rangeRowsRead,headers:data.headers,classCounts:data.classCounts,classNames:data.classNames,skippedRows:data.skippedRows,sampleStudents:data.students.slice(0,3).map(s => ({studentCode:s.studentCode,name:s.name,className:s.className,studyDays:s.studyDays,schoolName:s.schoolName})),studentsLoaded:data.students.length});
    return json_({students:data.students,source:'google_apps_script',sheetName:SHEET_NAME,headers:data.headers,rowCount:data.rowCount,totalRows:data.rowCount,rowsRead:data.rowsRead,rangeRowsRead:data.rangeRowsRead,classCounts:data.classCounts,classNames:data.classNames});
  } catch (err) { return json_({connected:false,error:err.message,spreadsheetId:SPREADSHEET_ID,sheetName:SHEET_NAME,range:SHEET_NAME+'!A:F'}); }
}
