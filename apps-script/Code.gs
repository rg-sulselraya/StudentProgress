const SPREADSHEET_ID = '1FcnkMsR24hU6A-bF5BYil2AX5iq-Ib38-Ta0dN_luBM';
const SHEET_NAME = 'siswa';

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
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'students';
    const data = students_();
    if (action === 'test') return json_({connected:true,spreadsheetId:SPREADSHEET_ID,sheetName:SHEET_NAME,range:SHEET_NAME+'!A:F',rowCount:data.rowCount,totalRows:data.rowCount,rowsRead:data.rowsRead,rangeRowsRead:data.rangeRowsRead,headers:data.headers,classCounts:data.classCounts,classNames:data.classNames,skippedRows:data.skippedRows,sampleStudents:data.students.slice(0,3).map(s => ({studentCode:s.studentCode,name:s.name,className:s.className,studyDays:s.studyDays,schoolName:s.schoolName})),studentsLoaded:data.students.length});
    return json_({students:data.students,source:'google_apps_script',sheetName:SHEET_NAME,headers:data.headers,rowCount:data.rowCount,totalRows:data.rowCount,rowsRead:data.rowsRead,rangeRowsRead:data.rangeRowsRead,classCounts:data.classCounts,classNames:data.classNames});
  } catch (err) { return json_({connected:false,error:err.message,spreadsheetId:SPREADSHEET_ID,sheetName:SHEET_NAME,range:SHEET_NAME+'!A:F'}); }
}
