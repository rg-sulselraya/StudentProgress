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
function weekFromNumber_(number) { const weekNumber = Number(number); if (!Number.isInteger(weekNumber) || weekNumber < 1) return null; const anchor = Date.parse(WEEK_ANCHOR + 'T00:00:00Z'); const start = new Date(anchor + (weekNumber - 1) * 7 * 86400000); const end = new Date(start.getTime() + 6 * 86400000); return {weekNumber, start:Utilities.formatDate(start, 'UTC', 'yyyy-MM-dd'), end:Utilities.formatDate(end, 'UTC', 'yyyy-MM-dd')}; }
function dayName_(date) { return {Sunday:'Minggu',Monday:'Senin',Tuesday:'Selasa',Wednesday:'Rabu',Thursday:'Kamis',Friday:'Jumat',Saturday:'Sabtu'}[Utilities.formatDate(new Date(date), APP_TIMEZONE, 'EEEE')]; }
function scanSheet_() { const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID); let sheet = spreadsheet.getSheetByName(SCAN_SHEET_NAME); if (!sheet) sheet = spreadsheet.insertSheet(SCAN_SHEET_NAME); if (sheet.getLastRow() === 0) sheet.appendRow(['Record ID','Scan Group ID','Student Code','Tanggal Scan','Jam Scan','Hari','Week Number','Week Start','Week End','Task Number','Task Count','Status']); return sheet; }
function existingScans_(sheet) { const values = sheet.getDataRange().getDisplayValues(); if (values.length < 2) return []; const headers = values[0].map(String); const ix = {}; headers.forEach((h,i) => ix[h] = i); return values.slice(1).map(r => { const date = r[ix['Tanggal Scan']] || ''; const time = r[ix['Jam Scan']] || '00:00'; return {id:r[ix.ID],scanGroupId:r[ix['Scan Group ID']],studentCode:r[ix['Student Code']],scannedAt:date ? `${date}T${time}:00+08:00` : null,date,time,dayOfWeek:r[ix.Hari],weekNumber:Number(r[ix['Week Number']]),weekStart:r[ix['Week Start']],weekEnd:r[ix['Week End']],taskNumber:Number(r[ix['Task Number']]),taskCount:Number(r[ix['Task Count']]),status:r[ix.Status]||'Tercatat',recordStatus:'active'}; }); }
 function saveScanLegacy_(e) { try { const body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); const code = String(body.studentCode || '').trim(); if (!code) return json_({error:'studentCode wajib diisi.'}); const data = students_(); const student = data.students.find(s => s.studentCode === code && allowedClass_(s.className)); if (!student) return json_({error:'Student code tidak ditemukan atau siswa tidak termasuk dalam cakupan program.'}); const scannedAt = body.scannedAt ? new Date(body.scannedAt) : new Date(); if (isNaN(scannedAt.getTime())) return json_({error:'Tanggal scan tidak valid.'}); const week = weekFromDate_(scannedAt); if (!week) return json_({error:'Tanggal scan berada sebelum Week 1 (3 Agustus 2026).'}); const selectedWeek = Number(body.weekNumber); if (Number.isInteger(selectedWeek) && selectedWeek !== week.weekNumber) return json_({error:'Week scan tidak sesuai dengan tanggal scan.',expectedWeek:week.weekNumber}); const taskCount = Math.max(1, Math.min(2, Number(body.taskCount) || 1)); const sheet = scanSheet_(); const previous = existingScans_(sheet).filter(row => row.studentCode === code && row.weekNumber === week.weekNumber && row.recordStatus !== 'cancelled'); const groupId = Utilities.getUuid(); const rows = []; for (let i = 1; i <= taskCount; i++) { const id = Utilities.getUuid(); const status = previous.length + i > 2 ? 'Tambahan' : 'Tercatat'; rows.push({id,scanGroupId:groupId,studentCode:code,scannedAt:Utilities.formatDate(scannedAt,APP_TIMEZONE,"yyyy-MM-dd'T'HH:mm:ssXXX"),date:Utilities.formatDate(scannedAt,APP_TIMEZONE,'yyyy-MM-dd'),time:Utilities.formatDate(scannedAt,APP_TIMEZONE,'HH:mm'),dayOfWeek:dayName_(scannedAt),weekNumber:week.weekNumber,weekStart:week.start,weekEnd:week.end,taskNumber:i,taskCount,status,recordStatus:'active'}); sheet.appendRow([id,groupId,code,rows[i-1].date,rows[i-1].time,rows[i-1].dayOfWeek,week.weekNumber,week.start,week.end,i,taskCount,status]); } return json_({ok:true,student,submission:rows[0],submissions:rows}); } catch (err) { return json_({ok:false,error:err.message}); } }
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'students';
    if (action === 'submissions') { const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SCAN_SHEET_NAME); if (sheet) ensureScanMetadataColumns_(sheet); return json_({submissions:sheet ? existingScans_(sheet) : []}); }
    const data = students_();
    if (action === 'test') return json_({connected:true,spreadsheetId:SPREADSHEET_ID,sheetName:SHEET_NAME,range:SHEET_NAME+'!A:F',rowCount:data.rowCount,totalRows:data.rowCount,rowsRead:data.rowsRead,rangeRowsRead:data.rangeRowsRead,headers:data.headers,classCounts:data.classCounts,classNames:data.classNames,skippedRows:data.skippedRows,sampleStudents:data.students.slice(0,3).map(s => ({studentCode:s.studentCode,name:s.name,className:s.className,studyDays:s.studyDays,schoolName:s.schoolName})),studentsLoaded:data.students.length});
    return json_({students:data.students,source:'google_apps_script',sheetName:SHEET_NAME,headers:data.headers,rowCount:data.rowCount,totalRows:data.rowCount,rowsRead:data.rowsRead,rangeRowsRead:data.rangeRowsRead,classCounts:data.classCounts,classNames:data.classNames});
  } catch (err) { return json_({connected:false,error:err.message,spreadsheetId:SPREADSHEET_ID,sheetName:SHEET_NAME,range:SHEET_NAME+'!A:F'}); }
}

// Stable record identifiers and mutation actions for the history screen.
function scanColumns_(sheet) {
  const lastColumn = Math.max(1, sheet.getLastColumn());
  const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0].map(String);
  const key = value => String(value || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');
  const find = names => { const wanted = names.map(key); return headers.findIndex(header => wanted.some(name => key(header) === name || key(header).indexOf(name) !== -1)); };
  const columns = {
    recordId: find(['recordid','id']), groupId: find(['scangroupid']), studentCode: find(['studentcode']), date: find(['tanggal_scan','tanggal_scan','inputdate']), time: find(['jam_scan','time']), day: find(['hari']), weekNumber: find(['weeknumber']), weekStart: find(['weekstart','weekstartdate']), weekEnd: find(['weekend','weekenddate']), taskNumber: find(['tasknumber']), taskCount: find(['taskcount','taskquantity']), status: find(['status']), updatedAt: find(['updatedat']), updatedBy: find(['updatedby']), cancelledAt: find(['cancelledat']), cancelledBy: find(['cancelledby']), cancelReason: find(['cancelreason'])
  };
  return {headers, columns};
}
function ensureScanMetadataColumns_(sheet) {
  let info = scanColumns_(sheet);
  const missing = [['Record ID','recordId'],['Updated At','updatedAt'],['Updated By','updatedBy'],['Cancelled At','cancelledAt'],['Cancelled By','cancelledBy'],['Cancel Reason','cancelReason']].filter(([,key]) => info.columns[key] < 0);
  if (missing.length) { const start = sheet.getLastColumn() + 1; sheet.getRange(1, start, 1, missing.length).setValues([missing.map(x => x[0])]); info = scanColumns_(sheet); }
  if (info.columns.recordId >= 0 && sheet.getLastRow() > 1) {
    const idRange = sheet.getRange(2, info.columns.recordId + 1, sheet.getLastRow() - 1, 1);
    const ids = idRange.getDisplayValues(); let changed = false;
    ids.forEach(row => { if (!String(row[0] || '').trim()) { row[0] = Utilities.getUuid(); changed = true; } });
    if (changed) idRange.setValues(ids);
  }
  return info;
}
function valueAt_(row, index) { return index >= 0 ? String(row[index] || '').trim() : ''; }
function existingScans_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getDisplayValues(); const info = scanColumns_(sheet); const c = info.columns;
  return values.slice(1).map((row, offset) => {
    const recordId = valueAt_(row, c.recordId) || `legacy-row-${offset + 2}`;
    const date = valueAt_(row, c.date), time = valueAt_(row, c.time) || '00:00';
    const rawStatus = valueAt_(row, c.status) || 'Tercatat'; const cancelled = /batal|cancel/i.test(rawStatus) || Boolean(valueAt_(row, c.cancelledAt));
    return {id:recordId,recordId,scanGroupId:valueAt_(row,c.groupId),studentCode:valueAt_(row,c.studentCode),scannedAt:date ? `${date}T${time}:00+08:00` : null,date,time,dayOfWeek:valueAt_(row,c.day),weekNumber:Number(valueAt_(row,c.weekNumber)),weekStart:valueAt_(row,c.weekStart),weekEnd:valueAt_(row,c.weekEnd),taskNumber:Number(valueAt_(row,c.taskNumber)) || 1,taskCount:Number(valueAt_(row,c.taskCount)) || 1,status:cancelled?'Dibatalkan':rawStatus,recordStatus:cancelled?'cancelled':'active',updatedAt:valueAt_(row,c.updatedAt)||null,updatedBy:valueAt_(row,c.updatedBy)||null,cancelledAt:valueAt_(row,c.cancelledAt)||null,cancelledBy:valueAt_(row,c.cancelledBy)||null,cancelReason:valueAt_(row,c.cancelReason)||''};
  });
}
function locateScanRecord_(sheet, recordId, info) {
  const values = sheet.getDataRange().getDisplayValues(); const c = info.columns; const wanted = String(recordId || '').trim();
  for (let i = 1; i < values.length; i++) { const current = valueAt_(values[i], c.recordId) || `legacy-row-${i + 1}`; if (current === wanted) return {rowNumber:i + 1, row:values[i]}; }
  return null;
}
function setCell_(sheet, rowNumber, column, value) { if (column >= 0) sheet.getRange(rowNumber, column + 1).setValue(value); }
function mutateScan_(body) {
  const action = String(body.action || '').trim(); const recordId = String(body.recordId || body.id || '').trim(); console.log('[SCAN MUTATION]', JSON.stringify({action,recordId,timestamp:new Date().toISOString()}));
  if (!recordId) return {success:false,message:'recordId wajib diisi.'};
  const sheet = scanSheet_(); const info = ensureScanMetadataColumns_(sheet); const c = info.columns; const found = locateScanRecord_(sheet, recordId, info);
  if (!found) { console.log('[SCAN MUTATION] record not found', recordId); return {success:false,message:'Record tidak ditemukan.'}; }
  const current = existingScans_(sheet).find(row => row.recordId === recordId); if (!current) return {success:false,message:'Record tidak ditemukan.'};
  if (current.recordStatus === 'cancelled') return {success:false,message:'Record sudah dibatalkan dan tidak dapat diubah.'};
  const now = new Date();
  if (action === 'cancelScan' || action === 'cancel') {
    const reason = String(body.cancelReason || '').trim(); if (!reason) return {success:false,message:'Alasan pembatalan wajib diisi.'};
    setCell_(sheet, found.rowNumber, c.status, 'Dibatalkan'); setCell_(sheet, found.rowNumber, c.cancelledAt, Utilities.formatDate(now, APP_TIMEZONE, "yyyy-MM-dd'T'HH:mm:ssXXX")); setCell_(sheet, found.rowNumber, c.cancelledBy, body.cancelledBy || 'Wali Kelas'); setCell_(sheet, found.rowNumber, c.cancelReason, reason); setCell_(sheet, found.rowNumber, c.updatedAt, Utilities.formatDate(now, APP_TIMEZONE, "yyyy-MM-dd'T'HH:mm:ssXXX")); setCell_(sheet, found.rowNumber, c.updatedBy, body.cancelledBy || 'Wali Kelas'); console.log('[SCAN MUTATION] cancelled', recordId); return {success:true,message:'Scan berhasil dibatalkan',recordId};
  }
  if (action !== 'editScan' && action !== 'edit') return {success:false,message:'Action scan tidak dikenali.'};
  const week = weekFromNumber_(Number(body.weekNumber)); if (!week) return {success:false,message:'Week tugas tidak valid.'};
  const taskCount = Math.max(1, Math.min(2, Number(body.taskQuantity ?? body.taskCount) || current.taskCount || 1));
  if (c.recordId >= 0 && !valueAt_(found.row, c.recordId)) setCell_(sheet, found.rowNumber, c.recordId, Utilities.getUuid());
  const values = sheet.getDataRange().getDisplayValues();
  const groupId = current.scanGroupId || `single-${recordId}`;
  const groupRows = values.slice(1).map((row, offset) => ({row, rowNumber:offset + 2, record:existingScans_(sheet)[offset]})).filter(item => item.record && item.record.recordStatus !== 'cancelled' && (item.record.scanGroupId || `single-${item.record.recordId}`) === groupId);
  const updateRow = item => { setCell_(sheet,item.rowNumber,c.weekNumber,week.weekNumber); setCell_(sheet,item.rowNumber,c.weekStart,week.start); setCell_(sheet,item.rowNumber,c.weekEnd,week.end); setCell_(sheet,item.rowNumber,c.taskCount,taskCount); setCell_(sheet,item.rowNumber,c.updatedAt,Utilities.formatDate(now,APP_TIMEZONE,"yyyy-MM-dd'T'HH:mm:ssXXX")); setCell_(sheet,item.rowNumber,c.updatedBy,body.updatedBy||'Wali Kelas'); };
  groupRows.forEach(updateRow);
  if (taskCount === 1) {
    groupRows.filter(item => item.rowNumber !== found.rowNumber).forEach(item => { setCell_(sheet,item.rowNumber,c.status,'Dibatalkan'); setCell_(sheet,item.rowNumber,c.cancelledAt,Utilities.formatDate(now,APP_TIMEZONE,"yyyy-MM-dd'T'HH:mm:ssXXX")); setCell_(sheet,item.rowNumber,c.cancelledBy,body.updatedBy||'Wali Kelas'); setCell_(sheet,item.rowNumber,c.cancelReason,'Jumlah tugas diubah menjadi 1'); });
  } else if (groupRows.length < 2) {
    const source = sheet.getRange(found.rowNumber,1,1,sheet.getLastColumn()).getValues()[0]; const newRow = source.slice();
    if (c.recordId >= 0) newRow[c.recordId] = Utilities.getUuid(); if (c.taskNumber >= 0) newRow[c.taskNumber] = Number(current.taskNumber) === 1 ? 2 : 1; if (c.taskCount >= 0) newRow[c.taskCount] = 2; if (c.status >= 0) newRow[c.status] = 'Tercatat'; if (c.updatedAt >= 0) newRow[c.updatedAt] = Utilities.formatDate(now,APP_TIMEZONE,"yyyy-MM-dd'T'HH:mm:ssXXX"); if (c.updatedBy >= 0) newRow[c.updatedBy] = body.updatedBy || 'Wali Kelas'; sheet.appendRow(newRow);
  }
  console.log('[SCAN MUTATION] edited', recordId); return {success:true,message:'Scan berhasil diedit',recordId};
}
function doPost(e) {
  try {
    const contents = e && e.postData && e.postData.contents; let body = {};
    if (contents) { try { body = JSON.parse(contents); } catch (_) { body = e.parameter || {}; } } else body = (e && e.parameter) || {};
    if (body.action === 'editScan' || body.action === 'cancelScan' || body.action === 'edit' || body.action === 'cancel') return json_(mutateScan_(body));
    return saveScanLegacy_({postData:{contents:JSON.stringify(body)}});
  } catch (err) { console.log('[SCAN MUTATION ERROR]', err.message); return json_({success:false,ok:false,message:err.message,error:err.message}); }
}
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'students';
    if (action === 'submissions') { const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SCAN_SHEET_NAME); return json_({submissions:sheet ? existingScans_(sheet) : []}); }
    const data = students_();
    if (action === 'test') return json_({connected:true,spreadsheetId:SPREADSHEET_ID,sheetName:SHEET_NAME,range:SHEET_NAME+'!A:F',rowCount:data.rowCount,totalRows:data.rowCount,rowsRead:data.rowsRead,rangeRowsRead:data.rangeRowsRead,headers:data.headers,classCounts:data.classCounts,classNames:data.classNames,skippedRows:data.skippedRows,sampleStudents:data.students.slice(0,3).map(s => ({studentCode:s.studentCode,name:s.name,className:s.className,studyDays:s.studyDays,schoolName:s.schoolName})),studentsLoaded:data.students.length});
    return json_({students:data.students,source:'google_apps_script',sheetName:SHEET_NAME,headers:data.headers,rowCount:data.rowCount,totalRows:data.rowCount,rowsRead:data.rowsRead,rangeRowsRead:data.rangeRowsRead,classCounts:data.classCounts,classNames:data.classNames});
  } catch (err) { return json_({connected:false,error:err.message,spreadsheetId:SPREADSHEET_ID,sheetName:SHEET_NAME,range:SHEET_NAME+'!A:F'}); }
}
