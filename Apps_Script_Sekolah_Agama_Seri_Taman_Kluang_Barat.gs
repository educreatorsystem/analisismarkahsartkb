/** BACKEND BARU (pilihan) — hanya gunakan jika Web App sedia ada tidak menyokong getMarkah / saveMarkah.
 * Letakkan dalam Apps Script yang terikat kepada Google Sheet simpanan markah.
 * Deploy > New deployment > Web app > Execute as Me.
 * Akses web app mesti dikawal mengikut dasar privasi sekolah.
 */
const SPREADSHEET_ID = '1bcgeY1mJbV8ovP0i371Z4FBQP9878sqKVnoh4cNYBN4';
const SHEET_NAME = 'MARKAH_KAFA_STKB';
const HEADERS = ['recordKey','studentId','ic','nama','kelas','pentaksiran','subjek','markah','gred','tafsiran','timestamp'];
function getGrade_(score) {
  if (score === 'TD') return ['TD','Tidak Daftar'];
  if (score === 'TH') return ['TH','Tidak Hadir'];
  const n = Number(score);
  if (score === '' || score === null || !Number.isFinite(n) || n < 0 || n > 100) return ['',''];
  if (n >= 85) return ['A','Cemerlang'];
  if (n >= 70) return ['B','Sangat Baik'];
  if (n >= 55) return ['C','Baik'];
  if (n >= 40) return ['D','Sederhana'];
  return ['E','Gagal'];
}
function sheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) { sh = ss.insertSheet(SHEET_NAME); sh.appendRow(HEADERS); }
  if (sh.getLastRow() === 0) sh.appendRow(HEADERS);
  return sh;
}
function records_() {
  const sh = sheet_(); const rows = sh.getDataRange().getDisplayValues(); const output = {};
  rows.slice(1).forEach(row => { const r = Object.fromEntries(HEADERS.map((h,i)=>[h,row[i] || ''])); if(r.recordKey){
    r.markah = /^(TH|TD)$/.test(r.markah) ? r.markah : r.markah === '' ? '' : Number(r.markah);
    const [grade,desc]=getGrade_(r.markah); r.gred=grade; r.tafsiran=desc; output[r.recordKey]=r;
  }});
  return output;
}
function json_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function doGet(e) {
  try { if(e && e.parameter && e.parameter.action && e.parameter.action !== 'getMarkah') throw Error('Action tidak dikenali');
    return json_({status:'success',data:records_()});
  } catch(err) { return json_({status:'error',message:String(err.message || err)}); }
}
function doPost(e) {
  const lock=LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const body=JSON.parse(e.postData.contents || '{}');
    if(body.action !== 'saveMarkah') throw Error('Action tidak dikenali');
    const sh=sheet_(), data=records_();
    (body.deleteKeys||[]).forEach(k=>delete data[String(k)]);
    (body.records||[]).forEach(r=>{
      const key=String(r.recordKey || '');
      if(!key) return;
      const score=String(r.markah || r.markah === 0 ? r.markah : '').toUpperCase();
      const markah = score === 'TD' || score === 'TH' ? score : Number(score);
      if (!['TD','TH'].includes(score) && (score === '' || !Number.isFinite(markah) || markah<0 || markah>100)) return;
      const [gred,tafsiran]=getGrade_(markah);
      data[key]={recordKey:key,studentId:String(r.studentId||''),ic:String(r.ic||''),nama:String(r.nama||''),kelas:String(r.kelas||''),pentaksiran:String(r.pentaksiran||''),subjek:String(r.subjek||''),markah,gred,tafsiran,timestamp:new Date().toISOString()};
    });
    if(sh.getLastRow()>1) sh.getRange(2,1,sh.getLastRow()-1,HEADERS.length).clearContent();
    const rows=Object.values(data).map(r=>HEADERS.map(h=>r[h]??''));
    if(rows.length) sh.getRange(2,1,rows.length,HEADERS.length).setValues(rows);
    return json_({status:'success',saved:(body.records||[]).length});
  }catch(err){return json_({status:'error',message:String(err.message||err)});}finally{lock.releaseLock();}
}
