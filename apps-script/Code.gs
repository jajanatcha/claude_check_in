/**
 * ระบบเช็คชื่อ Claude AI — ตัวเชื่อม Google Sheets
 * เวอร์ชัน 1.2.0 (14 พ.ย. 2569)
 * วิธีใช้: เปิด Google Sheet > ส่วนขยาย > Apps Script > วางโค้ดนี้ทั้งหมด > ทำให้ใช้งานได้ (Deploy) เป็น Web app
 * ชีตที่สร้างให้อัตโนมัติ: Roster (รายชื่อ), Checkins (การเช็คชื่อ), Config (ตั้งค่า)
 */
const VERSION = '1.2.0';
const SS = SpreadsheetApp.getActiveSpreadsheet();
const ROSTER_H = ['id', 'คำนำหน้า', 'ชื่อ-นามสกุล', 'Name (EN)', 'บริษัท/หน่วยงาน', 'ตำแหน่ง', 'อีเมล', 'เบอร์โทร', 'Line ID', 'หมายเหตุ', 'รายละเอียดเพิ่มเติม'];
const CHECK_H = ['วันที่', 'id', 'คำนำหน้า', 'ชื่อ-นามสกุล', 'บริษัท/หน่วยงาน', 'เวลาเช็คอิน', 'มาสาย (นาที)', 'timestamp', 'ลายเซ็น'];

function sheet_(name, header) {
  let s = SS.getSheetByName(name);
  if (!s) {
    s = SS.insertSheet(name);
    s.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold').setBackground('#e6defc');
    s.setFrozenRows(1);
    s.getRange('A:Z').setNumberFormat('@'); // เก็บเป็นข้อความ ไม่ให้แปลงวันที่/เวลา
  } else if (s.getLastColumn() < header.length) {
    // อัปเกรดชีตจากเวอร์ชันเก่า: เติมหัวคอลัมน์ที่เพิ่มใหม่
    s.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold').setBackground('#e6defc');
    s.getRange('A:Z').setNumberFormat('@');
  }
  return s;
}
function rows_(s) {
  const n = s.getLastRow() - 1;
  return n > 0 ? s.getRange(2, 1, n, s.getLastColumn()).getDisplayValues() : [];
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// รายละเอียดเพิ่มเติม เก็บในชีตเป็นบรรทัดละ 1 หัวข้อ รูปแบบ "หัวข้อ: รายละเอียด" (แก้ในชีตได้โดยตรง)
function formatExtra_(extra) {
  return (extra || []).filter(x => x && (x.k || x.v)).map(x => String(x.k || '').replace(/[:\n]/g, ' ') + ': ' + String(x.v || '').replace(/\n/g, ' ')).join('\n');
}
function parseExtra_(text) {
  return String(text || '').split('\n').filter(l => l.trim()).map(l => {
    const i = l.indexOf(':');
    return i >= 0 ? { k: l.slice(0, i).trim(), v: l.slice(i + 1).trim() } : { k: '', v: l.trim() };
  });
}
function readRoster_() {
  return rows_(sheet_('Roster', ROSTER_H)).filter(r => r[0] && r[2])
    .map(r => ({ id: r[0], prefix: r[1], name: r[2], en: r[3], org: r[4], pos: r[5], email: r[6] || '', phone: r[7] || '', line: r[8] || '', note: r[9] || '', extra: parseExtra_(r[10]) }));
}
function readCfg_() {
  const o = {};
  rows_(sheet_('Config', ['key', 'value'])).forEach(r => { if (r[0]) o[r[0]] = r[1]; });
  return o;
}
function readChecks_(day) {
  const o = {};
  rows_(sheet_('Checkins', CHECK_H)).forEach(r => {
    if (r[0] === day && r[1]) o[r[1]] = { pid: r[1], name: r[3], org: r[4], t: Number(r[7]), sig: r[8] };
  });
  return o;
}

function doGet(e) {
  try {
    const day = (e && e.parameter && e.parameter.day) || '';
    return out_({ ok: true, version: VERSION, people: readRoster_(), cfg: readCfg_(), checks: readChecks_(day) });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const b = JSON.parse(e.postData.contents);
    if (b.action === 'checkin') {
      const s = sheet_('Checkins', CHECK_H);
      const sig = String(b.sig || '').length < 49000 ? String(b.sig || '') : '';
      const row = [b.day, b.id, b.prefix || '', b.name || '', b.org || '', b.time || '', String(b.late ?? ''), String(b.t), sig];
      const idx = rows_(s).findIndex(r => r[0] === b.day && r[1] === b.id);
      if (idx >= 0) s.getRange(idx + 2, 1, 1, row.length).setValues([row]);
      else s.appendRow(row);
    } else if (b.action === 'checkout') {
      const s = sheet_('Checkins', CHECK_H);
      const idx = rows_(s).findIndex(r => r[0] === b.day && r[1] === b.id);
      if (idx >= 0) s.deleteRow(idx + 2);
    } else if (b.action === 'roster') {
      const s = sheet_('Roster', ROSTER_H);
      if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, ROSTER_H.length).clearContent();
      const vals = (b.people || []).map(p => [p.id, p.prefix || '', p.name || '', p.en || '', p.org || '', p.pos || '', p.email || '', p.phone || '', p.line || '', p.note || '', formatExtra_(p.extra)]);
      if (vals.length) s.getRange(2, 1, vals.length, ROSTER_H.length).setValues(vals);
    } else if (b.action === 'config') {
      const s = sheet_('Config', ['key', 'value']);
      if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, 2).clearContent();
      const vals = Object.keys(b.cfg || {}).map(k => [k, String(b.cfg[k])]);
      if (vals.length) s.getRange(2, 1, vals.length, 2).setValues(vals);
    } else {
      return out_({ ok: false, error: 'unknown action' });
    }
    return out_({ ok: true });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
