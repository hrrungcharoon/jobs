/**
 * ระบบใบสมัครงานออนไลน์ — บริษัท รุ่งจรูญ แอนด์ ซันส์ จำกัด
 *  - หน้าเว็บผู้สมัคร  : <ลิงก์ /exec>
 *  - หน้าผู้ดูแล (HR)  : <ลิงก์ /exec>?page=admin
 * ข้อมูลลง Google Sheet, ไฟล์แนบลง Google Drive, สร้างใบสมัคร Word/PDF จากแม่แบบ FR-HRM-01-05
 */

/***** ตั้งค่า *****/
const COMPANY_NAME = 'บริษัท รุ่งจรูญ แอนด์ ซันส์ จำกัด';
const SPREADSHEET_ID = '1wWKZDAZBkk8xCQni2SzLzt5iLLm8MuaP5VbsKSU-JC8'; // Google Sheet ที่ใช้เก็บข้อมูลผู้สมัคร
const DATA_SHEET = 'ข้อมูลผู้สมัคร';
const SETTINGS_SHEET = 'ตั้งค่าระบบ (ห้ามแก้ไข)';
const ROOT_FOLDER_NAME = 'ระบบใบสมัครงาน (เว็บ)';
const DEFAULT_ADMIN_PASSWORD = 'rsc1234'; // รหัสผ่านหน้าผู้ดูแลครั้งแรก — เปลี่ยนได้ในหน้า "ตั้งค่า"
const MAX_FILE_MB = 10;
const TZ = 'Asia/Bangkok';

const DEFAULT_SETTINGS = {
  intro: 'กรุณากรอกข้อมูลให้ครบถ้วนตามความเป็นจริง',
  positions: [
    { name: 'พนักงานขาย', detail: '', open: true },
    { name: 'พนักงานธุรการ', detail: '', open: true }
  ],
  lists: {
    titles: ['นาย', 'นาง', 'นางสาว'],
    religions: ['พุทธ', 'คริสต์', 'อิสลาม'],
    blood: ['A', 'B', 'O', 'AB', 'ไม่ทราบ'],
    status: ['ใหม่', 'กำลังพิจารณา', 'นัดสัมภาษณ์', 'ผ่าน', 'ไม่ผ่าน', 'สำรอง']
  },
  personality: {
    enabled: false,
    title: 'แบบประเมินบุคลิกภาพ',
    instruction: 'โปรดเลือกพฤติกรรมที่เกิดขึ้นกับตนเองเป็นประจำ',
    groups: [],
    items: []            // [{text:'...', group:'...'}]
  },
  notifyEmail: '',
  makePdf: true,
  templateId: '',
  senderName: '',              // ชื่อผู้ส่งอีเมล (เว้นว่าง = "ฝ่ายบุคคล <ชื่อบริษัท>")
  replyTo: '',                 // อีเมลรับการตอบกลับ (เว้นว่าง = อีเมลบัญชีนี้)
  autoReplyEnabled: false,     // ส่งอีเมลตอบรับอัตโนมัติเมื่อมีผู้สมัคร
  autoReplyTemplate: 'ได้รับใบสมัครแล้ว',
  emailTemplates: null         // null = ใช้แม่แบบเริ่มต้นใน Mailing.gs
};

const FIXED_HEAD = ['เลขที่ใบสมัคร', 'วันที่สมัคร'];
const FIXED_TAIL = ['ใบสมัคร (Google Doc)', 'ใบสมัคร (PDF)', 'สถานะ', 'หมายเหตุ HR', 'วันนัดสัมภาษณ์', 'รหัสนัดหมาย (ปฏิทิน)'];

/***** หน้าเว็บ *****/
function doGet(e) {
  const admin = e && e.parameter && e.parameter.page === 'admin';
  const t = HtmlService.createTemplateFromFile(admin ? 'Admin' : 'Index');
  t.companyName = COMPANY_NAME;
  t.config = admin ? '{}' : safeJson_(getPublicConfig_());
  return t.evaluate()
    .setTitle((admin ? 'ผู้ดูแลระบบใบสมัคร | ' : 'สมัครงาน | ') + COMPANY_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function safeJson_(o) { return JSON.stringify(o).replace(/</g, '\\u003c'); }

/** รันครั้งแรก 1 ครั้ง: อนุญาตสิทธิ์ + สร้างชีท/โฟลเดอร์ */
function setup() {
  const ss = getSpreadsheet_();
  getDataSheet_();
  saveSettings_(getSettings_());
  Logger.log('Google Sheet: ' + ss.getUrl());
  Logger.log('โฟลเดอร์ระบบ: ' + getFolder_('root').getUrl());
  Logger.log('รหัสผ่านหน้าผู้ดูแล (ครั้งแรก): ' + DEFAULT_ADMIN_PASSWORD);
}

function getPublicConfig_() {
  const s = getSettings_();
  const p = s.personality || {};
  const usePersonality = p.enabled && p.items && p.items.length;
  return {
    company: COMPANY_NAME,
    intro: String(s.intro || '').replace(/\s*ช่องที่มีเครื่องหมาย \* จำเป็นต้องกรอก/, ''),
    maxMb: MAX_FILE_MB,
    steps: getSchema_().filter(function (st) { return !st.personality || usePersonality; }),
    positions: (s.positions || []).filter(function (x) { return x.open && x.name; }).map(function (x) { return { name: x.name, detail: x.detail || '' }; }),
    lists: s.lists,
    personality: usePersonality ? { title: p.title, instruction: p.instruction, items: p.items.map(function (i) { return i.text; }) } : null,
    langs: LANGS, langSkills: LANG_SKILLS, langLevels: LANG_LEVELS
  };
}

/***** รับใบสมัคร *****/
function submitApplication(payload) {
  payload = payload || {};
  const ans = payload.answers || {};
  const files = payload.files || {};
  const settings = getSettings_();
  const fields = getFields_();
  const rec = {};

  fields.forEach(function (f) {
    if (f.type === 'file' || f.type === 'langtable') return;
    let v = ans[f.k];
    if (!isVisible_(f, ans)) v = '';
    if (f.type === 'checks' || f.type === 'personality') {
      v = Array.isArray(v) ? v.map(String).filter(String) : [];
    } else {
      v = v == null ? '' : String(v).trim();
      if (f.type === 'date' && v) v = toThaiDate_(v);
      if (f.type === 'thaiid' && v) {
        v = v.replace(/\D/g, '');
        if (!thaiIdValid_(v)) throw new Error('เลขบัตรประชาชนไม่ถูกต้อง');
      }
      if (f.type === 'check') v = v ? 'ใช่' : '';
    }
    const empty = Array.isArray(v) ? !v.length : v === '';
    if (f.req && empty && isVisible_(f, ans)) throw new Error('กรุณากรอก: ' + f.label);
    rec[f.k] = v;
  });
  LANGS.forEach(function (l) { LANG_SKILLS.forEach(function (s) {
    const k = 'lang_' + l[0] + '_' + s[0];
    rec[k] = LANG_LEVELS.indexOf(ans[k]) > -1 ? ans[k] : '';
  }); });
  const openPos = (settings.positions || []).filter(function (p) { return p.open; }).map(function (p) { return p.name; });
  if (rec.position && openPos.length && openPos.indexOf(rec.position) === -1) throw new Error('ตำแหน่งที่เลือกปิดรับสมัครแล้ว กรุณาเลือกใหม่');
  rec.personality_summary = personalitySummary_(rec.personality, settings);

  const ts = new Date();
  const name = (rec.title + rec.first_name + ' ' + rec.last_name).trim();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let rowNo, appNo, sheet, headers;
  try {
    sheet = getDataSheet_();
    headers = ensureHeaders_(sheet);
    rowNo = sheet.getLastRow() + 1;
    appNo = 'APP' + Utilities.formatDate(ts, TZ, 'yyMMdd') + '-' + ('000' + (rowNo - 1)).slice(-4);
    sheet.getRange(rowNo, 1, 1, 2).setValues([[appNo, ts]]);   // จองแถวไว้ก่อน
  } finally {
    lock.releaseLock();
  }

  // ไฟล์แนบ
  const fileLinks = {};
  const fileFields = fields.filter(function (f) { return f.type === 'file'; });
  const hasFiles = fileFields.some(function (f) { return (files[f.k] || []).length; });
  if (hasFiles) {
    const folder = getFolder_('files').createFolder(appNo + ' ' + name);
    fileFields.forEach(function (f) {
      fileLinks[f.k] = (files[f.k] || []).map(function (x) {
        const bytes = Utilities.base64Decode(x.data);
        if (bytes.length > MAX_FILE_MB * 1024 * 1024) throw new Error('ไฟล์ ' + x.name + ' ใหญ่เกิน ' + MAX_FILE_MB + ' MB');
        return folder.createFile(Utilities.newBlob(bytes, x.type || 'application/octet-stream', f.label + '_' + x.name)).getUrl();
      }).join('\n');
    });
  }

  const row = buildRow_(headers, rec, fileLinks, appNo, ts);
  sheet.getRange(rowNo, 1, 1, row.length).setValues([row]);
  sheet.getRange(rowNo, 2).setNumberFormat('dd/MM/yyyy HH:mm');

  // สร้างใบสมัคร Word/PDF (ถ้าล้มเหลว ใบสมัครยังถูกบันทึก)
  rec.photo = fileLinks.photo || '';
  try {
    const docs = generateDocs_(rec, ts, appNo, settings);
    writeDocLinks_(sheet, headers, rowNo, docs);
  } catch (e) { console.error('สร้างเอกสารไม่สำเร็จ', e); }
  try { rebuildMailMerge_(); } catch (e) { console.error('อัปเดต MailMerge ไม่สำเร็จ', e); }
  try { sendAutoReply_(rowNo, settings); } catch (e) { console.error('ส่งอีเมลตอบรับไม่สำเร็จ', e); }

  if (settings.notifyEmail) {
    try {
      MailApp.sendEmail(settings.notifyEmail, 'มีผู้สมัครงานใหม่: ' + name + ' (' + rec.position + ')',
        'เลขที่ใบสมัคร ' + appNo + '\nตำแหน่ง ' + rec.position + '\nเบอร์โทร ' + rec.mobile +
        '\n\nดูข้อมูลทั้งหมด: ' + getSpreadsheet_().getUrl());
    } catch (e) { console.error(e); }
  }
  return { ok: true, appNo: appNo };
}

function isVisible_(f, ans) {
  const c = f.showIf;
  if (!c) return true;
  const v = ans[c.k];
  if (c.eq !== undefined) return v === c.eq;
  if (c.ne !== undefined) return v !== c.ne;
  if (c.in) return c.in.indexOf(v) > -1;
  if (c.has) return Array.isArray(v) && v.indexOf(c.has) > -1;
  return true;
}

function thaiIdValid_(id) {
  if (!/^\d{13}$/.test(id)) return false;
  let s = 0;
  for (let i = 0; i < 12; i++) s += Number(id[i]) * (13 - i);
  return (11 - s % 11) % 10 === Number(id[12]);
}

/** yyyy-mm-dd → dd/mm/yyyy (พ.ศ.) */
function toThaiDate_(v) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(v);
  if (!m) return v;
  const y = Number(m[1]);
  return m[3] + '/' + m[2] + '/' + (y >= 2400 ? y : y + 543);   // ถ้ากรอกเป็น พ.ศ. มาแล้วไม่ต้องบวกซ้ำ
}

function personalitySummary_(selected, settings) {
  selected = selected || [];
  if (!selected.length) return '';
  const items = (settings.personality && settings.personality.items) || [];
  const count = {};
  selected.forEach(function (t) {
    const it = items.filter(function (i) { return i.text === t; })[0];
    const g = it && it.group ? it.group : 'ไม่ระบุกลุ่ม';
    count[g] = (count[g] || 0) + 1;
  });
  return Object.keys(count).sort(function (a, b) { return count[b] - count[a]; })
    .map(function (g) { return g + ' ' + count[g] + ' ข้อ'; }).join(', ');
}

/***** Google Sheet *****/
function getSpreadsheet_() {
  if (SPREADSHEET_ID) return SpreadsheetApp.openById(SPREADSHEET_ID);
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('SS_ID');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) { /* ถูกลบ */ } }
  const ss = SpreadsheetApp.create('ระบบใบสมัครงาน - ' + COMPANY_NAME);
  props.setProperty('SS_ID', ss.getId());
  try { DriveApp.getFileById(ss.getId()).moveTo(getFolder_('root')); } catch (e) { /* ignore */ }
  return ss;
}

function getDataSheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(DATA_SHEET);
  if (!sh) {
    sh = ss.insertSheet(DATA_SHEET, 0);
  }
  return sh;
}

function allHeaders_() {
  return FIXED_HEAD.concat(getColumns_().map(function (c) { return c.h; })).concat(FIXED_TAIL);
}

function ensureHeaders_(sheet) {
  const want = allHeaders_();
  let cur = sheet.getLastColumn() ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String) : [];
  if (!cur.filter(String).length) cur = [];
  const missing = want.filter(function (h) { return cur.indexOf(h) === -1; });
  if (missing.length) {
    sheet.getRange(1, cur.length + 1, 1, missing.length).setValues([missing]);
    cur = cur.concat(missing);
    sheet.getRange(1, 1, 1, cur.length).setFontWeight('bold').setBackground('#1f3a5f').setFontColor('#ffffff').setWrap(true).setVerticalAlignment('middle');
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(2);
  }
  return cur;
}

function buildRow_(headers, rec, fileLinks, appNo, ts) {
  const byHead = {};
  getColumns_().forEach(function (c) {
    let v = c.file ? (fileLinks[c.k] || '') : rec[c.k];
    if (Array.isArray(v)) v = v.join(', ');
    byHead[c.h] = v == null ? '' : v;
  });
  byHead['เลขที่ใบสมัคร'] = appNo;
  byHead['วันที่สมัคร'] = ts;
  byHead['สถานะ'] = 'ใหม่';
  return headers.map(function (h) {
    const v = byHead.hasOwnProperty(h) ? byHead[h] : '';
    return (typeof v === 'string' && v !== '') ? "'" + v : v;  // เก็บเป็นข้อความ (เลข 0 นำหน้าไม่หาย)
  });
}

function writeDocLinks_(sheet, headers, rowNo, docs) {
  const iw = headers.indexOf('ใบสมัคร (Google Doc)'), ip = headers.indexOf('ใบสมัคร (PDF)');
  if (iw > -1 && docs.docx) sheet.getRange(rowNo, iw + 1).setValue(docs.docx);
  if (ip > -1) sheet.getRange(rowNo, ip + 1).setValue(docs.pdf || '');
}

/** อ่านข้อมูลผู้สมัครจากแถวในชีท (ใช้ตอนสร้างเอกสารใหม่ หลัง HR แก้ข้อมูลในชีท) */
function recordFromRow_(sheet, rowNo) {
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const disp = sheet.getRange(rowNo, 1, 1, lastCol).getDisplayValues()[0];
  const vals = sheet.getRange(rowNo, 1, 1, lastCol).getValues()[0];
  const rec = {};
  getColumns_().forEach(function (c) {
    const i = headers.indexOf(c.h);
    let v = i > -1 ? disp[i] : '';
    if (c.multi) v = v ? v.split(/\s*,\s*/) : [];
    rec[c.k] = v;
  });
  const ti = headers.indexOf('วันที่สมัคร');
  const ts = ti > -1 && vals[ti] instanceof Date ? vals[ti] : new Date();
  return { rec: rec, ts: ts, appNo: disp[headers.indexOf('เลขที่ใบสมัคร')] || ('ROW' + rowNo), headers: headers };
}

/***** การตั้งค่า (เก็บในชีท "ตั้งค่าระบบ") *****/
function getSettings_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('settings');
  if (hit) return JSON.parse(hit);
  let s = {};
  try {
    const sh = getSpreadsheet_().getSheetByName(SETTINGS_SHEET);
    if (sh) s = JSON.parse(sh.getRange('A1').getValue() || '{}');
  } catch (e) { s = {}; }
  const merged = mergeSettings_(s);
  const json = JSON.stringify(merged);
  if (json.length < 90000) cache.put('settings', json, 600);
  return merged;
}

function mergeSettings_(s) {
  const d = JSON.parse(JSON.stringify(DEFAULT_SETTINGS));
  Object.keys(s || {}).forEach(function (k) {
    if (k === 'lists' || k === 'personality') d[k] = Object.assign(d[k], s[k] || {});
    else d[k] = s[k];
  });
  if (!Array.isArray(d.emailTemplates) || !d.emailTemplates.length) d.emailTemplates = JSON.parse(JSON.stringify(DEFAULT_EMAIL_TEMPLATES));
  return d;
}

function saveSettings_(s) {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(SETTINGS_SHEET);
  if (!sh) { sh = ss.insertSheet(SETTINGS_SHEET); sh.hideSheet(); }
  sh.getRange('A1').setValue(JSON.stringify(s));
  sh.getRange('A2').setValue('ชีทนี้เก็บการตั้งค่าของระบบ กรุณาแก้ไขผ่านหน้าผู้ดูแลเท่านั้น');
  CacheService.getScriptCache().remove('settings');
}

/***** โฟลเดอร์ใน Google Drive *****/
function getFolder_(kind) {
  const props = PropertiesService.getScriptProperties();
  const key = 'FOLDER_' + kind;
  const id = props.getProperty(key);
  if (id) { try { const f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) { /* สร้างใหม่ */ } }
  const names = { root: ROOT_FOLDER_NAME, files: 'ไฟล์แนบผู้สมัคร', docs: 'ใบสมัคร (Word-PDF)', template: 'แม่แบบใบสมัคร' };
  const f = kind === 'root' ? DriveApp.createFolder(names.root) : getFolder_('root').createFolder(names[kind]);
  props.setProperty(key, f.getId());
  return f;
}

/***** API หน้าผู้ดูแล *****/
function checkAdmin_(pw) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('adminFails') || 0);
  if (fails >= 10) throw new Error('ใส่รหัสผิดหลายครั้ง กรุณารอ 10 นาที');
  const real = PropertiesService.getScriptProperties().getProperty('ADMIN_PW') || DEFAULT_ADMIN_PASSWORD;
  if (String(pw) !== real) {
    cache.put('adminFails', String(fails + 1), 600);
    throw new Error('รหัสผ่านไม่ถูกต้อง');
  }
}

function adminLoad(pw) {
  checkAdmin_(pw);
  const s = getSettings_();
  let template = null;
  if (s.templateId) {
    try { const f = DriveApp.getFileById(s.templateId); template = { name: f.getName(), url: f.getUrl() }; } catch (e) { /* หาย */ }
  }
  return {
    settings: s,
    info: {
      sheetUrl: getSpreadsheet_().getUrl(),
      folderUrl: getFolder_('root').getUrl(),
      appUrl: ScriptApp.getService().getUrl(),
      template: template,
      usingDefaultPw: !PropertiesService.getScriptProperties().getProperty('ADMIN_PW'),
      langs: LANGS,
      emailVars: EMAIL_VARS,
      me: Session.getEffectiveUser().getEmail(),
      quota: MailApp.getRemainingDailyQuota()
    }
  };
}

function adminSave(pw, s) {
  checkAdmin_(pw);
  const cur = getSettings_();
  const clean = mergeSettings_(s);
  clean.templateId = cur.templateId;   // แม่แบบเปลี่ยนผ่านการอัปโหลดเท่านั้น
  clean.positions = (clean.positions || []).filter(function (p) { return p && String(p.name || '').trim(); })
    .map(function (p) { return { name: String(p.name).trim(), detail: String(p.detail || ''), open: !!p.open }; });
  Object.keys(clean.lists).forEach(function (k) {
    clean.lists[k] = (clean.lists[k] || []).map(function (x) { return String(x).trim(); }).filter(String);
  });
  clean.personality.items = (clean.personality.items || []).filter(function (i) { return i && String(i.text || '').trim(); })
    .map(function (i) { return { text: String(i.text).trim(), group: String(i.group || '').trim() }; });
  clean.emailTemplates = (clean.emailTemplates || []).filter(function (t) { return t && String(t.name || '').trim(); })
    .map(function (t) { return { name: String(t.name).trim(), subject: String(t.subject || ''), body: String(t.body || '') }; });
  if (!clean.emailTemplates.length) clean.emailTemplates = JSON.parse(JSON.stringify(DEFAULT_EMAIL_TEMPLATES));
  saveSettings_(clean);
  return adminLoad(pw);
}

function adminUploadTemplate(pw, file) {
  checkAdmin_(pw);
  const blob = Utilities.newBlob(Utilities.base64Decode(file.data), MimeType.MICROSOFT_WORD, file.name);
  const xml = readDocXml_(blob);
  if (xml.indexOf('{{fullname|') === -1) throw new Error('ไฟล์นี้ไม่ใช่แม่แบบที่มีช่องข้อมูล {{...}} กรุณาใช้ไฟล์ ใบสมัครงาน_TEMPLATE.docx');
  const f = getFolder_('template').createFile(blob);
  const s = getSettings_();
  if (s.templateId) { try { DriveApp.getFileById(s.templateId).setTrashed(true); } catch (e) { /* ignore */ } }
  s.templateId = f.getId();
  saveSettings_(s);
  return adminLoad(pw);
}

function adminChangePassword(pw, newPw) {
  checkAdmin_(pw);
  if (!newPw || String(newPw).length < 6) throw new Error('รหัสผ่านใหม่ต้องมีอย่างน้อย 6 ตัวอักษร');
  PropertiesService.getScriptProperties().setProperty('ADMIN_PW', String(newPw));
  return true;
}

function adminListApplicants(pw) {
  checkAdmin_(pw);
  const sh = getDataSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const from = Math.max(2, last - 499);
  const data = sh.getRange(from, 1, last - from + 1, headers.length).getDisplayValues();
  const col = function (h) { return headers.indexOf(h); };
  const c = {
    no: col('เลขที่ใบสมัคร'), date: col('วันที่สมัคร'), pos: col('ตำแหน่งที่ต้องการสมัคร'),
    t: col('คำนำหน้า'), fn: col('ชื่อ'), ln: col('นามสกุล'), mob: col('เบอร์มือถือ'), st: col('สถานะ'),
    w: col('ใบสมัคร (Google Doc)'), p: col('ใบสมัคร (PDF)'), em: col('อีเมล'), nk: col('ชื่อเล่น'), iv: col('วันนัดสัมภาษณ์')
  };
  const g = function (r, i) { return i > -1 ? r[i] : ''; };
  return data.map(function (r, i) {
    return {
      row: from + i, appNo: g(r, c.no), date: g(r, c.date), position: g(r, c.pos),
      name: (g(r, c.t) + g(r, c.fn) + ' ' + g(r, c.ln)).trim(), mobile: g(r, c.mob),
      status: g(r, c.st), docx: g(r, c.w), pdf: g(r, c.p),
      email: g(r, c.em), first: g(r, c.fn), nick: g(r, c.nk), interview: g(r, c.iv)
    };
  }).filter(function (x) { return x.appNo; }).reverse();
}

function adminGenerate(pw, rowNo) {
  checkAdmin_(pw);
  const sh = getDataSheet_();
  const r = recordFromRow_(sh, rowNo);
  const docs = generateDocs_(r.rec, r.ts, r.appNo, getSettings_());
  writeDocLinks_(sh, r.headers, rowNo, docs);
  return docs;
}

function adminSetStatus(pw, rowNo, status) {
  checkAdmin_(pw);
  const sh = getDataSheet_();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const i = headers.indexOf('สถานะ');
  if (i > -1) sh.getRange(rowNo, i + 1).setValue(status);
  return true;
}
