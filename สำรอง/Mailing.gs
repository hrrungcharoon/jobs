/**
 * ส่งอีเมลหาผู้สมัคร (Email Merge) — ใส่ชื่อ/ตำแหน่ง/วันนัด ของแต่ละคนอัตโนมัติ
 * ตัวแปรที่ใช้ในหัวเรื่อง/เนื้อความ: ดู EMAIL_VARS
 */
const LOG_SHEET = 'ประวัติการส่งอีเมล';
const EMAIL_VARS = ['ชื่อ-สกุล', 'ชื่อ', 'ชื่อเล่น', 'ตำแหน่ง', 'เลขที่ใบสมัคร', 'วันที่สมัคร', 'บริษัท',
  'ชื่อนามสกุล', 'วันนัด', 'วันที่สั้น', 'เวลา', 'สถานที่', 'ผู้สัมภาษณ์', 'ผู้ติดต่อ', 'เบอร์ติดต่อ', 'ที่อยู่บริษัท', 'อีเมลติดต่อ', 'LINE'];

const DEFAULT_EMAIL_TEMPLATES = [
  { name: 'ได้รับใบสมัครแล้ว', subject: 'ได้รับใบสมัครของคุณแล้ว ({{เลขที่ใบสมัคร}}) - {{บริษัท}}',
    body: 'เรียน {{ชื่อ-สกุล}}\n\nบริษัทฯ ได้รับใบสมัครงานตำแหน่ง {{ตำแหน่ง}} ของท่านเรียบร้อยแล้ว เลขที่ใบสมัคร {{เลขที่ใบสมัคร}}\nฝ่ายบุคคลจะพิจารณาและติดต่อกลับหากคุณสมบัติตรงกับตำแหน่งงาน\n\nขอแสดงความนับถือ\nฝ่ายบุคคล {{บริษัท}}' },
  { name: 'นัดสัมภาษณ์', subject: 'นัดสัมภาษณ์งานตำแหน่ง {{ตำแหน่ง}} วันที่ {{วันที่สั้น}} เวลา {{เวลา}}',
    body: 'เรียน คุณ {{ชื่อนามสกุล}}\nขอเชิญเข้าร่วมสัมภาษณ์งานกับ{{บริษัท}}\nตามรายละเอียดด้านล่าง\n\nตำแหน่งที่เสนอ: {{ตำแหน่ง}}\nวันที่: {{วันนัด}}\nเวลา: {{เวลา}}\nสถานที่: {{สถานที่}}\nผู้สัมภาษณ์: {{ผู้สัมภาษณ์}}\nโปรดยืนยันการตอบรับการสัมภาษณ์ของคุณด้วยการตอบกลับอีเมลนี้\nติดต่อ: {{ผู้ติดต่อ}}\nที่อยู่: {{ที่อยู่บริษัท}}\nโทร.: {{เบอร์ติดต่อ}}\nอีเมล: {{อีเมลติดต่อ}}\nLINE: {{LINE}}' },
  { name: 'แจ้งผลผ่านการคัดเลือก', subject: 'แจ้งผลการสมัครงานตำแหน่ง {{ตำแหน่ง}} - {{บริษัท}}',
    body: 'เรียน {{ชื่อ-สกุล}}\n\nบริษัทฯ ยินดีแจ้งให้ทราบว่าท่านผ่านการคัดเลือกในตำแหน่ง {{ตำแหน่ง}}\nขอให้ท่านมารายงานตัวในวันที่ {{วันนัด}} เวลา {{เวลา}} ณ {{สถานที่}}\n\nหากมีข้อสงสัย กรุณาติดต่อ {{ผู้ติดต่อ}} โทร {{เบอร์ติดต่อ}}\n\nขอแสดงความนับถือ\nฝ่ายบุคคล {{บริษัท}}' },
  { name: 'แจ้งผลไม่ผ่านการคัดเลือก', subject: 'แจ้งผลการสมัครงานตำแหน่ง {{ตำแหน่ง}} - {{บริษัท}}',
    body: 'เรียน {{ชื่อ-สกุล}}\n\nขอบคุณที่ให้ความสนใจสมัครงานกับ {{บริษัท}} ในตำแหน่ง {{ตำแหน่ง}}\nหลังจากพิจารณาแล้ว บริษัทฯ เสียใจที่ต้องแจ้งว่าในครั้งนี้ยังไม่สามารถรับท่านเข้าทำงานได้\nบริษัทฯ จะเก็บข้อมูลของท่านไว้พิจารณาหากมีตำแหน่งที่เหมาะสมในอนาคต\n\nขอแสดงความนับถือ\nฝ่ายบุคคล {{บริษัท}}' }
];

function fillEmail_(text, ctx) {
  return String(text || '').replace(/\{\{\s*([^}]+?)\s*\}\}/g, function (m, k) {
    return ctx.hasOwnProperty(k) && ctx[k] !== '' ? ctx[k] : (EMAIL_VARS.indexOf(k) > -1 ? '-' : m);
  });
}

function emailHtml_(text) {
  const esc = String(text).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\n/g, '<br>');
  return '<div style="font-family:Tahoma,Sarabun,sans-serif;font-size:15px;line-height:1.7;color:#1d2433">' + esc + '</div>';
}

/** ข้อมูลผู้สมัครแต่ละแถว สำหรับแทนค่าในอีเมล */
function emailContextFromRow_(headers, row, vars) {
  const g = function (h) { const i = headers.indexOf(h); return i > -1 ? String(row[i] || '').trim() : ''; };
  const ctx = {
    'ชื่อ-สกุล': (g('คำนำหน้า') + g('ชื่อ') + ' ' + g('นามสกุล')).trim(),
    'ชื่อ': g('ชื่อ'), 'ชื่อนามสกุล': (g('ชื่อ') + ' ' + g('นามสกุล')).trim(), 'ชื่อเล่น': g('ชื่อเล่น'), 'ตำแหน่ง': g('ตำแหน่งที่ต้องการสมัคร'),
    'เลขที่ใบสมัคร': g('เลขที่ใบสมัคร'), 'วันที่สมัคร': g('วันที่สมัคร').split(' ')[0], 'บริษัท': COMPANY_NAME
  };
  Object.keys(vars || {}).forEach(function (k) { ctx[k] = vars[k]; });
  return { ctx: ctx, email: g('อีเมล') };
}

function sendOne_(to, subject, body, settings) {
  MailApp.sendEmail({
    to: to, subject: subject, body: body, htmlBody: emailHtml_(body),
    name: settings.senderName || ('ฝ่ายบุคคล ' + COMPANY_NAME),
    replyTo: settings.replyTo || undefined
  });
}

function logEmail_(rows) {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(LOG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(LOG_SHEET);
    sh.appendRow(['เวลาที่ส่ง', 'เลขที่ใบสมัคร', 'ชื่อ-สกุล', 'อีเมล', 'หัวเรื่อง', 'ผล']);
    sh.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#1f3a5f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
}

/** อีเมลตอบรับอัตโนมัติหลังผู้สมัครส่งใบสมัคร */
function sendAutoReply_(rowNo, settings) {
  if (!settings.autoReplyEnabled) return;
  const tpl = (settings.emailTemplates || []).filter(function (t) { return t.name === settings.autoReplyTemplate; })[0];
  if (!tpl) return;
  const sh = getDataSheet_();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const row = sh.getRange(rowNo, 1, 1, headers.length).getDisplayValues()[0];
  const x = emailContextFromRow_(headers, row, {});
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(x.email)) return;
  const subject = fillEmail_(tpl.subject, x.ctx);
  try {
    sendOne_(x.email, subject, fillEmail_(tpl.body, x.ctx), settings);
    logEmail_([[new Date(), x.ctx['เลขที่ใบสมัคร'], x.ctx['ชื่อ-สกุล'], x.email, subject, 'ส่งแล้ว (อัตโนมัติ)']]);
  } catch (e) {
    logEmail_([[new Date(), x.ctx['เลขที่ใบสมัคร'], x.ctx['ชื่อ-สกุล'], x.email, subject, 'ไม่สำเร็จ: ' + e.message]]);
  }
}

/**
 * req = { rows:[เลขแถว], subject, body, vars:{วันนัด,...}, setStatus:'', test:true/false }
 */
function adminSendEmails(pw, req) {
  checkAdmin_(pw);
  const settings = getSettings_();
  const rows = (req.rows || []).map(Number).filter(function (n) { return n >= 2; });
  if (!rows.length) throw new Error('ยังไม่ได้เลือกผู้สมัคร');
  if (!String(req.subject || '').trim() || !String(req.body || '').trim()) throw new Error('กรุณาใส่หัวเรื่องและเนื้อความ');
  const sh = getDataSheet_();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const stCol = headers.indexOf('สถานะ');

  if (req.test) {
    const me = Session.getEffectiveUser().getEmail();
    const row = sh.getRange(rows[0], 1, 1, headers.length).getDisplayValues()[0];
    const x = emailContextFromRow_(headers, row, req.vars);
    sendOne_(me, '[ทดสอบ] ' + fillEmail_(req.subject, x.ctx), fillEmail_(req.body, x.ctx), settings);
    return { sent: 1, test: true, to: me, skipped: [], quotaLeft: MailApp.getRemainingDailyQuota() };
  }

  const quota = MailApp.getRemainingDailyQuota();
  if (rows.length > quota) throw new Error('โควตาส่งอีเมลวันนี้เหลือ ' + quota + ' ฉบับ แต่เลือกไว้ ' + rows.length + ' คน');
  const log = [], skipped = [];
  let sent = 0;
  rows.forEach(function (rowNo) {
    const row = sh.getRange(rowNo, 1, 1, headers.length).getDisplayValues()[0];
    const x = emailContextFromRow_(headers, row, req.vars);
    const subject = fillEmail_(req.subject, x.ctx);
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(x.email)) {
      skipped.push({ appNo: x.ctx['เลขที่ใบสมัคร'], name: x.ctx['ชื่อ-สกุล'], reason: 'ไม่มีอีเมล' });
      return;
    }
    try {
      sendOne_(x.email, subject, fillEmail_(req.body, x.ctx), settings);
      sent++;
      log.push([new Date(), x.ctx['เลขที่ใบสมัคร'], x.ctx['ชื่อ-สกุล'], x.email, subject, 'ส่งแล้ว']);
      if (req.setStatus && stCol > -1) sh.getRange(rowNo, stCol + 1).setValue(req.setStatus);
    } catch (e) {
      skipped.push({ appNo: x.ctx['เลขที่ใบสมัคร'], name: x.ctx['ชื่อ-สกุล'], reason: e.message });
      log.push([new Date(), x.ctx['เลขที่ใบสมัคร'], x.ctx['ชื่อ-สกุล'], x.email, subject, 'ไม่สำเร็จ: ' + e.message]);
    }
  });
  logEmail_(log);
  return { sent: sent, skipped: skipped, quotaLeft: MailApp.getRemainingDailyQuota() };
}
