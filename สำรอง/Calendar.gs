/**
 * นัดสัมภาษณ์ → Google Calendar
 * HR เลือกผู้สมัครในหน้าผู้ดูแล ใส่วัน เวลา สถานที่ → ระบบสร้างนัดในปฏิทินของบัญชีนี้
 * (เลือกหลายคน = จัดคิวต่อกันอัตโนมัติ) + ส่งคำเชิญปฏิทิน/อีเมลนัดสัมภาษณ์ให้ผู้สมัคร
 *
 * รูปแบบนัด:
 *   หัวเรื่อง: นัดสัมภาษณ์งานตำแหน่ง GRAPHIC วันที่ 7/11/2569 เวลา 10.00น.
 *   รายละเอียด: เรียน คุณ ... / ตำแหน่งที่เสนอ / วันที่ / เวลา / สถานที่ / ผู้สัมภาษณ์ / ติดต่อ / ที่อยู่ / โทร / อีเมล / LINE
 */
const TH_MONTHS = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];

function thaiDate_(d) { return d.getDate() + ' ' + TH_MONTHS[d.getMonth()] + ' ' + (d.getFullYear() + 543); }        // 7 พฤศจิกายน 2569
function thaiShortDate_(d) { return d.getDate() + '/' + (d.getMonth() + 1) + '/' + (d.getFullYear() + 543); }          // 7/11/2569
function thaiTime_(d) { return d.getHours() + '.' + ('0' + d.getMinutes()).slice(-2); }                                // 9.00

/** รายละเอียดในนัดปฏิทิน */
function interviewDescription_(c, extra) {
  return [
    'เรียน คุณ ' + c['ชื่อนามสกุล'],
    'ขอเชิญเข้าร่วมสัมภาษณ์งานกับ' + COMPANY_NAME,
    'ตามรายละเอียดด้านล่าง',
    '',
    'ตำแหน่งที่เสนอ: ' + c['ตำแหน่ง'],
    'วันที่: ' + c['วันนัด'],
    'เวลา: ' + c['เวลา'],
    'สถานที่: ' + c['สถานที่'],
    c['ผู้สัมภาษณ์'] ? 'ผู้สัมภาษณ์: ' + c['ผู้สัมภาษณ์'] : null,
    'โปรดยืนยันการตอบรับการสัมภาษณ์ของคุณด้วยการตอบกลับอีเมลนี้',
    c['ผู้ติดต่อ'] ? 'ติดต่อ: ' + c['ผู้ติดต่อ'] : null,
    c['ที่อยู่บริษัท'] ? 'ที่อยู่: ' + c['ที่อยู่บริษัท'] : null,
    c['เบอร์ติดต่อ'] ? 'โทร.: ' + c['เบอร์ติดต่อ'] : null,
    c['อีเมลติดต่อ'] ? 'อีเมล: ' + c['อีเมลติดต่อ'] : null,
    c['LINE'] ? 'LINE: ' + c['LINE'] : null
  ].filter(function (x) { return x !== null; }).join('\n') + (extra || '');
}

/**
 * req = { rows:[เลขแถว], date:'yyyy-mm-dd', time:'HH:mm', minutes:30, location, interviewer, contact, contactPhone,
 *         address, contactEmail, line, note, interviewers:'a@x.com, b@y.com', invite:true, sendEmail:true,
 *         templateName:'นัดสัมภาษณ์', setStatus:'นัดสัมภาษณ์' }
 */
function adminScheduleInterviews(pw, req) {
  checkAdmin_(pw);
  const rows = (req.rows || []).map(Number).filter(function (n) { return n >= 2; });
  if (!rows.length) throw new Error('ยังไม่ได้เลือกผู้สมัคร');
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(req.date || ''), t = /^(\d{1,2}):(\d{2})$/.exec(req.time || '');
  if (!m || !t) throw new Error('กรุณาเลือกวันที่และเวลาสัมภาษณ์');
  const minutes = Math.max(5, Number(req.minutes) || 30);
  let y = Number(m[1]); if (y > 2400) y -= 543;
  const first = new Date(y, Number(m[2]) - 1, Number(m[3]), Number(t[1]), Number(t[2]));

  const settings = getSettings_();
  // จำข้อมูลติดต่อ/สถานที่ล่าสุดไว้ใช้ครั้งหน้า
  const info = {
    location: String(req.location || '').trim(), interviewer: String(req.interviewer || '').trim(),
    contact: String(req.contact || '').trim(), phone: String(req.contactPhone || '').trim(),
    address: String(req.address || '').trim(), email: String(req.contactEmail || '').trim(), line: String(req.line || '').trim()
  };
  settings.interviewInfo = info;
  saveSettings_(settings);

  const tpl = req.sendEmail ? (settings.emailTemplates || []).filter(function (x) { return x.name === req.templateName; })[0] : null;
  if (req.sendEmail && !tpl) throw new Error('ไม่พบแม่แบบอีเมล "' + req.templateName + '"');
  if (req.sendEmail && rows.length > MailApp.getRemainingDailyQuota()) throw new Error('โควตาอีเมลวันนี้ไม่พอ');
  const interviewers = String(req.interviewers || '').split(/[,\s;]+/).filter(function (e) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e); });

  const cal = CalendarApp.getDefaultCalendar();
  const sh = getDataSheet_();
  ensureHeaders_(sh);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const col = function (h) { return headers.indexOf(h) + 1; };
  const results = [], log = [];

  rows.forEach(function (rowNo, i) {
    const start = new Date(first.getTime() + i * minutes * 60000);
    const end = new Date(start.getTime() + minutes * 60000);
    const row = sh.getRange(rowNo, 1, 1, headers.length).getDisplayValues()[0];
    const x = emailContextFromRow_(headers, row, {
      'วันนัด': thaiDate_(start), 'วันที่สั้น': thaiShortDate_(start), 'เวลา': thaiTime_(start) + ' น.',
      'สถานที่': info.location, 'ผู้สัมภาษณ์': info.interviewer, 'ผู้ติดต่อ': info.contact, 'เบอร์ติดต่อ': info.phone,
      'ที่อยู่บริษัท': info.address, 'อีเมลติดต่อ': info.email, 'LINE': info.line
    });
    const name = x.ctx['ชื่อนามสกุล'] || x.ctx['เลขที่ใบสมัคร'];
    const when = x.ctx['วันนัด'] + ' เวลา ' + x.ctx['เวลา'];
    const hasEmail = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(x.email);
    const inviteApplicant = req.invite && hasEmail;
    try {
      // ยกเลิกนัดเดิม (กรณีเลื่อนนัด)
      const oldId = col('รหัสนัดหมาย (ปฏิทิน)') ? String(row[col('รหัสนัดหมาย (ปฏิทิน)') - 1] || '') : '';
      if (oldId) { try { const ev = cal.getEventById(oldId); if (ev) ev.deleteEvent(); } catch (e) { /* ไม่พบนัดเดิม */ } }

      // ข้อมูลสำหรับ HR — ใส่เฉพาะเมื่อไม่ได้เชิญผู้สมัคร (ผู้ถูกเชิญมองเห็นรายละเอียดนัด)
      const docLink = col('ใบสมัคร (Google Doc)') ? row[col('ใบสมัคร (Google Doc)') - 1] : '';
      let extra = req.note ? '\n\n' + req.note : '';
      if (!inviteApplicant) extra += '\n\n[สำหรับ HR] เลขที่ใบสมัคร ' + x.ctx['เลขที่ใบสมัคร'] + (docLink ? '\nใบสมัคร: ' + docLink : '');

      const title = 'นัดสัมภาษณ์งานตำแหน่ง ' + x.ctx['ตำแหน่ง'] + ' วันที่ ' + x.ctx['วันที่สั้น'] + ' เวลา ' + thaiTime_(start) + 'น.';
      const guests = interviewers.concat(inviteApplicant ? [x.email] : []);
      const ev = cal.createEvent(title, start, end, {
        description: interviewDescription_(x.ctx, extra), location: info.location,
        guests: guests.join(','), sendInvites: !!req.invite
      });
      ev.addPopupReminder(30);

      if (col('วันนัดสัมภาษณ์')) sh.getRange(rowNo, col('วันนัดสัมภาษณ์')).setValue("'" + when);
      if (col('รหัสนัดหมาย (ปฏิทิน)')) sh.getRange(rowNo, col('รหัสนัดหมาย (ปฏิทิน)')).setValue(ev.getId());
      if (req.setStatus && col('สถานะ')) sh.getRange(rowNo, col('สถานะ')).setValue(req.setStatus);

      let mail = '';
      if (tpl) {
        if (hasEmail) {
          const subject = fillEmail_(tpl.subject, x.ctx);
          sendOne_(x.email, subject, fillEmail_(tpl.body, x.ctx), settings);
          log.push([new Date(), x.ctx['เลขที่ใบสมัคร'], name, x.email, subject, 'ส่งแล้ว (นัดสัมภาษณ์)']);
          mail = 'ส่งอีเมลแล้ว';
        } else mail = 'ไม่มีอีเมล';
      }
      results.push({ name: name, when: when, ok: true, mail: mail, invited: inviteApplicant });
    } catch (e) {
      results.push({ name: name, when: when, ok: false, error: e.message });
    }
  });
  if (log.length) logEmail_(log);
  return { results: results, calendarUrl: 'https://calendar.google.com/calendar/r/day/' + first.getFullYear() + '/' + (first.getMonth() + 1) + '/' + first.getDate() };
}

/** ยกเลิกนัดสัมภาษณ์ */
function adminCancelInterview(pw, rowNo) {
  checkAdmin_(pw);
  const sh = getDataSheet_();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const ci = headers.indexOf('รหัสนัดหมาย (ปฏิทิน)'), di = headers.indexOf('วันนัดสัมภาษณ์');
  if (ci < 0) return false;
  const id = String(sh.getRange(rowNo, ci + 1).getValue() || '');
  if (id) { try { const ev = CalendarApp.getDefaultCalendar().getEventById(id); if (ev) ev.deleteEvent(); } catch (e) { /* ignore */ } }
  sh.getRange(rowNo, ci + 1).setValue('');
  if (di > -1) sh.getRange(rowNo, di + 1).setValue('');
  return true;
}
