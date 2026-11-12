/**
 * นัดสัมภาษณ์ → Google Calendar
 * HR เลือกผู้สมัครในหน้าผู้ดูแล ใส่วัน เวลา สถานที่ → ระบบสร้างนัดในปฏิทินของบัญชีนี้
 * (เลือกหลายคน = จัดคิวต่อกันอัตโนมัติ) + ส่งคำเชิญปฏิทิน/อีเมลนัดสัมภาษณ์ให้ผู้สมัคร
 */
const TH_DAYS = ['อาทิตย์', 'จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์'];
const TH_MONTHS = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];

function thaiLongDate_(d) {
  return 'วัน' + TH_DAYS[d.getDay()] + 'ที่ ' + d.getDate() + ' ' + TH_MONTHS[d.getMonth()] + ' ' + (d.getFullYear() + 543);
}
function hhmm_(d) { return Utilities.formatDate(d, TZ, 'HH:mm') + ' น.'; }

/**
 * req = { rows:[เลขแถว], date:'yyyy-mm-dd', time:'HH:mm', minutes:30, location:'', interviewers:'a@x.com, b@y.com',
 *         invite:true, sendEmail:true, templateName:'นัดสัมภาษณ์', contact:'', contactPhone:'', setStatus:'นัดสัมภาษณ์' }
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
      'วันนัด': thaiLongDate_(start), 'เวลา': hhmm_(start) + ' - ' + hhmm_(end), 'สถานที่': req.location || '',
      'ผู้ติดต่อ': req.contact || '', 'เบอร์ติดต่อ': req.contactPhone || ''
    });
    const name = x.ctx['ชื่อ-สกุล'] || x.ctx['เลขที่ใบสมัคร'];
    const hasEmail = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(x.email);
    const inviteApplicant = req.invite && hasEmail;
    try {
      // ยกเลิกนัดเดิม (กรณีเลื่อนนัด)
      const oldId = col('รหัสนัดหมาย (ปฏิทิน)') ? String(row[col('รหัสนัดหมาย (ปฏิทิน)') - 1] || '') : '';
      if (oldId) { try { const ev = cal.getEventById(oldId); if (ev) ev.deleteEvent(); } catch (e) { /* ไม่พบนัดเดิม */ } }

      const docLink = col('ใบสมัคร (Google Doc)') ? row[col('ใบสมัคร (Google Doc)') - 1] : '';
      const desc = [
        'สัมภาษณ์งานตำแหน่ง: ' + x.ctx['ตำแหน่ง'],
        'ผู้สมัคร: ' + name + '  (เลขที่ใบสมัคร ' + x.ctx['เลขที่ใบสมัคร'] + ')',
        req.location ? 'สถานที่: ' + req.location : '',
        req.contact ? 'ติดต่อ: ' + req.contact + (req.contactPhone ? ' โทร ' + req.contactPhone : '') : '',
        // ลิงก์ใบสมัครใส่เฉพาะเมื่อไม่ได้เชิญผู้สมัคร (ผู้ถูกเชิญมองเห็นรายละเอียดนัด)
        !inviteApplicant && docLink ? '\nใบสมัคร: ' + docLink : '',
        req.note ? '\n' + req.note : ''
      ].filter(String).join('\n');
      const guests = interviewers.concat(inviteApplicant ? [x.email] : []);
      const ev = cal.createEvent('สัมภาษณ์งาน: ' + name + ' (' + x.ctx['ตำแหน่ง'] + ')', start, end, {
        description: desc, location: req.location || '', guests: guests.join(','), sendInvites: !!req.invite
      });
      ev.addPopupReminder(30);

      if (col('วันนัดสัมภาษณ์')) sh.getRange(rowNo, col('วันนัดสัมภาษณ์')).setValue("'" + x.ctx['วันนัด'] + ' ' + x.ctx['เวลา']);
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
      results.push({ name: name, when: x.ctx['วันนัด'] + ' ' + x.ctx['เวลา'], ok: true, mail: mail, invited: inviteApplicant });
    } catch (e) {
      results.push({ name: name, when: x.ctx['วันนัด'] + ' ' + x.ctx['เวลา'], ok: false, error: e.message });
    }
  });
  if (log.length) logEmail_(log);
  return { results: results, calendarUrl: 'https://calendar.google.com/calendar/r/day/' + Utilities.formatDate(first, TZ, 'yyyy/M/d') };
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
