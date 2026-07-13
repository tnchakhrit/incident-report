/**
 * ============================================================
 *  Incident Report — Google Apps Script Backend
 *  วางโค้ดนี้ใน Google Apps Script แล้ว Deploy เป็น Web App
 * ============================================================
 *
 *  ตั้งค่าก่อน Deploy (แก้ค่าในบล็อก CONFIG ด้านล่าง):
 *    SHEET_ID        — ID ของ Google Spreadsheet
 *    REVIEWER_EMAIL  — อีเมล QSHE Admin
 *    FORM_ID         — ID ของ Google Form สำหรับ Review/Approve
 *    FORM_ENTRY_REPORTNO  — entry ID ช่อง "เลขที่รายงาน" ใน Form
 *    FORM_ENTRY_ROLE      — entry ID ช่อง "บทบาท" ใน Form
 *
 *  วิธีหา entry ID: เปิด Form → ลิงค์ Preview → กรอกข้อมูล 1 ครั้ง
 *  แล้วดู URL ที่ redirect มา จะเห็น entry.XXXXXXXX=...
 * ============================================================
 */

// ─── CONFIG ───────────────────────────────────────────────
const CONFIG = {
  SHEET_ID: '1BlPrmdt4pYq0coB6PRRAh85lm6zm_2xA42UBLv3lZh4',
  DATA_SHEET_NAME: 'Incident Reports',        // ชื่อ Tab เก็บข้อมูล
  REVIEWER_EMAIL: 'tn.chakhrit@gmail.com',    // ← แก้ตรงนี้ (chakhrit.th@primo.co.th)
  REPORT_PREFIX: 'INC',                       // prefix เลขที่รายงาน

  // ─── Google Form สำหรับ Review/Approve ───
  FORM_ID:             '1FAIpQLScX9se6MqxBjQSWA3cJyEz4OmuZxm7KUK06L-j9f7h4TD6faA',
  FORM_ENTRY_REPORTNO:    'entry.493192896',   // ช่อง "เลขที่รายงาน"
  FORM_ENTRY_ROLE:        'entry.1900944920',  // ช่อง "บทบาทของคุณในรายงานนี้"
  FORM_ENTRY_REJECT_EMAIL:'entry.1253625947',  // ช่อง "Email ผู้ที่ต้องแก้ไข (กรณี Reject)"
};

// ─── CC LIST (รับ 1 ฉบับ FYI ต่อ case ไม่แสดงใน Frontend) ───
// หากใครใน list นี้เป็น Reviewer หรือ Approver ของ case นั้น
// จะถูกกรองออกอัตโนมัติ (เพราะได้รับอีเมลโดยตรงอยู่แล้ว)
const CC_EMAILS_LIST = [
  'chanida.m@primo.co.th',
  'prasit.c@primo.co.th',
  'chakhrit.th@primo.co.th',
  'panwadee.m@primo.co.th',
  'khathayut.w@primo.co.th',
  'nilobon.ch@primo.co.th',
  'sukanya.su@primo.co.th',
  'jomkhwan.s@primo.co.th',
  'kokiat.h@primo.co.th',
  'padipran.s@primo.co.th',
  'rungsarn.d@primo.co.th',
  'prompong.w@primo.co.th',
  'arnon.d@primo.co.th',
  'nirun.j@primo.co.th',
  'aittirit.r@primo.co.th',
];
// ──────────────────────────────────────────────────────────

// ─── FORM LINK HELPER ─────────────────────────────────────

/**
 * สร้างลิงค์ Google Form พร้อม prefill เลขรายงาน + บทบาท
 */
function buildFormUrl(reportNo, role) {
  if (!CONFIG.FORM_ID || CONFIG.FORM_ID === 'YOUR_GOOGLE_FORM_ID') return null;
  const base = `https://docs.google.com/forms/d/e/${CONFIG.FORM_ID}/viewform`;
  const params = [
    `usp=pp_url`,
    `${CONFIG.FORM_ENTRY_REPORTNO}=${encodeURIComponent(reportNo)}`,
    `${CONFIG.FORM_ENTRY_ROLE}=${encodeURIComponent(role)}`,
  ].join('&');
  return `${base}?${params}`;
}

// ─── FORM SUBMIT TRIGGER ───────────────────────────────────

/**
 * ฟังก์ชันนี้ต้องผูก Trigger: Extensions → Apps Script → Triggers
 * → Add Trigger → onFormSubmit → From spreadsheet → On form submit
 *
 * เมื่อ Reviewer/Approver กด Submit Form → อัปเดต Sheet + ส่งอีเมลแจ้ง QSHE
 */
function onFormSubmit(e) {
  try {
    const resp        = e.namedValues;
    const reportNo    = (resp['เลขที่รายงาน']                                    || [''])[0].trim();
    const role        = (resp['บทบาทของคุณในรายงานนี้']                          || [''])[0].trim();
    const result      = (resp['ผลการตรวจสอบ']                                    || [''])[0].trim();
    const comment     = (resp['หมายเหตุ / ข้อเสนอแนะ']                          || [''])[0].trim();
    const rejectEmail = (resp['กรุณากรอก Email ผู้ที่ต้องแก้ไข (กรณี Reject เท่านั้น)'] || [''])[0].trim();
    const submitter   = e.response.getRespondentEmail ? e.response.getRespondentEmail() : '';

    if (!reportNo) return;

    const sheet = getOrCreateSheet();
    const data  = sheet.getDataRange().getValues();
    const now   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');

    // หาแถวที่ตรงกับ reportNo (column 1 = index 0)
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() !== reportNo) continue;

      const statusCol = 29; // column "สถานะ" (1-based) ✅ แก้แล้ว

      // แปลผลเป็นสถานะภาษาไทย
      // ✅ เช็ค 'ไม่ผ่าน' ก่อน เพราะ 'ไม่ผ่าน'.includes('ผ่าน') === true
      let newStatus;
      const isReviewer = role.includes('Review') || role.includes('ตรวจสอบ');
      if (result.includes('ไม่ผ่าน')) {
        newStatus = isReviewer ? 'ส่งกลับแก้ไข ❌' : 'ไม่อนุมัติ ❌';
      } else if (result.includes('ผ่าน')) {
        newStatus = isReviewer ? 'ตรวจสอบแล้ว ✅' : 'อนุมัติแล้ว ✅';
      } else {
        newStatus = 'รอข้อมูลเพิ่มเติม ⏸️';
      }

      // อัปเดตสถานะใน Sheet
      sheet.getRange(i + 1, statusCol).setValue(newStatus);

      // เพิ่ม comment + วันที่ในคอลัมน์ถัดจากสถานะ (ถ้ามี)
      const noteText = `[${now}] ${role}: ${result}${comment ? '\nหมายเหตุ: ' + comment : ''}${submitter ? '\nโดย: ' + submitter : ''}`;
      const noteCell = sheet.getRange(i + 1, statusCol);
      const existing = noteCell.getNote() || '';
      noteCell.setNote(existing ? existing + '\n\n' + noteText : noteText);

      // ส่งอีเมลแจ้ง QSHE Admin
      sendUpdateNotification(reportNo, role, result, comment, submitter, data[i]);

      // ถ้า Reject → ส่งอีเมลแจ้งผู้ที่ต้องแก้ไขด้วย
      const isRejected = result.includes('ไม่ผ่าน') || result.includes('Reject');
      if (isRejected && rejectEmail) {
        sendRejectionEmail(reportNo, role, comment, rejectEmail, data[i]);
      }
      break;
    }
  } catch (err) {
    Logger.log('onFormSubmit error: ' + err.toString());
  }
}

/**
 * ส่งอีเมลแจ้ง QSHE เมื่อมีการ Review/Approve
 */
function sendUpdateNotification(reportNo, role, result, comment, submitter, rowData) {
  const project = rowData[3] || '';
  const subject = `[อัปเดต] ${reportNo} — ${role}: ${result}`;

  // ✅ เช็ค 'ไม่ผ่าน' ก่อน เพื่อให้สีถูกต้อง
  const resultColor = result.includes('ไม่ผ่าน') ? '#d62828' : result.includes('ผ่าน') ? '#2d6a4f' : '#e76f00';

  const body = `
<!DOCTYPE html><html lang="th"><head><meta charset="UTF-8">
<style>
  body{font-family:'Sarabun',Arial,sans-serif;background:#f0f4f8;margin:0;padding:20px}
  .card{max-width:560px;margin:0 auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 16px rgba(0,0,0,0.1)}
  .header{background:linear-gradient(135deg,#1a3a5c,#0d2137);color:#fff;padding:22px 28px}
  .header h1{margin:0;font-size:17px}
  .body{padding:24px 28px;font-size:14px}
  .badge{display:inline-block;background:${resultColor};color:#fff;font-weight:700;padding:5px 16px;border-radius:20px;font-size:14px;margin:12px 0}
  table{width:100%;border-collapse:collapse;margin:12px 0;font-size:13px}
  td{padding:9px 10px;border-bottom:1px solid #eee;vertical-align:top}
  td:first-child{font-weight:600;color:#1a3a5c;width:40%}
  .footer{background:#f8f9fa;padding:12px 28px;font-size:11px;color:#aaa;text-align:center}
</style></head><body>
<div class="card">
  <div class="header"><h1>อัปเดตรายงานอุบัติการณ์</h1></div>
  <div class="body">
    <p>มีการดำเนินการกับรายงาน <strong>${reportNo}</strong></p>
    <div class="badge">${role}: ${result}</div>
    <table>
      <tr><td>โครงการ</td><td>${project}</td></tr>
      <tr><td>ดำเนินการโดย</td><td>${submitter || role}</td></tr>
      <tr><td>หมายเหตุ</td><td>${comment || '-'}</td></tr>
    </table>
  </div>
  <div class="footer">ระบบ Incident Report — Design by Chakhrit</div>
</div></body></html>`;

  GmailApp.sendEmail(CONFIG.REVIEWER_EMAIL, subject, '', {
    htmlBody: body,
    name: 'Incident Report System',
  });
}

/**
 * ส่งอีเมลแจ้งผู้รายงานเมื่อถูก Reject
 */
function sendRejectionEmail(reportNo, reviewerRole, comment, toEmail, rowData) {
  const project = rowData[3] || '';
  // ✅ ลบ formUrl dead variable ออกแล้ว
  const subject = `[ต้องแก้ไข] รายงานอุบัติการณ์ ${reportNo} — ส่งกลับจาก ${reviewerRole}`;

  const body = `
<!DOCTYPE html><html lang="th"><head><meta charset="UTF-8">
<style>
  body{font-family:'Sarabun',Arial,sans-serif;background:#f0f4f8;margin:0;padding:20px}
  .card{max-width:560px;margin:0 auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 16px rgba(0,0,0,0.1)}
  .header{background:linear-gradient(135deg,#d62828,#8b0000);color:#fff;padding:22px 28px}
  .header h1{margin:0;font-size:17px}
  .body{padding:24px 28px;font-size:14px;color:#333}
  .report-no{display:inline-block;background:#fde8e8;color:#d62828;font-weight:700;font-size:16px;padding:7px 18px;border-radius:8px;margin:12px 0;letter-spacing:1px}
  .reason-box{background:#fff3f3;border-left:4px solid #d62828;padding:14px 18px;border-radius:6px;margin:16px 0;font-size:14px}
  table{width:100%;border-collapse:collapse;margin:12px 0;font-size:13px}
  td{padding:9px 10px;border-bottom:1px solid #eee;vertical-align:top}
  td:first-child{font-weight:600;color:#1a3a5c;width:40%}
  .btn{display:inline-block;background:#1a3a5c;color:#fff;text-decoration:none;padding:12px 28px;border-radius:8px;font-size:14px;font-weight:700;margin:16px 0}
  .footer{background:#f8f9fa;padding:12px 28px;font-size:11px;color:#aaa;text-align:center}
</style></head><body>
<div class="card">
  <div class="header"><h1>รายงานถูกส่งกลับเพื่อแก้ไข</h1></div>
  <div class="body">
    <p>รายงานอุบัติการณ์ของคุณถูกส่งกลับโดย <strong>${reviewerRole}</strong> เพื่อแก้ไขเพิ่มเติม</p>
    <div class="report-no">[${reportNo}]</div>
    <table>
      <tr><td>โครงการ</td><td>${project}</td></tr>
      <tr><td>ส่งกลับโดย</td><td>${reviewerRole}</td></tr>
    </table>
    ${comment ? `<div class="reason-box"><strong>เหตุผล / สิ่งที่ต้องแก้ไข:</strong><br>${comment}</div>` : ''}
    <p>กรุณากรอกรายงานใหม่พร้อมแก้ไขข้อมูลที่ระบุ แล้วส่งอีกครั้ง</p>
    <a href="https://test-incident-report-28.netlify.app" class="btn">กรอกรายงานใหม่</a>
  </div>
  <div class="footer">ระบบ Incident Report — Design by Chakhrit</div>
</div></body></html>`;

  GmailApp.sendEmail(toEmail, subject, '', {
    htmlBody: body,
    name: 'Incident Report System',
    replyTo: CONFIG.REVIEWER_EMAIL,
  });
}

// ──────────────────────────────────────────────────────────

/**
 * รับ POST request จากฟอร์ม HTML
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getOrCreateSheet();
    const reportNo = generateReportNo(sheet);

    // เพิ่มแถวข้อมูล
    appendRow(sheet, reportNo, data);

    // ส่งอีเมลแจ้ง Reviewer + Approver
    sendNotificationEmails(reportNo, data);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, reportNo: reportNo }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log('Error: ' + err.toString());
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * GET request:
 *  - ?reportNo=INC-2026-07-001  → หน้ารายงานฉบับเต็ม (public, ไม่ต้อง login)
 *  - ไม่มี query                → ping ทดสอบว่า Script ทำงานอยู่
 */
function doGet(e) {
  const reportNo = e && e.parameter && e.parameter.reportNo;
  if (reportNo) {
    return renderFullReport(reportNo.trim());
  }
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'Incident Report API is running.' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * สร้าง URL ไปหน้ารายงานฉบับเต็มของ Web App ที่ deploy อยู่
 */
function buildFullReportUrl(reportNo) {
  return `${ScriptApp.getService().getUrl()}?reportNo=${encodeURIComponent(reportNo)}`;
}

/**
 * escape ข้อความก่อนแทรกลง HTML — จำเป็นเพราะหน้านี้เป็นหน้าเว็บ public
 * ที่ render ข้อมูลซึ่งผู้ใช้กรอกเองมาโดยตรง (กัน stored XSS)
 */
function escapeHtml_(v) {
  return String(v == null ? '' : v)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * แปลง photoUrls string ("url1, url2, ...") → grid รูป thumbnail คลิกขยายได้
 * กรองเฉพาะ http/https กัน javascript: URI แอบมาแทรกใน href/src
 */
function renderPhotoGrid_(photoUrlsStr) {
  const urls = String(photoUrlsStr || '').split(',')
    .map(u => u.trim())
    .filter(u => /^https?:\/\//i.test(u));
  if (!urls.length) return '<p style="color:#aaa;font-size:13px">— ไม่มีรูปภาพประกอบ —</p>';
  return `<div style="display:flex;flex-wrap:wrap;gap:10px">${urls.map((url, i) => `
    <a href="${escapeHtml_(url)}" target="_blank" rel="noopener" style="display:block">
      <img src="${escapeHtml_(url)}" alt="รูปที่ ${i + 1}" style="width:120px;height:120px;object-fit:cover;border-radius:8px;border:1px solid #ddd">
    </a>`).join('')}</div>`;
}

/**
 * หาแถวข้อมูลจาก reportNo แล้ว render เป็นหน้ารายงานฉบับเต็ม (ทุกหัวข้อ)
 */
function renderFullReport(reportNo) {
  const sheet = getOrCreateSheet();
  const values = sheet.getDataRange().getValues();
  const headers = values[0];

  let row = null;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === reportNo) { row = values[i]; break; }
  }

  if (!row) {
    return HtmlService.createHtmlOutput(
      `<p style="font-family:sans-serif;padding:40px;text-align:center">ไม่พบรายงานเลขที่ "${escapeHtml_(reportNo)}"</p>`
    ).setTitle('ไม่พบรายงาน');
  }

  const rec = {};
  headers.forEach((h, i) => {
    const v = row[i];
    // Sheets จะแปลงค่าที่หน้าตาเหมือนวันที่ (เช่นจาก input type=date/time) เป็น Date object เอง
    rec[h] = (v instanceof Date) ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') : v;
  });

  const section = (title, rows) => `
    <h3 style="color:#1a3a5c;border-bottom:2px solid #1a3a5c;padding-bottom:6px;margin:24px 0 4px">${title}</h3>
    <table style="width:100%;border-collapse:collapse;margin:12px 0;font-size:14px">
      ${rows.filter(r => r).join('')}
    </table>`;
  const tr = (label, value) => value
    ? `<tr><td style="padding:9px 10px;border-bottom:1px solid #eee;vertical-align:top;font-weight:600;color:#1a3a5c;width:35%">${label}</td><td style="padding:9px 10px;border-bottom:1px solid #eee;vertical-align:top;white-space:pre-wrap">${escapeHtml_(value)}</td></tr>`
    : '';

  const html = `
<!DOCTYPE html><html lang="th"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<style>
  body{font-family:'Sarabun',Arial,sans-serif;background:#f0f4f8;margin:0;padding:20px}
  .card{max-width:760px;margin:0 auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 16px rgba(0,0,0,0.1)}
  .header{background:linear-gradient(135deg,#1a3a5c,#0d2137);color:#fff;padding:28px 32px}
  .header h1{margin:0 0 4px;font-size:20px}
  .header p{margin:0;opacity:0.7;font-size:13px}
  .body{padding:28px 32px}
  .report-no{display:inline-block;background:#e8f0f8;color:#1a3a5c;font-weight:700;font-size:18px;padding:8px 20px;border-radius:8px;margin-bottom:8px;letter-spacing:1px}
  .status-badge{display:inline-block;background:#1a3a5c;color:#fff;font-weight:700;padding:4px 14px;border-radius:20px;font-size:13px;margin-left:8px}
  .footer{background:#f8f9fa;padding:16px 32px;font-size:12px;color:#888;text-align:center}
</style></head><body>
<div class="card">
  <div class="header">
    <h1>เอกสารรายงานอุบัติการณ์ (ฉบับเต็ม)</h1>
    <p>Incident &amp; Near Miss Report — Full Detail</p>
  </div>
  <div class="body">
    <span class="report-no">[${escapeHtml_(reportNo)}]</span>
    <span class="status-badge">${escapeHtml_(rec['สถานะ'])}</span>

    ${section('ข้อมูลเหตุการณ์', [
      tr('วันที่บันทึก', rec['วันที่บันทึก']),
      tr('บริษัท', rec['บริษัท']),
      tr('ชื่อโครงการ', rec['ชื่อโครงการ']),
      tr('วันที่เกิดเหตุ', rec['วันที่เกิดเหตุ']),
      tr('เวลาเกิดเหตุ', rec['เวลาเกิดเหตุ']),
      tr('สถานที่เกิดเหตุ', rec['สถานที่เกิดเหตุ']),
      tr('ประเภทเหตุการณ์', rec['ประเภทเหตุการณ์']),
      tr('อื่นๆ (ระบุ)', rec['อื่นๆ (ระบุ)']),
      tr('ระดับความรุนแรง', rec['ระดับความรุนแรง']),
      tr('รายละเอียดเหตุการณ์', rec['รายละเอียดเหตุการณ์']),
    ])}

    ${section('สาเหตุและการดำเนินการ', [
      tr('สาเหตุเบื้องต้น', rec['สาเหตุเบื้องต้น']),
      tr('การดำเนินการทันที', rec['การดำเนินการทันที']),
      tr('RCA - Man (คน)', rec['RCA - Man (คน)']),
      tr('RCA - Machine (เครื่องจักร/อุปกรณ์)', rec['RCA - Machine (เครื่องจักร/อุปกรณ์)']),
      tr('RCA - Method (วิธีการ/ขั้นตอน)', rec['RCA - Method (วิธีการ/ขั้นตอน)']),
      tr('RCA - Material (วัสดุ/สารเคมี)', rec['RCA - Material (วัสดุ/สารเคมี)']),
      tr('มาตรการป้องกัน (Action Items)', rec['มาตรการป้องกัน (Action Items)']),
    ])}

    ${(rec['ผู้บาดเจ็บ/เสียชีวิต'] || rec['ทรัพย์สินเสียหาย'] || rec['ผลกระทบด้านสิ่งแวดล้อม']) ? section('ผลกระทบ', [
      tr('ผู้บาดเจ็บ/เสียชีวิต', rec['ผู้บาดเจ็บ/เสียชีวิต']),
      tr('ทรัพย์สินเสียหาย', rec['ทรัพย์สินเสียหาย']),
      tr('ค่าเสียหาย / ผลกระทบต่อการดำเนินงาน', rec['ผลกระทบด้านสิ่งแวดล้อม']),
    ]) : ''}

    ${section('ผู้เกี่ยวข้อง', [
      tr('ผู้จัดทำ', `${rec['ผู้จัดทำ - ชื่อ'] || ''} (${rec['ผู้จัดทำ - ตำแหน่ง'] || ''}) — ${rec['ผู้จัดทำ - วันที่'] || ''}`),
      tr('Reviewer', (rec['Reviewer - ชื่อ'] || rec['Reviewer - อีเมล']) ? `${rec['Reviewer - ชื่อ'] || ''} <${rec['Reviewer - อีเมล'] || ''}>` : ''),
      tr('Approver', (rec['Approver - ชื่อ'] || rec['Approver - อีเมล']) ? `${rec['Approver - ชื่อ'] || ''} <${rec['Approver - อีเมล'] || ''}>` : ''),
    ])}

    <h3 style="color:#1a3a5c;border-bottom:2px solid #1a3a5c;padding-bottom:6px;margin:24px 0 4px">รูปภาพประกอบ</h3>
    ${renderPhotoGrid_(rec['รูปภาพประกอบ (URLs)'])}
  </div>
  <div class="footer">ระบบ Incident Report — Design by Chakhrit</div>
</div>
</body></html>`;

  return HtmlService.createHtmlOutput(html)
    .setTitle(`รายงาน ${reportNo}`)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ─── SHEET HELPERS ────────────────────────────────────────

function getOrCreateSheet() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.DATA_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.DATA_SHEET_NAME);
    writeHeaders(sheet);
  }

  // เพิ่ม header ถ้ายังไม่มี (กรณี sheet ว่างเปล่า)
  if (sheet.getLastRow() === 0) {
    writeHeaders(sheet);
  }

  return sheet;
}

function writeHeaders(sheet) {
  const headers = [
    'เลขที่รายงาน',
    'วันที่บันทึก',
    'บริษัท',
    'ชื่อโครงการ',
    'วันที่เกิดเหตุ',
    'เวลาเกิดเหตุ',
    'สถานที่เกิดเหตุ',
    'ประเภทเหตุการณ์',
    'อื่นๆ (ระบุ)',
    'ระดับความรุนแรง',
    'รายละเอียดเหตุการณ์',
    'สาเหตุเบื้องต้น',
    'การดำเนินการทันที',
    'RCA - Man (คน)',
    'RCA - Machine (เครื่องจักร/อุปกรณ์)',
    'RCA - Method (วิธีการ/ขั้นตอน)',
    'RCA - Material (วัสดุ/สารเคมี)',
    'มาตรการป้องกัน (Action Items)',
    'ผู้บาดเจ็บ/เสียชีวิต',
    'ทรัพย์สินเสียหาย',
    'ผลกระทบด้านสิ่งแวดล้อม',
    'ผู้จัดทำ - ชื่อ',
    'ผู้จัดทำ - ตำแหน่ง',
    'ผู้จัดทำ - วันที่',
    'Reviewer - ชื่อ',
    'Reviewer - อีเมล',
    'Approver - ชื่อ',
    'Approver - อีเมล',
    'สถานะ',
    'รูปภาพประกอบ (URLs)',
  ];

  sheet.appendRow(headers);

  // จัด style header row
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#1a3a5c');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(11);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
}

/**
 * สร้างเลขที่รายงาน เช่น INC-2025-06-001
 */
function generateReportNo(sheet) {
  const now = new Date();
  const yyyy = now.getFullYear();
  const mm = String(now.getMonth() + 1).padStart(2, '0');
  const prefix = `${CONFIG.REPORT_PREFIX}-${yyyy}-${mm}-`;

  // นับจำนวนรายงานในเดือนนั้นๆ
  const lastRow = sheet.getLastRow();
  let count = 1;
  if (lastRow > 1) {
    const col1 = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const thisMonth = col1.filter(r => r[0] && String(r[0]).startsWith(prefix));
    count = thisMonth.length + 1;
  }

  return prefix + String(count).padStart(3, '0');
}

/**
 * เพิ่มแถวข้อมูลลง Sheet
 * (field names ตรงกับ name="" ใน HTML form)
 */
function appendRow(sheet, reportNo, d) {
  const now = new Date();

  // รวม action items จาก 3 arrays ที่ส่งมา
  const actionItems = formatActionItems(
    d['actionItem[]'] || '',
    d['actionPIC[]']  || '',
    d['actionDate[]'] || ''
  );

  // รวม immediate actions จาก 3 textarea
  const immediateAction = [
    d.emergencyResponse ? '🚨 การตอบสนองฉุกเฉิน:\n' + d.emergencyResponse : '',
    d.siteControl       ? '🚧 การควบคุมพื้นที่:\n'   + d.siteControl       : '',
    d.notification      ? '📢 การแจ้งเตือน:\n'       + d.notification      : '',
  ].filter(Boolean).join('\n\n');

  const row = [
    reportNo,
    Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
    d.company            || '',
    d.project            || '',
    d.incidentDate       || '',
    d.incidentTime       || '',
    d.specificLocation   || '',
    d.incidentCategory   || '',
    d.incidentCategoryOther || '',
    d.severityLevel      || '',
    d.chronology         || '',
    d.primaryRootCause   || '',
    immediateAction,
    d.rcaMan             || '',
    d.rcaMachine         || '',
    d.rcaMethod          || '',
    d.rcaMaterial        || '',
    actionItems,
    d.involvedPersons    || '',
    d.damageExtent       || '',
    [d.estimatedCost ? 'ค่าเสียหายโดยประมาณ: ' + d.estimatedCost + ' บาท' : '',
     d.operationImpact || ''].filter(Boolean).join('\n'),
    d.preparedByName     || '',
    d.preparedByPosition || '',
    d.preparedDate       || '',
    d.reviewedByName     || '',
    d.reviewedByEmail    || '',
    d.approvedByName     || '',
    d.approvedByEmail    || '',
    'รอตรวจสอบ',
    d.photoUrls          || '',
  ];

  sheet.appendRow(row);

  // จัด format แถวข้อมูล
  const newRow = sheet.getLastRow();
  sheet.getRange(newRow, 1, 1, row.length).setVerticalAlignment('top');
  sheet.getRange(newRow, 1, 1, 1).setFontWeight('bold').setFontColor('#1a3a5c');
}

/**
 * รวม action items จาก 3 string (คั่นด้วย ' | ')
 */
function formatActionItems(actionsStr, picsStr, datesStr) {
  if (!actionsStr) return '';
  const actions = actionsStr.split(' | ');
  const pics    = picsStr   ? picsStr.split(' | ')  : [];
  const dates   = datesStr  ? datesStr.split(' | ') : [];
  return actions.map((a, i) =>
    `${i + 1}. ${a}${pics[i] ? ' | ผู้รับผิดชอบ: ' + pics[i] : ''}${dates[i] ? ' | กำหนดเสร็จ: ' + dates[i] : ''}`
  ).join('\n');
}

// ─── EMAIL NOTIFICATIONS ──────────────────────────────────

function sendNotificationEmails(reportNo, d) {
  const subject = `[แจ้งเตือน] รายงานอุบัติการณ์ ${reportNo} — ${d.project || ''}`;

  // 1. แจ้ง QSHE Admin (ไม่มี CC)
  sendToQSHE(subject, reportNo, d);

  // 2. แจ้ง Reviewer (ไม่มี CC)
  if (d.reviewedByEmail) {
    sendToReviewer(subject, reportNo, d);
  }

  // 3. แจ้ง Approver (ไม่มี CC)
  if (d.approvedByEmail) {
    sendToApprover(subject, reportNo, d);
  }

  // 4. ส่ง 1 ฉบับรวมให้ CC List (กรอง Reviewer/Approver ออก ถ้าอยู่ในลิสต์)
  sendToCCList(subject, reportNo, d);
}

function sendToQSHE(subject, reportNo, d) {
  const body = buildEmailBody(reportNo, d, 'admin');
  GmailApp.sendEmail(CONFIG.REVIEWER_EMAIL, subject, '', {
    htmlBody: body,
    name: 'Incident Report System',
  });
}

function sendToReviewer(subject, reportNo, d) {
  const body = buildEmailBody(reportNo, d, 'reviewer');
  GmailApp.sendEmail(d.reviewedByEmail, subject, '', {
    htmlBody: body,
    name: 'Incident Report System',
    replyTo: CONFIG.REVIEWER_EMAIL,
  });
}

function sendToApprover(subject, reportNo, d) {
  const body = buildEmailBody(reportNo, d, 'approver');
  GmailApp.sendEmail(d.approvedByEmail, subject, '', {
    htmlBody: body,
    name: 'Incident Report System',
    replyTo: CONFIG.REVIEWER_EMAIL,
  });
}

/**
 * ส่ง 1 ฉบับรวม (FYI) ให้ CC List
 * กรองอีเมลออกถ้าคนนั้นเป็น Reviewer หรือ Approver ของ case นี้
 */
function sendToCCList(subject, reportNo, d) {
  const directRecipients = [
    (CONFIG.REVIEWER_EMAIL || '').toLowerCase(),
    (d.reviewedByEmail   || '').toLowerCase(),
    (d.approvedByEmail   || '').toLowerCase(),
  ].filter(Boolean);

  const filteredCC = CC_EMAILS_LIST.filter(
    email => !directRecipients.includes(email.toLowerCase())
  );

  if (filteredCC.length === 0) return;

  const body = buildEmailBody(reportNo, d, 'fyi');
  GmailApp.sendEmail(filteredCC[0], subject, '', {
    htmlBody: body,
    name: 'Incident Report System',
    replyTo: CONFIG.REVIEWER_EMAIL,
    cc: filteredCC.slice(1).join(','),
  });
}

/**
 * สร้าง HTML email template
 */
function buildEmailBody(reportNo, d, role) {
  const severityColor = {
    'ต่ำ': '#2d6a4f', 'ปานกลาง': '#e76f00', 'สูง': '#d62828', 'วิกฤต': '#6b0000'
  };
  const sColor = severityColor[d.severityLevel] || '#1a3a5c';

  const roleMsg = {
    admin:    'มีรายงานอุบัติการณ์ใหม่เข้ามาในระบบ',
    reviewer: `คุณได้รับมอบหมายให้ <strong>ตรวจสอบ (Review)</strong> รายงานนี้`,
    approver: `คุณได้รับมอบหมายให้ <strong>อนุมัติ (Approve)</strong> รายงานนี้`,
    fyi:      'แจ้งเพื่อทราบ — มีรายงานอุบัติการณ์ใหม่เข้ามาในระบบ (FYI)',
  }[role];

  return `
<!DOCTYPE html>
<html lang="th">
<head><meta charset="UTF-8">
<style>
  body{font-family:'Sarabun',Arial,sans-serif;background:#f0f4f8;margin:0;padding:20px}
  .card{max-width:600px;margin:0 auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 16px rgba(0,0,0,0.1)}
  .header{background:linear-gradient(135deg,#1a3a5c,#0d2137);color:#fff;padding:28px 32px}
  .header h1{margin:0 0 4px;font-size:20px}
  .header p{margin:0;opacity:0.7;font-size:13px}
  .body{padding:28px 32px}
  .report-no{display:inline-block;background:#e8f0f8;color:#1a3a5c;font-weight:700;font-size:18px;padding:8px 20px;border-radius:8px;margin-bottom:20px;letter-spacing:1px}
  .severity-badge{display:inline-block;background:${sColor};color:#fff;font-weight:700;padding:4px 14px;border-radius:20px;font-size:13px}
  table{width:100%;border-collapse:collapse;margin:16px 0;font-size:14px}
  td{padding:10px 12px;border-bottom:1px solid #eee;vertical-align:top}
  td:first-child{font-weight:600;color:#1a3a5c;white-space:nowrap;width:40%}
  .action-box{background:#fff3e0;border-left:4px solid #e76f00;padding:14px 18px;border-radius:6px;margin:20px 0;font-size:14px}
  .footer{background:#f8f9fa;padding:16px 32px;font-size:12px;color:#888;text-align:center}
  h3{font-size:14px;margin:20px 0 4px}
</style>
</head>
<body>
<div class="card">
  <div class="header">
    <h1>เอกสารรายงานอุบัติการณ์</h1>
    <p>Incident &amp; Near Miss Report</p>
  </div>
  <div class="body">
    <p>${roleMsg}</p>
    <div class="report-no">[${reportNo}]</div>

    <h3 style="color:#1a3a5c;border-bottom:2px solid #1a3a5c;padding-bottom:6px;margin-top:0">ข้อมูลเหตุการณ์</h3>
    <table>
      <tr><td>บริษัท / โครงการ</td><td>${d.company || ''} — ${d.project || ''}</td></tr>
      <tr><td>วันที่เกิดเหตุ</td><td>${d.incidentDate || ''} ${d.incidentTime ? 'เวลา ' + d.incidentTime : ''}</td></tr>
      <tr><td>สถานที่</td><td>${d.specificLocation || ''}</td></tr>
      <tr><td>ประเภทเหตุการณ์</td><td>${d.incidentCategory || ''}${d.incidentCategoryOther ? ' — ' + d.incidentCategoryOther : ''}</td></tr>
      <tr><td>ระดับความรุนแรง</td><td><span class="severity-badge">${d.severityLevel || '-'}</span></td></tr>
      <tr><td>รายละเอียดเหตุการณ์</td><td>${(d.chronology || '').substring(0, 400)}${(d.chronology || '').length > 400 ? '…' : ''}</td></tr>
    </table>

    <h3 style="color:#1a3a5c;border-bottom:2px solid #1a3a5c;padding-bottom:6px">สาเหตุและการดำเนินการ</h3>
    <table>
      <tr><td>สาเหตุเบื้องต้น</td><td>${d.primaryRootCause || '-'}</td></tr>
      ${d.rcaMan || d.rcaMachine || d.rcaMethod || d.rcaMaterial ? `
      <tr><td>RCA - Man (คน)</td><td>${d.rcaMan || '-'}</td></tr>
      <tr><td>RCA - Machine</td><td>${d.rcaMachine || '-'}</td></tr>
      <tr><td>RCA - Method</td><td>${d.rcaMethod || '-'}</td></tr>
      <tr><td>RCA - Material</td><td>${d.rcaMaterial || '-'}</td></tr>` : ''}
      ${d.emergencyResponse || d.siteControl || d.notification ? `
      <tr><td style="vertical-align:top">การดำเนินการทันที</td><td>
        ${d.emergencyResponse ? '<strong>การตอบสนองฉุกเฉิน:</strong><br>' + d.emergencyResponse.replace(/\n/g,'<br>') + '<br><br>' : ''}
        ${d.siteControl ? '<strong>การควบคุมพื้นที่:</strong><br>' + d.siteControl.replace(/\n/g,'<br>') + '<br><br>' : ''}
        ${d.notification ? '<strong>การแจ้งเตือน:</strong><br>' + d.notification.replace(/\n/g,'<br>') : ''}
      </td></tr>` : ''}
      ${(() => {
        const actionItems = formatActionItems(d['actionItem[]'] || '', d['actionPIC[]'] || '', d['actionDate[]'] || '');
        return actionItems ? `<tr><td style="vertical-align:top">มาตรการป้องกัน<br>(Action Items)</td><td>${actionItems.replace(/\n/g,'<br>')}</td></tr>` : '';
      })()}
    </table>

    ${(d.involvedPersons || d.damageExtent || d.estimatedCost || d.operationImpact) ? `
    <h3 style="color:#1a3a5c;border-bottom:2px solid #1a3a5c;padding-bottom:6px">ผลกระทบ</h3>
    <table>
      ${d.involvedPersons ? `<tr><td>ผู้บาดเจ็บ/เสียชีวิต</td><td>${escapeHtml_(d.involvedPersons)}</td></tr>` : ''}
      ${d.damageExtent ? `<tr><td>ทรัพย์สินเสียหาย</td><td>${escapeHtml_(d.damageExtent)}</td></tr>` : ''}
      ${d.estimatedCost ? `<tr><td>ค่าเสียหายโดยประมาณ</td><td>${escapeHtml_(d.estimatedCost)} บาท</td></tr>` : ''}
      ${d.operationImpact ? `<tr><td>ผลกระทบต่อการทำงาน/การอยู่อาศัย</td><td>${escapeHtml_(d.operationImpact)}</td></tr>` : ''}
    </table>` : ''}

    <h3 style="color:#1a3a5c;border-bottom:2px solid #1a3a5c;padding-bottom:6px">ผู้เกี่ยวข้อง</h3>
    <table>
      <tr><td>ผู้จัดทำ</td><td>${d.preparedByName || ''} (${d.preparedByPosition || ''})</td></tr>
      <tr><td>Reviewer</td><td>${d.reviewedByName || ''} &lt;${d.reviewedByEmail || ''}&gt;</td></tr>
      <tr><td>Approver</td><td>${d.approvedByName || ''} &lt;${d.approvedByEmail || ''}&gt;</td></tr>
    </table>

    <h3 style="color:#1a3a5c;border-bottom:2px solid #1a3a5c;padding-bottom:6px">รูปภาพประกอบ</h3>
    ${renderPhotoGrid_(d.photoUrls)}

    <div style="text-align:center;margin:24px 0 4px">
      <a href="${buildFullReportUrl(reportNo)}" style="display:inline-block;background:#fff;color:#1a3a5c;text-decoration:none;padding:11px 28px;border-radius:8px;font-size:14px;font-weight:700;border:2px solid #1a3a5c">
        📄 ดูรายงานฉบับเต็ม
      </a>
    </div>

    ${(role === 'reviewer' || role === 'approver') ? (() => {
      const roleLabel = role === 'reviewer' ? 'ตรวจสอบ (Review)' : 'อนุมัติ (Approve)';
      const roleValue = role === 'reviewer' ? 'ผู้ตรวจสอบ (Reviewer)' : 'ผู้อนุมัติ (Approver)';
      const formUrl   = buildFormUrl(reportNo, roleValue);
      const btnHtml   = formUrl
        ? `<div style="text-align:center;margin:20px 0">
             <a href="${formUrl}" style="display:inline-block;background:#1a3a5c;color:#fff;text-decoration:none;padding:13px 32px;border-radius:8px;font-size:15px;font-weight:700">
               ✏️ ${roleLabel} รายงานนี้
             </a>
             <div style="margin-top:8px;font-size:12px;color:#888">คลิกเพื่อเปิด Google Form → กรอกผล → กด Submit</div>
           </div>`
        : `<div style="font-size:12px;color:#aaa;text-align:center;margin:12px 0">(ยังไม่ได้ตั้งค่า Form URL — แก้ไข CONFIG.FORM_ID ใน Apps Script)</div>`;
      return `
    <div class="action-box">
      กรุณาดำเนินการภายใน <strong>48 ชั่วโมง</strong><br>
      หากมีข้อสงสัยกรุณาติดต่อทีม QSHE ที่ <strong>${CONFIG.REVIEWER_EMAIL}</strong>
    </div>
    ${btnHtml}`;
    })() : ''}
  </div>
  <div class="footer">
    อีเมลนี้ส่งโดยอัตโนมัติจากระบบ Incident Report — กรุณาอย่าตอบกลับอีเมลนี้โดยตรง
  </div>
</div>
</body>
</html>`;
}
