const SPREADSHEET_ID = '16aWlQbWYUTMt48mr7Md_p6eUs97540dE22OU63b7wuU';
const FROM_NAME = 'Revo Skill Bot';

function doGet() {
  return ContentService
    .createTextOutput('OK')
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {
    const data = parsePayload_(e);
    const sheetName = normalizeSheetName_(data.sheet);
    const name = trim_(data.name);
    const empId = trim_(data.empId || data.id);
    const email = trim_(data.email);
    const skillText = trim_(data.skillText || '');
    const shouldSendEmail = toBool_(data.sendEmail);

    if (!name || !empId || !email) {
      throw new Error('Thiếu dữ liệu bắt buộc: name, empId, email.');
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) sheet = ss.insertSheet(sheetName);

    ensureHeaders_(sheet);

    let mailStatus = 'skipped';
    let mailError = '';

    if (shouldSendEmail) {
      try {
        sendConfirmationEmail_({
          to: email,
          name,
          empId,
          sheetName,
          skillText,
        });
        mailStatus = 'sent';
      } catch (mailErr) {
        mailStatus = 'failed';
        mailError = String(mailErr.message || mailErr);
      }
    }

    sheet.appendRow([
      name,
      empId,
      email,
      skillText,
      mailStatus,
      mailError,
    ]);

    return json_({
      ok: true,
      sheet: sheetName,
      mail: {
        status: mailStatus,
        error: mailError,
      },
    });
  } catch (err) {
    return json_({
      ok: false,
      error: err.message || String(err),
    });
  }
}

function parsePayload_(e) {
  if (!e || !e.postData || !e.postData.contents) return {};

  const raw = e.postData.contents.trim();

  try {
    return JSON.parse(raw);
  } catch (jsonErr) {
    const obj = {};
    raw.split('&').forEach((pair) => {
      const idx = pair.indexOf('=');
      if (idx === -1) return;
      const key = decodeURIComponent(pair.slice(0, idx).replace(/\+/g, ' '));
      const val = decodeURIComponent(pair.slice(idx + 1).replace(/\+/g, ' '));
      obj[key] = val;
    });
    return obj;
  }
}

function normalizeSheetName_(sheetName) {
  const s = trim_(sheetName).toUpperCase();
  if (s === 'CAO') return 'CAO';
  if (s === 'SPARK') return 'SPARK';
  throw new Error('sheet phải là CAO hoặc SPARK.');
}

function ensureHeaders_(sheet) {
  const headers = ['Họ và tên', 'ID', 'Email', 'Skill', 'Mail status', 'Mail error'];
  const lastRow = sheet.getLastRow();

  if (lastRow === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    return;
  }

  const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const isBlank = current.every((v) => trim_(v) === '');
  if (isBlank) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
}

function sendConfirmationEmail_({ to, name, empId, sheetName, skillText }) {
  if (!to || !String(to).includes('@')) {
    throw new Error('Email nhận không hợp lệ: ' + to);
  }

  const subject = `[${sheetName}] Xác nhận nhận thông tin tạo Skill`;
  const safeName = escapeHtml_(name);
  const safeEmpId = escapeHtml_(empId);
  const safeEmail = escapeHtml_(to);
  const safeSkill = escapeHtml_(skillText || '(Chưa nhận được nội dung Skill)');

  const textBody = [
    `Chào ${name},`,
    '',
    `Mình đã nhận được thông tin của bạn và đã lưu vào tab ${sheetName} trong Google Sheet.`,
    `- Họ và tên: ${name}`,
    `- ID nhân viên: ${empId}`,
    `- Email: ${to}`,
    '',
    'Nội dung Skill bạn đã submit:',
    skillText || '(Chưa nhận được nội dung Skill)',
    '',
    'Cảm ơn bạn đã submit.',
  ].join('\n');

  const htmlBody = `
    <div style="margin:0;padding:0;background:#f4ead5;">
      <div style="max-width:680px;margin:0 auto;padding:32px 16px;font-family:Arial,Helvetica,sans-serif;color:#1a1410;">
        <div style="background:#1a1410;color:#f4ead5;padding:18px 24px;border-radius:14px 14px 0 0;">
          <div style="font-size:12px;letter-spacing:0.16em;text-transform:uppercase;opacity:0.8;">Revo Skill Bot</div>
          <h1 style="margin:8px 0 0;font-size:28px;line-height:1.2;">Đã nhận thông tin của bạn</h1>
        </div>

        <div style="background:#fbf6e9;padding:24px;border:1px solid #1a1410;border-top:none;border-radius:0 0 14px 14px;">
          <p style="margin:0 0 16px;font-size:16px;line-height:1.6;">Chào <strong>${safeName}</strong>,</p>

          <p style="margin:0 0 18px;font-size:15px;line-height:1.7;color:#3d2f24;">
            Mình đã ghi nhận dữ liệu của bạn vào tab <strong>${sheetName}</strong> trong Google Sheet.
          </p>

          <div style="background:#f4ead5;border:1px solid #ead9b3;padding:16px 18px;border-radius:12px;margin:0 0 20px;">
            <div style="font-size:12px;letter-spacing:0.12em;text-transform:uppercase;color:#d4502a;font-weight:700;margin-bottom:10px;">Thông tin đã nhận</div>
            <div style="font-size:14px;line-height:1.8;">
              <div><strong>Họ và tên:</strong> ${safeName}</div>
              <div><strong>ID nhân viên:</strong> ${safeEmpId}</div>
              <div><strong>Email:</strong> ${safeEmail}</div>
              <div><strong>Sheet:</strong> ${sheetName}</div>
            </div>
          </div>

          <div style="margin:0 0 12px;font-size:12px;letter-spacing:0.12em;text-transform:uppercase;color:#2d5f3f;font-weight:700;">
            Nội dung Skill bạn submit
          </div>

          <pre style="white-space:pre-wrap;word-wrap:break-word;background:#1a1410;color:#f4ead5;padding:18px;border-radius:12px;font-size:13px;line-height:1.7;margin:0 0 20px;font-family:Menlo,Monaco,Consolas,'Courier New',monospace;">${safeSkill}</pre>

          <p style="margin:0;font-size:14px;line-height:1.7;color:#3d2f24;">
            Cảm ơn bạn đã gửi thông tin. Nếu cần chỉnh sửa Skill, bạn chỉ cần submit lại từ trang web.
          </p>
        </div>

        <div style="text-align:center;font-size:12px;color:#3d2f24;margin-top:12px;opacity:0.75;">
          Email này được tạo tự động từ form CAO / SPARK
        </div>
      </div>
    </div>
  `;

  MailApp.sendEmail({
    to,
    subject,
    body: textBody,
    htmlBody,
    name: FROM_NAME,
  });
}

function escapeHtml_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function trim_(value) {
  return String(value == null ? '' : value).trim();
}

function toBool_(value) {
  if (value === true || value === 'true' || value === 1 || value === '1') return true;
  return false;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
