/**
 * send-gpt-expiry-reminder.js (FIXED)
 * - Đọc Google Sheet (JWT Service Account)
 * - Dò cột theo header (không lệch khi chèn/di chuyển cột)
 * - Múi giờ Việt Nam (UTC+7)
 * - RESET E/F khi daysLeft >= 6 (đặt ngay sau khi tính daysLeft, trước mọi continue)
 * - Gửi nhắc hạn 2 mốc: trước 5 ngày & trước 1 ngày
 *
 * ENV cần có:
 *  - SHEET_ID=...
 *  - GOOGLE_CLIENT_EMAIL=...@...iam.gserviceaccount.com
 *  - GOOGLE_PRIVATE_KEY="-----BEGIN PRIVATE KEY-----\n....\n-----END PRIVATE KEY-----\n"
 *  - GMAIL_USER=...
 *  - GMAIL_PASS=...
 *  - (optional) SHEET_NAME=Sheet1
 */

const { google } = require("googleapis");
const { JWT } = require("google-auth-library");
const nodemailer = require("nodemailer");
require("dotenv").config();

// ========= CẤU HÌNH =========
const SHEET_ID = process.env.SHEET_ID;
const SHEET_NAME = process.env.SHEET_NAME || "Sheet1";
const SHEET_READ_RANGE = `${SHEET_NAME}!1:10000`; // đọc cả header + data

// ========= HỖ TRỢ THỜI GIAN (VN) =========
function getTodayVN() {
  // 00:00 theo giờ VN (UTC+7)
  const now = new Date();
  return new Date(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate(), 7, 0, 0);
}
function daysBetween(date1, date2) {
  const d1 = new Date(date1.getFullYear(), date1.getMonth(), date1.getDate());
  const d2 = new Date(date2.getFullYear(), date2.getMonth(), date2.getDate());
  return Math.floor((d2 - d1) / (1000 * 60 * 60 * 24));
}

// ========= AUTH SHEETS =========
const auth = new JWT({
  email: process.env.GOOGLE_CLIENT_EMAIL,
  key: process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, "\n"),
  scopes: ["https://www.googleapis.com/auth/spreadsheets"],
});

// ========= UTILS =========
function colIndexToA1(colIndexZeroBased) {
  // 0->A, 1->B, ... 25->Z, 26->AA ...
  let n = colIndexZeroBased + 1, s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
function normalizeHeader(h) {
  return String(h || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ")
    .replace(/[:*()【】\[\]{}]/g, "");
}
function findColIdx(headers, aliases) {
  const normalized = headers.map(normalizeHeader);
  for (const alias of aliases) {
    const i = normalized.indexOf(normalizeHeader(alias));
    if (i !== -1) return i;
  }
  return -1;
}
function a1Cell(rowIdx0, colIdx0) {
  return `${SHEET_NAME}!${colIndexToA1(colIdx0)}${rowIdx0 + 1}`;
}

// ========= HEADER ALIASES =========
const HDR = {
  EMAIL: ["email", "email được phép sử dụng gpts", "địa chỉ email", "email cho phép sử dụng gpts", "email duoc phep su dung gpts"],
  EXPIRE: ["thời hạn sử dụng gpts", "thời hạn sử dụng", "ngày hết hạn", "hạn sử dụng", "thoi han su dung gpts", "han su dung"],
  GPT_ID: ["id", "gpt id", "mã gpt", "gpts id", "ma gpt"],
  GPT_NAME: ["tên gpts", "ten gpts", "ten gpt", "tên gpt"],
  SENT_5D: ["đã gửi trước 5 ngày", "nhắc trước 5 ngày", "nhắc hạn trước 5 ngày", "da gui truoc 5 ngay", "sent 5d", "nhac 5 ngay"],
  SENT_1D: ["đã gửi trước 1 ngày", "nhắc trước 5 ngày", "nhắc hạn trước 5 ngày", "da gui truoc 1 ngay", "sent 1d", "nhac 1 ngay"],
};

// ========= SHEETS HELPERS =========
async function getAllRows() {
  const sheets = google.sheets({ version: "v4", auth });
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: SHEET_READ_RANGE,
  });
  return res.data.values || [];
}
async function updateCell(rowIndexZeroBased, colIndexZeroBased, value) {
  const sheets = google.sheets({ version: "v4", auth });
  const range = a1Cell(rowIndexZeroBased, colIndexZeroBased);
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range,
    valueInputOption: "RAW",
    requestBody: { values: [[value]] },
  });
}

// ========= EMAIL SENDER =========
async function sendEmail(to, subject, text) {
  const transporter = nodemailer.createTransport({
    service: "gmail",
    auth: { user: process.env.GMAIL_USER, pass: process.env.GMAIL_PASS },
  });
  await transporter.sendMail({ from: process.env.GMAIL_USER, to, subject, text });
}

// ========= MAIN =========
(async () => {
  try {
    const rows = await getAllRows();
    if (!rows.length) {
      console.log("[INFO] Sheet trống.");
      return;
    }

    const headers = rows[0];
    const idxEmail = findColIdx(headers, HDR.EMAIL);
    const idxExpire = findColIdx(headers, HDR.EXPIRE);
    const idxId = findColIdx(headers, HDR.GPT_ID);
    const idxName = findColIdx(headers, HDR.GPT_NAME);
    const idxSent5 = findColIdx(headers, HDR.SENT_5D);
    const idxSent1 = findColIdx(headers, HDR.SENT_1D);

    const missing = [];
    if (idxEmail === -1) missing.push("Email");
    if (idxExpire === -1) missing.push("Thời hạn sử dụng");
    if (idxId === -1) missing.push("GPT ID");
    if (idxName === -1) missing.push("Tên GPTs");
    if (idxSent5 === -1) missing.push("Đã gửi trước 5 ngày");
    if (idxSent1 === -1) missing.push("Đã gửi trước 1 ngày");

    console.log(`[INFO] Header mapping -> Email:${idxEmail}, Expire:${idxExpire}, ID:${idxId}, Name:${idxName}, 5D:${idxSent5}, 1D:${idxSent1}`);

    if (missing.length) {
      console.error("[ERROR] Thiếu cột:", missing.join(", "));
      process.exit(1);
    }

    const today = getTodayVN();
    console.log(`[INFO] Today (VN): ${today.toLocaleDateString("vi-VN")} | Sheet: ${SHEET_NAME}`);

    for (let r = 1; r < rows.length; r++) {
      const row = rows[r] || [];
      const email = String(row[idxEmail] || "").trim();
      const expireStrRaw = String(row[idxExpire] || "").trim();
      const id = String(row[idxId] || "").trim();
      const gptName = String(row[idxName] || "").trim();
      const eRaw = String(row[idxSent5] || "").trim();
      const fRaw = String(row[idxSent1] || "").trim();

      if (!email || !expireStrRaw) {
        if (!email && !expireStrRaw && !id && !gptName) continue; // dòng trống
        console.log(`[SKIP] Row ${r + 1} thiếu email/expire. email="${email}" expire="${expireStrRaw}"`);
        continue;
      }

      // Parse dd/mm/yyyy (chấp nhận 1 chữ số d/m)
      const m = expireStrRaw.match(/^\s*(\d{1,2})\/(\d{1,2})\/(\d{4})\s*$/);
      if (!m) {
        console.warn(`[WARN] Row ${r + 1}: định dạng ngày không phải dd/mm/yyyy -> "${expireStrRaw}"`);
        continue;
      }
      const dd = parseInt(m[1], 10), mm = parseInt(m[2], 10), yyyy = parseInt(m[3], 10);
      const expireStr = `${String(dd).padStart(2,'0')}/${String(mm).padStart(2,'0')}/${yyyy}`;
      const expire = new Date(yyyy, mm - 1, dd);
      const daysLeft = daysBetween(today, expire);

      // ---------- RESET MARKER NGAY SAU KHI TÍNH daysLeft ----------
      if (daysLeft >= 6 && (eRaw || fRaw)) {
        await updateCell(r, idxSent5, "");
        await updateCell(r, idxSent1, "");
        console.log(`[RESET] Row ${r + 1}: cleared E/F vì daysLeft = ${daysLeft} | expire=${expireStr}`);
      }

      // Sau reset, xác định đã gửi (chỉ coi là đã gửi nếu daysLeft < 6 và ô không trống)
      const sent5 = (daysLeft < 6) && !!eRaw;
      const sent1 = (daysLeft < 6) && !!fRaw;

      // ---------- GỬI NHẮC ----------
      if (daysLeft === 5 && !sent5) {
        const subject = `[Nhắc hạn] GPT "${gptName}" sẽ hết hạn sau 5 ngày`;
        const body = `Chào bạn,
GPT "${gptName}" (ID: ${id}) sẽ hết hạn vào ngày ${expireStr}.
Vui lòng gia hạn để không bị gián đoạn.
Trân trọng!`;
        await sendEmail(email, subject, body);
        await updateCell(r, idxSent5, `Đã gửi@${expireStr} (${new Date().toISOString()})`);
        console.log(`[MAIL-5D] Row ${r + 1} -> ${email} | ${gptName} | ${expireStr}`);
      }

      if (daysLeft === 1 && !sent1) {
        const subject = `[Nhắc hạn] GPT "${gptName}" sẽ hết hạn NGÀY MAI!`;
        const body = `Chào bạn,
GPT "${gptName}" (ID: ${id}) sẽ hết hạn vào NGÀY MAI (${expireStr}).
Vui lòng gia hạn nếu muốn tiếp tục sử dụng.
Trân trọng!`;
        await sendEmail(email, subject, body);
        await updateCell(r, idxSent1, `Đã gửi@${expireStr} (${new Date().toISOString()})`);
        console.log(`[MAIL-1D] Row ${r + 1} -> ${email} | ${gptName} | ${expireStr}`);
      }
    }

    console.log("[DONE] Reminder job finished.");
  } catch (err) {
    console.error(`[ERROR] ${err?.message || err}`);
    process.exitCode = 1;
  }
})();
