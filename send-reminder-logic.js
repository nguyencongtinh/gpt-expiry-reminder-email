/**
 * send-gpt-expiry-reminder.js (Clean markers)
 * - Dò cột theo header (an toàn khi chèn cột)
 * - So sánh theo múi giờ VN (UTC+7)
 * - Gia hạn (daysLeft >= 6) -> tự xóa trắng E/F
 * - Gửi nhắc 5 ngày & 1 ngày -> chỉ ghi "Đã gửi" (không log ngày/ISO)
 */

const { google } = require("googleapis");
const { JWT } = require("google-auth-library");
const nodemailer = require("nodemailer");
require("dotenv").config();

// ===== Config =====
const SHEET_ID = process.env.SHEET_ID;
const SHEET_NAME = process.env.SHEET_NAME || "Sheet1";
const SHEET_READ_RANGE = `${SHEET_NAME}!1:10000`;

// ===== Time (VN) =====
function getTodayVN() {
  const now = new Date();
  return new Date(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate(), 7, 0, 0);
}
function daysBetween(d1, d2) {
  const a = new Date(d1.getFullYear(), d1.getMonth(), d1.getDate());
  const b = new Date(d2.getFullYear(), d2.getMonth(), d2.getDate());
  return Math.floor((b - a) / (1000 * 60 * 60 * 24));
}

// ===== Auth =====
const auth = new JWT({
  email: process.env.GOOGLE_CLIENT_EMAIL,
  key: process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, "\n"),
  scopes: ["https://www.googleapis.com/auth/spreadsheets"],
});

// ===== Utils =====
function colIndexToA1(idx0) {
  let n = idx0 + 1, s = "";
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
function normalizeHeader(h) {
  return String(h || "").trim().toLowerCase().replace(/\s+/g, " ").replace(/[:*()【】\[\]{}]/g, "");
}
function findColIdx(headers, aliases) {
  const norm = headers.map(normalizeHeader);
  for (const a of aliases) { const i = norm.indexOf(normalizeHeader(a)); if (i !== -1) return i; }
  return -1;
}
async function getAllRows() {
  const sheets = google.sheets({ version: "v4", auth });
  const res = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: SHEET_READ_RANGE });
  return res.data.values || [];
}
async function updateCell(row0, col0, value) {
  const sheets = google.sheets({ version: "v4", auth });
  const range = `${SHEET_NAME}!${colIndexToA1(col0)}${row0 + 1}`;
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range,
    valueInputOption: "RAW",
    requestBody: { values: [[value]] },
  });
}

// ===== Email =====
async function sendEmail(to, subject, text) {
  const transporter = nodemailer.createTransport({
    service: "gmail",
    auth: { user: process.env.GMAIL_USER, pass: process.env.GMAIL_PASS },
  });
  await transporter.sendMail({ from: process.env.GMAIL_USER, to, subject, text });
}

// ===== Header aliases =====
const HDR = {
  EMAIL: ["email", "email được phép sử dụng gpts", "email cho phép sử dụng gpts", "địa chỉ email", "email duoc phep su dung gpts"],
  EXPIRE: ["thời hạn sử dụng gpts", "thời hạn sử dụng", "ngày hết hạn", "hạn sử dụng", "thoi han su dung gpts", "han su dung"],
  GPT_ID: ["id", "gpt id", "gpts id", "mã gpt", "ma gpt"],
  GPT_NAME: ["tên gpts", "ten gpts", "ten gpt", "tên gpt"],
  SENT_5D: ["đã gửi trước 5 ngày", "nhắc trước 5 ngày", "nhắc hạn trước 5 ngày", "da gui truoc 5 ngay", "sent 5d", "nhac 5 ngay"],
  SENT_1D: ["đã gửi trước 1 ngày", "nhắc trước 1 ngày", "nhắc hạn trước 1 ngày", "da gui truoc 1 ngay", "sent 1d", "nhac 1 ngay"],
};

// ===== Main =====
(async () => {
  try {
    const rows = await getAllRows();
    if (!rows.length) { console.log("[INFO] Sheet trống."); return; }

    const headers = rows[0];
    const idxEmail = findColIdx(headers, HDR.EMAIL);
    const idxExpire = findColIdx(headers, HDR.EXPIRE);
    const idxId    = findColIdx(headers, HDR.GPT_ID);
    const idxName  = findColIdx(headers, HDR.GPT_NAME);
    const idx5     = findColIdx(headers, HDR.SENT_5D);
    const idx1     = findColIdx(headers, HDR.SENT_1D);

    const missing = [];
    if (idxEmail === -1) missing.push("Email");
    if (idxExpire === -1) missing.push("Thời hạn sử dụng");
    if (idxId === -1)    missing.push("GPT ID");
    if (idxName === -1)  missing.push("Tên GPTs");
    if (idx5 === -1)     missing.push("Đã gửi trước 5 ngày");
    if (idx1 === -1)     missing.push("Đã gửi trước 1 ngày");
    if (missing.length) { console.error("[ERROR] Thiếu cột:", missing.join(", ")); process.exit(1); }

    const today = getTodayVN();
    console.log(`[INFO] Today (VN): ${today.toLocaleDateString("vi-VN")}`);

    for (let r = 1; r < rows.length; r++) {
      const row = rows[r] || [];
      const email   = String(row[idxEmail] || "").trim();
      const expireS = String(row[idxExpire] || "").trim();
      const id      = String(row[idxId] || "").trim();
      const gptName = String(row[idxName] || "").trim();
      const mark5   = String(row[idx5] || "").trim();
      const mark1   = String(row[idx1] || "").trim();

      if (!email || !expireS) continue;

      const m = expireS.match(/^\s*(\d{1,2})\/(\d{1,2})\/(\d{4})\s*$/);
      if (!m) { console.warn(`[WARN] Row ${r + 1}: ngày không dd/mm/yyyy -> "${expireS}"`); continue; }
      const dd = +m[1], mm = +m[2], yyyy = +m[3];
      const expire = new Date(yyyy, mm - 1, dd);
      const expirePretty = `${String(dd).padStart(2,"0")}/${String(mm).padStart(2,"0")}/${yyyy}`;
      const daysLeft = daysBetween(today, expire);

      // 1) Gia hạn xa (>=6 ngày) -> xóa trắng E/F cho sạch
      if (daysLeft >= 6 && (mark5 || mark1)) {
        await updateCell(r, idx5, "");
        await updateCell(r, idx1, "");
        console.log(`[RESET] Row ${r + 1}: clear markers (daysLeft=${daysLeft})`);
      }

      // 2) Chỉ coi là đã gửi nếu còn <6 ngày và ô có chữ
      const sent5 = (daysLeft < 6) && !!mark5;
      const sent1 = (daysLeft < 6) && !!mark1;

      // 3) Gửi 5 ngày
      if (daysLeft === 5 && !sent5) {
        await sendEmail(
          email,
          `[Nhắc hạn] GPT "${gptName}" sẽ hết hạn sau 5 ngày`,
          `Chào bạn,\nGPT "${gptName}" (ID: ${id}) sẽ hết hạn vào ngày ${expirePretty}.\nVui lòng gia hạn để không bị gián đoạn.\nTrân trọng!`
        );
        await updateCell(r, idx5, "Đã gửi");
        console.log(`[MAIL-5D] Row ${r + 1} -> ${email} | ${gptName} | ${expirePretty}`);
      }

      // 4) Gửi 1 ngày
      if (daysLeft === 1 && !sent1) {
        await sendEmail(
          email,
          `[Nhắc hạn] GPT "${gptName}" sẽ hết hạn NGÀY MAI!`,
          `Chào bạn,\nGPT "${gptName}" (ID: ${id}) sẽ hết hạn vào NGÀY MAI (${expirePretty}).\nVui lòng gia hạn nếu muốn tiếp tục sử dụng.\nTrân trọng!`
        );
        await updateCell(r, idx1, "Đã gửi");
        console.log(`[MAIL-1D] Row ${r + 1} -> ${email} | ${gptName} | ${expirePretty}`);
      }
    }

    console.log("[DONE] Reminder job finished.");
  } catch (err) {
    console.error(`[ERROR] ${err?.message || err}`);
    process.exitCode = 1;
  }
})();
