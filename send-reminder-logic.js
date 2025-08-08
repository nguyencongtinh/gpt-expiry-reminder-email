const { google } = require("googleapis");
const nodemailer = require("nodemailer");
const { JWT } = require("google-auth-library");
require("dotenv").config();

const SHEET_ID = process.env.SHEET_ID;
const SHEET_RANGE = "Sheet1!A2:F"; // Đọc từ dòng 2 để bỏ dòng tiêu đề

// Hàm lấy ngày hôm nay theo giờ Việt Nam (UTC+7)
function getTodayVN() {
  const now = new Date();
  // UTC + 7 giờ (VN)
  return new Date(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate(), 7, 0, 0);
}

// Tính số ngày giữa 2 mốc (không bị lệch múi giờ, chỉ so sánh yyyy-mm-dd)
function daysBetween(date1, date2) {
  // Chỉ lấy phần ngày, không tính giờ/phút
  const d1 = new Date(date1.getFullYear(), date1.getMonth(), date1.getDate());
  const d2 = new Date(date2.getFullYear(), date2.getMonth(), date2.getDate());
  return Math.floor((d2 - d1) / (1000 * 60 * 60 * 24));
}

// Khởi tạo Google Sheets API
const auth = new JWT({
  email: process.env.GOOGLE_CLIENT_EMAIL,
  key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
  scopes: ["https://www.googleapis.com/auth/spreadsheets"],
});

async function getSheetData() {
  const sheets = google.sheets({ version: "v4", auth });
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: SHEET_RANGE,
  });
  return res.data.values || [];
}

async function updateSheet(rowIndex, colIndex, value) {
  const sheets = google.sheets({ version: "v4", auth });
  const range = `Sheet1!${String.fromCharCode(65 + colIndex)}${rowIndex + 2}`;
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range,
    valueInputOption: "RAW",
    requestBody: { values: [[value]] },
  });
}

async function sendEmail(to, subject, text) {
  const transporter = nodemailer.createTransport({
    service: "gmail",
    auth: { user: process.env.GMAIL_USER, pass: process.env.GMAIL_PASS },
  });
  await transporter.sendMail({ from: process.env.GMAIL_USER, to, subject, text });
}

(async () => {
  const data = await getSheetData();
  const today = getTodayVN();

  for (let i = 0; i < data.length; i++) {
    const [email, expireDate, id, gptName, sent5d, sent1d] = data[i];
    if (!email || !expireDate) continue;

    // Chuẩn hóa ngày hết hạn từ dd/mm/yyyy
    const [dd, mm, yyyy] = expireDate.split('/');
    const expire = new Date(parseInt(yyyy), parseInt(mm) - 1, parseInt(dd)); // YYYY, MM (0-11), DD

    const daysLeft = daysBetween(today, expire);

    // Trước 5 ngày
    if (daysLeft === 5 && !sent5d) {
      await sendEmail(
        email,
        `[Nhắc hạn] GPT "${gptName}" sẽ hết hạn sau 5 ngày`,
        `Chào bạn,\nGPT "${gptName}" sẽ hết hạn vào ngày ${expireDate}.\nVui lòng gia hạn để không bị gián đoạn.\nTrân trọng!`
      );
      await updateSheet(i, 4, "Đã gửi"); // cột E
    }

    // Trước 1 ngày
    if (daysLeft === 1 && !sent1d) {
      await sendEmail(
        email,
        `[Nhắc hạn] GPT "${gptName}" sẽ hết hạn NGÀY MAI!`,
        `Chào bạn,\nGPT "${gptName}" sẽ hết hạn vào NGÀY MAI (${expireDate}).\nVui lòng gia hạn nếu muốn tiếp tục sử dụng.\nTrân trọng!`
      );
      await updateSheet(i, 5, "Đã gửi"); // cột F
    }
  }
})();
