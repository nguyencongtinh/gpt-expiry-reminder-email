# GPT Expiry Reminder Email

## Mục đích
Tự động gửi email nhắc hạn sử dụng GPTs cho từng người dùng trong Google Sheets, chỉ gửi 2 lần (trước 5 ngày và 1 ngày hết hạn).

## Hướng dẫn sử dụng

### 1. Chuẩn bị

- Tạo Google Service Account, lấy `client_email` và `private_key`
- Share Google Sheet cho email Service Account
- Tạo ứng dụng "App password" cho Gmail hoặc dùng SMTP riêng
- Tạo repo GitHub cá nhân, upload toàn bộ file mã nguồn

### 2. Thiết lập biến môi trường (Secrets)

- SHEET_ID
- GOOGLE_CLIENT_EMAIL
- GOOGLE_PRIVATE_KEY
- GMAIL_USER
- GMAIL_PASS

### 3. Cấu trúc Google Sheet

| Email | Thời hạn sử dụng GPTs | ID | Tên GPTs | Đã gửi trước 5 ngày | Đã gửi trước 1 ngày |
|-------|----------------------|----|----------|---------------------|---------------------|

> Cột B: Định dạng ngày dd/mm/yyyy (ví dụ 10/08/2025)

### 4. Chạy tự động

- Dùng GitHub Actions (`.github/workflows/schedule.yml`) để chạy mỗi ngày
- Hoặc chạy thủ công bằng lệnh `node send-reminder-logic.js`

### 5. Xử lý lỗi
- Nếu lỗi key, kiểm tra key có dấu xuống dòng không đúng thì thay bằng `\\n`
- Nếu không gửi được email, kiểm tra App password hoặc cấu hình SMTP

## Liên hệ hỗ trợ
TINA – trợ lý AI: ChatGPT

