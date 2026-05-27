BẢN DỌN CODE V1

Mục tiêu:
- Giữ nguyên các chức năng đang có.
- Tách JavaScript ra khỏi HTML để dễ đọc và dễ sửa.
- Làm lại CSS cho gọn hơn, bỏ kiểu CSS lồng nhau dễ gây lỗi trên trình duyệt.
- Không thêm tính năng mới.

Cấu trúc file:
- index.html: màn hình import Excel, danh sách hồ sơ, modal in, lịch sử.
- index.css: giao diện trang quản lý.
- index.js: logic đọc Excel, lưu localStorage, mở/in hồ sơ, lưu lịch sử.
- A4.html: nội dung các biểu mẫu A4.
- A4.css: định dạng biểu mẫu và định dạng khi in.
- A4.js: logic tự điền dữ liệu, thêm/xóa dòng bảng, đồng bộ số lượng, xử lý chữ ký, paste plain text.

Lưu ý:
- Mở index.html để chạy dự án.
- File vẫn dùng thư viện XLSX qua CDN như bản cũ.
- Bản này chưa tự ý sửa nội dung văn bản hành chính trong biểu mẫu.
