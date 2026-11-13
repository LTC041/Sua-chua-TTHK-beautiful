QUẢN LÝ SỬA CHỮA TTHK

Cấu trúc file:
- index.html / index.css / index.js: màn hình import Excel, danh sách hồ sơ, cửa sổ soạn & in, lịch sử.
- history-store.js: kho lưu lịch sử hồ sơ đã in (IndexedDB + nén gzip).
- history-legacy.css: định dạng cũ, chỉ để xem / in lại hồ sơ lưu trước 03/11/2026.
- docx-export.js: xuất hồ sơ ra file Word (.docx), không cần thư viện ngoài.
- A4.html: nội dung 5 biểu mẫu (tờ trình, BB kiểm tra trước/sau sửa chữa, phiếu báo hỏng, BB thu hồi).
- A4.css: định dạng biểu mẫu, giao diện soạn thảo (chỉ trên màn hình) và định dạng khi in.
- A4.js: bảng điều khiển soạn hồ sơ (cấu hình mặc định nằm ở đầu file).

SOẠN HỒ SƠ (bảng điều khiển bên trái trang A4)
- Thông tin đơn vị: lấy từ Excel, sửa được cho hồ sơ đang soạn.
  Bấm "Lưu vào danh sách" để giữ thay đổi cho các lần sau (nhập lại Excel sẽ ghi đè).
- Thiết bị sửa chữa: nhập một lần, tự hiện ở cả hai biên bản kiểm tra và phiếu báo hỏng.
- Ngày tháng & số văn bản: năm tự lấy năm hiện tại; để trống để chừa chỗ viết tay.
  Ngày < 10 và tháng 1, 2 tự thêm số 0 (theo Nghị định 30/2020).
- Người ký cố định (Phòng KH, Phòng BĐKT, đại diện chức năng, kho, kế toán...):
  sửa cho hồ sơ này, hoặc bấm "Lưu làm mặc định" để dùng cho các hồ sơ sau.
- Quy tắc tự động (mỗi dòng một tên):
  + Lãnh đạo đơn vị luôn ẩn tên khi in.
  + Người đề nghị không cần cột ký đơn vị trên tờ trình.
  + Người đề nghị trùng kế toán tài sản thì bên giao BB thu hồi tự đổi sang người thay.
- Văn bản cần in: bỏ chọn trang không cần in (ghi nhớ cho lần sau).
- Ô "Cần kiểm tra" liệt kê mục còn thiếu; khi bấm In cũng sẽ nhắc lại.

TRÊN TRANG A4
- Chữ tô xanh: dữ liệu tự điền, bấm vào để nhảy tới ô nhập bên trái.
- Ô viền nét đứt: gõ trực tiếp. Với đoạn soạn sẵn (Căn cứ, Kính gửi, Nơi nhận, kết luận,
  chức danh ký...) sẽ hiện thanh "Lưu làm mẫu" / "Về mẫu gốc" để đổi nội dung mặc định
  mà không cần sửa code.
- Bảng: rê chuột vào dòng để thêm (+) / xoá (×) đúng dòng đó; Ctrl+Enter thêm dòng bên dưới.
  Bảng thu hồi tự điền tên, ĐVT, số lượng từ bảng vật tư của tờ trình (ô tô xanh lá);
  gõ vào ô thì ô đó thôi tự điền.
- Ctrl+P: in hồ sơ (vẫn lưu vào lịch sử).

XUẤT WORD
- Nút "Xuất Word" (cạnh nút In): tải file .docx gồm đúng các văn bản đang chọn in,
  với nội dung đang sửa trên trang A4. Mỗi văn bản là một trang riêng.
- File giữ nguyên bố cục: lề trang (trên/dưới/phải 2cm, trái 3cm), phông Times New Roman,
  cỡ chữ, bảng có viền, đầu trang và các cột chữ ký (dạng bảng không viền).
- Tên được ẩn khi in (vd. lãnh đạo đơn vị) nằm trong Word dưới dạng chữ ẩn:
  không hiện, không in; bật nút ¶ (Show/Hide) trong Word để xem.
- Trong cửa sổ Lịch sử, nút "Word" xuất lại file Word của hồ sơ đã in trước đó.

LỊCH SỬ HỒ SƠ ĐÃ IN
- Lưu trong IndexedDB của trình duyệt, mỗi hồ sơ được nén còn khoảng 10KB.
  Bản lưu chứa luôn định dạng (CSS) lúc in, nên sau này có sửa A4.css thì hồ sơ cũ
  vẫn xem / in lại đúng như bản đã in.
- In lại hồ sơ giống hệt bản gần nhất thì chỉ tăng "số lần in".
- "Xem": mở hồ sơ ở tab mới, có nút "In hồ sơ này" ở góc dưới (nút không in ra giấy).
- "In lại": mở thẳng hộp thoại in, không cần mở tab mới.
- Hồ sơ lưu trước bản nâng cấp 03/11/2026 dùng cấu trúc cũ: tự dùng định dạng cũ trong
  history-legacy.css (không sửa file này).
- "Xuất sao lưu" / "Nhập sao lưu" để backup hoặc chuyển sang máy khác.

Lưu ý:
- Mở index.html để chạy dự án; luôn mở theo cùng một cách (cùng trình duyệt, cùng
  đường dẫn / cổng Live Server) vì dữ liệu và cài đặt lưu theo trình duyệt.
- Thư viện XLSX tải qua CDN nên cần có mạng khi import Excel.
