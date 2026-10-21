# Hướng dẫn chạy bot thu thập dữ liệu Google Maps

Bạn **bỏ file Excel phân công của mình vào thư mục `input`, bấm đúp `CHAY_BOT`**, để máy tự chạy, rồi gửi lại file kết quả trong thư mục `output`. Bot tự tìm từng quán trên Google Maps và lấy **toạ độ mới, điểm đánh giá, số lượt đánh giá, khoảng giá (nếu Google có) và bình luận**.

Khoảng 1.000 quán mất **4–5 giờ**. Dừng lúc nào cũng được; chạy lại sẽ làm tiếp.

## 1. Cài đặt (một lần, khoảng 10 phút)

1. Cài **Google Chrome** (nếu chưa có).
2. Cài **Node.js** bản LTS: https://nodejs.org → tải bản "LTS" → cài như phần mềm bình thường.
3. Giải nén `HaiChau-crawl-bot.zip` → được thư mục **`HaiChau-crawl-bot`**.

## 2. Đăng nhập Google (một lần)

Chưa đăng nhập thì Google chỉ cho xem 5 bình luận mỗi quán; đăng nhập rồi bot lấy được tới 20.

1. Bấm đúp **`DANG_NHAP.bat`** (Windows) hoặc **`DANG_NHAP.command`** (macOS). Lần đầu sẽ mất 1–2 phút để cài thư viện.
2. Cửa sổ Chrome mở ra → bấm **Đăng nhập** ở góc phải → đăng nhập (nên dùng tài khoản Google phụ).
3. **Đóng cửa sổ Chrome đó.** Phiên đăng nhập được lưu cho các lần sau.

## 3. Chạy

1. Chép **file Excel phân công của bạn** vào thư mục **`input`** (giữ nguyên các cột: Tên quán, Địa chỉ, Vĩ độ, Kinh độ, storeId, Trạng thái…).
2. Bấm đúp **`CHAY_BOT.bat`** (Windows) hoặc **`CHAY_BOT.command`** (macOS).
3. Để máy chạy. Mỗi quán hiện một dòng, ví dụ:

```
[2/1085] Xôi Chiên                → done      ★4.3 (120) · 20 bình luận · lệch 1 m · còn ~250 phút
```

| Kết quả | Nghĩa |
|---|---|
| `done` | Lấy được dữ liệu |
| `not_found` | Không thấy quán khớp trên Google Maps (ghi kèm vài quán gần giống) |
| `check` | Tìm thấy quán nhưng tên/vị trí lệch nhiều, nhóm trưởng xem lại |
| `error` | Lỗi tạm thời (mạng, trang tải chậm); lần chạy sau bot tự làm lại |

**Trong lúc chạy:**

- **Để nguyên cửa sổ Chrome của bot**, không đóng, không dùng nó làm việc khác. Máy vẫn dùng cho việc khác được.
- **Không cho máy ngủ:** cắm sạc, tắt chế độ ngủ.
- **Dừng:** nhấn `Ctrl + C` trong cửa sổ đen. **Chạy tiếp:** bấm đúp `CHAY_BOT` lần nữa.
- **Captcha:** nếu máy kêu "bíp" và báo Google yêu cầu xác minh, làm captcha trong cửa sổ Chrome của bot; bot tự chạy tiếp.

## 4. Nộp kết quả

File kết quả nằm ở **`output/<tên file của bạn>_ketqua.xlsx`**. File được ghi lại sau mỗi 20 quán và khi chạy xong, nên **có thể gửi giữa chừng**. **Gửi file này cho nhóm trưởng.**

File gồm:

- **Sheet 1:** đúng bảng phân công của bạn, cột **Trạng thái** được cập nhật (Done / Không tìm thấy / Cần kiểm tra / Đóng cửa), thêm các cột **Vĩ độ mới, Kinh độ mới, Đánh giá, Số đánh giá, Giá thấp nhất, Giá cao nhất, Tên trên Google, Loại trên Google, Khớp, Lệch (m), Link Google Maps mới**.
- **Sheet "Bình luận":** mỗi bình luận một dòng: storeId, Tên quán, Số sao, Bình luận, Ngày, Người viết.

Nhiều quán sẽ trống cột giá vì Google không hiển thị khoảng giá cho mọi quán; đó không phải lỗi.

## Không gửi / không xoá

| Thư mục | |
|---|---|
| `output/*.jsonl` | Kết quả thô để chạy tiếp. **Đừng xoá** khi chưa xong |
| `.profile/` | Hồ sơ Chrome có phiên đăng nhập Google của bạn. **Không gửi cho ai** |

## Gặp lỗi

| Hiện tượng | Cách xử lý |
|---|---|
| "Chưa cài Node.js" | Cài Node.js bản LTS rồi bấm lại |
| macOS báo không mở được `.command` | Chuột phải file → **Mở** → **Mở** |
| "Không tìm thấy Google Chrome" | Cài Chrome |
| "Chưa có file Excel nào trong input" | Chép file phân công vào thư mục `input` |
| "không tìm thấy dòng tiêu đề" | File phải có cột **Tên quán** và **Vĩ độ** ở dòng tiêu đề |
| Nhiều quán `error` liên tiếp | Mạng chập chờn: `Ctrl + C`, chờ một lúc rồi chạy lại |
| Captcha hiện liên tục | Dừng 30–60 phút rồi chạy lại |
| File kết quả đang mở trong Excel | Bot ghi sang file khác cùng tên kèm số; đóng Excel trước khi gửi |
