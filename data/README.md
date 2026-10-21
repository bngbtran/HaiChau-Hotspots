# Dữ liệu

| Thư mục / file | Nội dung | Ai tạo |
|---|---|---|
| `DaNangData_full.xlsm` | Dữ liệu gốc 3.256 cửa hàng (không sửa tay) | Nhóm |
| `sim/prices.csv`, `sim/reviews.csv` | Giá và bình luận **giả lập** | `scripts/simulate_price_comment.py` |
| `crawl/MAU_crawl.xlsx` | File mẫu cho thành viên crawl (có sheet Hướng dẫn) | Có sẵn |
| `crawl/parts/<tên>.xlsx` | File crawl của từng thành viên | Thành viên |
| `crawl/prices.csv`, `crawl/reviews.csv` | Giá và bình luận **crawl thật**, đã gộp và kiểm tra | `scripts/merge_crawl.py` |
| `processed/` | Bảng trung gian của notebook | Notebook |
| `web/areas.json` | Kết quả cho web `index.html` | Notebook (bước 9) |

Notebook đọc `crawl/` nếu có `prices.csv` / `reviews.csv`, nếu không thì dùng `sim/`. Không trộn giả lập với dữ liệu thật: quán chưa crawl được sẽ để trống giá trên web.

## Crawl tự động bằng bot Google Maps (khuyến nghị)

Bot `tools/gmaps-bot/` tự tìm từng quán trên Google Maps và lấy đánh giá, số lượt đánh giá, khoảng giá (nếu Google hiển thị), toạ độ mới, loại quán, bình luận. Hướng dẫn cho thành viên: [tools/gmaps-bot/HUONG_DAN.md](../tools/gmaps-bot/HUONG_DAN.md).

1. Gửi cho mỗi thành viên file `HaiChau-crawl-bot.zip` (bot + nút `CHAY_BOT` / `DANG_NHAP` + hướng dẫn) cùng file Excel phân công của họ.
2. Thành viên bấm `DANG_NHAP` (một lần), chép file phân công vào `input/`, bấm `CHAY_BOT`, rồi gửi lại `output/<tên file>_ketqua.xlsx`
   (bảng phân công giữ nguyên cột + Vĩ độ mới, Kinh độ mới, Đánh giá, Số đánh giá, Giá thấp/cao nhất…; sheet "Bình luận").
3. Nhóm trưởng chép các file `*_ketqua.xlsx` vào `data/crawl/parts/` rồi chạy `scripts/update_web.py`:
   giá → `crawl/prices.csv`, bình luận → `crawl/reviews.csv`, đánh giá tổng + toạ độ mới → `crawl/places.csv`; notebook dùng toạ độ mới và đánh giá Google cho các quán khớp.

Quán `Cần kiểm tra` (tên/vị trí lệch) và `Không tìm thấy` không được dùng; có thể sửa tay trong file Excel (điền cột và đổi `Trạng thái` thành `Done`).

## Quy trình crawl cho 3 thành viên (nhập tay)

1. Mỗi người giữ bảng quán được giao (sheet đầu tiên), **không sửa cột `storeId`**. Đây là khoá để gộp.
2. Thêm 4 cột vào bảng đó: `Trạng thái`, `Giá thấp nhất`, `Giá cao nhất`, `Nguồn` (xem `crawl/MAU_crawl.xlsx`).
   - `Trạng thái`: `Not Yet` / `Done` / `Không tìm thấy` / `Đóng cửa`
   - Giá mỗi người: ghi `30000`, `30.000đ`, `30k` đều được; hoặc ghi cả khoảng `30.000 - 50.000đ` vào cột `Giá thấp nhất`.
3. Thêm sheet **`Bình luận`**, mỗi bình luận một dòng: `storeId | Số sao | Bình luận | Ngày | Nguồn`.
4. Lưu file tên mình (`tran.xlsx`, `minh.xlsx`, …) vào `crawl/parts/`. Nộp lại bản mới bất cứ lúc nào, chỉ cần ghi đè file cũ.

Script gộp tự kiểm tra: `storeId` không có trong file gốc, một quán bị nhiều người crawl (giữ bản `Done`), giá hoặc số sao không hợp lệ, bình luận trùng; và in bảng tiến độ của từng người.

## Nạp dữ liệu mới cho web

```
# Windows
D:\LIBRARY\MINICONDA\Conda\envs\msa\python.exe scripts/update_web.py
# macOS
/opt/miniconda3/envs/msa/bin/python scripts/update_web.py
```

Lần đầu trên máy mới: cài `openpyxl` và đăng ký kernel `msa` cho Jupyter (script chạy notebook bằng kernel này):

```
<python của msa> -m pip install openpyxl
<python của msa> -m ipykernel install --user --name msa --display-name "Python (msa)"
```

Lệnh này gộp `crawl/parts/*.xlsx` → chạy lại notebook → ghi `web/areas.json`. Sau đó:

- **Live Server:** tải lại trang.
- **Vercel:** deploy lại cả thư mục dự án (web cần `index.html` và `data/web/areas.json`).
- **Mở thử file JSON khác:** trên web mở menu tài khoản (góc phải trên) → *Nạp dữ liệu*.

Thêm hoặc bớt quán trong danh sách gốc: cập nhật `DaNangData_full.xlsm` (giữ nguyên tên cột), rồi chạy lại lệnh trên.

## Định dạng file sau khi gộp

`prices.csv`, một dòng mỗi cửa hàng:

| Cột | Kiểu | Ghi chú |
|---|---|---|
| `storeId` | chuỗi | Khớp cột `storeId` trong file gốc |
| `price_min`, `price_max` | số nguyên (VNĐ) | Giá mỗi người, ví dụ `30000`, `55000` |
| `source` | chuỗi | `simulated`, `shopeefood`, `googlemaps`… |

`reviews.csv`, một dòng mỗi bình luận:

| Cột | Kiểu | Ghi chú |
|---|---|---|
| `storeId` | chuỗi | Khớp `storeId` |
| `review_id` | chuỗi | Duy nhất (script tự tạo) |
| `rating` | số nguyên 1–5 | |
| `comment` | chuỗi | Nội dung bình luận (UTF-8) |
| `created_at` | `YYYY-MM-DD` | Để trống nếu không lấy được |
| `source` | chuỗi | Như trên |
