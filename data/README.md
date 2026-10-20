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

## Quy trình crawl cho 3 thành viên

1. Mỗi người giữ bảng quán được giao (sheet đầu tiên), **không sửa cột `storeId`**. Đây là khoá để gộp.
2. Thêm 4 cột vào bảng đó: `Trạng thái`, `Giá thấp nhất`, `Giá cao nhất`, `Nguồn` (xem `crawl/MAU_crawl.xlsx`).
   - `Trạng thái`: `Not Yet` / `Done` / `Không tìm thấy` / `Đóng cửa`
   - Giá mỗi người: ghi `30000`, `30.000đ`, `30k` đều được; hoặc ghi cả khoảng `30.000 - 50.000đ` vào cột `Giá thấp nhất`.
3. Thêm sheet **`Bình luận`**, mỗi bình luận một dòng: `storeId | Số sao | Bình luận | Ngày | Nguồn`.
4. Lưu file tên mình (`tran.xlsx`, `minh.xlsx`, …) vào `crawl/parts/`. Nộp lại bản mới bất cứ lúc nào, chỉ cần ghi đè file cũ.

Script gộp tự kiểm tra: `storeId` không có trong file gốc, một quán bị nhiều người crawl (giữ bản `Done`), giá hoặc số sao không hợp lệ, bình luận trùng; và in bảng tiến độ của từng người.

## Nạp dữ liệu mới cho web

```
D:\LIBRARY\MINICONDA\Conda\envs\msa\python.exe scripts/update_web.py
```

Lệnh này gộp `crawl/parts/*.xlsx` → chạy lại notebook → ghi `web/areas.json`. Sau đó:

- **Live Server:** tải lại trang.
- **Vercel:** deploy lại cả thư mục dự án (web cần `index.html` và `data/web/areas.json`).
- **Mở thử file JSON khác:** bấm *Nạp dữ liệu* trên web.

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
