# HaiChau Hotspots

Phân loại các khu vực ẩm thực ở trung tâm Đà Nẵng (Hải Châu và lân cận) theo cơ cấu ngành hàng, từ dữ liệu đánh giá Google Maps, và hiển thị kết quả trên bản đồ 3D tương tác.

Dự án kết thúc môn **Data Mining (DAM501)**.

## Thành viên

- Nguyễn Văn Tú
- Bùi Ngọc Bảo Trân
- Lê Ngọc Hiếu
- Lê Bảo Nguyên

## Bài toán

Mỗi khu vực thuộc "kiểu" gì nếu nhìn vào cơ cấu các loại hình ăn uống của nó?

Dữ liệu không có sẵn nhãn "kiểu khu vực", nên đây là bài toán **học không giám sát (phân cụm)**. Kết quả phục vụ người muốn mở quán, người đi ăn và người nghiên cứu đô thị.

## Dữ liệu

Dữ liệu được crawl từ Google Maps bằng [Apify](https://apify.com) (actor `compass/google-maps-reviews-scraper`), trong các ngày 27–29/10/2026.

| Chỉ số | Giá trị |
|---|---:|
| Số đánh giá | 301.566 |
| Số quán | 1.602 |
| Đánh giá có bình luận tiếng Việt (gồm cả bản Google dịch) | 221.278 |
| Quán có thông tin giá mỗi người | 1.266 |
| Thời gian của đánh giá | 07/2012 – 10/2026 |

File dữ liệu `data/cleaned/data_cleaned.xlsx` nặng khoảng 160MB, vượt giới hạn của GitHub nên **không nằm trong repo**. Web không cần file này, nhưng muốn chạy lại notebook thì cần xin từ nhóm. Sheet `Ghi chú` trong file liệt kê các quy tắc đã dùng để làm sạch dữ liệu.

## Phương pháp

Notebook đi theo quy trình 10 bước, từ định nghĩa bài toán đến triển khai và phản hồi.

1. **Tiền xử lý:** bỏ quán đã đóng cửa vĩnh viễn và địa điểm không phải ăn uống, còn **1.581 quán**. Khoảng 150 category của Google được gộp thành 11 nhóm món và 7 nhóm ngành hàng.
2. **Đơn vị phân tích:** lưới ô 250 m, giữ các ô có từ 5 quán trở lên, được **70 ô** chứa 1.537 quán.
3. **Đặc trưng:** tỷ lệ 7 nhóm ngành hàng của mỗi ô, làm mượt không gian với bán kính 250 m, biến đổi Hellinger rồi giảm còn 3 chiều bằng PCA.
4. **Mô hình:** so sánh K-Means, Hierarchical (Ward) và GMM với k từ 3 đến 7.
5. **Đánh giá:** Silhouette, Davies–Bouldin, Calinski–Harabasz, độ ổn định bằng bootstrap và độ đồng thuận giữa các mô hình (ARI).
6. **Đặt tên:** dựa trên 1–2 nhóm ngành hàng vượt mức chung nhiều nhất của mỗi cụm.

Giá, điểm và bình luận **không đưa vào phân cụm**, chỉ dùng để mô tả các cụm sau đó.

## Kết quả

Mô hình được chọn là **K-Means với k = 7**, cho **7 kiểu khu vực** và **11 vùng liền mạch**, tổng diện tích khoảng 4,26 km² (chỉ tính phần đất).

| Kiểu khu vực | Số ô | Số quán | Điểm TB |
|---|---:|---:|---:|
| Khu lẩu nướng · nhậu – món quốc tế | 13 | 383 | 4,54 |
| Khu bún phở · bánh mì – món quốc tế | 14 | 363 | 4,38 |
| Khu cà phê | 15 | 269 | 4,33 |
| Khu cà phê – bún phở · bánh mì | 9 | 222 | 4,52 |
| Khu ăn vặt – cơm · món Việt | 11 | 182 | 4,41 |
| Khu cơm · món Việt – lẩu nướng · nhậu | 4 | 76 | 4,46 |
| Khu lẩu nướng · nhậu – cà phê | 4 | 42 | 4,35 |

Tên mô tả cơ cấu quán trong khu, không nói lên đối tượng khách hàng.

### Lựa chọn số cụm

| k | Silhouette | Davies–Bouldin | Calinski–Harabasz | Độ ổn định |
|---:|---:|---:|---:|---:|
| 3 | 0,332 | 1,027 | 36,5 | 0,532 |
| 4 | 0,384 | 0,942 | 43,6 | 0,732 |
| 5 | 0,357 | 0,948 | 43,0 | 0,678 |
| 6 | 0,362 | 0,893 | 42,8 | 0,602 |
| **7** | **0,361** | **0,827** | **44,6** | **0,630** |

Nhóm chọn k = 7 để bản đồ chia khu vực chi tiết hơn. k = 7 có Davies–Bouldin và Calinski–Harabasz tốt nhất, nhưng k = 4 có Silhouette và độ ổn định cao hơn. Ở k = 7, độ đồng thuận với Hierarchical và GMM là ARI 0,58 và 0,57.

## Web

`index.html` hiển thị kết quả trên bản đồ MapLibre GL:

- Bản đồ các vùng với biểu đồ cột 3D, có chế độ bản đồ và vệ tinh, giao diện sáng và tối.
- Bộ lọc theo địa điểm, diện tích, loại hình kinh doanh và kiểu khu vực.
- Khoanh vùng tuỳ ý trên bản đồ và tìm kiếm khu vực hoặc quán.
- Chi tiết từng vùng:
  - Cơ cấu loại hình và lý do vùng có tên đó.
  - So sánh với mức chung và xu hướng bình luận theo thời gian.
  - Giá, điểm và từ khoá bình luận.
  - Quán theo mức đánh giá và bình luận tích cực.

## Cấu trúc thư mục

```
├── index.html                    Web
├── HaiChau_Hotspots.ipynb        Phân tích và mô hình
├── scripts/
│   ├── update_web.py             Chạy lại notebook để cập nhật dữ liệu web
│   └── crawl_apify_reviews.py    Crawl đánh giá bằng Apify
├── data/
│   ├── README.md                 Mô tả các file dữ liệu
│   ├── cleaned/                  data_cleaned.xlsx (không có trong repo)
│   ├── geo/water.geojson         Sông và biển, để cắt hình vùng theo bờ
│   └── web/areas.json            Dữ liệu cho web, do notebook xuất ra
└── brand/                        Logo và icon
```

## Cách chạy

**Xem web:** mở `index.html` bằng Live Server hoặc một web server bất kỳ. Không mở trực tiếp bằng `file://`, vì trình duyệt sẽ chặn đọc dữ liệu. Ví dụ:

```bash
python -m http.server 8000
# rồi mở http://localhost:8000
```

Bản đồ nền cần có Internet.

**Chạy lại phân tích:** cần Python 3.12 và file `data/cleaned/data_cleaned.xlsx`.

```bash
pip install pandas numpy scipy scikit-learn matplotlib seaborn openpyxl shapely jupyter ipykernel
python -m ipykernel install --user --name msa --display-name "Python (msa)"
python scripts/update_web.py
```

Lệnh cuối chạy lại notebook và ghi đè `data/web/areas.json`.

**Deploy:** web là trang tĩnh, chỉ cần `index.html`, `brand/` và `data/web/areas.json`. File `.vercelignore` loại file dữ liệu lớn khi deploy lên Vercel.

## Hạn chế

- **Mẫu nhỏ:** mô hình chỉ có 70 ô, hai kiểu khu vực chỉ gồm 4 ô, và kết quả có thể đổi theo kích thước ô hoặc bán kính làm mượt.
- **Dữ liệu chưa đều:** danh sách thiên về quán có trên Google Maps, và số đánh giá crawl được khác nhau giữa các quán (119 quán dưới 50% tổng số đánh giá).
- **Giá chỉ có ở một phần quán:** web giữ nguyên khoảng giá người đánh giá chọn trên Google và để trống nếu quán không có. Khoảng này khá rộng nên gần như mọi vùng đều phổ biến nhất ở mức "đến 100k".
- **Bình luận tích cực chưa phải phân tích cảm xúc:** tỷ lệ tính theo số sao, còn bình luận mẫu được lọc bằng từ khoá nên có thể sót câu lẫn ý chê.
- **Xu hướng có thể thiên lệch:** một số đợt crawl ưu tiên bình luận mới nhất, nên không nên xem là mức tăng trưởng thật.

## Nguồn dữ liệu và công cụ

- Đánh giá: Google Maps, thu thập qua Apify.
- Bản đồ nền và ranh giới sông, biển: © [OpenStreetMap](https://www.openstreetmap.org/copyright) contributors, qua OpenFreeMap.
- Ảnh vệ tinh: Esri, Maxar, Earthstar Geographics.
- Bản đồ web: MapLibre GL JS.
