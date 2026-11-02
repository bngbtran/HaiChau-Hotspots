# Dữ liệu

| File | Nội dung | Ai tạo |
|---|---|---|
| `cleaned/data_cleaned.xlsx` | Nguồn dữ liệu duy nhất của dự án: đánh giá Google Maps đã làm sạch. Sheet `Data` (mỗi dòng một đánh giá), `Quán` (mỗi dòng một quán), `Ghi chú` (các quy tắc đã clean) | Nhóm |
| `web/areas.json` | Kết quả phân cụm cho web `index.html` | Notebook (bước 9) |

## Nạp dữ liệu mới cho web

Sau khi cập nhật `cleaned/data_cleaned.xlsx`, chạy:

```
# macOS
/opt/miniconda3/envs/msa/bin/python scripts/update_web.py
# Windows
D:\LIBRARY\MINICONDA\Conda\envs\msa\python.exe scripts/update_web.py
```

Lệnh này chạy lại `HaiChau_Hotspots.ipynb` và ghi đè `web/areas.json`. Sau đó tải lại trang (Live Server) hoặc deploy lại thư mục dự án (Vercel chỉ cần `index.html` và `data/web/areas.json`, không cần file xlsx).

Lần đầu trên máy mới: cài `openpyxl` và đăng ký kernel `msa` cho Jupyter:

```
<python của msa> -m pip install openpyxl
<python của msa> -m ipykernel install --user --name msa --display-name "Python (msa)"
```
