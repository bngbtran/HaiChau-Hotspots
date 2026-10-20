"""Gộp file crawl của các thành viên thành data/crawl/prices.csv và data/crawl/reviews.csv.

Mỗi thành viên nộp MỘT file Excel vào data/crawl/parts/ (ví dụ parts/tran.xlsx), gồm:

  Sheet đầu tiên (bảng quán đang crawl, giữ nguyên các cột cũ, chỉ thêm cột):
    storeId | Trạng thái | Giá thấp nhất | Giá cao nhất | Nguồn | (các cột khác được bỏ qua)
    - Trạng thái: Not Yet / Done / Không tìm thấy / Đóng cửa
    - Giá: số (30000) hoặc chữ ("30.000đ", "30k"); có thể ghi cả khoảng "30.000 - 50.000đ" vào cột "Giá thấp nhất"

  Sheet "Bình luận" (mỗi bình luận một dòng):
    storeId | Số sao | Bình luận | Ngày | Nguồn

Script kiểm tra lỗi (storeId lạ, trùng giữa các thành viên, giá/sao không hợp lệ, bình luận trùng),
in báo cáo tiến độ từng người, rồi ghi file theo đúng định dạng trong data/README.md.

Chạy:  python scripts/merge_crawl.py            (cần pandas + openpyxl, ví dụ môi trường conda msa)
       python scripts/merge_crawl.py --parts <thư mục> --out <thư mục>
"""
import argparse
import re
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
STATUS_DONE = "done"

# Tên cột chấp nhận (không phân biệt hoa thường, dấu cách) -> tên chuẩn
STORE_COLS = {"storeid": "storeId", "trạng thái": "status", "trang thai": "status", "status": "status",
              "giá thấp nhất": "price_min", "gia thap nhat": "price_min", "price_min": "price_min", "giá": "price_min",
              "giá cao nhất": "price_max", "gia cao nhat": "price_max", "price_max": "price_max",
              "nguồn": "source", "nguon": "source", "source": "source"}
REVIEW_COLS = {"storeid": "storeId", "số sao": "rating", "so sao": "rating", "rating": "rating", "sao": "rating",
               "bình luận": "comment", "binh luan": "comment", "comment": "comment",
               "ngày": "created_at", "ngay": "created_at", "created_at": "created_at", "nguồn": "source", "nguon": "source", "source": "source"}


def rename(df, mapping):
    cols = {c: mapping[str(c).strip().lower()] for c in df.columns if str(c).strip().lower() in mapping}
    return df.rename(columns=cols)[list(dict.fromkeys(cols.values()))]


def parse_money(v):
    """'30.000đ' -> 30000, '30k' -> 30000, 30 -> 30000 (số nhỏ hơn 1000 hiểu là nghìn đồng)."""
    if pd.isna(v) or str(v).strip() == "":
        return None
    s = str(v).lower().replace(" ", "")
    m = re.search(r"(\d+(?:[.,]\d+)*)(k|nghìn|ngàn)?", s)
    if not m:
        return None
    num = float(m.group(1).replace(".", "").replace(",", "")) if not m.group(2) else float(m.group(1).replace(",", "."))
    if m.group(2) or num < 1000:
        num *= 1000
    return int(num)


def split_range(v):
    """'30.000 - 50.000đ' -> (30000, 50000), '30-50k' -> (30000, 50000); một giá trị -> (giá trị, None)."""
    parts = re.split(r"\s*[-–~]\s*|\s+đến\s+", str(v)) if isinstance(v, str) else [v]
    return parse_money(parts[0]), parse_money(parts[1]) if len(parts) > 1 else None


def load_master():
    df = pd.read_excel(ROOT / "data" / "DaNangData_full.xlsm", engine="openpyxl", usecols=["storeId"])
    return set(df.storeId.astype(str))


def read_part(path):
    xl = pd.ExcelFile(path, engine="openpyxl")
    stores = rename(xl.parse(xl.sheet_names[0], dtype=str), STORE_COLS)
    review_sheet = next((s for s in xl.sheet_names if s.strip().lower() in ("bình luận", "binh luan", "reviews")), None)
    reviews = rename(xl.parse(review_sheet, dtype=str), REVIEW_COLS) if review_sheet else pd.DataFrame(columns=["storeId"])
    return stores, reviews


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--parts", default=ROOT / "data" / "crawl" / "parts", type=Path)
    ap.add_argument("--out", default=ROOT / "data" / "crawl", type=Path)
    args = ap.parse_args()

    files = sorted(p for p in args.parts.glob("*.xlsx") if not p.name.startswith("~$"))
    if not files:
        sys.exit(f"Không có file .xlsx nào trong {args.parts}")
    master = load_master()
    warns, P, R, progress = [], [], [], []

    for f in files:
        who = f.stem
        stores, reviews = read_part(f)
        if "storeId" not in stores:
            warns.append(f"[{who}] sheet đầu không có cột storeId → bỏ qua file"); continue
        stores = stores.dropna(subset=["storeId"]).assign(storeId=lambda d: d.storeId.str.strip(), member=who)
        st = stores.get("status", pd.Series("", index=stores.index)).fillna("Not Yet").str.strip()
        progress.append({"thành viên": who, "số quán": len(stores), **st.str.lower().value_counts().to_dict()})

        unknown = stores[~stores.storeId.isin(master)]
        for sid in unknown.storeId.head(5):
            warns.append(f"[{who}] storeId không có trong file gốc: {sid}")
        if len(unknown) > 5:
            warns.append(f"[{who}] ... và {len(unknown) - 5} storeId lạ khác")
        stores = stores[stores.storeId.isin(master)]  # storeId lạ: chỉ cảnh báo, không đưa vào kết quả

        for _, r in stores.iterrows():
            if "price_min" not in stores or pd.isna(r.get("price_min")):
                continue
            lo, hi = split_range(r["price_min"])
            hi2 = parse_money(r.get("price_max")) if "price_max" in stores else None
            hi = hi2 if hi2 is not None else hi if hi is not None else lo
            if lo is None or not (5_000 <= lo <= hi <= 5_000_000):
                warns.append(f"[{who}] {r.storeId}: giá không hợp lệ ({r.get('price_min')} – {r.get('price_max')})"); continue
            P.append({"storeId": r.storeId, "price_min": lo, "price_max": hi,
                      "source": (r.get("source") or who) if pd.notna(r.get("source")) else who, "member": who,
                      "done": str(r.get("status", "")).strip().lower() == STATUS_DONE})

        if "comment" in reviews:
            reviews = reviews.dropna(subset=["storeId", "comment"]).assign(storeId=lambda d: d.storeId.str.strip())
            reviews = reviews[reviews.storeId.isin(master)]
            for _, r in reviews.iterrows():
                try:
                    rating = int(float(str(r.get("rating")).replace(",", ".")))
                    assert 1 <= rating <= 5
                except (ValueError, AssertionError):
                    warns.append(f"[{who}] {r.storeId}: số sao không hợp lệ ({r.get('rating')})"); continue
                raw_date = str(r.get("created_at") or "")
                # Ngày kiểu Việt Nam (20/10/2026) hoặc ISO (2026-10-20)
                date = pd.to_datetime(raw_date, errors="coerce", dayfirst=not re.match(r"\d{4}-", raw_date))
                R.append({"storeId": r.storeId, "rating": rating, "comment": str(r.comment).strip(),
                          "created_at": "" if pd.isna(date) else date.date().isoformat(),
                          "source": r.get("source") if pd.notna(r.get("source")) else who, "member": who})

    prices = pd.DataFrame(P)
    if len(prices):
        dup = prices[prices.duplicated("storeId", keep=False)]
        for sid, g in dup.groupby("storeId"):
            warns.append(f"storeId {sid} có giá từ nhiều người ({', '.join(g.member)}) → giữ bản đánh dấu Done")
        prices = prices.sort_values("done", ascending=False).drop_duplicates("storeId")
    reviews = pd.DataFrame(R)
    if len(reviews):
        n0 = len(reviews)
        reviews = reviews.drop_duplicates(["storeId", "comment"])
        if n0 > len(reviews):
            warns.append(f"Bỏ {n0 - len(reviews)} bình luận trùng (cùng quán, cùng nội dung)")
        reviews["review_id"] = reviews.storeId + "-R" + (reviews.groupby("storeId").cumcount() + 1).astype(str).str.zfill(2)

    args.out.mkdir(parents=True, exist_ok=True)
    prices.reindex(columns=["storeId", "price_min", "price_max", "source"]).to_csv(args.out / "prices.csv", index=False, encoding="utf-8-sig")
    reviews.reindex(columns=["storeId", "review_id", "rating", "comment", "created_at", "source"]).to_csv(args.out / "reviews.csv", index=False, encoding="utf-8-sig")

    print("TIẾN ĐỘ")
    table = pd.DataFrame(progress).fillna(0)
    print(table.astype({c: int for c in table.columns if c != "thành viên"}).to_string(index=False))
    print(f"\nĐã ghi {args.out / 'prices.csv'}: {len(prices)} quán có giá / {len(master)} quán trong file gốc")
    print(f"Đã ghi {args.out / 'reviews.csv'}: {len(reviews)} bình luận cho {reviews.storeId.nunique() if len(reviews) else 0} quán")
    if warns:
        print(f"\nCẢNH BÁO ({len(warns)}):")
        for w in warns[:40]:
            print(" -", w)
        if len(warns) > 40:
            print(f" ... và {len(warns) - 40} cảnh báo khác")


if __name__ == "__main__":
    main()
