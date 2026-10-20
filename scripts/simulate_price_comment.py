"""Giả lập trường GIÁ và BÌNH LUẬN cho từng cửa hàng trong data/DaNangData_full.xlsm.

Dùng tạm cho tới khi crawl được dữ liệu thật. Đầu ra có ĐÚNG định dạng mà dữ liệu crawl phải theo,
nên khi có data/crawl/prices.csv và data/crawl/reviews.csv thì notebook chỉ cần đọc thư mục đó thay cho data/sim/.

    data/sim/prices.csv   storeId, price_min, price_max, source
    data/sim/reviews.csv  storeId, review_id, rating, comment, created_at, source

Chạy:  py scripts/simulate_price_comment.py      (chỉ cần thư viện chuẩn của Python)
Kết quả cố định nhờ seed, chạy lại cho ra cùng dữ liệu.
"""
import csv
import math
import random
import re
import zipfile
from datetime import date, timedelta
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "data" / "DaNangData_full.xlsm"
OUT = ROOT / "data" / "sim"
SEED = 501
NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def read_xlsm(path):
    """Đọc sheet đầu tiên của file .xlsm/.xlsx thành list[dict] mà không cần openpyxl (không chạy macro)."""
    with zipfile.ZipFile(path) as z:
        shared = [
            "".join(t.text or "" for t in si.iter(f"{{{NS['m']}}}t"))
            for si in ET.fromstring(z.read("xl/sharedStrings.xml")).findall("m:si", NS)
        ]
        sheet = ET.fromstring(z.read("xl/worksheets/sheet1.xml"))
    rows = []
    for row in sheet.find("m:sheetData", NS).findall("m:row", NS):
        vals = {}
        for c in row.findall("m:c", NS):
            col = re.sub(r"\d", "", c.get("r"))
            v = c.find("m:v", NS)
            if v is None:
                t = c.find(".//m:t", NS)
                vals[col] = t.text if t is not None else ""
            else:
                vals[col] = shared[int(v.text)] if c.get("t") == "s" else v.text
        rows.append(vals)
    header = rows[0]
    keys = sorted(header, key=lambda k: (len(k), k))
    return [{header[k]: r.get(k, "") for k in keys} for r in rows[1:]]


# Nhóm món: lấy từ khoá đầu tiên của productType khớp với bảng dưới
GROUP_RULES = [
    ("lau_nuong", ["Hotpot & Grill", "Beef", "BBQ"]),
    ("tra_sua", ["Milk Tea"]),
    ("ca_phe", ["Coffee - Tea - Juice"]),
    ("bun_pho", ["Noodles", "Congee"]),
    ("com", ["Rice", "Chicken"]),
    ("banh_mi", ["Bread", "Bánh Truyền Thống"]),
    ("an_vat", ["Snack", "Dessert", "Chè"]),
    ("quoc_te", ["International Food", "Pizza", "Korean", "Japanese", "Dimsum", "Fast Food"]),
    ("healthy", ["Healthy Food", "Vegetarian"]),
]


def food_group(category, product_type):
    if category == "SHOP":
        return "cua_hang"
    first = product_type.split(",")[0].strip()
    for g, keys in GROUP_RULES:
        if any(k in first for k in keys):
            return g
    for g, keys in GROUP_RULES:  # thử các nhãn phía sau nếu nhãn đầu là "Others"
        if any(k in product_type for k in keys):
            return g
    return "khac"


# Giá trung vị mỗi người (nghìn đồng) và độ lệch log theo nhóm
PRICE = {
    "ca_phe": (32, .30), "tra_sua": (32, .25), "bun_pho": (38, .25), "com": (42, .30), "banh_mi": (22, .30),
    "an_vat": (25, .35), "lau_nuong": (170, .35), "quoc_te": (120, .45), "healthy": (65, .30),
    "cua_hang": (55, .60), "khac": (45, .45),
}

# Câu bình luận mẫu theo nhóm món và theo cảm xúc (tốt / trung bình / chưa tốt)
DISH = {
    "ca_phe": ["cà phê muối", "bạc xỉu", "cà phê sữa đá", "trà đào", "nước ép"],
    "tra_sua": ["trà sữa trân châu", "trà sữa matcha", "hồng trà kem cheese", "trà trái cây"],
    "bun_pho": ["mì Quảng", "bún bò", "phở bò", "bún chả cá", "cháo lòng"],
    "com": ["cơm gà", "cơm tấm", "cơm niêu", "cơm văn phòng"],
    "banh_mi": ["bánh mì thịt", "bánh mì chả cá", "bánh xèo", "bánh bèo"],
    "an_vat": ["chè", "bánh tráng trộn", "ốc", "kem", "xiên que"],
    "lau_nuong": ["lẩu thái", "bò nướng", "lẩu hải sản", "nướng ngói", "lẩu gà lá é"],
    "quoc_te": ["pizza", "gà rán", "sushi", "mì cay", "dimsum"],
    "healthy": ["salad", "bowl ức gà", "nước ép detox", "đồ chay"],
    "cua_hang": ["đồ khô", "bánh kẹo", "trái cây", "đặc sản"],
    "khac": ["món ăn", "đồ ăn", "món đặc biệt"],
}
GOOD = ["{dish} ngon, đậm vị", "{dish} rất ngon, sẽ quay lại", "Quán sạch sẽ, {dish} chuẩn vị", "Nhân viên thân thiện, {dish} ngon",
        "{dish} ở đây ngon nhất khu này", "Không gian thoáng, {dish} ổn áp"]
MID = ["{dish} tạm được", "{dish} ăn được, không quá đặc biệt", "Quán hơi đông, {dish} ổn", "{dish} bình thường, phục vụ hơi chậm"]
BAD = ["{dish} hơi nhạt", "Chờ lâu, {dish} không như kỳ vọng", "{dish} không ngon như review", "Quán hơi bẩn, {dish} nguội"]
PRICE_GOOD = ["giá hợp lý", "giá rẻ", "đáng tiền", "giá sinh viên"]
PRICE_BAD = ["hơi đắt", "giá cao so với chất lượng", "giá hơi chát"]
TIME = {
    "ca_phe": ["ngồi làm việc buổi sáng", "hẹn bạn buổi chiều", "view sông đẹp"],
    "tra_sua": ["đi chơi buổi tối", "đông học sinh"],
    "bun_pho": ["ăn sáng", "ăn sáng trước giờ làm"],
    "com": ["ăn trưa văn phòng", "giao hàng nhanh"],
    "banh_mi": ["ăn sáng nhanh", "tiện mang đi"],
    "an_vat": ["ăn vặt buổi tối", "đi chợ ghé ăn"],
    "lau_nuong": ["đi nhậu buổi tối", "tụ tập bạn bè", "mở cửa khuya"],
    "quoc_te": ["đi cùng gia đình", "khách du lịch đông"],
    "healthy": ["ăn kiêng", "sau giờ tập"],
    "cua_hang": ["mua làm quà", "gần chợ"],
    "khac": ["tiện đường"],
}


def make_comment(rng, group, rating, price_level):
    dish = rng.choice(DISH[group])
    base = rng.choice(GOOD if rating >= 4 else MID if rating >= 3 else BAD).format(dish=dish)
    parts = [base[0].upper() + base[1:]]
    if rng.random() < .55:
        parts.append(rng.choice(PRICE_BAD if price_level > 1.15 and rng.random() < .6 else PRICE_GOOD))
    if rng.random() < .45:
        parts.append(rng.choice(TIME[group]))
    return ", ".join(parts) + "."


def main():
    rng = random.Random(SEED)
    stores = read_xlsm(SRC)
    OUT.mkdir(parents=True, exist_ok=True)
    start = date(2025, 10, 1)
    prices, reviews = [], []
    for s in stores:
        g = food_group(s["category"], s["productType"])
        try:
            lat, lng = float(s["latitude"]), float(s["longitude"])
        except ValueError:
            lat, lng = 0.0, 0.0
        # Gần bờ sông Hàn (đường Bạch Đằng ~ kinh độ 108.224) và biển Mỹ Khê giá cao hơn
        d_river = abs(lng - 108.2240) * 107  # km
        d_beach = abs(lng - 108.2450) * 107
        level = 1 + .25 * math.exp(-d_river / .6) + .30 * math.exp(-d_beach / .5)
        med, sd = PRICE[g]
        mid = med * level * math.exp(rng.gauss(0, sd))
        lo = max(5, round(mid * rng.uniform(.6, .85)))
        hi = max(lo + 5, round(mid * rng.uniform(1.15, 1.6)))
        prices.append({"storeId": s["storeId"], "price_min": lo * 1000, "price_max": hi * 1000, "source": "simulated"})

        quality = min(4.9, max(2.8, rng.gauss(4.1, .35)))
        n = 0 if rng.random() < .08 else min(25, 1 + int(rng.expovariate(1 / 6)))
        for k in range(n):
            r = int(min(5, max(1, round(rng.gauss(quality, .8)))))
            reviews.append({
                "storeId": s["storeId"], "review_id": f'{s["storeId"]}-R{k + 1:02d}', "rating": r,
                "comment": make_comment(rng, g, r, mid / med),
                "created_at": (start + timedelta(days=rng.randrange(365))).isoformat(), "source": "simulated",
            })

    for name, rows in [("prices.csv", prices), ("reviews.csv", reviews)]:
        with open(OUT / name, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=list(rows[0]))
            w.writeheader()
            w.writerows(rows)
    with_reviews = len({r["storeId"] for r in reviews})
    print(f"{len(stores)} cửa hàng -> {OUT / 'prices.csv'} ({len(prices)} dòng)")
    print(f"{len(reviews)} bình luận cho {with_reviews} cửa hàng -> {OUT / 'reviews.csv'}")


if __name__ == "__main__":
    main()
