"""Thu thập dữ liệu quán từ Google Places API (New) cho danh sách DaNangData_full.xlsm.

Ba bước, chạy lần lượt (mỗi bước chạy lại được, bỏ qua quán đã có dữ liệu):

  1. ids      Text Search Essentials (IDs Only)          → place_id                              MIỄN PHÍ, không giới hạn
  2. verify   Place Details Pro                            → tên, địa chỉ, toạ độ mới, loại quán   miễn phí 5.000 lượt/tháng
  3. reviews  Place Details Enterprise + Atmosphere        → đánh giá, số đánh giá, khoảng giá,    miễn phí 1.000 lượt/tháng
                                                             mức giá, tối đa 5 bình luận           (sau đó $25 / 1.000 lượt)
     (thêm --no-reviews: Place Details Enterprise, không lấy bình luận: miễn phí 1.000/tháng, sau đó $20 / 1.000)

  4. export   Ghi data/crawl/parts/googlemaps_api.xlsx (sheet "Quán" + "Bình luận") để scripts/update_web.py gộp.

Script đếm số lượt gọi từng SKU trong tháng (data/crawl/api/usage.json) và DỪNG khi chạm mức miễn phí,
trừ khi thêm --allow-paid. Bảng giá: https://developers.google.com/maps/billing-and-pricing/pricing (cập nhật 2026-10-20).

Chuẩn bị: tạo project Google Cloud, bật "Places API (New)", bật billing, tạo API key rồi đặt biến môi trường
(không ghi key vào code hay file trong repo):
    export GOOGLE_MAPS_API_KEY="..."          # macOS / Linux
    set GOOGLE_MAPS_API_KEY=...               # Windows cmd

Ví dụ (môi trường conda msa):
    python scripts/crawl_places_api.py ids
    python scripts/crawl_places_api.py verify
    python scripts/crawl_places_api.py reviews --max 1000
    python scripts/crawl_places_api.py export
    python scripts/crawl_places_api.py status
Thêm --dry-run để xem trước yêu cầu sẽ gửi mà không gọi API; --limit N để thử trên N quán đầu.

Lưu ý điều khoản: Google chỉ cho lưu place_id lâu dài; nội dung khác (đánh giá, giá, bình luận) bị giới hạn lưu trữ
và phải hiển thị kèm ghi nguồn Google. Xem phần Feedback trong notebook.
"""
import argparse
import json
import os
import re
import sys
import time
import unicodedata
from datetime import datetime
from math import asin, cos, radians, sin, sqrt
from pathlib import Path

import pandas as pd
import requests

ROOT = Path(__file__).resolve().parents[1]
API_DIR = ROOT / "data" / "crawl" / "api"
STORE_FILE = API_DIR / "places.jsonl"      # mỗi dòng: kết quả mới nhất của một storeId
USAGE_FILE = API_DIR / "usage.json"
OUT_XLSX = ROOT / "data" / "crawl" / "parts" / "googlemaps_api.xlsx"
BASE = "https://places.googleapis.com/v1"

# SKU → (mức miễn phí mỗi tháng, giá USD / 1.000 lượt ở bậc đầu)
SKU = {
    "text_search_ids": (None, 0.0),
    "details_pro": (5_000, 17.0),
    "details_enterprise": (1_000, 20.0),
    "details_enterprise_atmosphere": (1_000, 25.0),
}
FIELDS_PRO = ["id", "displayName", "formattedAddress", "location", "types", "primaryType", "businessStatus", "googleMapsUri"]
FIELDS_ENT = FIELDS_PRO + ["rating", "userRatingCount", "priceLevel", "priceRange"]
FIELDS_ATM = FIELDS_ENT + ["reviews"]
PRICE_LEVEL = {"PRICE_LEVEL_FREE": "Miễn phí", "PRICE_LEVEL_INEXPENSIVE": "Rẻ", "PRICE_LEVEL_MODERATE": "Vừa phải",
               "PRICE_LEVEL_EXPENSIVE": "Đắt", "PRICE_LEVEL_VERY_EXPENSIVE": "Rất đắt"}


# ---------- tiện ích ----------
def norm(s):
    s = unicodedata.normalize("NFD", str(s or "")).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9 ]", " ", s.replace("đ", "d"))


STOP = {"quan", "nha", "hang", "tiem", "cua", "shop", "cafe", "coffee", "ca", "phe", "va", "da", "nang", "chi", "nhanh", "cs", "co", "so"}


def name_sim(a, b):
    """Tỷ lệ từ chung trên tên ngắn hơn: 'Bún chả cá 109' ⊂ 'Bún Chả Cá 109 - Nguyễn Chí Thanh' → 1.0"""
    A, B = ({t for t in norm(x).split() if t not in STOP} for x in (a, b))
    return len(A & B) / min(len(A), len(B)) if A and B else 0.0


def dist_m(la1, ln1, la2, ln2):
    la1, ln1, la2, ln2 = map(radians, (la1, ln1, la2, ln2))
    h = sin((la2 - la1) / 2) ** 2 + cos(la1) * cos(la2) * sin((ln2 - ln1) / 2) ** 2
    return 2 * 6_371_000 * asin(sqrt(h))


def load_records():
    rec = {}
    if STORE_FILE.exists():
        for line in STORE_FILE.read_text(encoding="utf-8").splitlines():
            if line.strip():
                r = json.loads(line)
                rec[r["storeId"]] = r
    return rec


def save_records(rec):
    API_DIR.mkdir(parents=True, exist_ok=True)
    tmp = STORE_FILE.with_suffix(".tmp")
    tmp.write_text("\n".join(json.dumps(r, ensure_ascii=False) for r in rec.values()) + "\n", encoding="utf-8")
    tmp.replace(STORE_FILE)


class Usage:
    """Đếm lượt gọi từng SKU theo tháng để không vượt mức miễn phí."""

    def __init__(self):
        self.month = datetime.now().strftime("%Y-%m")
        data = json.loads(USAGE_FILE.read_text()) if USAGE_FILE.exists() else {}
        self.data = data
        self.cur = data.setdefault(self.month, {})

    def used(self, sku):
        return self.cur.get(sku, 0)

    def left_free(self, sku):
        cap = SKU[sku][0]
        return float("inf") if cap is None else max(0, cap - self.used(sku))

    def add(self, sku, n=1):
        self.cur[sku] = self.used(sku) + n
        API_DIR.mkdir(parents=True, exist_ok=True)
        USAGE_FILE.write_text(json.dumps(self.data, indent=1))


def api_key():
    key = os.environ.get("GOOGLE_MAPS_API_KEY", "").strip()
    if not key:
        sys.exit("Chưa có API key: đặt biến môi trường GOOGLE_MAPS_API_KEY (xem đầu file).")
    return key


def call(method, url, field_mask, body=None, params=None, dry=False):
    if dry:
        print(f"  [dry-run] {method} {url}  fields={field_mask}  body={json.dumps(body, ensure_ascii=False) if body else ''}")
        return None
    headers = {"X-Goog-Api-Key": api_key(), "X-Goog-FieldMask": field_mask, "Content-Type": "application/json"}
    for attempt in range(5):
        r = requests.request(method, url, headers=headers, json=body, params=params, timeout=30)
        if r.status_code == 200:
            return r.json()
        if r.status_code in (429, 500, 502, 503, 504):
            time.sleep(2 ** attempt)
            continue
        msg = r.json().get("error", {}).get("message", r.text[:200]) if r.headers.get("content-type", "").startswith("application/json") else r.text[:200]
        if r.status_code in (400, 403):
            sys.exit(f"Lỗi {r.status_code} từ Google: {msg}\n(Kiểm tra: đã bật Places API (New) và billing cho project, key không bị giới hạn sai.)")
        return {"_error": f"{r.status_code}: {msg}"}
    return {"_error": "quá nhiều lần thử lại"}


def load_master(limit=None):
    d = pd.read_excel(ROOT / "data" / "DaNangData_full.xlsm", engine="openpyxl")
    d = d.drop_duplicates("storeId")
    return d.head(limit) if limit else d


# ---------- bước 1: place_id (miễn phí) ----------
def step_ids(args):
    master, rec, usage = load_master(args.limit), load_records(), Usage()
    todo = [r for _, r in master.iterrows() if not rec.get(r.storeId, {}).get("place_id") and rec.get(r.storeId, {}).get("id_status") != "not_found"]
    print(f"Text Search (IDs Only, miễn phí): {len(todo)} quán chưa có place_id")
    for i, s in enumerate(todo, 1):
        addr = ", ".join(str(s.address).split(",")[:2])
        body = {"textQuery": f"{s.storeName}, {addr}, Đà Nẵng", "languageCode": "vi", "regionCode": "VN", "pageSize": 3}
        if pd.notna(s.latitude) and pd.notna(s.longitude):
            body["locationBias"] = {"circle": {"center": {"latitude": float(s.latitude), "longitude": float(s.longitude)}, "radius": 500.0}}
        res = call("POST", f"{BASE}/places:searchText", "places.id", body=body, dry=args.dry_run)
        if args.dry_run:
            if i >= 3: break
            continue
        usage.add("text_search_ids")
        r = rec.setdefault(s.storeId, {"storeId": s.storeId})
        if res and res.get("places"):
            r.update(place_id=res["places"][0]["id"], candidates=[p["id"] for p in res["places"]], id_status="found")
        elif res and "_error" in res:
            r.update(id_status="error", error=res["_error"])
        else:
            r.update(id_status="not_found")
        if i % 50 == 0 or i == len(todo):
            save_records(rec); print(f"  {i}/{len(todo)}")
        time.sleep(args.sleep)
    save_records(rec)


# ---------- bước 2 và 3: Place Details ----------
def parse_details(p, s):
    out = {"g_name": (p.get("displayName") or {}).get("text"), "g_address": p.get("formattedAddress"),
           "g_types": ",".join(p.get("types", [])), "g_primary": p.get("primaryType"), "g_status": p.get("businessStatus"),
           "g_url": p.get("googleMapsUri")}
    loc = p.get("location") or {}
    out["lat_new"], out["lng_new"] = loc.get("latitude"), loc.get("longitude")
    sim = name_sim(s.storeName, out["g_name"])
    d = dist_m(s.latitude, s.longitude, out["lat_new"], out["lng_new"]) if out["lat_new"] and pd.notna(s.latitude) else None
    out.update(name_sim=round(sim, 2), dist_m=None if d is None else round(d))
    out["match"] = "ok" if sim >= 0.5 and (d is None or d <= 400) else "check"
    if "rating" in p or "userRatingCount" in p or "priceLevel" in p or "priceRange" in p or "reviews" in p:
        pr = p.get("priceRange") or {}
        units = lambda m: int(m.get("units", 0)) if m else None
        out.update(rating=p.get("rating"), rating_count=p.get("userRatingCount"), price_level=PRICE_LEVEL.get(p.get("priceLevel"), ""),
                   price_min=units(pr.get("startPrice")), price_max=units(pr.get("endPrice")))
        out["reviews"] = [{"rating": v.get("rating"), "comment": ((v.get("text") or v.get("originalText") or {}).get("text") or "").strip(),
                           "date": (v.get("publishTime") or "")[:10], "author": (v.get("authorAttribution") or {}).get("displayName"),
                           "uri": v.get("googleMapsUri")} for v in p.get("reviews", [])]
    return out


def step_details(args, tier):
    sku = {"pro": "details_pro", "ent": "details_enterprise", "atm": "details_enterprise_atmosphere"}[tier]
    fields = {"pro": FIELDS_PRO, "ent": FIELDS_ENT, "atm": FIELDS_ATM}[tier]
    done_key = {"pro": "verified_at", "ent": "rated_at", "atm": "rated_at"}[tier]
    master, rec, usage = load_master(args.limit).set_index("storeId", drop=False), load_records(), Usage()
    todo = [sid for sid, r in rec.items() if r.get("place_id") and not r.get(done_key) and sid in master.index]
    if tier != "pro":   # bước đánh giá: bỏ quán đã biết là khớp sai, ưu tiên quán đã kiểm tra khớp
        skipped = sum(1 for sid in todo if rec[sid].get("match") == "check")
        todo = sorted((sid for sid in todo if rec[sid].get("match") != "check"), key=lambda sid: rec[sid].get("match") != "ok")
        if skipped:
            print(f"Bỏ qua {skipped} quán khớp sai ở bước verify (xem cột Khớp trong file export, sửa place_id tay nếu cần)")
    free = usage.left_free(sku)
    budget = len(todo) if args.allow_paid else min(len(todo), free)
    if args.max is not None:
        budget = min(budget, args.max)
    price = SKU[sku][1]
    paid = max(0, budget - free)
    print(f"{sku}: còn {len(todo)} quán · tháng {usage.month} đã dùng {usage.used(sku)} lượt, còn {free if free != float('inf') else '∞'} lượt miễn phí")
    print(f"→ lần chạy này gọi {budget} lượt" + (f" (trong đó {paid} lượt trả phí ≈ ${paid * price / 1000:.2f})" if paid else " (miễn phí)"))
    if budget < len(todo) and not args.allow_paid:
        print(f"  Dừng ở mức miễn phí. Chạy lại tháng sau, hoặc thêm --allow-paid để trả phí ${price:.0f}/1.000 lượt.")
    for i, sid in enumerate(todo[:budget], 1):
        r, s = rec[sid], master.loc[sid]
        res = call("GET", f"{BASE}/places/{r['place_id']}", ",".join(fields), params={"languageCode": "vi", "regionCode": "VN"}, dry=args.dry_run)
        if args.dry_run:
            if i >= 3: break
            continue
        usage.add(sku)
        if res and "_error" not in res:
            r.update(parse_details(res, s))
            r[done_key] = datetime.now().isoformat(timespec="seconds")
            if tier == "pro":
                r["verified_at"] = r[done_key]
        else:
            r["error"] = (res or {}).get("_error")
        if i % 50 == 0 or i == budget:
            save_records(rec); print(f"  {i}/{budget}")
        time.sleep(args.sleep)
    save_records(rec)


# ---------- xuất Excel cho merge_crawl.py ----------
def step_export(args):
    master, rec = load_master(), load_records()
    rows, revs = [], []
    for _, s in master.iterrows():
        r = rec.get(s.storeId, {})
        if r.get("rated_at"):
            status = "Done"
        elif r.get("id_status") == "not_found":
            status = "Không tìm thấy"
        elif r.get("g_status") == "CLOSED_PERMANENTLY":
            status = "Đóng cửa"
        else:
            status = "Not Yet"
        ok = r.get("match") == "ok"   # chỉ dùng toạ độ / dữ liệu khi tên và vị trí khớp
        rows.append({"storeId": s.storeId, "storeName": s.storeName, "address": s.address,
                     "Trạng thái": status if ok or status != "Done" else "Cần kiểm tra",
                     "Đánh giá": r.get("rating") if ok else None, "Số đánh giá": r.get("rating_count") if ok else None,
                     "Giá thấp nhất": r.get("price_min") if ok else None, "Giá cao nhất": r.get("price_max") if ok else None,
                     "Mức giá": r.get("price_level") if ok else None,
                     "Vĩ độ mới": r.get("lat_new") if ok else None, "Kinh độ mới": r.get("lng_new") if ok else None,
                     "place_id": r.get("place_id"), "Tên trên Google": r.get("g_name"), "Khớp": r.get("match"),
                     "Lệch (m)": r.get("dist_m"), "Link Google Maps": r.get("g_url"), "Nguồn": "googlemaps_api"})
        if ok:
            for v in r.get("reviews", []):
                if v.get("comment"):
                    revs.append({"storeId": s.storeId, "Số sao": v["rating"], "Bình luận": v["comment"], "Ngày": v["date"],
                                 "Người viết": v.get("author"), "Link": v.get("uri"), "Nguồn": "googlemaps_api"})
    OUT_XLSX.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as w:
        pd.DataFrame(rows).to_excel(w, sheet_name="Quán", index=False)
        pd.DataFrame(revs, columns=["storeId", "Số sao", "Bình luận", "Ngày", "Người viết", "Link", "Nguồn"]).to_excel(w, sheet_name="Bình luận", index=False)
    df = pd.DataFrame(rows)
    print(f"Đã ghi {OUT_XLSX}: {len(df)} quán · {df['Trạng thái'].value_counts().to_dict()} · {len(revs)} bình luận")
    print("Tiếp theo: python scripts/update_web.py")


def step_status(args):
    master, rec, usage = load_master(), load_records(), Usage()
    s = pd.Series({
        "quán trong file gốc": len(master),
        "đã có place_id": sum(1 for r in rec.values() if r.get("place_id")),
        "không tìm thấy": sum(1 for r in rec.values() if r.get("id_status") == "not_found"),
        "đã kiểm tra (verify)": sum(1 for r in rec.values() if r.get("verified_at")),
        "  khớp tên + vị trí": sum(1 for r in rec.values() if r.get("match") == "ok"),
        "  cần kiểm tra tay": sum(1 for r in rec.values() if r.get("match") == "check"),
        "đã có đánh giá/giá": sum(1 for r in rec.values() if r.get("rated_at")),
        "  có khoảng giá (VNĐ)": sum(1 for r in rec.values() if r.get("price_min")),
        "  có bình luận": sum(1 for r in rec.values() if r.get("reviews")),
    })
    print(s.to_string())
    print(f"\nLượt gọi tháng {usage.month}:")
    for k, (cap, price) in SKU.items():
        print(f"  {k:32s} {usage.used(k):6d} / {'∞' if cap is None else cap} miễn phí")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("step", choices=["ids", "verify", "reviews", "export", "status"])
    ap.add_argument("--max", type=int, help="số lượt gọi tối đa lần chạy này")
    ap.add_argument("--limit", type=int, help="chỉ xét N quán đầu của file gốc (chạy thử)")
    ap.add_argument("--allow-paid", action="store_true", help="cho phép vượt mức miễn phí của tháng (tính phí)")
    ap.add_argument("--no-reviews", action="store_true", help="bước reviews: chỉ lấy đánh giá + giá (SKU Enterprise, $20/1.000)")
    ap.add_argument("--sleep", type=float, default=0.1, help="nghỉ giữa các lượt gọi (giây)")
    ap.add_argument("--dry-run", action="store_true", help="in yêu cầu mẫu, không gọi API")
    args = ap.parse_args()
    if args.step == "ids":
        step_ids(args)
    elif args.step == "verify":
        step_details(args, "pro")
    elif args.step == "reviews":
        step_details(args, "ent" if args.no_reviews else "atm")
    elif args.step == "export":
        step_export(args)
    else:
        step_status(args)


if __name__ == "__main__":
    main()
