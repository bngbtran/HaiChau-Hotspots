"""Crawl bình luận Google Maps bằng Apify actor compass/google-maps-reviews-scraper cho phần việc của một thành viên.

Đầu vào: một sheet của Full_dataset_FB_4nguoi_40acc.xlsx (mặc định sheet "Trân"). Sheet chia thành các "Nhóm",
mỗi nhóm ~8.000 bình luận = chạy bằng MỘT tài khoản Apify (gói Free có $5 / tháng). Mỗi quán lấy tối đa 1.035 bình luận.

Các bước (chạy lại được, nhóm đã xong sẽ bỏ qua; run đang chạy dở sẽ được theo dõi tiếp, không chạy lại tốn tiền):
  plan     In danh sách nhóm, số quán, số bình luận dự kiến.                     Không gọi API
  check    Kiểm tra từng token: đúng tài khoản không, tháng này đã dùng bao nhiêu $.
  run      Chạy actor cho từng nhóm, tải kết quả về data/crawl/apify/<tên>/raw/<nhóm>.json
  export   Mỗi nhóm một file: data/crawl/apify/<tên>/xlsx/<tên>_<nhóm>.xlsx và csv/<tên>_<nhóm>.csv
  status   Bảng tiến độ: nhóm nào xong, bao nhiêu bình luận / dự kiến

Chuẩn bị token (KHÔNG ghi token vào code, không commit): tạo file data/crawl/apify/tokens.json (đã có trong .gitignore)
    {"Nhóm 2": "apify_api_xxx", "Nhóm 7": "apify_api_yyy", ...}
Token lấy ở https://console.apify.com/settings/integrations của từng tài khoản.

Ví dụ:
    python scripts/crawl_apify_reviews.py plan
    python scripts/crawl_apify_reviews.py check
    python scripts/crawl_apify_reviews.py run --group "Nhóm 2" --test 1   # chạy thử 1 quán, 20 bình luận
    python scripts/crawl_apify_reviews.py run                             # chạy tất cả nhóm có token
    python scripts/crawl_apify_reviews.py export
"""
import argparse
import json
import re
import sys
import time
import unicodedata
from pathlib import Path

import pandas as pd
import requests

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_XLSX = ROOT / "downloads" / "Full_dataset_FB_4nguoi_40acc.xlsx"
APIFY_DIR = ROOT / "data" / "crawl" / "apify"
TOKENS_FILE = APIFY_DIR / "tokens.json"
API = "https://api.apify.com/v2"
ACTOR = "compass~google-maps-reviews-scraper"
MAX_PER_PLACE = 1035
REVIEW_COLS = ["Nhóm", "placeId", "title", "stars", "text", "textTranslated", "publishedAtDate", "likesCount",
               "responseFromOwnerText", "responseFromOwnerDate", "originalLanguage", "reviewOrigin", "reviewId",
               "reviewImageUrls", "reviewContext", "reviewDetailedRating", "totalScore", "reviewsCount", "categoryName",
               "address", "lat", "lng"]


# ---------- tiện ích ----------
def slug(s):
    s = unicodedata.normalize("NFD", str(s)).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9]+", "_", s.replace("đ", "d")).strip("_")


def load_sheet(xlsx, sheet):
    """Các dòng quán của sheet (bỏ dòng 'Tổng nhóm' và dòng trống), kèm placeId tách từ url."""
    d = pd.read_excel(xlsx, sheet_name=sheet)
    d = d[d["Nhóm"].notna() & d["url"].notna()].copy()
    d["placeId"] = d["url"].str.extract(r"query_place_id=([A-Za-z0-9_-]+)", expand=False)
    cap = next(c for c in d.columns if c.startswith("reviewsCount (max"))
    d["expected"] = d[cap].fillna(0).astype(int)
    bad = d["placeId"].isna().sum()
    if bad:
        print(f"Cảnh báo: {bad} dòng không tách được placeId từ url → bỏ qua")
    d = d[d["placeId"].notna()]
    order = sorted(d["Nhóm"].unique(), key=lambda g: int(re.sub(r"\D", "", g) or 0))
    return d, order


def out_dir(sheet):
    return APIFY_DIR / slug(sheet)


def load_json(path, default):
    return json.loads(path.read_text(encoding="utf-8")) if path.exists() else default


def save_json(path, obj):
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(obj, ensure_ascii=False, indent=1), encoding="utf-8")
    tmp.replace(path)


def load_tokens():
    if not TOKENS_FILE.exists():
        sys.exit(f"Chưa có {TOKENS_FILE}. Tạo file dạng {{\"Nhóm 2\": \"apify_api_...\"}} (xem đầu file).")
    return load_json(TOKENS_FILE, {})


def api(method, path, token, **kw):
    for attempt in range(5):
        r = requests.request(method, f"{API}{path}", headers={"Authorization": f"Bearer {token}"}, timeout=90, **kw)
        if r.status_code in (200, 201):
            return r.json()
        if r.status_code in (429, 500, 502, 503, 504):
            time.sleep(2 ** attempt)
            continue
        sys.exit(f"Lỗi {r.status_code} từ Apify ({path}): {r.text[:300]}")
    sys.exit(f"Apify không phản hồi ({path})")


# ---------- plan / check ----------
def step_plan(args):
    d, order = load_sheet(args.xlsx, args.sheet)
    tokens = load_json(TOKENS_FILE, {})
    t = d.groupby("Nhóm").agg(so_quan=("placeId", "size"), binh_luan_du_kien=("expected", "sum")).loc[order]
    t["có token"] = ["✓" if g in tokens else "" for g in t.index]
    print(t.to_string())
    print(f"\nTổng: {len(d)} quán · {d.expected.sum():,} bình luận dự kiến · {len(order)} nhóm = {len(order)} tài khoản Apify")


def step_check(args):
    _, order = load_sheet(args.xlsx, args.sheet)
    tokens = load_tokens()
    seen = {}
    for g in order:
        if g not in tokens:
            print(f"{g:9s} — chưa có token"); continue
        for t in token_list(tokens, g):
            me = api("GET", "/users/me", t)["data"]
            lim = api("GET", "/users/me/limits", t)["data"]
            used, cap = lim["current"]["monthlyUsageUsd"], lim["limits"]["maxMonthlyUsageUsd"]
            dup = f"  ⚠ trùng tài khoản với {seen[me['id']]}" if me["id"] in seen else ""
            seen.setdefault(me["id"], g)
            print(f"{g:9s} {me['username']:25s} đã dùng ${used:.2f} / ${cap:.2f} tháng này{dup}")


# ---------- run ----------
def run_input(place_ids, max_reviews, args):
    return {"placeIds": place_ids, "maxReviews": max_reviews, "reviewsSort": args.sort, "language": args.language,
            "reviewsOrigin": "google", "personalData": False}


def crawled_counts(base):
    """Số bình luận (không trùng) đã lấy cho từng placeId, gộp mọi file raw trừ file chạy thử;
    kèm số bình luận hiện có trên Google lúc crawl."""
    ids, now = {}, {}
    for f in (base / "raw").glob("*.json"):
        if not f.stem.endswith("_test"):
            for it in load_json(f, []):
                ids.setdefault(it.get("placeId"), set()).add(it.get("reviewId"))
                if it.get("reviewsCount") is not None:
                    now[it.get("placeId")] = max(now.get(it.get("placeId"), 0), it["reviewsCount"])
    return {k: len(v) for k, v in ids.items()}, now


def credit_left(token):
    lim = api("GET", "/users/me/limits", token)["data"]
    return lim["limits"]["maxMonthlyUsageUsd"] - lim["current"]["monthlyUsageUsd"]


def token_list(tokens, g):
    """Một nhóm có thể có một token hoặc danh sách token (tài khoản phụ để chạy bù khi tài khoản chính hết tiền)."""
    t = tokens.get(g)
    return [t] if isinstance(t, str) else list(t or [])


def fetch_items(token, dataset_id):
    items, offset = [], 0
    while True:
        page = requests.get(f"{API}/datasets/{dataset_id}/items", headers={"Authorization": f"Bearer {token}"},
                            params={"format": "json", "clean": "true", "offset": offset, "limit": 1000}, timeout=120).json()
        items += page
        if len(page) < 1000:
            return items
        offset += 1000


def step_run(args):
    d, order = load_sheet(args.xlsx, args.sheet)
    tokens = load_json(TOKENS_FILE, {}) if args.dry_run else load_tokens()
    base = out_dir(args.sheet)
    runs_file = base / "runs.json"
    runs = load_json(runs_file, {})
    groups = [args.group] if args.group else order
    for g in groups:
        if g not in order:
            sys.exit(f"Không có {g!r} trong sheet {args.sheet}. Các nhóm: {', '.join(order)}")
        key = f"{g} (test)" if args.test else f"{g} (retry {args.sort})" if args.retry else g
        raw = base / "raw" / f"{slug(key)}.json"
        if raw.exists() and not args.force:
            print(f"{key}: đã có {raw.name}, bỏ qua (thêm --force để chạy lại)"); continue
        if g not in tokens and not args.dry_run:
            print(f"{g}: chưa có token trong tokens.json, bỏ qua"); continue
        toks, places = token_list(tokens, g), d[d["Nhóm"] == g]
        if args.test:
            places = places.head(args.test)
        if args.retry:   # chỉ chạy lại quán lấy thiếu so với dự kiến
            got, now = crawled_counts(base)
            places = places.assign(expected=[min(e, now.get(p, e)) for p, e in zip(places["placeId"], places["expected"])])
            places = places[places["placeId"].map(lambda p: got.get(p, 0)) < places["expected"]]
            if places.empty:
                print(f"{g}: không còn quán nào thiếu"); continue
            for _, p in places.iterrows():
                print(f"  thiếu: {p.title} · {got.get(p.placeId, 0)}/{p.expected}")
        max_reviews = 20 if args.test else MAX_PER_PLACE
        r = runs.get(key)
        if r and r.get("status") in ("READY", "RUNNING"):
            token = toks[r.get("tok", 0)]
            print(f"{key}: theo dõi tiếp run {r['run_id']}")
        else:
            body = run_input(places["placeId"].tolist(), max_reviews, args)
            if args.dry_run:
                print(f"[dry-run] {key}: {len(places)} quán, maxReviews={max_reviews}, giới hạn ${args.max_usd}\n  {json.dumps(body, ensure_ascii=False)[:300]}…")
                continue
            # ưu tiên tài khoản đủ tiền cho cả run (≈ $0.6 / 1.000 bình luận); không có thì lấy tài khoản còn tiền đầu tiên
            need = len(places) * max_reviews * 0.0006 + 0.05
            lefts = [credit_left(t) for t in toks]
            tok = next((i for i, l in enumerate(lefts) if l >= need), next((i for i, l in enumerate(lefts) if l > 0.15), None))
            max_usd = round(min(args.max_usd, lefts[tok] - 0.05), 2) if tok is not None else 0
            if tok is None:
                print(f"{g}: mọi tài khoản của nhóm đã hết tiền tháng này, bỏ qua (thêm token mới vào tokens.json)"); continue
            token = toks[tok]
            data = api("POST", f"/acts/{ACTOR}/runs", token, json=body, params={"maxTotalChargeUsd": max_usd})["data"]
            r = runs[key] = {"run_id": data["id"], "dataset_id": data["defaultDatasetId"], "status": data["status"],
                             "places": len(places), "expected": int(places["expected"].clip(upper=max_reviews).sum()), "tok": tok}
            print(f"  tài khoản #{tok + 1} của nhóm · giới hạn chi phí run này: ${max_usd}")
            save_json(runs_file, runs)
            print(f"{key}: bắt đầu run {r['run_id']} · {r['places']} quán · dự kiến {r['expected']:,} bình luận")
            print(f"  xem trên web: https://console.apify.com/actors/runs/{r['run_id']}")
        while True:
            data = api("GET", f"/actor-runs/{r['run_id']}", token, params={"waitForFinish": 60})["data"]
            r["status"] = data["status"]
            save_json(runs_file, runs)
            if data["status"] not in ("READY", "RUNNING"):
                break
            n = api("GET", f"/datasets/{r['dataset_id']}", token)["data"].get("itemCount", 0)
            print(f"  … {data['status']} · {n:,} bình luận")
        items = fetch_items(token, r["dataset_id"])
        save_json(raw, items)
        r.update(got=len(items), usd=data.get("usageTotalUsd"))
        save_json(runs_file, runs)
        print(f"{key}: {r['status']} · {len(items):,}/{r['expected']:,} bình luận · ${r.get('usd') or 0:.2f} → {raw.relative_to(ROOT)}")
        if r["status"] != "SUCCEEDED":
            print(f"  ⚠ Run kết thúc {r['status']}: kiểm tra log trên console. Dữ liệu một phần vẫn đã lưu; xoá {raw.name} rồi chạy lại nếu cần.")


# ---------- export / status ----------
def step_export(args):
    """Mỗi nhóm một file <tên>_<nhóm>.xlsx/.csv (vd. tran_nhom2), giữ MỌI trường Apify trả về."""
    d, order = load_sheet(args.xlsx, args.sheet)
    base = out_dir(args.sheet)
    files = [f for f in sorted((base / "raw").glob("*.json")) if not f.stem.endswith("_test")]
    if not files:
        sys.exit("Chưa có kết quả nào trong raw/. Chạy bước run trước.")
    group_of = d.set_index("placeId")["Nhóm"].to_dict()
    rows = []
    for f in files:
        for it in load_json(f, []):
            row = {k: json.dumps(v, ensure_ascii=False) if isinstance(v, (list, dict)) and v else (None if v in ([], {}) else v)
                   for k, v in it.items() if k != "location"}
            loc = it.get("location") or {}
            row.update(Nhóm=group_of.get(it.get("placeId")), lat=loc.get("lat"), lng=loc.get("lng"))
            rows.append(row)
    rv = pd.DataFrame(rows).drop_duplicates("reviewId")
    rv = rv[[c for c in REVIEW_COLS if c in rv] + sorted(c for c in rv if c not in REVIEW_COLS)]
    got = rv.groupby("placeId").size()
    now = pd.to_numeric(rv["reviewsCount"], errors="coerce").groupby(rv["placeId"]).max()   # số bình luận trên Google lúc crawl
    for g in ([args.group] if args.group else order):
        places = d.loc[d["Nhóm"] == g, ["Nhóm", "placeId", "title", "street", "totalScore", "expected", "url"]].copy()
        places["reviewsCount hiện tại"] = places["placeId"].map(now)
        # quán bị xoá bớt bình luận từ lúc lập danh sách: chỉ cần lấy đủ số hiện có
        places["expected"] = places[["expected", "reviewsCount hiện tại"]].min(axis=1).astype(int)
        places["crawled"] = places["placeId"].map(got).fillna(0).astype(int)
        if not places["crawled"].any():
            continue
        places["Trạng thái"] = ["Done" if n >= e else "Thiếu" if n else "Not Yet" for n, e in zip(places.crawled, places.expected)]
        grv = rv[rv["Nhóm"] == g]
        name = f"{slug(args.sheet)}_{slug(g).replace('_', '')}"
        out, out_csv = base / "xlsx" / f"{name}.xlsx", base / "csv" / f"{name}.csv"
        out.parent.mkdir(exist_ok=True)
        out_csv.parent.mkdir(exist_ok=True)
        try:
            with pd.ExcelWriter(out, engine="openpyxl") as w:
                places.to_excel(w, sheet_name="Quán", index=False)
                grv.to_excel(w, sheet_name="Bình luận", index=False)
            grv.to_csv(out_csv, index=False, encoding="utf-8-sig")
        except PermissionError as e:
            sys.exit(f"Không ghi được {Path(e.filename).name}: file đang mở trong Excel. Đóng file rồi chạy lại export.")
        miss = places[places.crawled < places.expected]
        print(f"Đã ghi {out.relative_to(ROOT)} + {out_csv.relative_to(ROOT)}: {len(grv):,} bình luận · {grv.shape[1]} cột · "
              f"{(places.crawled >= places.expected).sum()}/{len(places)} quán đủ"
              + (f" · thiếu: {', '.join(miss.title)} → chạy run --group \"{g}\" --retry --sort newest --language en"
                 f" (vẫn thiếu thì thử --sort mostRelevant / lowestRanking)" if len(miss) else ""))


def step_status(args):
    d, order = load_sheet(args.xlsx, args.sheet)
    runs = load_json(out_dir(args.sheet) / "runs.json", {})
    t = pd.DataFrame([{"Nhóm": g, "quán": (d["Nhóm"] == g).sum(), "dự kiến": d.loc[d["Nhóm"] == g, "expected"].sum(),
                       "trạng thái": runs.get(g, {}).get("status", "chưa chạy"), "đã lấy": runs.get(g, {}).get("got"),
                       "$": runs.get(g, {}).get("usd")} for g in order])
    print(t.to_string(index=False))


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("step", choices=["plan", "check", "run", "export", "status"])
    ap.add_argument("--xlsx", type=Path, default=DEFAULT_XLSX)
    ap.add_argument("--sheet", default="Trân")
    ap.add_argument("--group", help='chỉ chạy một nhóm, ví dụ "Nhóm 2"')
    ap.add_argument("--test", type=int, metavar="N", help="chạy thử N quán đầu của nhóm, 20 bình luận/quán")
    ap.add_argument("--retry", action="store_true", help="chỉ chạy lại quán lấy thiếu so với dự kiến (gộp, bỏ trùng khi export)")
    ap.add_argument("--sort", default="newest", choices=["newest", "mostRelevant", "highestRanking", "lowestRanking"],
                    help="thứ tự lấy bình luận")
    ap.add_argument("--language", default="vi", help="ngôn ngữ hiển thị/ dịch bình luận (vi, en…)")
    ap.add_argument("--max-usd", type=float, default=4.8, help="trần chi phí mỗi run (USD), actor tự dừng khi chạm")
    ap.add_argument("--force", action="store_true", help="chạy lại nhóm đã có kết quả")
    ap.add_argument("--dry-run", action="store_true", help="in input sẽ gửi, không gọi API")
    args = ap.parse_args()
    {"plan": step_plan, "check": step_check, "run": step_run, "export": step_export, "status": step_status}[args.step](args)


if __name__ == "__main__":
    main()
