// ==UserScript==
// @name         HaiChau Collector · Google Maps
// @namespace    haichau-hotspots
// @version      1.0.0
// @description  Hỗ trợ nhóm DAM501 thu thập đánh giá, khoảng giá, bình luận và toạ độ của quán đang mở trên Google Maps, khớp với storeId và xuất Excel cho scripts/merge_crawl.py
// @match        https://www.google.com/maps*
// @match        https://www.google.com.vn/maps*
// @require      https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js
// @grant        GM_getValue
// @grant        GM_setValue
// @run-at       document-idle
// @noframes
// ==/UserScript==

/* Cách dùng (chi tiết trong data/README.md):
   1. Bấm "Nạp danh sách" → chọn file Excel quán được giao (sheet đầu có cột storeId, storeName, address, latitude, longitude).
   2. Bấm "Tìm trên Maps" (Alt+N) → chọn đúng quán trong kết quả nếu Google chưa tự mở.
   3. Kiểm tra các ô đã tự điền (sửa tay nếu cần) → "Lấy bình luận" (Alt+R) → "Lưu" (Alt+S), công cụ tự mở quán tiếp theo.
   4. Xong (hoặc giữa chừng) bấm "Xuất Excel" → chép file vào data/crawl/parts/.
   Công cụ chỉ đọc nội dung trang bạn đang xem; không tự gửi yêu cầu hàng loạt tới Google. */
(function () {
  'use strict';
  const XLSX_ = typeof XLSX !== 'undefined' ? XLSX : window.XLSX;
  const KEY = 'hch-collector-v1';
  const SOURCE = 'googlemaps';
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  /* ---------- Lưu trạng thái (GM storage, còn nguyên khi tải lại trang) ---------- */
  const gmGet = typeof GM_getValue === 'function' ? GM_getValue : (k, d) => { try { return JSON.parse(localStorage.getItem(k)) ?? d; } catch (e) { return d; } };
  const gmSet = typeof GM_setValue === 'function' ? GM_setValue : (k, v) => localStorage.setItem(k, JSON.stringify(v));
  const S = Object.assign({ stores: [], cols: [], idx: 0, results: {}, reviews: {}, member: '', fileName: '', maxReviews: 10, autoNext: true, min: false }, gmGet(KEY, {}));
  const save = () => gmSet(KEY, S);

  /* ---------- Chuẩn hoá chữ, so khớp tên + khoảng cách ---------- */
  const norm = (s) => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'd').toLowerCase();
  const STOP = new Set(['quan', 'nha', 'hang', 'tiem', 'cua', 'shop', 'cafe', 'coffee', 'ca', 'phe', 'the', 'and', 'va', 'da', 'nang', 'chi', 'nhanh', 'co', 'so', 'cs']);
  const tokens = (s) => norm(s).replace(/[^a-z0-9 ]/g, ' ').split(/\s+/).filter((t) => t && !STOP.has(t));
  function nameSim(a, b) {
    const A = new Set(tokens(a)), B = new Set(tokens(b));
    if (!A.size || !B.size) return 0;
    let inter = 0; A.forEach((t) => { if (B.has(t)) inter++; });
    return inter / Math.min(A.size, B.size);   // tỷ lệ chứa: "Bún chả cá 109" ⊂ "Bún Chả Cá 109 - Nguyễn Chí Thanh"
  }
  function distM(la1, ln1, la2, ln2) {
    const R = 6371000, r = Math.PI / 180, dLa = (la2 - la1) * r, dLn = (ln2 - ln1) * r;
    const h = Math.sin(dLa / 2) ** 2 + Math.cos(la1 * r) * Math.cos(la2 * r) * Math.sin(dLn / 2) ** 2;
    return 2 * R * Math.asin(Math.sqrt(h));
  }
  const num = (v) => { const x = parseFloat(String(v ?? '').replace(',', '.')); return Number.isFinite(x) ? x : null; };

  /* ---------- Đọc thông tin quán đang mở trên Google Maps ---------- */
  // "100–200 N ₫", "₫1–100 N", "200.000–300.000 ₫", "1-100.000 ₫", "Trên 1 Tr ₫" → [min, max] (VNĐ)
  function parsePriceRange(text) {
    const unit = (u) => !u ? 1 : /^tr|triệu/i.test(u) ? 1e6 : 1e3;
    const val = (n, u) => { const s = String(n); const v = u ? parseFloat(s.replace(/\./g, '').replace(',', '.')) : parseFloat(s.replace(/[.,]/g, '')); return Math.round(v * unit(u)); };
    for (const seg of String(text || '').split(/\n|·/)) {
      if (!/₫|vnd|đ\b/i.test(seg)) continue;
      const s = seg.replace(/₫|vnđ|vnd/gi, ' ').replace(/\s+/g, ' ');
      let m = s.match(/(\d[\d.,]*)\s*(N|Tr|K|nghìn|ngàn|triệu)?\s*[–—-]\s*(\d[\d.,]*)\s*(N|Tr|K|nghìn|ngàn|triệu)?/i);
      if (m) {
        const u2 = m[4], u1 = m[2] || (u2 && !/[.,]\d{3}/.test(m[1]) ? u2 : '');
        const lo = val(m[1], u1), hi = val(m[3], u2);
        if (lo > 0 && hi >= lo) return [lo < 1000 ? lo * 1000 : lo, hi < 1000 ? hi * 1000 : hi];
      }
      m = s.match(/(?:trên|từ|over)\s*(\d[\d.,]*)\s*(N|Tr|K|nghìn|triệu)?/i);
      if (m) { const lo = val(m[1], m[2]); if (lo > 0) return [lo, null]; }
    }
    return [null, null];
  }
  // "2 tuần trước", "một năm trước", "Đã chỉnh sửa 3 tháng trước" → YYYY-MM-DD (ước lượng)
  function relDate(text) {
    const t = norm(text);
    if (/vua xong|vua moi/.test(t)) return new Date().toISOString().slice(0, 10);
    const m = t.match(/(\d+|mot|hai|ba)\s*(phut|gio|ngay|tuan|thang|nam)\s*truoc/);
    if (!m) { const d = t.match(/(\d{1,2})[/-](\d{1,2})[/-](\d{4})/); return d ? `${d[3]}-${d[2].padStart(2, '0')}-${d[1].padStart(2, '0')}` : ''; }
    const n = { mot: 1, hai: 2, ba: 3 }[m[1]] || parseInt(m[1], 10);
    const days = { phut: 0, gio: 0, ngay: 1, tuan: 7, thang: 30, nam: 365 }[m[2]] * n;
    const d = new Date(); d.setDate(d.getDate() - days); return d.toISOString().slice(0, 10);
  }
  const isPlacePage = () => /\/maps\/place\//.test(location.pathname);
  function readCoords() {
    const u = decodeURIComponent(location.href);
    let m = u.match(/!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)/g);
    if (m) { const last = m[m.length - 1].match(/!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)/); return { lat: +last[1], lng: +last[2], exact: true }; }
    m = u.match(/@(-?\d+\.\d+),(-?\d+\.\d+)/);
    return m ? { lat: +m[1], lng: +m[2], exact: false } : { lat: null, lng: null, exact: false };
  }
  function mainPanel() {
    const h1 = document.querySelector('h1');
    return h1?.closest('div[role="main"]') || document.querySelector('div[role="main"]') || document.body;
  }
  function readPlace() {
    const main = mainPanel(), out = { name: '', rating: null, ratingCount: null, priceMin: null, priceMax: null, address: '', closed: '', ...readCoords() };
    out.name = (main.querySelector('h1')?.innerText || '').trim();
    // Điểm: span aria-label "4,2 sao" hoặc số trong khối F7nice
    const starEl = [...main.querySelectorAll('[aria-label]')].find((e) => /^\s*\d[.,]\d\s*sao/i.test(e.getAttribute('aria-label')));
    out.rating = num(starEl?.getAttribute('aria-label')) ?? num(main.querySelector('div.F7nice span[aria-hidden="true"]')?.textContent);
    // Số đánh giá: aria-label "1.234 bài đánh giá" hoặc "(1.234)" cạnh điểm
    const cntEl = [...main.querySelectorAll('[aria-label]')].find((e) => /\d[\d.,]*\s*(bài )?đánh giá|reviews?/i.test(e.getAttribute('aria-label')) && !/viết|write/i.test(e.getAttribute('aria-label')));
    const cntTxt = cntEl?.getAttribute('aria-label') || (main.querySelector('div.F7nice')?.textContent.match(/\(([\d.,]+)\)/) || [])[1];
    if (cntTxt) { const c = parseInt(String(cntTxt).replace(/[^\d]/g, ''), 10); if (c > 0) out.ratingCount = c; }
    // Khoảng giá: phần tử có aria-label "Giá…" hoặc dòng có ₫ trong khung tổng quan / "Giá mỗi người"
    const priceAria = [...main.querySelectorAll('[aria-label]')].map((e) => e.getAttribute('aria-label')).filter((a) => /giá|price/i.test(a) && /\d/.test(a)).join('\n');
    const perPerson = (main.innerText.match(/Giá mỗi người[\s\S]{0,80}/i) || [''])[0];
    [out.priceMin, out.priceMax] = parsePriceRange(priceAria);
    if (out.priceMin == null) [out.priceMin, out.priceMax] = parsePriceRange(perPerson);
    if (out.priceMin == null) [out.priceMin, out.priceMax] = parsePriceRange(main.innerText.slice(0, 1500));
    out.address = (main.querySelector('button[data-item-id="address"]')?.getAttribute('aria-label') || '').replace(/^Địa chỉ:\s*/i, '').trim();
    const t = main.innerText;
    out.closed = /đóng cửa vĩnh viễn|permanently closed/i.test(t) ? 'Đóng cửa' : /tạm thời đóng cửa|temporarily closed/i.test(t) ? 'Tạm đóng cửa' : '';
    return out;
  }
  // Bình luận: mở tab "Bài đánh giá", cuộn danh sách tới khi đủ số lượng, mở rộng "Thêm" rồi đọc
  async function collectReviews(max) {
    const tab = [...document.querySelectorAll('button[role="tab"]')].find((b) => /đánh giá|reviews/i.test((b.getAttribute('aria-label') || '') + b.innerText));
    if (!tab) return { list: [], note: 'Không thấy tab "Bài đánh giá" (đăng nhập Google để xem đầy đủ).' };
    if (tab.getAttribute('aria-selected') !== 'true') { tab.click(); await sleep(1800); }
    const cards = () => [...document.querySelectorAll('div[data-review-id]')].filter((e) => e.querySelector('[role="img"][aria-label*="sao"], [role="img"][aria-label*="star"]'));
    for (let k = 0; k < 8 && cards().length < max; k++) {
      const last = cards().pop();
      const box = last && (() => { let e = last.parentElement; while (e && e !== document.body && !(e.scrollHeight > e.clientHeight + 20 && /auto|scroll/.test(getComputedStyle(e).overflowY))) e = e.parentElement; return e; })();
      if (!box || box === document.body) break;
      box.scrollTop = box.scrollHeight; await sleep(1200);
    }
    cards().slice(0, max).forEach((c) => c.querySelectorAll('button').forEach((b) => { if (/^(thêm|xem thêm|more|see more)$/i.test((b.getAttribute('aria-label') || b.innerText || '').trim())) b.click(); }));
    await sleep(500);
    const seen = new Set(), list = [];
    for (const c of cards()) {
      const id = c.getAttribute('data-review-id'); if (seen.has(id)) continue; seen.add(id);
      const stars = parseInt((c.querySelector('[role="img"][aria-label*="sao"], [role="img"][aria-label*="star"]')?.getAttribute('aria-label') || '').match(/\d/)?.[0], 10);
      const textEl = c.querySelector('.wiI7pd') || [...c.querySelectorAll('span')].sort((a, b) => b.innerText.length - a.innerText.length)[0];
      const comment = (textEl?.innerText || '').replace(/\s+/g, ' ').trim();
      const dateTxt = c.querySelector('.rsqaWe')?.innerText || [...c.querySelectorAll('span')].map((s) => s.innerText).find((x) => /trước|vừa xong|ago/i.test(x) && x.length < 40) || '';
      if (stars >= 1 && stars <= 5 && comment) list.push({ rating: stars, comment, date: relDate(dateTxt) });
      if (list.length >= max) break;
    }
    return { list, note: '' };
  }

  /* ---------- Excel ---------- */
  const NEW_COLS = ['Trạng thái', 'Đánh giá', 'Số đánh giá', 'Giá thấp nhất', 'Giá cao nhất', 'Vĩ độ mới', 'Kinh độ mới', 'Link Google Maps', 'Nguồn'];
  function importFile(file) {
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const wb = XLSX_.read(new Uint8Array(reader.result), { type: 'array' });
        const rows = XLSX_.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '', raw: true });
        if (!rows.length || !('storeId' in rows[0])) throw new Error('Sheet đầu tiên phải có cột storeId');
        const cols = Object.keys(rows[0]).filter((c) => !NEW_COLS.includes(c));
        const results = {}, reviews = {};
        rows.forEach((r) => {   // file đã làm dở: nạp lại kết quả cũ
          const st = String(r['Trạng thái'] || '').trim();
          if (st && st.toLowerCase() !== 'not yet') results[r.storeId] = { status: st, rating: num(r['Đánh giá']), ratingCount: num(r['Số đánh giá']), priceMin: num(r['Giá thấp nhất']), priceMax: num(r['Giá cao nhất']), lat: num(r['Vĩ độ mới']), lng: num(r['Kinh độ mới']), url: r['Link Google Maps'] || '' };
        });
        const rs = wb.SheetNames.find((n) => /bình luận|binh luan|reviews/i.test(n));
        if (rs) XLSX_.utils.sheet_to_json(wb.Sheets[rs], { defval: '' }).forEach((r) => { (reviews[r.storeId] = reviews[r.storeId] || []).push({ rating: num(r['Số sao']), comment: r['Bình luận'], date: r['Ngày'] }); });
        Object.assign(S, { stores: rows.map((r) => Object.fromEntries(cols.map((c) => [c, r[c]]))), cols, results, reviews, fileName: file.name, member: S.member || file.name.replace(/\.xlsx?$/i, '') });
        S.idx = Math.max(0, S.stores.findIndex((s) => !S.results[s.storeId]));
        save(); render(); toast(`Đã nạp ${rows.length} quán từ ${file.name}`);
      } catch (e) { toast('Không đọc được file: ' + e.message, true); }
    };
    reader.readAsArrayBuffer(file);
  }
  function exportFile() {
    if (!S.stores.length) return toast('Chưa nạp danh sách', true);
    const member = (prompt('Tên file (tên thành viên, không dấu):', S.member || 'thanhvien') || '').trim();
    if (!member) return;
    S.member = member; save();
    const rows = S.stores.map((s) => {
      const r = S.results[s.storeId] || {};
      return { ...s, 'Trạng thái': r.status || 'Not Yet', 'Đánh giá': r.rating ?? '', 'Số đánh giá': r.ratingCount ?? '', 'Giá thấp nhất': r.priceMin ?? '', 'Giá cao nhất': r.priceMax ?? '',
        'Vĩ độ mới': r.lat ?? '', 'Kinh độ mới': r.lng ?? '', 'Link Google Maps': r.url || '', 'Nguồn': r.status ? SOURCE : '' };
    });
    const revRows = [];
    S.stores.forEach((s) => (S.reviews[s.storeId] || []).forEach((v) => revRows.push({ storeId: s.storeId, 'Số sao': v.rating, 'Bình luận': v.comment, 'Ngày': v.date || '', 'Nguồn': SOURCE })));
    const wb = XLSX_.utils.book_new();
    XLSX_.utils.book_append_sheet(wb, XLSX_.utils.json_to_sheet(rows, { header: [...S.cols, ...NEW_COLS] }), 'Quán');
    XLSX_.utils.book_append_sheet(wb, XLSX_.utils.json_to_sheet(revRows, { header: ['storeId', 'Số sao', 'Bình luận', 'Ngày', 'Nguồn'] }), 'Bình luận');
    XLSX_.writeFile(wb, `${member}.xlsx`);
    toast(`Đã xuất ${member}.xlsx · chép vào data/crawl/parts/`);
  }

  /* ---------- Điều hướng ---------- */
  const cur = () => S.stores[S.idx];
  const nextTodo = (from) => { for (let i = from + 1; i < S.stores.length; i++) if (!S.results[S.stores[i].storeId]) return i; return -1; };
  function openSearch() {
    const s = cur(); if (!s) return;
    const addr = String(s.address || '').split(',').slice(0, 2).join(',');
    const lat = num(s.latitude), lng = num(s.longitude);
    location.href = `https://www.google.com/maps/search/${encodeURIComponent(`${s.storeName} ${addr}`.trim())}` + (lat && lng ? `/@${lat},${lng},17z` : '') + '?hl=vi';
  }
  function go(i, search = true) { if (i < 0 || i >= S.stores.length) return; S.idx = i; save(); dirty = {}; render(); if (search) openSearch(); }

  /* ---------- Giao diện (shadow DOM để không lẫn CSS của Google) ---------- */
  const host = document.createElement('div'); host.id = 'hch-collector';
  const root = host.attachShadow({ mode: 'open' });
  document.documentElement.appendChild(host);
  root.innerHTML = `<style>
    :host{all:initial}
    .card{position:fixed;right:72px;bottom:24px;z-index:2147483000;width:340px;max-height:calc(100vh - 48px);overflow:auto;background:#fff;color:#1F3A4D;border-radius:14px;box-shadow:0 12px 36px rgba(17,23,29,.28);font:13px/1.45 Inter,system-ui,-apple-system,"Segoe UI",Roboto,sans-serif}
    .hd{display:flex;align-items:center;gap:8px;padding:10px 12px;background:#11171D;color:#F3EEE6;border-radius:14px 14px 0 0;cursor:default}
    .hd b{font-size:13.5px}.hd b span{color:#7FB7D6}.hd .sp{flex:1}
    .dots{display:grid;grid-template-columns:repeat(3,5px);gap:2px}.dots i{width:5px;height:5px;border-radius:50%}
    .ib{border:0;background:transparent;color:#9AA8B4;cursor:pointer;font-size:16px;line-height:1;padding:2px 4px;border-radius:6px}.ib:hover{color:#fff;background:#1F2A35}
    .bd{padding:12px}
    .prog{height:6px;background:#ECE5D8;border-radius:99px;overflow:hidden;margin:2px 0 6px}.prog i{display:block;height:100%;background:#4A90B5;transition:width .3s}
    .muted{color:#56697A;font-size:12px}
    .store{background:#F8F4EC;border:1px solid #ECE5D8;border-radius:10px;padding:9px 10px;margin:8px 0}
    .store .nm{font-weight:600;font-size:14px}
    .row{display:flex;gap:6px;align-items:center}
    .match{display:inline-block;font-size:11.5px;font-weight:600;padding:2px 8px;border-radius:99px;margin-top:6px}
    .ok{background:#E2F3EF;color:#1D6B5E}.warn{background:#FCE7DB;color:#A4441A}.none{background:#ECE5D8;color:#56697A}
    .grid{display:grid;grid-template-columns:1fr 1fr;gap:6px 8px;margin:8px 0}
    label{display:block;font-size:11px;color:#56697A}
    input[type=text],input[type=number]{width:100%;box-sizing:border-box;border:1px solid #DDD2C0;border-radius:7px;padding:5px 7px;font:inherit;color:#1F3A4D;background:#fff}
    input.auto{background:#F1F8FB;border-color:#BFDCEA}
    button.b{font:inherit;border:1px solid #DDD2C0;background:#fff;color:#1F3A4D;border-radius:8px;padding:6px 9px;cursor:pointer;white-space:nowrap}
    button.b:hover{background:#F8F4EC}
    button.p{background:#2E6E92;border-color:#2E6E92;color:#fff;font-weight:600;flex:1}button.p:hover{background:#2A6688}
    button.w{color:#A4441A}
    .revs{margin:6px 0 8px;font-size:12px}
    .revs li{margin:3px 0;color:#56697A}
    .ft{display:flex;gap:6px;flex-wrap:wrap;border-top:1px solid #ECE5D8;padding-top:9px;margin-top:4px}
    .kbd{font-size:10.5px;color:#9AA8B4}
    .toast{position:fixed;right:72px;bottom:8px;z-index:2147483001;background:#1F3A4D;color:#fff;padding:7px 12px;border-radius:99px;font:12.5px Inter,system-ui,sans-serif;box-shadow:0 4px 14px rgba(0,0,0,.25)}
    .toast.err{background:#A4441A}
    .mini{width:auto}
  </style><div class="card" id="card"></div><input type="file" id="file" accept=".xlsx,.xls" hidden>`;
  const $ = (s) => root.querySelector(s);
  $('#file').addEventListener('change', (e) => { const f = e.target.files[0]; e.target.value = ''; if (f) importFile(f); });
  let toastT = 0;
  function toast(msg, err) {
    root.querySelectorAll('.toast').forEach((t) => t.remove());
    const t = document.createElement('div'); t.className = 'toast' + (err ? ' err' : ''); t.textContent = msg; root.appendChild(t);
    clearTimeout(toastT); toastT = setTimeout(() => t.remove(), 3500);
  }
  const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  let place = null, dirty = {}, pendingReviews = null;
  const FIELDS = [['rating', 'Đánh giá (sao)'], ['ratingCount', 'Số đánh giá'], ['priceMin', 'Giá thấp nhất (₫)'], ['priceMax', 'Giá cao nhất (₫)'], ['lat', 'Vĩ độ mới'], ['lng', 'Kinh độ mới']];
  function matchInfo(s, p) {
    if (!p || !isPlacePage() || !p.name) return { cls: 'none', txt: 'Chưa mở trang quán trên Maps' };
    const sim = nameSim(s.storeName, p.name), lat = num(s.latitude), lng = num(s.longitude);
    const d = p.lat != null && lat != null ? distM(lat, lng, p.lat, p.lng) : null;
    const dTxt = d == null ? '' : d < 1000 ? ` · lệch ${Math.round(d)} m` : ` · lệch ${(d / 1000).toFixed(1)} km`;
    if (sim >= 0.5 && (d == null || d <= 400)) return { cls: 'ok', txt: `Khớp: “${p.name}”${dTxt}` };
    return { cls: 'warn', txt: `Kiểm tra lại: “${p.name}”${dTxt}${sim < 0.5 ? ' · tên khác' : ''}` };
  }
  function render() {
    const card = $('#card');
    if (S.min) { card.className = 'card mini'; card.innerHTML = `<div class="hd" style="border-radius:14px"><div class="dots">${dots()}</div><b>HaiChau <span>Collector</span></b><button class="ib" id="max" title="Mở rộng">▢</button></div>`; $('#max').onclick = () => { S.min = false; save(); render(); }; return; }
    card.className = 'card';
    const total = S.stores.length, done = S.stores.filter((s) => S.results[s.storeId]).length, s = cur();
    let body;
    if (!total) {
      body = `<p>Nạp file Excel danh sách quán được giao (sheet đầu có cột <b>storeId</b>, storeName, address, latitude, longitude).</p>
        <div class="ft" style="border:0"><button class="b p" id="load">Nạp danh sách (.xlsx)</button></div>`;
    } else {
      const r = S.results[s.storeId], m = matchInfo(s, place), rv = pendingReviews ?? S.reviews[s.storeId] ?? [];
      const val = (k) => k in dirty ? dirty[k] : place && isPlacePage() && place[k] != null ? place[k] : r?.[k] ?? '';
      body = `<div class="row muted"><span>${done}/${total} quán đã xử lý</span><span style="flex:1"></span><span>${esc(S.fileName)}</span></div>
        <div class="prog"><i style="width:${total ? done / total * 100 : 0}%"></i></div>
        <div class="store"><div class="row"><span class="muted">#${S.idx + 1} · ${esc(s.storeId)}</span><span style="flex:1"></span>${r ? `<span class="match ok" style="margin:0">${esc(r.status)}</span>` : ''}</div>
          <div class="nm">${esc(s.storeName)}</div><div class="muted">${esc(s.address)}</div>
          <span class="match ${m.cls}">${esc(m.txt)}</span>${place?.closed ? ` <span class="match warn">${esc(place.closed)}</span>` : ''}</div>
        <div class="row"><button class="b" id="prev" title="Quán trước (Alt+B)">◀</button><button class="b" id="search" style="flex:1" title="Alt+N">Tìm trên Maps</button><button class="b" id="next" title="Quán chưa làm tiếp theo">▶</button></div>
        <div class="grid">${FIELDS.map(([k, l]) => `<div><label>${l}</label><input type="text" data-k="${k}" value="${esc(val(k))}" class="${!(k in dirty) && place && isPlacePage() && place[k] != null ? 'auto' : ''}"></div>`).join('')}</div>
        <div class="row"><button class="b" id="reread" title="Đọc lại trang">↻ Đọc lại</button><button class="b" id="revs" style="flex:1" title="Alt+R">Lấy bình luận (≤ <span id="maxr">${S.maxReviews}</span>)</button></div>
        <div class="revs">${rv.length ? `<b>${rv.length} bình luận</b><ul style="margin:4px 0 0;padding-left:16px">${rv.slice(0, 3).map((v) => `<li>${'★'.repeat(v.rating)} ${esc(String(v.comment).slice(0, 70))}${String(v.comment).length > 70 ? '…' : ''}</li>`).join('')}</ul>` : '<span class="muted">Chưa có bình luận</span>'}</div>
        <div class="row"><button class="b p" id="save" title="Alt+S">Lưu ✓</button><button class="b w" id="nf" title="Alt+K">Không tìm thấy</button><button class="b w" id="closed">Đóng cửa</button></div>
        <div class="ft"><button class="b" id="export">Xuất Excel</button><button class="b" id="load">Nạp danh sách</button>
          <label style="display:flex;align-items:center;gap:5px;font-size:12px;color:#1F3A4D"><input type="checkbox" id="auto" ${S.autoNext ? 'checked' : ''}> Tự mở quán tiếp</label></div>
        <div class="kbd">Alt+N tìm · Alt+R bình luận · Alt+S lưu · Alt+K không tìm thấy · Alt+B quay lại</div>`;
    }
    card.innerHTML = `<div class="hd"><div class="dots">${dots()}</div><b>HaiChau <span>Collector</span></b><span class="sp"></span><button class="ib" id="min" title="Thu nhỏ">–</button></div><div class="bd">${body}</div>`;
    $('#min').onclick = () => { S.min = true; save(); render(); };
    $('#load').onclick = () => $('#file').click();
    if (!total) return;
    $('#prev').onclick = () => go(S.idx - 1);
    $('#next').onclick = () => { const i = nextTodo(S.idx); i >= 0 ? go(i) : toast('Đã xử lý hết danh sách'); };
    $('#search').onclick = openSearch;
    $('#reread').onclick = () => { dirty = {}; refresh(true); };
    $('#revs').onclick = getReviews;
    $('#save').onclick = () => commit('Done');
    $('#nf').onclick = () => commit('Không tìm thấy');
    $('#closed').onclick = () => commit('Đóng cửa');
    $('#export').onclick = exportFile;
    $('#auto').onchange = (e) => { S.autoNext = e.target.checked; save(); };
    $('#maxr').ondblclick = () => { const n = parseInt(prompt('Số bình luận tối đa mỗi quán:', S.maxReviews), 10); if (n > 0) { S.maxReviews = n; save(); render(); } };
    root.querySelectorAll('input[data-k]').forEach((i) => i.addEventListener('input', () => { dirty[i.dataset.k] = i.value; i.classList.remove('auto'); }));
  }
  const dots = () => ['#4A90B5', '#F7CF8F', '#5CC3C1', '#4A90B5', '#F28C5B', '#5CC3C1', '#4A90B5', '#F7CF8F', '#5CC3C1'].map((c) => `<i style="background:${c}"></i>`).join('');
  async function getReviews() {
    if (!isPlacePage()) return toast('Hãy mở trang quán trên Google Maps trước', true);
    toast('Đang lấy bình luận…');
    const { list, note } = await collectReviews(S.maxReviews);
    pendingReviews = list; render();
    toast(list.length ? `Lấy được ${list.length} bình luận` : note || 'Không đọc được bình luận nào', !list.length);
  }
  function commit(status) {
    const s = cur(); if (!s) return;
    const get = (k) => num(root.querySelector(`input[data-k="${k}"]`)?.value);
    const rec = { status, rating: get('rating'), ratingCount: get('ratingCount'), priceMin: get('priceMin'), priceMax: get('priceMax'), lat: get('lat'), lng: get('lng'), url: isPlacePage() ? location.href.split('?')[0] : '' };
    if (status === 'Done') {
      if (!isPlacePage()) return toast('Chưa mở trang quán: chọn đúng quán trong kết quả tìm kiếm', true);
      if (rec.priceMin != null && rec.priceMax != null && rec.priceMax < rec.priceMin) return toast('Giá cao nhất nhỏ hơn giá thấp nhất', true);
      if (rec.rating != null && (rec.rating < 1 || rec.rating > 5)) return toast('Đánh giá phải từ 1 đến 5', true);
    } else { Object.assign(rec, { rating: null, ratingCount: null, priceMin: null, priceMax: null, lat: status === 'Đóng cửa' ? rec.lat : null, lng: status === 'Đóng cửa' ? rec.lng : null }); }
    S.results[s.storeId] = rec;
    if (pendingReviews) S.reviews[s.storeId] = pendingReviews;
    pendingReviews = null; save();
    toast(`Đã lưu ${s.storeName} (${status})`);
    const i = nextTodo(S.idx);
    if (i < 0) { render(); return toast('🎉 Đã xử lý hết danh sách, bấm Xuất Excel'); }
    go(i, S.autoNext);
  }
  function refresh(force) {
    const p = isPlacePage() ? readPlace() : null;
    const changed = JSON.stringify(p) !== JSON.stringify(place);
    if (changed || force) { place = p; if (changed) pendingReviews = null; render(); }
  }

  /* ---------- Theo dõi trang (Google Maps là SPA: đổi URL không tải lại trang) ---------- */
  let lastUrl = '';
  setInterval(() => {
    if (location.href !== lastUrl) { lastUrl = location.href; dirty = {}; setTimeout(() => refresh(true), 1200); }
    else if (isPlacePage() && (!place || !place.rating || !place.name)) refresh(false);   // nội dung tải chậm
  }, 1000);
  document.addEventListener('keydown', (e) => {
    if (!e.altKey || !S.stores.length) return;
    const k = e.code;
    const act = { KeyN: openSearch, KeyR: getReviews, KeyS: () => commit('Done'), KeyK: () => commit('Không tìm thấy'), KeyB: () => go(S.idx - 1) }[k];
    if (act) { e.preventDefault(); e.stopPropagation(); act(); }
  }, true);
  render();

  // Cho phép kiểm thử tự động các hàm đọc dữ liệu
  window.__hchCollector = { parsePriceRange, relDate, nameSim, readPlace, state: S };
})();
