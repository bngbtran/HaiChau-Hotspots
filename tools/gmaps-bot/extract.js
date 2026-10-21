// Hàm đọc dữ liệu chạy BÊN TRONG trang Google Maps (được bot.js tiêm vào bằng page.evaluate).
// Viết bằng JS thuần, không phụ thuộc thư viện; mỗi giá trị có vài cách đọc dự phòng vì class của Google hay đổi.
/* eslint-disable no-undef */
module.exports = function installExtractors() {
  if (window.__hch) return;
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const num = (v) => { const x = parseFloat(String(v ?? '').replace(/\s/g, '').replace(',', '.')); return Number.isFinite(x) ? x : null; };
  const norm = (s) => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/gi, 'd').toLowerCase();

  // "100–200 N ₫", "₫1–100 N", "200.000–300.000 ₫", "1-100.000 ₫", "Trên 1 Tr ₫" → [min, max] VNĐ
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
  // "2 tuần trước", "một năm trước", "Đã chỉnh sửa 3 tháng trước" → YYYY-MM-DD (ước lượng theo ngày chạy)
  function relDate(text) {
    const t = norm(text);
    if (/vua xong|vua moi/.test(t)) return new Date().toISOString().slice(0, 10);
    const m = t.match(/(\d+|mot|hai|ba)\s*(phut|gio|ngay|tuan|thang|nam)\s*truoc/);
    if (!m) return '';
    const n = { mot: 1, hai: 2, ba: 3 }[m[1]] || parseInt(m[1], 10);
    const d = new Date(); d.setDate(d.getDate() - { phut: 0, gio: 0, ngay: 1, tuan: 7, thang: 30, nam: 365 }[m[2]] * n);
    return d.toISOString().slice(0, 10);
  }
  function coordsFrom(url) {
    const u = decodeURIComponent(url || location.href);
    const all = u.match(/!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)/g);
    if (all) { const m = all[all.length - 1].match(/!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)/); return { lat: +m[1], lng: +m[2] }; }
    return { lat: null, lng: null };
  }
  const main = () => document.querySelector('h1')?.closest('div[role="main"]') || document.querySelector('div[role="main"]');

  // Trạng thái trang: 'place' (đang mở một quán), 'list' (danh sách kết quả), 'captcha', 'empty', 'loading'
  function pageState() {
    if (/\/sorry\//.test(location.href) || /unusual traffic|lưu lượng truy cập bất thường|not a robot/i.test(document.body?.innerText || '')) return 'captcha';
    if (/\/maps\/place\//.test(location.pathname) && main()?.querySelector('h1')) return 'place';
    if (document.querySelector('div[role="feed"] a[href*="/maps/place/"]')) return 'list';
    if (/không tìm thấy|did not match|no results/i.test(document.querySelector('div[role="main"]')?.innerText || '')) return 'empty';
    return 'loading';
  }
  // Danh sách kết quả: tên + toạ độ + link từng quán
  function listResults() {
    return [...document.querySelectorAll('div[role="feed"] a[href*="/maps/place/"]')].map((a) => ({ name: a.getAttribute('aria-label') || '', href: a.href, ...coordsFrom(a.href) }));
  }
  function readPlace() {
    const m = main(); if (!m) return null;
    const out = { name: (m.querySelector('h1')?.innerText || '').trim(), rating: null, ratingCount: null, priceMin: null, priceMax: null, priceText: '',
      category: '', address: '', closed: '', url: location.href.split('?')[0], ...coordsFrom() };
    const head = m.querySelector('div.F7nice')?.innerText || '';
    const hm = head.match(/(\d[.,]\d)\s*\(?\s*([\d.,]+)?\s*\)?/);
    if (hm) { out.rating = num(hm[1]); if (hm[2]) out.ratingCount = parseInt(hm[2].replace(/[^\d]/g, ''), 10) || null; }
    const aria = [...m.querySelectorAll('[aria-label]')].map((e) => e.getAttribute('aria-label'));
    if (out.rating == null) out.rating = num((aria.find((a) => /^\s*\d[.,]\d\s*sao/i.test(a)) || '').match(/\d[.,]\d/)?.[0]);
    if (out.ratingCount == null) { const c = aria.find((a) => /^[\d.,]+\s*(bài )?đánh giá/i.test(a)); if (c) out.ratingCount = parseInt(c.replace(/[^\d]/g, ''), 10) || null; }
    // Khoảng giá: phần tử aria "Giá…", dòng "Giá mỗi người", hoặc dòng có ₫ ở phần đầu khung thông tin
    const priceAria = aria.filter((a) => /giá|price/i.test(a) && /\d/.test(a) && !/đánh giá/i.test(a)).join('\n');
    const perPerson = (m.innerText.match(/Giá mỗi người[^\n]*\n?[^\n]*/i) || [''])[0];
    for (const src of [priceAria, perPerson, m.innerText.slice(0, 600)]) {
      const [lo, hi] = parsePriceRange(src);
      if (lo != null) { out.priceMin = lo; out.priceMax = hi; out.priceText = src.split('\n').find((l) => /₫/.test(l))?.trim() || ''; break; }
    }
    out.category = (m.querySelector('button[jsaction*="category"]')?.innerText || '').trim();
    out.address = (m.querySelector('button[data-item-id="address"]')?.getAttribute('aria-label') || '').replace(/^Địa chỉ:\s*/i, '').trim();
    const t = m.innerText;
    out.closed = /đóng cửa vĩnh viễn|permanently closed/i.test(t) ? 'Đóng cửa vĩnh viễn' : /tạm thời đóng cửa|temporarily closed/i.test(t) ? 'Tạm đóng cửa' : '';
    return out;
  }
  // Bình luận: mở tab "Bài đánh giá", cuộn tới khi đủ, mở rộng "Thêm", đọc sao + nội dung + ngày
  async function collectReviews(max) {
    const tab = [...document.querySelectorAll('button[role="tab"]')].find((b) => /đánh giá|reviews/i.test((b.getAttribute('aria-label') || '') + b.innerText));
    if (!tab) return [];
    if (tab.getAttribute('aria-selected') !== 'true') { tab.click(); }
    const cards = () => [...document.querySelectorAll('div[data-review-id][aria-label]')];
    for (let k = 0; k < 20 && !cards().length; k++) await sleep(300);
    // Cuộn tới bình luận cuối để Google tải thêm (tải lười); dừng khi đủ hoặc không tải thêm được
    for (let k = 0, stall = 0; k < 15 && cards().length < max && stall < 3; k++) {
      const before = cards().length, last = cards().pop(); if (!last) break;
      last.scrollIntoView({ block: 'end' });
      let box = last.parentElement;
      while (box && box !== document.body && !(box.scrollHeight > box.clientHeight + 20 && /auto|scroll/.test(getComputedStyle(box).overflowY))) box = box.parentElement;
      if (box && box !== document.body) box.scrollBy(0, 2000);
      for (let w = 0; w < 12 && cards().length === before; w++) await sleep(250);
      stall = cards().length === before ? stall + 1 : 0;
    }
    cards().slice(0, max).forEach((c) => c.querySelectorAll('button').forEach((b) => { if (/^(thêm|xem thêm|more|see more)$/i.test((b.innerText || b.getAttribute('aria-label') || '').trim())) b.click(); }));
    await sleep(400);
    const out = [];
    for (const c of cards()) {
      const starsEl = c.querySelector('[role="img"][aria-label*="sao"], [role="img"][aria-label*="star"]');
      let stars = parseInt((starsEl?.getAttribute('aria-label') || '').match(/\d/)?.[0], 10);
      if (!stars) { const sm = (c.innerText.match(/(\d)\/5/) || [])[1]; stars = parseInt(sm, 10); }
      const comment = (c.querySelector('.wiI7pd')?.innerText || '').replace(/\s+/g, ' ').trim();
      const dateTxt = c.querySelector('.rsqaWe')?.innerText || [...c.querySelectorAll('span')].map((s) => s.innerText).find((x) => /trước|vừa xong/i.test(x) && x.length < 40) || '';
      if (stars >= 1 && stars <= 5 && comment) out.push({ rating: stars, comment, date: relDate(dateTxt), author: c.getAttribute('aria-label') || '' });
      if (out.length >= max) break;
    }
    return out;
  }
  window.__hch = { pageState, listResults, readPlace, collectReviews, parsePriceRange, relDate };
};
