#!/usr/bin/env node
/* HaiChau Hotspots · bot thu thập dữ liệu quán từ Google Maps

   Cách dùng cho thành viên (chi tiết trong HUONG_DAN.md):
     1. Bỏ file Excel phân công của bạn vào thư mục input/
        (cần các cột: Tên quán, Địa chỉ, Vĩ độ, Kinh độ, storeId — tên cột tiếng Việt hoặc storeName/address/latitude/longitude)
     2. Bấm đúp CHAY_BOT (Windows: CHAY_BOT.bat · macOS: CHAY_BOT.command), hoặc gõ:  node bot.js run
     3. Kết quả: output/<tên file>_ketqua.xlsx = bảng của bạn + cột Vĩ độ mới, Kinh độ mới, Đánh giá, Số đánh giá,
        Giá thấp nhất, Giá cao nhất, Tên trên Google, Khớp, Lệch (m), Link Google Maps mới; sheet "Bình luận".
        File được ghi lại sau mỗi 20 quán và khi xong, nên dừng giữa chừng vẫn có kết quả.

   Lệnh:
     node bot.js login                 đăng nhập Google một lần (để lấy được hơn 5 bình luận mỗi quán)
     node bot.js run [file.xlsx ...]   crawl các file chỉ định, hoặc mọi file .xlsx trong input/; chạy lại sẽ làm tiếp
     node bot.js status                xem tiến độ
     node bot.js export                ghi lại file kết quả từ dữ liệu đã crawl
   Tuỳ chọn: --reviews 20 (số bình luận mỗi quán, 0 = bỏ qua) · --delay 4-7 (giây nghỉ giữa các quán) · --limit 5 (chạy thử)
             --redo not_found,check (làm lại các quán đó)

   Bot mở cửa sổ Chrome thật (hồ sơ riêng trong .profile/, giữ phiên đăng nhập). Gặp captcha / "lưu lượng bất thường"
   thì bot dừng, kêu "bíp" và chờ bạn tự xử lý trong cửa sổ đó rồi chạy tiếp. */
const fs = require('fs');
const path = require('path');
const puppeteer = require('puppeteer-core');
const XLSX = require('xlsx');
const install = require('./extract');

const HERE = __dirname;
const IN_DIR = path.join(HERE, 'input');
const OUT_DIR = path.join(HERE, 'output');
const PROFILE = path.join(HERE, '.profile');
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

/* ---------- tham số ---------- */
const [cmd = 'run', ...rest] = process.argv.slice(2);
const opt = { reviews: 20, delay: '4-7', limit: Infinity, redo: '' }, files = [];
for (let i = 0; i < rest.length; i++) {
  const k = rest[i].replace(/^--/, '');
  if (rest[i].startsWith('--') && k in opt) { opt[k] = typeof opt[k] === 'number' ? Number(rest[i + 1]) : rest[i + 1]; i++; }
  else if (/\.xlsx?$/i.test(rest[i])) files.push(path.resolve(rest[i]));
  else { console.error(`Không hiểu: ${rest[i]}`); process.exit(1); }
}
const [dMin, dMax] = String(opt.delay).split('-').map(Number);

/* ---------- đọc file phân công ---------- */
const norm = (s) => String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/gi, 'd').toLowerCase().trim();
const COLS = {
  name: ['ten quan', 'storename', 'ten', 'name'],
  address: ['dia chi', 'address'],
  lat: ['vi do', 'latitude', 'lat'],
  lng: ['kinh do', 'longitude', 'lng', 'lon'],
  id: ['storeid', 'store id', 'ma quan'],
  status: ['trang thai', 'status'],
  placeId: ['place_id', 'placeid'],
};
const toNum = (v) => { if (typeof v === 'number') return v; const x = parseFloat(String(v ?? '').trim().replace(/\s/g, '').replace(',', '.')); return Number.isFinite(x) ? x : null; };
function inputFiles() {
  if (files.length) return files;
  fs.mkdirSync(IN_DIR, { recursive: true });
  const list = fs.readdirSync(IN_DIR).filter((f) => /\.xlsx?$/i.test(f) && !f.startsWith('~$')).map((f) => path.join(IN_DIR, f));
  if (!list.length) { console.error(`Chưa có file Excel nào trong ${IN_DIR}\n→ Bỏ file phân công của bạn vào thư mục input/ rồi chạy lại.`); process.exit(1); }
  return list;
}
function readInput(file) {
  const wb = XLSX.readFile(file);
  const sheetName = wb.SheetNames[0], ws = wb.Sheets[sheetName];
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true });
  // dòng tiêu đề = dòng đầu tiên có cột tên quán và cột vĩ độ
  const hRow = grid.findIndex((r) => r.some((c) => COLS.name.includes(norm(c))) && r.some((c) => COLS.lat.includes(norm(c))));
  if (hRow < 0) throw new Error(`${path.basename(file)}: không tìm thấy dòng tiêu đề có cột "Tên quán" và "Vĩ độ"`);
  const header = grid[hRow].map((h, i) => String(h || '').trim() || `Cột ${i + 1}`);
  const col = Object.fromEntries(Object.entries(COLS).map(([k, names]) => [k, header.findIndex((h) => names.includes(norm(h)))]));
  if (col.lat < 0 || col.lng < 0) throw new Error(`${path.basename(file)}: thiếu cột Vĩ độ / Kinh độ`);
  const rows = [];
  for (let r = hRow + 1; r < grid.length; r++) {
    const g = grid[r];
    if (!g || !String(g[col.name] ?? '').trim()) continue;
    // giữ nguyên các ô; ô có link (vd. "Mở bản đồ") thì lấy địa chỉ link
    const orig = header.map((_, c) => { const cell = ws[XLSX.utils.encode_cell({ r, c })]; return cell?.l?.Target || g[c]; });
    rows.push({ orig, key: String(col.id >= 0 && g[col.id] ? g[col.id] : `row${r + 1}`).trim(),
      name: String(g[col.name]).trim(), address: col.address >= 0 ? String(g[col.address] || '') : '',
      lat: toNum(g[col.lat]), lng: toNum(g[col.lng]), status: col.status >= 0 ? String(g[col.status] || '').trim() : '',
      placeId: col.placeId >= 0 ? String(g[col.placeId] || '').trim() : '' });
  }
  const base = path.basename(file).replace(/\.xlsx?$/i, '');
  return { file, base, sheetName, header, col, rows, results: path.join(OUT_DIR, `${base}.jsonl`), out: path.join(OUT_DIR, `${base}_ketqua.xlsx`) };
}
function loadJsonl(file) {
  const m = new Map();
  if (fs.existsSync(file)) for (const line of fs.readFileSync(file, 'utf8').split('\n')) if (line.trim()) { const r = JSON.parse(line); m.set(r.key, r); }
  return m;
}

/* ---------- so khớp tên + khoảng cách ---------- */
const STOP = new Set(['quan', 'nha', 'hang', 'tiem', 'cua', 'shop', 'cafe', 'coffee', 'ca', 'phe', 'va', 'da', 'nang', 'chi', 'nhanh', 'cs', 'co', 'so']);
const toks = (s) => new Set(norm(s).replace(/[^a-z0-9 ]/g, ' ').split(/\s+/).filter((t) => t && !STOP.has(t)));
function nameSim(a, b) { const A = toks(a), B = toks(b); if (!A.size || !B.size) return 0; let n = 0; A.forEach((t) => { if (B.has(t)) n++; }); return n / Math.min(A.size, B.size); }
function distM(la1, ln1, la2, ln2) {
  if ([la1, ln1, la2, ln2].some((v) => v == null || Number.isNaN(+v))) return null;
  const r = Math.PI / 180, h = Math.sin((la2 - la1) * r / 2) ** 2 + Math.cos(la1 * r) * Math.cos(la2 * r) * Math.sin((ln2 - ln1) * r / 2) ** 2;
  return Math.round(2 * 6371000 * Math.asin(Math.sqrt(h)));
}
// Khớp khi tên giống và ở gần (≤ 400 m). Tên rất giống và đặc trưng (≥ 2 từ riêng, vd. "KUnivercity Korean BBQ")
// thì chấp nhận cả khi cách xa tới 15 km: toạ độ cũ sai hoặc quán đã chuyển chỗ, đúng là trường hợp cần toạ độ mới.
// Tên chung chung ("Milano Coffee", "Xôi Chiên") thì bắt buộc phải gần để không nhầm chi nhánh / quán trùng tên.
function judge(name, sim, d) {
  if (sim >= 0.5 && (d == null || d <= 400)) return 'ok';
  if (sim >= 0.8 && toks(name).size >= 2 && d != null && d <= 15000) return 'ok';
  return 'check';
}

/* ---------- trình duyệt ---------- */
function chromePath() {
  if (process.env.CHROME_PATH) return process.env.CHROME_PATH;
  const c = { darwin: ['/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'],
    win32: [`${process.env['PROGRAMFILES']}\\Google\\Chrome\\Application\\chrome.exe`, `${process.env['PROGRAMFILES(X86)']}\\Google\\Chrome\\Application\\chrome.exe`, `${process.env.LOCALAPPDATA}\\Google\\Chrome\\Application\\chrome.exe`],
    linux: ['/usr/bin/google-chrome', '/usr/bin/chromium', '/usr/bin/chromium-browser'] }[process.platform] || [];
  const p = c.find((x) => x && fs.existsSync(x));
  if (!p) { console.error('Không tìm thấy Google Chrome. Hãy cài Chrome (https://www.google.com/chrome).'); process.exit(1); }
  return p;
}
const launch = (w = 1280, h = 900) => puppeteer.launch({ executablePath: chromePath(), headless: false, userDataDir: PROFILE, defaultViewport: null, args: ['--lang=vi-VN', `--window-size=${w},${h}`] });
async function state(page) { try { await page.evaluate(install); return await page.evaluate(() => window.__hch.pageState()); } catch (e) { return 'loading'; } }
async function waitState(page, ms = 15000) {
  const t0 = Date.now(); let s = 'loading';
  while (Date.now() - t0 < ms) {
    s = await state(page);
    if (s === 'captcha') { await waitHuman(page); continue; }
    if (s !== 'loading') { await sleep(1200); return await state(page); }
    await sleep(400);
  }
  return s;
}
async function waitHuman(page) {
  process.stdout.write('\x07\n⚠️  Google yêu cầu xác minh (captcha). Hãy xử lý trong cửa sổ Chrome, bot sẽ tự chạy tiếp...\n');
  for (let i = 0; i < 360; i++) { await sleep(5000); if ((await state(page)) !== 'captcha') { console.log('   Đã qua, chạy tiếp.'); await sleep(3000); return; } }
  throw new Error('Chờ xác minh quá 30 phút');
}

/* ---------- xử lý một quán ---------- */
// Mở một câu tìm kiếm; trả về trang quán khớp nhất (hoặc null) và các ứng viên đã thấy
async function tryQuery(page, s, url) {
  await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 45000 });
  let st = await waitState(page), seen = [];
  if (st === 'list') {
    seen = (await page.evaluate(() => window.__hch.listResults())).map((c) => ({ ...c, sim: nameSim(s.name, c.name), dist: distM(s.lat, s.lng, c.lat, c.lng) }));
    const ok = seen.filter((c) => judge(s.name, c.sim, c.dist) === 'ok').sort((a, b) => b.sim - a.sim || (a.dist ?? 9e9) - (b.dist ?? 9e9));
    if (!ok.length) return { place: null, seen };
    await page.goto(ok[0].href, { waitUntil: 'domcontentloaded', timeout: 45000 });
    st = await waitState(page);
  }
  if (st !== 'place') return { place: null, seen };
  const p = await page.evaluate(() => window.__hch.readPlace());
  p.sim = nameSim(s.name, p.name); p.dist = distM(s.lat, s.lng, p.lat, p.lng);
  return { place: p, seen: [...seen, { name: p.name, sim: p.sim, dist: p.dist }] };
}
async function crawlOne(page, s) {
  const rec = { key: s.key, name: s.name, crawled_at: new Date().toISOString().slice(0, 19) };
  // Thử lần lượt: place_id (nếu có) → tên + số nhà/đường → tên + "Đà Nẵng" → chỉ tên; dừng ở lần đầu tìm được quán khớp
  const street = s.address.split(',').find((x) => /\d/.test(x)) || '';
  const at = `/@${s.lat || 16.065},${s.lng || 108.222},17z?hl=vi`, search = (q) => `https://www.google.com/maps/search/${encodeURIComponent(q)}${at}`;
  const urls = [...new Set([
    s.placeId && `https://www.google.com/maps/place/?q=place_id:${s.placeId}&hl=vi`,
    street && search(`${s.name} ${street.trim()}`),
    search(`${s.name} Đà Nẵng`),
    search(s.name),
  ].filter(Boolean))];
  let p = null, seen = [];
  for (const url of urls) {
    const r = await tryQuery(page, s, url);
    seen.push(...r.seen);
    if (r.place && judge(s.name, r.place.sim, r.place.dist) === 'ok') { p = r.place; break; }
    if (r.place && !p) p = r.place;   // giữ ứng viên đầu tiên để ghi "Cần kiểm tra" nếu không lần nào khớp
  }
  if (!p) {
    const best = seen.sort((a, b) => b.sim - a.sim).slice(0, 3);
    return { ...rec, status: 'not_found', candidates: best.map((c) => `${c.name} (${c.dist == null ? '?' : c.dist >= 1000 ? (c.dist / 1000).toFixed(1) + ' km' : c.dist + ' m'})`) };
  }
  Object.assign(rec, { g_name: p.name, g_category: p.category, g_address: p.address, g_url: p.url, rating: p.rating, rating_count: p.ratingCount,
    price_min: p.priceMin, price_max: p.priceMax, lat_new: p.lat, lng_new: p.lng, name_sim: +p.sim.toFixed(2), dist_m: p.dist, closed: p.closed });
  rec.match = judge(s.name, p.sim, p.dist);
  if (rec.match !== 'ok') rec.candidates = [`${p.name} (${p.dist ?? '?'} m)`];
  rec.status = rec.match !== 'ok' ? 'check' : /vĩnh viễn/.test(p.closed) ? 'closed' : 'done';
  rec.reviews = rec.status === 'done' && opt.reviews > 0 ? await page.evaluate((n) => window.__hch.collectReviews(n), opt.reviews) : [];
  return rec;
}

/* ---------- ghi file kết quả ---------- */
const ST = { done: 'Done', not_found: 'Không tìm thấy', closed: 'Đóng cửa', check: 'Cần kiểm tra' };
const NEW_COLS = ['Vĩ độ mới', 'Kinh độ mới', 'Đánh giá', 'Số đánh giá', 'Giá thấp nhất', 'Giá cao nhất', 'Tên trên Google', 'Loại trên Google', 'Khớp', 'Lệch (m)', 'Link Google Maps mới', 'Nguồn'];
function writeOutput(inp) {
  const res = loadJsonl(inp.results);
  const header = [...inp.header.filter((h) => !NEW_COLS.includes(h)), ...(inp.col.status < 0 ? ['Trạng thái'] : []), ...NEW_COLS];
  const aoa = [header], revs = [['storeId', 'Tên quán', 'Số sao', 'Bình luận', 'Ngày', 'Người viết', 'Nguồn']];
  for (const s of inp.rows) {
    const r = res.get(s.key), ok = r?.match === 'ok';
    const row = inp.header.map((h, i) => (i === inp.col.status && r ? ST[r.status] : s.orig[i])).filter((_, i) => !NEW_COLS.includes(inp.header[i]));
    if (inp.col.status < 0) row.push(r ? ST[r.status] : 'Not Yet');
    row.push(ok ? r.lat_new ?? '' : '', ok ? r.lng_new ?? '' : '', ok ? r.rating ?? '' : '', ok ? r.rating_count ?? '' : '', ok ? r.price_min ?? '' : '', ok ? r.price_max ?? '' : '',
      r?.g_name || (r?.candidates ? `Gần giống: ${r.candidates.join('; ')}` : ''), r?.g_category || '', r?.match || '', r?.dist_m ?? '', r?.g_url || '', r ? 'googlemaps' : '');
    aoa.push(row);
    if (ok) for (const v of r.reviews || []) revs.push([s.key, s.name, v.rating, v.comment, v.date, v.author, 'googlemaps']);
  }
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = header.map((h) => ({ wch: /Tên|Địa chỉ|Link|Google/.test(h) ? 32 : /Phường|Loại/.test(h) ? 24 : 13 }));
  XLSX.utils.book_append_sheet(wb, ws, inp.sheetName.slice(0, 31) || 'Quán');
  const wr = XLSX.utils.aoa_to_sheet(revs); wr['!cols'] = [{ wch: 15 }, { wch: 28 }, { wch: 7 }, { wch: 90 }, { wch: 11 }, { wch: 22 }, { wch: 11 }];
  XLSX.utils.book_append_sheet(wb, wr, 'Bình luận');
  fs.mkdirSync(OUT_DIR, { recursive: true });
  try { XLSX.writeFile(wb, inp.out); }
  catch (e) { const alt = inp.out.replace(/\.xlsx$/, `_${Date.now()}.xlsx`); XLSX.writeFile(wb, alt); console.log(`(File ${path.basename(inp.out)} đang mở trong Excel nên đã ghi sang ${path.basename(alt)})`); }
  return { n: inp.rows.length, done: [...res.values()].filter((r) => r.status === 'done').length, reviews: revs.length - 1 };
}

/* ---------- lệnh ---------- */
async function run() {
  const inputs = inputFiles().map(readInput);
  const redo = new Set(String(opt.redo).split(',').filter(Boolean));
  const jobs = [];
  for (const inp of inputs) {
    const done = loadJsonl(inp.results);
    const skipStatus = (s) => s.status && !/^(not yet|chua|)$/i.test(norm(s.status));   // dòng đã điền tay (Done, ...) thì bỏ qua
    const todo = inp.rows.filter((s) => (!done.has(s.key) || redo.has(done.get(s.key).status)) && !skipStatus(s));
    console.log(`📄 ${path.basename(inp.file)}: ${inp.rows.length} quán · đã làm ${inp.rows.filter((s) => done.has(s.key)).length} · cần làm ${todo.length}`);
    todo.forEach((s) => jobs.push([inp, s]));
  }
  const list = jobs.slice(0, opt.limit);
  if (!list.length) { inputs.forEach(writeOutput); console.log('Không còn quán nào cần làm. Kết quả trong thư mục output/'); return; }
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const browser = await launch();
  const page = (await browser.pages())[0] || await browser.newPage();
  await page.setExtraHTTPHeaders({ 'Accept-Language': 'vi-VN,vi;q=0.9' });
  const t0 = Date.now(), tally = {};
  let stop = false;
  process.on('SIGINT', () => { if (stop) process.exit(1); stop = true; console.log('\nĐang dừng sau quán hiện tại... (Ctrl+C lần nữa để thoát ngay)'); });
  browser.on('disconnected', () => { if (!stop) { stop = true; console.log('\nCửa sổ Chrome đã bị đóng → dừng. Chạy lại để làm tiếp.'); } });
  for (let i = 0; i < list.length && !stop; i++) {
    const [inp, s] = list[i];
    let rec;
    try { rec = await crawlOne(page, s); }
    catch (e) { if (stop) break; rec = { key: s.key, status: 'error', error: String(e.message).slice(0, 120) }; }
    if (rec.status !== 'error') fs.appendFileSync(inp.results, JSON.stringify(rec) + '\n');   // lỗi tạm thời: không ghi, lần sau làm lại
    tally[rec.status] = (tally[rec.status] || 0) + 1;
    const eta = (Date.now() - t0) / (i + 1) * (list.length - i - 1) / 60000;
    console.log(`[${i + 1}/${list.length}] ${s.name.slice(0, 30).padEnd(30)} → ${rec.status.padEnd(9)}${rec.error ? ' ' + rec.error : ''}${rec.rating ? ` ★${rec.rating} (${rec.rating_count ?? '?'})` : ''}${rec.price_min ? ` · ${rec.price_min}–${rec.price_max ?? ''}₫` : ''}${rec.reviews?.length ? ` · ${rec.reviews.length} bình luận` : ''}${rec.dist_m != null ? ` · lệch ${rec.dist_m} m` : ''}${rec.candidates ? ` · gần giống: ${rec.candidates[0]}` : ''} · còn ~${eta.toFixed(0)} phút`);
    if ((i + 1) % 20 === 0) writeOutput(inp);
    if ((i + 1) % 150 === 0) { console.log('Nghỉ 60 giây cho đỡ bị chặn...'); await sleep(60000); }
    if (i < list.length - 1 && !stop) await sleep((dMin + Math.random() * (dMax - dMin)) * 1000);
  }
  const wasStopped = stop; stop = true;   // đóng Chrome có chủ đích: không báo "cửa sổ bị đóng"
  try { await browser.close(); } catch (e) { /* đã đóng */ }
  if (wasStopped) console.log('Đã dừng. Chạy lại để làm tiếp.');
  console.log('\nKết quả lần chạy:', tally, `· ${((Date.now() - t0) / 60000).toFixed(1)} phút`);
  for (const inp of inputs) { const w = writeOutput(inp); console.log(`✅ ${path.relative(HERE, inp.out)}: ${w.done}/${w.n} quán có dữ liệu · ${w.reviews} bình luận`); }
  console.log('Gửi file *_ketqua.xlsx trong thư mục output/ cho nhóm trưởng.');
}
async function login() {
  const browser = await launch(1100, 800);
  const page = (await browser.pages())[0] || await browser.newPage();
  await page.goto('https://www.google.com/maps?hl=vi', { waitUntil: 'domcontentloaded' });
  console.log('Bấm "Đăng nhập" ở góc phải cửa sổ Chrome vừa mở và đăng nhập Google như bình thường.\nXong thì ĐÓNG cửa sổ Chrome đó; các lần chạy bot sau sẽ dùng phiên đăng nhập này.');
  await new Promise((r) => browser.on('disconnected', r));
  console.log('Đã lưu phiên đăng nhập.');
}
function status() {
  for (const inp of inputFiles().map(readInput)) {
    const res = loadJsonl(inp.results), c = {};
    inp.rows.forEach((s) => { const r = res.get(s.key); const k = r ? r.status : 'chưa làm'; c[k] = (c[k] || 0) + 1; });
    const v = [...res.values()];
    console.log(`📄 ${path.basename(inp.file)}: ${inp.rows.length} quán`, c, `· ${v.reduce((a, r) => a + (r.reviews?.length || 0), 0)} bình luận`);
  }
}
function exportAll() { for (const inp of inputFiles().map(readInput)) { const w = writeOutput(inp); console.log(`✅ ${path.relative(HERE, inp.out)}: ${w.done}/${w.n} quán có dữ liệu · ${w.reviews} bình luận`); } }

Promise.resolve(({ run, login, status, export: exportAll }[cmd] || (() => console.log('Lệnh: run | login | status | export')))()).catch((e) => { console.error('Lỗi:', e.message); process.exit(1); });
