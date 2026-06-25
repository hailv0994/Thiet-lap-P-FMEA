#!/usr/bin/env node
/**
 * check-pagebreak.js — Kiểm tra tự động vị trí ngắt trang trong file Excel xuất ra.
 * Chạy: node tools/check-pagebreak.js
 *
 * Script sẽ:
 * 1. Tạo state mẫu (5 công đoạn × 3 yêu cầu × 2 nguyên nhân)
 * 2. Gọi buildFromTemplate() để xuất .xlsx
 * 3. Giải nén và đọc XML để kiểm tra:
 *    - Chiều cao các hàng tiêu đề (rows 1–9)
 *    - Vị trí ngắt trang (rowBreaks)
 *    - Tổng chiều cao data per page vs. khả năng in A4 landscape
 */
'use strict';

const fs   = require('fs');
const path = require('path');

const ROOT   = path.join(__dirname, '..');
const VENDOR = path.join(ROOT, 'vendor');
const DATA   = path.join(ROOT, 'data');
const JS     = path.join(ROOT, 'js');

// ── 1. fflate (CommonJS export) ──────────────────────────────────────────────
const fflate = require(path.join(VENDOR, 'fflate.min.js'));

// ── 2. Template base64 (trích từ data/template.js bằng regex, tránh eval file lớn) ─
const templateSrc = fs.readFileSync(path.join(DATA, 'template.js'), 'utf8');
const b64Match    = templateSrc.match(/PFMEA_TEMPLATE_B64\s*=\s*"([A-Za-z0-9+/=]+)"/);
if (!b64Match) { console.error('Không tìm thấy PFMEA_TEMPLATE_B64 trong template.js'); process.exit(1); }
const templateBytes = new Uint8Array(Buffer.from(b64Match[1], 'base64'));

// ── 3. export-template.js (sets global.TemplateExport via `this` in Node) ────
global.window = undefined; // đảm bảo export-template dùng `this` (= global) làm root
const exportSrc = fs.readFileSync(path.join(JS, 'export-template.js'), 'utf8');
// eslint-disable-next-line no-eval
(function () { eval(exportSrc); }).call(global);
const { buildFromTemplate } = global.TemplateExport;

// ── 4. State mẫu (đủ dữ liệu để tạo nhiều trang) ───────────────────────────
function makeCause(pi, ri, ci) {
  return {
    id: `c${pi}_${ri}_${ci}`, category: 'Thiết bị/con người',
    cause: `Nguyên nhân ${ci + 1}: vật liệu không đúng tiêu chuẩn, không kiểm soát đúng thông số gia công gây ra lỗi`,
    pastTrouble: '',
    occurrence: 4,
    prevention: `Biện pháp phòng ngừa ${ci + 1}: kiểm soát theo tiêu chuẩn, huấn luyện vận hành`,
    detectCause: `Kiểm tra bằng dưỡng / đo bằng thước cặp sau mỗi lô sản xuất`,
    detectExtra: '',
    detection: 5,
    action: '', responsible: '', actionTaken: '',
    s2: '', o2: '', d2: ''
  };
}
function makeReq(pi, ri) {
  return {
    id: `r${pi}_${ri}`, splitId: null, mergeId: null,
    reqText: `Yêu cầu chất lượng ${ri + 1}: kích thước nằm trong dung sai ±0.05mm, bề mặt không trầy xước`,
    failureMode: `Dạng hỏng hóc ${ri + 1}: kích thước ngoài dung sai, bề mặt bị trầy xước`,
    effectAnalysis: `Ảnh hưởng ${ri + 1}: sản phẩm không lắp được, ảnh hưởng chức năng giảm chấn`,
    effectStdText: '',
    severity: 7, classification: '',
    detectFailureAuto: `Kiểm tra bằng dưỡng\nĐo bằng thước cặp\nKiểm tra 100% tại cuối chuyền`,
    causes: [makeCause(pi, ri, 0), makeCause(pi, ri, 1)]
  };
}
function makeProc(pi) {
  return {
    no: pi + 1,
    name: `Công đoạn ${pi + 1}: ${['Cắt phôi', 'Gia công CNC', 'Hàn MIG', 'Mài + Đánh bóng', 'Kiểm tra cuối'][pi] || 'Lắp ráp'}`,
    func: `Chức năng: tạo hình dạng và kích thước theo bản vẽ, đảm bảo yêu cầu cơ học`,
    reqs: [makeReq(pi, 0), makeReq(pi, 1), makeReq(pi, 2)]
  };
}
const state = {
  meta: { dept: 'PRO1', product: 'Giảm xóc trước', line: 'Line A', model: 'CHECK-001' },
  processes: Array.from({ length: 5 }, (_, i) => makeProc(i))
};

// ── 5. Xuất Excel ─────────────────────────────────────────────────────────────
console.log('Đang xuất Excel mẫu...');
const result = buildFromTemplate(state, templateBytes, fflate);
const outPath = path.join(ROOT, 'tools', '_check-pagebreak-output.xlsx');
fs.writeFileSync(outPath, result);
console.log(`Đã ghi: ${outPath} (${(result.length / 1024).toFixed(0)} KB)\n`);

// ── 6. Phân tích XML ──────────────────────────────────────────────────────────
const files = fflate.unzipSync(result);
const dec   = new TextDecoder('utf-8');
const xml   = dec.decode(files['xl/worksheets/sheet1.xml']);

// 6a. Chiều cao hàng tiêu đề (rows 1–9) từ template
console.log('=== CHIỀU CAO HÀNG TIÊU ĐỀ (rows 1–9, từ template) ===');
let headerTotal = 0;
for (let r = 1; r <= 9; r++) {
  const m = xml.match(new RegExp(`<row r="${r}"[^>]* ht="([\\d.]+)"`));
  const h = m ? +m[1] : 15;
  headerTotal += h;
  console.log(`  Row ${r}: ${h.toFixed(2)} pt`);
}
console.log(`  Tổng header (rows 1–9): ${headerTotal.toFixed(2)} pt\n`);

// 6b. Ngắt trang
const rbMatch = xml.match(/<rowBreaks[^>]*>([\s\S]*?)<\/rowBreaks>/);
const breaks  = rbMatch
  ? [...rbMatch[1].matchAll(/<brk id="(\d+)"/g)].map((m) => +m[1])
  : [];
console.log(`=== VỊ TRÍ NGẮT TRANG (rowBreaks) ===`);
if (breaks.length === 0) {
  console.log('  (không có ngắt trang — toàn bộ dữ liệu nằm trên 1 trang)\n');
} else {
  console.log(`  ${breaks.length} ngắt: rows ${breaks.join(', ')}\n`);
}

// 6c. Chiều cao data rows — dùng ' ht=' (có khoảng trắng) để tránh khớp customHeight
const rowMatches = [...xml.matchAll(/<row r="(\d+)"[^>]* ht="([\d.]+)"/g)];
const heights    = {};
rowMatches.forEach(([, r, h]) => { heights[+r] = +h; });
const dataRows   = Object.keys(heights).map(Number).filter((r) => r >= 10).sort((a, b) => a - b);
const endRow     = dataRows.length ? dataRows[dataRows.length - 1] : 10;

// Tổng chiều cao data
const totalDataH = dataRows.reduce((s, r) => s + (heights[r] || 15), 0);
console.log(`=== DATA ROWS (10 – ${endRow}): ${dataRows.length} hàng, tổng = ${totalDataH.toFixed(1)} pt ===\n`);

// 6d. Thống kê per page
const A4_H = 537.68; // A4 landscape printable height (pt) với margins 0.28/0.3/0.1/0.12 in
const A4_W = 841.68; // A4 landscape printable width (pt), lề trái=phải=0
const MDW_PT = 5.25; // pt per Excel char unit (Arial 10pt, 7px/96DPI × 72)
const HEADER_H = headerTotal;

// Ước lượng tổng chiều rộng cột và scale in
const colsXml = xml.match(/<cols>([\s\S]*?)<\/cols>/);
let totalColW = 0;
if (colsXml) [...colsXml[1].matchAll(/<col min="(\d+)" max="\d+" width="([\d.]+)"/g)]
              .forEach(([,, w]) => { totalColW += +w; });
const estScale = Math.min(0.95, A4_W / (totalColW * MDW_PT));
const computedPageCap = Math.floor(A4_H / estScale - HEADER_H);

// Dùng endRow thực (không tính placeholder rows của template)
const actualLastRow = breaks.length ? Math.max(...breaks) + 6 : endRow;
const pageStarts = [10, ...breaks.map((b) => b + 1)];
const pageEnds   = [...breaks, Math.min(endRow, actualLastRow)];

console.log(`=== PHÂN TÍCH TỪNG TRANG ===`);
for (let i = 0; i < pageStarts.length; i++) {
  const from = pageStarts[i], to = pageEnds[i];
  let dataH = 0;
  for (let r = from; r <= to; r++) dataH += heights[r] || 0;
  const totalH = dataH + HEADER_H;
  const paperH = totalH * estScale;
  const util   = (paperH / A4_H * 100).toFixed(0);
  let warn = '';
  if (paperH > A4_H * 1.02) warn = '  ⚠️  TRÀN TRANG! Nội dung vượt A4 printable area.';
  else if (dataH < computedPageCap * 0.5 && i < pageStarts.length - 1) warn = '  ⚠️  Quá ít data, trang trống nhiều.';
  console.log(`  Trang ${i + 1}: rows ${from}–${to}  data=${dataH.toFixed(0)}pt  trên giấy~${paperH.toFixed(0)}pt  (${util}% A4)${warn}`);
}

// 6e. Tổng kết
console.log(`\n=== THÔNG SỐ TÍNH TOÁN ===`);
console.log(`  A4 landscape printable height:   ${A4_H.toFixed(2)} pt`);
console.log(`  Header rows 1–9 (từ template):  ${HEADER_H.toFixed(2)} pt`);
console.log(`  Tổng chiều rộng cột:             ${totalColW.toFixed(1)} char → ${(totalColW*MDW_PT).toFixed(0)} pt`);
console.log(`  Ước lượng scale in:              ${(estScale*100).toFixed(1)}%`);
console.log(`  pageCap tính được:               ${computedPageCap} pt (data/trang, không tính header)`);
console.log(`  => Mỗi trang in ~${computedPageCap} pt data (trên giấy ~${(computedPageCap*estScale).toFixed(0)}pt) + header ${(HEADER_H*estScale).toFixed(0)}pt`);
console.log('');
