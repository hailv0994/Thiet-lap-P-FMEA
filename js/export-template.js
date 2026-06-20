/*
 * export-template.js — Xuất P-FMEA: đổ dữ liệu vào template gốc, giữ định dạng
 * tiêu đề + định dạng vùng dữ liệu cho dễ đọc, in A4.
 *
 * Tính năng:
 * - Font dữ liệu tiếng Việt: Arial 10pt; đặc tính đặc thù (E): Arial 20pt canh giữa.
 * - Cột số canh giữa; tự cân đối độ rộng cột theo nội dung (giữ tổng -> vừa A4).
 * - Chiều cao dòng tự tính theo nội dung -> không mất chữ.
 * - Ô "phát hiện ra": chỉ xuất nội dung (bỏ nhãn "-Phát hiện ra ...").
 * - Khi nội dung tràn sang trang A4 khác: tách ô gộp tại ranh giới trang và
 *   LẶP LẠI nội dung Quy trình/Dạng hỏng/Ảnh hưởng ở đầu trang sau (ngắt trang
 *   thủ công để Excel ngắt đúng chỗ).
 * - Lặp tiêu đề (dòng 1-9) mỗi trang + tự đánh số trang "1/N" góc trên phải.
 *
 * Hàm chính: buildFromTemplate(state, templateBytes, fflate) -> Uint8Array
 */
(function (root) {
  'use strict';

  const COLS = 19;
  const colLetter = (c) => {
    let s = '';
    while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = (c - m - 1) / 26; }
    return s;
  };
  const xmlEsc = (s) => String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/\r\n/g, '\n');

  const rpnOf = (req, cause) => {
    const s = +req.severity, o = +cause.occurrence, d = +cause.detection;
    return (s && o && d) ? s * o * d : '';
  };
  const causeText = (c) => ((c.category ? c.category + ': ' : '') + (c.cause || '')).trim();

  // cột -> loại: T text, N số, S đặc tính đặc thù
  const COL_KIND = { 1: 'T', 2: 'T', 3: 'T', 4: 'N', 5: 'S', 6: 'T', 7: 'T', 8: 'N', 9: 'T', 10: 'T', 11: 'N', 12: 'N', 13: 'T', 14: 'T', 15: 'T', 16: 'N', 17: 'N', 18: 'N', 19: 'N' };
  const TEXT_COLS = [1, 2, 3, 6, 7, 9, 10, 13, 14, 15];
  // Cột số canh giữa: 4(S) và 16(S sau) rộng hơn để tiêu đề "nghiêm trọng" không vỡ chữ
  const FIXED_W = { 4: 7, 5: 6, 8: 5, 11: 5, 12: 6, 16: 7, 17: 5, 18: 5, 19: 5 };
  // 13 (Biện pháp đề xuất) gấp ~2 lần; 3/6/9 (Ảnh hưởng/Nguyên nhân/Dự phòng) giảm ~30%
  const MIN_W = { 1: 16, 2: 12, 3: 11, 6: 10, 7: 7, 9: 8, 10: 16, 13: 20, 14: 10, 15: 10 };
  const MAX_W = { 1: 32, 2: 24, 3: 21, 6: 21, 7: 14, 9: 17, 10: 32, 13: 44, 14: 22, 15: 22 };
  const ORIG_TOTAL = 222.9; // tổng độ rộng cột gốc (để giữ vừa khổ A4 scale 63%)

  // ---- Tính dữ liệu cần ghi từ state ----
  function computeData(state, startRow) {
    const rows = {};
    const merges = []; // [c1,r1,c2,r2]
    const put = (r, c, v, num) => {
      if (v === '' || v == null) return;
      (rows[r] || (rows[r] = {}))[c] = { v, num: !!num };
    };

    let row = startRow;
    for (const p of state.processes) {
      const procStart = row;
      const totalRows = p.reqs.reduce((n, r) => n + (r.causes.length || 1), 0) || 1;
      const reqList = p.reqs.map((r, i) => `${i + 1}.${r.reqText}`).join('\n');
      const aText = `${p.no ? p.no + '.' : ''}${p.name}\n\n-Chức năng: \n${p.func || ''}\n\n-Yêu cầu: \n${reqList}`;
      put(procStart, 1, aText);
      if (totalRows > 1) merges.push([1, procStart, 1, procStart + totalRows - 1]);

      for (const r of p.reqs) {
        const reqStart = row;
        const rs = r.causes.length || 1;
        put(reqStart, 2, r.failureMode);
        const effect = r.effectAnalysis
          ? (r.effectStdText ? r.effectAnalysis + '\n=>' + r.effectStdText : r.effectAnalysis)
          : (r.effectStdText || '');
        put(reqStart, 3, effect);
        put(reqStart, 4, r.severity, true);
        put(reqStart, 5, r.classification);
        if (rs > 1) [2, 3, 4, 5].forEach((c) => merges.push([c, reqStart, c, reqStart + rs - 1]));

        (r.causes.length ? r.causes : [{}]).forEach((c) => {
          put(row, 6, causeText(c));
          put(row, 7, c.pastTrouble);
          put(row, 8, c.occurrence, true);
          put(row, 9, c.prevention);
          // Ô phát hiện ra: chỉ nội dung, KHÔNG nhãn
          const det = [c.detectCause, r.detectFailureAuto].map((s) => (s || '').trim()).filter(Boolean).join('\n');
          put(row, 10, det);
          put(row, 11, c.detection, true);
          const rpn = rpnOf(r, c); if (rpn) put(row, 12, rpn, true);
          put(row, 13, c.action);
          put(row, 14, c.responsible);
          put(row, 15, c.actionTaken);
          put(row, 16, c.s2, true); put(row, 17, c.o2, true); put(row, 18, c.d2, true);
          const rpn2 = (+c.s2 && +c.o2 && +c.d2) ? (+c.s2) * (+c.o2) * (+c.d2) : '';
          if (rpn2) put(row, 19, rpn2, true);
          row++;
        });
      }
    }
    return { rows, merges, endRow: row - 1 };
  }

  // ---- Độ rộng cột tự cân đối theo nội dung (giữ tổng = ORIG_TOTAL) ----
  function longestLine(text) {
    let m = 0; String(text).split('\n').forEach((s) => { if (s.length > m) m = s.length; });
    return m;
  }
  function computeWidths(rows) {
    const W = {};
    for (const c in FIXED_W) W[+c] = FIXED_W[c];
    // desired theo nội dung
    const desired = {};
    let sumDesired = 0;
    TEXT_COLS.forEach((c) => {
      let mx = 0;
      for (const r in rows) { const cell = rows[r][c]; if (cell && !cell.num) mx = Math.max(mx, longestLine(cell.v)); }
      const d = Math.min(MAX_W[c], Math.max(MIN_W[c], mx + 2));
      desired[c] = d; sumDesired += d;
    });
    const budget = ORIG_TOTAL - Object.values(FIXED_W).reduce((a, b) => a + b, 0);
    const scale = sumDesired > 0 ? budget / sumDesired : 1;
    TEXT_COLS.forEach((c) => { W[c] = Math.max(8, Math.round(desired[c] * scale * 100) / 100); });
    return W;
  }

  const LH = 14.5, PAD = 5; // 1 dòng Arial 10pt + đệm
  function cellLines(text, col, widths) {
    const cpl = Math.max(4, Math.floor(widths[col] || 8));
    let lines = 0;
    String(text).split('\n').forEach((seg) => { lines += Math.max(1, Math.ceil(seg.length / cpl)); });
    return Math.max(1, lines);
  }
  const clampH = (h) => Math.min(409, Math.max(15, Math.round(h * 100) / 100));

  function baseHeights(rows, merges, startRow, lastRow, widths) {
    const mergeTop = {};
    for (const m of merges) if (m[1] !== m[3]) mergeTop[m[1] + ',' + m[0]] = true;
    const H = {};
    for (let r = startRow; r <= lastRow; r++) {
      const vals = rows[r] || {};
      let maxLines = 1, hasSC = false;
      for (const c in vals) {
        if (mergeTop[r + ',' + c]) continue;
        const cell = vals[c];
        if (cell.num) continue;
        if (+c === 5) { hasSC = true; continue; }
        maxLines = Math.max(maxLines, cellLines(cell.v, +c, widths));
      }
      let h = maxLines * LH + PAD;
      if (hasSC) h = Math.max(h, 30);
      H[r] = clampH(h);
    }
    return H;
  }
  function applyMergeDeficit(H, merges, rows, widths) {
    for (const [c, r1, , r2] of merges) {
      const cell = (rows[r1] || {})[c]; if (!cell) continue;
      let need;
      if (+c === 5) need = 30;
      else if (cell.num) need = LH + PAD;
      else need = cellLines(cell.v, +c, widths) * LH + PAD;
      let avail = 0; for (let r = r1; r <= r2; r++) avail += H[r] || 15;
      if (need > avail) H[r1] = clampH((H[r1] || 15) + (need - avail));
    }
  }

  // ---- Ngắt trang + tách ô gộp tại ranh giới trang, lặp nội dung ----
  const PAGE_CAP = 670; // chiều cao nội dung (pt) cho vùng dữ liệu mỗi trang A4
  function paginate(H, startRow, lastRow) {
    const segments = []; const breaks = [];
    let pageStart = startRow, cum = 0;
    for (let r = startRow; r <= lastRow; r++) {
      const h = H[r] || 15;
      if (r > pageStart && cum + h > PAGE_CAP) {
        segments.push([pageStart, r - 1]); breaks.push(r - 1); pageStart = r; cum = 0;
      }
      cum += h;
    }
    segments.push([pageStart, lastRow]);
    return { segments, breaks };
  }
  // Tách ô gộp dọc theo các trang; lặp giá trị neo ở đầu mỗi trang.
  function splitMergesByPage(merges, rows, segments) {
    const out = [];
    for (const m of merges) {
      const [c, r1, c2, r2] = m;
      if (r1 === r2) { out.push(m); continue; } // không phải gộp dọc
      for (const [s, e] of segments) {
        const top = Math.max(r1, s), bot = Math.min(r2, e);
        if (top > bot) continue;
        if (top > r1) { // lặp lại nội dung neo ở đầu trang
          const src = (rows[r1] || {})[c];
          if (src) (rows[top] || (rows[top] = {}))[c] = src;
        }
        if (bot > top) out.push([c, top, c2, bot]);
      }
    }
    return out;
  }

  // ---- styles.xml: thêm font Arial + 3 xf ----
  function injectStyles(stylesXml) {
    const fm = /<fonts count="(\d+)"([^>]*)>/.exec(stylesXml);
    const fontCount = +fm[1];
    const idFont10 = fontCount, idFont20 = fontCount + 1;
    const f10 = '<font><sz val="10"/><name val="Arial"/><family val="2"/></font>';
    const f20 = '<font><sz val="20"/><name val="Arial"/><family val="2"/></font>';
    stylesXml = stylesXml.replace(/<\/fonts>/, f10 + f20 + '</fonts>')
      .replace(fm[0], `<fonts count="${fontCount + 2}"${fm[2]}>`);

    const xm = /<cellXfs count="(\d+)">/.exec(stylesXml);
    const xfCount = +xm[1];
    const idText = xfCount, idNum = xfCount + 1, idSC = xfCount + 2;
    const xfText = `<xf numFmtId="0" fontId="${idFont10}" fillId="0" borderId="8" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="top" wrapText="1"/></xf>`;
    const xfNum = `<xf numFmtId="0" fontId="${idFont10}" fillId="0" borderId="8" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>`;
    const xfSC = `<xf numFmtId="0" fontId="${idFont20}" fillId="0" borderId="8" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>`;
    stylesXml = stylesXml.replace(/<\/cellXfs>/, xfText + xfNum + xfSC + '</cellXfs>')
      .replace(xm[0], `<cellXfs count="${xfCount + 3}">`);
    return { xml: stylesXml, idText, idNum, idSC };
  }

  function genCell(ref, sId, cell) {
    const sAttr = (sId != null) ? ` s="${sId}"` : '';
    if (!cell || cell.v === '' || cell.v == null) return `<c r="${ref}"${sAttr}/>`;
    if (cell.num && !isNaN(+cell.v)) return `<c r="${ref}"${sAttr}><v>${+cell.v}</v></c>`;
    return `<c r="${ref}"${sAttr} t="inlineStr"><is><t xml:space="preserve">${xmlEsc(cell.v)}</t></is></c>`;
  }
  function genRow(rn, valsByCol, ids, h) {
    let cells = '';
    for (let c = 1; c <= COLS; c++) {
      const kind = COL_KIND[c];
      const sId = kind === 'N' ? ids.idNum : (kind === 'S' ? ids.idSC : ids.idText);
      cells += genCell(colLetter(c) + rn, sId, valsByCol[c]);
    }
    return `<row r="${rn}" spans="1:19" ht="${h}" customHeight="1">${cells}</row>`;
  }

  function buildColsXml(W) {
    let s = '<cols>';
    for (let c = 1; c <= COLS; c++) s += `<col min="${c}" max="${c}" width="${W[c]}" customWidth="1"/>`;
    return s + '</cols>';
  }

  // Xóa sheet "VÍ DỤ" (sheet2) + mọi tham chiếu -> file xuất chỉ còn sheet FORMAT
  function stripExampleSheet(files, dec, enc) {
    delete files['xl/worksheets/sheet2.xml'];
    delete files['xl/worksheets/_rels/sheet2.xml.rels'];
    delete files['xl/drawings/drawing2.xml'];
    delete files['xl/printerSettings/printerSettings2.bin'];
    delete files['xl/calcChain.xml']; // calcChain chỉ trỏ ô trong VÍ DỤ

    let wb = dec.decode(files['xl/workbook.xml']);
    wb = wb.replace(/<sheet [^>]*name="VÍ DỤ"[^>]*\/>/, '');
    // mọi definedName gắn sheet index 1 (VÍ DỤ) sẽ thành sai -> bỏ
    wb = wb.replace(/<definedName [^>]*localSheetId="1"[^>]*>[^<]*<\/definedName>/g, '');
    files['xl/workbook.xml'] = enc.encode(wb);

    let rels = dec.decode(files['xl/_rels/workbook.xml.rels']);
    rels = rels.replace(/<Relationship [^>]*Target="worksheets\/sheet2\.xml"[^>]*\/>/, '');
    rels = rels.replace(/<Relationship [^>]*Target="calcChain\.xml"[^>]*\/>/, '');
    files['xl/_rels/workbook.xml.rels'] = enc.encode(rels);

    let ct = dec.decode(files['[Content_Types].xml']);
    ct = ct.replace(/<Override PartName="\/xl\/worksheets\/sheet2\.xml"[^>]*\/>/, '');
    ct = ct.replace(/<Override PartName="\/xl\/drawings\/drawing2\.xml"[^>]*\/>/, '');
    ct = ct.replace(/<Override PartName="\/xl\/calcChain\.xml"[^>]*\/>/, '');
    files['[Content_Types].xml'] = enc.encode(ct);

    if (files['docProps/app.xml']) {
      let app = dec.decode(files['docProps/app.xml']);
      app = app.replace(/<TitlesOfParts>[\s\S]*?<\/TitlesOfParts>/,
        '<TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>FORMAT</vt:lpstr></vt:vector></TitlesOfParts>');
      app = app.replace(/<HeadingPairs>[\s\S]*?<\/HeadingPairs>/,
        '<HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs>');
      files['docProps/app.xml'] = enc.encode(app);
    }
  }

  function buildFromTemplate(state, templateBytes, fflate) {
    const files = fflate.unzipSync(templateBytes);
    const SHEET = 'xl/worksheets/sheet1.xml';
    const dec = new TextDecoder('utf-8'), enc = new TextEncoder();
    stripExampleSheet(files, dec, enc);
    let xml = dec.decode(files[SHEET]);
    // Bỏ nhãn tĩnh "Trang ページ：    /" ở ô R2 (số trang sẽ in qua header để đúng từng trang)
    xml = xml.replace(/<c r="R2"( s="\d+")?[^>]*>\s*<v>24<\/v>\s*<\/c>/, '<c r="R2"$1/>');

    const st = injectStyles(dec.decode(files['xl/styles.xml']));
    files['xl/styles.xml'] = enc.encode(st.xml);
    const ids = { idText: st.idText, idNum: st.idNum, idSC: st.idSC };

    const START = 10;
    const { rows, merges, endRow } = computeData(state, START);
    const lastRow = Math.max(endRow, START);
    const widths = computeWidths(rows);

    // chiều cao -> ngắt trang -> tách ô gộp & lặp nội dung -> chiều cao cuối
    let H = baseHeights(rows, merges, START, lastRow, widths);
    applyMergeDeficit(H, merges, rows, widths);
    const { segments, breaks } = paginate(H, START, lastRow);
    const finalMerges = splitMergesByPage(merges, rows, segments);
    H = baseHeights(rows, finalMerges, START, lastRow, widths);
    applyMergeDeficit(H, finalMerges, rows, widths);

    // ghi đè các dòng dữ liệu
    for (let rn = START; rn <= lastRow; rn++) {
      const newRow = genRow(rn, rows[rn] || {}, ids, H[rn] || 15);
      const re = new RegExp(`<row r="${rn}"[^>]*(?:/>|>[\\s\\S]*?</row>)`);
      if (re.test(xml)) xml = xml.replace(re, newRow);
      else xml = xml.replace('</sheetData>', newRow + '</sheetData>');
    }

    // merge dữ liệu (đã tách theo trang)
    if (finalMerges.length) {
      const mm = /<mergeCells count="(\d+)">([\s\S]*?)<\/mergeCells>/.exec(xml);
      if (mm) {
        let add = '';
        for (const [c1, r1, c2, r2] of finalMerges) add += `<mergeCell ref="${colLetter(c1)}${r1}:${colLetter(c2)}${r2}"/>`;
        xml = xml.replace(mm[0], `<mergeCells count="${+mm[1] + finalMerges.length}">${mm[2]}${add}</mergeCells>`);
      }
    }

    // độ rộng cột
    xml = xml.replace(/<cols>[\s\S]*?<\/cols>/, buildColsXml(widths));

    // lề + scale in A4 + đánh số trang (vào dòng "Trang ページ") + ngắt trang thủ công
    xml = xml.replace(/<pageMargins[^>]*\/>/, '<pageMargins left="0" right="0" top="0.5" bottom="0.2" header="0.3" footer="0"/>');
    // scale 63% -> 70% cho vừa khổ A4; bỏ r:id để scale trong XML có hiệu lực
    xml = xml.replace(/<pageSetup[^>]*\/>/, '<pageSetup paperSize="9" scale="70" orientation="landscape"/>');
    // Số trang in vào ĐÚNG dòng "Trang ページ：" (góc trên phải), đúng từng trang (&P/&N)
    const headerFooter = '<headerFooter><oddHeader>&amp;R&amp;&quot;Arial&quot;&amp;10Trang ページ： &amp;P／&amp;N</oddHeader></headerFooter>';
    // Template ĐÃ CÓ sẵn 1 thẻ <headerFooter/>; phải THAY THẾ (không chèn thêm),
    // nếu không sẽ trùng 2 thẻ -> Excel báo "We found a problem with content".
    if (/<headerFooter[^>]*\/>/.test(xml)) {
      xml = xml.replace(/<headerFooter[^>]*\/>/, headerFooter);
    } else if (/<headerFooter[\s\S]*?<\/headerFooter>/.test(xml)) {
      xml = xml.replace(/<headerFooter[\s\S]*?<\/headerFooter>/, headerFooter);
    } else {
      xml = xml.replace(/(<pageSetup[^>]*\/>)/, `$1${headerFooter}`);
    }
    // rowBreaks phải đứng NGAY SAU headerFooter (đúng thứ tự schema worksheet)
    if (breaks.length) {
      const rowBreaks = `<rowBreaks count="${breaks.length}" manualBreakCount="${breaks.length}">` +
        breaks.map((r) => `<brk id="${r}" max="16383" man="1"/>`).join('') + '</rowBreaks>';
      xml = xml.replace(/(<\/headerFooter>)/, `$1${rowBreaks}`);
    }

    files[SHEET] = enc.encode(xml);

    // Print_Titles (lặp tiêu đề) + mở rộng Print_Area
    let wbxml = dec.decode(files['xl/workbook.xml']);
    const areaEnd = Math.max(lastRow, 60);
    wbxml = wbxml.replace(
      /<definedName name="_xlnm.Print_Area" localSheetId="0">[^<]*<\/definedName>/,
      `<definedName name="_xlnm.Print_Area" localSheetId="0">FORMAT!$A$1:$S$${areaEnd}</definedName>` +
      `<definedName name="_xlnm.Print_Titles" localSheetId="0">FORMAT!$1:$9</definedName>`
    );
    files['xl/workbook.xml'] = enc.encode(wbxml);

    return fflate.zipSync(files, { level: 6 });
  }

  const api = { buildFromTemplate, computeData, computeWidths, paginate };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.TemplateExport = api;
})(typeof window !== 'undefined' ? window : this);
