/*
 * export-template.js — Xuất P-FMEA bằng cách ĐỔ DỮ LIỆU vào file template gốc,
 * GIỮ NGUYÊN định dạng tiêu đề + thêm định dạng vùng dữ liệu cho dễ đọc/in A4.
 *
 * - Header (dòng 1–9) giữ nguyên 100% style của template.
 * - Vùng dữ liệu dùng bộ style riêng (thêm vào styles.xml): font 10pt đồng nhất,
 *   wrap text, viền hộp; cột số canh giữa; cột "đặc tính đặc thù" (E) font 20pt
 *   canh giữa (gấp đôi cỡ chữ điểm S).
 * - Chiều cao mỗi dòng tự tính theo nội dung -> không mất chữ.
 * - Thiết lập in: A4 ngang (sẵn có) + lặp tiêu đề (dòng 1–9) ở mỗi trang.
 *
 * Hàm chính: buildFromTemplate(state, templateBytes, fflate) -> Uint8Array
 */
(function (root) {
  'use strict';

  const COLS = 19;
  const WIDTHS = [24.3, 18.7, 18.7, 5.7, 5.7, 22.7, 5.7, 5.7, 18.7, 18.7, 5.7, 5.7, 18.7, 12.7, 12.7, 5.7, 5.7, 5.7, 5.7];

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

  // Gộp 4M + nội dung nguyên nhân
  const causeText = (c) => {
    const cat = c.category ? c.category + ': ' : '';
    return (cat + (c.cause || '')).trim();
  };

  // ---- Tính dữ liệu cần ghi từ state ----
  function computeData(state, startRow) {
    const rows = {};
    const merges = [];
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
          const det = `-Phát hiện ra nguyên nhân:\n${c.detectCause || ''}\n-Phát hiện ra dạng hỏng hóc:\n${r.detectFailureAuto || ''}`;
          put(row, 10, det);
          put(row, 11, c.detection, true);
          const rpn = rpnOf(r, c);
          if (rpn) put(row, 12, rpn, true);
          put(row, 13, c.action);
          put(row, 14, c.responsible);
          put(row, 15, c.actionTaken);
          put(row, 16, c.s2, true);
          put(row, 17, c.o2, true);
          put(row, 18, c.d2, true);
          const rpn2 = (+c.s2 && +c.o2 && +c.d2) ? (+c.s2) * (+c.o2) * (+c.d2) : '';
          if (rpn2) put(row, 19, rpn2, true);
          row++;
        });
      }
    }
    return { rows, merges, endRow: row - 1 };
  }

  // ---- Thêm style cho vùng dữ liệu vào styles.xml ----
  function injectStyles(stylesXml) {
    // font 18pt (gấp đôi 9pt = cỡ chữ điểm S) cho đặc tính đặc thù
    const fm = /<fonts count="(\d+)"([^>]*)>/.exec(stylesXml);
    const fontCount = +fm[1];
    const newFontId = fontCount;
    const bigFont = '<font><sz val="18"/><name val="ＭＳ Ｐゴシック"/><family val="3"/><charset val="128"/></font>';
    stylesXml = stylesXml.replace(/<\/fonts>/, bigFont + '</fonts>')
      .replace(fm[0], `<fonts count="${fontCount + 1}"${fm[2]}>`);

    // 3 cellXf: TEXT, NUM, SC (borderId=8 = viền hộp thin đầy đủ)
    const xm = /<cellXfs count="(\d+)">/.exec(stylesXml);
    const xfCount = +xm[1];
    const idText = xfCount, idNum = xfCount + 1, idSC = xfCount + 2;
    const xfText = '<xf numFmtId="0" fontId="10" fillId="0" borderId="8" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="top" wrapText="1"/></xf>';
    const xfNum = '<xf numFmtId="0" fontId="10" fillId="0" borderId="8" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>';
    const xfSC = `<xf numFmtId="0" fontId="${newFontId}" fillId="0" borderId="8" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>`;
    stylesXml = stylesXml.replace(/<\/cellXfs>/, xfText + xfNum + xfSC + '</cellXfs>')
      .replace(xm[0], `<cellXfs count="${xfCount + 3}">`);

    return { xml: stylesXml, idText, idNum, idSC };
  }

  // cột -> loại style: T=text, N=number, S=special-char
  const COL_KIND = { 1: 'T', 2: 'T', 3: 'T', 4: 'N', 5: 'S', 6: 'T', 7: 'T', 8: 'N', 9: 'T', 10: 'T', 11: 'N', 12: 'N', 13: 'T', 14: 'T', 15: 'T', 16: 'N', 17: 'N', 18: 'N', 19: 'N' };

  const LH = 13.6, PAD = 5; // chiều cao 1 dòng (9pt) + đệm

  // ---- Ước lượng số dòng văn bản trong 1 ô ----
  function cellLines(text, col) {
    const cpl = Math.max(4, Math.floor(WIDTHS[col - 1])); // ký tự / dòng (ước lượng an toàn)
    let lines = 0;
    String(text).split('\n').forEach((seg) => {
      lines += Math.max(1, Math.ceil(seg.length / cpl));
    });
    return Math.max(1, lines);
  }
  const clampH = (h) => Math.min(409, Math.max(15, Math.round(h * 100) / 100));

  /*
   * Tính chiều cao từng dòng, có PHÂN BỔ cho ô gộp dọc (A theo công đoạn,
   * B/C/D/E theo yêu cầu): trước hết tính theo nội dung từng dòng (cột không
   * gộp), sau đó nếu nội dung ô gộp cao hơn tổng các dòng nó chiếm thì bù phần
   * thiếu vào dòng đầu của ô gộp -> dòng cân đối, không mất chữ.
   */
  function buildRowHeights(rows, merges, startRow, lastRow) {
    const mergeTop = {}; // "r,c" -> true (ô gộp dọc bắt đầu ở đây)
    const vmerges = [];
    for (const m of merges) {
      if (m[1] !== m[3]) { mergeTop[m[1] + ',' + m[0]] = true; vmerges.push(m); }
    }

    const H = {};
    for (let r = startRow; r <= lastRow; r++) {
      const vals = rows[r] || {};
      let maxLines = 1, hasSC = false;
      for (const c in vals) {
        if (mergeTop[r + ',' + c]) continue;      // ô gộp dọc -> xử lý sau
        const cell = vals[c];
        if (cell.num) continue;
        if (+c === 5) { hasSC = true; continue; }  // đặc tính đặc thù (font lớn)
        maxLines = Math.max(maxLines, cellLines(cell.v, +c));
      }
      let h = maxLines * LH + PAD;
      if (hasSC) h = Math.max(h, 28);
      H[r] = clampH(h);
    }

    // Bù chiều cao cho ô gộp dọc nếu nội dung không đủ chỗ
    for (const [c, r1, , r2] of vmerges) {
      const cell = (rows[r1] || {})[c];
      if (!cell) continue;
      let need;
      if (+c === 5) need = 28;                      // đặc tính đặc thù: 1 ký tự font 18pt
      else if (cell.num) need = LH + PAD;
      else need = cellLines(cell.v, +c) * LH + PAD;
      let avail = 0;
      for (let r = r1; r <= r2; r++) avail += H[r] || 15;
      if (need > avail) H[r1] = clampH((H[r1] || 15) + (need - avail));
    }
    return H;
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

  function buildFromTemplate(state, templateBytes, fflate) {
    const files = fflate.unzipSync(templateBytes);
    const SHEET = 'xl/worksheets/sheet1.xml'; // FORMAT
    const dec = new TextDecoder('utf-8');
    const enc = new TextEncoder();
    let xml = dec.decode(files[SHEET]);

    // 1) Thêm style vùng dữ liệu
    const st = injectStyles(dec.decode(files['xl/styles.xml']));
    files['xl/styles.xml'] = enc.encode(st.xml);
    const ids = { idText: st.idText, idNum: st.idNum, idSC: st.idSC };

    // 2) Dữ liệu
    const START = 10;
    const { rows, merges, endRow } = computeData(state, START);
    const lastRow = Math.max(endRow, START);
    const heights = buildRowHeights(rows, merges, START, lastRow);

    // 3) Ghi đè từng dòng dữ liệu (giữ style header, thay vùng data)
    for (let rn = START; rn <= lastRow; rn++) {
      const newRow = genRow(rn, rows[rn] || {}, ids, heights[rn] || 15);
      const re = new RegExp(`<row r="${rn}"[^>]*(?:/>|>[\\s\\S]*?</row>)`);
      if (re.test(xml)) xml = xml.replace(re, newRow);
      else xml = xml.replace('</sheetData>', newRow + '</sheetData>');
    }

    // 4) Thêm vùng merge dữ liệu
    if (merges.length) {
      const mm = /<mergeCells count="(\d+)">([\s\S]*?)<\/mergeCells>/.exec(xml);
      if (mm) {
        let add = '';
        for (const [c1, r1, c2, r2] of merges) add += `<mergeCell ref="${colLetter(c1)}${r1}:${colLetter(c2)}${r2}"/>`;
        xml = xml.replace(mm[0], `<mergeCells count="${+mm[1] + merges.length}">${mm[2]}${add}</mergeCells>`);
      }
    }
    files[SHEET] = enc.encode(xml);

    // 5) Thiết lập in: cập nhật Print_Area + thêm Print_Titles (lặp tiêu đề)
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

  const api = { buildFromTemplate, computeData, injectStyles };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.TemplateExport = api;
})(typeof window !== 'undefined' ? window : this);
