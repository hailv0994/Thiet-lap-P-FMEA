/*
 * export-template.js — Xuất P-FMEA bằng cách ĐỔ DỮ LIỆU vào file template gốc,
 * GIỮ NGUYÊN 100% định dạng (viền, màu, font, merge, độ rộng cột).
 *
 * Nguyên tắc: không dựng lại workbook. Ta giải nén template (.xlsx) bằng fflate,
 * chỉ sửa giá trị các ô trong sheet "FORMAT" (giữ nguyên thuộc tính style `s`
 * của từng ô), thêm các vùng merge cho dữ liệu, rồi nén lại.
 *
 * Hàm chính: buildFromTemplate(state, templateBytes, fflate) -> Uint8Array
 */
(function (root) {
  'use strict';

  const COLS = 19; // A..S
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

  // ---- Tính dữ liệu cần ghi từ state ----
  // Trả về { rows: { [rowNo]: { [col1based]: {v, num} } }, merges: [[c1,r1,c2,r2]] }
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
      const aText = `${p.no ? p.no + '.' : ''}${p.name}\n\n-Chức năng: \n${p.func}\n\n-Yêu cầu: \n${reqList}`;
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
          put(row, 6, c.cause);
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

  // ---- Đọc style `s` và thuộc tính từng dòng của template ----
  function parseTemplate(xml) {
    const styleMap = {};   // rowNo -> {col1based -> s}
    const rowAttrs = {};   // rowNo -> chuỗi thuộc tính của <row ...>
    const rowRe = /<row r="(\d+)"([^>]*)(\/>|>([\s\S]*?)<\/row>)/g;
    let m;
    while ((m = rowRe.exec(xml))) {
      const rn = +m[1];
      rowAttrs[rn] = m[2].trim();
      const body = m[4] || '';
      const map = {};
      const cRe = /<c r="([A-Z]+)\d+"(?:\s+s="(\d+)")?/g;
      let cm;
      while ((cm = cRe.exec(body))) {
        let col = 0; for (const ch of cm[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
        map[col] = cm[2] != null ? cm[2] : null;
      }
      styleMap[rn] = map;
    }
    return { styleMap, rowAttrs };
  }

  function genCell(ref, s, cell) {
    const sAttr = (s != null) ? ` s="${s}"` : '';
    if (!cell || cell.v === '' || cell.v == null) return `<c r="${ref}"${sAttr}/>`;
    if (cell.num && cell.v !== '' && !isNaN(+cell.v)) {
      return `<c r="${ref}"${sAttr}><v>${+cell.v}</v></c>`;
    }
    return `<c r="${ref}"${sAttr} t="inlineStr"><is><t xml:space="preserve">${xmlEsc(cell.v)}</t></is></c>`;
  }

  function genRow(rn, attrs, styleRow, defStyleRow, valsByCol) {
    let cells = '';
    for (let c = 1; c <= COLS; c++) {
      const ref = colLetter(c) + rn;
      let s = styleRow && (c in styleRow) ? styleRow[c] : (defStyleRow ? defStyleRow[c] : null);
      cells += genCell(ref, s, valsByCol[c]);
    }
    return `<row r="${rn}" ${attrs}>${cells}</row>`;
  }

  function buildFromTemplate(state, templateBytes, fflate) {
    const files = fflate.unzipSync(templateBytes);
    const SHEET = 'xl/worksheets/sheet1.xml'; // FORMAT
    const dec = new TextDecoder('utf-8');
    const enc = new TextEncoder();
    let xml = dec.decode(files[SHEET]);

    const { styleMap, rowAttrs } = parseTemplate(xml);
    const START = 10;
    const { rows, merges } = computeData(state, START);

    const dataRowNos = Object.keys(rows).map(Number);
    const maxDataRow = dataRowNos.length ? Math.max(...dataRowNos) : START - 1;

    // Style/định dạng mặc định cho dòng thân (dùng cho dòng tràn ngoài template)
    const defStyle = styleMap[12] || styleMap[11] || {};
    const defAttrs = rowAttrs[12] || rowAttrs[11] || 'spans="1:19"';

    // Ghi đè / chèn từng dòng dữ liệu, giữ nguyên style template
    for (let rn = START; rn <= maxDataRow; rn++) {
      const valsByCol = rows[rn] || {};
      const styleRow = styleMap[rn] || defStyle;
      const attrs = rowAttrs[rn] || defAttrs;
      const newRow = genRow(rn, attrs, styleRow, defStyle, valsByCol);

      if (rowAttrs[rn] != null) {
        // dòng đã có trong template -> thay thế tại chỗ
        const re = new RegExp(`<row r="${rn}"[^>]*(?:/>|>[\\s\\S]*?</row>)`);
        xml = xml.replace(re, newRow);
      } else {
        // dòng vượt quá template -> chèn trước </sheetData>
        xml = xml.replace('</sheetData>', newRow + '</sheetData>');
      }
    }

    // Thêm vùng merge cho dữ liệu vào mergeCells sẵn có
    if (merges.length) {
      const mm = /<mergeCells count="(\d+)">([\s\S]*?)<\/mergeCells>/.exec(xml);
      if (mm) {
        const base = +mm[1];
        let add = '';
        for (const [c1, r1, c2, r2] of merges) {
          add += `<mergeCell ref="${colLetter(c1)}${r1}:${colLetter(c2)}${r2}"/>`;
        }
        xml = xml.replace(mm[0],
          `<mergeCells count="${base + merges.length}">${mm[2]}${add}</mergeCells>`);
      }
    }

    files[SHEET] = enc.encode(xml);
    return fflate.zipSync(files, { level: 6 });
  }

  const api = { buildFromTemplate, computeData, parseTemplate };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.TemplateExport = api;
})(typeof window !== 'undefined' ? window : this);
