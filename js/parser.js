/*
 * parser.js — Đọc file Control Plan (CP) và rút dữ liệu cần cho P-FMEA.
 *
 * CP chỉ là NGUỒN dữ liệu. Trình đọc dò theo TÊN tiêu đề (không phụ thuộc cứng
 * vào vị trí cột), nên vẫn chạy được khi bố cục lệch đôi chút. Phần nào dò
 * không chuẩn thì người dùng sửa ở bước "Xem lại & chỉnh tay".
 *
 * Khối được dùng: "Quản lý đặc tính chất lượng" (Quality characteristics),
 * tức các hạng mục chất lượng -> mỗi hạng mục thành 1 yêu cầu / 1 dạng hỏng hóc.
 */
(function () {
  'use strict';

  const norm = (s) => (s == null ? '' : String(s)).replace(/\s+/g, ' ').trim();
  // Lấy dòng đầu tiên (thường là tiếng Việt) của một ô nhiều dòng VN/EN
  const firstLine = (s) => (s == null ? '' : String(s)).split('\n')[0].trim();

  // Ký tự có dấu tiếng Việt — dùng để nhận biết dòng tiếng Việt nối tiếp
  const VN_DIACRITIC = /[àáảãạăằắẳẵặâầấẩẫậèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵđ]/i;
  // Tên/nội dung tiếng Việt có thể bị xuống dòng GIỮA CHỪNG trong ô CP
  // (vd "KIỂM TRA PHÔI\nĐẦU VÀO\n入荷検査"). Lấy dòng đầu rồi GHÉP tiếp các dòng
  // tiếng Việt (có dấu) liền sau; dừng khi gặp dòng tiếng Anh/Nhật (không dấu VN).
  const vnText = (s) => {
    if (s == null) return '';
    const parts = String(s).split('\n').map((x) => x.trim()).filter(Boolean);
    if (!parts.length) return '';
    const out = [parts[0]];
    for (let i = 1; i < parts.length; i++) {
      if (VN_DIACRITIC.test(parts[i])) out.push(parts[i]); else break;
    }
    return out.join(' ').trim();
  };

  function decodeAddr(addr) {
    const m = /^([A-Z]+)(\d+)$/.exec(addr);
    if (!m) return null;
    let col = 0;
    for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
    return { col: col - 1, row: parseInt(m[2], 10) - 1 }; // 0-based
  }

  // Tìm tất cả ô có giá trị chuỗi chứa `needle`
  function findCells(ws, needle) {
    const out = [];
    for (const addr in ws) {
      if (addr[0] === '!') continue;
      const v = ws[addr] && ws[addr].v;
      if (typeof v === 'string' && v.includes(needle)) {
        const rc = decodeAddr(addr);
        if (rc) out.push({ addr, row: rc.row, col: rc.col, v });
      }
    }
    return out;
  }

  // Giá trị ô theo (row,col) 0-based, có truy ngược merge để lấy ô gốc.
  // Nếu ô dùng font Symbol (cellStyles: true) thì chuyển ký tự đặc biệt sang Unicode.
  function cellRC(ws, row, col, merges) {
    const a = XLSX.utils.encode_cell({ r: row, c: col });
    if (ws[a] && ws[a].v != null && norm(ws[a].v) !== '') return applySymbolFont(ws[a], ws[a].v);
    if (merges) {
      for (const m of merges) {
        if (row >= m.s.r && row <= m.e.r && col >= m.s.c && col <= m.e.c) {
          const ga = XLSX.utils.encode_cell({ r: m.s.r, c: m.s.c });
          if (ws[ga]) return applySymbolFont(ws[ga], ws[ga].v);
          return undefined;
        }
      }
    }
    return ws[a] ? applySymbolFont(ws[a], ws[a].v) : undefined;
  }

  // Cột kết thúc của vùng merge chứa (row,col); nếu không nằm trong merge nào
  // thì trả về chính col đó.
  function mergeEndCol(merges, row, col) {
    for (const m of merges) {
      if (row >= m.s.r && row <= m.e.r && col >= m.s.c && col <= m.e.c) return m.e.c;
    }
    return col;
  }

  // Tên công đoạn: tìm nhãn "Tên công đoạn"/"Process name", lấy ô có giá trị
  // bên phải trong cùng hàng (bỏ qua vùng merge của chính nhãn).
  function findProcessName(ws, merges) {
    const labels = findCells(ws, 'Tên công đoạn').concat(findCells(ws, 'Process name'));
    for (const lab of labels) {
      const startC = mergeEndCol(merges, lab.row, lab.col) + 1;
      for (let c = startC; c <= startC + 14; c++) {
        const v = cellRC(ws, lab.row, c, merges);
        if (typeof v === 'string' && norm(v) && !/Tên công đoạn|Process name/.test(v)) {
          return vnText(v);
        }
      }
    }
    return '';
  }

  // Xác định cột theo header trong phạm vi [colMin, colMax)
  function pickHeaderCol(cells, colMin, colMax) {
    const inRange = cells.filter((c) => c.col >= colMin && c.col < colMax);
    if (!inRange.length) return -1;
    inRange.sort((a, b) => a.col - b.col);
    return inRange[0].col;
  }

  function looksLikeTolerance(s) {
    s = norm(s);
    return !!s && /[+\-±]/.test(s);
  }

  // Trong CP, ký hiệu đường kính Ø thường được gõ là chữ "F" với font Symbol.
  // Chuyển 'F'/'f' dùng làm ký hiệu đường kính sang "Ø":
  //   (a) đứng trước chữ số, dấu thập phân, hoặc ký hiệu dung sai ±
  //   (b) đứng ở cuối chuỗi (vd "Max F" → "Max Ø")
  //   (c) là toàn bộ giá trị (chỉ một ký tự F/f)
  function fixDiameter(s) {
    s = (s == null ? '' : String(s));
    s = s.replace(/(^|[^0-9A-Za-zÀ-ỹ])[Ff](?=\s*[0-9.±])/g, '$1Ø');
    s = s.replace(/(^|[^0-9A-Za-zÀ-ỹ])[Ff](?=\s*$)/g, '$1Ø');
    if (/^[Ff]$/.test(s.trim())) s = 'Ø';
    return s;
  }

  // Nếu SheetJS đọc được thông tin font (cellStyles: true), chuyển ký tự
  // trong ô dùng font Symbol sang Unicode tương ứng ('F' → 'Ø').
  // Chỉ chuyển 'F' đứng độc lập (không nằm trong từ chữ cái), tránh sai từ.
  function applySymbolFont(cell, val) {
    if (!cell || !cell.s || !cell.s.font || typeof val !== 'string') return val;
    if (!/symbol/i.test(String(cell.s.font.name || ''))) return val;
    return val.replace(/(^|[^A-Za-zÀ-ỹ])F(?=[^A-Za-zÀ-ỹ]|$)/g, '$1Ø');
  }

  // Fallback cho bố cục GL SQS0831: khối "Quản lý đặc tính chất lượng" xếp DỌC
  // (cùng cột với "Quản lý điều kiện chế tạo"), tiêu đề nằm NGAY DƯỚI nhãn khối;
  // cột giá trị tên là "Giá trị quản lý"; mỗi hạng mục có dòng tiếng Việt + dòng
  // tiếng Anh riêng. Chỉ dùng khi cách dò theo cột-biên (parseSheet) không ra cột tên.
  function parseStacked(ws, merges, sheetName, processName) {
    const q = findCells(ws, 'Quản lý đặc tính')[0];
    if (!q) return null;
    let spanS = q.col, spanE = 1e9;
    for (const m of merges) {
      if (q.row >= m.s.r && q.row <= m.e.r && q.col >= m.s.c && q.col <= m.e.c) { spanS = m.s.c; spanE = m.e.c; break; }
    }
    const qRow = q.row;
    let blockEnd = XLSX.utils.decode_range(ws['!ref']).e.r;
    ['Quản lý điều kiện', 'Hạng mục cấm', 'Following rules'].forEach((n) =>
      findCells(ws, n).forEach((c) => { if (c.row > qRow && c.row - 1 < blockEnd) blockEnd = c.row - 1; }));
    const below = (cells) => {
      const f = cells.filter((c) => c.row > qRow && c.col >= spanS && c.col <= spanE)
        .sort((a, b) => a.row - b.row || a.col - b.col);
      return f[0] || null;
    };
    const nameH = below(findCells(ws, 'Hạng mục quản lý'));
    if (!nameH) return null;
    const nameCol = nameH.col, headerRow = nameH.row;
    const inRow = (cells) => {
      const f = cells.filter((c) => c.col >= spanS && c.col <= spanE && Math.abs(c.row - headerRow) <= 2)
        .sort((a, b) => Math.abs(a.row - headerRow) - Math.abs(b.row - headerRow) || a.col - b.col);
      return f[0] ? f[0].col : -1;
    };
    const specCol = inRow(findCells(ws, 'Giá trị quản lý')
      .concat(findCells(ws, 'Giá trị tiêu chuẩn')).concat(findCells(ws, 'Control value')).concat(findCells(ws, 'Spec value')));
    const methodCol = inRow(findCells(ws, 'Phương pháp xác nhận').concat(findCells(ws, 'Check method')).concat(findCells(ws, 'Check\nMethod')));
    const freqCol = inRow(findCells(ws, 'Tần suất xác nhận').concat(findCells(ws, 'Check frequency')).concat(findCells(ws, 'Check\nFrequency')));
    const scCol = inRow(findCells(ws, 'Đặc tính đặc thù').concat(findCells(ws, 'S.C')));

    // Hàng bắt đầu hạng mục: ô tên có chữ tiếng Việt (bỏ dòng dịch tiếng Anh riêng).
    const starts = [];
    for (let r = headerRow + 1; r <= blockEnd; r++) {
      const nv = ws[XLSX.utils.encode_cell({ r, c: nameCol })];
      const txt = nv ? norm(nv.v) : '';
      if (txt && VN_DIACRITIC.test(txt)) starts.push(r);
    }
    const items = [];
    const specEnd = methodCol > specCol ? methodCol : spanE + 1;
    for (let i = 0; i < starts.length; i++) {
      const r0 = starts[i];
      const r1 = (i + 1 < starts.length ? starts[i + 1] : blockEnd + 1) - 1;
      const name = fixDiameter(vnText(cellRC(ws, r0, nameCol, merges)));
      if (!name) continue;
      // spec/tol: quét vùng [specCol, methodCol); bỏ R/L, Max/Min; phân loại dung sai vs trị số.
      const specParts = [], tolParts = [];
      if (specCol >= 0) {
        for (let c = specCol; c < specEnd; c++) {
          let val = '';
          for (let r = r0; r <= r1; r++) {
            const a = XLSX.utils.encode_cell({ r, c });
            if (ws[a] && norm(ws[a].v)) { val = norm(ws[a].v); break; }
          }
          if (!val || /^(R\/L|L\/R|Max\.?|Min\.?)$/i.test(val)) continue;
          if (looksLikeTolerance(val)) tolParts.push(val); else specParts.push(fixDiameter(val));
        }
      }
      const spec = specParts.join(' ');
      const tol = tolParts.join('/');
      const method = methodCol >= 0 ? vnText(cellRC(ws, r0, methodCol, merges)) : '';
      const freq = freqCol >= 0 ? String(cellRC(ws, r0, freqCol, merges) || '').replace(/\n+/g, ' ').replace(/\s+/g, ' ').trim() : '';
      // SC: chỉ đọc ĐÚNG ô tại hàng tên hạng mục (r0), KHÔNG theo merge
      // (tránh kế thừa SC của hạng mục phía trên qua ô gộp dọc, và tránh vớ chữ "S" lạc ở dưới).
      const sc = scCol >= 0 ? norm(cellRC(ws, r0, scCol, [])) : '';
      let requirement = name;
      if (spec) requirement += ': ' + spec + (tol ? '(' + tol + ')' : '');
      items.push({ no: items.length + 1, name, spec, tol, requirement, method, freq, sc });
    }
    return { sheetName, processName, items };
  }

  /*
   * Trả về:
   * { sheetName, processName, items: [ {no, name, spec, tol, requirement,
   *   method, freq, sc} ] }
   */
  function parseSheet(wb, sheetName) {
    const ws = wb.Sheets[sheetName];
    const merges = ws['!merges'] || [];
    const processName = findProcessName(ws, merges);

    // Ranh giới khối chất lượng: từ "Quản lý đặc tính chất lượng" đến
    // "Quản lý điều kiện chế tạo" (nếu có).
    const qBlock = findCells(ws, 'Quản lý đặc tính');
    const cBlock = findCells(ws, 'Quản lý điều kiện');
    const colMin = qBlock.length ? Math.min(...qBlock.map((c) => c.col)) : 0;
    const colMax = cBlock.length ? Math.min(...cBlock.map((c) => c.col)) : 9999;

    // Các cột dữ liệu, dò theo tên tiêu đề trong khối chất lượng
    const nameCol = pickHeaderCol(findCells(ws, 'Hạng mục quản lý'), colMin, colMax);
    const specCol = pickHeaderCol(
      findCells(ws, 'Giá trị tiêu chuẩn').concat(findCells(ws, 'Spec value')),
      colMin, colMax
    );
    const methodCol = pickHeaderCol(
      findCells(ws, 'Phương pháp xác nhận').concat(findCells(ws, 'Check\nMethod')),
      colMin, colMax
    );
    const freqCol = pickHeaderCol(
      findCells(ws, 'Tần suất xác nhận').concat(findCells(ws, 'Check\nFrequency')),
      colMin, colMax
    );
    const scCol = pickHeaderCol(
      findCells(ws, 'Đặc tính đặc thù').concat(findCells(ws, 'S.C')),
      colMin, colMax
    );

    if (nameCol < 0) {
      // Thử bố cục GL SQS0831 (khối xếp dọc) trước khi báo lỗi.
      const alt = parseStacked(ws, merges, sheetName, processName);
      if (alt && alt.items.length) return alt;
      return { sheetName, processName, items: [], error: 'Không tìm thấy cột "Hạng mục quản lý" trong khối đặc tính chất lượng.' };
    }

    // Hàng tiêu đề = hàng của ô "Hạng mục quản lý"
    const headerCell = findCells(ws, 'Hạng mục quản lý').filter((c) => c.col === nameCol)[0];
    const headerRow = headerCell ? headerCell.row : 0;

    const range = XLSX.utils.decode_range(ws['!ref']);
    const maxRow = range.e.r;

    // Tìm cột số thứ tự (№ / No / STT) trong khối chất lượng, bên trái cột tên.
    // CP đánh số mỗi hạng mục chất lượng ở cột này; 1 hạng mục có thể trải nhiều
    // dòng (chỉ dòng đầu mới có số). Nếu tìm thấy → dùng số làm ranh giới hạng mục
    // (chính xác hơn việc dựa vào ô tên, tránh đếm dư các dòng nối tiếp/lặp số).
    let noCol = -1;
    for (const addr in ws) {
      if (addr[0] === '!') continue;
      const v = ws[addr] && ws[addr].v;
      if (v == null) continue;
      if (/^\s*(№|no\.?|stt|s\.?t\.?t\.?|số\s*tt)\s*$/i.test(String(v))) {
        const rc = decodeAddr(addr);
        if (rc && rc.col >= colMin && rc.col < nameCol && Math.abs(rc.row - headerRow) <= 2) {
          noCol = rc.col; break;
        }
      }
    }

    // Tìm các hàng bắt đầu hạng mục.
    const starts = [];
    if (noCol >= 0) {
      // Theo cột số:
      //  • Hàng có TÊN và SỐ mới (chưa gặp)  -> hạng mục mới.
      //  • Hàng có TÊN nhưng số TRỐNG/LẶP:
      //       - cùng tên với hạng mục đang xét  -> dòng nối tiếp (gộp lên trên).
      //       - khác tên (tác giả quên đánh số)  -> vẫn là hạng mục mới.
      const seenNo = new Set();
      let curName = null;
      for (let r = headerRow + 1; r <= maxRow; r++) {
        const nv = ws[XLSX.utils.encode_cell({ r, c: nameCol })];
        const txt = nv ? norm(nv.v) : '';
        if (!txt || txt === 'Control Item') continue;
        const nameKey = txt.toLowerCase();
        const noV = ws[XLSX.utils.encode_cell({ r, c: noCol })];
        const noTxt = noV ? norm(noV.v) : '';
        if (noTxt && !seenNo.has(noTxt)) {
          seenNo.add(noTxt); starts.push(r); curName = nameKey; // số mới
        } else if (nameKey !== curName) {
          starts.push(r); curName = nameKey;                    // khác tên -> hạng mục mới
        }
        // còn lại: số trống/lặp + cùng tên -> nối tiếp, bỏ qua
      }
    } else {
      // Không có cột số: dựa vào ô tên không rỗng (như cũ).
      for (let r = headerRow + 1; r <= maxRow; r++) {
        const nv = ws[XLSX.utils.encode_cell({ r, c: nameCol })];
        const txt = nv ? norm(nv.v) : '';
        if (txt && txt !== 'Control Item') starts.push(r);
      }
    }

    const items = [];
    for (let i = 0; i < starts.length; i++) {
      const r0 = starts[i];
      const r1 = (i + 1 < starts.length ? starts[i + 1] : maxRow + 1) - 1; // hàng cuối của hạng mục

      const name = fixDiameter(vnText(cellRC(ws, r0, nameCol, merges)));
      if (!name) continue;

      // spec: gom TẤT CẢ ô có giá trị trong specCol (r0..r1), nối lại.
      // Một số hạng mục có spec trải 2 ô (vd "Ø14.5" ở ô trên và "±0.2" ở ô dưới).
      // Dùng Set để loại trùng (bản dịch Anh/Nhật trùng số với bản tiếng Việt).
      let spec = '';
      if (specCol >= 0) {
        const specParts = [];
        for (let r = r0; r <= r1; r++) {
          const v = cellRC(ws, r, specCol, [] /*không truy merge để tránh lặp*/);
          const fv = norm(v) ? fixDiameter(firstLine(v)) : '';
          if (fv && !specParts.includes(fv)) specParts.push(fv);
        }
        spec = specParts.join(' ');
      }

      // tolerance / giá trị kề spec: gom các ô ở vài cột bên phải spec.
      // Nếu ô có dấu ±/-/+ → dung sai (ghi vào tol, hiển thị trong ngoặc).
      // Nếu không có dấu đó (vd "0.1", "Rz12.5") → ghép vào spec để không mất.
      let tols = [], specExtras = [];
      if (specCol >= 0) {
        for (let r = r0; r <= r1; r++) {
          for (let c = specCol + 1; c <= Math.min(specCol + 5, colMax - 1); c++) {
            const a = XLSX.utils.encode_cell({ r, c });
            const v = ws[a] ? ws[a].v : undefined;
            const nv = norm(v);
            if (!nv) continue;
            if (looksLikeTolerance(v)) {
              tols.push(nv);
            } else {
              const fv = fixDiameter(nv);
              if (!specExtras.includes(fv) && !specParts.includes(fv)) specExtras.push(fv);
            }
          }
        }
      }
      tols = [...new Set(tols)];
      const tol = tols.join('/');
      // Ghép thêm các giá trị kề không phải dung sai vào cuối spec
      if (specExtras.length) spec = [spec, ...specExtras].filter(Boolean).join(' ');

      const method = methodCol >= 0 ? vnText(cellRC(ws, r0, methodCol, merges)) : '';
      const freq = freqCol >= 0 ? String(cellRC(ws, r0, freqCol, merges) || '').replace(/\n+/g, ' ').replace(/\s+/g, ' ').trim() : '';
      // SC: chỉ đọc ĐÚNG ô tại hàng tên hạng mục (r0), KHÔNG theo merge.
      // Không theo merge → tránh kế thừa SC của hạng mục PHÍA TRÊN (ô gộp dọc).
      // Không quét xuống r1 → tránh vớ phải chữ "S" lạc ở dưới khi hạng mục cuối có r1=maxRow.
      const sc = scCol >= 0 ? norm(cellRC(ws, r0, scCol, [])) : '';

      // Chuỗi yêu cầu: "Tên: spec(tol)"
      let requirement = name;
      if (spec) requirement += ': ' + spec + (tol ? '(' + tol + ')' : '');

      items.push({
        no: items.length + 1,
        name, spec, tol, requirement, method, freq, sc,
      });
    }

    return { sheetName, processName, items };
  }

  window.CPParser = {
    listSheets(wb) {
      return wb.SheetNames.filter((n) => {
        const ws = wb.Sheets[n];
        return ws && ws['!ref'] && ws['!ref'] !== 'A1';
      });
    },
    parseSheet,
  };
})();
