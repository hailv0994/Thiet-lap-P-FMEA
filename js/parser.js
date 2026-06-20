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

  // Giá trị ô theo (row,col) 0-based, có truy ngược merge để lấy ô gốc
  function cellRC(ws, row, col, merges) {
    const a = XLSX.utils.encode_cell({ r: row, c: col });
    if (ws[a] && ws[a].v != null && norm(ws[a].v) !== '') return ws[a].v;
    if (merges) {
      for (const m of merges) {
        if (row >= m.s.r && row <= m.e.r && col >= m.s.c && col <= m.e.c) {
          const ga = XLSX.utils.encode_cell({ r: m.s.r, c: m.s.c });
          return ws[ga] ? ws[ga].v : undefined;
        }
      }
    }
    return ws[a] ? ws[a].v : undefined;
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
      return { sheetName, processName, items: [], error: 'Không tìm thấy cột "Hạng mục quản lý" trong khối đặc tính chất lượng.' };
    }

    // Hàng tiêu đề = hàng của ô "Hạng mục quản lý"
    const headerCell = findCells(ws, 'Hạng mục quản lý').filter((c) => c.col === nameCol)[0];
    const headerRow = headerCell ? headerCell.row : 0;

    const range = XLSX.utils.decode_range(ws['!ref']);
    const maxRow = range.e.r;

    // Tìm các hàng bắt đầu hạng mục (ô tên không rỗng), bỏ qua hàng tiêu đề phụ
    const starts = [];
    for (let r = headerRow + 1; r <= maxRow; r++) {
      const nv = ws[XLSX.utils.encode_cell({ r, c: nameCol })];
      const txt = nv ? norm(nv.v) : '';
      if (txt && txt !== 'Control Item') starts.push(r);
    }

    const items = [];
    for (let i = 0; i < starts.length; i++) {
      const r0 = starts[i];
      const r1 = (i + 1 < starts.length ? starts[i + 1] : maxRow + 1) - 1; // hàng cuối của hạng mục

      const name = vnText(cellRC(ws, r0, nameCol, merges));
      if (!name) continue;

      // spec: ô đầu tiên có giá trị trong specCol
      let spec = '';
      if (specCol >= 0) {
        for (let r = r0; r <= r1; r++) {
          const v = cellRC(ws, r, specCol, [] /*không truy merge để tránh lặp*/);
          if (norm(v)) { spec = firstLine(v); break; }
        }
      }

      // tolerance: gom các ô dạng dung sai ở vài cột bên phải spec, trong span
      let tols = [];
      if (specCol >= 0) {
        for (let r = r0; r <= r1; r++) {
          for (let c = specCol + 1; c <= Math.min(specCol + 5, colMax - 1); c++) {
            const a = XLSX.utils.encode_cell({ r, c });
            const v = ws[a] ? ws[a].v : undefined;
            if (looksLikeTolerance(v)) tols.push(norm(v));
          }
        }
      }
      tols = [...new Set(tols)];
      const tol = tols.join('/');

      const method = methodCol >= 0 ? vnText(cellRC(ws, r0, methodCol, merges)) : '';
      const freq = freqCol >= 0 ? String(cellRC(ws, r0, freqCol, merges) || '').replace(/\n+/g, ' ').replace(/\s+/g, ' ').trim() : '';
      const sc = scCol >= 0 ? norm(cellRC(ws, r0, scCol, merges)) : '';

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
