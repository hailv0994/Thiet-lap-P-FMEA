/* =====================================================================
 * app.js — P-FMEA Builder
 * Quản lý state, dựng bảng P-FMEA, chấm điểm S theo tiêu chuẩn, xuất Excel.
 * ===================================================================== */
(function () {
  'use strict';

  // ----------------------------- State -----------------------------
  let UID = 1;
  const uid = (p) => p + (UID++);
  const state = {
    meta: { dept: '', product: '', line: '', model: '' },
    processes: [],
  };
  let workbook = null;               // workbook CP đang mở
  const LS_PROJECTS = 'pfmea_projects_v1';
  const LS_AUTOSAVE = 'pfmea_autosave_v1'; // tự lưu phiên làm việc (meta + nội dung)
  const LS_GEMINI_KEY = 'pfmea_gemini_key_v1';
  const LS_GEMINI_MODEL = 'pfmea_gemini_model_v1';
  const LS_CONTEXT = 'pfmea_context_v1';   // bối cảnh AI theo bộ phận
  const LS_PHRASES = 'pfmea_phrases_v1';   // bộ nhớ câu đã nhập theo cột + bộ phận

  // ----- Cấu hình đồng bộ đám mây (Supabase) — anon key là khóa CÔNG KHAI -----
  const SB_URL = 'https://iccrgjtkaizosocrxsql.supabase.co';
  const SB_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImljY3JnanRrYWl6b3NvY3J4c3FsIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODIwMTcyNzEsImV4cCI6MjA5NzU5MzI3MX0.oOHGGZOfteIF7he7u8IKaIejLfDNDCLG1sa3uqQftfc';
  let sb = null;

  const $ = (s) => document.querySelector(s);
  const esc = (s) => (s == null ? '' : String(s))
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');

  // --------------------- Tạo dữ liệu mặc định ----------------------
  function newCause() {
    return { id: uid('c'), category: '', cause: '', pastTrouble: '', occurrence: '',
      prevention: '', detectCause: '', detectExtra: '', detection: '',
      action: '', responsible: '', actionTaken: '', s2: '', o2: '', d2: '' };
  }
  const FOUR_M = ['Man', 'Machine', 'Method', 'Material'];

  function detectAuto(item) {
    let s = 'Kiểm tra ' + (item.name || '');
    if (item.method) s += ' bằng ' + item.method;
    if (item.freq) s += ' theo tần suất ' + item.freq;
    return s;
  }

  function reqFromItem(item) {
    return {
      id: uid('r'),
      reqText: item.requirement || item.name || '',
      failureMode: (item.name || '') + ' không đạt',
      effectAnalysis: '',
      effectStdText: '',
      effectScope: '',
      severity: '',
      classification: item.sc || '',
      detectFailureAuto: detectAuto(item),
      causes: [newCause()],
    };
  }

  // Xác định danh sách dạng hỏng hóc dựa vào spec và tolerance của hạng mục CP
  function failureModesFor(item) {
    const name = norm(item.name || '');
    const spec = norm(item.spec || '');
    const tol  = norm(item.tol  || '');
    const specDisplay = spec ? ' ' + spec : '';
    // Quản lý 1 phía (max/min/≤/≥) → 1 dạng hỏng
    if (/^(max|min|[≤≥]|tối\s*đa|tối\s*thiểu|không\s*quá|ít\s*nhất)/i.test(spec)) {
      return [name + specDisplay + ' không đạt'];
    }
    // Không có spec, hoặc spec toàn chữ (không số, không ~) → 1 dạng hỏng
    if (!spec || (!tol && !/[\d~]/.test(spec))) {
      return [name + ' không đạt'];
    }
    // Có dung sai (±/+x-y) hoặc dạng khoảng (4~5) → 2 phía → 2 dạng hỏng
    if (tol || /~/.test(spec)) {
      const base = name + specDisplay;
      return [base + ' lớn hơn tiêu chuẩn', base + ' nhỏ hơn tiêu chuẩn'];
    }
    // Spec là số nhưng không có dung sai → 1 dạng hỏng
    return [name + specDisplay + ' không đạt'];
  }

  // Tạo 1 hoặc 2 yêu cầu từ 1 hạng mục CP (tùy số dạng hỏng).
  // Khi tách 2 dạng hỏng (lớn hơn / nhỏ hơn) → 2 yêu cầu ĐỘC LẬP (mỗi cái có
  // nguyên nhân/ S / O / D riêng) nhưng cùng splitId để cột "Yêu cầu" chỉ hiện 1 dòng.
  function reqsFromItem(item) {
    const modes = failureModesFor(item);
    const splitId = modes.length > 1 ? uid('split') : '';
    return modes.map((fm) => ({
      id: uid('r'),
      splitId,
      reqText: item.requirement || item.name || '',
      failureMode: fm,
      effectAnalysis: '', effectStdText: '', effectScope: '', severity: '',
      classification: item.sc || '',
      detectFailureAuto: detectAuto(item),
      mergeId: '',
      causes: [newCause()],
    }));
  }

  // Đánh số yêu cầu: các yêu cầu cùng splitId (tách dạng hỏng) dùng CHUNG 1 số.
  // Trả về map { reqId -> số thứ tự yêu cầu }.
  function reqNumbers(reqs) {
    const map = {}; const seen = {}; let n = 0;
    reqs.forEach((r) => {
      const key = r.splitId || r.id;
      if (!(key in seen)) { n++; seen[key] = n; }
      map[r.id] = seen[key];
    });
    return map;
  }

  // ------------------------ Bộ chọn ảnh hưởng (S) -------------------
  function severityOptions(selectedIdx) {
    const groups = { product: [], process: [] };
    window.SEVERITY_TABLE.forEach((row, i) => {
      const lbl = `S=${row.rank} · ${row.category} — ${row.text}`;
      const sel = (i === selectedIdx) ? ' selected' : '';
      groups[row.scope].push(
        `<option value="${i}"${sel}>${esc(lbl.length > 90 ? lbl.slice(0, 90) + '…' : lbl)}</option>`
      );
    });
    return `<option value="">— Chọn ảnh hưởng (ý②) để tự chấm S —</option>
      <optgroup label="Ảnh hưởng đến SẢN PHẨM (khách hàng)">${groups.product.join('')}</optgroup>
      <optgroup label="Ảnh hưởng đến CÔNG ĐOẠN (chế tạo/lắp ráp)">${groups.process.join('')}</optgroup>`;
  }

  function findSeverityIdx(req) {
    if (!req.effectStdText) return -1;
    return window.SEVERITY_TABLE.findIndex(
      (r) => r.text === req.effectStdText && r.scope === req.effectScope
    );
  }

  // -------------------------- Tra cứu state ------------------------
  function getProc(pid) { return state.processes.find((p) => p.id === pid); }
  function getReq(pid, rid) { const p = getProc(pid); return p && p.reqs.find((r) => r.id === rid); }
  function getCause(pid, rid, cid) { const r = getReq(pid, rid); return r && r.causes.find((c) => c.id === cid); }

  const rpnOf = (req, cause) => {
    const s = +req.severity, o = +cause.occurrence, d = +cause.detection;
    return (s && o && d) ? s * o * d : '';
  };

  // Nhận diện "kiểm tra bằng giác quan" (tay/mắt/đếm…) từ nội dung phát hiện.
  // Quét cả ① Phát hiện ra nguyên nhân (detectCause) và ② Phát hiện ra dạng hỏng hóc
  // (detectFailureAuto — chứa "bằng [phương pháp]"). Nếu có từ khóa giác quan → true.
  const VISUAL_KW = ['tay', 'mắt', 'đếm', 'nhìn', 'quan sát', 'thị giác', 'cảm quan',
    'cảm nhận', 'sờ', 'nghe', 'ngửi', 'nếm', 'trực quan', 'mục kiểm', 'thủ công'];
  function hasWord(text, kw) {
    try {
      const re = new RegExp('(^|[^\\p{L}\\p{N}])' + kw + '($|[^\\p{L}\\p{N}])', 'u');
      return re.test(text);
    } catch (e) { return text.indexOf(kw) >= 0; }
  }
  function isVisualDetect(req, cause) {
    const text = ((cause.detectCause || '') + ' ' + (req.detectFailureAuto || '')).toLowerCase();
    return VISUAL_KW.some((kw) => hasWord(text, kw));
  }

  // Tiêu chuẩn thực hiện biện pháp đề xuất (GL): ô Biện pháp đề xuất cần đề xuất khi
  //   S ≥ 9, hoặc O ≥ 4, hoặc (S ≥ 7 và D ≥ 6), hoặc (S ≤ 6 và D ≥ 7 và KHÔNG giác quan).
  // Khi đúng → viền đỏ trên web nhắc người làm; KHÔNG in ra (chỉ là gợi ý màn hình).
  function needsAction(req, cause) {
    const s = +req.severity || 0, o = +cause.occurrence || 0, d = +cause.detection || 0;
    if (s >= 9) return true;
    if (o >= 4) return true;
    if (s >= 7 && d >= 6) return true;       // trường hợp 1
    // trường hợp 2: S≤6 và D≥7, chỉ áp dụng khi phát hiện KHÔNG bằng giác quan
    if (s >= 1 && s <= 6 && d >= 7 && !isVisualDetect(req, cause)) return true;
    return false;
  }

  // ---------------- Gộp dạng hỏng hóc (merge group) ----------------
  const norm = (v) => String(v == null ? '' : v).trim();

  // Tách chuỗi "Kiểm tra [tên] bằng [method] theo tần suất [freq]"
  // → { name: 'tên', suffix: ' bằng method theo tần suất freq' }
  // Nếu không có " bằng " → suffix = '' (không thể so sánh → dùng toàn bộ chuỗi).
  function parseDetect(str) {
    const s = norm(str);
    const bang = s.indexOf(' bằng ');
    const prefix = 'Kiểm tra ';
    if (bang < 0) return { name: s.startsWith(prefix) ? s.slice(prefix.length) : s, suffix: '' };
    return {
      name: s.startsWith(prefix) ? s.slice(prefix.length, bang) : s.slice(0, bang),
      suffix: s.slice(bang), // " bằng Height gage theo tần suất đầu ca"
    };
  }
  // Tên hạng mục từ failureMode (bỏ đuôi " không đạt").
  const itemNameFrom = (r) => norm(r.failureMode).replace(/\s*không đạt\s*$/i, '').trim();

  // Chữ ký để so sánh: chỉ phần method+tần suất (sau " bằng ").
  // Hai hạng mục có thể gộp khi cùng dụng cụ + tần suất kiểm tra, dù tên khác nhau.
  function reqSig(r) {
    const { suffix } = parseDetect(r.detectFailureAuto);
    return norm(suffix) || norm(r.detectFailureAuto);
  }

  const VN_DIACR = /[àáảãạăằắẳẵặâầấẩẫậèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵđ]/i;
  // Tên đo lường gốc của dạng hỏng: bỏ đuôi "lớn/nhỏ hơn tiêu chuẩn" và "không đạt".
  const baseMeasureName = (s) => norm(s)
    .replace(/\s*(lớn hơn tiêu chuẩn|nhỏ hơn tiêu chuẩn|không đạt)\s*$/i, '').trim();
  // Tiền tố chung (theo từ) của một danh sách tên. VD ["kích thước A","kích thước B"]
  // -> "kích thước".
  function commonPrefixWords(names) {
    if (!names.length) return '';
    const split = names.map((n) => norm(n).split(' '));
    const first = split[0];
    let k = 0;
    for (; k < first.length; k++) {
      const w = first[k].toLowerCase();
      if (!split.every((arr) => arr[k] && arr[k].toLowerCase() === w)) break;
    }
    return first.slice(0, k).join(' ').trim();
  }
  // Gộp danh sách tên đo lường thành chủ ngữ cho câu "Kiểm tra …".
  //  • Cùng loại đo lường (tiền tố ≥2 từ), phần khác nhau là MÃ NGẮN (A,B…) hoặc
  //    chứa GIÁ TRỊ SỐ (203.3, Ø24, 191.9…)  -> "các <tiền tố chung>".
  //  • Khác loại thật sự (cứng / nhám, tiền tố 1 từ như "Độ") -> liệt kê "X và Y".
  function summarizeNames(names) {
    const prefix = commonPrefixWords(names);
    const wc = prefix ? prefix.split(' ').length : 0;
    if (prefix && wc >= 2) {
      const rems = names.map((n) => norm(n).slice(prefix.length).trim());
      const ok = rems.every((rm) =>
        rm === '' || /\d/.test(rm) || (rm.length <= 4 && !VN_DIACR.test(rm)));
      if (ok) return 'các ' + prefix.charAt(0).toLowerCase() + prefix.slice(1);
    }
    return names.join(' và ');
  }
  // Xây lại chuỗi detectFailureAuto gộp từ danh sách thành viên nhóm.
  function buildGroupDetect(members, suffix) {
    const names = members.map((m) => baseMeasureName(itemNameFrom(m))).filter(Boolean);
    if (!names.length || !suffix) return members[0] ? members[0].detectFailureAuto : '';
    if (names.length === 1) return 'Kiểm tra ' + names[0] + suffix;
    return 'Kiểm tra ' + summarizeNames(names) + suffix;
  }
  // Gom các yêu cầu cùng mergeId thành nhóm; giữ thứ tự, đại diện = phần tử đầu.
  function reqGroups(reqs) {
    const seen = new Set(); const groups = [];
    reqs.forEach((r, ri) => {
      if (seen.has(r.id)) return;
      const members = [{ r, ri }]; seen.add(r.id);
      if (r.mergeId) reqs.forEach((r2, ri2) => {
        if (!seen.has(r2.id) && r2.mergeId === r.mergeId) { members.push({ r: r2, ri: ri2 }); seen.add(r2.id); }
      });
      groups.push(members);
    });
    return groups;
  }
  // Đồng bộ dữ liệu chung (mọi cột trừ reqText/failureMode) từ đại diện -> thành viên,
  // để khi tách nhóm mỗi yêu cầu vẫn giữ đầy đủ phân tích.
  function syncMergeGroup(p, rep) {
    if (!p || !rep || !rep.mergeId) return;
    p.reqs.forEach((r) => {
      if (r === rep || r.mergeId !== rep.mergeId) return;
      r.effectAnalysis = rep.effectAnalysis; r.effectStdText = rep.effectStdText;
      r.effectScope = rep.effectScope; r.severity = rep.severity;
      r.classification = rep.classification;
      // detectFailureAuto KHÔNG sync ở đây — được xây gộp riêng khi merge/unmerge.
      r.causes = rep.causes.map((c) => Object.assign({}, c, { id: uid('c') }));
    });
  }

  // ============================ RENDER =============================
  function buildHeader() {
    const H = (txt, cls = '') => `<th class="${cls}">${txt}</th>`;
    const thead = $('#fmea thead');
    thead.innerHTML = `
      <tr>
        ${H('Quy trình / Bước / Chức năng\nプロセス ステップ/機能\n(Hạng mục yêu cầu 要求事項)')}
        ${H('Dạng hỏng hóc mang tính tiềm ẩn\n潜在的故障モード')}
        ${H('Ảnh hưởng của hỏng hóc tiềm ẩn\n潜在的故障影響')}
        ${H('Mức độ\nnghiêm trọng\n厳しさ (S)')}
        ${H('Phân loại\n分類')}
        ${H('Nguyên nhân của hỏng hóc\n潜在的故障原因')}
        ${H('Phản ánh lỗi\nquá khứ\n過去トラ反映')}
        ${H('Tần suất\nphát sinh\n発生頻度 (O)')}
        ${H('Quản lý hiện tại - Dự phòng\n現行のプロセス管理 予防')}
        ${H('Quản lý hiện tại - Phát hiện ra\n現行のプロセス管理 検出')}
        ${H('Phát hiện\n検出 (D)')}
        ${H('RPN')}
        ${H('Biện pháp đề xuất\n推奨処置')}
        ${H('Người chịu trách nhiệm\n& thời hạn\n責任者及び目標完了期限')}
        ${H('Kết quả xử lý　処置結果', 'grp')}
      </tr>
      <tr>
        ${H('Biện pháp được sử dụng\nvà ngày hoàn thành\n取られた処置及び完了日', 'grp')}
        ${H('S\n厳しさ', 'grp')}
        ${H('O\n発生頻度', 'grp')}
        ${H('D\n検出', 'grp')}
        ${H('RPN', 'grp')}
      </tr>`;
    // rowspan cho 14 cột đầu (A..N)
    const ths = thead.rows[0].cells;
    for (let i = 0; i < 14; i++) ths[i].rowSpan = 2;
    ths[14].colSpan = 5; // nhóm Kết quả xử lý
  }

  function procCellHTML(p, reqNo) {
    reqNo = reqNo || reqNumbers(p.reqs);
    const seen = new Set();
    const reqLines = p.reqs.map((r) => {
      const key = r.splitId || r.id;
      if (seen.has(key)) return '';   // yêu cầu đã tách dạng hỏng → chỉ hiện 1 dòng
      seen.add(key);
      return `
      <div class="req-line" data-proc="${p.id}" data-req="${r.id}">
        <span class="idx">${reqNo[r.id]}.</span>
        <textarea data-field="reqText" rows="2" placeholder="Yêu cầu (điều kiện chất lượng)">${esc(r.reqText)}</textarea>
        <button class="mini-btn danger" data-action="del-req" title="Xóa yêu cầu này">✕</button>
      </div>`;
    }).join('');
    return `<div class="proc-cell" data-proc="${p.id}">
        <div class="proc-head">
          <span class="proc-head-left">
            <input data-field="no" class="inp-no" style="width:23px" value="${esc(p.no)}" placeholder="STT" />.
            <input data-field="name" class="inp-pname" value="${esc(p.name)}" placeholder="Tên công đoạn" />
          </span>
          <span class="proc-move">
            <button class="mini-btn" data-action="move-up" title="Lên trên">▲</button>
            <button class="mini-btn" data-action="move-down" title="Xuống dưới">▼</button>
          </span>
        </div>
        <div class="lbl">-Chức năng:</div>
        <textarea data-field="func" rows="2" placeholder="Chức năng của công đoạn">${esc(p.func)}</textarea>
        <div class="lbl">-Yêu cầu:</div>
        ${reqLines}
        <button class="mini-btn" data-proc="${p.id}" data-action="add-req" style="margin-top:6px">＋ Thêm yêu cầu</button>
        <button class="mini-btn danger" data-proc="${p.id}" data-action="del-proc" style="margin-top:6px">🗑 Xóa công đoạn</button>
      </div>`;
  }

  // Nút AI hỗ trợ phân tích cho 1 ô
  const aiBtn = (field) => `<button class="ai-btn" data-action="ai" data-ai-field="${field}" title="AI hỗ trợ phân tích">✨ AI</button>`;

  // Ô cột B (Dạng hỏng hóc): liệt kê tất cả dạng hỏng hóc trong nhóm gộp.
  // Mỗi dòng có số thứ tự (theo yêu cầu ở cột A), nút 🔗 Gộp và 🔓 Tách (nếu đang gộp).
  function fmCellHTML(p, grp, reqNo) {
    const grouped = grp.length > 1;
    return grp.map(({ r }) => `
      <div class="fm-line" data-proc="${p.id}" data-req="${r.id}">
        <span class="fm-idx">${reqNo[r.id]}.</span>
        <div class="cell-edit" contenteditable="true" data-field="failureMode">${esc(r.failureMode)}</div>
        <span class="fm-tools">
          <button class="mini-btn" data-action="merge-open" title="Gộp các dạng hỏng hóc giống nhau">🔗</button>
          ${grouped ? '<button class="mini-btn danger" data-action="unmerge" title="Tách khỏi nhóm gộp">🔓</button>' : ''}
        </span>
      </div>`).join('');
  }

  function effectCellHTML(p, r) {
    const idx = findSeverityIdx(r);
    const chosen = !!r.effectStdText;
    return `<td class="auto" rowspan="@RS@" data-proc="${p.id}" data-req="${r.id}">
      <div class="effect-cell">
        <textarea data-field="effectAnalysis" rows="2" placeholder="① Tự phân tích ảnh hưởng…">${esc(r.effectAnalysis)}</textarea>
        ${aiBtn('effectAnalysis')}
        <div class="effect-std" data-proc="${p.id}" data-req="${r.id}">
          <div class="std-pick"${chosen ? ' style="display:none"' : ''}>
            <div class="effect-arrow">=&gt; (② chọn theo tiêu chuẩn đánh giá S)</div>
            <select data-field="effectStd">${severityOptions(idx)}</select>
          </div>
          <div class="std-show"${chosen ? '' : ' style="display:none"'}>
            <span class="std-arrow">=&gt;</span>
            <span class="std-text" id="text-${r.id}">${esc(r.effectStdText || '')}</span>
            <button class="mini-btn edit-eff" data-action="edit-effect" title="Đổi lựa chọn">✎ đổi</button>
          </div>
        </div>
      </div></td>`;
  }

  function detectCellHTML(p, r, c) {
    const isVis = isVisualDetect(r, c);
    const hasSC = !!norm(r.classification); // có đặc tính đặc thù (S/A…) → hiện ô bổ sung
    return `<td data-proc="${p.id}" data-req="${r.id}" data-cause="${c.id}">
      <div class="detect-cell">
        <div class="detect-label">① Phát hiện ra nguyên nhân (tự phân tích):</div>
        <textarea data-field="detectCause" rows="2" placeholder="…">${esc(c.detectCause)}</textarea>
        ${aiBtn('detectCause')}
        <span class="visual-auto" id="vis-${c.id}"${isVis ? '' : ' hidden'}
              title="Tự nhận từ nội dung: phát hiện bằng giác quan (tay/mắt/đếm…) → khi D≥7 không bắt buộc đề xuất biện pháp">🖐️ Giác quan (tự nhận)</span>
        <div class="detect-label">② Phát hiện ra dạng hỏng hóc (tự động từ CP):</div>
        <div class="detect-auto" contenteditable="true" data-field="detectFailureAuto"
             data-proc="${p.id}" data-req="${r.id}">${esc(r.detectFailureAuto)}</div>
        <div class="detect-label detect-extra-label" id="dxl-${c.id}"${hasSC ? '' : ' hidden'}>③ Bổ sung cho đặc tính đặc thù (tự điền):</div>
        <textarea class="detect-extra" id="dx-${c.id}" data-field="detectExtra" rows="2"${hasSC ? '' : ' hidden'} placeholder="Nội dung phát hiện bổ sung…">${esc(c.detectExtra)}</textarea>
      </div></td>`;
  }

  function txtTD(p, r, c, field, ph, ai) {
    return `<td data-proc="${p.id}" data-req="${r.id}" data-cause="${c.id}">
      <div class="cell-edit" contenteditable="true" data-field="${field}" data-ph="${ph || ''}">${esc(c[field])}</div>${ai ? aiBtn(field) : ''}</td>`;
  }

  function numTD(p, r, c, field, idAttr) {
    const id = idAttr ? ` id="${idAttr}"` : '';
    return `<td class="num" data-proc="${p.id}" data-req="${r.id}" data-cause="${c.id}">
      <input${id} class="num-inp" type="number" min="1" max="10" data-field="${field}" value="${esc(c[field])}" /></td>`;
  }

  function render() {
    if (aiPop) closeAIPop();
    if (mergePop) closeMergePop();
    const tbody = $('#fmea tbody');
    const empty = $('#emptyState');
    if (!state.processes.length) {
      tbody.innerHTML = '';
      empty.hidden = false;
      $('#fmea').style.display = 'none';
      $('#btnExport').disabled = true;
      return;
    }
    empty.hidden = true;
    $('#fmea').style.display = '';
    $('#btnExport').disabled = false;

    const rows = [];
    for (const p of state.processes) {
      const groups = reqGroups(p.reqs);
      const reqNo = reqNumbers(p.reqs);
      const totalRows = groups.reduce((n, g) => n + (g[0].r.causes.length || 1), 0) || 1;
      let firstProcRow = true;

      groups.forEach((grp) => {
        const r = grp[0].r;            // đại diện nhóm: cung cấp các cột C–S
        const rs = r.causes.length || 1;
        r.causes.forEach((c, ci) => {
          let tr = `<tr class="${firstProcRow ? 'proc-sep' : ''}">`;

          // Cột A — chỉ ở hàng đầu của công đoạn
          if (firstProcRow) {
            tr += `<td rowspan="${totalRows}">${procCellHTML(p, reqNo)}</td>`;
          }
          // B,C,D,E — chỉ ở hàng đầu của yêu cầu
          if (ci === 0) {
            tr += `<td class="auto" rowspan="${rs}">${fmCellHTML(p, grp, reqNo)}</td>`;
            tr += effectCellHTML(p, r).replace('@RS@', rs);
            tr += `<td class="num" rowspan="${rs}"><div class="score-box" id="sev-${r.id}">${esc(r.severity)}</div></td>`;
            tr += `<td rowspan="${rs}" data-proc="${p.id}" data-req="${r.id}">
                     <div class="cell-edit" contenteditable="true" data-field="classification">${esc(r.classification)}</div></td>`;
          }

          // F — nguyên nhân: 4M + nội dung + nút thêm/xóa
          const mOpts = ['<option value="">— 4M —</option>']
            .concat(FOUR_M.map((m) => `<option${c.category === m ? ' selected' : ''}>${m}</option>`)).join('');
          tr += `<td data-proc="${p.id}" data-req="${r.id}" data-cause="${c.id}">
                   <select class="four-m" data-field="category">${mOpts}</select>
                   <div class="cell-edit" contenteditable="true" data-field="cause" data-ph="Nguyên nhân ${ci + 1}">${esc(c.cause)}</div>
                   ${aiBtn('cause')}
                   <div class="cause-toolbar">
                     <button class="mini-btn" data-action="add-cause">＋ NN</button>
                     ${rs > 1 ? '<button class="mini-btn danger" data-action="del-cause">✕ NN</button>' : ''}
                   </div></td>`;
          // G phản ánh lỗi quá khứ
          tr += txtTD(p, r, c, 'pastTrouble', 'Lỗi quá khứ');
          // H tần suất O
          tr += numTD(p, r, c, 'occurrence');
          // I dự phòng
          tr += txtTD(p, r, c, 'prevention', 'Quản lý dự phòng (tự phân tích)', true);
          // J phát hiện (2 ý)
          tr += detectCellHTML(p, r, c);
          // K phát hiện D
          tr += numTD(p, r, c, 'detection');
          // L RPN
          tr += `<td class="num"><div class="rpn-box" id="rpn-${c.id}">${esc(rpnOf(r, c))}</div></td>`;
          // M (Biện pháp đề xuất) — viền đỏ khi S/O/D rơi vào tiêu chuẩn cần đề xuất
          const needCls = needsAction(r, c) ? ' need-action' : '';
          tr += `<td id="act-${c.id}" class="act-cell${needCls}" data-proc="${p.id}" data-req="${r.id}" data-cause="${c.id}">
                   <div class="cell-edit" contenteditable="true" data-field="action" data-ph="Biện pháp đề xuất">${esc(c.action)}</div></td>`;
          tr += txtTD(p, r, c, 'responsible', 'Người chịu trách nhiệm');
          tr += txtTD(p, r, c, 'actionTaken', '');
          tr += numTD(p, r, c, 's2');
          tr += numTD(p, r, c, 'o2');
          tr += numTD(p, r, c, 'd2');
          tr += `<td class="num"></td>`; // RPN sau (P*Q*R) — để trống, sẽ tính khi nhập

          tr += '</tr>';
          rows.push(tr);
          firstProcRow = false;
        });
      });
    }
    tbody.innerHTML = rows.join('');
    scheduleAutosave();
  }

  // ===================== Cập nhật không re-render =====================
  // Bật/tắt viền đỏ ô Biện pháp đề xuất theo tiêu chuẩn S/O/D.
  function refreshActionFlag(r, c) {
    const cell = $(`#act-${c.id}`);
    if (cell) cell.classList.toggle('need-action', needsAction(r, c));
    const vis = $(`#vis-${c.id}`);
    if (vis) vis.hidden = !isVisualDetect(r, c);
  }
  function refreshReqScores(pid, rid) {
    const r = getReq(pid, rid);
    if (!r) return;
    const sev = $(`#sev-${rid}`);
    if (sev) sev.textContent = r.severity || '';
    r.causes.forEach((c) => {
      const box = $(`#rpn-${c.id}`);
      if (box) box.textContent = rpnOf(r, c);
      refreshActionFlag(r, c); // S đổi -> xét lại mọi nguyên nhân của yêu cầu
    });
  }
  function refreshCauseRPN(pid, rid, cid) {
    const r = getReq(pid, rid), c = getCause(pid, rid, cid);
    if (!r || !c) return;
    const box = $(`#rpn-${cid}`);
    if (box) box.textContent = rpnOf(r, c);
    refreshActionFlag(r, c);
  }

  // ============================ Sự kiện ============================
  function dataset(el) {
    // tìm proc/req/cause gần nhất
    let pid = '', rid = '', cid = '';
    let n = el;
    while (n && n !== document) {
      if (!cid && n.dataset && n.dataset.cause) cid = n.dataset.cause;
      if (!rid && n.dataset && n.dataset.req) rid = n.dataset.req;
      if (!pid && n.dataset && n.dataset.proc) pid = n.dataset.proc;
      n = n.parentNode;
    }
    return { pid, rid, cid };
  }

  function onInput(e) {
    const el = e.target;
    const field = el.dataset && el.dataset.field;
    if (!field) return;
    const { pid, rid, cid } = dataset(el);
    const val = (el.tagName === 'TEXTAREA' || el.tagName === 'INPUT') ? el.value : el.textContent;

    if (field === 'no' || field === 'name' || field === 'func') {
      const p = getProc(pid); if (p) p[field] = val; return;
    }
    if (field === 'reqText' || field === 'failureMode' || field === 'classification'
        || field === 'effectAnalysis' || field === 'detectFailureAuto') {
      const r = getReq(pid, rid); if (r) r[field] = val;
      // reqText/failureMode là riêng từng yêu cầu; các cột còn lại là CHUNG -> đồng bộ nhóm
      if (r && r.mergeId && field !== 'reqText' && field !== 'failureMode') syncMergeGroup(getProc(pid), r);
      // Đổi ② Phát hiện dạng hỏng hóc -> xét lại giác quan & viền đỏ cho mọi nguyên nhân
      if (r && field === 'detectFailureAuto') r.causes.forEach((c) => refreshActionFlag(r, c));
      // Đổi Phân loại (đặc tính đặc thù S/A) -> bật/tắt ô bổ sung ③ ở cột Phát hiện ra
      if (r && field === 'classification') {
        const show = !!norm(val);
        r.causes.forEach((c) => {
          const box = $(`#dx-${c.id}`), lbl = $(`#dxl-${c.id}`);
          if (box) box.hidden = !show;
          if (lbl) lbl.hidden = !show;
        });
      }
      scheduleAutosave(); return;
    }
    // cấp nguyên nhân (cột CHUNG khi gộp)
    const c = getCause(pid, rid, cid); if (!c) return;
    c[field] = val;
    if (field === 'occurrence' || field === 'detection') refreshCauseRPN(pid, rid, cid);
    // Đổi ① Phát hiện nguyên nhân -> xét lại giác quan & viền đỏ
    if (field === 'detectCause') { const r = getReq(pid, rid); if (r) refreshActionFlag(r, c); }
    const rr = getReq(pid, rid); if (rr && rr.mergeId) syncMergeGroup(getProc(pid), rr);
    scheduleAutosave();
  }

  function onChange(e) {
    const el = e.target;
    const field = el.dataset && el.dataset.field;
    if (field === 'category') {
      const { pid, rid, cid } = dataset(el);
      const c = getCause(pid, rid, cid);
      if (c) { c.category = el.value; const r = getReq(pid, rid); if (r && r.mergeId) syncMergeGroup(getProc(pid), r); scheduleAutosave(); }
      return;
    }
    if (field === 'effectStd') {
      const { pid, rid } = dataset(el);
      const r = getReq(pid, rid); if (!r) return;
      const idx = el.value === '' ? -1 : +el.value;
      const container = el.closest('.effect-std');
      if (idx < 0) {
        r.effectStdText = ''; r.effectScope = ''; r.severity = '';
        // không chọn gì -> vẫn để picker hiển thị
      } else {
        const row = window.SEVERITY_TABLE[idx];
        r.effectStdText = row.text; r.effectScope = row.scope; r.severity = row.rank;
        const txt = $(`#text-${rid}`); if (txt) txt.textContent = r.effectStdText;
        if (container) {
          container.querySelector('.std-pick').style.display = 'none';
          container.querySelector('.std-show').style.display = '';
        }
      }
      refreshReqScores(pid, rid);
      if (r.mergeId) syncMergeGroup(getProc(pid), r);
      scheduleAutosave();
    }
  }

  function onClick(e) {
    const btn = e.target.closest('[data-action]');
    if (!btn) return;
    const action = btn.dataset.action;
    if (action === 'ai') { onAIClick(btn); return; }
    const { pid, rid, cid } = dataset(btn);

    if (action === 'edit-effect') {
      const container = btn.closest('.effect-std');
      if (container) {
        container.querySelector('.std-show').style.display = 'none';
        container.querySelector('.std-pick').style.display = '';
        const sel = container.querySelector('select');
        if (sel) sel.focus();
      }
      return;
    }
    if (action === 'merge-open') { openMergePop(btn, pid, rid); return; }
    if (action === 'unmerge') { unmergeReq(pid, rid); return; }
    if (action === 'add-cause') {
      const r = getReq(pid, rid); if (r) { r.causes.push(newCause()); render(); }
    } else if (action === 'del-cause') {
      const r = getReq(pid, rid);
      if (r && r.causes.length > 1) { r.causes = r.causes.filter((c) => c.id !== cid); render(); }
    } else if (action === 'add-req') {
      const p = getProc(pid);
      if (p) { p.reqs.push(reqFromItem({ name: '', requirement: '' })); render(); }
    } else if (action === 'del-req') {
      const p = getProc(pid);
      const target = p && p.reqs.find((r) => r.id === rid);
      // Xóa cả cụm tách dạng hỏng (cùng splitId) — vì cột Yêu cầu chỉ hiện 1 dòng.
      const toDel = target && target.splitId
        ? p.reqs.filter((r) => r.splitId === target.splitId).map((r) => r.id)
        : [rid];
      if (p && p.reqs.length > toDel.length) {
        p.reqs = p.reqs.filter((r) => toDel.indexOf(r.id) < 0); render();
      } else if (p) { alert('Mỗi công đoạn cần ít nhất 1 yêu cầu.'); }
    } else if (action === 'del-proc') {
      if (confirm('Xóa công đoạn này?')) {
        state.processes = state.processes.filter((p) => p.id !== pid); render();
      }
    } else if (action === 'move-up') {
      moveProc(pid, -1);
    } else if (action === 'move-down') {
      moveProc(pid, +1);
    }
  }

  // ============================ Tải CP =============================
  function onFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    $('#fileName').textContent = file.name;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        workbook = XLSX.read(new Uint8Array(ev.target.result), { type: 'array' });
      } catch (err) {
        alert('Không đọc được file: ' + err.message); return;
      }
      const sheets = window.CPParser.listSheets(workbook);
      $('#sheetInfo').textContent = `Đã đọc ${sheets.length} sheet — bấm để nạp tất cả công đoạn.`;
      $('#sheetGroup').hidden = false;
      $('#btnAddProc').hidden = false;
      $('#btnClear').hidden = false;
    };
    reader.readAsArrayBuffer(file);
  }

  // Đọc TẤT CẢ sheet -> danh sách công đoạn mới (fresh từ CP)
  // Các sheet có cùng "Tên công đoạn" được gộp thành 1 công đoạn,
  // số thứ tự hạng mục tiếp nối liên tục qua các sheet.
  function parseAllProcs() {
    const sheets = window.CPParser.listSheets(workbook);
    // Dùng Map để giữ thứ tự sheet; key = tên công đoạn đã chuẩn hóa
    const groups = new Map(); // normKey → { name: displayName, items: [] }
    const skipped = [];
    sheets.forEach((sheet) => {
      let res;
      try { res = window.CPParser.parseSheet(workbook, sheet); }
      catch (err) { skipped.push(sheet); return; }
      if (res.error || !res.items.length) { skipped.push(sheet); return; }
      const displayName = res.processName || sheet;
      const key = normKey(displayName);
      if (!groups.has(key)) groups.set(key, { name: displayName, items: [] });
      groups.get(key).items.push(...res.items);
    });
    const procs = [];
    groups.forEach(({ name, items }) => {
      // Đánh số thứ tự liên tục qua tất cả sheet của cùng công đoạn
      items.forEach((item, idx) => { item.no = idx + 1; });
      procs.push({
        id: uid('p'),
        no: String(procs.length + 1),
        name,
        func: '',
        reqs: items.flatMap(reqsFromItem),
      });
    });
    return { procs, skipped };
  }

  // Khóa so khớp tên (công đoạn / hạng mục) — bỏ dấu cách thừa, không phân biệt hoa thường
  const normKey = (s) => (s == null ? '' : String(s)).replace(/\s+/g, ' ').trim().toLowerCase();
  // Tên hạng mục lấy từ phần trước dấu ":" của yêu cầu (ổn định dù kích thước đổi)
  function reqNameKey(r) {
    let n = (r.reqText || '').split(':')[0].trim();
    if (!n) n = (r.failureMode || '').replace(/\s*không đạt\s*$/i, '').trim();
    return normKey(n);
  }

  // Ghép CP mới theo Model base: cấu trúc lấy theo CP mới; hạng mục trùng tên thì
  // GIỮ phân tích của base, chỉ cập nhật kích thước (reqText) theo CP mới; hạng mục
  // mới thì giữ nguyên bản tự sinh từ CP; hạng mục base không có trong CP thì bỏ.
  function mergeWithBase(baseProcs, newProcs) {
    const baseByName = {};
    baseProcs.forEach((p) => { baseByName[normKey(p.name)] = p; });
    return newProcs.map((np, i) => {
      const bp = baseByName[normKey(np.name)];
      let reqs;
      if (bp) {
        const baseByItem = {};
        bp.reqs.forEach((r) => { baseByItem[reqNameKey(r)] = r; });
        reqs = np.reqs.map((nr) => {
          const br = baseByItem[reqNameKey(nr)];
          if (!br) return nr; // hạng mục mới -> giữ bản tự sinh từ CP
          // trùng tên -> giữ phân tích base, chỉ cập nhật kích thước theo CP mới
          return {
            id: uid('r'),
            reqText: nr.reqText,
            failureMode: br.failureMode,
            effectAnalysis: br.effectAnalysis,
            effectStdText: br.effectStdText,
            effectScope: br.effectScope,
            severity: br.severity,
            classification: br.classification,
            detectFailureAuto: br.detectFailureAuto,
            causes: (br.causes && br.causes.length ? br.causes : [newCause()])
              .map((c) => Object.assign({}, c, { id: uid('c') })),
          };
        });
      } else {
        reqs = np.reqs; // công đoạn mới hoàn toàn
      }
      return { id: uid('p'), no: String(i + 1), name: np.name, func: bp ? bp.func : np.func, reqs };
    });
  }

  function onLoadProc() {
    if (!workbook) return;
    const { procs, skipped } = parseAllProcs();
    if (!procs.length) {
      alert('Không tìm thấy công đoạn nào có hạng mục chất lượng trong file.');
      return;
    }
    // Đang mở sẵn dữ liệu (Model base) -> ghép theo base; nếu muốn nạp mới hoàn
    // toàn thì bấm "Xóa hết" trước rồi nạp lại.
    if (state.processes.length) {
      const ok = confirm(
        'Đang có dữ liệu (Model base) đang mở.\n\n' +
        'OK = GHÉP CP mới theo Model base:\n' +
        '  • Giữ nguyên phân tích của base, chỉ cập nhật kích thước theo CP mới.\n' +
        '  • Thêm hạng mục mới mà base chưa có (tự sinh từ CP).\n' +
        '  • Bỏ hạng mục base không có trong CP mới.\n\n' +
        'Cancel = Hủy (muốn nạp mới hoàn toàn thì bấm "Xóa hết" trước).'
      );
      if (!ok) return;
      state.processes = mergeWithBase(state.processes, procs);
      render();
      flash('Đã ghép CP theo Model base (' + state.processes.length + ' công đoạn). Hãy đổi Model sang model mới rồi bấm Lưu để tạo bản mới.');
      document.querySelector('.sheet-wrap').scrollIntoView({ behavior: 'smooth' });
      return;
    }
    state.processes = procs;
    render();
    flash(`Đã nạp ${procs.length} công đoạn` + (skipped.length ? ` (bỏ qua ${skipped.length} sheet không phải công đoạn)` : '') + '.');
    document.querySelector('.sheet-wrap').scrollIntoView({ behavior: 'smooth' });
  }

  // Đổi thứ tự công đoạn (lên/xuống) + đánh lại STT tự động
  function moveProc(pid, dir) {
    const i = state.processes.findIndex((p) => p.id === pid);
    const j = i + dir;
    if (i < 0 || j < 0 || j >= state.processes.length) return;
    const arr = state.processes;
    [arr[i], arr[j]] = [arr[j], arr[i]];
    arr.forEach((p, k) => { if (/^\d+$/.test((p.no || '').trim()) || !p.no) p.no = String(k + 1); });
    render();
  }

  function onAddProc() {
    state.processes.push({
      id: uid('p'), no: '', name: '', func: '',
      reqs: [reqFromItem({ name: '', requirement: '' })],
    });
    render();
  }

  function onClear() {
    if (state.processes.length && !confirm('Xóa toàn bộ dữ liệu P-FMEA?')) return;
    state.processes = []; render();
  }

  // ====================== Lưu / Mở dự án (localStorage) ======================
  function snapshot() {
    return JSON.stringify({ meta: state.meta, processes: state.processes });
  }
  // Phân tích reqText dạng "Tên: spec(tol)" → { name, spec, tol }
  // Dùng trong migration để tái sinh dạng hỏng hóc từ dữ liệu cũ.
  function parseReqText(reqText) {
    const s = norm(reqText || '');
    const colonIdx = s.indexOf(': ');
    if (colonIdx < 0) return { name: s, spec: '', tol: '' };
    const name = s.slice(0, colonIdx);
    const rest = s.slice(colonIdx + 2).trim();
    if (!rest) return { name, spec: '', tol: '' };
    const parenMatch = rest.match(/^(.+?)\(([^)]+)\)\s*$/);
    if (parenMatch) return { name, spec: parenMatch[1].trim(), tol: parenMatch[2].trim() };
    return { name, spec: rest, tol: '' };
  }

  // Bổ sung field mới VÀ tái sinh dạng hỏng hóc cho dữ liệu cũ.
  function migrateState(obj) {
    // Đẩy UID vượt qua mọi id đã có trong obj để tránh trùng lúc tạo ID mới.
    let mx = UID;
    JSON.stringify(obj.processes || []).replace(/[a-z](\d+)/g, (_, n) => { mx = Math.max(mx, +n); return _; });
    UID = mx + 1;

    (obj.processes || []).forEach((p) => {
      const newReqs = [];
      (p.reqs || []).forEach((r) => {
        // 1. Bổ sung field còn thiếu
        if (r.splitId       === undefined) r.splitId       = '';
        if (r.mergeId       === undefined) r.mergeId       = '';
        if (r.classification=== undefined) r.classification = '';
        if (r.detectFailureAuto === undefined) r.detectFailureAuto = '';
        (r.causes || []).forEach((c) => {
          if (c.detectExtra === undefined) c.detectExtra = '';
        });

        // 2. Tái sinh dạng hỏng hóc nếu req chưa được tách (dạng cũ).
        //    Nhận biết dạng cũ: không có splitId VÀ failureMode không chứa
        //    "lớn hơn tiêu chuẩn" / "nhỏ hơn tiêu chuẩn".
        if (!r.splitId && !/lớn hơn tiêu chuẩn|nhỏ hơn tiêu chuẩn/.test(r.failureMode)) {
          const parsed = parseReqText(r.reqText);
          const modes  = failureModesFor(parsed);
          if (modes.length === 2) {
            // Tách thành 2 req độc lập cùng splitId
            const splitId = uid('split');
            newReqs.push(Object.assign({}, r, { splitId, failureMode: modes[0] }));
            newReqs.push(Object.assign({}, r, {
              id: uid('r'), splitId, failureMode: modes[1],
              causes: (r.causes || []).map((c) => Object.assign({}, c, { id: uid('c') })),
            }));
            return; // không push r gốc
          }
          // 1 mode: cập nhật failureMode sang định dạng mới (có spec)
          r.failureMode = modes[0] || r.failureMode;
        }
        newReqs.push(r);
      });
      p.reqs = newReqs;

      // Cập nhật nội dung "Phát hiện ra dạng hỏng hóc" của nhóm đã gộp sang định
      // dạng mới "Kiểm tra các <tên chung> bằng…" nếu còn ở dạng nối cũ "X và Y".
      const seenM = new Set();
      p.reqs.forEach((r) => {
        if (!r.mergeId || seenM.has(r.mergeId)) return;
        seenM.add(r.mergeId);
        const members = p.reqs.filter((x) => x.mergeId === r.mergeId);
        if (members.length < 2) return;
        const auto = r.detectFailureAuto || '';
        if (/^Kiểm tra .+ và .+ bằng /.test(auto)) {
          const rebuilt = buildGroupDetect(members, parseDetect(auto).suffix);
          members.forEach((m) => { m.detectFailureAuto = rebuilt; });
        }
      });
    });
    return obj;
  }

  function applySnapshot(obj) {
    obj = migrateState(obj);
    state.meta = Object.assign({ dept: '', product: '', line: '', model: '' }, obj.meta || {});
    state.processes = obj.processes || [];
    reindexUID();
    writeMetaInputs();
    render();
  }

  let autosaveTimer = null;
  function scheduleAutosave() {
    clearTimeout(autosaveTimer);
    autosaveTimer = setTimeout(() => {
      // Tự lưu TOÀN BỘ phiên làm việc (meta + nội dung). KHÔNG bao giờ xóa.
      try { localStorage.setItem(LS_AUTOSAVE, snapshot()); } catch (e) { /* hết dung lượng */ }
    }, 600);
  }

  function readProjects() {
    try { return JSON.parse(localStorage.getItem(LS_PROJECTS) || '{}'); } catch (e) { return {}; }
  }
  function writeProjects(map) {
    localStorage.setItem(LS_PROJECTS, JSON.stringify(map));
  }
  function projectKey(meta) {
    return [meta.model, meta.line, meta.product, meta.dept].map((s) => (s || '').trim()).join(' | ').replace(/(\s\|\s)+$/, '');
  }
  // Khóa định danh dự án hiện tại — CHỈ hợp lệ khi đã có Model (model = tên dự án).
  function currentKey() { return state.meta.model.trim() ? projectKey(state.meta) : ''; }

  // Đẩy UID vượt qua mọi id số sẵn có (tránh trùng khi nạp dữ liệu đã lưu)
  function reindexUID() {
    let mx = 0;
    JSON.stringify(state.processes).replace(/[a-z](\d+)/g, (_, n) => { mx = Math.max(mx, +n); return _; });
    UID = mx + 1;
  }

  function onSave() {
    readMetaInputs();
    if (!state.meta.model.trim()) {
      alert('Hãy nhập Model trước khi lưu (dùng làm tên dự án).');
      $('#mModel').focus(); return;
    }
    if (!state.processes.length) { alert('Chưa có dữ liệu để lưu.'); return; }
    const map = readProjects();
    const key = projectKey(state.meta);
    map[key] = { meta: state.meta, processes: state.processes, savedAt: new Date().toISOString() };
    writeProjects(map);
    cloudPushProject(key, map[key]);   // đẩy lên dữ liệu chung
    refreshProjectUI();
    $('#projSelect').value = key;
    flash('Đã lưu: ' + key + (sb ? ' (đã đồng bộ chung)' : ''));
  }

  function onPickProject() {
    const key = $('#projSelect').value;
    if (!key) return;
    const map = readProjects();
    const proj = map[key];
    if (!proj) return;
    if (state.processes.length && !confirm('Mở dự án "' + key + '"? Dữ liệu hiện tại chưa lưu sẽ bị thay thế.')) {
      $('#projSelect').value = projectKey(state.meta); return;
    }
    applySnapshot(proj);
    flash('Đã mở: ' + key);
  }

  function onDeleteProject() {
    const key = $('#projSelect').value;
    if (!key) { alert('Chọn một dự án để xóa.'); return; }
    if (!confirm('Xóa dự án đã lưu "' + key + '"?')) return;
    const map = readProjects();
    delete map[key];
    writeProjects(map);
    refreshProjectUI();
    flash('Đã xóa: ' + key);
  }

  function onExportJson() {
    const data = { projects: readProjects(), current: JSON.parse(snapshot()), exportedAt: new Date().toISOString() };
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = `P-FMEA_backup_${new Date().toISOString().slice(0, 10)}.json`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function onImportJson(e) {
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const data = JSON.parse(ev.target.result);
        const map = readProjects();
        const incoming = data.projects || {};
        Object.assign(map, incoming);
        writeProjects(map);
        // Đẩy các dự án vừa nạp lên DB chung (để chuyển model base cũ lên web cho cả nhóm)
        Object.keys(incoming).forEach((k) => cloudPushProject(k, incoming[k]));
        // Khôi phục luôn NỘI DUNG đang làm dở (current) nếu có — để mang dữ liệu
        // sang file mới mà không mất gì.
        const cur = data.current;
        if (cur && cur.processes && cur.processes.length &&
            (!state.processes.length ||
             confirm('Nạp luôn nội dung đang làm dở trong file sao lưu (thay nội dung hiện tại)?'))) {
          applySnapshot(cur);
        }
        refreshProjectUI();
        flash('Đã nạp file sao lưu (' + Object.keys(data.projects || {}).length + ' dự án).');
      } catch (err) { alert('File sao lưu không hợp lệ: ' + err.message); }
    };
    reader.readAsText(file);
    e.target.value = '';
  }

  function refreshProjectUI() {
    const map = readProjects();
    const keys = Object.keys(map).sort();
    const sel = $('#projSelect');
    const cur = sel.value;
    sel.innerHTML = '<option value="">— Chọn —</option>' +
      keys.map((k) => `<option value="${esc(k)}">${esc(k)}</option>`).join('');
    if (keys.includes(cur)) sel.value = cur;
    fillModelDatalist();
  }

  // ----- Dropdown phân cấp Bộ phận > Sản phẩm > Dây chuyền (CHỈ theo Material) -----
  const TREE = () => window.PFMEA_MATERIAL || {};
  // Đổ option cho 1 <select>. CHỈ liệt kê đúng dữ liệu Material (không thêm giá trị lạ).
  // Nếu 'keep' nằm trong danh sách thì chọn sẵn; nếu không thì để trống.
  function fillSelect(id, items, keep, placeholder) {
    const sel = $(id);
    const has = keep && items.includes(keep);
    sel.innerHTML = `<option value="">${placeholder}</option>` +
      items.map((v) => `<option value="${esc(v)}"${v === keep ? ' selected' : ''}>${esc(v)}</option>`).join('');
    sel.value = has ? keep : '';
  }
  function fillDept(keep) { fillSelect('#mDept', Object.keys(TREE()), keep, '— Chọn bộ phận —'); }
  function fillProduct(dept, keep) {
    const prods = TREE()[dept] ? Object.keys(TREE()[dept]) : [];
    fillSelect('#mProduct', prods, keep, '— Chọn sản phẩm —');
  }
  function fillLine(dept, product, keep) {
    const lines = (TREE()[dept] && TREE()[dept][product]) ? TREE()[dept][product] : [];
    fillSelect('#mLine', lines, keep, lines.length ? '— Chọn dây chuyền —' : '— (không có) —');
  }
  // Model: gõ tự do, gợi ý từ dự án đã lưu
  function fillModelDatalist() {
    const map = readProjects();
    const models = new Set();
    Object.keys(map).forEach((k) => { const m = (map[k].meta || {}).model; if (m) models.add(m); });
    $('#dlModel').innerHTML = [...models].map((v) => `<option value="${esc(v)}">`).join('');
  }

  function readMetaInputs() {
    state.meta.dept = $('#mDept').value.trim();
    state.meta.product = $('#mProduct').value.trim();
    state.meta.line = $('#mLine').value.trim();
    state.meta.model = $('#mModel').value.trim();
  }
  function writeMetaInputs() {
    fillDept(state.meta.dept || '');
    fillProduct(state.meta.dept || '', state.meta.product || '');
    fillLine(state.meta.dept || '', state.meta.product || '', state.meta.line || '');
    $('#mModel').value = state.meta.model || '';
    fillModelDatalist();
  }

  let flashTimer = null;
  function flash(msg) {
    let el = $('#flash');
    if (!el) {
      el = document.createElement('div'); el.id = 'flash'; el.className = 'flash';
      document.body.appendChild(el);
    }
    el.textContent = msg; el.classList.add('show');
    clearTimeout(flashTimer);
    flashTimer = setTimeout(() => el.classList.remove('show'), 2200);
  }

  // ============================ AI (Gemini) ========================
  function readContexts() {
    try { return JSON.parse(localStorage.getItem(LS_CONTEXT) || '{}'); } catch (e) { return {}; }
  }
  function ctxKeyName() { return (state.meta.dept || '').trim() || '(chung)'; }
  function getContext() { return readContexts()[ctxKeyName()] || {}; }
  function contextText() {
    const c = getContext();
    const parts = [];
    if (c.dept) parts.push('- Bộ phận sản xuất: ' + c.dept);
    if (c.product) parts.push('- Sản phẩm/loại giảm xóc: ' + c.product);
    if (c.machine) parts.push('- Máy móc thiết bị: ' + c.machine);
    if (c.process) parts.push('- Quy trình gia công: ' + c.process);
    if (c.other) parts.push('- Thông tin khác: ' + c.other);
    return parts.join('\n');
  }
  const getGeminiKey = () => (localStorage.getItem(LS_GEMINI_KEY) || '').trim();
  const getGeminiModel = () => (localStorage.getItem(LS_GEMINI_MODEL) || '').trim();

  // ----- Modal bối cảnh -----
  function openAIModal() {
    $('#ctxDeptLabel').textContent = 'Bối cảnh cho bộ phận: ' + ctxKeyName() + ' (đổi Bộ phận để nhập cho bộ phận khác)';
    cloudPullContexts(); cloudPullPhrases();   // làm tươi dữ liệu chung khi mở
    $('#aiKey').value = getGeminiKey();
    $('#aiModel').value = getGeminiModel();
    const c = getContext();
    $('#ctxDept').value = c.dept || '';
    $('#ctxProduct').value = c.product || '';
    $('#ctxMachine').value = c.machine || '';
    $('#ctxProcess').value = c.process || '';
    $('#ctxOther').value = c.other || '';
    $('#aiModal').hidden = false;
  }
  function closeAIModal() { $('#aiModal').hidden = true; }
  function saveAIContext() {
    localStorage.setItem(LS_GEMINI_KEY, $('#aiKey').value.trim());
    localStorage.setItem(LS_GEMINI_MODEL, $('#aiModel').value);
    const map = readContexts();
    map[ctxKeyName()] = {
      dept: $('#ctxDept').value.trim(), product: $('#ctxProduct').value.trim(),
      machine: $('#ctxMachine').value.trim(), process: $('#ctxProcess').value.trim(),
      other: $('#ctxOther').value.trim(),
    };
    localStorage.setItem(LS_CONTEXT, JSON.stringify(map));
    cloudPushContext(ctxKeyName(), map[ctxKeyName()]);   // đẩy lên chung
    closeAIModal();
    flash('Đã lưu bối cảnh AI cho bộ phận ' + ctxKeyName() + (sb ? ' (đã đồng bộ chung)' : '') + '.');
  }

  // ----- Gọi Gemini (1 model) -----
  async function callGeminiOnce(prompt, key, model) {
    const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent?key=' + encodeURIComponent(key);
    const res = await fetch(url, {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] }),
    });
    if (!res.ok) {
      let detail = '';
      try { const e = await res.json(); if (e.error && e.error.message) detail = e.error.message; } catch (_) { /* noop */ }
      const err = new Error('HTTP ' + res.status + (detail ? (' — ' + detail) : ''));
      err.status = res.status;
      err.quotaZero = /limit:\s*0/i.test(detail);          // model không có hạn mức free
      err.notFound = res.status === 404 || /not found|NOT_FOUND/i.test(detail);
      throw err;
    }
    const data = await res.json();
    const cand = data && data.candidates && data.candidates[0];
    const parts = cand && cand.content && cand.content.parts;
    return (parts || []).map((p) => p.text || '').join('').trim();
  }
  // Gọi Gemini: nếu người dùng chọn model thì dùng đúng model đó; nếu để "Tự động"
  // thì thử lần lượt cho tới khi 1 model có hạn mức free hoạt động.
  async function callGemini(prompt, key) {
    const chosen = getGeminiModel();
    const models = chosen ? [chosen]
      : ['gemini-2.5-flash', 'gemini-2.0-flash', 'gemini-2.0-flash-lite', 'gemini-1.5-flash'];
    let lastErr;
    for (const m of models) {
      try { return await callGeminiOnce(prompt, key, m); }
      catch (e) {
        lastErr = e;
        if (e.quotaZero || e.notFound) continue;  // model này không khả dụng -> thử model khác
        throw e;                                   // lỗi khác (rate limit thật, key sai...) -> dừng
      }
    }
    if (lastErr && lastErr.quotaZero) {
      throw new Error('Các model đều báo hết hạn mức miễn phí (limit 0). API key của bạn có thể CHƯA bật gói free — thường do key tạo từ Google Cloud có bật thanh toán, hoặc khu vực chưa hỗ trợ. Hãy tạo key MỚI tại aistudio.google.com/apikey (tài khoản Google cá nhân, không bật billing) rồi thử lại.');
    }
    throw lastErr || new Error('Không gọi được AI.');
  }

  // Mô tả ngắn cho từng cột (để AI gợi ý đúng trọng tâm)
  const AI_ASK_SHORT = {
    effectAnalysis: 'ẢNH HƯỞNG của dạng hỏng hóc (đến công đoạn sau / sản phẩm / khách hàng dùng xe)',
    cause: 'NGUYÊN NHÂN tiềm ẩn gây ra dạng hỏng hóc (theo 4M)',
    prevention: 'BIỆN PHÁP DỰ PHÒNG để ngăn nguyên nhân xảy ra',
    detectCause: 'cách PHÁT HIỆN RA NGUYÊN NHÂN (phương pháp kiểm soát/kiểm tra)',
  };
  const FIELD_LABEL = { effectAnalysis: 'Ảnh hưởng', cause: 'Nguyên nhân', prevention: 'Dự phòng', detectCause: 'Phát hiện nguyên nhân' };

  // Yêu cầu AI trả về ~5 mẫu câu ngắn gọn (mỗi câu 1 dòng)
  function buildSuggestPrompt(field, p, r, c) {
    const ctx = contextText();
    const L = [];
    L.push('Bạn là chuyên gia P-FMEA (FMEA công đoạn) trong nhà máy sản xuất GIẢM XÓC XE MÁY.');
    if (ctx) L.push('Bối cảnh nhà máy/bộ phận:\n' + ctx);
    L.push('Công đoạn: ' + (p.name || '') + (p.func ? (' — chức năng: ' + p.func) : ''));
    if (r) {
      if (r.reqText) L.push('Hạng mục yêu cầu: ' + r.reqText);
      if (r.failureMode) L.push('Dạng hỏng hóc tiềm ẩn: ' + r.failureMode);
      if (r.effectStdText) L.push('Ảnh hưởng (theo tiêu chuẩn): ' + r.effectStdText);
    }
    if (c && (field === 'prevention' || field === 'detectCause')) {
      L.push('Nguyên nhân' + (c.category ? (' (' + c.category + ')') : '') + ': ' + (c.cause || '(chưa nêu)'));
    }
    L.push('Hãy đưa ra ĐÚNG 5 mẫu câu cho phần: ' + (AI_ASK_SHORT[field] || '') + '.');
    L.push('Mỗi câu trên 1 DÒNG riêng. NGẮN GỌN, đủ ý logic kỹ thuật, KHÔNG lan man, KHÔNG đánh số, KHÔNG gạch đầu dòng, KHÔNG giải thích thêm. Tiếng Việt.');
    return L.join('\n');
  }
  function parseSuggestions(text) {
    return (text || '').split('\n')
      .map((s) => s.replace(/^\s*(?:\d+[.)]|[-*•])\s*/, '').trim())
      .filter(Boolean).slice(0, 8);
  }

  function applyAIResult(field, pid, rid, cid, text) {
    text = (text || '').trim();
    if (!text) return;
    if (field === 'effectAnalysis') {
      const r = getReq(pid, rid); if (r) r.effectAnalysis = text;
    } else {
      const c = getCause(pid, rid, cid); if (c) c[field] = text;
    }
    render();
    scheduleAutosave();
  }

  // ----- Bộ nhớ câu đã nhập: theo CỘT + BỘ PHẬN -----
  function readPhrases() { try { return JSON.parse(localStorage.getItem(LS_PHRASES) || '{}'); } catch (e) { return {}; } }
  const phraseKey = (field) => field + '|' + ctxKeyName();
  function getSavedPhrases(field) { return readPhrases()[phraseKey(field)] || []; }
  function savePhrase(field, text) {
    text = (text || '').replace(/\s+/g, ' ').trim();
    if (!text || text.length < 2) return;
    const all = readPhrases(); const k = phraseKey(field); const list = all[k] || [];
    if (list.some((p) => p.trim().toLowerCase() === text.toLowerCase())) return;
    list.unshift(text); if (list.length > 50) list.length = 50;
    all[k] = list; localStorage.setItem(LS_PHRASES, JSON.stringify(all));
    cloudPushPhrase(field, ctxKeyName(), text);   // đẩy câu mới lên chung
  }

  // ----- Popup gộp dạng hỏng hóc -----
  let mergePop = null;
  function closeMergePop() {
    if (mergePop) { mergePop.remove(); mergePop = null; document.removeEventListener('mousedown', onMergeOutside, true); }
  }
  function onMergeOutside(e) {
    if (mergePop && !mergePop.contains(e.target) && !e.target.closest('[data-action="merge-open"]')) closeMergePop();
  }
  function openMergePop(btn, pid, rid) {
    closeMergePop();
    const p = getProc(pid), r = getReq(pid, rid);
    if (!p || !r) return;
    const sig = reqSig(r);
    // Ứng viên gộp: cùng công đoạn, chữ ký GIỐNG HỆT (mọi cột trừ dạng hỏng hóc/yêu cầu),
    // và chưa nằm cùng nhóm với yêu cầu này. Khác mục kiểm tra/tần suất -> chữ ký khác -> bị loại.
    const cands = p.reqs.filter((r2) => r2.id !== rid && reqSig(r2) === sig
      && !(r.mergeId && r2.mergeId === r.mergeId));
    mergePop = document.createElement('div');
    mergePop.className = 'ai-pop merge-pop';
    let body;
    if (!cands.length) {
      body = '<div class="ai-pop-sub">Không có hạng mục nào dùng cùng <b>dụng cụ &amp; tần suất kiểm tra</b> để gộp.</div>';
    } else {
      body = '<div class="ai-pop-sub">Cùng dụng cụ &amp; tần suất kiểm tra — chọn để gộp:</div>'
        + cands.map((r2) => {
            const i = p.reqs.indexOf(r2) + 1;
            return `<label class="merge-item"><input type="checkbox" value="${r2.id}" /><span>${i}. ${esc(r2.failureMode || r2.reqText || '(trống)')}</span></label>`;
          }).join('');
    }
    mergePop.innerHTML =
      `<div class="ai-pop-head">🔗 Gộp dạng hỏng hóc<button class="ai-pop-x" title="Đóng">✕</button></div>
       <div class="ai-pop-scroll">${body}</div>`
      + (cands.length ? '<div class="ai-pop-add"><button class="btn btn-primary merge-do">Gộp đã chọn</button></div>' : '');
    document.body.appendChild(mergePop);
    positionPop(mergePop, btn);
    mergePop.querySelector('.ai-pop-x').addEventListener('click', closeMergePop);
    const doBtn = mergePop.querySelector('.merge-do');
    if (doBtn) doBtn.addEventListener('click', () => {
      const ids = Array.from(mergePop.querySelectorAll('input:checked')).map((i) => i.value);
      if (!ids.length) { closeMergePop(); return; }
      const mid = r.mergeId || uid('m');
      r.mergeId = mid;
      ids.forEach((id) => { const t = p.reqs.find((x) => x.id === id); if (t) t.mergeId = mid; });
      // Xây lại câu "Kiểm tra A và B bằng ... theo tần suất ..." cho toàn nhóm.
      const allMembers = p.reqs.filter((x) => x.mergeId === mid);
      const { suffix } = parseDetect(r.detectFailureAuto);
      const combined = buildGroupDetect(allMembers, suffix);
      allMembers.forEach((m) => { m.detectFailureAuto = combined; });
      syncMergeGroup(p, r);
      closeMergePop();
      render();
    });
    document.addEventListener('mousedown', onMergeOutside, true);
  }
  // Tách 1 yêu cầu khỏi nhóm gộp, khôi phục câu detectFailureAuto riêng.
  function unmergeReq(pid, rid) {
    const p = getProc(pid), r = getReq(pid, rid);
    if (!p || !r || !r.mergeId) return;
    const mid = r.mergeId;
    const { suffix } = parseDetect(r.detectFailureAuto);
    r.mergeId = '';
    // Khôi phục câu riêng: "Kiểm tra [tên hạng mục] bằng ..."
    if (suffix) r.detectFailureAuto = 'Kiểm tra ' + baseMeasureName(itemNameFrom(r)) + suffix;
    const rest = p.reqs.filter((x) => x.mergeId === mid);
    if (rest.length === 1) {
      // Chỉ còn 1 thành viên → giải tán nhóm, khôi phục câu riêng của thành viên đó
      const last = rest[0]; last.mergeId = '';
      if (suffix) last.detectFailureAuto = 'Kiểm tra ' + baseMeasureName(itemNameFrom(last)) + suffix;
    } else if (rest.length > 0 && suffix) {
      // Xây lại câu gộp cho phần còn lại
      const combined = buildGroupDetect(rest, suffix);
      rest.forEach((m) => { m.detectFailureAuto = combined; });
    }
    render();
  }

  // ----- Popup gợi ý -----
  let aiPop = null;
  function closeAIPop() {
    if (aiPop) { aiPop.remove(); aiPop = null; document.removeEventListener('mousedown', onPopOutside, true); }
  }
  function onPopOutside(e) {
    if (aiPop && !aiPop.contains(e.target) && !e.target.closest('.ai-btn')) closeAIPop();
  }
  function positionPop(pop, btn) {
    const r = btn.getBoundingClientRect();
    let left = r.left, top = r.bottom + 4;
    const pw = 340;
    if (left + pw > window.innerWidth - 8) left = Math.max(8, window.innerWidth - pw - 8);
    if (top + 260 > window.innerHeight && r.top - 4 > 260) top = r.top - 264; // lật lên nếu thiếu chỗ
    pop.style.left = left + 'px'; pop.style.top = top + 'px';
  }
  function chooseSuggestion(field, pid, rid, cid, text) {
    closeAIPop();
    applyAIResult(field, pid, rid, cid, text); // thay thế nội dung ô
  }

  async function onAIClick(btn) {
    const field = btn.dataset.aiField;
    const { pid, rid, cid } = dataset(btn);
    const p = getProc(pid), r = getReq(pid, rid);
    if (!p || !r) return;
    closeAIPop();

    aiPop = document.createElement('div');
    aiPop.className = 'ai-pop';
    aiPop.innerHTML =
      `<div class="ai-pop-head">Gợi ý: <b>${esc(FIELD_LABEL[field] || field)}</b><button class="ai-pop-x" title="Đóng">✕</button></div>
       <div class="ai-pop-scroll">
         <div class="ai-pop-saved"></div>
         <div class="ai-pop-ai"><div class="ai-pop-sub">⏳ Đang lấy gợi ý AI…</div></div>
       </div>
       <div class="ai-pop-add"><input type="text" placeholder="Tự nhập câu của bạn…" /><button class="btn btn-primary">Dùng</button></div>`;
    document.body.appendChild(aiPop);
    positionPop(aiPop, btn);

    // Câu đã lưu (đúng cột + bộ phận)
    const saved = getSavedPhrases(field);
    if (saved.length) {
      aiPop.querySelector('.ai-pop-saved').innerHTML =
        '<div class="ai-pop-sub">Câu đã lưu (cột này):</div>' +
        saved.map((s) => `<div class="ai-pop-item" data-text="${esc(s)}">${esc(s)}</div>`).join('');
    }

    // Sự kiện chọn item
    aiPop.querySelector('.ai-pop-x').addEventListener('click', closeAIPop);
    aiPop.querySelector('.ai-pop-scroll').addEventListener('click', (e) => {
      const it = e.target.closest('.ai-pop-item');
      if (it) chooseSuggestion(field, pid, rid, cid, it.getAttribute('data-text'));
    });
    const inp = aiPop.querySelector('.ai-pop-add input');
    const doAdd = () => {
      const v = inp.value.trim(); if (!v) return;
      savePhrase(field, v);                          // câu tự nhập -> lưu vào bộ nhớ cột
      chooseSuggestion(field, pid, rid, cid, v);
    };
    aiPop.querySelector('.ai-pop-add button').addEventListener('click', doAdd);
    inp.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); doAdd(); } });
    setTimeout(() => document.addEventListener('mousedown', onPopOutside, true), 0);

    // Gợi ý AI
    const aiBox = aiPop.querySelector('.ai-pop-ai');
    const key = getGeminiKey();
    if (!key) {
      aiBox.innerHTML = '<div class="ai-pop-sub">Chưa có API key — bấm "🧠 Bối cảnh AI" để nhập (vẫn dùng được câu đã lưu / tự nhập).</div>';
      return;
    }
    try {
      const c = cid ? getCause(pid, rid, cid) : null;
      const text = await callGemini(buildSuggestPrompt(field, p, r, c), key);
      if (!aiPop) return;
      const sugs = parseSuggestions(text);
      aiBox.innerHTML = '<div class="ai-pop-sub">Gợi ý AI:</div>' +
        (sugs.length ? sugs.map((s) => `<div class="ai-pop-item" data-text="${esc(s)}">${esc(s)}</div>`).join('')
          : '<div class="ai-pop-sub">(AI không trả về gợi ý)</div>');
    } catch (e) {
      if (aiPop) aiBox.innerHTML = '<div class="ai-pop-sub ai-err">Lỗi AI: ' + esc(e.message) + '</div>';
    }
  }

  // ============================ Xuất Excel =========================
  // Giải mã base64 -> Uint8Array (template .xlsx nhúng sẵn)
  function b64ToBytes(b64) {
    const bin = atob(b64);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return bytes;
  }

  function exportXlsx() {
    if (!state.processes.length) return;
    if (!window.PFMEA_TEMPLATE_B64 || !window.fflate || !window.TemplateExport) {
      alert('Thiếu template hoặc thư viện nén. Không thể xuất.');
      return;
    }
    // Đổ dữ liệu vào file template gốc -> GIỮ NGUYÊN 100% định dạng P-FMEA
    const templateBytes = b64ToBytes(window.PFMEA_TEMPLATE_B64);
    const out = window.TemplateExport.buildFromTemplate(state, templateBytes, window.fflate);

    const blob = new Blob([out], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    const today = new Date().toISOString().slice(0, 10);
    a.href = url;
    a.download = `P-FMEA_${today}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function onExportClick() { readMetaInputs(); exportXlsx(); }

  // ======================= Đồng bộ đám mây (Supabase) =======================
  function initCloud() {
    try {
      if (window.supabase && SB_URL && SB_KEY) sb = window.supabase.createClient(SB_URL, SB_KEY);
    } catch (e) { sb = null; }
    if (sb) cloudStatus('online');
  }
  function cloudStatus(state) {
    const el = $('#cloudDot'); if (!el) return;
    el.className = 'cloud-dot ' + state;
    el.title = state === 'online' ? 'Đã kết nối dữ liệu chung' : 'Chưa kết nối dữ liệu chung (đang dùng cục bộ)';
  }

  // --- Kéo dữ liệu chung về (gộp vào localStorage) ---
  async function cloudPullProjects() {
    if (!sb) return;
    try {
      const { data, error } = await sb.from('projects').select('key,meta,processes,updated_at');
      if (error || !data) return;
      const map = readProjects();
      data.forEach((r) => { map[r.key] = { meta: r.meta, processes: r.processes, savedAt: r.updated_at }; });
      writeProjects(map); refreshProjectUI();
    } catch (e) { /* offline */ }
  }
  async function cloudPullContexts() {
    if (!sb) return;
    try {
      const { data, error } = await sb.from('contexts').select('dept,data');
      if (error || !data) return;
      const map = readContexts();
      data.forEach((r) => { if (r.data) map[r.dept] = r.data; });
      localStorage.setItem(LS_CONTEXT, JSON.stringify(map));
    } catch (e) { /* offline */ }
  }
  async function cloudPullPhrases() {
    if (!sb) return;
    try {
      const { data, error } = await sb.from('phrases').select('field,dept,body');
      if (error || !data) return;
      const all = readPhrases();
      data.forEach((r) => {
        const k = r.field + '|' + r.dept; const list = all[k] || [];
        if (r.body && !list.some((p) => p.trim().toLowerCase() === r.body.trim().toLowerCase())) list.push(r.body);
        all[k] = list;
      });
      localStorage.setItem(LS_PHRASES, JSON.stringify(all));
    } catch (e) { /* offline */ }
  }

  // --- Đẩy thay đổi lên dữ liệu chung (fire-and-forget) ---
  function cloudPushProject(key, proj) {
    if (!sb || !key || !proj) return;
    try {
      sb.from('projects').upsert({ key, meta: proj.meta, processes: proj.processes, updated_at: new Date().toISOString() })
        .then(() => {}, () => {});
    } catch (e) { /* noop */ }
  }
  function cloudPushContext(dept, data) {
    if (!sb || !dept) return;
    try { sb.from('contexts').upsert({ dept, data, updated_at: new Date().toISOString() }).then(() => {}, () => {}); }
    catch (e) { /* noop */ }
  }
  function cloudPushPhrase(field, dept, body) {
    if (!sb || !field || !body) return;
    try { sb.from('phrases').insert({ field, dept, body }).then(() => {}, () => {}); } // unique -> tự bỏ trùng
    catch (e) { /* noop */ }
  }

  // ============================ Khởi tạo ===========================
  function init() {
    buildHeader();

    // Mở trang mới = KHÔNG hiển thị dữ liệu gì. Dữ liệu chỉ xuất hiện khi người
    // dùng CHỌN (dự án đã lưu / đủ Bộ phận-SP-Dây chuyền-Model khớp dự án) hoặc
    // UPLOAD Control Plan. Không tự khôi phục phiên cũ.
    writeMetaInputs();   // meta rỗng -> dropdown để trống
    refreshProjectUI();  // chỉ nạp danh sách dự án đã lưu vào ô "Mở dự án"
    render();            // bảng trống
    $('#btnAddProc').hidden = false; $('#btnClear').hidden = false;

    $('#fileCP').addEventListener('change', onFile);
    $('#btnLoad').addEventListener('click', onLoadProc);
    $('#btnAddProc').addEventListener('click', onAddProc);
    $('#btnClear').addEventListener('click', onClear);
    $('#btnExport').addEventListener('click', onExportClick);
    $('#btnSave').addEventListener('click', onSave);
    $('#projSelect').addEventListener('change', onPickProject);
    $('#btnDeleteProj').addEventListener('click', onDeleteProject);
    $('#btnExportJson').addEventListener('click', onExportJson);
    $('#fileJson').addEventListener('change', onImportJson);
    // Bộ phận/Sản phẩm/Dây chuyền xác định ngữ cảnh P-FMEA (mỗi tổ hợp = 1 bản riêng).
    // Đổi bất kỳ 3 ô này → chuyển sang bản P-FMEA khác:
    //   • Nếu đang có dữ liệu → hỏi xác nhận trước khi xóa.
    //   • Sau khi xóa / nếu bảng trống → tự mở dự án đã lưu nếu khớp tổ hợp mới.
    // Model là nhãn tự do — không kích hoạt xét ngữ cảnh.
    function onMetaChange() {
      readMetaInputs();
      if (!state.processes.length) {
        const proj = readProjects()[currentKey()];
        if (proj && proj.processes && proj.processes.length) {
          state.processes = migrateState({ processes: proj.processes }).processes; reindexUID(); render();
          $('#projSelect').value = currentKey();
        }
      }
      scheduleAutosave();
    }
    // Xác nhận chuyển ngữ cảnh (nếu đang có dữ liệu); trả về false nếu người dùng hủy.
    function confirmContext() {
      if (!state.processes.length) return true;
      return confirm('Bạn đang chuyển sang bản P-FMEA khác.\nDữ liệu hiện tại chưa lưu sẽ bị xóa.\nTiếp tục?');
    }
    $('#mDept').addEventListener('change', () => {
      const newDept = $('#mDept').value;
      if (!confirmContext()) {
        // Hoàn tác: state.meta chưa được readMetaInputs() cập nhật → vẫn là giá trị cũ
        $('#mDept').value = state.meta.dept || '';
        return;
      }
      if (state.processes.length) { state.processes = []; render(); }
      fillProduct(newDept, '');
      fillLine(newDept, '', '');
      onMetaChange();
    });
    $('#mProduct').addEventListener('change', () => {
      const newProduct = $('#mProduct').value;
      if (!confirmContext()) {
        $('#mProduct').value = state.meta.product || '';
        return;
      }
      if (state.processes.length) { state.processes = []; render(); }
      fillLine($('#mDept').value, newProduct, '');
      onMetaChange();
    });
    $('#mLine').addEventListener('change', () => {
      if (!confirmContext()) {
        $('#mLine').value = state.meta.line || '';
        return;
      }
      if (state.processes.length) { state.processes = []; render(); }
      onMetaChange();
    });
    // Model: gõ tự do; khi nhập xong (change) thử tự mở dự án đã lưu nếu bảng trống.
    $('#mModel').addEventListener('change', onMetaChange);
    $('#mModel').addEventListener('input', () => { readMetaInputs(); scheduleAutosave(); });

    const tbody = $('#fmea tbody');
    tbody.addEventListener('input', onInput);
    tbody.addEventListener('change', onChange);
    tbody.addEventListener('click', onClick);
    // Tự gõ trực tiếp vào ô -> lưu câu vào bộ nhớ của cột đó (khi rời ô)
    tbody.addEventListener('focusout', (e) => {
      const el = e.target; const field = el.dataset && el.dataset.field;
      if (field === 'effectAnalysis' || field === 'cause' || field === 'prevention' || field === 'detectCause') {
        savePhrase(field, (el.tagName === 'TEXTAREA' || el.tagName === 'INPUT') ? el.value : el.textContent);
      }
    });

    // Modal bối cảnh AI
    $('#btnAIContext').addEventListener('click', openAIModal);
    $('#aiModalClose').addEventListener('click', closeAIModal);
    $('#aiModalCancel').addEventListener('click', closeAIModal);
    $('#aiCtxSave').addEventListener('click', saveAIContext);
    $('#aiModal').addEventListener('click', (e) => { if (e.target.id === 'aiModal') closeAIModal(); });

    // Đồng bộ đám mây: kết nối + kéo dữ liệu chung về; làm tươi khi mở dropdown dự án
    initCloud();
    cloudPullProjects(); cloudPullContexts(); cloudPullPhrases();
    $('#projSelect').addEventListener('focus', cloudPullProjects);

    // Tab switching
    document.querySelectorAll('.tab-btn').forEach(function(btn) {
      btn.addEventListener('click', function() {
        var targetId = this.dataset.tab;
        document.querySelectorAll('.tab-btn').forEach(function(b) { b.classList.remove('active'); });
        this.classList.add('active');
        document.querySelectorAll('.tab-panel').forEach(function(p) { p.hidden = true; });
        document.getElementById(targetId).hidden = false;
      });
    });

    // Dựng tab Hướng dẫn
    setupGuide();
  }

  /* =====================================================================
   * TAB HƯỚNG DẪN — dựng động theo từng cột P-FMEA
   * ===================================================================== */

  // Nhãn từng cột (khớp với buildHeader) + ví dụ minh họa
  // Ví dụ lấy từ CP G823-00 (PRO2 — giảm xóc HÀN CAP)
  const G_COLS = {
    A: { vi: 'Quy trình / Bước /\nChức năng\n(Hạng mục yêu cầu)', jp: 'プロセス ステップ/機能\n要求事項',
      ex: '<b>2. HÀN CAP</b>\n<i>Chức năng:</i> Liên kết damper cap với main pipe\n\n<i>Yêu cầu (từ CP):</i>\n1. Kích thước chiều cao tổng: 203.3(+0.3/-1.1)\n2. Chiều rộng mối hàn ≥ 3 mm\n3. Độ kín khí: không rò rỉ' },
    B: { vi: 'Dạng hỏng hóc\ntiềm ẩn', jp: '潜在的故障モード',
      ex: '<i>Yêu cầu 1 — dung sai 2 phía → 2 dạng hỏng:</i>\n① Kích thước chiều cao tổng 203.3 <b>lớn hơn tiêu chuẩn</b>\n② Kích thước chiều cao tổng 203.3 <b>nhỏ hơn tiêu chuẩn</b>\n\n<i>Yêu cầu 3 — văn bản → 1 dạng hỏng:</i>\n③ Độ kín khí không đạt' },
    C: { vi: 'Ảnh hưởng của\nhỏng hóc tiềm ẩn', jp: '潜在的故障影響',
      ex: '① Không lắp được vào thân giảm xóc ở công đoạn sau / Phải sửa ngoài dây chuyền\n② "Một số sản phẩm phải sửa ngoài dây chuyền hoặc phế phẩm" → S = 5' },
    D: { vi: 'Mức độ\nnghiêm trọng (S)', jp: '厳しさ', ex: '5\n(chọn từ câu kết luận ở cột C)' },
    E: { vi: 'Phân loại\n(Đặc tính đặc thù)', jp: '分類',
      ex: 'S\n(Safety — tự động từ CP)' },
    F: { vi: 'Nguyên nhân\ncủa hỏng hóc', jp: '潜在的故障原因',
      ex: '<i>Machine:</i> Máy hàn lệch tọa độ trục Z\n<i>Man:</i> Cài đặt sai thông số hàn\n<i>Method:</i> Gá phôi không đúng vị trí\n<i>Material:</i> Phôi cap sai chiều cao đầu vào' },
    G: { vi: 'Phản ánh\nlỗi quá khứ', jp: '過去不具合の反映',
      ex: '<i>Nguyên nhân “Máy hàn lệch tọa độ trục Z”</i>\n→ đã từng phát sinh lỗi thực tế trước đây\n→ điền <b>“o”</b>\n\n<i>Nguyên nhân chưa từng phát sinh</i>\n→ để trống' },
    H: { vi: 'Tần suất\nphát sinh (O)', jp: '発生頻度', ex: '3' },
    I: { vi: 'Quản lý hiện tại\n— Dự phòng', jp: '現行管理 予防',
      ex: 'Bảo dưỡng máy hàn định kỳ hàng tháng\nKiểm tra tọa độ bằng mẫu chuẩn đầu mỗi ca\nKiểm tra phôi cap đầu vào theo CP' },
    J: { vi: 'Quản lý hiện tại\n— Phát hiện ra', jp: '現行管理 検出',
      ex: '① Kiểm tra tọa độ máy hàn bằng mẫu chuẩn (đầu ca)\n② Đo chiều cao tổng bằng thước cặp, 5 sp/lần, theo TC QC\n③ [SC] Kiểm tra 100% độ kín khí bằng thiết bị đo khí nén' },
    K: { vi: 'Phát hiện (D)', jp: '検出', ex: '4' },
    L: { vi: 'RPN', jp: '', ex: 'S×O×D = 5×3×4 = 60' },
    M: { vi: 'Biện pháp đề xuất\n+ Kết quả xử lý', jp: '推奨処置 / 処置結果',
      ex: 'Bổ sung cảm biến đo chiều cao tự động trong máy hàn\n→ S′=5, O′=2, D′=2\n→ RPN′ = 20' },
  };

  function gColChip(tag) {
    const c = G_COLS[tag];
    if (!c) return '';
    return '<div class="gcol-chip"><div class="gcol-head">'
      + '<span class="gcol-tag">' + tag + '</span>'
      + '<span class="gcol-name">' + esc(c.vi) + '</span>'
      + (c.jp ? '<span class="gcol-jp">' + esc(c.jp) + '</span>' : '')
      + '</div>'
      + (c.ex ? '<div class="gcol-ex"><div class="gcol-ex-label">Ví dụ</div><div class="gcol-ex-body">' + c.ex + '</div></div>' : '')
      + '</div>';
  }

  // Bảng tiêu chuẩn Mức độ nghiêm trọng S — render trực tiếp từ data thực
  function gSeverityTable(scope) {
    const rows = (window.SEVERITY_TABLE || []).filter(r => r.scope === scope).slice()
      .sort((a, b) => b.rank - a.rank);
    const body = rows.map(r =>
      '<tr><td class="s-rank s-' + r.rank + '">' + r.rank + '</td><td>' + esc(r.category) + '</td><td>' + esc(r.text) + '</td></tr>'
    ).join('');
    return '<table class="guide-table guide-s-table"><thead><tr><th>Rank S</th><th>Nhóm ảnh hưởng</th><th>Tiêu chuẩn đánh giá</th></tr></thead><tbody>' + body + '</tbody></table>';
  }

  const G_OFFICIAL = '<div class="guide-official">📘 <b>Bảng tiêu chuẩn chính thức</b> — Bảng-2 Tiêu chuẩn đánh giá P-FMEA (GL SQS0811). Trường hợp khách hàng có tiêu chuẩn riêng thì thực hiện theo tiêu chuẩn của khách hàng.</div>';

  // (2) Tiêu chuẩn đánh giá Tần suất phát sinh — nguyên văn theo Bảng-2
  function gOccurrenceTable() {
    return G_OFFICIAL + '<table class="guide-table"><thead><tr><th>Rank O</th><th>Khả năng của hỏng hóc</th><th>Tần suất phát sinh của nguyên nhân<br><small>(số vấn đề / mặt hàng · xe)</small></th><th>Giá trị Cp</th></tr></thead><tbody>'
      + '<tr><td class="s-rank s-10">10</td><td rowspan="1">Rất cao</td><td>100/1.000 trở lên &nbsp;·&nbsp; 1/10 trở lên</td><td>—</td></tr>'
      + '<tr><td class="s-rank s-9">9</td><td>Rất cao</td><td>50/1.000 &nbsp;·&nbsp; 1/20</td><td>Cp &gt; 0,67</td></tr>'
      + '<tr><td class="s-rank s-8">8</td><td rowspan="2">Cao</td><td>20/1.000 &nbsp;·&nbsp; 1/50</td><td>—</td></tr>'
      + '<tr><td class="s-rank s-7">7</td><td>10/1.000 &nbsp;·&nbsp; 1/100</td><td>—</td></tr>'
      + '<tr><td class="s-rank s-6">6</td><td rowspan="3">Mức độ trung bình</td><td>2/1.000 &nbsp;·&nbsp; 1/500</td><td>Cp &gt; 1,0</td></tr>'
      + '<tr><td class="s-rank s-5">5</td><td>0,5/1.000 &nbsp;·&nbsp; 1/2.000</td><td>—</td></tr>'
      + '<tr><td class="s-rank s-4">4</td><td>0,1/1.000 &nbsp;·&nbsp; 1/10.000</td><td>—</td></tr>'
      + '<tr><td class="s-rank s-3">3</td><td rowspan="2">Thấp</td><td>0,01/1.000 &nbsp;·&nbsp; 1/100.000</td><td>Cp ≥ 1,33</td></tr>'
      + '<tr><td class="s-rank s-2">2</td><td>0,001/1.000 trở xuống &nbsp;·&nbsp; 1/1.000.000</td><td>Cp ≥ 1,67</td></tr>'
      + '<tr><td class="s-rank s-1">1</td><td>Rất thấp</td><td>Hỏng hóc được loại bỏ bằng quản lý dự phòng</td><td>—</td></tr>'
      + '</tbody></table>';
  }

  // (3) Tiêu chuẩn đánh giá Phát hiện ra — nguyên văn theo Bảng-2
  function gDetectionTable() {
    return G_OFFICIAL + '<table class="guide-table guide-d-table"><thead><tr><th>Rank D</th><th>Cơ hội phát hiện ra</th><th>Tiêu chuẩn: phương pháp phát hiện</th><th>Khả năng</th></tr></thead><tbody>'
      + '<tr><td class="s-rank s-10">10</td><td>Không có cơ hội phát hiện ra</td><td>Không quản lý công đoạn hiện tại. Không thể phát hiện ra hoặc không thể phân tích được</td><td>Hầu như không thể</td></tr>'
      + '<tr><td class="s-rank s-9">9</td><td>Ở bất kỳ giai đoạn nào, khả năng phát hiện gần như không có (VD: đánh giá xác suất)</td><td>Không thể dễ dàng phát hiện ra dạng hỏng hóc và/hoặc lỗi (nguyên nhân)</td><td>Rất ít</td></tr>'
      + '<tr><td class="s-rank s-8">8</td><td><b>Sau khi kết thúc gia công</b></td><td>Phát hiện dạng hỏng hóc sau khi kết thúc gia công bằng <b>thị giác / xúc giác / thính giác</b> của người thao tác</td><td>Ít</td></tr>'
      + '<tr><td class="s-rank s-7">7</td><td><b>Tại thời điểm phát sinh</b></td><td>Phát hiện bằng <b>giác quan người</b> tại hiện trường phát sinh; hoặc dưỡng <b>GO/NO-GO</b> sau khi kết thúc gia công</td><td>Rất thấp</td></tr>'
      + '<tr><td class="s-rank s-6">6</td><td><b>Sau khi kết thúc gia công</b></td><td>Phát hiện bằng <b>dưỡng đo giá trị số</b> sau gia công; hoặc dưỡng GO/NO-GO tại hiện trường phát sinh</td><td>Thấp</td></tr>'
      + '<tr><td class="s-rank s-5">5</td><td><b>Tại thời điểm phát sinh</b></td><td>Dưỡng đo giá trị số tại hiện trường; hoặc <b>quản lý tự động</b> báo đèn/còi khi có linh kiện lỗi; kiểm tra 2 lần do người</td><td>Trung bình</td></tr>'
      + '<tr><td class="s-rank s-4">4</td><td><b>Sau khi kết thúc gia công</b></td><td>Phát hiện dạng hỏng hóc bằng <b>quản lý tự động</b> sau gia công (chặn đứng, tránh gia công thêm)</td><td>Tương đối cao</td></tr>'
      + '<tr><td class="s-rank s-3">3</td><td><b>Tại thời điểm phát sinh</b></td><td>Phát hiện dạng hỏng hóc bằng <b>quản lý tự động</b> tại hiện trường phát sinh (chặn đứng)</td><td>Cao</td></tr>'
      + '<tr><td class="s-rank s-2">2</td><td>Phát hiện lỗi và/hoặc dự phòng vấn đề</td><td>Phát hiện <b>lỗi (nguyên nhân)</b> bằng quản lý tự động tại hiện trường — dự phòng việc sản xuất ra linh kiện lỗi</td><td>Rất cao</td></tr>'
      + '<tr><td class="s-rank s-1">1</td><td>Không áp dụng phát hiện: dự phòng lỗi</td><td><b>Poka-yoke</b> — dự phòng lỗi bằng thiết kế đồ gá / máy / linh kiện nên không sản xuất ra linh kiện lỗi</td><td>Gần như chắc chắn phát hiện được</td></tr>'
      + '</tbody></table>';
  }

  // Tiêu chuẩn thực hiện biện pháp đề xuất (theo file tiêu chuẩn của công ty)
  function gActionTable() {
    return '<table class="guide-table guide-action-table"><thead><tr><th style="min-width:120px">Phân loại</th><th style="min-width:160px">Tiêu chuẩn thực hiện biện pháp</th><th>Cách suy nghĩ khi thực hiện biện pháp đề xuất</th></tr></thead><tbody>'
      + '<tr><td><b>Mức độ nghiêm trọng (S)</b></td><td style="text-align:center;font-weight:700">9 trở lên</td>'
      + '<td>Khi cấp độ nghiêm trọng là <b>9 trở lên</b>, biện pháp đề xuất phải lấy cơ sở là <b>thay đổi thiết kế, thay đổi công đoạn</b>.<br><small>(Phòng chống “Lỗi làm mất chức năng cơ bản của xe — chạy, rẽ, dừng”, không điều khiển được mà không có dấu hiệu báo trước, và lỗi gây cháy/bốc khói/dẫn tới tai nạn.)</small></td></tr>'
      + '<tr><td><b>Tần suất phát sinh (O)</b></td><td style="text-align:center;font-weight:700">4 trở lên</td>'
      + '<td>Năng lực công đoạn <b>Cp và Cpk ≥ 1,33</b> là tiêu chuẩn của tập đoàn. Nếu đảm bảo được năng lực công đoạn thì tần suất phát sinh là từ <b>3 trở xuống</b>.</td></tr>'
      + '<tr><td><b>Phát hiện ra (D)</b></td><td>TH1: S ≥ 7 <b>và</b> D ≥ 6<br>TH2: S ≤ 6 và (ngoài kiểm tra bằng giác quan) D ≥ 7</td>'
      + '<td>Với các yếu tố ảnh hưởng dẫn tới tổn hại chức năng nghiêm trọng, thực hiện biện pháp giảm mức độ phát hiện ra <b>xuống dưới 5</b>.<br><small>(Lý do loại bỏ kiểm tra bằng giác quan: kiểm tra giác quan thuộc lĩnh vực tính thương phẩm, ít ảnh hưởng đến chức năng.)</small></td></tr>'
      + '</tbody></table>';
  }

  // Bảng format tổng quan (trang đầu)
  function gFormatTable() {
    const top = [
      ['A', 'Quy trình / Bước /\nChức năng'], ['B', 'Dạng hỏng hóc\ntiềm ẩn'],
      ['C', 'Ảnh hưởng của\nhỏng hóc'], ['D', 'Mức độ\nnghiêm trọng\n(S)'],
      ['E', 'Phân loại'], ['F', 'Nguyên nhân'], ['G', 'Phản ánh\nlỗi quá khứ'],
      ['H', 'Tần suất\nphát sinh (O)'], ['I', 'Quản lý\nDự phòng'],
      ['J', 'Quản lý\nPhát hiện ra'], ['K', 'Phát hiện\n(D)'], ['L', 'RPN'],
      ['M', 'Biện pháp\nđề xuất'], ['N', 'Trách nhiệm\n& thời hạn'],
    ];
    const cells = top.map(c => '<th><span class="ff-tag">' + c[0] + '</span>\n' + esc(c[1]) + '</th>').join('');
    const sub = [['O', 'Biện pháp\nđã thực hiện'], ['P', "S'"], ['Q', "O'"], ['R', "D'"], ['S', "RPN'"]];
    const subCells = sub.map(c => '<th class="ff-grp"><span class="ff-tag">' + c[0] + '</span>\n' + esc(c[1]) + '</th>').join('');
    return '<div class="guide-format-scroll"><table class="guide-format-table"><thead>'
      + '<tr>' + cells + '<th class="ff-grp" colspan="5">Kết quả xử lý 処置結果</th></tr>'
      + '<tr>' + subCells + '</tr>'
      + '</thead></table></div>';
  }

  // Định nghĩa các trang hướng dẫn (dữ liệu ví dụ từ CP G823-00)
  function gPagesThink() {
    return [

      // ===== Slide 1: Tổng thể =====
      {
        tag: '', title: 'Tổng thể nội dung P-FMEA', menu: 'Tổng thể — Format P-FMEA',
        full: true,
        body:
          '<div class="g-block"><p><strong>P-FMEA</strong> (Process Failure Mode and Effects Analysis) là phương pháp phân tích có hệ thống nhằm phát hiện và đánh giá các <em>dạng hỏng hóc tiềm ẩn</em> của quá trình sản xuất, từ đó đưa ra biện pháp kiểm soát và cải tiến phòng ngừa <em>trước khi</em> sản xuất hàng loạt.</p>'
          + '<ul style="font-size:13px;line-height:1.8">'
          + '<li><b>Nguồn dữ liệu chính:</b> Control Plan (CP) — mỗi hạng mục chất lượng trong CP thành một hoặc hai yêu cầu phân tích trong P-FMEA.</li>'
          + '<li><b>Đầu ra:</b> Bảng P-FMEA đầy đủ, dùng quản lý rủi ro chất lượng và lập kế hoạch kiểm soát.</li>'
          + '<li><b>Áp dụng:</b> Lập mới khi có sản phẩm mới / thay đổi quy trình; cập nhật khi phát sinh lỗi thực tế.</li>'
          + '</ul></div>'
          + '<div class="g-block"><h5>📋 Bố cục bảng P-FMEA (19 cột A→S)</h5>'
          + gFormatTable()
          + '</div>'
          + '<div class="g-block"><h5>🔄 Luồng phân tích — thứ tự điền cột</h5>'
          + '<ol style="font-size:13px;color:var(--ink-soft);line-height:2;padding-left:20px">'
          + '<li><b>A</b> — Xác định quy trình, chức năng và yêu cầu (từ CP, tự động)</li>'
          + '<li><b>B</b> — Liệt kê dạng hỏng hóc (phủ định của yêu cầu, tự động)</li>'
          + '<li><b>C</b> — Phân tích ảnh hưởng → chọn câu kết luận theo bảng tiêu chuẩn</li>'
          + '<li><b>D</b> — Điểm S (tự động từ câu kết luận ở C)</li>'
          + '<li><b>E</b> — Ký hiệu SC — đặc tính đặc thù (tự động từ CP)</li>'
          + '<li><b>F/G</b> — Nguyên nhân 4M, phản ánh lỗi quá khứ</li>'
          + '<li><b>H</b> — Điểm O (tần suất phát sinh)</li>'
          + '<li><b>I</b> — Biện pháp dự phòng hiện tại</li>'
          + '<li><b>J</b> — Biện pháp phát hiện (tự động điền phần ② từ CP)</li>'
          + '<li><b>K</b> — Điểm D (khả năng phát hiện)</li>'
          + '<li><b>L</b> — RPN = S×O×D (tự động)</li>'
          + '<li><b>M→S</b> — Biện pháp đề xuất, kết quả xử lý, S′/O′/D′/RPN′</li>'
          + '</ol></div>'
          + '<div class="g-block"><h5>🧭 Cách dùng tab Hướng dẫn này</h5><ul style="font-size:13px;line-height:1.8">'
          + '<li>Hướng dẫn có <b>2 lựa chọn</b> (nút chọn ở phía trên): <b>🧠 Hiểu &amp; tư duy đúng</b> — giải thích ý nghĩa &amp; cách tư duy từng cột; và <b>🖱️ Thao tác trên web</b> — các bước thực hiện trực tiếp trên công cụ.</li>'
          + '<li>Chuyển slide bằng <b>‹ Trang trước</b> / <b>Trang sau ›</b>, hoặc chọn nhanh ở ô danh sách.</li>'
          + '<li>Mỗi slide cột: <b>bên trái</b> hiện tên cột + ví dụ thực tế từ P-FMEA <b>G823-00</b> (Damper case comp); <b>bên phải</b> là cách hiểu &amp; cách tư duy.</li>'
          + '<li>Toàn bộ điểm S / O / D chấm theo <b>Bảng-2 Tiêu chuẩn đánh giá P-FMEA (GL SQS0811)</b>.</li>'
          + '</ul></div>',
      },

      // ===== Slide 2: Cột A =====
      {
        tag: 'A', title: 'Cột A — Quy trình / Bước / Chức năng / Yêu cầu',
        body:
          '<div class="g-block"><h5>🔎 Cách hiểu</h5>'
          + '<p>Cột A xác định <b>công đoạn</b> đang phân tích và <b>danh sách yêu cầu chất lượng</b> của công đoạn đó. Đây là nền tảng để liệt kê dạng hỏng hóc ở cột B.</p></div>'
          + '<div class="g-block"><h5>✍️ Cách làm</h5><ul>'
          + '<li><b>Số thứ tự và tên công đoạn</b> phải <b>đồng nhất với Quy trình công nghệ (QTCN)</b> — không tự đặt tên khác.</li>'
          + '<li><b>Chức năng công đoạn:</b> mô tả ngắn gọn công đoạn làm gì (thường 1–2 câu).</li>'
          + '<li><b>Yêu cầu:</b> lấy trực tiếp từ cột <em>”Hạng mục quản lý”</em> trong Control Plan — không viết lại theo cách khác. Mỗi hạng mục CP → 1 dòng yêu cầu, tool tự điền khi nạp CP.</li>'
          + '<li>Sau khi nạp CP, có thể thêm yêu cầu bằng nút <b>＋ Thêm yêu cầu</b>, xóa bằng nút <b>✕</b>.</li>'
          + '</ul></div>'
          + '<div class="g-eg"><b>Ví dụ G823-00 — Công đoạn 2 “HÀN CAP”:</b><br>'
          + '<i>Chức năng:</i> Liên kết giữa damper cap và main pipe<br><br>'
          + '<i>Yêu cầu từ CP:</i><br>'
          + '&nbsp;&nbsp;1. Kiểm tra số lượng, chủng loại: Đủ, đúng chủng loại<br>'
          + '&nbsp;&nbsp;2. Kích thước chiều cao tổng: 203.3(+0.3/-1.1) mm<br>'
          + '&nbsp;&nbsp;3. Chiều rộng mối hàn ≥ 3 mm<br>'
          + '&nbsp;&nbsp;4. Độ kín khí: không rò rỉ</div>'
          + '<div class="guide-note">💡 Tên công đoạn và số thứ tự có thể chỉnh tay trong ô STT và ô tên. Chỉnh tên không ảnh hưởng đến dữ liệu yêu cầu.</div>',
      },

      // ===== Trang 3: Cột B =====
      {
        tag: 'B', title: 'Cột B — Dạng hỏng hóc tiềm ẩn',
        body:
          '<div class="g-block"><h5>🔎 Cách hiểu</h5>'
          + '<p>Dạng hỏng hóc là trạng thái mà công đoạn <b>không đáp ứng được yêu cầu</b> — tức là <b>phủ định của yêu cầu</b>. Phải nêu rõ theo hướng nào không đạt (lớn hơn / nhỏ hơn / vượt ngưỡng / sai kiểu…).</p></div>'
          + '<div class="g-block"><h5>✍️ Quy tắc xác định số dạng hỏng hóc</h5>'
          + '<table class="guide-table"><thead><tr><th style="min-width:180px">Loại yêu cầu từ CP</th><th style="min-width:60px;text-align:center">Số dạng hỏng</th><th>Cách viết</th></tr></thead><tbody>'
          + '<tr><td>Dung sai <b>2 phía</b><br><small style="color:var(--muted)">VD: 203.3(+0.3/-1.1) hoặc ±0.5</small></td>'
          + '<td style="text-align:center;font-weight:800;color:#1d4ed8;font-size:16px">2</td>'
          + '<td>① &lt;tên&gt; <b>lớn hơn tiêu chuẩn</b><br>② &lt;tên&gt; <b>nhỏ hơn tiêu chuẩn</b></td></tr>'
          + '<tr><td>Dung sai / giá trị <b>1 phía</b><br><small style="color:var(--muted)">VD: max 0.3, min 14 kN, ≤ 5 mm</small></td>'
          + '<td style="text-align:center;font-weight:700">1</td>'
          + '<td>&lt;tên&gt; <b>không đạt</b></td></tr>'
          + '<tr><td>Yêu cầu <b>văn bản</b> (Pass/Fail)<br><small style="color:var(--muted)">VD: không rò rỉ, đúng chủng loại</small></td>'
          + '<td style="text-align:center;font-weight:700">1</td>'
          + '<td>&lt;tên&gt; <b>không đạt</b></td></tr>'
          + '</tbody></table></div>'
          + '<div class="g-eg"><b>Ví dụ G823-00:</b><br>'
          + 'Yêu cầu: <i>”Kích thước chiều cao tổng: 203.3(+0.3/-1.1)”</i> → dung sai 2 phía:<br>'
          + '&nbsp;&nbsp;① Kích thước chiều cao tổng 203.3 <b>lớn hơn tiêu chuẩn</b><br>'
          + '&nbsp;&nbsp;② Kích thước chiều cao tổng 203.3 <b>nhỏ hơn tiêu chuẩn</b><br><br>'
          + 'Yêu cầu: <i>”Kiểm tra số lượng, chủng loại: Đủ, đúng chủng loại”</i> → văn bản:<br>'
          + '&nbsp;&nbsp;① Kiểm tra số lượng, chủng loại <b>không đạt</b></div>'
          + '<div class="guide-note">💡 Tool tự phân tích spec và tolerance từ CP để tạo 1 hoặc 2 dạng hỏng. Trong cột A hiển thị <b>1 dòng yêu cầu</b> nhưng cột B có thể có <b>2 dạng hỏng</b> (không nhân đôi yêu cầu). Dùng <b>🔗 Gộp</b> để gộp các dạng hỏng tương tự khi chúng có cùng nguyên nhân và phương pháp kiểm soát.</div>',
      },

      // ===== Trang 4: Cột C =====
      {
        tag: 'C', title: 'Cột C — Ảnh hưởng của dạng hỏng hóc',
        body:
          '<div class="g-block"><h5>🔎 Cách hiểu</h5>'
          + '<p>Cột C đòi hỏi <b>kiến thức thực tế về dây chuyền và sản phẩm</b>. Ảnh hưởng gồm 2 phần:</p>'
          + '<ul>'
          + '<li><b>① Ảnh hưởng đến công đoạn / quá trình:</b> bao gồm toàn bộ quá trình từ sản xuất đến lắp ráp (kể cả lắp ráp tại khách hàng mà mình cung cấp hàng). Hỏng hóc có gây ra sự cố gì trong quá trình gia công, lắp ráp không?</li>'
          + '<li><b>② Ảnh hưởng đến sản phẩm cuối / người dùng:</b> với ngành giảm xóc, sản phẩm cuối là <b>chiếc xe máy</b>, người dùng là <b>người lái xe</b>. Nếu hỏng hóc lọt ra ngoài thì ảnh hưởng gì đến người sử dụng?</li>'
          + '</ul>'
          + '<p>⚠ <b>Nếu hỏng hóc được ngăn chặn 100% trước khi lắp lên xe</b> → chỉ viết ảnh hưởng đến công đoạn; không cần viết ảnh hưởng đến sản phẩm.</p></div>'
          + '<div class="g-block"><h5>✍️ Cách làm</h5><ul>'
          + '<li>Câu <b>②</b> phải <b>trích dẫn nguyên văn từ bảng tiêu chuẩn đánh giá S</b> — đây là cơ sở để tool tự điền điểm S. Chọn câu phù hợp từ dropdown trong ô ảnh hưởng.</li>'
          + '<li>Câu <b>①</b> mô tả cụ thể ảnh hưởng đến dây chuyền: không lắp được, phải sửa ngoài, gây dừng máy, phế phẩm…</li>'
          + '<li>Chọn đúng phạm vi ảnh hưởng: công đoạn nội bộ hay ảnh hưởng đến tận khách hàng / người dùng?</li>'
          + '</ul></div>'
          + '<div class="g-eg"><b>Ví dụ G823-00 — “Chiều cao tổng lớn hơn tiêu chuẩn”:</b><br>'
          + '① Không lắp ráp được vào thân giảm xóc ở công đoạn sau / Phải sửa ngoài dây chuyền<br>'
          + '② <i>”Một số sản phẩm phải sửa ngoài dây chuyền hoặc phế phẩm”</i> → S = 5</div>'
          + '<div class="g-block"><h5>📊 Bảng tiêu chuẩn — Ảnh hưởng đến CÔNG ĐOẠN</h5>' + gSeverityTable('process') + '</div>'
          + '<div class="g-block"><h5>📊 Bảng tiêu chuẩn — Ảnh hưởng đến SẢN PHẨM (người dùng xe)</h5>' + gSeverityTable('product') + '</div>',
      },

      // ===== Trang 5: Cột D (S score) =====
      {
        tag: 'D', title: 'Cột D — Mức độ nghiêm trọng (S)',
        body:
          '<div class="g-block"><h5>🔎 Cách hiểu</h5>'
          + '<p>Điểm <b>S (Severity)</b> đánh giá <b>mức độ nghiêm trọng</b> của ảnh hưởng đã nêu ở cột C. Thang 1–10, điểm càng cao thì ảnh hưởng càng nghiêm trọng.</p></div>'
          + '<div class="g-block"><h5>✍️ Cách làm</h5><ul>'
          + '<li>Trong tool: chọn câu kết luận ở <b>ý ② của cột C</b> → điểm S <b>tự động điền</b> đúng theo bảng tiêu chuẩn. Không tự chấm cảm tính.</li>'
          + '<li>S <b>không thể giảm</b> bằng biện pháp kiểm soát — chỉ giảm được bằng thay đổi thiết kế sản phẩm/quy trình.</li>'
          + '<li>S không thay đổi dù có hay không có biện pháp phát hiện tốt đến mức nào.</li>'
          + '</ul></div>'
          + '<div class="guide-note">⚠ <b>Khi S ≥ 9:</b> Nếu điểm S từ 9 trở lên do <b>tính công nghệ của R&amp;D thiết kế</b> (không thể thay đổi bằng cải tiến sản xuất), phải tổng hợp hạng mục đó vào <b>trang tổng hợp đặc biệt</b> và gửi về <b>công ty mẹ tại Nhật để xin phê duyệt</b>.</div>'
          + '<div class="g-block"><h5>📊 Bảng tiêu chuẩn S — Ảnh hưởng đến CÔNG ĐOẠN</h5>' + gSeverityTable('process') + '</div>'
          + '<div class="g-block"><h5>📊 Bảng tiêu chuẩn S — Ảnh hưởng đến SẢN PHẨM</h5>' + gSeverityTable('product') + '</div>',
      },

      // ===== Slide 6: Cột F (Nguyên nhân) =====
      {
        tag: 'F', title: 'Cột F — Nguyên nhân của hỏng hóc (Phân tích 4M)',
        body:
          '<div class="g-block"><h5>🔎 Cách hiểu</h5>'
          + '<p>Cột F đòi hỏi <b>kiến thức thực tế về dây chuyền sản xuất</b>. Phải tìm ra các <b>nguyên nhân gốc rễ</b> có thể gây ra dạng hỏng hóc đã nêu ở cột B.</p>'
          + '<p>⚙ <b>Nguyên tắc cốt lõi:</b> nguyên nhân phải dựa trên <b>các điều kiện chế tạo của công đoạn</b>. Chỉ khi một <em>điều kiện chế tạo không đạt tiêu chuẩn</em> thì mới phát sinh ra dạng hỏng hóc. Vì vậy hãy đi từ điều kiện chế tạo (thông số máy, gá lắp, vật liệu đầu vào, thao tác chuẩn…) — điều kiện nào sai lệch sẽ trở thành nguyên nhân.</p></div>'
          + '<div class="g-block"><h5>✍️ Cách làm — Rà soát lần lượt 4M</h5>'
          + '<p>Soi từng <b>điều kiện chế tạo</b> theo <b>4 nhóm nguyên nhân</b>. Một dạng hỏng không nhất thiết có đủ cả 4 nhóm, nhưng phải <em>tư duy theo từng nhóm</em> để không bỏ sót:</p></div>'
          + '<div class="guide-4m-grid">'
          + '<div class="guide-4m-card m-man"><div class="m-title">👤 Man — Con người</div>'
          + '<p>Thao tác sai, thiếu kỹ năng, không theo SOP, mệt mỏi, nhầm lẫn…</p>'
          + '<div class="m-examples"><b>VD G823-00:</b><br>Cài đặt sai thông số hàn<br>Gá phôi không đúng vị trí</div></div>'
          + '<div class="guide-4m-card m-machine"><div class="m-title">⚙️ Machine — Máy móc</div>'
          + '<p>Dụng cụ mòn, thiết bị trục trặc, lệch tọa độ, không ổn định, cần hiệu chỉnh…</p>'
          + '<div class="m-examples"><b>VD G823-00:</b><br>Máy hàn lệch tọa độ trục Z<br>Điện cực hàn bị mòn</div></div>'
          + '<div class="guide-4m-card m-method"><div class="m-title">📋 Method — Phương pháp</div>'
          + '<p>Điều kiện gia công chưa tối ưu, thứ tự thao tác sai, thiếu bước kiểm tra, thông số không phù hợp…</p>'
          + '<div class="m-examples"><b>VD G823-00:</b><br>Thông số dòng điện hàn không phù hợp<br>Thiếu bước kiểm tra gá lắp trước hàn</div></div>'
          + '<div class="guide-4m-card m-material"><div class="m-title">📦 Material — Vật liệu</div>'
          + '<p>Phôi/linh kiện đầu vào sai kích thước, kém chất lượng, sai chủng loại, bề mặt nhiễm bẩn…</p>'
          + '<div class="m-examples"><b>VD G823-00:</b><br>Phôi cap sai chiều cao đầu vào<br>Bề mặt hàn nhiễm dầu/oxy hóa</div></div>'
          + '</div>'
          + '<div class="guide-note">💡 Trong tool, mỗi nguyên nhân là một hàng riêng. Thêm bằng nút <b>＋ Thêm nguyên nhân</b>; chọn nhóm 4M từ dropdown. Nút <b>✨ AI</b> gợi ý nguyên nhân dựa trên bối cảnh bộ phận đã cung cấp.</div>',
      },

      // ===== Trang 8: Cột H (O score) =====
      {
        tag: 'H', title: 'Cột H — Tần suất phát sinh (O)',
        body:
          '<div class="g-block"><h5>🔎 Cách hiểu</h5>'
          + '<p>Điểm <b>O (Occurrence)</b> đánh giá <b>xác suất nguyên nhân xảy ra</b> trong điều kiện sản xuất bình thường — <em>với các biện pháp dự phòng đang áp dụng</em>. Thang 1–10; điểm càng thấp thì nguyên nhân càng hiếm xảy ra.</p></div>'
          + '<div class="g-block"><h5>✍️ Cách làm</h5><ul>'
          + '<li>Chấm dựa trên <b>dữ liệu lỗi thực tế</b> (PPM, tỷ lệ hư hỏng lịch sử) hoặc <b>kinh nghiệm</b> với nguyên nhân tương tự.</li>'
          + '<li>Xét trong bối cảnh <b>đang có biện pháp dự phòng (cột I)</b> — không phải tần suất khi không có kiểm soát gì.</li>'
          + '<li>O có thể giảm bằng cách cải thiện biện pháp dự phòng (thay dao sớm hơn, bảo dưỡng máy thường xuyên hơn…).</li>'
          + '</ul></div>'
          + '<div class="g-block"><h5>📊 Bảng tiêu chuẩn đánh giá tần suất phát sinh</h5>' + gOccurrenceTable() + '</div>',
      },

      // ===== Slide 8: Cột G (Phản ánh lỗi quá khứ) =====
      {
        tag: 'G', title: 'Cột G — Phản ánh lỗi quá khứ',
        body:
          '<div class="g-block"><h5>🔎 Cách hiểu</h5>'
          + '<p>Cột G đánh dấu những <b>nguyên nhân đã từng phát sinh lỗi trong thực tế</b> (lỗi quá khứ). Mục đích: ưu tiên kiểm soát các nguyên nhân đã có tiền lệ, không để lỗi cũ lặp lại.</p></div>'
          + '<div class="g-block"><h5>✍️ Cách làm</h5><ul>'
          + '<li>Nguyên nhân nào <b>đã từng phát sinh</b> (có lịch sử lỗi, claim, NG thực tế) → điền ký hiệu <b>“o”</b> vào ô tương ứng.</li>'
          + '<li>Nguyên nhân <b>chưa từng phát sinh</b> → để trống.</li>'
          + '<li>Căn cứ: hồ sơ lỗi nội bộ, báo cáo 4M change, claim khách hàng, dữ liệu NG của công đoạn.</li>'
          + '</ul></div>'
          + '<div class="g-eg"><b>Ví dụ G823-00 — công đoạn HÀN CAP:</b><br>'
          + 'Nguyên nhân <i>“Máy hàn lệch tọa độ trục Z”</i> đã từng gây lỗi chiều cao tổng trong quá khứ → điền <b>“o”</b>.<br>'
          + 'Nguyên nhân <i>“Bề mặt hàn nhiễm dầu”</i> chưa từng phát sinh → để trống.</div>'
          + '<div class="guide-note">💡 Cột G giúp người xét duyệt nhận diện nhanh các điểm rủi ro đã có tiền lệ — thường đi kèm yêu cầu siết chặt biện pháp dự phòng (cột I) và phát hiện (cột J).</div>',
      },

      // ===== Slide 9: Cột I (Dự phòng) =====
      {
        tag: 'I', title: 'Cột I — Quản lý hiện tại: Dự phòng',
        body:
          '<div class="g-block"><h5>🔎 Cách hiểu</h5>'
          + '<p>Ghi <b>biện pháp đang thực sự được áp dụng</b> để ngăn <em>nguyên nhân</em> không xảy ra. Đây là biện pháp <b>phòng ngừa từ gốc</b> (khác với phát hiện ở cột J).</p></div>'
          + '<div class="g-block"><h5>✍️ Cách làm</h5><ul>'
          + '<li>Tự hỏi: <em>”Hiện tại đang làm gì để nguyên nhân này không xảy ra?”</em></li>'
          + '<li>Chỉ ghi biện pháp <b>đang thực hiện thực tế tại dây chuyền</b> — không ghi mong muốn hoặc kế hoạch tương lai.</li>'
          + '<li>Ghi cụ thể: tên biện pháp + chu kỳ + người thực hiện (nếu có).</li>'
          + '</ul></div>'
          + '<div class="g-eg"><b>Ví dụ G823-00 — Nguyên nhân “Máy hàn lệch tọa độ trục Z”:</b><br>'
          + 'Dự phòng: <i>”Bảo dưỡng máy hàn định kỳ hàng tháng; kiểm tra tọa độ bằng mẫu chuẩn đầu mỗi ca sản xuất”</i><br><br>'
          + '<b>Nguyên nhân “Phôi cap sai chiều cao đầu vào”:</b><br>'
          + 'Dự phòng: <i>”Kiểm tra chiều cao phôi cap tại công đoạn KIỂM TRA PHÔI ĐẦU VÀO theo CP trước khi đưa vào HÀN CAP”</i></div>',
      },

      // ===== Trang 10: Cột J (Phát hiện) =====
      {
        tag: 'J', title: 'Cột J — Quản lý hiện tại: Phát hiện ra',
        body:
          '<div class="g-block"><h5>🔎 Cách hiểu</h5>'
          + '<p>Cột J ghi cách phát hiện vấn đề, gồm <b>3 phần</b>:</p>'
          + '<ul>'
          + '<li><b>① Phát hiện ra nguyên nhân:</b> làm thế nào để biết nguyên nhân đang xảy ra? (VD: kiểm tra tọa độ máy bằng mẫu chuẩn đầu ca)</li>'
          + '<li><b>② Phát hiện ra dạng hỏng hóc:</b> làm thế nào để phát hiện sản phẩm lỗi? — <em>Tự động điền</em> từ CP (phương pháp + tần suất kiểm tra của hạng mục). Vẫn chỉnh được.</li>'
          + '<li><b>③ Kiểm tra đặc biệt cho SC</b> (chỉ xuất hiện khi hạng mục có ký hiệu SC): biện pháp kiểm tra bổ sung bắt buộc do tính an toàn/quan trọng của đặc tính.</li>'
          + '</ul></div>'
          + '<div class="g-block"><h5>✍️ Cách làm</h5><ul>'
          + '<li>Phần ② phải ghi rõ: <b>kiểm tra cái gì — bằng phương pháp gì — tần suất ra sao</b> — khớp với CP/tiêu chuẩn kiểm tra QC.</li>'
          + '<li>Hỏng hóc được phát hiện <b>ngay tại công đoạn</b> → D thấp hơn so với phát hiện <b>ở công đoạn sau</b>.</li>'
          + '<li>Kiểm tra bằng <b>giác quan</b> (mắt, tay, đếm) → D thường cao (7–9). Kiểm tra <b>tự động / thiết bị</b> → D thấp (2–5). Poka-yoke / không thể lọt → D = 1.</li>'
          + '<li>Phần ③ (SC): mô tả biện pháp kiểm tra 100% hoặc thiết bị đặc biệt dành riêng cho đặc tính an toàn.</li>'
          + '</ul></div>'
          + '<div class="g-eg"><b>Ví dụ G823-00 — Hạng mục “Kích thước chiều cao tổng” (SC = S):</b><br>'
          + '① Kiểm tra tọa độ máy hàn bằng mẫu chuẩn, đầu mỗi ca<br>'
          + '② Đo chiều cao tổng bằng thước cặp, tần suất 5 sản phẩm/lần, theo tiêu chuẩn kiểm tra QC<br>'
          + '③ [SC] Kiểm tra 100% độ kín khí bằng thiết bị đo khí nén chuyên dụng</div>'
          + '<div class="g-block"><h5>📊 Bảng tiêu chuẩn đánh giá phát hiện (D)</h5>' + gDetectionTable() + '</div>',
      },

      // ===== Slide 10: Cột K (D score) =====
      {
        tag: 'K', title: 'Cột K — Phát hiện ra (D)',
        body:
          '<div class="g-block"><h5>🔎 Cách hiểu</h5>'
          + '<p>Điểm <b>D (Detection)</b> đánh giá <b>khả năng phát hiện</b> dạng hỏng hóc hoặc nguyên nhân <em>trước khi</em> sản phẩm lỗi đến công đoạn sau hoặc tay khách hàng. D = 1 chắc chắn phát hiện, D = 10 gần như không phát hiện được.</p></div>'
          + '<div class="g-block"><h5>✍️ Cách làm — chọn đúng "cơ hội phát hiện"</h5>'
          + '<p>Trước khi chấm điểm, xác định dạng hỏng được phát hiện <b>ở đâu</b> — đây là cột "Cơ hội phát hiện ra" trong bảng tiêu chuẩn:</p><ul>'
          + '<li>Phát hiện <b>ngay tại công đoạn đang làm</b> → chấm theo dòng <b>“tại hiện trường phát sinh”</b> (điểm D thấp hơn — lẻ: 3, 5, 7).</li>'
          + '<li>Phát hiện <b>ở công đoạn sau / sau khi rời công đoạn</b> → chấm theo dòng <b>“sau khi kết thúc gia công”</b> (điểm D cao hơn — chẵn: 4, 6, 8).</li>'
          + '</ul></div>'
          + '<div class="g-block"><h5>✍️ Mức điểm theo phương pháp kiểm soát</h5><ul>'
          + '<li>Kiểm tra bằng <b>giác quan</b> (mắt nhìn, tay sờ, đếm số) → D = 7–9 vì độ tin cậy thấp.</li>'
          + '<li>Kiểm tra bằng <b>dưỡng đo / dụng cụ thủ công</b> (thước, caliper, GO/NO-GO) → D = 5–7.</li>'
          + '<li>Kiểm tra <b>tự động / cảm biến</b> chặn đứng → D = 3–4.</li>'
          + '<li>Quản lý tự động phát hiện <b>lỗi (nguyên nhân)</b>, dự phòng sản xuất hàng lỗi → D = 2.</li>'
          + '<li><b>Poka-yoke</b> (thiết kế chống lỗi, không thể xảy ra/lọt qua) → D = 1.</li>'
          + '</ul></div>'
          + '<div class="guide-note">💡 <b>Trường hợp đặc biệt S ≤ 6 và D ≥ 7:</b> Nếu D cao do bản chất phương pháp kiểm tra <b>bằng giác quan</b> và không thể thay thế bằng thiết bị tự động (VD: kiểm tra màu sắc, ngoại quan phức tạp), tình trạng này có thể chấp nhận — không bắt buộc phải có biện pháp đề xuất để hạ D.</div>'
          + '<div class="g-block"><h5>📊 Bảng tiêu chuẩn đánh giá phát hiện</h5>' + gDetectionTable() + '</div>',
      },

      // ===== Trang 12: Cột L (RPN) =====
      {
        tag: 'L', title: 'Cột L — RPN (Chỉ số ưu tiên rủi ro)',
        body:
          '<div class="g-block"><h5>🔎 Cách hiểu</h5>'
          + '<p>RPN = <b>S × O × D</b>. Giá trị càng cao thì tổ hợp rủi ro càng lớn.</p>'
          + '<div class="guide-rpn-formula"><span class="rpn-box">S</span><span class="rpn-op">×</span><span class="rpn-box">O</span><span class="rpn-op">×</span><span class="rpn-box">D</span><span class="rpn-op">=</span><span class="rpn-box rpn-result">RPN</span></div>'
          + '<p style="color:var(--muted);margin-top:8px"><b>Ví dụ G823-00:</b> S=5, O=3, D=4 → RPN = 60</p></div>'
          + '<div class="guide-note">⚠ <b>RPN không phải tiêu chí duy nhất</b> để quyết định có cần biện pháp đề xuất. Việc thực hiện biện pháp được xét theo <b>tiêu chuẩn kết hợp S và D</b> của công ty (xem trang cột M), không chỉ dựa vào RPN. Ô biện pháp đề xuất trong tool sẽ <b>tô viền đỏ</b> khi cần hành động.</div>'
          + '<div class="g-block"><p>RPN <b>tự động tính</b> trong tool ngay khi nhập đủ S, O, D. RPN sau cải tiến (RPN′) cũng tự tính từ S′, O′, D′.</p></div>',
      },

      // ===== Slide 12: Cột M (Biện pháp) =====
      {
        tag: 'M', title: 'Cột M — Biện pháp đề xuất & Kết quả xử lý',
        body:
          '<div class="g-block"><h5>🔎 Cách hiểu</h5>'
          + '<p>Cột M ghi <b>biện pháp cải tiến</b> nhằm giảm S, O hoặc D. Kết quả sau khi thực hiện được ghi vào cột O (biện pháp đã làm) và cột P/Q/R/S (S′/O′/D′/RPN′).</p></div>'
          + '<div class="g-block"><h5>✍️ Tiêu chuẩn thực hiện biện pháp đề xuất</h5>'
          + '<p>Dựa vào <b>bảng tiêu chuẩn</b> dưới đây để xác định nội dung nào <b>bắt buộc</b> phải thực hiện biện pháp đề xuất (xét theo từng yếu tố S, O, D — <b>không</b> chỉ dựa vào RPN):</p>'
          + gActionTable()
          + '<p style="font-size:12px;color:var(--muted);margin-top:6px">Trong tool: ô biện pháp đề xuất <b>tô viền đỏ</b> khi hạng mục chạm ngưỡng cần hành động.</p></div>'
          + '<div class="g-block"><h5>✍️ Cách làm</h5><ul>'
          + '<li>Biện pháp phải cụ thể: <em>làm gì, ai làm, thời hạn nào</em>.</li>'
          + '<li>Mục tiêu: giảm O (cải thiện dự phòng, đảm bảo Cp/Cpk ≥ 1,33), giảm D (thêm kiểm soát tự động — đưa D xuống dưới 5), hoặc giảm S (thay đổi thiết kế / công đoạn).</li>'
          + '<li>Sau khi thực hiện: ghi kết quả vào cột O, chấm lại S′/O′/D′ → RPN′ tự tính.</li>'
          + '</ul></div>'
          + '<div class="g-eg"><b>Ví dụ G823-00 — Chiều cao tổng (S=5, O=3, D=6 → RPN=90):</b><br>'
          + '<i>Biện pháp đề xuất:</i> “Bổ sung cảm biến đo chiều cao trong máy hàn, phát hiện 100% tự động ngay tại công đoạn”<br>'
          + '<i>Kết quả:</i> S′=5, O′=3, D′=2 → <b>RPN′ = 30</b></div>'
          + '<div class="guide-note">⚠ <b>Bắt buộc phê duyệt công ty mẹ:</b> Nếu <b>không có biện pháp đề xuất</b> nào để giảm được điểm <b>S, O hoặc D</b> về ngưỡng tiêu chuẩn (ví dụ S ≥ 9 do tính công nghệ thiết kế R&amp;D, không thể thay đổi bằng cải tiến sản xuất), thì phải tổng hợp hạng mục đó và <b>gửi xin phê duyệt từ công ty mẹ</b> trước khi chốt.</div>',
      },
    ];
  }

  // ===================================================================
  // Chế độ 2: CÁCH THAO TÁC TRỰC TIẾP TRÊN WEB
  // ===================================================================
  function gPagesWeb() {
    return [

      // ----- Web slide 1: Quy trình 6 bước -----
      {
        tag: '', title: 'Quy trình thao tác (6 bước)', menu: 'Quy trình 6 bước',
        full: true,
        body:
          '<ul class="guide-steps">'
          + '<li><span class="step-num">1</span><div>'
          + '<strong>Chọn thông tin dự án</strong><br>'
          + 'Ở thanh trên chọn <b>Bộ phận → Sản phẩm → Dây chuyền</b>, rồi nhập <b>Model</b> (VD: <code>G823-00</code>).<br>'
          + '<b style="color:#c00">⚠ Quan trọng:</b> đổi Bộ phận / Sản phẩm / Dây chuyền sẽ tạo P-FMEA <em>mới hoàn toàn</em> — dữ liệu đang nhập bị xóa. Chọn đúng <b>trước</b> khi tải CP.'
          + '</div></li>'
          + '<li><span class="step-num">2</span><div>'
          + '<strong>Tải file Control Plan (.xlsx)</strong><br>'
          + 'Nhấn <b>📂 Tải Control Plan</b> → chọn file CP. Tool tự đọc <em>tất cả sheet</em> và báo số công đoạn tìm thấy.<br>'
          + '<span style="color:var(--muted)">Tên sheet trong CP không ảnh hưởng đến STT/tên công đoạn trong P-FMEA.</span>'
          + '</div></li>'
          + '<li><span class="step-num">3</span><div>'
          + '<strong>Nạp công đoạn vào bảng</strong><br>'
          + 'Nhấn <b>＋ Nạp tất cả công đoạn vào P-FMEA</b> → tool tự dựng bảng và điền sẵn các cột tự động (xem slide “Cột tự động”).'
          + '</div></li>'
          + '<li><span class="step-num">4</span><div>'
          + '<strong>(Tùy chọn) Bật AI hỗ trợ</strong><br>'
          + 'Nhấn <b>🧠 Cung cấp thêm bối cảnh cho AI</b> → dán API key Gemini (miễn phí tại <b>aistudio.google.com/apikey</b>) và mô tả bộ phận / máy móc / quy trình. Key lưu trong trình duyệt, <b>không gửi về server</b>.'
          + '</div></li>'
          + '<li><span class="step-num">5</span><div>'
          + '<strong>Phân tích &amp; điền các cột tự làm</strong><br>'
          + 'Đi từ trái sang phải, điền các cột tool không tự điền được (ảnh hưởng, nguyên nhân, O, dự phòng, phát hiện, D, biện pháp). Xem slide “Cột phải tự phân tích”.'
          + '</div></li>'
          + '<li><span class="step-num">6</span><div>'
          + '<strong>Lưu &amp; Xuất Excel</strong><br>'
          + '<ul>'
          + '<li><b>💾 Lưu</b> → lưu vào trình duyệt (tự phân theo Bộ phận / Sản phẩm / Dây chuyền).</li>'
          + '<li><b>⬇ Xuất Excel (.xlsx)</b> → xuất theo form chuẩn công ty, giữ định dạng A4.</li>'
          + '<li><b>⬇ Sao lưu / ⬆ Nạp sao lưu</b> → tải/khôi phục file JSON ngoài trình duyệt.</li>'
          + '<li><b>Model base</b> → chọn model đã lưu để tham khảo / copy sang model mới.</li>'
          + '</ul>'
          + '</div></li>'
          + '</ul>',
      },

      // ----- Web slide 2: Cần tải gì & chọn gì -----
      {
        tag: '', title: 'Cần tải gì lên & chọn gì', menu: 'Tải gì · Chọn gì',
        full: true,
        body:
          '<div class="g-block"><h5>📤 File cần tải lên</h5><ul>'
          + '<li><b>Control Plan (.xlsx)</b> — bắt buộc. Đây là nguồn dữ liệu chính để tool rút hạng mục chất lượng, dạng hỏng, SC và phương pháp kiểm tra.</li>'
          + '<li><b>File sao lưu (.json)</b> — tùy chọn, chỉ khi muốn khôi phục một P-FMEA đã làm dở trước đó.</li>'
          + '</ul></div>'
          + '<div class="g-block"><h5>🔽 Cần chọn / nhập gì</h5>'
          + '<table class="guide-table"><thead><tr><th>Mục</th><th>Ý nghĩa</th></tr></thead><tbody>'
          + '<tr><td><b>Bộ phận</b></td><td>Phân xưởng đang lập P-FMEA (VD: PRO2). Quyết định cây dữ liệu Sản phẩm/Dây chuyền.</td></tr>'
          + '<tr><td><b>Sản phẩm</b></td><td>Loại sản phẩm (VD: giảm xóc / damper case comp).</td></tr>'
          + '<tr><td><b>Dây chuyền</b></td><td>Dây chuyền cụ thể trong bộ phận.</td></tr>'
          + '<tr><td><b>Model</b></td><td>Mã model (VD: G823-00). Dùng để lưu &amp; tra cứu.</td></tr>'
          + '</tbody></table>'
          + '<p style="color:var(--muted);font-size:12px;margin-top:6px">Bộ ba Bộ phận / Sản phẩm / Dây chuyền là “khóa” phân loại — đổi bộ ba này = mở P-FMEA khác.</p></div>'
          + '<div class="guide-note">💡 Đảm bảo CP đúng định dạng chuẩn (có cột Hạng mục quản lý, Spec/dung sai, Đặc tính đặc thù SC, Phương pháp &amp; tần suất kiểm tra) thì tool mới rút đủ dữ liệu.</div>',
      },

      // ----- Web slide 3: Cột tự động -----
      {
        tag: '', title: 'Chỗ nào TỰ ĐỘNG (tool tự điền)', menu: 'Cột tự động',
        full: true,
        body:
          '<div class="g-block"><p>Sau khi nạp CP, tool <b>tự điền</b> các cột sau (ô nền <b>xám nhạt</b>). Vẫn chỉnh tay được nếu cần:</p>'
          + '<table class="guide-table"><thead><tr><th>Cột</th><th>Tool tự làm gì</th></tr></thead><tbody>'
          + '<tr><td><b>A</b> — Công đoạn / Yêu cầu</td><td>Điền STT, tên công đoạn, chức năng và danh sách <b>yêu cầu</b> từ CP.</td></tr>'
          + '<tr><td><b>B</b> — Dạng hỏng hóc</td><td>Tạo dạng hỏng từ yêu cầu: dung sai <b>2 phía</b> → 2 dạng (“lớn hơn”/“nhỏ hơn”); 1 phía/văn bản → 1 dạng (“không đạt”).</td></tr>'
          + '<tr><td><b>D</b> — Điểm S</td><td><b>Tự điền</b> ngay khi bạn chọn câu kết luận ảnh hưởng ở cột C (theo bảng tiêu chuẩn S).</td></tr>'
          + '<tr><td><b>E</b> — Phân loại (SC)</td><td>Điền ký hiệu đặc tính đặc thù (S, A…) từ cột S.C trong CP. Có SC → mở thêm ô ③ ở cột J.</td></tr>'
          + '<tr><td><b>J</b> — ô ②</td><td>Điền sẵn phương pháp + tần suất kiểm tra của hạng mục (từ CP).</td></tr>'
          + '<tr><td><b>L / S′</b> — RPN</td><td><b>Tự tính</b> S×O×D và RPN′ = S′×O′×D′ ngay khi đủ số.</td></tr>'
          + '</tbody></table></div>'
          + '<div class="guide-note">💡 Nút <b>🔗 Gộp</b> ở cột B: gộp các dạng hỏng cùng nhóm (cùng kiểu kích thước, cùng cách kiểm tra) → ô phát hiện ② tự viết câu chung “Kiểm tra các kích thước … theo tần suất …”.</div>',
      },

      // ----- Web slide 4: Cột tự phân tích -----
      {
        tag: '', title: 'Chỗ nào PHẢI TỰ PHÂN TÍCH', menu: 'Cột phải tự phân tích',
        full: true,
        body:
          '<div class="g-block"><p>Các cột dưới đây cần <b>kiến thức thực tế</b> về dây chuyền — tool/AI chỉ <em>gợi ý</em>, người làm phải quyết định:</p>'
          + '<table class="guide-table"><thead><tr><th>Cột</th><th>Bạn phải làm</th></tr></thead><tbody>'
          + '<tr><td><b>C</b> — Ảnh hưởng</td><td>Phân tích ảnh hưởng đến công đoạn &amp; sản phẩm, rồi <b>chọn câu kết luận</b> từ dropdown (→ S tự điền).</td></tr>'
          + '<tr><td><b>F</b> — Nguyên nhân</td><td>Liệt kê nguyên nhân theo <b>4M</b>, dựa trên điều kiện chế tạo thực tế. Nút <b>✨ AI</b> gợi ý.</td></tr>'
          + '<tr><td><b>G</b> — Lỗi quá khứ</td><td>Điền <b>“o”</b> cho nguyên nhân đã từng phát sinh.</td></tr>'
          + '<tr><td><b>H</b> — Điểm O</td><td>Chấm tần suất phát sinh theo bảng tiêu chuẩn + dữ liệu lỗi thực tế.</td></tr>'
          + '<tr><td><b>I</b> — Dự phòng</td><td>Ghi biện pháp ngăn nguyên nhân, đang áp dụng thực tế tại hiện trường.</td></tr>'
          + '<tr><td><b>J</b> — ô ① &amp; ③</td><td>① cách phát hiện nguyên nhân; ③ kiểm tra đặc biệt cho hạng mục có SC.</td></tr>'
          + '<tr><td><b>K</b> — Điểm D</td><td>Chấm khả năng phát hiện theo bảng tiêu chuẩn (tại hiện trường / sau gia công).</td></tr>'
          + '<tr><td><b>M</b> — Biện pháp</td><td>Đề xuất biện pháp khi chạm ngưỡng (ô viền đỏ); ghi kết quả S′/O′/D′.</td></tr>'
          + '</tbody></table></div>'
          + '<div class="guide-note">💡 Chi tiết cách tư duy từng cột xem ở lựa chọn <b>🧠 Hiểu &amp; tư duy đúng</b> phía trên.</div>',
      },

    ];
  }

  let GUIDE_IDX = 0;
  let GUIDE_PAGES = null;

  function renderGuidePage(i) {
    const pages = GUIDE_PAGES;
    GUIDE_IDX = Math.max(0, Math.min(i, pages.length - 1));
    const pg = pages[GUIDE_IDX];
    const root = $('#guideRoot');
    const titleHTML = '<div class="guide-page-title">'
      + (pg.tag ? '<span class="gpt-tag">' + pg.tag + '</span>' : '')
      + '<span>' + esc(pg.title) + '</span></div>';
    if (pg.full) {
      root.innerHTML = titleHTML + '<div class="guide-right">' + pg.body + '</div>';
    } else {
      root.innerHTML = titleHTML + '<div class="guide-two-col"><div class="guide-left">'
        + gColChip(pg.tag) + '</div><div class="guide-right">' + pg.body + '</div></div>';
    }
    $('#guidePageInfo').textContent = 'Trang ' + (GUIDE_IDX + 1) + '/' + pages.length;
    $('#guideJump').value = String(GUIDE_IDX);
    $('#guidePrev').disabled = GUIDE_IDX === 0;
    $('#guideNext').disabled = GUIDE_IDX === pages.length - 1;
    if (root.scrollIntoView) window.scrollTo({ top: 0 });
  }

  function setupGuide() {
    let currentMode = 'think';
    const jump = $('#guideJump');

    function loadMode(mode) {
      currentMode = mode;
      GUIDE_PAGES = mode === 'web' ? gPagesWeb() : gPagesThink();
      jump.innerHTML = GUIDE_PAGES.map((p, i) =>
        '<option value="' + i + '">' + esc((i === 0 ? '' : (i) + '. ') + (p.menu || ((p.tag ? p.tag + ' — ' : '') + p.title))) + '</option>'
      ).join('');
      renderGuidePage(0);
      document.querySelectorAll('.guide-mode-btn').forEach(b => {
        b.classList.toggle('active', b.dataset.mode === mode);
      });
    }

    jump.addEventListener('change', () => renderGuidePage(parseInt(jump.value, 10) || 0));
    $('#guidePrev').addEventListener('click', () => renderGuidePage(GUIDE_IDX - 1));
    $('#guideNext').addEventListener('click', () => renderGuidePage(GUIDE_IDX + 1));

    document.querySelectorAll('.guide-mode-btn').forEach(b => {
      b.addEventListener('click', () => loadMode(b.dataset.mode));
    });

    loadMode('think');
  }

  document.addEventListener('DOMContentLoaded', init);
})();
