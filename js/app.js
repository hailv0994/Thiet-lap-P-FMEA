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
      prevention: '', detectCause: '', detection: '',
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

  function procCellHTML(p) {
    const reqLines = p.reqs.map((r, i) => `
      <div class="req-line" data-proc="${p.id}" data-req="${r.id}">
        <span class="idx">${i + 1}.</span>
        <textarea data-field="reqText" rows="2" placeholder="Yêu cầu (điều kiện chất lượng)">${esc(r.reqText)}</textarea>
        <button class="mini-btn danger" data-action="del-req" title="Xóa yêu cầu này">✕</button>
      </div>`).join('');
    return `<div class="proc-cell" data-proc="${p.id}">
        <div class="proc-head">
          <span class="proc-move">
            <button class="mini-btn" data-action="move-up" title="Lên trên">▲</button>
            <button class="mini-btn" data-action="move-down" title="Xuống dưới">▼</button>
          </span>
          <input data-field="no" class="inp-no" style="width:46px" value="${esc(p.no)}" placeholder="STT" />.
          <input data-field="name" style="width:150px" value="${esc(p.name)}" placeholder="Tên công đoạn" />
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
    return `<td data-proc="${p.id}" data-req="${r.id}" data-cause="${c.id}">
      <div class="detect-cell">
        <div class="detect-label">① Phát hiện ra nguyên nhân (tự phân tích):</div>
        <textarea data-field="detectCause" rows="2" placeholder="…">${esc(c.detectCause)}</textarea>
        ${aiBtn('detectCause')}
        <div class="detect-label">② Phát hiện ra dạng hỏng hóc (tự động từ CP):</div>
        <div class="detect-auto" contenteditable="true" data-field="detectFailureAuto"
             data-proc="${p.id}" data-req="${r.id}">${esc(r.detectFailureAuto)}</div>
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
      const totalRows = p.reqs.reduce((n, r) => n + r.causes.length, 0) || 1;
      let firstProcRow = true;

      p.reqs.forEach((r) => {
        const rs = r.causes.length || 1;
        r.causes.forEach((c, ci) => {
          let tr = `<tr class="${firstProcRow ? 'proc-sep' : ''}">`;

          // Cột A — chỉ ở hàng đầu của công đoạn
          if (firstProcRow) {
            tr += `<td rowspan="${totalRows}">${procCellHTML(p)}</td>`;
          }
          // B,C,D,E — chỉ ở hàng đầu của yêu cầu
          if (ci === 0) {
            tr += `<td class="auto" rowspan="${rs}" data-proc="${p.id}" data-req="${r.id}">
                     <div class="cell-edit" contenteditable="true" data-field="failureMode">${esc(r.failureMode)}</div></td>`;
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
          // M..S
          tr += txtTD(p, r, c, 'action', 'Biện pháp đề xuất');
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
  function refreshReqScores(pid, rid) {
    const r = getReq(pid, rid);
    if (!r) return;
    const sev = $(`#sev-${rid}`);
    if (sev) sev.textContent = r.severity || '';
    r.causes.forEach((c) => {
      const box = $(`#rpn-${c.id}`);
      if (box) box.textContent = rpnOf(r, c);
    });
  }
  function refreshCauseRPN(pid, rid, cid) {
    const r = getReq(pid, rid), c = getCause(pid, rid, cid);
    if (!r || !c) return;
    const box = $(`#rpn-${cid}`);
    if (box) box.textContent = rpnOf(r, c);
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
      const r = getReq(pid, rid); if (r) r[field] = val; return;
    }
    // cấp nguyên nhân
    const c = getCause(pid, rid, cid); if (!c) return;
    c[field] = val;
    if (field === 'occurrence' || field === 'detection') refreshCauseRPN(pid, rid, cid);
    scheduleAutosave();
  }

  function onChange(e) {
    const el = e.target;
    const field = el.dataset && el.dataset.field;
    if (field === 'category') {
      const { pid, rid, cid } = dataset(el);
      const c = getCause(pid, rid, cid); if (c) { c.category = el.value; scheduleAutosave(); }
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
      if (p && p.reqs.length > 1) { p.reqs = p.reqs.filter((r) => r.id !== rid); render(); }
      else if (p) { alert('Mỗi công đoạn cần ít nhất 1 yêu cầu.'); }
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
  function parseAllProcs() {
    const sheets = window.CPParser.listSheets(workbook);
    const procs = [];
    const skipped = [];
    sheets.forEach((sheet) => {
      let res;
      try { res = window.CPParser.parseSheet(workbook, sheet); }
      catch (err) { skipped.push(sheet); return; }
      if (res.error || !res.items.length) { skipped.push(sheet); return; }
      procs.push({
        id: uid('p'),
        no: String(procs.length + 1),
        name: res.processName || sheet,
        func: '',
        reqs: res.items.map(reqFromItem),
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
  function applySnapshot(obj) {
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
    // Đổi meta = đổi NHÃN cho phiên hiện tại; KHÔNG xóa nội dung đang làm.
    // Chỉ khi bảng đang TRỐNG và có dự án đã lưu khớp Model thì mới tự mở ra.
    function onMetaChange() {
      readMetaInputs();
      if (!state.processes.length) {
        const proj = readProjects()[currentKey()];
        if (proj && proj.processes && proj.processes.length) {
          state.processes = proj.processes; reindexUID(); render();
          $('#projSelect').value = currentKey();
        }
      }
      scheduleAutosave();
    }
    $('#mDept').addEventListener('change', () => {
      const dept = $('#mDept').value;
      fillProduct(dept, '');                  // đổ lại sản phẩm theo bộ phận
      fillLine(dept, '', '');                 // reset dây chuyền
      onMetaChange();
    });
    $('#mProduct').addEventListener('change', () => {
      fillLine($('#mDept').value, $('#mProduct').value, ''); // đổ lại dây chuyền
      onMetaChange();
    });
    $('#mLine').addEventListener('change', onMetaChange);
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
  }

  document.addEventListener('DOMContentLoaded', init);
})();
