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

  // ---------------- Gộp dạng hỏng hóc (merge group) ----------------
  // Chữ ký so sánh: MỌI cột trừ reqText & failureMode. Chỉ những yêu cầu có chữ ký
  // GIỐNG HỆT nhau (Ảnh hưởng, Nguyên nhân, Dự phòng, Phát hiện ra — kể cả mục
  // kiểm tra + tần suất ở "② tự động từ CP") mới được phép gộp.
  const norm = (v) => String(v == null ? '' : v).trim();
  function reqSig(r) {
    const head = [r.effectAnalysis, r.effectStdText, r.effectScope, r.severity,
      r.classification, r.detectFailureAuto].map(norm).join('|');
    const cs = (r.causes || []).map((c) => [c.category, c.cause, c.pastTrouble,
      c.occurrence, c.prevention, c.detectCause, c.detection, c.action,
      c.responsible, c.actionTaken, c.s2, c.o2, c.d2].map(norm).join('|')).join('§');
    return head + '#' + cs;
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
      r.classification = rep.classification; r.detectFailureAuto = rep.detectFailureAuto;
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

  // Ô cột B (Dạng hỏng hóc): liệt kê tất cả dạng hỏng hóc trong nhóm gộp.
  // Mỗi dòng có số thứ tự (theo yêu cầu ở cột A), nút 🔗 Gộp và 🔓 Tách (nếu đang gộp).
  function fmCellHTML(p, grp) {
    const grouped = grp.length > 1;
    return grp.map(({ r, ri }) => `
      <div class="fm-line" data-proc="${p.id}" data-req="${r.id}">
        <span class="fm-idx">${ri + 1}.</span>
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
      const totalRows = groups.reduce((n, g) => n + (g[0].r.causes.length || 1), 0) || 1;
      let firstProcRow = true;

      groups.forEach((grp) => {
        const r = grp[0].r;            // đại diện nhóm: cung cấp các cột C–S
        const rs = r.causes.length || 1;
        r.causes.forEach((c, ci) => {
          let tr = `<tr class="${firstProcRow ? 'proc-sep' : ''}">`;

          // Cột A — chỉ ở hàng đầu của công đoạn
          if (firstProcRow) {
            tr += `<td rowspan="${totalRows}">${procCellHTML(p)}</td>`;
          }
          // B,C,D,E — chỉ ở hàng đầu của yêu cầu
          if (ci === 0) {
            tr += `<td class="auto" rowspan="${rs}">${fmCellHTML(p, grp)}</td>`;
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
      const r = getReq(pid, rid); if (r) r[field] = val;
      // reqText/failureMode là riêng từng yêu cầu; các cột còn lại là CHUNG -> đồng bộ nhóm
      if (r && r.mergeId && field !== 'reqText' && field !== 'failureMode') syncMergeGroup(getProc(pid), r);
      scheduleAutosave(); return;
    }
    // cấp nguyên nhân (cột CHUNG khi gộp)
    const c = getCause(pid, rid, cid); if (!c) return;
    c[field] = val;
    if (field === 'occurrence' || field === 'detection') refreshCauseRPN(pid, rid, cid);
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
      body = '<div class="ai-pop-sub">Không có dạng hỏng hóc nào <b>giống hệt</b> (Ảnh hưởng · Nguyên nhân · Dự phòng · Phát hiện ra — kể cả mục kiểm tra &amp; tần suất) để gộp.</div>';
    } else {
      body = '<div class="ai-pop-sub">Chọn dạng hỏng hóc giống hệt để gộp chung 1 ô:</div>'
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
      syncMergeGroup(p, r);
      closeMergePop();
      render();
    });
    document.addEventListener('mousedown', onMergeOutside, true);
  }
  // Tách 1 yêu cầu khỏi nhóm gộp (dữ liệu đã được đồng bộ nên giữ nguyên đủ phân tích).
  function unmergeReq(pid, rid) {
    const p = getProc(pid), r = getReq(pid, rid);
    if (!p || !r || !r.mergeId) return;
    const mid = r.mergeId;
    r.mergeId = '';
    const rest = p.reqs.filter((x) => x.mergeId === mid);
    if (rest.length === 1) rest[0].mergeId = ''; // nhóm còn 1 thành viên -> không còn là nhóm
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
  const G_COLS = {
    A: { vi: 'Quy trình / Bước /\nChức năng\n(Hạng mục yêu cầu)', jp: 'プロセス ステップ/機能', ex: '<b>1. Kiểm tra phôi đầu vào</b>\n• Bước 1: Đo đường kính ngoài\n<i>Yêu cầu:</i> Ø20 ±0,1 mm' },
    B: { vi: 'Dạng hỏng hóc\ntiềm ẩn', jp: '潜在的故障モード', ex: 'Đường kính ngoài\nngoài dung sai\n(Ø < 19,9 hoặc > 20,1)' },
    C: { vi: 'Ảnh hưởng của\nhỏng hóc tiềm ẩn', jp: '潜在的故障影響', ex: '① Không lắp được vào\nthân giảm xóc ở công đoạn sau\n② "Một số sản phẩm phải\nsửa ngoài dây chuyền…"' },
    D: { vi: 'Mức độ\nnghiêm trọng (S)', jp: '厳しさ', ex: '5' },
    E: { vi: 'Phân loại\n(Đặc tính đặc thù)', jp: '分類', ex: '◎ (theo bản vẽ)' },
    F: { vi: 'Nguyên nhân\ncủa hỏng hóc', jp: '潜在的故障原因', ex: '<i>Machine:</i> dao tiện mòn\n<i>Method:</i> sai bù dao\n<i>Material:</i> phôi sai cỡ' },
    H: { vi: 'Tần suất\nphát sinh (O)', jp: '発生頻度', ex: '3' },
    I: { vi: 'Quản lý hiện tại\n— Dự phòng', jp: '現行管理 予防', ex: 'Thay dao theo chu kỳ\n500 sản phẩm/lần\n(đang áp dụng)' },
    J: { vi: 'Quản lý hiện tại\n— Phát hiện ra', jp: '現行管理 検出', ex: 'Đo Ø bằng panme\n100% tại công đoạn,\ntheo tiêu chuẩn kiểm tra' },
    K: { vi: 'Phát hiện (D)', jp: '検出', ex: '4' },
    L: { vi: 'RPN', jp: '', ex: 'S×O×D = 5×3×4 = 60' },
    M: { vi: 'Biện pháp đề xuất\n+ Kết quả xử lý', jp: '推奨処置 / 処置結果', ex: 'Lắp cảm biến đo Ø tự động\n→ chấm lại S/O/D sau cải tiến' },
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

  const G_PENDING = '<div class="guide-pending">⚠ <b>Bảng tham khảo.</b> Sẽ thay bằng <b>bảng tiêu chuẩn chính thức của công ty</b> ngay khi có file. Hiện tạm dùng làm định hướng chấm điểm.</div>';

  function gOccurrenceTable() {
    return G_PENDING + '<table class="guide-table"><thead><tr><th>Rank O</th><th>Khả năng xảy ra</th><th>Tỷ lệ phát sinh tham khảo</th></tr></thead><tbody>'
      + '<tr><td class="s-rank s-10">10</td><td>Rất cao — gần như chắc chắn</td><td>≥ 1 lần / 2 sản phẩm</td></tr>'
      + '<tr><td class="s-rank s-9">9</td><td>Rất cao</td><td>≈ 1 / 3</td></tr>'
      + '<tr><td class="s-rank s-8">8</td><td>Cao</td><td>≈ 1 / 8</td></tr>'
      + '<tr><td class="s-rank s-7">7</td><td>Cao — thỉnh thoảng</td><td>≈ 1 / 20</td></tr>'
      + '<tr><td class="s-rank s-6">6</td><td>Trung bình</td><td>≈ 1 / 80</td></tr>'
      + '<tr><td class="s-rank s-5">5</td><td>Trung bình thấp</td><td>≈ 1 / 400</td></tr>'
      + '<tr><td class="s-rank s-4">4</td><td>Tương đối thấp</td><td>≈ 1 / 2.000</td></tr>'
      + '<tr><td class="s-rank s-3">3</td><td>Thấp</td><td>≈ 1 / 15.000</td></tr>'
      + '<tr><td class="s-rank s-2">2</td><td>Rất thấp</td><td>≈ 1 / 150.000</td></tr>'
      + '<tr><td class="s-rank s-1">1</td><td>Hầu như không xảy ra</td><td>≤ 1 / 1.500.000</td></tr>'
      + '</tbody></table>';
  }

  function gDetectionTable() {
    return G_PENDING + '<table class="guide-table"><thead><tr><th>Rank D</th><th>Khả năng phát hiện</th><th>Phương pháp kiểm soát tham khảo</th></tr></thead><tbody>'
      + '<tr><td class="s-rank s-10">10</td><td>Gần như không phát hiện được</td><td>Không kiểm tra / kiểm tra ngẫu nhiên không đáng tin</td></tr>'
      + '<tr><td class="s-rank s-9">9</td><td>Rất khó phát hiện</td><td>Chỉ quan sát bằng mắt thường</td></tr>'
      + '<tr><td class="s-rank s-8">8</td><td>Khó phát hiện</td><td>Kiểm tra thị giác kép / dùng ảnh chuẩn</td></tr>'
      + '<tr><td class="s-rank s-7">7</td><td>Thấp</td><td>Đo thủ công bằng thước/caliper; gá go–no/go đơn giản</td></tr>'
      + '<tr><td class="s-rank s-6">6</td><td>Trung bình thấp</td><td>Đo bằng dụng cụ chuyên dụng thủ công</td></tr>'
      + '<tr><td class="s-rank s-5">5</td><td>Trung bình</td><td>SPC / đo tự động sau công đoạn</td></tr>'
      + '<tr><td class="s-rank s-4">4</td><td>Tương đối cao</td><td>Phát hiện ngay trong công đoạn bằng cảm biến tự động</td></tr>'
      + '<tr><td class="s-rank s-3">3</td><td>Cao</td><td>Kiểm tra 100% tự động sớm trong công đoạn</td></tr>'
      + '<tr><td class="s-rank s-2">2</td><td>Rất cao</td><td>Kiểm tra 100% tự động + cảnh báo; jig kiểm lỗi</td></tr>'
      + '<tr><td class="s-rank s-1">1</td><td>Chắc chắn phát hiện / phòng tránh</td><td>Poka-yoke: không thể xảy ra hoặc không thể lọt qua</td></tr>'
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

  // Định nghĩa các trang hướng dẫn
  function gPages() {
    return [
      // ---------- Trang tổng quan ----------
      {
        tag: '', title: 'Tổng quan format P-FMEA', menu: 'Tổng quan — Format P-FMEA',
        full: true,
        body:
          '<div class="g-block"><p><strong>P-FMEA</strong> (Process Failure Mode and Effects Analysis) là phương pháp phân tích có hệ thống nhằm phát hiện và đánh giá các <em>dạng hỏng hóc tiềm ẩn</em> của quá trình sản xuất, từ đó đưa ra biện pháp kiểm soát và cải tiến phòng ngừa <em>trước khi</em> sản xuất hàng loạt.</p></div>'
          + '<div class="g-block"><h5>📋 Bố cục bảng P-FMEA</h5><p>Bảng gồm các cột từ <b>A</b> đến <b>S</b>. Các trang sau hướng dẫn chi tiết cách hiểu và cách điền cho từng cột.</p>'
          + gFormatTable() + '</div>'
          + '<div class="g-block"><h5>🧭 Cách dùng tài liệu này</h5><ul>'
          + '<li>Chuyển trang bằng <b>‹ Trang trước</b> / <b>Trang sau ›</b>, hoặc chọn nhanh ở ô danh sách phía trên.</li>'
          + '<li>Mỗi trang: <b>bên trái</b> hiển thị đúng cột đang nói tới (kèm ví dụ), <b>bên phải</b> là cách hiểu &amp; cách làm.</li>'
          + '<li>Các cột có nền nhạt trong tool là cột <b>tự động điền</b> từ Control Plan.</li>'
          + '</ul></div>',
      },

      // ---------- A ----------
      {
        tag: 'A', title: 'Quy trình / Bước / Chức năng / Yêu cầu',
        body:
          '<div class="g-block"><h5>🔎 Cách hiểu</h5><p>Cột này xác định <b>công đoạn</b> đang phân tích và <b>yêu cầu chất lượng</b> cần đạt của công đoạn đó.</p></div>'
          + '<div class="g-block"><h5>✍️ Cách làm</h5><ul>'
          + '<li><b>Số thứ tự công đoạn, tên công đoạn và chức năng</b> phải <b>đồng nhất với Quy trình công nghệ (QTCN)</b> — không tự đặt tên khác.</li>'
          + '<li><b>Bước (step):</b> nếu một công đoạn gồm nhiều nguyên công thì <b>mỗi nguyên công là một bước</b> riêng.</li>'
          + '<li><b>Yêu cầu (hạng mục yêu cầu):</b> phải <b>trùng khớp với hạng mục chất lượng trong Control Plan</b> hoặc trong tiêu chuẩn kiểm tra.</li>'
          + '</ul></div>'
          + '<div class="g-eg"><b>Ví dụ:</b> Công đoạn <b>“Tiện thô”</b> gồm 2 nguyên công → tách thành Bước 1 (tiện mặt đầu), Bước 2 (tiện đường kính). Yêu cầu lấy đúng theo CP: <b>Ø20 ±0,1 mm</b>.</div>',
      },

      // ---------- B ----------
      {
        tag: 'B', title: 'Dạng hỏng hóc tiềm ẩn',
        body:
          '<div class="g-block"><h5>🔎 Cách hiểu</h5><p>Dạng hỏng hóc là trạng thái mà công đoạn <b>không đáp ứng được yêu cầu</b> đã nêu ở cột A — tức là <b>phủ định của yêu cầu</b>.</p></div>'
          + '<div class="g-block"><h5>✍️ Cách làm</h5><p>Viết dạng hỏng hóc là <b>phủ định trực tiếp của yêu cầu</b>, nêu rõ ngưỡng/điều kiện không đạt.</p></div>'
          + '<div class="g-eg"><b>Ví dụ:</b> Yêu cầu là <b>“Lực kéo phá hủy min 14 kN”</b> → Dạng hỏng hóc là <b>“Lực kéo phá hủy &lt; 14 kN”</b> (hoặc “không đạt”).</div>'
          + '<div class="g-block"><p class="muted">Trong tool, cột này được <b>tự động đề xuất</b> dạng “&lt;tên hạng mục&gt; không đạt”; bạn chỉnh lại cho sát ngưỡng thực tế.</p></div>',
      },

      // ---------- C ----------
      {
        tag: 'C', title: 'Ảnh hưởng của dạng hỏng hóc',
        body:
          '<div class="g-block"><h5>🔎 Cách hiểu</h5><p>Đây là cột đòi hỏi <b>kiến thức và kinh nghiệm</b> về dây chuyền sản xuất, về lắp ráp sản phẩm và về <b>nhận định nguy hiểm đối với người sử dụng xe</b>. Ảnh hưởng được chia làm hai loại:</p>'
          + '<ul>'
          + '<li><b>Ảnh hưởng đến công đoạn:</b> hiểu là <b>toàn bộ quá trình sản xuất và lắp ráp</b> trước khi hình thành một chiếc xe hoàn chỉnh — bao gồm <b>cả lắp ráp nội bộ nhà máy lẫn lắp ráp tại khách hàng</b> mà mình xuất hàng cho họ.</li>'
          + '<li><b>Ảnh hưởng đến sản phẩm (khách hàng):</b> “sản phẩm” ở đây là <b>sản phẩm cuối cùng</b>. Với ngành sản xuất giảm xóc, sản phẩm cuối cùng là <b>chiếc xe máy</b>, còn khách hàng là <b>người sử dụng xe</b>.</li>'
          + '</ul></div>'
          + '<div class="g-block"><h5>✍️ Cách làm</h5><ul>'
          + '<li>Khi phân tích một dạng hỏng hóc, phải xác định dạng hỏng đó <b>có được ngăn chặn trong suốt quá trình từ sản xuất đến khi lắp lên xe hay không</b>.</li>'
          + '<li>Nếu <b>được ngăn chặn 100%</b> → chỉ phân tích theo hướng <b>ảnh hưởng đến công đoạn</b>.</li>'
          + '<li>Nếu công đoạn <b>không thể ngăn chặn được</b> → mới tiếp tục phân tích <b>ảnh hưởng đến sản phẩm</b> (đến chiếc xe / người dùng).</li>'
          + '<li>Mỗi câu phân tích ảnh hưởng phải <b>gắn với một câu kết luận tương ứng trong Bảng tiêu chuẩn đánh giá mức độ nghiêm trọng</b> (ý ② của cột) — đây là cơ sở để chấm điểm S ở cột D.</li>'
          + '</ul></div>'
          + '<div class="g-block"><h5>📊 Bảng tiêu chuẩn — Ảnh hưởng đến CÔNG ĐOẠN</h5>' + gSeverityTable('process') + '</div>'
          + '<div class="g-block"><h5>📊 Bảng tiêu chuẩn — Ảnh hưởng đến SẢN PHẨM (khách hàng)</h5>' + gSeverityTable('product') + '</div>',
      },

      // ---------- E ----------
      {
        tag: 'E', title: 'Phân loại — Đặc tính đặc thù',
        body:
          '<div class="g-block"><h5>🔎 Cách hiểu</h5><p>Cột này ghi <b>ký hiệu đặc tính đặc thù</b> (Special Characteristic) của hạng mục — ví dụ đặc tính an toàn / đặc tính quan trọng.</p></div>'
          + '<div class="g-block"><h5>✍️ Cách làm</h5><ul>'
          + '<li>Ký hiệu phải <b>khớp đúng với bản vẽ</b> — không tự gán.</li>'
          + '<li>Trong tool, nếu Control Plan có sẵn ký hiệu đặc tính thì cột này được điền tự động; vẫn kiểm tra lại đối chiếu bản vẽ.</li>'
          + '</ul></div>',
      },

      // ---------- D ----------
      {
        tag: 'D', title: 'Mức độ nghiêm trọng (S)',
        body:
          '<div class="g-block"><h5>🔎 Cách hiểu</h5><p>Điểm <b>S</b> thể hiện mức độ nghiêm trọng của ảnh hưởng đã nêu ở cột C. Điểm càng cao thì ảnh hưởng càng nghiêm trọng (thang 1–10).</p></div>'
          + '<div class="g-block"><h5>✍️ Cách làm</h5><ul>'
          + '<li>Chấm điểm <b>theo Bảng tiêu chuẩn đánh giá mức độ nghiêm trọng</b>.</li>'
          + '<li>Khi cột Ảnh hưởng (C) đã <b>gắn câu kết luận lấy từ bảng tiêu chuẩn</b>, thì ở cột này chỉ việc <b>nhập đúng số điểm tương ứng</b> với câu kết luận đó — không tự chấm cảm tính.</li>'
          + '</ul><p class="muted">Trong tool: chọn câu kết luận ở ý ② cột Ảnh hưởng → điểm S tự điền theo Rank.</p></div>'
          + '<div class="g-block"><h5>📊 Bảng tiêu chuẩn — Ảnh hưởng đến CÔNG ĐOẠN</h5>' + gSeverityTable('process') + '</div>'
          + '<div class="g-block"><h5>📊 Bảng tiêu chuẩn — Ảnh hưởng đến SẢN PHẨM (khách hàng)</h5>' + gSeverityTable('product') + '</div>',
      },

      // ---------- F ----------
      {
        tag: 'F', title: 'Nguyên nhân của hỏng hóc',
        body:
          '<div class="g-block"><h5>🔎 Cách hiểu</h5><p>Cột này cũng đòi hỏi <b>kiến thức và kinh nghiệm về dây chuyền sản xuất</b>. Cần tìm ra <b>nguyên nhân gốc rễ</b> gây ra dạng hỏng hóc.</p></div>'
          + '<div class="g-block"><h5>✍️ Cách làm</h5><ul>'
          + '<li>Phân tích <b>theo 4M</b> (Man / Machine / Method / Material) để <b>không bỏ sót nguyên nhân</b>.</li>'
          + '<li>Một dạng hỏng hóc <b>có thể không đủ cả 4 nguyên nhân</b> theo 4M — nhưng vẫn phải <b>tư duy lần lượt theo 4M</b> để rà soát.</li>'
          + '<li>Phải <b>dựa vào các điều kiện chế tạo thực tế của công đoạn đó</b> để phân tích (thông số máy, dụng cụ, vật liệu, thao tác…).</li>'
          + '</ul></div>'
          + '<div class="guide-4m-grid">'
          + '<div class="guide-4m-card m-man"><div class="m-title">👤 Man (Con người)</div><p>Thao tác sai, thiếu kỹ năng, không theo SOP, nhầm lẫn…</p></div>'
          + '<div class="guide-4m-card m-machine"><div class="m-title">⚙️ Machine (Máy móc)</div><p>Dụng cụ mòn, thiết bị trục trặc, cài đặt/điều chỉnh sai…</p></div>'
          + '<div class="guide-4m-card m-method"><div class="m-title">📋 Method (Phương pháp)</div><p>Điều kiện gia công chưa tối ưu, thứ tự thao tác sai, thiếu bước…</p></div>'
          + '<div class="guide-4m-card m-material"><div class="m-title">📦 Material (Vật liệu)</div><p>Phôi/linh kiện đầu vào sai cỡ, kém chất lượng…</p></div>'
          + '</div>',
      },

      // ---------- H ----------
      {
        tag: 'H', title: 'Tần suất phát sinh (O)',
        body:
          '<div class="g-block"><h5>🔎 Cách hiểu</h5><p>Điểm <b>O</b> đánh giá <b>xác suất nguyên nhân xảy ra</b> trong điều kiện sản xuất bình thường (thang 1–10).</p></div>'
          + '<div class="g-block"><h5>✍️ Cách làm</h5><p>Chấm điểm <b>theo Bảng tiêu chuẩn đánh giá tần suất phát sinh</b>, dựa trên dữ liệu lỗi thực tế hoặc kinh nghiệm với các nguyên nhân tương tự.</p></div>'
          + '<div class="g-block"><h5>📊 Bảng tiêu chuẩn đánh giá tần suất phát sinh</h5>' + gOccurrenceTable() + '</div>',
      },

      // ---------- I ----------
      {
        tag: 'I', title: 'Quản lý hiện tại — Dự phòng',
        body:
          '<div class="g-block"><h5>🔎 Cách hiểu</h5><p>Ghi <b>phương pháp quản lý đang được thực hiện để nguyên nhân không xảy ra</b> (phòng ngừa nguyên nhân).</p></div>'
          + '<div class="g-block"><h5>✍️ Cách làm</h5><ul>'
          + '<li>Suy nghĩ: <b>quản lý như thế nào để nguyên nhân đó không xảy ra?</b></li>'
          + '<li>Biện pháp ghi vào phải là biện pháp <b>đang thực sự được áp dụng tại dây chuyền</b> — không ghi biện pháp mong muốn hoặc chưa triển khai.</li>'
          + '</ul></div>'
          + '<div class="g-eg"><b>Ví dụ:</b> Nguyên nhân “dao tiện mòn” → Dự phòng: <b>“Thay dao theo chu kỳ 500 sản phẩm/lần, có ghi nhật ký”</b> (đang áp dụng).</div>',
      },

      // ---------- J ----------
      {
        tag: 'J', title: 'Quản lý hiện tại — Phát hiện ra',
        body:
          '<div class="g-block"><h5>🔎 Cách hiểu</h5><p>Cột này gồm <b>2 ý</b>: biện pháp <b>phát hiện ra nguyên nhân</b> và biện pháp <b>phát hiện ra dạng hỏng hóc</b>.</p></div>'
          + '<div class="g-block"><h5>✍️ Cách làm</h5><ul>'
          + '<li>Phải ghi <b>cụ thể</b>: <b>kiểm tra cái gì — bằng phương pháp gì — tần suất ra sao</b>, và phải <b>khớp với tiêu chuẩn kiểm tra hoặc Control Plan</b>.</li>'
          + '<li>Chủ yếu quan tâm: ảnh hưởng/nguyên nhân được <b>phát hiện ngay tại công đoạn đang phân tích</b> (tại hiện trường phát sinh), hay phát hiện ở <b>công đoạn sau đó</b> (sau khi kết thúc gia công).</li>'
          + '</ul></div>'
          + '<div class="g-eg"><b>Ví dụ:</b> <b>“Đo đường kính bằng panme, kiểm tra 100% tại công đoạn, theo tiêu chuẩn kiểm tra QC-…”.</b></div>'
          + '<div class="g-block"><h5>📊 Bảng tiêu chuẩn đánh giá phát hiện ra</h5>' + gDetectionTable() + '</div>',
      },

      // ---------- L ----------
      {
        tag: 'L', title: 'RPN — Chỉ số ưu tiên rủi ro',
        body:
          '<div class="g-block"><h5>🔎 Cách hiểu</h5><p>RPN là <b>tích của S, O và D</b>.</p>'
          + '<div class="guide-rpn-formula"><span class="rpn-box">S</span><span class="rpn-op">×</span><span class="rpn-box">O</span><span class="rpn-op">×</span><span class="rpn-box">D</span><span class="rpn-op">=</span><span class="rpn-box rpn-result">RPN</span></div></div>'
          + '<div class="g-block"><h5>✍️ Cách làm</h5><p>Trong tool, RPN được <b>tự động tính</b> ngay khi nhập đủ S, O, D. Lưu ý: việc <b>có phải thực hiện biện pháp đề xuất hay không</b> được xét theo <b>tiêu chuẩn riêng</b> (xem trang cột M), <b>không</b> chỉ dựa vào RPN.</p></div>',
      },

      // ---------- M + kết quả xử lý ----------
      {
        tag: 'M', title: 'Biện pháp đề xuất & Kết quả xử lý',
        body:
          '<div class="g-block"><h5>🔎 Cách hiểu</h5><p>Cột này ghi <b>biện pháp cải tiến</b> nhằm giảm điểm S, O hoặc D, kèm <b>kết quả sau khi thực hiện</b> (các cột O–S: biện pháp đã thực hiện và S′/O′/D′/RPN′ sau cải tiến).</p></div>'
          + '<div class="g-block"><h5>✍️ Cách làm</h5><ul>'
          + '<li>Việc thực hiện biện pháp phải áp dụng <b>theo Tiêu chuẩn thực hiện biện pháp đề xuất</b> — <b>KHÔNG</b> áp dụng theo RPN.</li>'
          + '<li>Trong tiêu chuẩn sẽ quy định rõ: với mức <b>S, O, D</b> như thế nào thì <b>bắt buộc phải thực hiện biện pháp</b> để giảm điểm.</li>'
          + '<li>Nếu thực hiện được biện pháp → <b>ghi biện pháp</b>, rồi <b>chấm lại S, O, D sau khi thực hiện</b> vào các ô phía sau (S′/O′/D′ → RPN′ tự tính).</li>'
          + '</ul></div>'
          + '<div class="guide-pending">⚠ <b>Trường hợp đặc biệt — S ≥ 9:</b> nếu điểm <b>S từ 9 trở lên</b> mà do <b>tính công nghệ do bên R&amp;D thiết kế</b>, không thể áp dụng biện pháp thay đổi thiết kế, thì phải <b>tổng hợp hạng mục S đó vào trang tổng hợp</b> và <b>gửi sang công ty mẹ tại Nhật để xin phê duyệt</b>.</div>'
          + '<div class="guide-pending">📄 <b>Bảng “Tiêu chuẩn thực hiện biện pháp đề xuất”</b> sẽ được chèn vào đây khi có file chính thức.</div>',
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
    GUIDE_PAGES = gPages();
    const jump = $('#guideJump');
    jump.innerHTML = GUIDE_PAGES.map((p, i) =>
      '<option value="' + i + '">' + esc((i === 0 ? '' : (i) + '. ') + (p.menu || ((p.tag ? p.tag + ' — ' : '') + p.title))) + '</option>'
    ).join('');
    jump.addEventListener('change', () => renderGuidePage(parseInt(jump.value, 10) || 0));
    $('#guidePrev').addEventListener('click', () => renderGuidePage(GUIDE_IDX - 1));
    $('#guideNext').addEventListener('click', () => renderGuidePage(GUIDE_IDX + 1));
    renderGuidePage(0);
  }

  document.addEventListener('DOMContentLoaded', init);
})();
