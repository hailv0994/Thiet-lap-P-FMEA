/* =====================================================================
 * app.js — P-FMEA Builder
 * Quản lý state, dựng bảng P-FMEA, chấm điểm S theo tiêu chuẩn, xuất Excel.
 * ===================================================================== */
(function () {
  'use strict';

  // ----------------------------- State -----------------------------
  let UID = 1;
  const uid = (p) => p + (UID++);
  const state = { processes: [] };   // [{id,no,name,func,reqs:[...]}]
  let workbook = null;               // workbook CP đang mở

  const $ = (s) => document.querySelector(s);
  const esc = (s) => (s == null ? '' : String(s))
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');

  // --------------------- Tạo dữ liệu mặc định ----------------------
  function newCause() {
    return { id: uid('c'), cause: '', pastTrouble: '', occurrence: '',
      prevention: '', detectCause: '', detection: '',
      action: '', responsible: '', actionTaken: '', s2: '', o2: '', d2: '' };
  }

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

  function effectCellHTML(p, r) {
    const idx = findSeverityIdx(r);
    const chosen = !!r.effectStdText;
    return `<td class="auto" rowspan="@RS@" data-proc="${p.id}" data-req="${r.id}">
      <div class="effect-cell">
        <textarea data-field="effectAnalysis" rows="2" placeholder="① Tự phân tích ảnh hưởng…">${esc(r.effectAnalysis)}</textarea>
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
        <div class="detect-label">② Phát hiện ra dạng hỏng hóc (tự động từ CP):</div>
        <div class="detect-auto" contenteditable="true" data-field="detectFailureAuto"
             data-proc="${p.id}" data-req="${r.id}">${esc(r.detectFailureAuto)}</div>
      </div></td>`;
  }

  function txtTD(p, r, c, field, ph) {
    return `<td data-proc="${p.id}" data-req="${r.id}" data-cause="${c.id}">
      <div class="cell-edit" contenteditable="true" data-field="${field}" data-ph="${ph || ''}">${esc(c[field])}</div></td>`;
  }

  function numTD(p, r, c, field, idAttr) {
    const id = idAttr ? ` id="${idAttr}"` : '';
    return `<td class="num" data-proc="${p.id}" data-req="${r.id}" data-cause="${c.id}">
      <input${id} class="num-inp" type="number" min="1" max="10" data-field="${field}" value="${esc(c[field])}" /></td>`;
  }

  function render() {
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

          // F — nguyên nhân + nút thêm/xóa
          tr += `<td data-proc="${p.id}" data-req="${r.id}" data-cause="${c.id}">
                   <div class="cell-edit" contenteditable="true" data-field="cause" data-ph="Nguyên nhân ${ci + 1}">${esc(c.cause)}</div>
                   <div class="cause-toolbar">
                     <button class="mini-btn" data-action="add-cause">＋ NN</button>
                     ${rs > 1 ? '<button class="mini-btn danger" data-action="del-cause">✕ NN</button>' : ''}
                   </div></td>`;
          // G phản ánh lỗi quá khứ
          tr += txtTD(p, r, c, 'pastTrouble', 'Lỗi quá khứ');
          // H tần suất O
          tr += numTD(p, r, c, 'occurrence');
          // I dự phòng
          tr += txtTD(p, r, c, 'prevention', 'Quản lý dự phòng (tự phân tích)');
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
  }

  function onChange(e) {
    const el = e.target;
    if (el.dataset && el.dataset.field === 'effectStd') {
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
    }
  }

  function onClick(e) {
    const btn = e.target.closest('[data-action]');
    if (!btn) return;
    const action = btn.dataset.action;
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
      const sel = $('#sheetSelect');
      sel.innerHTML = sheets.map((s) => `<option>${esc(s)}</option>`).join('');
      $('#sheetGroup').hidden = false;
      $('#btnAddProc').hidden = false;
      $('#btnClear').hidden = false;
    };
    reader.readAsArrayBuffer(file);
  }

  function onLoadProc() {
    if (!workbook) return;
    const sheet = $('#sheetSelect').value;
    const res = window.CPParser.parseSheet(workbook, sheet);
    if (res.error) { alert(res.error); return; }
    if (!res.items.length) { alert('Không tìm thấy hạng mục chất lượng nào trong sheet này.'); return; }

    const proc = {
      id: uid('p'),
      no: $('#procNo').value.trim(),
      name: res.processName || sheet,
      func: '',
      reqs: res.items.map(reqFromItem),
    };
    state.processes.push(proc);
    render();
    document.querySelector('.sheet-wrap').scrollIntoView({ behavior: 'smooth' });
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

  // ============================ Khởi tạo ===========================
  function init() {
    buildHeader();
    render();
    $('#fileCP').addEventListener('change', onFile);
    $('#btnLoad').addEventListener('click', onLoadProc);
    $('#btnAddProc').addEventListener('click', onAddProc);
    $('#btnClear').addEventListener('click', onClear);
    $('#btnExport').addEventListener('click', exportXlsx);

    const tbody = $('#fmea tbody');
    tbody.addEventListener('input', onInput);
    tbody.addEventListener('change', onChange);
    tbody.addEventListener('click', onClick);
  }

  document.addEventListener('DOMContentLoaded', init);
})();
