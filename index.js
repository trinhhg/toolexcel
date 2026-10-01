/**
 * VOCAB STUDIO PRO - COMPLETE JAVASCRIPT REBUILD
 * Kiến trúc cột động (Dynamic Columns Engine), Bộ quét Universal Parser,
 * Hỗ trợ TTS Tiếng Trung Giản Thể & Bảo vệ công thức khi xuất file Excel.
 */

// ==========================================
// 1. TRẠNG THÁI ỨNG DỤNG (APPLICATION STATE)
// ==========================================
const App = {
  // Dữ liệu từ vựng (Record Object động)
  records: [
    { id: 'VOC-0001', stt: 1, c_q: 'xin chào', c_inp: '', c_res: '', c_a: '你好', c_mask: 'Lời chào xã giao', c_pinyin: 'nǐ hǎo', c_extra: 'Chào hỏi' },
    { id: 'VOC-0002', stt: 2, c_q: 'cảm ơn', c_inp: '', c_res: '', c_a: '谢谢', c_mask: 'Biểu thị lòng biết ơn', c_pinyin: 'xièxie', c_extra: 'Lịch sự' },
    { id: 'VOC-0003', stt: 3, c_q: 'trường học', c_inp: '', c_res: '', c_a: '学校', c_mask: 'Nơi có thầy cô & bạn bè', c_pinyin: 'xuéxiào', c_extra: 'Địa điểm' },
    { id: 'VOC-0004', stt: 4, c_q: 'giáo viên', c_inp: '', c_res: '', c_a: '老师', c_mask: 'Người truyền đạt kiến thức', c_pinyin: 'lǎoshī', c_extra: 'Nghề nghiệp' },
    { id: 'VOC-0005', stt: 5, c_q: 'học sinh', c_inp: '', c_res: '', c_a: '学生', c_mask: 'Người tiếp thu bài học', c_pinyin: 'xuéshēng', c_extra: 'Người học' }
  ],
  originalRecords: null,

  // Cấu hình Cột ĐỘNG HOÀN TOÀN (Tùy biến 100% trên giao diện)
  columns: [
    { key: 'stt', title: 'STT', role: 'stt', width: 50, readonly: true },
    { key: 'id', title: 'Mã từ', role: 'id', width: 85, readonly: true },
    { key: 'c_q', title: 'Đề bài', role: 'q', width: 160, readonly: false },
    { key: 'c_inp', title: 'Ô làm bài', role: 'input', width: 160, readonly: false },
    { key: 'c_res', title: 'Kết quả', role: 'res', width: 90, readonly: true },
    { key: 'c_a', title: 'Đáp án chuẩn', role: 'a', width: 160, readonly: false },
    { key: 'c_mask', title: 'Ô che mở khóa', role: 'mask', width: 180, readonly: false },
    { key: 'c_pinyin', title: 'Phiên âm / Pinyin', role: 'pinyin', width: 140, readonly: false },
    { key: 'c_extra', title: 'Ghi chú thêm', role: 'extra', width: 140, readonly: false }
  ],

  // Cấu hình ứng dụng
  config: {
    mode: 'practice',          // 'practice' | 'exam'
    maskRule: 'on_correct',    // 'on_correct' (đúng mới hiện) | 'on_submit' (nộp bài mới hiện) | 'always'
    caseInsensitive: true,
    multipleAnswers: true,
    hideIdOnExport: true,      // Ẩn cột ID trong Excel (Mặc định BẬT)
    hideAnswerOnExport: false,
    lockSheet: true,
    sheetPassword: '',
    isExamSubmitted: false,
    ttsLang: 'zh-CN'           // Mặc định Tiếng Trung Giản Thể
  },

  // Quản lý vùng chọn (Bôi đen nhiều hàng / nhiều cột / khối ô)
  selection: {
    type: 'cells',             // 'cells' | 'rows' | 'cols'
    startRow: 0, startCol: 0,
    endRow: 0, endCol: 0,
    activeRow: 0, activeCol: 0
  },
  isDragging: false,
  counterId: 6,

  // Mini game Flashcard & Quiz
  flashcardIdx: 0,
  quizCurrentRow: null,
  quizScore: 0,
  quizStreak: 0
};

// Chuyển số cột sang chữ cái Excel: 0 -> A, 1 -> B, 26 -> AA...
function colLetter(n) {
  let s = '';
  for (n++; n; n = Math.floor((n - 1) / 26)) {
    s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  }
  return s;
}

const $ = id => document.getElementById(id);

// ==========================================
// 2. KHỞI TẠO & RENDER BẢNG TÍNH ĐỘNG
// ==========================================
function initApp() {
  App.originalRecords = JSON.parse(JSON.stringify(App.records));

  bindControlEvents();
  bindSelectionEvents();
  bindModalsEvents();

  renderTable();
  updateScoring();
}

/**
 * Render toàn bộ cấu trúc Table động
 */
function renderTable() {
  const table = $('sheet-table');
  const { columns, records } = App;

  // 1. HEADER (THEAD) - CHO PHÉP ĐỔI TÊN & VAI TRÒ TRỰC TIẾP TẠI MỖI CỘT
  let theadHtml = '<thead><tr><th class="corner-header" title="Chọn toàn bộ bảng">◢</th>';
  columns.forEach((col, cIdx) => {
    const isProtected = ['stt', 'id'].includes(col.role);
    theadHtml += `
      <th class="col-header" data-c="${cIdx}" style="width: ${col.width}px;" title="Bấm để chọn cột ${colLetter(cIdx)}">
        <div class="col-header-box">
          <div class="col-top-row">
            <span class="col-letter">${colLetter(cIdx)}</span>
            ${!isProtected ? `<button class="col-del-btn" data-c="${cIdx}" title="Xóa cột này">✖</button>` : ''}
          </div>
          <input type="text" class="col-name-input" data-c="${cIdx}" value="${escapeHtml(col.title)}" ${isProtected ? 'readonly' : ''}>
          <select class="col-role-select" data-c="${cIdx}" ${isProtected ? 'disabled' : ''}>
            <option value="q" ${col.role === 'q' ? 'selected' : ''}>🎯 Đề bài</option>
            <option value="input" ${col.role === 'input' ? 'selected' : ''}>✏️ Ô làm bài</option>
            <option value="a" ${col.role === 'a' ? 'selected' : ''}>🔑 Đáp án</option>
            <option value="mask" ${col.role === 'mask' ? 'selected' : ''}>👁️ Ô che</option>
            <option value="res" ${col.role === 'res' ? 'selected' : ''}>📊 Kết quả</option>
            <option value="pinyin" ${col.role === 'pinyin' ? 'selected' : ''}>🗣️ Pinyin</option>
            <option value="extra" ${col.role === 'extra' ? 'selected' : ''}>📝 Ghi chú</option>
            ${isProtected ? `<option value="${col.role}" selected>${col.title}</option>` : ''}
          </select>
        </div>
      </th>`;
  });
  theadHtml += '</tr></thead>';

  // 2. NỘI DUNG (TBODY)
  let tbodyHtml = '<tbody>';
  records.forEach((row, rIdx) => {
    tbodyHtml += `<tr><th class="row-header" data-r="${rIdx}" title="Bấm hoặc kéo để chọn hàng ${rIdx + 1}">${rIdx + 1}</th>`;
    columns.forEach((col, cIdx) => {
      let val = '';
      let isEditable = !col.readonly;
      let extraClass = '';

      if (col.role === 'stt') {
        val = rIdx + 1;
      } else if (col.role === 'id') {
        val = row.id;
      } else if (col.role === 'res') {
        val = row.status || '';
        extraClass = 'col-res-cell';
      } else if (col.role === 'input') {
        val = row[col.key] || '';
        extraClass = 'col-input-cell';
      } else if (col.role === 'mask') {
        extraClass = 'col-mask-cell';
        val = getMaskDisplayValue(row, col.key);
        if (isCellMasked(row)) {
          extraClass += ' masked';
          isEditable = false;
        } else {
          extraClass += ' revealed';
        }
      } else {
        val = row[col.key] ?? '';
      }

      tbodyHtml += `
        <td data-r="${rIdx}" data-c="${cIdx}" class="${extraClass}"
            data-status="${getStatusStyle(row.status, col.role)}"
            ${isEditable ? 'contenteditable="true"' : ''}
            style="width: ${col.width}px;">${escapeHtml(String(val))}</td>`;
    });
    tbodyHtml += '</tr>';
  });
  tbodyHtml += '</tbody>';

  table.innerHTML = theadHtml + tbodyHtml;

  bindHeaderActions();
  highlightSelection();
  updateFxBar();
}

function escapeHtml(text) {
  return String(text).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function getStatusStyle(status, role) {
  if (role !== 'input' && role !== 'res') return '';
  if (status === 'ĐÚNG') return 'ok';
  if (status === 'SAI') return 'bad';
  return '';
}

function isCellMasked(row) {
  const { maskRule, mode, isExamSubmitted } = App.config;
  if (maskRule === 'always') return false;
  if (maskRule === 'on_submit') {
    return mode === 'exam' && !isExamSubmitted;
  }
  return row.status !== 'ĐÚNG';
}

function getMaskDisplayValue(row, key) {
  if (isCellMasked(row)) return '🔒 ••••••';
  return row[key] ?? '';
}

// ==========================================
// 3. THAO TÁC CỘT TRỰC TIẾP TRÊN HEADER
// ==========================================
function bindHeaderActions() {
  // Đổi tên cột
  document.querySelectorAll('.col-name-input').forEach(input => {
    input.onchange = e => {
      const c = +e.target.dataset.c;
      App.columns[c].title = e.target.value.trim() || `Cột ${colLetter(c)}`;
      updateFxBar();
    };
  });

  // Đổi vai trò cột tức thì
  document.querySelectorAll('.col-role-select').forEach(select => {
    select.onchange = e => {
      const c = +e.target.dataset.c;
      const newRole = e.target.value;
      App.columns[c].role = newRole;
      App.columns[c].readonly = ['stt', 'id', 'res'].includes(newRole);
      renderTable();
      updateScoring();
    };
  });

  // Nút xóa cột trực tiếp trên từng header
  document.querySelectorAll('.col-del-btn').forEach(btn => {
    btn.onclick = e => {
      e.stopPropagation();
      const c = +btn.dataset.c;
      deleteColumnAtIndex(c);
    };
  });
}

function deleteColumnAtIndex(c) {
  if (App.columns.length <= 3) {
    alert('Bảng cần tối thiểu 3 cột để hoạt động!');
    return;
  }
  App.columns.splice(c, 1);
  App.selection.activeCol = Math.max(0, c - 1);
  renderTable();
  updateScoring();
}

// ==========================================
// 4. BÔI ĐEN NHIỀU HÀNG / NHIỀU CỘT / KHỐI Ô
// ==========================================
function bindSelectionEvents() {
  const table = $('sheet-table');

  table.addEventListener('mousedown', e => {
    const td = e.target.closest('td');
    const thRow = e.target.closest('th.row-header');
    const thCol = e.target.closest('th.col-header');
    const thCorner = e.target.closest('th.corner-header');

    if (e.target.tagName === 'INPUT' || e.target.tagName === 'SELECT' || e.target.tagName === 'BUTTON') {
      return;
    }

    if (td) {
      const r = +td.dataset.r;
      const c = +td.dataset.c;
      App.selection = { type: 'cells', startRow: r, startCol: c, endRow: r, endCol: c, activeRow: r, activeCol: c };
      App.isDragging = true;
      highlightSelection();
      updateFxBar();
    } else if (thRow) {
      const r = +thRow.dataset.r;
      App.selection = { type: 'rows', startRow: r, startCol: 0, endRow: r, endCol: App.columns.length - 1, activeRow: r, activeCol: 0 };
      App.isDragging = true;
      highlightSelection();
      updateFxBar();
    } else if (thCol) {
      const c = +thCol.dataset.c;
      App.selection = { type: 'cols', startRow: 0, startCol: c, endRow: App.records.length - 1, endCol: c, activeRow: 0, activeCol: c };
      App.isDragging = true;
      highlightSelection();
      updateFxBar();
    } else if (thCorner) {
      App.selection = { type: 'cells', startRow: 0, startCol: 0, endRow: App.records.length - 1, endCol: App.columns.length - 1, activeRow: 0, activeCol: 0 };
      highlightSelection();
    }
  });

  table.addEventListener('mouseover', e => {
    if (!App.isDragging) return;
    const td = e.target.closest('td');
    const thRow = e.target.closest('th.row-header');
    const thCol = e.target.closest('th.col-header');

    if (App.selection.type === 'rows' && thRow) {
      App.selection.endRow = +thRow.dataset.r;
      highlightSelection();
    } else if (App.selection.type === 'cols' && thCol) {
      App.selection.endCol = +thCol.dataset.c;
      highlightSelection();
    } else if (App.selection.type === 'cells' && td) {
      App.selection.endRow = +td.dataset.r;
      App.selection.endCol = +td.dataset.c;
      highlightSelection();
    }
  });

  document.addEventListener('mouseup', () => { App.isDragging = false; });

  // Nhập dữ liệu trực tiếp trong ô
  table.addEventListener('input', e => {
    const td = e.target.closest('td');
    if (!td) return;
    const r = +td.dataset.r;
    const c = +td.dataset.c;
    const col = App.columns[c];
    const val = td.textContent.trim();

    if (col && !col.readonly) {
      App.records[r][col.key] = val;
    }

    evaluateRow(App.records[r]);
    updateScoring();
    refreshRowDisplay(r);
    updateFxBar();
  });

  // Phím Delete: Xóa thông minh theo vùng chọn
  document.addEventListener('keydown', e => {
    if ((e.key === 'Delete' || e.key === 'Backspace') && !['INPUT', 'TEXTAREA'].includes(document.activeElement.tagName)) {
      e.preventDefault();
      if (App.selection.type === 'rows') deleteSelectedRows();
      else if (App.selection.type === 'cols') deleteSelectedColumns();
      else clearSelectedCells();
    }
  });

  // Đồng bộ thanh FX
  $('fx-input').addEventListener('input', e => {
    const { activeRow, activeCol } = App.selection;
    const col = App.columns[activeCol];
    if (!col || col.readonly) return;

    const val = e.target.value;
    App.records[activeRow][col.key] = val;

    const td = getCell(activeRow, activeCol);
    if (td) td.textContent = val;

    evaluateRow(App.records[activeRow]);
    updateScoring();
    refreshRowDisplay(activeRow);
  });
}

function getCell(r, c) {
  return document.querySelector(`td[data-r="${r}"][data-c="${c}"]`);
}

function highlightSelection() {
  document.querySelectorAll('.cell-selected, .cell-focus, th.row-selected, th.col-selected').forEach(el => {
    el.classList.remove('cell-selected', 'cell-focus', 'row-selected', 'col-selected');
  });

  const { startRow, startCol, endRow, endCol, activeRow, activeCol, type } = App.selection;
  const minR = Math.min(startRow, endRow);
  const maxR = Math.max(startRow, endRow);
  const minC = Math.min(startCol, endCol);
  const maxC = Math.max(startCol, endCol);

  for (let r = minR; r <= maxR; r++) {
    for (let c = minC; c <= maxC; c++) {
      const td = getCell(r, c);
      if (td) {
        td.classList.add('cell-selected');
        if (r === activeRow && c === activeCol) td.classList.add('cell-focus');
      }
    }
    const thRow = document.querySelector(`th.row-header[data-r="${r}"]`);
    if (thRow && (type === 'rows' || type === 'cells')) thRow.classList.add('row-selected');
  }

  for (let c = minC; c <= maxC; c++) {
    const thCol = document.querySelector(`th.col-header[data-c="${c}"]`);
    if (thCol && (type === 'cols' || type === 'cells')) thCol.classList.add('col-selected');
  }
}

function updateFxBar() {
  const { activeRow, activeCol } = App.selection;
  const coord = `${colLetter(activeCol)}${activeRow + 1}`;
  $('active-cell-coord').textContent = coord;

  const col = App.columns[activeCol];
  const row = App.records[activeRow];
  let val = '';

  if (col) {
    $('active-cell-role-badge').textContent = `[${col.title}]`;
    if (row) {
      if (col.role === 'stt') val = activeRow + 1;
      else if (col.role === 'id') val = row.id;
      else if (col.role === 'res') val = row.status || '';
      else val = row[col.key] ?? '';
    }
  }

  $('fx-input').value = val;
  $('fx-input').disabled = col ? col.readonly : true;
}

function refreshRowDisplay(r) {
  const row = App.records[r];
  App.columns.forEach((col, c) => {
    const td = getCell(r, c);
    if (!td) return;
    if (col.role === 'res') {
      td.textContent = row.status || '';
      td.setAttribute('data-status', getStatusStyle(row.status, 'res'));
    } else if (col.role === 'input') {
      td.setAttribute('data-status', getStatusStyle(row.status, 'input'));
    } else if (col.role === 'mask') {
      td.textContent = getMaskDisplayValue(row, col.key);
      const masked = isCellMasked(row);
      td.className = `col-mask-cell ${masked ? 'masked' : 'revealed'}`;
      td.contentEditable = !masked && !col.readonly;
    }
  });
}

function clearSelectedCells() {
  const { startRow, startCol, endRow, endCol } = App.selection;
  const minR = Math.min(startRow, endRow);
  const maxR = Math.max(startRow, endRow);
  const minC = Math.min(startCol, endCol);
  const maxC = Math.max(startCol, endCol);

  for (let r = minR; r <= maxR; r++) {
    for (let c = minC; c <= maxC; c++) {
      const col = App.columns[c];
      if (col && !col.readonly) {
        App.records[r][col.key] = '';
        const td = getCell(r, c);
        if (td) td.textContent = '';
      }
    }
    evaluateRow(App.records[r]);
    refreshRowDisplay(r);
  }
  updateScoring();
  updateFxBar();
}

// ==========================================
// 5. CHẤM ĐIỂM & ĐÁNH GIÁ KẾT QUẢ
// ==========================================
function evaluateRow(row) {
  const colInp = App.columns.find(c => c.role === 'input');
  const colA = App.columns.find(c => c.role === 'a');

  if (!colInp || !colA) {
    row.status = '';
    return;
  }

  const input = (row[colInp.key] || '').trim();
  const answer = (row[colA.key] || '').trim();

  if (App.config.mode === 'exam' && !App.config.isExamSubmitted) {
    row.status = '';
    return;
  }

  if (!input) {
    row.status = '';
    return;
  }

  const { caseInsensitive, multipleAnswers } = App.config;
  const normalize = s => caseInsensitive ? s.toLowerCase().trim() : s.trim();
  const userAns = normalize(input);

  if (multipleAnswers) {
    const valids = answer.split('|').map(normalize);
    row.status = valids.includes(userAns) ? 'ĐÚNG' : 'SAI';
  } else {
    row.status = (userAns === normalize(answer)) ? 'ĐÚNG' : 'SAI';
  }
}

function updateScoring() {
  App.records.forEach(row => evaluateRow(row));

  const colA = App.columns.find(c => c.role === 'a');
  const colInp = App.columns.find(c => c.role === 'input');

  let correct = 0, wrong = 0, answered = 0, total = 0;

  App.records.forEach(r => {
    if (colA && (r[colA.key] || '').trim()) total++;
    if (colInp && (r[colInp.key] || '').trim()) answered++;
    if (r.status === 'ĐÚNG') correct++;
    else if (r.status === 'SAI') wrong++;
  });

  const pending = total - correct - wrong;
  const rate = total ? Math.round((correct / total) * 100) : 0;
  const score10 = total ? ((correct / total) * 10).toFixed(2) : '0.00';
  const progressPercent = total ? Math.min(100, Math.round((answered / total) * 100)) : 0;

  $('statCorrect').textContent = correct;
  $('statWrong').textContent = wrong;
  $('statPending').textContent = pending;
  $('statScore').textContent = score10;
  $('progressText').textContent = `${answered}/${total} câu (${progressPercent}%)`;
  $('progressBarFill').style.width = `${progressPercent}%`;

  const badge = $('examStatusBadge');
  if (App.config.mode === 'exam') {
    badge.textContent = App.config.isExamSubmitted ? 'Đã nộp bài' : 'Đang làm bài thi';
    badge.style.background = App.config.isExamSubmitted ? '#dcfce7' : '#fef08a';
    badge.style.color = App.config.isExamSubmitted ? '#166534' : '#854d0e';
    $('examActionArea').style.display = 'flex';
  } else {
    badge.textContent = 'Đang luyện tập';
    badge.style.background = '#e2e8f0';
    badge.style.color = '#475569';
    $('examActionArea').style.display = 'none';
  }
}

// ==========================================
// 6. UNIVERSAL MULTI-LANGUAGE PARSER
// ==========================================
let currentDelim = 'auto';
let parsedImportData = [];

function detectDelimiter(text) {
  const sample = text.split(/\r?\n/).slice(0, 15).filter(l => l.trim());
  const counts = { '\t': 0, ',': 0, ';': 0, '|': 0, '-': 0, ':': 0 };

  sample.forEach(l => {
    if (l.includes('\t')) counts['\t']++;
    if (l.includes(',')) counts[',']++;
    if (l.includes(';')) counts[';']++;
    if (l.includes('|')) counts['|']++;
    if (/\s+[-–—]\s+/.test(l)) counts['-']++;
    if (/\s*[:：]\s*/.test(l)) counts[':']++;
  });

  let best = '\t', max = 0;
  for (const [d, c] of Object.entries(counts)) {
    if (c > max) { max = c; best = d; }
  }
  return max > 0 ? best : '\t';
}

function parseTextLines(text, delimMode) {
  const lines = text.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  if (!lines.length) return [];

  const delim = delimMode === 'auto' ? detectDelimiter(text) : delimMode;

  return lines.map(line => {
    let parts = [];
    if (delim === 'space') {
      parts = line.split(/\s{2,}|\t/);
      if (parts.length === 1) parts = line.split(/\s+/);
    } else if (delim === '-') {
      parts = line.split(/\s+[-–—]\s+/);
    } else if (delim === ':') {
      parts = line.split(/\s*[:：]\s*/);
    } else if (delim === ',') {
      const regex = /(?:^|,)(?:"([^"]*(?:""[^"]*)*)"|([^",]*))/g;
      let match;
      while ((match = regex.exec(line))) {
        if (match[0] === '' && regex.lastIndex >= line.length) break;
        let v = match[1] ? match[1].replace(/""/g, '"') : match[2];
        parts.push(v ? v.trim() : '');
      }
    } else {
      parts = line.split(delim);
    }
    return parts.map(p => p.trim());
  });
}

function analyzeImportText() {
  const text = $('importTextarea').value;
  parsedImportData = parseTextLines(text, currentDelim);

  if (!parsedImportData.length) {
    alert('Vui lòng dán dữ liệu từ vựng vào ô!');
    return;
  }

  const maxCols = Math.max(...parsedImportData.map(r => r.length));
  const previewTable = $('previewTable');

  // Gợi ý vai trò cột thông minh theo dữ liệu thực tế
  const suggestedRoles = [];
  for (let c = 0; c < maxCols; c++) {
    const samples = parsedImportData.slice(0, 10).map(r => r[c] || '');
    const isChinese = samples.filter(s => /[\u4e00-\u9fff]/.test(s)).length >= samples.length * 0.4;
    const isPinyin = samples.filter(s => /[āáǎàēéěèīíǐìōóǒòūúǔùǖǘǚǜü\/\[\]]/.test(s)).length >= samples.length * 0.4;

    if (c === 0) suggestedRoles.push('q');
    else if (isChinese || c === 1) suggestedRoles.push('a');
    else if (isPinyin) suggestedRoles.push('pinyin');
    else if (c === 2) suggestedRoles.push('mask');
    else suggestedRoles.push('extra');
  }

  let thead = '<thead><tr>';
  for (let c = 0; c < maxCols; c++) {
    thead += `
      <th>
        <select class="preview-role-select" data-col="${c}">
          <option value="q" ${suggestedRoles[c] === 'q' ? 'selected' : ''}>🎯 Đề bài</option>
          <option value="a" ${suggestedRoles[c] === 'a' ? 'selected' : ''}>🔑 Đáp án</option>
          <option value="mask" ${suggestedRoles[c] === 'mask' ? 'selected' : ''}>👁️ Ô che</option>
          <option value="pinyin" ${suggestedRoles[c] === 'pinyin' ? 'selected' : ''}>🗣️ Pinyin / Phiên âm</option>
          <option value="extra" ${suggestedRoles[c] === 'extra' ? 'selected' : ''}>📝 Ghi chú / Khác</option>
          <option value="ignore">❌ Bỏ qua</option>
        </select>
      </th>`;
  }
  thead += '</tr></thead><tbody>';

  parsedImportData.slice(0, 6).forEach(row => {
    thead += '<tr>';
    for (let c = 0; c < maxCols; c++) {
      thead += `<td>${escapeHtml(row[c] || '')}</td>`;
    }
    thead += '</tr>';
  });
  thead += '</tbody>';

  previewTable.innerHTML = thead;
  $('previewArea').style.display = 'block';
  $('btnConfirmImport').style.display = 'inline-flex';
}

function confirmImport() {
  const selects = document.querySelectorAll('.preview-role-select');
  const mappings = Array.from(selects).map(s => s.value);

  // Tạo các cột mới dựa trên cấu trúc vừa nạp
  const newCols = [
    { key: 'stt', title: 'STT', role: 'stt', width: 50, readonly: true },
    { key: 'id', title: 'Mã từ', role: 'id', width: 85, readonly: true }
  ];

  let hasInput = false;
  let hasRes = false;

  mappings.forEach((role, idx) => {
    if (role === 'ignore') return;
    const key = `col_${idx}`;
    let title = 'Cột ' + (idx + 1);
    if (role === 'q') title = 'Đề bài';
    else if (role === 'a') title = 'Đáp án chuẩn';
    else if (role === 'mask') title = 'Ô che mở khóa';
    else if (role === 'pinyin') title = 'Pinyin / Phiên âm';
    else if (role === 'extra') title = 'Ghi chú';

    newCols.push({ key, title, role, width: 160, readonly: false });

    // Tự động chèn Ô làm bài và Kết quả ngay sau Đề bài
    if (role === 'q' && !hasInput) {
      newCols.push({ key: 'c_user_inp', title: 'Ô làm bài', role: 'input', width: 160, readonly: false });
      newCols.push({ key: 'c_user_res', title: 'Kết quả', role: 'res', width: 90, readonly: true });
      hasInput = true;
      hasRes = true;
    }
  });

  // Nếu chưa có ô làm bài thì thêm vào
  if (!hasInput) {
    newCols.splice(3, 0,
      { key: 'c_user_inp', title: 'Ô làm bài', role: 'input', width: 160, readonly: false },
      { key: 'c_user_res', title: 'Kết quả', role: 'res', width: 90, readonly: true }
    );
  }

  // Nạp dữ liệu vào records
  const newRecords = parsedImportData.map((row, idx) => {
    const rec = {
      id: `VOC-${String(App.counterId++).padStart(4, '0')}`,
      stt: idx + 1,
      c_user_inp: '',
      c_user_res: '',
      status: ''
    };

    row.forEach((val, cIdx) => {
      const role = mappings[cIdx];
      if (role !== 'ignore') {
        rec[`col_${cIdx}`] = val;
      }
    });

    return rec;
  });

  App.columns = newCols;
  App.records = newRecords;
  App.originalRecords = JSON.parse(JSON.stringify(newRecords));

  $('importModal').hidden = true;
  renderTable();
  updateScoring();
}

// ==========================================
// 7. THÊM / XÓA HÀNG CỘT & BẢNG TÍNH
// ==========================================
function addNewRow() {
  const newId = `VOC-${String(App.counterId++).padStart(4, '0')}`;
  const newRow = { id: newId, stt: App.records.length + 1, status: '' };

  App.columns.forEach(col => {
    if (col.key !== 'id' && col.key !== 'stt') {
      newRow[col.key] = '';
    }
  });

  App.records.push(newRow);
  renderTable();
  updateScoring();
}

function deleteSelectedRows() {
  const { startRow, endRow } = App.selection;
  const minR = Math.min(startRow, endRow);
  const maxR = Math.max(startRow, endRow);

  if (App.records.length <= 1) {
    alert('Bảng phải có ít nhất 1 hàng!');
    return;
  }

  App.records.splice(minR, maxR - minR + 1);
  App.records.forEach((r, idx) => r.stt = idx + 1);

  App.selection.startRow = Math.min(minR, App.records.length - 1);
  App.selection.endRow = App.selection.startRow;
  App.selection.activeRow = App.selection.startRow;

  renderTable();
  updateScoring();
}

function addNewColumn() {
  const key = `custom_${Date.now()}`;
  App.columns.push({
    key,
    title: `Cột ${colLetter(App.columns.length)}`,
    role: 'extra',
    width: 150,
    readonly: false
  });
  renderTable();
}

function deleteSelectedColumns() {
  const { startCol, endCol } = App.selection;
  const minC = Math.min(startCol, endCol);
  const maxC = Math.max(startCol, endCol);

  for (let c = minC; c <= maxC; c++) {
    if (['stt', 'id'].includes(App.columns[c]?.role)) {
      alert('Không thể xóa cột STT và Mã từ!');
      return;
    }
  }

  App.columns.splice(minC, maxC - minC + 1);
  App.selection.activeCol = Math.max(0, minC - 1);
  renderTable();
}

function shuffleAllRows() {
  App.records = shuffleArray(App.records);
  App.records.forEach((r, idx) => r.stt = idx + 1);
  renderTable();
  updateScoring();
}

function restoreOriginalOrder() {
  if (!App.originalRecords) return;
  App.records = JSON.parse(JSON.stringify(App.originalRecords));
  renderTable();
  updateScoring();
}

function pickRandomSubset(n) {
  if (!n || n < 1 || n > App.originalRecords.length) {
    alert(`Vui lòng nhập số câu từ 1 đến ${App.originalRecords.length}`);
    return;
  }
  const shuffled = shuffleArray(App.originalRecords);
  App.records = shuffled.slice(0, n);
  App.records.forEach((r, idx) => r.stt = idx + 1);
  renderTable();
  updateScoring();
}

function shuffleArray(arr) {
  const a = [...arr];
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}

// ==========================================
// 8. IMPORT NGƯỢC FILE EXCEL CŨ ĐỂ SỬA
// ==========================================
async function handleImportExcelFile(e) {
  const file = e.target.files[0];
  if (!file) return;

  try {
    const wb = new ExcelJS.Workbook();
    const arrayBuffer = await file.arrayBuffer();
    await wb.xlsx.load(arrayBuffer);

    const ws = wb.worksheets[0];
    if (!ws) {
      alert('File Excel không hợp lệ!');
      return;
    }

    // Tìm dòng header
    let headerRowIdx = 1;
    for (let r = 1; r <= 8; r++) {
      const vals = ws.getRow(r).values;
      if (Array.isArray(vals) && vals.some(v => String(v).includes('Đề bài') || String(v).includes('STT') || String(v).includes('Mã từ'))) {
        headerRowIdx = r;
        break;
      }
    }

    const headerRow = ws.getRow(headerRowIdx);
    const colCount = ws.columnCount || headerRow.cellCount;
    const newCols = [];
    const keyMap = [];

    for (let c = 1; c <= colCount; c++) {
      const title = String(headerRow.getCell(c).value || '').trim();
      if (!title) continue;

      let role = 'extra';
      let key = `col_${c}`;

      if (title.includes('STT')) { role = 'stt'; key = 'stt'; }
      else if (title.includes('Mã từ') || title.includes('ID')) { role = 'id'; key = 'id'; }
      else if (title.includes('Đề bài')) { role = 'q'; key = 'c_q'; }
      else if (title.includes('Ô làm bài') || title.includes('Ô nhập')) { role = 'input'; key = 'c_inp'; }
      else if (title.includes('Kết quả')) { role = 'res'; key = 'c_res'; }
      else if (title.includes('Đáp án')) { role = 'a'; key = 'c_a'; }
      else if (title.includes('Ô che') || title.includes('Gợi ý')) { role = 'mask'; key = 'c_mask'; }
      else if (title.includes('Pinyin') || title.includes('Phiên âm')) { role = 'pinyin'; key = 'c_pinyin'; }

      newCols.push({ key, title, role, width: 150, readonly: ['stt', 'id', 'res'].includes(role) });
      keyMap.push({ colIndex: c, key, role });
    }

    const newRecords = [];
    for (let r = headerRowIdx + 1; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      if (!row.hasValues) continue;

      const rec = { id: `VOC-${String(App.counterId++).padStart(4, '0')}`, stt: newRecords.length + 1, status: '' };
      keyMap.forEach(({ colIndex, key, role }) => {
        let val = row.getCell(colIndex).value;
        if (val && typeof val === 'object') val = val.result || val.text || '';
        if (role === 'id' && val) rec.id = String(val);
        else if (role !== 'res' && role !== 'input') {
          rec[key] = val != null ? String(val).trim() : '';
        }
      });
      newRecords.push(rec);
    }

    App.columns = newCols;
    App.records = newRecords;
    App.originalRecords = JSON.parse(JSON.stringify(newRecords));

    renderTable();
    updateScoring();
    alert(`Đã nạp thành công ${newRecords.length} câu từ file Excel! Bạn có thể tiếp tục chỉnh sửa.`);
  } catch (err) {
    console.error(err);
    alert('Không thể đọc file Excel này! Vui lòng kiểm tra định dạng .xlsx');
  } finally {
    e.target.value = '';
  }
}

// ==========================================
// 9. XUẤT EXCEL (.XLSX) GIỮ NGUYÊN HIỆU LỰC
// ==========================================
async function exportToExcel() {
  const colQIdx = App.columns.findIndex(c => c.role === 'q');
  const colAIdx = App.columns.findIndex(c => c.role === 'a');
  const colInpIdx = App.columns.findIndex(c => c.role === 'input');
  const colResIdx = App.columns.findIndex(c => c.role === 'res');
  const colIdIdx = App.columns.findIndex(c => c.role === 'id');
  const colMaskIdx = App.columns.findIndex(c => c.role === 'mask');

  if (colQIdx === -1 || colAIdx === -1 || colInpIdx === -1 || colResIdx === -1) {
    alert('Vui lòng đảm bảo bảng có đủ các cột: Đề bài, Đáp án, Ô làm bài và Kết quả!');
    return;
  }

  const wb = new ExcelJS.Workbook();
  wb.creator = 'Vocab Studio Pro';
  wb.created = new Date();

  const numVariants = Math.max(1, Math.min(10, +$('numVariants').value || 1));
  const password = $('sheetPassword').value.trim();

  for (let v = 0; v < numVariants; v++) {
    const ws = wb.addWorksheet(numVariants === 1 ? 'Bài tập từ vựng' : `Đề số ${v + 1}`, {
      views: [{ state: 'frozen', ySplit: 5 }]
    });

    let variantRecords = JSON.parse(JSON.stringify(App.records));
    if (v > 0) {
      variantRecords = shuffleArray(variantRecords);
      variantRecords.forEach((r, idx) => r.stt = idx + 1);
    }

    const rowStart = 6;
    const rowEnd = rowStart + variantRecords.length - 1;

    const L_INP = colLetter(colInpIdx);
    const L_RES = colLetter(colResIdx);
    const L_A = colLetter(colAIdx);

    // 1. DASHBOARD TIẾN ĐỘ & ĐIỂM SỐ (HÀNG 1 - 4)
    ws.mergeCells('A1:D1');
    ws.getCell('A1').value = 'BẢNG KẾT QUẢ ĐÁNH GIÁ TỪ VỰNG';
    ws.getCell('A1').font = { bold: true, size: 12, color: { argb: 'FF107C41' } };

    ws.getCell('E1').value = 'Trạng thái:';
    ws.getCell('E1').font = { bold: true };
    const cellStatus = ws.getCell('F1');
    if (App.config.mode === 'exam') {
      cellStatus.value = 'Chưa nộp';
      cellStatus.protection = { locked: false };
      cellStatus.dataValidation = {
        type: 'list',
        allowBlank: false,
        formulae: ['"Chưa nộp,Đã nộp"']
      };
    } else {
      cellStatus.value = 'Luyện tập tự do';
    }
    cellStatus.font = { bold: true, color: { argb: 'FFD97706' } };

    // Hàng 2: Thống kê Đúng/Sai
    ws.getCell('A2').value = '✓ Đúng:';
    ws.getCell('B2').value = { formula: `COUNTIF(${L_RES}${rowStart}:${L_RES}${rowEnd},"ĐÚNG")` };
    ws.getCell('C2').value = '✗ Sai:';
    ws.getCell('D2').value = { formula: `COUNTIF(${L_RES}${rowStart}:${L_RES}${rowEnd},"SAI")` };
    ws.getCell('E2').value = '○ Chưa làm:';
    ws.getCell('F2').value = { formula: `COUNTBLANK(${L_INP}${rowStart}:${L_INP}${rowEnd})` };

    // Hàng 3: Điểm số /10 & Tỷ lệ
    ws.getCell('A3').value = '📝 Tổng câu:';
    ws.getCell('B3').value = { formula: `COUNTA(${L_A}${rowStart}:${L_A}${rowEnd})` };
    ws.getCell('C3').value = '⭐ Điểm /10:';
    ws.getCell('D3').value = { formula: `IF(B3=0,0,ROUND(B2/B3*10,2))` };
    ws.getCell('E3').value = '📈 Tỷ lệ:';
    ws.getCell('F3').value = { formula: `IF(B3=0,0,B2/B3)` };
    ws.getCell('F3').numFmt = '0.0%';

    // Hàng 4: Thanh tiến độ khối tương thích 100% mọi Excel
    ws.getCell('A4').value = 'Đã làm:';
    ws.getCell('B4').value = { formula: `COUNTA(${L_INP}${rowStart}:${L_INP}${rowEnd})` };
    ws.getCell('C4').value = 'Tiến độ:';
    ws.mergeCells('D4:F4');
    ws.getCell('D4').value = {
      formula: `REPT("■",MIN(10,ROUND(B4/MAX(B3,1)*10,0)))&REPT("□",10-MIN(10,ROUND(B4/MAX(B3,1)*10,0)))&" "&TEXT(IF(B3=0,0,B4/B3),"0%")`
    };

    for (let r = 1; r <= 4; r++) {
      for (let c = 1; c <= Math.max(6, App.columns.length); c++) {
        const cell = ws.getCell(r, c);
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
        cell.border = {
          top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
          bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } }
        };
      }
    }

    // 2. HEADER HÀNG 5
    App.columns.forEach((col, cIdx) => {
      const cell = ws.getCell(5, cIdx + 1);
      cell.value = col.title;
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF107C41' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      ws.getColumn(cIdx + 1).width = Math.max(13, Math.round(col.width / 7.2));
    });

    // 3. DỮ LIỆU & CÔNG THỨC HÀNG 6+
    variantRecords.forEach((rec, idx) => {
      const rIdx = rowStart + idx;

      App.columns.forEach((col, cIdx) => {
        const cell = ws.getCell(rIdx, cIdx + 1);

        if (col.role === 'stt') {
          cell.value = idx + 1;
          cell.alignment = { horizontal: 'center' };
        } else if (col.role === 'id') {
          cell.value = rec.id;
          cell.alignment = { horizontal: 'center' };
        } else if (col.role === 'input') {
          cell.value = '';
          cell.protection = { locked: false };
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCFBF2' } };
        } else if (col.role === 'res') {
          const curInp = `${L_INP}${rIdx}`;
          const curA = `${L_A}${rIdx}`;
          const norm = s => App.config.caseInsensitive ? `LOWER(TRIM(${s}))` : `TRIM(${s})`;

          let checkFormula = '';
          if (App.config.multipleAnswers) {
            checkFormula = `ISNUMBER(SEARCH("|"&${norm(curInp)}&"|","|"&SUBSTITUTE(SUBSTITUTE(${norm(curA)}," |","|"),"| ","|")&"|"))`;
          } else {
            checkFormula = App.config.caseInsensitive ? `${norm(curInp)}=${norm(curA)}` : `EXACT(${norm(curInp)},${norm(curA)})`;
          }

          if (App.config.mode === 'exam') {
            cell.value = { formula: `IF(OR(${curInp}="",$F$1<>"Đã nộp"),"",IF(${checkFormula},"ĐÚNG","SAI"))` };
          } else {
            cell.value = { formula: `IF(${curInp}="","",IF(${checkFormula},"ĐÚNG","SAI"))` };
          }
          cell.alignment = { horizontal: 'center' };
          cell.font = { bold: true };
        } else if (col.role === 'mask') {
          // Ô CHE: NHẬP ĐÚNG MỚI HIỆN HOẶC NỘP BÀI MỚI HIỆN
          const curInp = `${L_INP}${rIdx}`;
          const curA = `${L_A}${rIdx}`;
          const rawVal = (rec[col.key] || '').replace(/"/g, '""');

          if (App.config.maskRule === 'on_correct') {
            cell.value = {
              formula: `IF(AND(${curInp}<>"",LOWER(TRIM(${curInp}))=LOWER(TRIM(${curA}))),"${rawVal}","🔒 [Nhập đúng để mở]")`
            };
          } else if (App.config.maskRule === 'on_submit') {
            cell.value = {
              formula: `IF($F$1="Đã nộp","${rawVal}","🔒 [Nộp bài để xem]")`
            };
          } else {
            cell.value = rec[col.key] || '';
          }
        } else {
          cell.value = rec[col.key] || '';
        }

        cell.border = {
          top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
          bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
          left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
          right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
        };
      });
    });

    // 4. CONDITIONAL FORMATTING MÀU XANH / ĐỎ
    const rangeRes = `${L_RES}${rowStart}:${L_RES}${rowEnd}`;
    const rangeInp = `${L_INP}${rowStart}:${L_INP}${rowEnd}`;

    [rangeRes, rangeInp].forEach(ref => {
      ws.addConditionalFormatting({
        ref,
        rules: [
          {
            type: 'expression',
            formulae: [`$${L_RES}${rowStart}="ĐÚNG"`],
            style: {
              fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFDCFCE7' } },
              font: { color: { argb: 'FF166534' }, bold: true }
            }
          },
          {
            type: 'expression',
            formulae: [`$${L_RES}${rowStart}="SAI"`],
            style: {
              fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFEE2E2' } },
              font: { color: { argb: 'FF991B1B' }, bold: true }
            }
          }
        ]
      });
    });

    // 5. ẨN CỘT MÃ TỪ VÀ ĐÁP ÁN
    if (App.config.hideIdOnExport && colIdIdx !== -1) {
      ws.getColumn(colIdIdx + 1).hidden = true;
    }
    if (App.config.hideAnswerOnExport && colAIdx !== -1) {
      ws.getColumn(colAIdx + 1).hidden = true;
    }

    // 6. KHÓA SHEET BẢO VỆ CÔNG THỨC
    if (App.config.lockSheet) {
      await ws.protect(password, {
        selectLockedCells: true,
        selectUnlockedCells: true
      });
    }
  }

  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `De_Tu_Vung_${new Date().toISOString().slice(0, 10)}.xlsx`;
  a.click();
  URL.revokeObjectURL(url);
}

// ==========================================
// 10. PHÁT ÂM TTS (CHUYÊN TIẾNG TRUNG GIẢN THỂ)
// ==========================================
function speakText(text) {
  if (!('speechSynthesis' in window)) {
    alert('Trình duyệt không hỗ trợ Web Speech TTS!');
    return;
  }
  window.speechSynthesis.cancel();
  const clean = text.split('|')[0].trim();
  const utter = new SpeechSynthesisUtterance(clean);
  utter.lang = App.config.ttsLang;
  utter.rate = 0.9; // Tốc độ chuẩn dễ nghe
  window.speechSynthesis.speak(utter);
}

// ==========================================
// 11. THẺ FLASHCARD 3D
// ==========================================
function openFlashcardModal() {
  if (!App.records.length) {
    alert('Chưa có từ vựng nào để luyện!');
    return;
  }
  App.flashcardIdx = 0;
  $('flashcardModal').hidden = false;
  renderFlashcard();
}

function renderFlashcard() {
  const card = App.records[App.flashcardIdx];
  const colQ = App.columns.find(c => c.role === 'q');
  const colA = App.columns.find(c => c.role === 'a');
  const colPinyin = App.columns.find(c => c.role === 'pinyin');
  const colMask = App.columns.find(c => c.role === 'mask');

  $('flashcardInner').classList.remove('is-flipped');
  $('fcCardIndex').textContent = `Thẻ ${App.flashcardIdx + 1} / ${App.records.length}`;

  $('fcFrontText').textContent = colQ ? card[colQ.key] : 'Không có đề';
  $('fcBackText').textContent = colA ? card[colA.key] : '';
  $('fcBackPinyin').textContent = colPinyin ? card[colPinyin.key] : '';
  $('fcBackMask').textContent = colMask ? card[colMask.key] : '';
}

function flipFlashcard() {
  $('flashcardInner').classList.toggle('is-flipped');
}

// ==========================================
// 12. TRẮC NGHIỆM 4 ĐÁP ÁN (QUIZ)
// ==========================================
function openQuizModal() {
  if (App.records.length < 4) {
    alert('Cần tối thiểu 4 từ vựng trong bảng để tạo trắc nghiệm!');
    return;
  }
  App.quizScore = 0;
  App.quizStreak = 0;
  $('quizModal').hidden = false;
  nextQuizQuestion();
}

function nextQuizQuestion() {
  $('btnQuizNext').style.display = 'none';
  $('quizFeedbackArea').style.display = 'none';

  const colQ = App.columns.find(c => c.role === 'q');
  const colA = App.columns.find(c => c.role === 'a');

  const currIdx = Math.floor(Math.random() * App.records.length);
  App.quizCurrentRow = App.records[currIdx];

  $('quizQuestionText').textContent = App.quizCurrentRow[colQ.key];
  $('quizScoreText').textContent = `Điểm: ${App.quizScore}`;
  $('quizStreakText').textContent = `🔥 Chuỗi đúng: ${App.quizStreak}`;

  const otherAnswers = App.records
    .filter((_, idx) => idx !== currIdx)
    .map(r => (r[colA.key] || '').split('|')[0].trim())
    .filter(Boolean);

  const shuffledOthers = shuffleArray(otherAnswers).slice(0, 3);
  const correct = (App.quizCurrentRow[colA.key] || '').split('|')[0].trim();
  const allOpts = shuffleArray([correct, ...shuffledOthers]);

  const grid = $('quizOptionsGrid');
  grid.innerHTML = '';

  allOpts.forEach(opt => {
    const btn = document.createElement('button');
    btn.className = 'quiz-opt-btn';
    btn.textContent = opt;
    btn.onclick = () => handleQuizAnswer(opt, btn, correct);
    grid.appendChild(btn);
  });
}

function handleQuizAnswer(selected, buttonEl, correct) {
  document.querySelectorAll('.quiz-opt-btn').forEach(b => b.disabled = true);
  const isCorrect = selected.toLowerCase() === correct.toLowerCase();
  const fb = $('quizFeedbackArea');
  fb.style.display = 'block';

  if (isCorrect) {
    buttonEl.classList.add('correct');
    App.quizScore += 10;
    App.quizStreak += 1;
    fb.style.background = '#dcfce7';
    fb.style.color = '#166534';
    fb.textContent = '🎉 CHÍNH XÁC!';
  } else {
    buttonEl.classList.add('wrong');
    App.quizStreak = 0;
    document.querySelectorAll('.quiz-opt-btn').forEach(b => {
      if (b.textContent.toLowerCase() === correct.toLowerCase()) b.classList.add('correct');
    });
    fb.style.background = '#fee2e2';
    fb.style.color = '#991b1b';
    fb.textContent = `❌ Chưa đúng! Đáp án là: ${correct}`;
  }

  $('quizScoreText').textContent = `Điểm: ${App.quizScore}`;
  $('quizStreakText').textContent = `🔥 Chuỗi đúng: ${App.quizStreak}`;
  $('btnQuizNext').style.display = 'inline-flex';
}

// ==========================================
// 13. GẮN SỰ KIỆN GIAO DIỆN
// ==========================================
function bindControlEvents() {
  // Nạp & Mở file
  $('btnOpenImportModal').onclick = () => { $('importModal').hidden = false; $('importTextarea').focus(); };
  $('btnTriggerUpload').onclick = () => $('fileImportExcel').click();
  $('fileImportExcel').onchange = handleImportExcelFile;
  $('btnExportExcel').onclick = exportToExcel;

  // Thao tác hàng cột
  $('btnAddRow').onclick = addNewRow;
  $('btnDelRow').onclick = deleteSelectedRows;
  $('btnAddCol').onclick = addNewColumn;
  $('btnDelCol').onclick = deleteSelectedColumns;
  $('btnClearCells').onclick = clearSelectedCells;

  // Trộn & Lấy subset
  $('btnShuffle').onclick = shuffleAllRows;
  $('btnRestore').onclick = restoreOriginalOrder;
  $('btnPickSubset').onclick = () => pickRandomSubset(+$('numPick').value);

  // Chế độ & Quy tắc
  $('examMode').onchange = e => {
    App.config.mode = e.target.value;
    App.config.isExamSubmitted = false;
    $('btnSubmitExam').style.display = 'inline-flex';
    $('btnResetExam').style.display = 'none';
    updateScoring();
    renderTable();
  };
  $('maskRevealRule').onchange = e => {
    App.config.maskRule = e.target.value;
    renderTable();
  };
  $('chkCaseInsensitive').onchange = e => {
    App.config.caseInsensitive = e.target.checked;
    updateScoring();
    renderTable();
  };
  $('chkMultipleAnswers').onchange = e => {
    App.config.multipleAnswers = e.target.checked;
    updateScoring();
    renderTable();
  };
  $('ttsLangSelect').onchange = e => { App.config.ttsLang = e.target.value; };
  $('chkHideId').onchange = e => App.config.hideIdOnExport = e.target.checked;
  $('chkHideAnswer').onchange = e => App.config.hideAnswerOnExport = e.target.checked;

  $('btnToggleLock').onclick = e => {
    App.config.lockSheet = !App.config.lockSheet;
    e.target.textContent = `🔒 Khóa Sheet: ${App.config.lockSheet ? 'BẬT' : 'TẮT'}`;
    e.target.classList.toggle('on', App.config.lockSheet);
  };

  // Nộp bài
  $('btnSubmitExam').onclick = () => {
    App.config.isExamSubmitted = true;
    $('btnSubmitExam').style.display = 'none';
    $('btnResetExam').style.display = 'inline-flex';
    updateScoring();
    renderTable();
  };
  $('btnResetExam').onclick = () => {
    App.config.isExamSubmitted = false;
    const colInp = App.columns.find(c => c.role === 'input');
    if (colInp) App.records.forEach(r => { r[colInp.key] = ''; r.status = ''; });
    $('btnSubmitExam').style.display = 'inline-flex';
    $('btnResetExam').style.display = 'none';
    updateScoring();
    renderTable();
  };

  // Nghe phát âm ô đang chọn
  $('btnSpeakActiveCell').onclick = () => {
    const { activeRow, activeCol } = App.selection;
    const col = App.columns[activeCol];
    const row = App.records[activeRow];
    if (col && row) speakText(row[col.key] || '');
  };

  // Mini games
  $('btnOpenFlashcard').onclick = openFlashcardModal;
  $('btnOpenQuiz').onclick = openQuizModal;
}

function bindModalsEvents() {
  // Modal Import
  $('btnCloseImportModal').onclick = () => $('importModal').hidden = true;
  $('btnCancelImport').onclick = () => $('importModal').hidden = true;
  $('btnAnalyzeText').onclick = analyzeImportText;
  $('btnConfirmImport').onclick = confirmImport;

  document.querySelectorAll('.btn-delimit').forEach(btn => {
    btn.onclick = () => {
      document.querySelectorAll('.btn-delimit').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      currentDelim = btn.dataset.delim;
      if ($('importTextarea').value.trim()) analyzeImportText();
    };
  });

  // Modal Flashcard
  $('btnCloseFlashcardModal').onclick = () => $('flashcardModal').hidden = true;
  $('flashcardWrapper').onclick = flipFlashcard;
  $('btnFcFlip').onclick = flipFlashcard;
  $('btnFcNext').onclick = () => {
    App.flashcardIdx = (App.flashcardIdx + 1) % App.records.length;
    renderFlashcard();
  };
  $('btnFcPrev').onclick = () => {
    App.flashcardIdx = (App.flashcardIdx - 1 + App.records.length) % App.records.length;
    renderFlashcard();
  };
  $('btnFcSpeak').onclick = () => {
    const card = App.records[App.flashcardIdx];
    const colA = App.columns.find(c => c.role === 'a');
    if (card && colA) speakText(card[colA.key]);
  };

  document.addEventListener('keydown', e => {
    if (!$('flashcardModal').hidden) {
      if (e.code === 'Space') { e.preventDefault(); flipFlashcard(); }
      else if (e.key === 'ArrowRight') { App.flashcardIdx = (App.flashcardIdx + 1) % App.records.length; renderFlashcard(); }
      else if (e.key === 'ArrowLeft') { App.flashcardIdx = (App.flashcardIdx - 1 + App.records.length) % App.records.length; renderFlashcard(); }
      else if (e.key === 'p' || e.key === 'P') {
        const card = App.records[App.flashcardIdx];
        const colA = App.columns.find(c => c.role === 'a');
        if (card && colA) speakText(card[colA.key]);
      }
    }
  });

  // Modal Quiz
  $('btnCloseQuizModal').onclick = () => $('quizModal').hidden = true;
  $('btnQuizExit').onclick = () => $('quizModal').hidden = true;
  $('btnQuizNext').onclick = nextQuizQuestion;
  $('btnQuizSpeak').onclick = () => {
    const colQ = App.columns.find(c => c.role === 'q');
    if (App.quizCurrentRow && colQ) speakText(App.quizCurrentRow[colQ.key]);
  };
}

// Khởi chạy khi DOM tải xong
window.addEventListener('DOMContentLoaded', initApp);
