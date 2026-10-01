/**
 * VOCAB SHEET STUDIO - INDEX.JS
 * Kiến trúc nâng cấp:
 * 1. Bộ quét Universal Multi-Engine Scanner nhận diện chính xác mọi ngôn ngữ & ký tự
 * 2. Hỗ trợ bôi đen nhiều hàng/nhiều cột rồi xóa (Multi-Select & Delete)
 * 3. Ẩn cột Mã từ (ID) khi xuất Excel & Import ngược file XLSX/CSV để chỉnh sửa tiếp
 * 4. Ký hiệu chuẩn hóa không lỗi font/đen trắng trong Excel
 * 5. Cơ chế Ô che (Mask Cell) mở khóa khi gõ đúng hoặc nộp bài
 * 6. Tích hợp chế độ Thẻ Flashcard 3D, Trắc nghiệm 4 câu & Phát âm Web Speech TTS
 */

// ==========================================
// 1. TRẠNG THÁI ỨNG DỤNG (APPLICATION STATE)
// ==========================================
const App = {
  // Danh sách các bản ghi từ vựng (Record Object)
  records: [
    { id: 'VOC-0001', stt: 1, q: 'xin chào', a: 'hello', mask: 'Lời chào xã giao', pinyin: '/həˈləʊ/', extra: 'Chào hỏi', input: '', status: '' },
    { id: 'VOC-0002', stt: 2, q: 'cảm ơn', a: 'thank you | thanks', mask: 'Biểu thị lòng biết ơn', pinyin: '/θæŋk juː/', extra: 'Lịch sự', input: '', status: '' },
    { id: 'VOC-0003', stt: 3, q: 'trường học', a: 'school', mask: 'Nơi có thầy cô & bạn bè', pinyin: '/skuːl/', extra: 'Địa điểm', input: '', status: '' },
    { id: 'VOC-0004', stt: 4, q: 'giáo viên', a: 'teacher', mask: 'Người truyền đạt kiến thức', pinyin: '/ˈtiːtʃə/', extra: 'Nghề nghiệp', input: '', status: '' },
    { id: 'VOC-0005', stt: 5, q: 'học sinh', a: 'student', mask: 'Người tiếp thu bài học', pinyin: '/ˈstjuːdnt/', extra: 'Người học', input: '', status: '' }
  ],
  originalRecords: null, // Sao lưu để khôi phục thứ tự gốc

  // Cấu hình các cột trong bảng tính
  columns: [
    { key: 'stt', title: 'STT', role: 'stt', width: 55, readonly: true },
    { key: 'id', title: 'Mã từ', role: 'id', width: 90, readonly: true },
    { key: 'q', title: 'Đề bài', role: 'q', width: 170, readonly: false },
    { key: 'input', title: 'Ô làm bài', role: 'input', width: 170, readonly: false },
    { key: 'res', title: 'Kết quả', role: 'res', width: 90, readonly: true },
    { key: 'mask', title: 'Ô che / Gợi ý', role: 'mask', width: 170, readonly: false },
    { key: 'a', title: 'Đáp án chuẩn', role: 'a', width: 160, readonly: false },
    { key: 'pinyin', title: 'Phiên âm / Pinyin', role: 'pinyin', width: 140, readonly: false },
    { key: 'extra', title: 'Ghi chú thêm', role: 'extra', width: 150, readonly: false }
  ],

  // Cấu hình làm bài & Xuất Excel
  config: {
    mode: 'practice',          // 'practice' | 'exam'
    maskRule: 'on_correct',    // 'on_correct' | 'on_submit' | 'always'
    caseInsensitive: true,     // Bỏ qua chữ hoa/thường
    multipleAnswers: true,     // Cho phép nhiều đáp án a | b
    hideIdOnExport: true,      // Ẩn cột Mã từ khi xuất Excel (Mặc định BẬT)
    hideAnswerOnExport: false, // Ẩn cột đáp án khi xuất Excel
    lockSheet: true,           // Khóa Sheet Excel
    sheetPassword: '',         // Mật khẩu khóa
    isExamSubmitted: false,    // Đã nộp bài kiểm tra hay chưa
    ttsLang: 'en-US'           // Ngôn ngữ phát âm Text-to-Speech
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

  // Trạng thái cho mini game Flashcard & Quiz
  flashcardIdx: 0,
  quizCurrentRow: null,
  quizScore: 0,
  quizStreak: 0
};

// Hàm chuyển số thứ tự cột sang chữ cái Excel: 0 -> A, 1 -> B, 25 -> Z, 26 -> AA...
function colLetter(n) {
  let s = '';
  for (n++; n; n = Math.floor((n - 1) / 26)) {
    s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  }
  return s;
}

const $ = id => document.getElementById(id);

// ==========================================
// 2. KHỞI TẠO & RENDER BẢNG TÍNH
// ==========================================
function initApp() {
  App.originalRecords = JSON.parse(JSON.stringify(App.records));

  bindToolbarEvents();
  bindSelectionEvents();
  bindModalsEvents();

  renderTable();
  updateScoring();
}

/**
 * Render cấu trúc bảng tính Table (Header + Rows)
 */
function renderTable() {
  const table = $('sheet-table');
  const { columns, records } = App;

  // 1. Tạo Header Thead
  let theadHtml = '<thead><tr><th class="corner-header" title="Chọn toàn bộ bảng">◢</th>';
  columns.forEach((col, cIdx) => {
    const roleBadge = getRoleBadgeName(col.role);
    theadHtml += `
      <th class="col-header" data-c="${cIdx}" style="width: ${col.width}px;" title="Bấm hoặc kéo để chọn cột ${colLetter(cIdx)}">
        ${colLetter(cIdx)}
        <span class="col-role-tag">${roleBadge}</span>
      </th>`;
  });
  theadHtml += '</tr></thead>';

  // 2. Tạo Body Tbody
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
        val = row.input || '';
        extraClass = 'col-input-cell';
      } else if (col.role === 'mask') {
        extraClass = 'col-mask-cell';
        val = getMaskDisplayValue(row, col.key);
        if (isCellMasked(row)) {
          extraClass += ' masked';
          isEditable = false; // Khi đang bị che thì không sửa trực tiếp
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

  highlightSelection();
  updateFxBar();
}

function escapeHtml(text) {
  return String(text).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function getRoleBadgeName(role) {
  switch (role) {
    case 'q': return '🎯 Đề bài';
    case 'a': return '🔑 Đáp án';
    case 'input': return '✏️ Ô nhập';
    case 'mask': return '👁️ Ô che';
    case 'res': return '📊 Kết quả';
    case 'stt': return 'STT';
    case 'id': return 'Mã từ';
    case 'pinyin': return 'Phiên âm';
    default: return 'Ghi chú';
  }
}

function getStatusStyle(status, role) {
  if (role !== 'input' && role !== 'res') return '';
  if (status === 'ĐÚNG') return 'ok';
  if (status === 'SAI') return 'bad';
  return '';
}

// Kiểm tra xem ô che có đang bị che giấu không
function isCellMasked(row) {
  const { maskRule, mode, isExamSubmitted } = App.config;
  if (maskRule === 'always') return false;
  if (maskRule === 'on_submit') {
    return mode === 'exam' && !isExamSubmitted;
  }
  // on_correct: Đúng thì hiện, chưa đúng thì che
  return row.status !== 'ĐÚNG';
}

function getMaskDisplayValue(row, key) {
  if (isCellMasked(row)) {
    return '🔒 ••••••';
  }
  return row[key] ?? '';
}

// ==========================================
// 3. TƯƠNG TÁC CHUỘT & BÔI ĐEN NHIỀU HÀNG/CỘT
// ==========================================
function bindSelectionEvents() {
  const table = $('sheet-table');

  // Chuột nhấn xuống bắt đầu vùng chọn
  table.addEventListener('mousedown', e => {
    const td = e.target.closest('td');
    const thCol = e.target.closest('th.col-header');
    const thRow = e.target.closest('th.row-header');
    const thCorner = e.target.closest('th.corner-header');

    if (td) {
      // Bôi đen các ô (cells)
      const r = +td.dataset.r;
      const c = +td.dataset.c;
      App.selection = {
        type: 'cells',
        startRow: r, startCol: c,
        endRow: r, endCol: c,
        activeRow: r, activeCol: c
      };
      App.isDragging = true;
      highlightSelection();
      updateFxBar();
    } else if (thRow) {
      // Bôi đen cả hàng (rows)
      const r = +thRow.dataset.r;
      App.selection = {
        type: 'rows',
        startRow: r, startCol: 0,
        endRow: r, endCol: App.columns.length - 1,
        activeRow: r, activeCol: 0
      };
      App.isDragging = true;
      highlightSelection();
      updateFxBar();
    } else if (thCol) {
      // Bôi đen cả cột (cols)
      const c = +thCol.dataset.c;
      App.selection = {
        type: 'cols',
        startRow: 0, startCol: c,
        endRow: App.records.length - 1, endCol: c,
        activeRow: 0, activeCol: c
      };
      App.isDragging = true;
      highlightSelection();
      updateFxBar();
    } else if (thCorner) {
      // Chọn tất cả
      App.selection = {
        type: 'cells',
        startRow: 0, startCol: 0,
        endRow: App.records.length - 1, endCol: App.columns.length - 1,
        activeRow: 0, activeCol: 0
      };
      highlightSelection();
    }
  });

  // Kéo chuột bôi đen nhiều hàng / nhiều cột liên tục
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

  document.addEventListener('mouseup', () => {
    App.isDragging = false;
  });

  // Nhập trực tiếp vào ô
  table.addEventListener('input', e => {
    const td = e.target.closest('td');
    if (!td) return;
    const r = +td.dataset.r;
    const c = +td.dataset.c;
    const col = App.columns[c];
    const val = td.textContent.trim();

    if (col.role === 'input') {
      App.records[r].input = val;
    } else if (col.key && !col.readonly) {
      App.records[r][col.key] = val;
    }

    evaluateRow(App.records[r]);
    updateScoring();
    refreshRowDisplay(r);
    updateFxBar();
  });

  // Phím Delete / Backspace: Tự động xóa theo phạm vi chọn
  document.addEventListener('keydown', e => {
    if ((e.key === 'Delete' || e.key === 'Backspace') && !['INPUT', 'TEXTAREA'].includes(document.activeElement.tagName)) {
      e.preventDefault();
      if (App.selection.type === 'rows') {
        deleteSelectedRows();
      } else if (App.selection.type === 'cols') {
        deleteSelectedColumns();
      } else {
        clearSelectedCells();
      }
    }
  });

  // Đồng bộ thanh FX
  $('fx-input').addEventListener('input', e => {
    const { activeRow, activeCol } = App.selection;
    const col = App.columns[activeCol];
    if (!col || col.readonly) return;

    const val = e.target.value;
    if (col.role === 'input') {
      App.records[activeRow].input = val;
    } else {
      App.records[activeRow][col.key] = val;
    }

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

/**
 * Hiển thị bôi đen vùng chọn
 */
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
        if (r === activeRow && c === activeCol) {
          td.classList.add('cell-focus');
        }
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
      else if (col.role === 'input') val = row.input || '';
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

// Xóa trắng dữ liệu ô đang bôi đen
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
        if (col.role === 'input') App.records[r].input = '';
        else App.records[r][col.key] = '';

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
// 4. CHẤM ĐIỂM & ĐÁNH GIÁ KẾT QUẢ
// ==========================================
function evaluateRow(row) {
  const input = (row.input || '').trim();
  const answer = (row.a || '').trim();

  // Chế độ kiểm tra và chưa nộp bài -> Không hiện kết quả
  if (App.config.mode === 'exam' && !App.config.isExamSubmitted) {
    row.status = '';
    return;
  }

  // Chưa nhập -> Chưa làm
  if (!input) {
    row.status = '';
    return;
  }

  const { caseInsensitive, multipleAnswers } = App.config;
  const normalize = str => caseInsensitive ? str.toLowerCase().trim() : str.trim();
  const userAns = normalize(input);

  if (multipleAnswers) {
    const validAnswers = answer.split('|').map(normalize);
    row.status = validAnswers.includes(userAns) ? 'ĐÚNG' : 'SAI';
  } else {
    row.status = (userAns === normalize(answer)) ? 'ĐÚNG' : 'SAI';
  }
}

function updateScoring() {
  App.records.forEach(row => evaluateRow(row));

  let correct = 0, wrong = 0, answered = 0, totalWithAnswer = 0;

  App.records.forEach(r => {
    if ((r.a || '').trim()) totalWithAnswer++;
    if ((r.input || '').trim()) answered++;
    if (r.status === 'ĐÚNG') correct++;
    else if (r.status === 'SAI') wrong++;
  });

  const pending = totalWithAnswer - correct - wrong;
  const rate = totalWithAnswer ? Math.round((correct / totalWithAnswer) * 100) : 0;
  const score10 = totalWithAnswer ? ((correct / totalWithAnswer) * 10).toFixed(2) : '0.00';
  const progressPercent = totalWithAnswer ? Math.min(100, Math.round((answered / totalWithAnswer) * 100)) : 0;

  $('statCorrect').textContent = correct;
  $('statWrong').textContent = wrong;
  $('statPending').textContent = pending;
  $('statTotal').textContent = totalWithAnswer;
  $('statScore').textContent = score10;
  $('statRate').textContent = `${rate}%`;

  $('progressText').textContent = `${answered}/${totalWithAnswer} câu (${progressPercent}%)`;
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
// 5. BỘ QUÉT UNIVERSAL MULTI-ENGINE PARSER
// ==========================================
let currentDelimiterMode = 'auto';
let parsedImportRows = [];

function detectBestDelimiter(text) {
  const sampleLines = text.split(/\r?\n/).slice(0, 15).filter(l => l.trim());
  const counts = { '\t': 0, ',': 0, ';': 0, '|': 0, '-': 0, ':': 0 };

  sampleLines.forEach(l => {
    if (l.includes('\t')) counts['\t']++;
    if (l.includes(',')) counts[',']++;
    if (l.includes(';')) counts[';']++;
    if (l.includes('|')) counts['|']++;
    if (l.includes(' - ') || l.includes(' – ') || l.includes(' — ')) counts['-']++;
    if (l.includes(' : ') || l.includes(':')) counts[':']++;
  });

  let best = '\t';
  let maxVal = 0;
  for (const [delim, cnt] of Object.entries(counts)) {
    if (cnt > maxVal) {
      maxVal = cnt;
      best = delim;
    }
  }
  return maxVal > 0 ? best : '\t';
}

function parseTextLines(text, forcedDelim = 'auto') {
  const lines = text.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  if (!lines.length) return [];

  const delim = forcedDelim === 'auto' ? detectBestDelimiter(text) : forcedDelim;

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
      // Hỗ trợ CSV có ngoặc kép
      const regex = /(?:^|,)(?:"([^"]*(?:""[^"]*)*)"|([^",]*))/g;
      let match;
      while ((match = regex.exec(line))) {
        if (match[0] === '' && regex.lastIndex >= line.length) break;
        let val = match[1] ? match[1].replace(/""/g, '"') : match[2];
        parts.push(val ? val.trim() : '');
      }
    } else {
      parts = line.split(delim);
    }
    return parts.map(p => p.trim());
  });
}

function analyzeImportText() {
  const text = $('importTextarea').value;
  parsedImportRows = parseTextLines(text, currentDelimiterMode);

  if (!parsedImportRows.length) {
    alert('Vui lòng dán dữ liệu từ vựng vào ô văn bản!');
    return;
  }

  const maxCols = Math.max(...parsedImportRows.map(r => r.length));
  const previewTable = $('previewTable');

  // Gợi ý vai trò cột thông minh
  const suggestedRoles = [];
  for (let c = 0; c < maxCols; c++) {
    const colSamples = parsedImportRows.slice(0, 10).map(r => r[c] || '');
    const isChinese = colSamples.filter(s => /[\u4e00-\u9fff]/.test(s)).length >= colSamples.length * 0.4;
    const isPinyin = colSamples.filter(s => /[āáǎàēéěèīíǐìōóǒòūúǔùǖǘǚǜü\/\[\]]/.test(s)).length >= colSamples.length * 0.4;

    if (c === 0) suggestedRoles.push('q');             // Cột 1 mặc định là Đề bài
    else if (c === 1) suggestedRoles.push('a');        // Cột 2 là Đáp án chuẩn
    else if (isPinyin) suggestedRoles.push('pinyin');  // Phiên âm
    else if (isChinese && !suggestedRoles.includes('a')) suggestedRoles.push('a');
    else if (c === 2) suggestedRoles.push('mask');     // Cột 3 là Ô che/Gợi ý
    else suggestedRoles.push('extra');
  }

  // Header chọn vai trò
  let theadHtml = '<thead><tr>';
  for (let c = 0; c < maxCols; c++) {
    theadHtml += `
      <th>
        <select class="preview-role-select" data-col="${c}">
          <option value="q" ${suggestedRoles[c] === 'q' ? 'selected' : ''}>🎯 Cột Đề bài</option>
          <option value="a" ${suggestedRoles[c] === 'a' ? 'selected' : ''}>🔑 Cột Đáp án</option>
          <option value="mask" ${suggestedRoles[c] === 'mask' ? 'selected' : ''}>👁️ Cột Ô che/Gợi ý</option>
          <option value="pinyin" ${suggestedRoles[c] === 'pinyin' ? 'selected' : ''}>🗣️ Phiên âm/Pinyin</option>
          <option value="extra" ${suggestedRoles[c] === 'extra' ? 'selected' : ''}>📝 Ghi chú/Khác</option>
          <option value="ignore">❌ Bỏ qua</option>
        </select>
      </th>`;
  }
  theadHtml += '</tr></thead><tbody>';

  // Hiển thị 5 dòng xem trước
  parsedImportRows.slice(0, 6).forEach(row => {
    theadHtml += '<tr>';
    for (let c = 0; c < maxCols; c++) {
      theadHtml += `<td>${escapeHtml(row[c] || '')}</td>`;
    }
    theadHtml += '</tr>';
  });
  theadHtml += '</tbody>';

  previewTable.innerHTML = theadHtml;
  $('previewArea').style.display = 'block';
  $('btnConfirmImport').style.display = 'inline-flex';
}

function confirmImport() {
  const selects = document.querySelectorAll('.preview-role-select');
  const colMappings = Array.from(selects).map(s => s.value);

  const newRecords = parsedImportRows.map((row, idx) => {
    const record = {
      id: `VOC-${String(App.counterId++).padStart(4, '0')}`,
      stt: idx + 1,
      q: '', a: '', mask: '', pinyin: '', extra: '', input: '', status: ''
    };

    row.forEach((val, cIdx) => {
      const role = colMappings[cIdx];
      if (role === 'q') record.q = val;
      else if (role === 'a') record.a = val;
      else if (role === 'mask') record.mask = val;
      else if (role === 'pinyin') record.pinyin = val;
      else if (role === 'extra') record.extra = record.extra ? `${record.extra} | ${val}` : val;
    });

    return record;
  });

  App.records = newRecords;
  App.originalRecords = JSON.parse(JSON.stringify(newRecords));

  $('importModal').hidden = true;
  renderTable();
  updateScoring();
}

// ==========================================
// 6. THAO TÁC XÓA NHIỀU HÀNG / NHIỀU CỘT
// ==========================================
function addNewRow() {
  const newId = `VOC-${String(App.counterId++).padStart(4, '0')}`;
  const newRow = {
    id: newId,
    stt: App.records.length + 1,
    q: 'Từ mới',
    a: 'Đáp án',
    mask: 'Gợi ý',
    pinyin: '',
    extra: '',
    input: '',
    status: ''
  };
  App.records.push(newRow);
  renderTable();
  updateScoring();
}

// Xóa tất cả các hàng đang được bôi đen
function deleteSelectedRows() {
  const { startRow, endRow } = App.selection;
  const minR = Math.min(startRow, endRow);
  const maxR = Math.max(startRow, endRow);

  if (App.records.length <= 1) {
    alert('Bảng phải có ít nhất 1 hàng dữ liệu!');
    return;
  }

  const deleteCount = maxR - minR + 1;
  App.records.splice(minR, deleteCount);
  App.records.forEach((r, idx) => r.stt = idx + 1);

  App.selection.startRow = Math.min(minR, App.records.length - 1);
  App.selection.endRow = App.selection.startRow;
  App.selection.activeRow = App.selection.startRow;

  renderTable();
  updateScoring();
}

function addNewColumn() {
  const colId = `custom_${Date.now()}`;
  App.columns.push({
    key: colId,
    title: `Cột ${colLetter(App.columns.length)}`,
    role: 'extra',
    width: 140,
    readonly: false
  });
  renderTable();
}

// Xóa tất cả các cột đang được bôi đen
function deleteSelectedColumns() {
  const { startCol, endCol } = App.selection;
  const minC = Math.min(startCol, endCol);
  const maxC = Math.max(startCol, endCol);

  // Không cho xóa nếu chỉ còn 3 cột cốt lõi
  if (App.columns.length <= 3) {
    alert('Không thể xóa thêm cột! Bảng cần tối thiểu các cột cơ bản.');
    return;
  }

  // Kiểm tra không xóa các cột bảo vệ hệ thống
  const protectedRoles = ['stt', 'id'];
  for (let c = minC; c <= maxC; c++) {
    if (protectedRoles.includes(App.columns[c]?.role)) {
      alert('Không thể xóa cột STT và cột Mã từ!');
      return;
    }
  }

  const deleteCount = maxC - minC + 1;
  App.columns.splice(minC, deleteCount);
  App.selection.activeCol = Math.max(0, minC - 1);
  App.selection.startCol = App.selection.activeCol;
  App.selection.endCol = App.selection.activeCol;

  renderTable();
}

function setColumnRole(role) {
  const { activeCol } = App.selection;
  const targetCol = App.columns[activeCol];

  if (['stt', 'id'].includes(targetCol.role)) {
    alert('Không thể thay đổi vai trò của cột STT và Mã từ!');
    return;
  }

  if (['q', 'a', 'input', 'res'].includes(role)) {
    const existing = App.columns.find(c => c.role === role);
    if (existing && existing !== targetCol) existing.role = 'extra';
  }

  targetCol.role = role;
  renderTable();
  updateScoring();
}

function swapQuestionAndAnswer() {
  const colQ = App.columns.find(c => c.role === 'q');
  const colA = App.columns.find(c => c.role === 'a');

  if (!colQ || !colA) {
    alert('Bảng chưa có đủ cặp cột Đề bài và Đáp án!');
    return;
  }

  colQ.role = 'a';
  colA.role = 'q';

  App.records.forEach(r => {
    [r.q, r.a] = [r.a, r.q];
  });

  renderTable();
  updateScoring();
}

// ==========================================
// 7. TRỘN NGUYÊN HÀNG & PHỤC HỒI
// ==========================================
function shuffleArray(arr) {
  const a = [...arr];
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
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

// ==========================================
// 8. IMPORT NGƯỢC FILE EXCEL/CSV ĐỂ SỬA
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
      alert('File Excel không có sheet nào hợp lệ!');
      return;
    }

    // Tìm hàng tiêu đề (Nếu là file tool xuất thì header ở dòng 5, nếu file ngoài thì ở dòng 1)
    let headerRowIdx = 1;
    for (let r = 1; r <= 10; r++) {
      const rowVals = ws.getRow(r).values;
      if (Array.isArray(rowVals) && rowVals.some(v => String(v).includes('Đề bài') || String(v).includes('Đáp án') || String(v).includes('STT'))) {
        headerRowIdx = r;
        break;
      }
    }

    const headerRow = ws.getRow(headerRowIdx);
    const colCount = ws.columnCount || headerRow.cellCount;
    const importedCols = [];
    const keyMap = [];

    // Nhận diện lại cấu trúc cột
    for (let c = 1; c <= colCount; c++) {
      const title = String(headerRow.getCell(c).value || '').trim();
      if (!title) continue;

      let role = 'extra';
      let key = `col_${c}`;

      if (title.includes('STT')) { role = 'stt'; key = 'stt'; }
      else if (title.includes('Mã từ') || title.includes('ID')) { role = 'id'; key = 'id'; }
      else if (title.includes('Đề bài') || title.includes('Question')) { role = 'q'; key = 'q'; }
      else if (title.includes('Ô làm bài') || title.includes('Ô nhập')) { role = 'input'; key = 'input'; }
      else if (title.includes('Kết quả')) { role = 'res'; key = 'res'; }
      else if (title.includes('Đáp án')) { role = 'a'; key = 'a'; }
      else if (title.includes('Ô che') || title.includes('Gợi ý')) { role = 'mask'; key = 'mask'; }
      else if (title.includes('Phiên âm') || title.includes('Pinyin')) { role = 'pinyin'; key = 'pinyin'; }

      importedCols.push({
        key,
        title,
        role,
        width: 150,
        readonly: ['stt', 'id', 'res'].includes(role)
      });
      keyMap.push({ colIndex: c, key, role });
    }

    // Đọc các hàng dữ liệu
    const importedRecords = [];
    const dataStartRow = headerRowIdx + 1;

    for (let r = dataStartRow; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      // Bỏ qua dòng trống
      if (!row.hasValues) continue;

      const rec = {
        id: `VOC-${String(App.counterId++).padStart(4, '0')}`,
        stt: importedRecords.length + 1,
        input: '',
        status: ''
      };

      keyMap.forEach(({ colIndex, key, role }) => {
        let cellVal = row.getCell(colIndex).value;
        // Nếu cell là công thức/object
        if (cellVal && typeof cellVal === 'object') {
          cellVal = cellVal.result || cellVal.text || '';
        }
        if (role === 'id' && cellVal) rec.id = String(cellVal);
        else if (role !== 'res' && role !== 'input') {
          rec[key] = cellVal != null ? String(cellVal).trim() : '';
        }
      });

      importedRecords.push(rec);
    }

    if (!importedRecords.length) {
      alert('Không tìm thấy dòng dữ liệu nào trong file Excel!');
      return;
    }

    App.columns = importedCols;
    App.records = importedRecords;
    App.originalRecords = JSON.parse(JSON.stringify(importedRecords));

    renderTable();
    updateScoring();
    alert(`Đã nạp thành công ${importedRecords.length} câu từ file Excel vào bảng tính!`);
  } catch (err) {
    console.error(err);
    alert('Không thể đọc file Excel này! Vui lòng đảm bảo file định dạng .xlsx hợp lệ.');
  } finally {
    e.target.value = '';
  }
}

// ==========================================
// 9. XUẤT EXCEL (.XLSX) CHUẨN MỰC
// ==========================================
async function exportToExcel() {
  const colQIdx = App.columns.findIndex(c => c.role === 'q');
  const colAIdx = App.columns.findIndex(c => c.role === 'a');
  const colInpIdx = App.columns.findIndex(c => c.role === 'input');
  const colResIdx = App.columns.findIndex(c => c.role === 'res');
  const colMaskIdx = App.columns.findIndex(c => c.role === 'mask');
  const colIdIdx = App.columns.findIndex(c => c.role === 'id');

  if (colQIdx === -1 || colAIdx === -1 || colInpIdx === -1 || colResIdx === -1) {
    alert('Vui lòng đảm bảo bảng đã có đủ các cột: Đề bài, Đáp án, Ô làm bài và Kết quả!');
    return;
  }

  const wb = new ExcelJS.Workbook();
  wb.creator = 'Vocab Sheet Studio';
  wb.created = new Date();

  const numVariants = Math.max(1, Math.min(10, +$('numVariants').value || 1));
  const password = $('sheetPassword').value.trim();

  for (let v = 0; v < numVariants; v++) {
    const sheetName = numVariants === 1 ? 'Bài tập từ vựng' : `Đề số ${v + 1}`;
    const ws = wb.addWorksheet(sheetName, {
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

    // 1. DASHBOARD BẢNG ĐIỂM (HÀNG 1 - 4) - SỬ DỤNG CHUẨN TYPOGRAPHY KHÔNG LỖI FONT
    ws.mergeCells('A1:D1');
    ws.getCell('A1').value = 'BẢNG ĐÁNH GIÁ KẾT QUẢ TỪ VỰNG';
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

    // Nền xám nhạt cho Dashboard
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

    // 2. TIÊU ĐỀ CỘT HÀNG 5
    App.columns.forEach((col, cIdx) => {
      const cell = ws.getCell(5, cIdx + 1);
      cell.value = col.title;
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF107C41' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      ws.getColumn(cIdx + 1).width = Math.max(12, Math.round(col.width / 7.5));
    });

    // 3. DỮ LIỆU & CÔNG THỨC CHẤM TỰ ĐỘNG HÀNG 6+
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
            cell.value = {
              formula: `IF(OR(${curInp}="",$F$1<>"Đã nộp"),"",IF(${checkFormula},"ĐÚNG","SAI"))`
            };
          } else {
            cell.value = {
              formula: `IF(${curInp}="","",IF(${checkFormula},"ĐÚNG","SAI"))`
            };
          }
          cell.alignment = { horizontal: 'center' };
          cell.font = { bold: true };
        } else if (col.role === 'mask') {
          // CÔNG THỨC Ô CHE: NHẬP ĐÚNG MỚI HIỆN HOẶC NỘP BÀI MỚI HIỆN
          const curInp = `${L_INP}${rIdx}`;
          const curA = `${L_A}${rIdx}`;
          const rawMaskVal = (rec[col.key] || '').replace(/"/g, '""');

          if (App.config.maskRule === 'on_correct') {
            cell.value = {
              formula: `IF(AND(${curInp}<>"",LOWER(TRIM(${curInp}))=LOWER(TRIM(${curA}))),"${rawMaskVal}","🔒 [Nhập đúng để mở]")`
            };
          } else if (App.config.maskRule === 'on_submit') {
            cell.value = {
              formula: `IF($F$1="Đã nộp","${rawMaskVal}","🔒 [Nộp bài để xem]")`
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

    // 4. CONDITIONAL FORMATTING MÀU XANH/ĐỎ CHUẨN
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

    // 5. CHE CỘT MÃ TỪ VÀ ĐÁP ÁN (NẾU CẤU HÌNH)
    if (App.config.hideIdOnExport && colIdIdx !== -1) {
      ws.getColumn(colIdIdx + 1).hidden = true;
    }
    if (App.config.hideAnswerOnExport && colAIdx !== -1) {
      ws.getColumn(colAIdx + 1).hidden = true;
    }

    // 6. KHÓA BẢO VỆ BẢNG TÍNH
    if (App.config.lockSheet) {
      await ws.protect(password, {
        selectLockedCells: true,
        selectUnlockedCells: true
      });
    }
  }

  // Tải file về máy tính
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
// 10. TEXT-TO-SPEECH (PHÁT ÂM TỪ VỰNG)
// ==========================================
function speakText(text) {
  if (!('speechSynthesis' in window)) {
    alert('Trình duyệt của bạn không hỗ trợ tính năng đọc phát âm!');
    return;
  }
  window.speechSynthesis.cancel();
  const clean = text.split('|')[0].trim();
  const utter = new SpeechSynthesisUtterance(clean);
  utter.lang = App.config.ttsLang;
  window.speechSynthesis.speak(utter);
}

// ==========================================
// 11. MINI GAME: THẺ FLASHCARD LẬT 3D
// ==========================================
function openFlashcardModal() {
  if (!App.records.length) {
    alert('Chưa có từ vựng nào để ôn tập!');
    return;
  }
  App.flashcardIdx = 0;
  $('flashcardModal').hidden = false;
  renderFlashcard();
}

function renderFlashcard() {
  const card = App.records[App.flashcardIdx];
  $('flashcardInner').classList.remove('is-flipped');

  $('fcCardIndex').textContent = `Thẻ ${App.flashcardIdx + 1} / ${App.records.length}`;
  $('fcFrontText').textContent = card.q || 'Không có đề bài';
  $('fcFrontSub').textContent = 'Nhấn hoặc phím Space để xem đáp án';

  $('fcBackText').textContent = card.a || '';
  $('fcBackPinyin').textContent = card.pinyin || '';
  $('fcBackMask').textContent = card.mask || card.extra || '';
}

function flipFlashcard() {
  $('flashcardInner').classList.toggle('is-flipped');
}

// ==========================================
// 12. MINI GAME: TRẮC NGHIỆM 4 ĐÁP ÁN (QUIZ)
// ==========================================
function openQuizModal() {
  if (App.records.length < 4) {
    alert('Cần tối thiểu 4 từ vựng trong bảng để tạo bài trắc nghiệm 4 đáp án!');
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

  // Chọn ngẫu nhiên 1 câu làm câu hỏi
  const currentIdx = Math.floor(Math.random() * App.records.length);
  App.quizCurrentRow = App.records[currentIdx];

  $('quizQuestionText').textContent = App.quizCurrentRow.q;
  $('quizScoreText').textContent = `Điểm: ${App.quizScore}`;
  $('quizStreakText').textContent = `🔥 Chuỗi đúng: ${App.quizStreak}`;

  // Lấy 3 đáp án sai từ các dòng khác
  const otherAnswers = App.records
    .filter((_, idx) => idx !== currentIdx)
    .map(r => r.a.split('|')[0].trim());
  const shuffledOthers = shuffleArray(otherAnswers).slice(0, 3);

  const correctAnswer = App.quizCurrentRow.a.split('|')[0].trim();
  const allOptions = shuffleArray([correctAnswer, ...shuffledOthers]);

  const grid = $('quizOptionsGrid');
  grid.innerHTML = '';

  allOptions.forEach(opt => {
    const btn = document.createElement('button');
    btn.className = 'quiz-opt-btn';
    btn.textContent = opt;
    btn.onclick = () => handleQuizAnswer(opt, btn, correctAnswer);
    grid.appendChild(btn);
  });
}

function handleQuizAnswer(selected, buttonEl, correctAnswer) {
  const allBtns = document.querySelectorAll('.quiz-opt-btn');
  allBtns.forEach(b => b.disabled = true);

  const isCorrect = selected.toLowerCase() === correctAnswer.toLowerCase();
  const fb = $('quizFeedbackArea');
  fb.style.display = 'block';

  if (isCorrect) {
    buttonEl.classList.add('correct');
    App.quizScore += 10;
    App.quizStreak += 1;
    fb.style.background = '#dcfce7';
    fb.style.color = '#166534';
    fb.textContent = '🎉 CHÍNH XÁC! Bạn làm rất tốt!';
  } else {
    buttonEl.classList.add('wrong');
    App.quizStreak = 0;
    allBtns.forEach(b => {
      if (b.textContent.toLowerCase() === correctAnswer.toLowerCase()) b.classList.add('correct');
    });
    fb.style.background = '#fee2e2';
    fb.style.color = '#991b1b';
    fb.textContent = `❌ Chưa đúng! Đáp án chính xác là: ${correctAnswer}`;
  }

  $('quizScoreText').textContent = `Điểm: ${App.quizScore}`;
  $('quizStreakText').textContent = `🔥 Chuỗi đúng: ${App.quizStreak}`;
  $('btnQuizNext').style.display = 'inline-flex';
}

// ==========================================
// 13. GẮN SỰ KIỆN GIAO DIỆN & MODAL
// ==========================================
function bindToolbarEvents() {
  // 1. Dữ liệu & File
  $('btnOpenImportModal').onclick = () => {
    $('importModal').hidden = false;
    $('importTextarea').focus();
  };
  $('btnTriggerUpload').onclick = () => $('fileImportExcel').click();
  $('fileImportExcel').onchange = handleImportExcelFile;
  $('btnExportExcel').onclick = exportToExcel;

  // 2. Hàng & Cột
  $('btnAddRow').onclick = addNewRow;
  $('btnDelRow').onclick = deleteSelectedRows;
  $('btnAddCol').onclick = addNewColumn;
  $('btnDelCol').onclick = deleteSelectedColumns;
  $('btnClearCells').onclick = clearSelectedCells;

  // 3. Vai trò cột
  $('btnSetRoleQ').onclick = () => setColumnRole('q');
  $('btnSetRoleInput').onclick = () => setColumnRole('input');
  $('btnSetRoleA').onclick = () => setColumnRole('a');
  $('btnSetRoleMask').onclick = () => setColumnRole('mask');
  $('btnSetRoleRes').onclick = () => setColumnRole('res');
  $('btnSwapQA').onclick = swapQuestionAndAnswer;

  // 4. Trộn đề
  $('btnShuffle').onclick = shuffleAllRows;
  $('btnRestore').onclick = restoreOriginalOrder;
  $('btnPickSubset').onclick = () => pickRandomSubset(+$('numPick').value);

  // 5. Chế độ luyện tập
  $('btnOpenFlashcard').onclick = openFlashcardModal;
  $('btnOpenQuiz').onclick = openQuizModal;
  $('ttsLangSelect').onchange = e => App.config.ttsLang = e.target.value;

  // 6. Cấu hình & Bảo vệ
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
  $('chkHideId').onchange = e => App.config.hideIdOnExport = e.target.checked;
  $('chkHideAnswer').onchange = e => App.config.hideAnswerOnExport = e.target.checked;

  $('btnToggleLock').onclick = e => {
    App.config.lockSheet = !App.config.lockSheet;
    e.target.textContent = `🔒 Khóa Sheet: ${App.config.lockSheet ? 'BẬT' : 'TẮT'}`;
    e.target.classList.toggle('on', App.config.lockSheet);
  };

  // Nộp bài kiểm tra
  $('btnSubmitExam').onclick = () => {
    App.config.isExamSubmitted = true;
    $('btnSubmitExam').style.display = 'none';
    $('btnResetExam').style.display = 'inline-flex';
    updateScoring();
    renderTable();
  };
  $('btnResetExam').onclick = () => {
    App.config.isExamSubmitted = false;
    App.records.forEach(r => { r.input = ''; r.status = ''; });
    $('btnSubmitExam').style.display = 'inline-flex';
    $('btnResetExam').style.display = 'none';
    updateScoring();
    renderTable();
  };

  // Phát âm dòng đang chọn
  $('btnSpeakActiveRow').onclick = () => {
    const row = App.records[App.selection.activeRow];
    if (row) speakText(row.a || row.q);
  };
}

function bindModalsEvents() {
  // Modal Quét từ vựng
  $('btnCloseImportModal').onclick = () => $('importModal').hidden = true;
  $('btnCancelImport').onclick = () => $('importModal').hidden = true;
  $('btnAnalyzeText').onclick = analyzeImportText;
  $('btnConfirmImport').onclick = confirmImport;

  document.querySelectorAll('.btn-delimit').forEach(btn => {
    btn.onclick = () => {
      document.querySelectorAll('.btn-delimit').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      currentDelimiterMode = btn.dataset.delim;
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
    if (card) speakText(card.a || card.q);
  };

  // Phím tắt Flashcard (Space: Lật, Mũi tên trái/phải: Đổi thẻ)
  document.addEventListener('keydown', e => {
    if (!$('flashcardModal').hidden) {
      if (e.code === 'Space') {
        e.preventDefault();
        flipFlashcard();
      } else if (e.key === 'ArrowRight') {
        App.flashcardIdx = (App.flashcardIdx + 1) % App.records.length;
        renderFlashcard();
      } else if (e.key === 'ArrowLeft') {
        App.flashcardIdx = (App.flashcardIdx - 1 + App.records.length) % App.records.length;
        renderFlashcard();
      }
    }
  });

  // Modal Quiz
  $('btnCloseQuizModal').onclick = () => $('quizModal').hidden = true;
  $('btnQuizExit').onclick = () => $('quizModal').hidden = true;
  $('btnQuizNext').onclick = nextQuizQuestion;
  $('btnQuizSpeak').onclick = () => {
    if (App.quizCurrentRow) speakText(App.quizCurrentRow.q);
  };
}

// Khởi chạy ứng dụng khi DOM tải xong
window.addEventListener('DOMContentLoaded', initApp);
