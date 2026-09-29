/**
 * VOCAB SHEET STUDIO - INDEX.JS
 * Xây dựng kiến trúc Record-based chuẩn:
 * - Dữ liệu là các Object có ID cố định (VOC-0001)
 * - Trộn đề (Shuffle) nguyên hàng: Đề + Đáp án luôn đi cùng nhau
 * - Chấm điểm trực quan: Đúng / Sai / Chưa làm / Điểm 10 / Tiến độ
 * - Xuất XLSX với công thức Excel động & Khóa bảo vệ Sheet
 */

// ==========================================
// 1. TRẠNG THÁI ỨNG DỤNG (APPLICATION STATE)
// ==========================================
const App = {
  // Mảng dữ liệu các dòng từ vựng (Record Object)
  records: [
    { id: 'VOC-0001', stt: 1, q: 'xin chào', a: '你好', pinyin: 'nǐ hǎo', extra: 'Lời chào thân mật', input: '', status: '' },
    { id: 'VOC-0002', stt: 2, q: 'cảm ơn', a: '谢谢', pinyin: 'xièxie', extra: 'Cảm tạ', input: '', status: '' },
    { id: 'VOC-0003', stt: 3, q: 'trường học', a: '学校', pinyin: 'xuéxiào', extra: 'Nơi học tập', input: '', status: '' },
    { id: 'VOC-0004', stt: 4, q: 'giáo viên', a: '老师', pinyin: 'lǎoshī', extra: 'Người dạy học', input: '', status: '' },
    { id: 'VOC-0005', stt: 5, q: 'học sinh', a: '学生', pinyin: 'xuéshēng', extra: 'Người đi học', input: '', status: '' }
  ],
  originalRecords: null, // Lưu thứ tự gốc để khôi phục

  // Cấu hình các cột hiển thị trong bảng
  columns: [
    { key: 'stt', title: 'STT', role: 'stt', width: 60, readonly: true },
    { key: 'id', title: 'Mã từ', role: 'id', width: 90, readonly: true },
    { key: 'q', title: 'Đề bài', role: 'q', width: 180, readonly: false },
    { key: 'input', title: 'Ô nhập đáp án', role: 'input', width: 180, readonly: false },
    { key: 'res', title: 'Kết quả', role: 'res', width: 100, readonly: true },
    { key: 'a', title: 'Đáp án chuẩn', role: 'a', width: 180, readonly: false },
    { key: 'pinyin', title: 'Pinyin', role: 'extra', width: 140, readonly: false },
    { key: 'extra', title: 'Ghi chú / Nghĩa phụ', role: 'extra', width: 180, readonly: false }
  ],

  // Cấu hình bài kiểm tra & chấm điểm
  config: {
    mode: 'practice',          // 'practice' | 'exam'
    caseInsensitive: true,     // Bỏ qua hoa thường
    multipleAnswers: true,     // Cho phép nhiều đáp án a | b
    hideAnswerOnExport: false, // Ẩn cột đáp án khi xuất XLSX
    lockSheet: true,           // Khóa Sheet Excel bảo vệ công thức
    sheetPassword: '',         // Mật khẩu khóa
    isExamSubmitted: false     // Trạng thái đã nộp bài ở chế độ kiểm tra
  },

  // Trạng thái chọn ô (Spreadsheet Selection)
  selection: {
    startRow: 0, startCol: 0,
    endRow: 0, endCol: 0,
    activeRow: 0, activeCol: 0
  },
  isDragging: false,
  counterId: 6 // Bộ đếm sinh VOC-ID tiếp theo
};

// Hàm chuyển số thành chữ cái cột Excel: 0 -> A, 1 -> B, 25 -> Z, 26 -> AA...
function colLetter(n) {
  let s = '';
  for (n++; n; n = Math.floor((n - 1) / 26)) {
    s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  }
  return s;
}

// Shortcut lấy element theo ID
const $ = id => document.getElementById(id);

// ==========================================
// 2. KHỞI TẠO & RENDER BẢNG TÍNH (SPREADSHEET ENGINE)
// ==========================================
function initApp() {
  // Sao lưu bản gốc
  App.originalRecords = JSON.parse(JSON.stringify(App.records));

  // Gán sự kiện
  bindToolbarEvents();
  bindSelectionEvents();

  // Render bảng và tính điểm ban đầu
  renderTable();
  updateScoring();
}

/**
 * Render toàn bộ cấu trúc Table (Header + Data Rows)
 */
function renderTable() {
  const table = $('sheet-table');
  const { columns, records } = App;

  // 1. Tạo THEAD
  let theadHtml = '<thead><tr><th class="corner-header" title="Chọn toàn bộ bảng">◢</th>';
  columns.forEach((col, cIdx) => {
    const roleName = getRoleBadgeName(col.role);
    theadHtml += `
      <th class="col-header" data-c="${cIdx}" style="width: ${col.width}px;" title="Nhấp để chọn cột ${colLetter(cIdx)}">
        ${colLetter(cIdx)}
        <span class="col-role-tag">${roleName}</span>
      </th>`;
  });
  theadHtml += '</tr></thead>';

  // 2. Tạo TBODY
  let tbodyHtml = '<tbody>';
  records.forEach((row, rIdx) => {
    tbodyHtml += `<tr><th class="row-header" data-r="${rIdx}">${rIdx + 1}</th>`;
    columns.forEach((col, cIdx) => {
      let val = '';
      let isEditable = !col.readonly;
      let extraClass = '';

      if (col.role === 'stt') val = rIdx + 1;
      else if (col.role === 'id') val = row.id;
      else if (col.role === 'res') {
        val = row.status || '';
        extraClass = 'col-res-cell';
      } else if (col.role === 'input') {
        val = row.input || '';
        extraClass = 'col-input-cell';
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
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function getRoleBadgeName(role) {
  switch (role) {
    case 'q': return '🎯 Đề bài';
    case 'a': return '🔑 Đáp án';
    case 'input': return '✏️ Ô nhập';
    case 'res': return '📊 Kết quả';
    case 'stt': return 'STT';
    case 'id': return 'Mã từ';
    default: return 'Ghi chú';
  }
}

function getStatusStyle(status, role) {
  // Chỉ đổi màu cột Nhập và cột Kết quả
  if (role !== 'input' && role !== 'res') return '';
  if (status === 'ĐÚNG') return 'ok';
  if (status === 'SAI') return 'bad';
  return '';
}

// ==========================================
// 3. SELECTION & TƯƠNG TÁC Ô TRỰC QUAN
// ==========================================
function bindSelectionEvents() {
  const table = $('sheet-table');

  // Chuột nhấn xuống bắt đầu bôi đen
  table.addEventListener('mousedown', e => {
    const td = e.target.closest('td');
    const thCol = e.target.closest('th.col-header');
    const thRow = e.target.closest('th.row-header');
    const thCorner = e.target.closest('th.corner-header');

    if (td) {
      const r = +td.dataset.r;
      const c = +td.dataset.c;
      App.selection = { startRow: r, startCol: c, endRow: r, endCol: c, activeRow: r, activeCol: c };
      App.isDragging = true;
      highlightSelection();
      updateFxBar();
    } else if (thCol) {
      // Chọn cả cột
      const c = +thCol.dataset.c;
      App.selection = { startRow: 0, startCol: c, endRow: App.records.length - 1, endCol: c, activeRow: 0, activeCol: c };
      highlightSelection();
      updateFxBar();
    } else if (thRow) {
      // Chọn cả hàng
      const r = +thRow.dataset.r;
      App.selection = { startRow: r, startCol: 0, endRow: r, endCol: App.columns.length - 1, activeRow: r, activeCol: 0 };
      highlightSelection();
      updateFxBar();
    } else if (thCorner) {
      // Chọn toàn bộ bảng
      App.selection = { startRow: 0, startCol: 0, endRow: App.records.length - 1, endCol: App.columns.length - 1, activeRow: 0, activeCol: 0 };
      highlightSelection();
    }
  });

  // Kéo chuột bôi đen vùng (Range Drag)
  table.addEventListener('mouseover', e => {
    if (!App.isDragging) return;
    const td = e.target.closest('td');
    if (td) {
      App.selection.endRow = +td.dataset.r;
      App.selection.endCol = +td.dataset.c;
      highlightSelection();
    }
  });

  document.addEventListener('mouseup', () => {
    App.isDragging = false;
  });

  // Chỉnh sửa trực tiếp tại ô
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

    // Cập nhật điểm & trạng thái
    evaluateRow(App.records[r]);
    updateScoring();
    refreshRowDisplay(r);
    updateFxBar();
  });

  // Điều hướng bằng bàn phím (Enter: Xuống hàng, Tab: Sang ô phải, Delete: Xóa nội dung)
  table.addEventListener('keydown', e => {
    const td = e.target.closest('td');
    if (!td) return;
    const r = +td.dataset.r;
    const c = +td.dataset.c;

    if (e.key === 'Enter') {
      e.preventDefault();
      const nextTd = getCell(r + 1, c);
      if (nextTd) {
        nextTd.focus();
        App.selection = { startRow: r + 1, startCol: c, endRow: r + 1, endCol: c, activeRow: r + 1, activeCol: c };
        highlightSelection();
        updateFxBar();
      }
    } else if (e.key === 'Tab') {
      e.preventDefault();
      const nextTd = getCell(r, c + 1);
      if (nextTd) {
        nextTd.focus();
        App.selection = { startRow: r, startCol: c + 1, endRow: r, endCol: c + 1, activeRow: r, activeCol: c + 1 };
        highlightSelection();
        updateFxBar();
      }
    }
  });

  // Bắt phím Delete để xóa nhanh các ô bôi đen
  document.addEventListener('keydown', e => {
    if (e.key === 'Delete' && !['INPUT', 'TEXTAREA'].includes(document.activeElement.tagName)) {
      clearSelectedCells();
    }
  });

  // Đồng bộ thanh Fx Bar với ô
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
 * Bôi đen vùng ô đang được chọn
 */
function highlightSelection() {
  document.querySelectorAll('.cell-selected, .cell-focus, th.selected').forEach(el => {
    el.classList.remove('cell-selected', 'cell-focus', 'selected');
  });

  const { startRow, startCol, endRow, endCol, activeRow, activeCol } = App.selection;
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
    if (thRow) thRow.classList.add('selected');
  }

  for (let c = minC; c <= maxC; c++) {
    const thCol = document.querySelector(`th.col-header[data-c="${c}"]`);
    if (thCol) thCol.classList.add('selected');
  }
}

function updateFxBar() {
  const { activeRow, activeCol } = App.selection;
  const coord = `${colLetter(activeCol)}${activeRow + 1}`;
  $('active-cell-coord').textContent = coord;

  const col = App.columns[activeCol];
  const row = App.records[activeRow];
  let val = '';
  if (row && col) {
    if (col.role === 'stt') val = activeRow + 1;
    else if (col.role === 'id') val = row.id;
    else if (col.role === 'res') val = row.status || '';
    else if (col.role === 'input') val = row.input || '';
    else val = row[col.key] ?? '';
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
// 4. CHẤM ĐIỂM & ĐÁNH GIÁ (SCORING ENGINE)
// ==========================================
function evaluateRow(row) {
  const input = (row.input || '').trim();
  const answer = (row.a || '').trim();

  // Ở chế độ kiểm tra và chưa nộp bài -> Không hiện kết quả
  if (App.config.mode === 'exam' && !App.config.isExamSubmitted) {
    row.status = '';
    return;
  }

  // Chưa nhập -> Chưa làm (Không tính sai)
  if (!input) {
    row.status = '';
    return;
  }

  const { caseInsensitive, multipleAnswers } = App.config;
  const normalize = str => caseInsensitive ? str.toLowerCase().trim() : str.trim();

  const userAns = normalize(input);

  if (multipleAnswers) {
    // Tách các đáp án hợp lệ theo dấu |
    const validAnswers = answer.split('|').map(normalize);
    row.status = validAnswers.includes(userAns) ? 'ĐÚNG' : 'SAI';
  } else {
    row.status = (userAns === normalize(answer)) ? 'ĐÚNG' : 'SAI';
  }
}

function updateScoring() {
  // Đánh giá lại tất cả các dòng
  App.records.forEach(row => evaluateRow(row));

  let correct = 0;
  let wrong = 0;
  let answered = 0;
  let totalWithAnswer = 0;

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

  // Hiển thị lên Dashboard
  $('statCorrect').textContent = correct;
  $('statWrong').textContent = wrong;
  $('statPending').textContent = pending;
  $('statTotal').textContent = totalWithAnswer;
  $('statScore').textContent = score10;
  $('statRate').textContent = `${rate}%`;

  $('progressText').textContent = `${answered}/${totalWithAnswer} câu (${progressPercent}%)`;
  $('progressBarFill').style.width = `${progressPercent}%`;

  // Cập nhật Badge chế độ
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
// 5. TRỘN NGUYÊN HÀNG & TẠO ĐỀ THI (SHUFFLE)
// ==========================================
/**
 * Thuật toán Fisher-Yates Shuffle xáo trộn nguyên mảng Record.
 * Cả dòng [STT, ID, Đề, Đáp án, Pinyin, Ghi chú] di chuyển cùng nhau!
 */
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
  // Đánh lại số thứ tự hiển thị
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
  if (n < 1 || n > App.originalRecords.length) {
    alert(`Vui lòng nhập số câu hợp lệ từ 1 đến ${App.originalRecords.length}`);
    return;
  }
  const shuffled = shuffleArray(App.originalRecords);
  App.records = shuffled.slice(0, n);
  App.records.forEach((r, idx) => r.stt = idx + 1);
  renderTable();
  updateScoring();
}

// ==========================================
// 6. QUÉT DỮ LIỆU THÔNG MINH (SMART PARSER)
// ==========================================
let parsedImportData = [];

function openImportModal() {
  $('importModal').hidden = false;
  $('importTextarea').focus();
}

function closeImportModal() {
  $('importModal').hidden = true;
  $('previewArea').style.display = 'none';
  $('btnConfirmImport').style.display = 'none';
  $('btnAnalyze').style.display = 'inline-flex';
}

/**
 * Phân tích dữ liệu paste vào: Tự nhận diện Tab, phẩy, gạch |, hoặc dấu cách
 */
function parseRawText(text) {
  const lines = text.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  if (!lines.length) return [];

  const CJK_REGEX = /[\u4e00-\u9fff]/;
  const PINYIN_REGEX = /[āáǎàēéěèīíǐìōóǒòūúǔùǖǘǚǜü]/i;

  return lines.map(line => {
    // 1. Phân tách theo ký tự ngăn cách nếu có
    let parts = null;
    if (line.includes('\t')) parts = line.split('\t');
    else if (line.includes('|')) parts = line.split('|');
    else if (line.includes(';')) parts = line.split(';');
    else if (/\s{2,}/.test(line)) parts = line.split(/\s{2,}/);

    // 2. Nếu không có dấu ngăn cách, phân tích dạng "你好 nǐ hǎo xin chào"
    if (!parts) {
      const match = line.match(/^([\u4e00-\u9fff，。？！、…]+)\s+(.+)$/);
      if (match) {
        const hanzi = match[1];
        const restWords = match[2].split(/\s+/);
        let pinyinArr = [];

        // Lấy các từ pinyin tương ứng số lượng chữ Hán
        while (restWords.length > 1 && (pinyinArr.length < hanzi.length || PINYIN_REGEX.test(restWords[0]))) {
          pinyinArr.push(restWords.shift());
        }
        parts = [hanzi, pinyinArr.join(' '), restWords.join(' ')];
      } else {
        parts = line.split(/\s+/);
      }
    }

    return parts.map(p => p.trim());
  });
}

function analyzeImportText() {
  const text = $('importTextarea').value;
  parsedImportData = parseRawText(text);

  if (!parsedImportData.length) {
    alert('Không tìm thấy dữ liệu từ vựng nào! Vui lòng kiểm tra lại văn bản dán.');
    return;
  }

  const maxCols = Math.max(...parsedImportData.map(r => r.length));
  const previewTable = $('previewTable');

  // Gợi ý vai trò từng cột
  const suggestedRoles = [];
  for (let c = 0; c < maxCols; c++) {
    const colSamples = parsedImportData.slice(0, 10).map(r => r[c] || '');
    const isChinese = colSamples.filter(s => /[\u4e00-\u9fff]/.test(s)).length >= colSamples.length * 0.4;
    const isPinyin = colSamples.filter(s => /[āáǎàēéěèīíǐìōóǒòūúǔùǖǘǚǜü]/i.test(s)).length >= colSamples.length * 0.4;

    if (isChinese) suggestedRoles.push('a');       // Mặc định Tiếng Trung là Đáp án
    else if (isPinyin) suggestedRoles.push('pinyin');
    else if (c === maxCols - 1 || suggestedRoles.includes('a')) suggestedRoles.push('q'); // Tiếng Việt là Đề
    else suggestedRoles.push('extra');
  }

  // Tạo bảng Preview
  let theadHtml = '<thead><tr>';
  for (let c = 0; c < maxCols; c++) {
    theadHtml += `
      <th>
        <select class="preview-role-select" data-col="${c}">
          <option value="q" ${suggestedRoles[c] === 'q' ? 'selected' : ''}>🎯 Cột Đề bài</option>
          <option value="a" ${suggestedRoles[c] === 'a' ? 'selected' : ''}>🔑 Cột Đáp án</option>
          <option value="pinyin" ${suggestedRoles[c] === 'pinyin' ? 'selected' : ''}>Pinyin</option>
          <option value="extra" ${suggestedRoles[c] === 'extra' ? 'selected' : ''}>Ghi chú / Khác</option>
          <option value="ignore">❌ Bỏ qua</option>
        </select>
      </th>`;
  }
  theadHtml += '</tr></thead><tbody>';

  // Hiển thị 5 dòng đầu xem trước
  parsedImportData.slice(0, 5).forEach(row => {
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
  $('btnAnalyze').style.display = 'none';
}

function confirmImport() {
  const selects = document.querySelectorAll('.preview-role-select');
  const colMappings = Array.from(selects).map(sel => sel.value);

  const newRecords = parsedImportData.map((row, idx) => {
    const record = {
      id: `VOC-${String(App.counterId++).padStart(4, '0')}`,
      stt: idx + 1,
      q: '',
      a: '',
      pinyin: '',
      extra: '',
      input: '',
      status: ''
    };

    row.forEach((val, cIdx) => {
      const role = colMappings[cIdx];
      if (role === 'q') record.q = val;
      else if (role === 'a') record.a = val;
      else if (role === 'pinyin') record.pinyin = val;
      else if (role === 'extra') record.extra = record.extra ? `${record.extra} | ${val}` : val;
    });

    return record;
  });

  App.records = newRecords;
  App.originalRecords = JSON.parse(JSON.stringify(newRecords));

  closeImportModal();
  renderTable();
  updateScoring();
}

// ==========================================
// 7. THAO TÁC CẤU TRÚC BẢNG (THÊM/XÓA HÀNG & CỘT)
// ==========================================
function addNewRow() {
  const newId = `VOC-${String(App.counterId++).padStart(4, '0')}`;
  const newRow = {
    id: newId,
    stt: App.records.length + 1,
    q: 'Từ mới',
    a: 'Đáp án',
    pinyin: '',
    extra: '',
    input: '',
    status: ''
  };
  App.records.push(newRow);
  renderTable();
  updateScoring();
}

function deleteSelectedRows() {
  const { startRow, endRow } = App.selection;
  const minR = Math.min(startRow, endRow);
  const maxR = Math.max(startRow, endRow);

  if (App.records.length <= 1) {
    alert('Bảng phải có ít nhất 1 hàng dữ liệu!');
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

function deleteSelectedColumn() {
  const { activeCol } = App.selection;
  const col = App.columns[activeCol];

  if (['stt', 'id', 'q', 'a', 'input', 'res'].includes(col.role)) {
    alert(`Không thể xóa cột cốt lõi: ${col.title}! Bạn chỉ có thể xóa các cột phụ thêm.`);
    return;
  }

  App.columns.splice(activeCol, 1);
  App.selection.activeCol = Math.max(0, activeCol - 1);
  renderTable();
}

// Gán vai trò cho cột đang được chọn
function setColumnRole(role) {
  const { activeCol } = App.selection;
  const targetCol = App.columns[activeCol];

  if (['stt', 'id'].includes(targetCol.role)) {
    alert('Không thể thay đổi vai trò của cột STT và Mã từ!');
    return;
  }

  // Nếu vai trò đã tồn tại ở cột khác (q, a, input, res), gán cột cũ về extra
  if (['q', 'a', 'input', 'res'].includes(role)) {
    const existing = App.columns.find(c => c.role === role);
    if (existing) existing.role = 'extra';
  }

  targetCol.role = role;
  renderTable();
  updateScoring();
}

function swapQuestionAndAnswer() {
  const colQ = App.columns.find(c => c.role === 'q');
  const colA = App.columns.find(c => c.role === 'a');

  if (!colQ || !colA) {
    alert('Chưa xác định đủ cột Đề bài và cột Đáp án!');
    return;
  }

  colQ.role = 'a';
  colA.role = 'q';

  // Đổi dữ liệu giữa q và a trong tất cả các record
  App.records.forEach(r => {
    [r.q, r.a] = [r.a, r.q];
  });

  renderTable();
  updateScoring();
}

// ==========================================
// 8. GẮN SỰ KIỆN TOOLBAR
// ==========================================
function bindToolbarEvents() {
  // Modal Dán dữ liệu
  $('btnImportModal').onclick = openImportModal;
  $('btnModalClose').onclick = closeImportModal;
  $('btnCancelImport').onclick = closeImportModal;
  $('btnAnalyze').onclick = analyzeImportText;
  $('btnConfirmImport').onclick = confirmImport;

  // Thêm/Xóa hàng & cột
  $('btnAddRow').onclick = addNewRow;
  $('btnDelRow').onclick = deleteSelectedRows;
  $('btnAddCol').onclick = addNewColumn;
  $('btnDelCol').onclick = deleteSelectedColumn;

  // Vai trò cột
  $('btnSetRoleQ').onclick = () => setColumnRole('q');
  $('btnSetRoleA').onclick = () => setColumnRole('a');
  $('btnSetRoleInput').onclick = () => setColumnRole('input');
  $('btnSetRoleRes').onclick = () => setColumnRole('res');
  $('btnSwapQA').onclick = swapQuestionAndAnswer;

  // Cấu hình chế độ & quy tắc
  $('examMode').onchange = e => {
    App.config.mode = e.target.value;
    App.config.isExamSubmitted = false;
    $('btnSubmitExam').style.display = 'inline-flex';
    $('btnResetExam').style.display = 'none';
    updateScoring();
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

  $('chkHideAnswer').onchange = e => {
    App.config.hideAnswerOnExport = e.target.checked;
  };

  // Trộn & Tạo đề
  $('btnShuffle').onclick = shuffleAllRows;
  $('btnRestore').onclick = restoreOriginalOrder;
  $('btnPickSubset').onclick = () => {
    const n = +$('numPick').value;
    pickRandomSubset(n);
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

  // Nút khóa Sheet
  $('btnToggleLock').onclick = e => {
    App.config.lockSheet = !App.config.lockSheet;
    e.target.textContent = `🔒 Khóa Sheet: ${App.config.lockSheet ? 'BẬT' : 'TẮT'}`;
    e.target.classList.toggle('on', App.config.lockSheet);
  };

  // Xuất Excel
  $('btnExportExcel').onclick = exportToExcel;
}

// ==========================================
// 9. XUẤT FILE EXCEL (.XLSX) CHUẨN MỰC
// ==========================================
async function exportToExcel() {
  const colQIdx = App.columns.findIndex(c => c.role === 'q');
  const colAIdx = App.columns.findIndex(c => c.role === 'a');
  const colInpIdx = App.columns.findIndex(c => c.role === 'input');
  const colResIdx = App.columns.findIndex(c => c.role === 'res');

  if (colQIdx === -1 || colAIdx === -1 || colInpIdx === -1 || colResIdx === -1) {
    alert('Vui lòng đảm bảo bảng đã có đủ 4 cột: Đề bài, Đáp án chuẩn, Ô nhập, và Kết quả!');
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
      views: [{ state: 'frozen', ySplit: 5 }] // Cố định 5 hàng đầu (Dashboard + Header)
    });

    // Dữ liệu cho từng đề (Nếu nhiều đề, xáo trộn nguyên hàng độc lập)
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

    // 1. VẼ DASHBOARD BẢNG ĐIỂM (HÀNG 1 - 4)
    // Hàng 1: Tiêu đề Dashboard & Trạng thái nộp bài
    ws.mergeCells('A1:D1');
    ws.getCell('A1').value = 'BẢNG KẾT QUẢ ĐÁNH GIÁ TỪ VỰNG';
    ws.getCell('A1').font = { bold: true, size: 13, color: { argb: 'FF107C41' } };

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

    // Hàng 2: Thống kê Đúng, Sai, Chưa làm, Tổng câu
    ws.getCell('A2').value = '✅ Đúng:';
    ws.getCell('B2').value = { formula: `COUNTIF(${L_RES}${rowStart}:${L_RES}${rowEnd},"ĐÚNG")` };
    ws.getCell('C2').value = '❌ Sai:';
    ws.getCell('D2').value = { formula: `COUNTIF(${L_RES}${rowStart}:${L_RES}${rowEnd},"SAI")` };
    ws.getCell('E2').value = '⚪ Chưa làm:';
    ws.getCell('F2').value = { formula: `COUNTBLANK(${L_INP}${rowStart}:${L_INP}${rowEnd})` };

    // Hàng 3: Điểm số & Tỷ lệ hoàn thành
    ws.getCell('A3').value = '📝 Tổng câu:';
    ws.getCell('B3').value = { formula: `COUNTA(${L_A}${rowStart}:${L_A}${rowEnd})` };
    ws.getCell('C3').value = '📊 Điểm /10:';
    ws.getCell('D3').value = { formula: `IF(B3=0,0,ROUND(B2/B3*10,2))` };
    ws.getCell('E3').value = '📈 Tỷ lệ:';
    ws.getCell('F3').value = { formula: `IF(B3=0,0,B2/B3)` };
    ws.getCell('F3').numFmt = '0.0%';

    // Hàng 4: Thanh tiến độ trực quan
    ws.getCell('A4').value = 'Đã hoàn thành:';
    ws.getCell('B4').value = { formula: `COUNTA(${L_INP}${rowStart}:${L_INP}${rowEnd})` };
    ws.getCell('C4').value = 'Tiến độ:';
    ws.mergeCells('D4:F4');
    ws.getCell('D4').value = {
      formula: `REPT("█",MIN(10,ROUND(B4/MAX(B3,1)*10,0)))&REPT("░",10-MIN(10,ROUND(B4/MAX(B3,1)*10,0)))&" "&TEXT(IF(B3=0,0,B4/B3),"0%")`
    };

    // Định dạng màu nền cho Dashboard
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

    // 2. HÀNG 5: TIÊU ĐỀ BẢNG DỮ LIỆU (HEADER)
    App.columns.forEach((col, cIdx) => {
      const cell = ws.getCell(5, cIdx + 1);
      cell.value = col.title;
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF107C41' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      ws.getColumn(cIdx + 1).width = Math.max(12, Math.round(col.width / 7.5));
    });

    // 3. HÀNG 6+: GHI DỮ LIỆU & CÔNG THỨC CHẤM TỪNG HÀNG
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
          // Ô HỌC SINH NHẬP -> MỞ KHÓA BẢO VỆ
          cell.value = '';
          cell.protection = { locked: false };
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCFBF2' } };
        } else if (col.role === 'res') {
          // CÔNG THỨC CHẤM TỰ ĐỘNG
          const curInp = `${L_INP}${rIdx}`;
          const curA = `${L_A}${rIdx}`;

          const norm = s => App.config.caseInsensitive ? `LOWER(TRIM(${s}))` : `TRIM(${s})`;
          let checkFormula = '';

          if (App.config.multipleAnswers) {
            // Kiểm tra nhiều đáp án chứa dấu |
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

    // 4. CONDITIONAL FORMATTING (TÔ MÀU ĐÚNG / SAI)
    const rangeRes = `${L_RES}${rowStart}:${L_RES}${rowEnd}`;
    const rangeInp = `${L_INP}${rowStart}:${L_INP}${rowEnd}`;

    [rangeRes, rangeInp].forEach(ref => {
      ws.addConditionalFormatting({
        ref: ref,
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

    // Ẩn cột đáp án nếu được chọn
    if (App.config.hideAnswerOnExport) {
      ws.getColumn(colAIdx + 1).hidden = true;
    }

    // 5. KHÓA BẢO VỆ SHEET (CHỈ CHO PHÉP NHẬP Ô ĐÁP ÁN)
    if (App.config.lockSheet) {
      await ws.protect(password, {
        selectLockedCells: true,
        selectUnlockedCells: true
      });
    }
  }

  // Tải file về máy
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `De_Tu_Vung_${new Date().toISOString().slice(0, 10)}.xlsx`;
  a.click();
  URL.revokeObjectURL(url);
}

// Khởi chạy khi tải xong trang
window.addEventListener('DOMContentLoaded', initApp);
