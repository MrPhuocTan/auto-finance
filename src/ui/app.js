/**
 * Giao diện Điều khiển Chính - Công Cụ Đối Soát Và Gom Khoản Tiền Excel
 * Thiết kế chuyên biệt cho kế toán Việt Nam (chuẩn phần mềm văn phòng Excel / MISA).
 */

import {
  DEFAULT_TOLERANCE,
  DEFAULT_MAX_ITEMS,
  DEFAULT_MAX_RESULTS,
  DEFAULT_TIME_LIMIT_MS,
} from '../core/contracts.js';

import {
  parseVietnameseAmount,
  formatNumberVN,
  formatInputAmountVN,
  formatDateVN,
  isoToUtcTimestamp,
  readAmountInVietnameseWords,
} from '../core/vnFormat.js';

import {
  loadWorkbookFromBuffer,
  getSheetNames,
  inspectSheet,
  inspectRawGrid,
  extractTransactions,
  exportReconciledWorkbook,
} from '../core/excelParser.js';

import {
  normalizeVietnameseSearch,
  findCombinations,
} from '../core/reconcileEngine.js';

import { getSampleDataset } from '../core/sampleData.js';

/* ==========================================================================
   TRẠNG THÁI ỨNG DỤNG (CENTRAL STATE)
   ========================================================================== */
const state = {
  // Tệp & Trang tính
  workbook: null,
  sheetNames: [],
  activeSheetName: '',
  fileName: '',
  isSampleData: false,
  samplePresets: [],
  activePresetIndex: null,

  // Cấu trúc cột & dữ liệu trích xuất
  inspection: null,
  headerRowIndex: 0,
  mapping: {
    dateColIndex: null,
    amountColIndex: null,
    debitColIndex: null,
    creditColIndex: null,
    amountMode: 'single',
    descColIndex: null,
    refColIndex: null,
  },
  partnerColIndex: null,
  isMappingOpen: false,
  includeSummaryRows: false,

  items: [],
  itemById: new Map(),
  summaryRowCount: 0,
  skippedEmptyOrInvalidCount: 0,
  minDateIso: null,
  maxDateIso: null,
  availableMonths: [],

  // Điều kiện lọc & tìm kiếm (Bước 2)
  dateFromIso: '',
  dateToIso: '',
  activeMonthKey: 'all',
  targetAmount: null,
  tolerance: DEFAULT_TOLERANCE,
  keyword: '',
  maxItems: DEFAULT_MAX_ITEMS,

  // Tương tác dòng & Đối soát nhiều đợt
  pinnedIds: new Set(),
  manualExcludedIds: new Set(),
  checkedIds: new Set(),
  inspectedIds: new Set(),
  inspectLabel: '',

  reconciledBatches: [],
  reconciledMap: new Map(), // itemId -> batch
  nextBatchSeq: 1,

  // Tiến trình tìm kiếm & Kết quả (Bước 3)
  isSearching: false,
  activeRequestId: 0,
  cancelFallbackFlag: false,
  worker: null,
  workerDisabled: false,
  cancelFallbackTimer: null,

  searchResult: null,
  lastSearchContext: null,
  resultsFilter: 'all',
  expandedSolutionIds: new Set(),
  solutionDraftNotes: new Map(),

  // Chế độ hiển thị Sổ chi tiết
  ledgerView: 'filtered',
  ledgerPageSize: 200,
  ledgerPage: 1,
};

let toastTimer = null;

/* ==========================================================================
   THAM CHIẾU DOM
   ========================================================================== */
const dom = {};

function cacheDomElements() {
  // Topbar & Hướng dẫn
  dom.btnFontNormal = document.getElementById('btnFontNormal');
  dom.btnFontLarge = document.getElementById('btnFontLarge');
  dom.btnToggleGuide = document.getElementById('btnToggleGuide');
  dom.btnCloseGuide = document.getElementById('btnCloseGuide');
  dom.quickGuidePanel = document.getElementById('quickGuidePanel');
  dom.statusToast = document.getElementById('statusToast');

  // Bước 1: Nạp file & Cột
  dom.dropZone = document.getElementById('dropZone');
  dom.excelFileInput = document.getElementById('excelFileInput');
  dom.btnPickFile = document.getElementById('btnPickFile');
  dom.btnLoadSample = document.getElementById('btnLoadSample');

  dom.loadedFileBar = document.getElementById('loadedFileBar');
  dom.fileNameDisplay = document.getElementById('fileNameDisplay');
  dom.sheetSelectWrap = document.getElementById('sheetSelectWrap');
  dom.sheetSelect = document.getElementById('sheetSelect');
  dom.extractSummaryText = document.getElementById('extractSummaryText');
  dom.mappedColsSummary = document.getElementById('mappedColsSummary');
  dom.btnToggleMapping = document.getElementById('btnToggleMapping');

  dom.mappingPanel = document.getElementById('mappingPanel');
  dom.mappingWarning = document.getElementById('mappingWarning');
  dom.headerRowSelect = document.getElementById('headerRowSelect');
  dom.dateColSelect = document.getElementById('dateColSelect');
  dom.amountModeSelect = document.getElementById('amountModeSelect');
  dom.singleAmountWrap = document.getElementById('singleAmountWrap');
  dom.amountColSelect = document.getElementById('amountColSelect');
  dom.debitColWrap = document.getElementById('debitColWrap');
  dom.debitColSelect = document.getElementById('debitColSelect');
  dom.creditColWrap = document.getElementById('creditColWrap');
  dom.creditColSelect = document.getElementById('creditColSelect');
  dom.descColSelect = document.getElementById('descColSelect');
  dom.refColSelect = document.getElementById('refColSelect');
  dom.includeSummaryCheckbox = document.getElementById('includeSummaryCheckbox');
  dom.btnApplyMapping = document.getElementById('btnApplyMapping');

  // Bước 2: Điều kiện lọc & Tìm tổ hợp
  dom.btnResetFilters = document.getElementById('btnResetFilters');
  dom.samplePresetsBar = document.getElementById('samplePresetsBar');
  dom.presetButtonsList = document.getElementById('presetButtonsList');

  dom.dateFromInput = document.getElementById('dateFromInput');
  dom.dateToInput = document.getElementById('dateToInput');
  dom.dateFromVN = document.getElementById('dateFromVN');
  dom.dateToVN = document.getElementById('dateToVN');
  dom.monthQuickBar = document.getElementById('monthQuickBar');

  dom.targetAmountInput = document.getElementById('targetAmountInput');
  dom.btnClearAmount = document.getElementById('btnClearAmount');
  dom.amountWordsDisplay = document.getElementById('amountWordsDisplay');

  dom.toleranceInput = document.getElementById('toleranceInput');
  dom.toleranceQuickGroup = document.getElementById('toleranceQuickGroup');

  dom.keywordInput = document.getElementById('keywordInput');
  dom.btnClearKeyword = document.getElementById('btnClearKeyword');
  dom.maxItemsSelect = document.getElementById('maxItemsSelect');

  dom.candidatePreviewInfo = document.getElementById('candidatePreviewInfo');
  dom.btnRunSearch = document.getElementById('btnRunSearch');
  dom.btnCancelSearch = document.getElementById('btnCancelSearch');
  dom.searchProgressStrip = document.getElementById('searchProgressStrip');
  dom.searchProgressText = document.getElementById('searchProgressText');

  // Quản lý các Đợt đã gom
  dom.batchesPanel = document.getElementById('batchesPanel');
  dom.batchesPanelTitle = document.getElementById('batchesPanelTitle');
  dom.batchesSummaryBar = document.getElementById('batchesSummaryBar');
  dom.batchesTableBody = document.getElementById('batchesTableBody');
  dom.chkExportUnallocatedOnly = document.getElementById('chkExportUnallocatedOnly');
  dom.btnExportExcelTop = document.getElementById('btnExportExcelTop');
  dom.btnClearAllBatches = document.getElementById('btnClearAllBatches');

  // Bước 3A: Kết quả phương án
  dom.resultsFilterBar = document.getElementById('resultsFilterBar');
  dom.tabResAll = document.getElementById('tabResAll');
  dom.tabResExact = document.getElementById('tabResExact');
  dom.tabResNear = document.getElementById('tabResNear');
  dom.resultsSummaryBanner = document.getElementById('resultsSummaryBanner');
  dom.solutionsContainer = document.getElementById('solutionsContainer');

  // Bước 3B: Bảng Sổ chi tiết gốc
  dom.ledgerTotalBadge = document.getElementById('ledgerTotalBadge');
  dom.btnExportExcelLedger = document.getElementById('btnExportExcelLedger');
  dom.ledgerFilterTabs = document.getElementById('ledgerFilterTabs');
  dom.tabLedgerFiltered = document.getElementById('tabLedgerFiltered');
  dom.tabLedgerAll = document.getElementById('tabLedgerAll');
  dom.tabLedgerUnreconciled = document.getElementById('tabLedgerUnreconciled');
  dom.tabLedgerReconciled = document.getElementById('tabLedgerReconciled');
  dom.tabLedgerInspect = document.getElementById('tabLedgerInspect');

  dom.btnClearPins = document.getElementById('btnClearPins');
  dom.btnClearExcludes = document.getElementById('btnClearExcludes');
  dom.ledgerPageSize = document.getElementById('ledgerPageSize');
  dom.ledgerTableWrap = document.getElementById('ledgerTableWrap');
  dom.chkSelectAllVisible = document.getElementById('chkSelectAllVisible');
  dom.ledgerTableBody = document.getElementById('ledgerTableBody');

  dom.ledgerPaginationBar = document.getElementById('ledgerPaginationBar');
  dom.ledgerPageInfo = document.getElementById('ledgerPageInfo');
  dom.btnPrevPage = document.getElementById('btnPrevPage');
  dom.btnNextPage = document.getElementById('btnNextPage');

  dom.ledgerStatusBar = document.getElementById('ledgerStatusBar');
  dom.statusBarLeft = document.getElementById('statusBarLeft');
  dom.selectionStatusText = document.getElementById('selectionStatusText');
  dom.btnClearChecked = document.getElementById('btnClearChecked');
}

/* ==========================================================================
   TIỆN ÍCH ĐỊNH DẠNG & BẢO VỆ HTML
   ========================================================================== */
function escapeHtml(value) {
  if (value === null || value === undefined) return '';
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function cleanTsvCell(value) {
  if (value === null || value === undefined) return '';
  return String(value).replace(/[\t\r\n]+/g, ' ').trim();
}

function formatAmountHtml(amount, showPlusSign = false) {
  const num = Number(amount || 0);
  const formatted = escapeHtml(formatNumberVN(num, { showPlusSign }));
  const negClass = num < 0 ? ' num-neg' : '';
  return `<span class="num-val${negClass}">${formatted}</span>`;
}

function formatTimestampNowVN() {
  const now = new Date();
  const dd = String(now.getDate()).padStart(2, '0');
  const mm = String(now.getMonth() + 1).padStart(2, '0');
  const yyyy = now.getFullYear();
  const hh = String(now.getHours()).padStart(2, '0');
  const min = String(now.getMinutes()).padStart(2, '0');
  return `${dd}/${mm}/${yyyy} ${hh}:${min}`;
}

function showToast(message, level = 'info', durationMs = 4200) {
  if (!dom.statusToast) return;
  if (toastTimer) {
    clearTimeout(toastTimer);
    toastTimer = null;
  }
  dom.statusToast.textContent = message;
  dom.statusToast.classList.remove('hidden', 'toast-error', 'toast-warn');
  if (level === 'error') {
    dom.statusToast.classList.add('toast-error');
  } else if (level === 'warn') {
    dom.statusToast.classList.add('toast-warn');
  }
  toastTimer = setTimeout(() => {
    dom.statusToast.classList.add('hidden');
  }, durationMs);
}

/**
 * Phát hiện cột Đối tác / Khách hàng phụ (nếu file có cột riêng ngoài cột Diễn giải)
 * để hiển thị đầy đủ thông tin cho kế toán dễ đối chiếu.
 */
function detectAuxiliaryPartnerColumn(columns, mapping) {
  if (!Array.isArray(columns)) return null;
  const usedCols = new Set([
    mapping.dateColIndex,
    mapping.amountColIndex,
    mapping.debitColIndex,
    mapping.creditColIndex,
    mapping.descColIndex,
    mapping.refColIndex,
  ]);

  for (const col of columns) {
    if (usedCols.has(col.index)) continue;
    const norm = normalizeVietnameseSearch(col.label || '');
    if (
      norm.includes('khach hang') ||
      norm.includes('doi tac') ||
      norm.includes('nha cung cap') ||
      norm.includes('ten don vi') ||
      norm.includes('nguoi mua')
    ) {
      return col.index;
    }
  }
  return null;
}

function getDisplayDescription(item) {
  const baseDesc = (item.description || '').trim();
  if (
    state.partnerColIndex !== null &&
    Array.isArray(item.rawCells) &&
    item.rawCells[state.partnerColIndex] !== undefined &&
    item.rawCells[state.partnerColIndex] !== null
  ) {
    const partnerText = String(item.rawCells[state.partnerColIndex]).trim();
    if (partnerText) {
      const normDesc = normalizeVietnameseSearch(baseDesc);
      const normPartner = normalizeVietnameseSearch(partnerText);
      if (normPartner && !normDesc.includes(normPartner)) {
        return baseDesc ? `${partnerText} - ${baseDesc}` : partnerText;
      }
    }
  }
  return baseDesc;
}

function buildItemSearchHaystack(item) {
  const parts = [item.reference || '', item.description || ''];
  if (item.extraCells && typeof item.extraCells === 'object') {
    for (const val of Object.values(item.extraCells)) {
      if (val !== null && val !== undefined && val !== '') {
        parts.push(String(val));
      }
    }
  }
  return normalizeVietnameseSearch(parts.join('\n'));
}

/* ==========================================================================
   CỠ CHỮ: CHUẨN | LỚN (LƯU LOCALSTORAGE)
   ========================================================================== */
function applyFontScale(scale) {
  const mode = scale === 'large' ? 'large' : 'normal';
  document.documentElement.setAttribute('data-font-size', mode);
  dom.btnFontNormal.classList.toggle('active', mode === 'normal');
  dom.btnFontLarge.classList.toggle('active', mode === 'large');
  try {
    localStorage.setItem('af_font_scale', mode);
  } catch (e) {
    // Bỏ qua nếu trình duyệt chặn localStorage
  }
}

function initFontScale() {
  let initial = 'normal';
  try {
    if (localStorage.getItem('af_font_scale') === 'large') {
      initial = 'large';
    }
  } catch (e) {
    initial = document.documentElement.getAttribute('data-font-size') || 'normal';
  }
  applyFontScale(initial);
}

/* ==========================================================================
   BƯỚC 1: NẠP FILE EXCEL, KIỂM TRA CỘT & TRÍCH XUẤT DỮ LIỆU
   ========================================================================== */
function confirmBeforeReplacingData() {
  if (state.reconciledBatches.length > 0) {
    return window.confirm(
      `Bạn đang có ${state.reconciledBatches.length} đợt đã đối soát chưa xuất ra Excel. Mở bảng dữ liệu mới sẽ xóa danh sách đợt hiện tại. Bạn có chắc chắn muốn tiếp tục?`
    );
  }
  return true;
}

function resetReconciliationSession() {
  cancelRunningSearch(true);
  state.pinnedIds.clear();
  state.manualExcludedIds.clear();
  state.checkedIds.clear();
  state.inspectedIds.clear();
  state.inspectLabel = '';
  state.reconciledBatches = [];
  state.reconciledMap.clear();
  state.nextBatchSeq = 1;
  state.searchResult = null;
  state.lastSearchContext = null;
  state.expandedSolutionIds.clear();
  state.solutionDraftNotes.clear();
  state.ledgerView = 'filtered';
  state.ledgerPage = 1;
}

async function handleFileSelected(file) {
  if (!file) return;
  if (!confirmBeforeReplacingData()) {
    dom.excelFileInput.value = '';
    return;
  }

  try {
    const buffer = await file.arrayBuffer();
    const workbook = loadWorkbookFromBuffer(buffer);
    const sheetNames = getSheetNames(workbook);

    if (!Array.isArray(sheetNames) || sheetNames.length === 0) {
      showToast('Tệp Excel không có trang tính (Sheet) nào hợp lệ.', 'error');
      return;
    }

    resetReconciliationSession();
    state.workbook = workbook;
    state.sheetNames = sheetNames;
    state.activeSheetName = sheetNames[0];
    state.fileName = file.name;
    state.isSampleData = false;
    state.samplePresets = [];
    state.activePresetIndex = null;

    renderSamplePresetsBar();
    inspectAndExtractCurrentSheet(null, true);
    showToast(`Đã mở tệp "${file.name}" thành công.`);
  } catch (err) {
    console.error('Lỗi đọc tệp Excel:', err);
    showToast(
      'Không thể đọc tệp đã chọn. Vui lòng kiểm tra tệp có đúng định dạng Excel (.xlsx, .xls, .csv) và không bị đặt mật khẩu bảo vệ.',
      'error',
      6000
    );
  } finally {
    dom.excelFileInput.value = '';
  }
}

function handleLoadSampleDataset() {
  if (!confirmBeforeReplacingData()) return;

  try {
    const sample = getSampleDataset();
    resetReconciliationSession();

    state.workbook = null;
    state.fileName = sample.fileName || 'So_Chi_Tiet_Cong_No_Mau.xlsx';
    state.sheetNames = [sample.sheetName || 'Sheet1'];
    state.activeSheetName = state.sheetNames[0];
    state.isSampleData = true;
    state.samplePresets = Array.isArray(sample.presets) ? sample.presets : [];
    state.activePresetIndex = null;

    const inspection = inspectRawGrid(sample.rawGrid, state.activeSheetName, null);
    applyInspectionResult(inspection, true);
    renderSamplePresetsBar();

    showToast(
      'Đã nạp bảng kê mẫu. Bạn có thể bấm chọn ngay 1 trong 3 ví dụ thực hành ở Bước 2 để xem kết quả.'
    );
  } catch (err) {
    console.error('Lỗi nạp dữ liệu mẫu:', err);
    showToast('Không thể khởi tạo dữ liệu mẫu. Vui lòng tải lại trang.', 'error');
  }
}

function inspectAndExtractCurrentSheet(customHeaderRowIndex = null, resetDates = true) {
  try {
    let inspection;
    if (state.workbook) {
      inspection = inspectSheet(state.workbook, state.activeSheetName, customHeaderRowIndex);
    } else if (state.inspection && Array.isArray(state.inspection.rawGrid)) {
      inspection = inspectRawGrid(
        state.inspection.rawGrid,
        state.activeSheetName,
        customHeaderRowIndex
      );
    } else {
      return;
    }
    applyInspectionResult(inspection, resetDates);
  } catch (err) {
    console.error('Lỗi phân tích trang tính:', err);
    showToast(
      'Không thể phân tích cấu trúc bảng tính này. Vui lòng kiểm tra lại dòng tiêu đề hoặc chọn Sheet khác.',
      'error'
    );
  }
}

function evaluateMappingCompleteness(mapping) {
  const missing = [];
  const hasDate = Number.isInteger(mapping.dateColIndex) && mapping.dateColIndex >= 0;
  const hasDesc = Number.isInteger(mapping.descColIndex) && mapping.descColIndex >= 0;

  let hasAmount = false;
  if (mapping.amountMode === 'debit_minus_credit') {
    hasAmount =
      (Number.isInteger(mapping.debitColIndex) && mapping.debitColIndex >= 0) ||
      (Number.isInteger(mapping.creditColIndex) && mapping.creditColIndex >= 0);
  } else {
    hasAmount = Number.isInteger(mapping.amountColIndex) && mapping.amountColIndex >= 0;
  }

  if (!hasDate) missing.push('Cột Ngày chứng từ');
  if (!hasAmount) missing.push('Cột Số tiền');
  if (!hasDesc) missing.push('Cột Diễn giải');

  return {
    hasDate,
    hasAmount,
    hasDesc,
    allAutoDetected: hasDate && hasAmount && hasDesc,
    missing,
  };
}

function applyInspectionResult(inspection, resetDates) {
  state.inspection = inspection;
  state.headerRowIndex = inspection.detectedHeaderRowIndex ?? 0;

  const suggested = inspection.suggestedMapping || {};
  state.mapping = {
    dateColIndex: suggested.dateColIndex ?? null,
    amountColIndex: suggested.amountColIndex ?? null,
    debitColIndex: suggested.debitColIndex ?? null,
    creditColIndex: suggested.creditColIndex ?? null,
    amountMode: suggested.amountMode === 'debit_minus_credit' ? 'debit_minus_credit' : 'single',
    descColIndex: suggested.descColIndex ?? null,
    refColIndex: suggested.refColIndex ?? null,
  };

  const completeness = evaluateMappingCompleteness(state.mapping);
  // Mặc định thu gọn nếu đã nhận diện đủ Ngày + Số tiền + Diễn giải; tự mở nếu thiếu cột!
  state.isMappingOpen = !completeness.allAutoDetected;

  renderSheetSelector();
  renderMappingControls();
  runDataExtraction(resetDates);
}

function renderSheetSelector() {
  dom.loadedFileBar.classList.remove('hidden');
  dom.fileNameDisplay.textContent = state.fileName || '-';

  if (state.sheetNames.length > 1) {
    dom.sheetSelectWrap.classList.remove('hidden');
  } else {
    // Vẫn hiển thị tên sheet gọn gàng để kế toán dễ kiểm tra
    dom.sheetSelectWrap.classList.remove('hidden');
  }

  dom.sheetSelect.innerHTML = state.sheetNames
    .map(
      (name) =>
        `<option value="${escapeHtml(name)}"${
          name === state.activeSheetName ? ' selected' : ''
        }>${escapeHtml(name)}</option>`
    )
    .join('');
  dom.sheetSelect.disabled = state.sheetNames.length <= 1;
}

function buildColumnOptionsHtml(columns, selectedIndex, allowEmpty, emptyLabel) {
  const opts = [];
  if (allowEmpty) {
    const sel = selectedIndex === null || selectedIndex === undefined ? ' selected' : '';
    opts.push(`<option value=""${sel}>${escapeHtml(emptyLabel)}</option>`);
  }
  for (const col of columns) {
    const sampleText =
      Array.isArray(col.sampleValues) && col.sampleValues.length > 0
        ? ` (VD: ${String(col.sampleValues[0]).slice(0, 26)})`
        : '';
    const display = `[Cột ${col.excelCol}] ${col.label || '(Không tiêu đề)'}${sampleText}`;
    const isSelected = Number(selectedIndex) === col.index ? ' selected' : '';
    opts.push(`<option value="${col.index}"${isSelected}>${escapeHtml(display)}</option>`);
  }
  return opts.join('');
}

function renderMappingControls() {
  if (!state.inspection) return;
  const { previewTopRows = [], columns = [] } = state.inspection;

  // Dòng tiêu đề trong Excel
  const rowOptions = previewTopRows.map((row) => {
    const nonEmptyCells = (row.cells || []).filter((c) => String(c || '').trim() !== '');
    const snippet = nonEmptyCells.slice(0, 4).join(' | ').slice(0, 75) || '(Dòng trống)';
    const selected = row.rowIndex === state.headerRowIndex ? ' selected' : '';
    return `<option value="${row.rowIndex}"${selected}>Dòng ${row.excelRowNumber}: ${escapeHtml(
      snippet
    )}</option>`;
  });
  // Đảm bảo headerRowIndex luôn có mặt kể cả ngoài previewTopRows
  if (!previewTopRows.some((r) => r.rowIndex === state.headerRowIndex)) {
    rowOptions.unshift(
      `<option value="${state.headerRowIndex}" selected>Dòng ${state.headerRowIndex + 1}</option>`
    );
  }
  dom.headerRowSelect.innerHTML = rowOptions.join('');

  // Các cột dữ liệu
  dom.dateColSelect.innerHTML = buildColumnOptionsHtml(
    columns,
    state.mapping.dateColIndex,
    true,
    '-- Chưa chọn cột Ngày --'
  );
  dom.amountModeSelect.value = state.mapping.amountMode;
  dom.amountColSelect.innerHTML = buildColumnOptionsHtml(
    columns,
    state.mapping.amountColIndex,
    true,
    '-- Chọn cột Số tiền --'
  );
  dom.debitColSelect.innerHTML = buildColumnOptionsHtml(
    columns,
    state.mapping.debitColIndex,
    true,
    '-- Chọn cột Phát sinh Nợ (+) --'
  );
  dom.creditColSelect.innerHTML = buildColumnOptionsHtml(
    columns,
    state.mapping.creditColIndex,
    true,
    '-- Chọn cột Phát sinh Có (-) --'
  );
  dom.descColSelect.innerHTML = buildColumnOptionsHtml(
    columns,
    state.mapping.descColIndex,
    true,
    '-- Không dùng / Chưa chọn --'
  );
  dom.refColSelect.innerHTML = buildColumnOptionsHtml(
    columns,
    state.mapping.refColIndex,
    true,
    '-- Không có cột Số CT/HĐ --'
  );

  dom.includeSummaryCheckbox.checked = state.includeSummaryRows;
  syncAmountModeVisibility();
  syncMappingPanelVisibility();
}

function syncAmountModeVisibility() {
  const isDebitCredit = state.mapping.amountMode === 'debit_minus_credit';
  dom.singleAmountWrap.classList.toggle('hidden', isDebitCredit);
  dom.debitColWrap.classList.toggle('hidden', !isDebitCredit);
  dom.creditColWrap.classList.toggle('hidden', !isDebitCredit);
}

function syncMappingPanelVisibility() {
  dom.mappingPanel.classList.toggle('hidden', !state.isMappingOpen);
  dom.btnToggleMapping.setAttribute('aria-expanded', state.isMappingOpen ? 'true' : 'false');
  dom.btnToggleMapping.textContent = state.isMappingOpen
    ? 'Thu gọn thiết lập cột'
    : 'Thiết lập cột & dòng tiêu đề';

  const completeness = evaluateMappingCompleteness(state.mapping);
  if (completeness.missing.length > 0) {
    dom.mappingWarning.classList.remove('hidden');
    dom.mappingWarning.textContent = `Hệ thống cần kiểm tra thêm cột: ${completeness.missing.join(
      ', '
    )}. Vui lòng chọn đúng tên cột trong file Excel của bạn ở bên dưới.`;
  } else {
    dom.mappingWarning.classList.add('hidden');
    dom.mappingWarning.textContent = '';
  }

  updateMappedColsSummaryText();
}

function getColShortTag(colIndex) {
  if (colIndex === null || colIndex === undefined || !state.inspection) return 'Chưa chọn';
  const col = state.inspection.columns.find((c) => c.index === Number(colIndex));
  if (!col) return 'Chưa chọn';
  const shortLabel = (col.label || '').slice(0, 22);
  return `[${col.excelCol}] ${shortLabel}`;
}

function updateMappedColsSummaryText() {
  if (!state.inspection) {
    dom.mappedColsSummary.textContent = '';
    return;
  }
  const dateTag = getColShortTag(state.mapping.dateColIndex);
  const amountTag =
    state.mapping.amountMode === 'debit_minus_credit'
      ? `Nợ ${getColShortTag(state.mapping.debitColIndex)} - Có ${getColShortTag(
          state.mapping.creditColIndex
        )}`
      : getColShortTag(state.mapping.amountColIndex);
  const descTag = getColShortTag(state.mapping.descColIndex);

  dom.mappedColsSummary.textContent = `Ngày: ${dateTag} | Số tiền: ${amountTag} | Diễn giải: ${descTag}`;
}

function parseOptionalIndex(selectEl) {
  const val = selectEl.value;
  if (val === '' || val === null || val === undefined) return null;
  const num = Number(val);
  return Number.isInteger(num) && num >= 0 ? num : null;
}

function readMappingFromControls() {
  state.mapping = {
    dateColIndex: parseOptionalIndex(dom.dateColSelect),
    amountMode:
      dom.amountModeSelect.value === 'debit_minus_credit' ? 'debit_minus_credit' : 'single',
    amountColIndex: parseOptionalIndex(dom.amountColSelect),
    debitColIndex: parseOptionalIndex(dom.debitColSelect),
    creditColIndex: parseOptionalIndex(dom.creditColSelect),
    descColIndex: parseOptionalIndex(dom.descColSelect),
    refColIndex: parseOptionalIndex(dom.refColSelect),
  };
  state.includeSummaryRows = Boolean(dom.includeSummaryCheckbox.checked);
}

function hasAmountMappingChanged(prev, next) {
  if (!prev || !next) return false;
  if (prev.amountMode !== next.amountMode) return true;
  if (next.amountMode === 'debit_minus_credit') {
    return (
      prev.debitColIndex !== next.debitColIndex ||
      prev.creditColIndex !== next.creditColIndex
    );
  }
  return prev.amountColIndex !== next.amountColIndex;
}

function clearBatchesAndResultsOnStructuralChange(reasonMessage = '') {
  const hadBatches = state.reconciledBatches.length > 0;
  cancelRunningSearch(true);
  state.reconciledBatches = [];
  state.reconciledMap.clear();
  state.nextBatchSeq = 1;
  state.searchResult = null;
  state.lastSearchContext = null;
  state.expandedSolutionIds.clear();
  state.solutionDraftNotes.clear();
  state.inspectedIds.clear();
  state.inspectLabel = '';
  if (state.ledgerView === 'inspect') {
    state.ledgerView = 'filtered';
  }
  if (hadBatches && reasonMessage) {
    showToast(reasonMessage, 'info');
  }
}

function runDataExtraction(resetDates = false) {
  if (!state.inspection) return;

  const completeness = evaluateMappingCompleteness(state.mapping);
  syncMappingPanelVisibility();

  state.partnerColIndex = detectAuxiliaryPartnerColumn(
    state.inspection.columns,
    state.mapping
  );

  if (!completeness.hasAmount) {
    state.items = [];
    state.itemById.clear();
    state.summaryRowCount = 0;
    state.availableMonths = [];
    state.pinnedIds.clear();
    state.manualExcludedIds.clear();
    state.checkedIds.clear();
    clearBatchesAndResultsOnStructuralChange();
    dom.extractSummaryText.textContent =
      'Chưa chọn cột Số tiền - Vui lòng chọn cột Số tiền ở khung thiết lập bên dưới.';
    renderMonthQuickButtons();
    updateCandidatePreviewAndLedger();
    renderBatchesPanel();
    renderResultsArea();
    return;
  }

  try {
    const extracted = extractTransactions(
      state.inspection.rawGrid,
      state.headerRowIndex,
      state.mapping,
      state.inspection.columns
    );

    state.items = Array.isArray(extracted.items) ? extracted.items : [];
    state.itemById = new Map(state.items.map((it) => [it.id, it]));
    state.summaryRowCount = extracted.summaryRowCount || 0;
    state.skippedEmptyOrInvalidCount = extracted.skippedEmptyOrInvalidCount || 0;
    state.minDateIso = extracted.minDateIso || null;
    state.maxDateIso = extracted.maxDateIso || null;
    state.availableMonths = Array.isArray(extracted.availableMonths)
      ? extracted.availableMonths
      : [];

    const activeBusinessCount = state.items.filter((it) => !it.isSummaryRow).length;
    dom.extractSummaryText.textContent = `Đã đọc ${formatNumberVN(
      activeBusinessCount
    )} dòng nghiệp vụ (tự động bỏ qua ${formatNumberVN(
      state.summaryRowCount
    )} dòng tổng cộng)`;

    // Dọn các ID ghim / tích chọn không còn hợp lệ (bao gồm dòng Tổng cộng khi tắt includeSummaryRows)
    cleanupOrphanRowSelections();

    if (resetDates) {
      state.dateFromIso = state.minDateIso || '';
      state.dateToIso = state.maxDateIso || '';
      state.activeMonthKey = 'all';
      dom.dateFromInput.value = state.dateFromIso;
      dom.dateToInput.value = state.dateToIso;
    }

    syncDateReadoutsAndMonthHighlight();
    renderMonthQuickButtons();
    dom.btnExportExcelLedger.disabled = state.items.length === 0;

    state.ledgerPage = 1;
    updateCandidatePreviewAndLedger();
    renderBatchesPanel();
    renderResultsArea();
  } catch (err) {
    console.error('Lỗi trích xuất dòng nghiệp vụ:', err);
    showToast(
      'Có lỗi khi đọc các dòng số liệu từ cấu hình cột hiện tại. Vui lòng kiểm tra lại cột Ngày và cột Số tiền.',
      'error'
    );
  }
}

function cleanupOrphanRowSelections() {
  const isRowSelectableNow = (id) => {
    const item = state.itemById.get(id);
    if (!item) return false;
    if (!state.includeSummaryRows && item.isSummaryRow) return false;
    return true;
  };

  for (const id of Array.from(state.pinnedIds)) {
    if (!isRowSelectableNow(id)) state.pinnedIds.delete(id);
  }
  for (const id of Array.from(state.manualExcludedIds)) {
    if (!isRowSelectableNow(id)) state.manualExcludedIds.delete(id);
  }
  for (const id of Array.from(state.checkedIds)) {
    if (!isRowSelectableNow(id)) state.checkedIds.delete(id);
  }
  for (const id of Array.from(state.inspectedIds)) {
    if (!state.itemById.has(id)) state.inspectedIds.delete(id);
  }

  // Đồng bộ tham chiếu các đợt đã đối soát với danh sách items vừa trích xuất lại
  if (state.reconciledBatches.length > 0) {
    const validBatches = [];
    state.reconciledMap.clear();

    for (const batch of state.reconciledBatches) {
      const freshItems = (batch.itemIds || [])
        .map((id) => state.itemById.get(id))
        .filter((it) => Boolean(it) && (state.includeSummaryRows || !it.isSummaryRow));

      if (freshItems.length > 0 && freshItems.length === batch.itemIds.length) {
        batch.items = freshItems;
        validBatches.push(batch);
        for (const it of freshItems) {
          state.reconciledMap.set(it.id, batch);
        }
      }
    }

    state.reconciledBatches = validBatches;
  }
}

/* ==========================================================================
   BƯỚC 2: BỘ LỌC KỲ THÁNG, SỐ TIỀN BẰNG CHỮ, SAI SỐ & PRESETS MẪU
   ========================================================================== */
function renderSamplePresetsBar() {
  if (!state.isSampleData || state.samplePresets.length === 0) {
    dom.samplePresetsBar.classList.add('hidden');
    dom.presetButtonsList.innerHTML = '';
    return;
  }

  dom.samplePresetsBar.classList.remove('hidden');
  dom.presetButtonsList.innerHTML = state.samplePresets
    .map((preset, idx) => {
      const activeCls = state.activePresetIndex === idx ? ' active' : '';
      return `<button type="button" class="preset-btn${activeCls}" data-preset-idx="${idx}">${escapeHtml(
        preset.label
      )}</button>`;
    })
    .join('');
}

function applySamplePreset(presetIndex) {
  const preset = state.samplePresets[presetIndex];
  if (!preset) return;

  state.activePresetIndex = presetIndex;
  renderSamplePresetsBar();

  state.dateFromIso = preset.dateFromIso || state.minDateIso || '';
  state.dateToIso = preset.dateToIso || state.maxDateIso || '';
  dom.dateFromInput.value = state.dateFromIso;
  dom.dateToInput.value = state.dateToIso;

  setTargetAmountProgrammatically(preset.targetAmount);
  setToleranceProgrammatically(
    preset.tolerance !== undefined ? preset.tolerance : DEFAULT_TOLERANCE
  );

  state.keyword = preset.keyword || '';
  dom.keywordInput.value = state.keyword;
  dom.btnClearKeyword.classList.toggle('hidden', !state.keyword);

  syncDateReadoutsAndMonthHighlight();
  updateCandidatePreviewAndLedger();
  executeSearch();
}

function formatMonthButtonLabel(monthObj) {
  const raw = String(monthObj.label || monthObj.key || '').trim();
  if (/^tháng/i.test(raw)) {
    return `${raw} (${monthObj.count})`;
  }
  return `Tháng ${raw} (${monthObj.count})`;
}

function renderMonthQuickButtons() {
  if (!state.items.length || state.availableMonths.length === 0) {
    dom.monthQuickBar.innerHTML =
      '<span class="muted-note">Không có thông tin phân kỳ tháng trong dữ liệu hiện tại</span>';
    return;
  }

  const allActive = state.activeMonthKey === 'all' ? ' active' : '';
  const buttons = [
    `<button type="button" class="quick-chip${allActive}" data-month-key="all">Toàn bộ ngày</button>`,
  ];

  for (const m of state.availableMonths) {
    const isActive = state.activeMonthKey === m.key ? ' active' : '';
    buttons.push(
      `<button type="button" class="quick-chip${isActive}" data-month-key="${escapeHtml(
        m.key
      )}">${escapeHtml(formatMonthButtonLabel(m))}</button>`
    );
  }

  dom.monthQuickBar.innerHTML = buttons.join('');
}

function syncDateReadoutsAndMonthHighlight() {
  const vnFrom = state.dateFromIso ? formatDateVN(state.dateFromIso) : '';
  const vnTo = state.dateToIso ? formatDateVN(state.dateToIso) : '';

  dom.dateFromVN.textContent = `Định dạng VN: ${vnFrom || '--/--/----'}`;
  dom.dateToVN.textContent = `Định dạng VN: ${vnTo || '--/--/----'}`;

  // Tự động xác định nút kỳ tháng nào đang khớp
  let matchedKey = '';
  if (
    state.minDateIso &&
    state.maxDateIso &&
    state.dateFromIso === state.minDateIso &&
    state.dateToIso === state.maxDateIso
  ) {
    matchedKey = 'all';
  } else {
    const matchedMonth = state.availableMonths.find(
      (m) => m.fromIso === state.dateFromIso && m.toIso === state.dateToIso
    );
    if (matchedMonth) matchedKey = matchedMonth.key;
  }
  state.activeMonthKey = matchedKey;

  const chips = dom.monthQuickBar.querySelectorAll('.quick-chip[data-month-key]');
  chips.forEach((btn) => {
    btn.classList.toggle('active', btn.getAttribute('data-month-key') === matchedKey);
  });
}

/**
 * Định dạng trực tiếp ô nhập số tiền chuẩn Việt Nam và giữ nguyên vị trí con trỏ
 */
function handleLiveTargetAmountInput() {
  const input = dom.targetAmountInput;
  const raw = input.value;
  const trimmed = raw.trim();

  if (trimmed === '' || trimmed === '-' || trimmed === '+') {
    state.targetAmount = null;
    input.classList.toggle('is-negative', trimmed === '-');
    updateAmountWordsDisplay();
    updateLedgerStatusBar();
    return;
  }

  const caretPos = input.selectionStart ?? raw.length;
  const sigBeforeCaret = (raw.slice(0, caretPos).match(/[0-9-]/g) || []).length;

  const { numeric, formatted } = formatInputAmountVN(raw);
  state.targetAmount = numeric;
  input.value = formatted;
  input.classList.toggle('is-negative', numeric !== null && numeric < 0);

  // Khôi phục vị trí con trỏ mượt mà
  if (document.activeElement === input) {
    let seenSig = 0;
    let newCaret = formatted.length;
    if (sigBeforeCaret === 0) {
      newCaret = 0;
    } else {
      for (let i = 0; i < formatted.length; i++) {
        if (/[0-9-]/.test(formatted[i])) {
          seenSig++;
          if (seenSig >= sigBeforeCaret) {
            newCaret = i + 1;
            break;
          }
        }
      }
    }
    try {
      input.setSelectionRange(newCaret, newCaret);
    } catch (e) {
      // Bỏ qua nếu input không hỗ trợ
    }
  }

  updateAmountWordsDisplay();
  updateLedgerStatusBar();
}

function setTargetAmountProgrammatically(amount) {
  if (amount === null || amount === undefined || amount === '') {
    state.targetAmount = null;
    dom.targetAmountInput.value = '';
    dom.targetAmountInput.classList.remove('is-negative');
  } else {
    const parsed = typeof amount === 'number' ? amount : parseVietnameseAmount(amount);
    state.targetAmount = parsed;
    dom.targetAmountInput.value = parsed !== null ? formatNumberVN(parsed) : '';
    dom.targetAmountInput.classList.toggle('is-negative', parsed !== null && parsed < 0);
  }
  updateAmountWordsDisplay();
  updateLedgerStatusBar();
}

function updateAmountWordsDisplay() {
  if (state.targetAmount === null || !Number.isFinite(state.targetAmount)) {
    dom.amountWordsDisplay.textContent =
      'Nhập số tiền khách thanh toán trên sao kê ngân hàng để kiểm tra bằng chữ';
    dom.amountWordsDisplay.classList.add('is-empty');
    return;
  }
  const words = readAmountInVietnameseWords(state.targetAmount);
  dom.amountWordsDisplay.textContent = `Bằng chữ: ${words}`;
  dom.amountWordsDisplay.classList.remove('is-empty');
}

function setToleranceProgrammatically(tolValue) {
  const num = Math.max(0, Math.round(Math.abs(Number(tolValue) || 0)));
  state.tolerance = num;
  dom.toleranceInput.value = formatNumberVN(num);
  syncToleranceQuickChips();
}

function handleLiveToleranceInput() {
  const raw = dom.toleranceInput.value.trim();
  if (raw === '') {
    state.tolerance = 0;
    syncToleranceQuickChips();
    return;
  }
  const { numeric } = formatInputAmountVN(raw);
  const absVal = numeric !== null ? Math.max(0, Math.round(Math.abs(numeric))) : 0;
  state.tolerance = absVal;
  dom.toleranceInput.value = formatNumberVN(absVal);
  syncToleranceQuickChips();
}

function syncToleranceQuickChips() {
  const chips = dom.toleranceQuickGroup.querySelectorAll('.quick-chip[data-tolerance]');
  chips.forEach((chip) => {
    const chipVal = Number(chip.getAttribute('data-tolerance'));
    chip.classList.toggle('active', chipVal === state.tolerance);
  });
}

/* ==========================================================================
   KIỂM TRA DÒNG PHÙ HỢP BỘ LỌC (DÙNG CHO PREVIEW & SỔ CHI TIẾT)
   ========================================================================== */
function isFullSheetDateRangeSelected() {
  if (state.activeMonthKey === 'all') return true;
  const coversMin =
    !state.dateFromIso || !state.minDateIso || state.dateFromIso <= state.minDateIso;
  const coversMax =
    !state.dateToIso || !state.maxDateIso || state.dateToIso >= state.maxDateIso;
  return coversMin && coversMax;
}

function doesItemMatchStep2Filter(item) {
  if (!state.includeSummaryRows && item.isSummaryRow) {
    return false;
  }

  const fromTs = state.dateFromIso ? isoToUtcTimestamp(state.dateFromIso) : null;
  const toTs = state.dateToIso ? isoToUtcTimestamp(state.dateToIso) : null;

  if (fromTs !== null || toTs !== null) {
    if (item.timestamp === null || item.timestamp === undefined) {
      if (!isFullSheetDateRangeSelected()) {
        return false;
      }
    } else {
      if (fromTs !== null && item.timestamp < fromTs) return false;
      if (toTs !== null && item.timestamp > toTs) return false;
    }
  }

  const normKw = normalizeVietnameseSearch(state.keyword || '');
  if (normKw) {
    const haystack = buildItemSearchHaystack(item);
    if (!haystack.includes(normKw)) return false;
  }

  return true;
}

function updateCandidatePreviewAndLedger() {
  if (state.items.length === 0) {
    dom.candidatePreviewInfo.textContent = 'Chưa nạp sổ dữ liệu.';
    renderLedgerSection();
    return;
  }

  let eligibleCount = 0;
  for (const item of state.items) {
    if (state.reconciledMap.has(item.id)) continue;
    if (state.manualExcludedIds.has(item.id)) continue;
    if (!state.includeSummaryRows && item.isSummaryRow) continue;
    if (state.pinnedIds.has(item.id) || doesItemMatchStep2Filter(item)) {
      eligibleCount++;
    }
  }

  const pinCount = state.pinnedIds.size;
  const reconciledCount = state.reconciledMap.size;
  const parts = [`Đủ điều kiện xét gom: ${formatNumberVN(eligibleCount)} dòng`];
  if (pinCount > 0) {
    parts.push(`Đang ghim bắt buộc: ${pinCount} dòng`);
  }
  if (reconciledCount > 0) {
    parts.push(`Đã khóa (${state.reconciledBatches.length} đợt): ${reconciledCount} dòng`);
  }
  dom.candidatePreviewInfo.textContent = parts.join(' | ');

  renderLedgerSection();
}

/* ==========================================================================
   CƠ CHẾ TÌM TỔ HỢP: WEB WORKER + FALLBACK TRỰC TIẾP AN TOÀN
   ========================================================================== */
function ensureReconcileWorker() {
  if (state.workerDisabled) return null;
  if (state.worker) return state.worker;

  try {
    const worker = new Worker(new URL('../core/reconcileWorker.js', import.meta.url), {
      type: 'module',
    });

    worker.onmessage = (event) => {
      const msg = event.data;
      if (!msg || msg.requestId !== state.activeRequestId) return;

      if (msg.type === 'PROGRESS') {
        dom.searchProgressText.textContent = `Đang rà soát tổ hợp (${
          msg.currentSize || 1
        } chứng từ)... Đã kiểm tra ${formatNumberVN(
          msg.exploredNodes || 0
        )} nhánh, tìm thấy ${formatNumberVN(msg.foundCount || 0)} phương án.`;
      } else if (msg.type === 'RESULT') {
        finishSearchWithResult(msg.result);
      } else if (msg.type === 'ERROR') {
        finishSearchWithError(
          msg.message || 'Có lỗi xảy ra trong quá trình tìm tổ hợp chứng từ.'
        );
      }
    };

    worker.onerror = (err) => {
      console.warn('Web Worker không khả dụng, chuyển sang chạy trực tiếp:', err);
      try {
        worker.terminate();
      } catch (e) {
        // Bỏ qua
      }
      state.worker = null;
      state.workerDisabled = true;

      if (state.isSearching && state.pendingSearchOptions) {
        runMainThreadFallback(state.activeRequestId, state.pendingSearchOptions);
      }
    };

    state.worker = worker;
    return worker;
  } catch (err) {
    console.warn('Không thể khởi tạo Web Worker, dùng chế độ tính trực tiếp:', err);
    state.workerDisabled = true;
    state.worker = null;
    return null;
  }
}

function runMainThreadFallback(requestId, options) {
  state.cancelFallbackFlag = false;
  dom.searchProgressText.textContent = 'Đang tính toán tổ hợp trực tiếp trên trình duyệt...';

  setTimeout(() => {
    if (requestId !== state.activeRequestId || !state.isSearching) return;
    try {
      const result = findCombinations(
        options,
        (prog) => {
          if (prog && typeof prog.exploredNodes === 'number') {
            dom.searchProgressText.textContent = `Đang kiểm tra ${formatNumberVN(
              prog.exploredNodes
            )} nhánh...`;
          }
        },
        () => state.cancelFallbackFlag
      );
      if (requestId !== state.activeRequestId) return;
      finishSearchWithResult(result);
    } catch (err) {
      console.error('Lỗi tìm kiếm trực tiếp:', err);
      finishSearchWithError('Không thể hoàn tất phép tìm tổ hợp. Vui lòng kiểm tra lại dữ liệu đầu vào.');
    }
  }, 25);
}

function executeSearch() {
  if (state.items.length === 0) {
    showToast(
      'Vui lòng chọn file Excel từ máy hoặc bấm "Nạp bảng kê mẫu để thử" ở Bước 1 trước khi tìm.',
      'warn'
    );
    return;
  }

  if (state.targetAmount === null || !Number.isFinite(state.targetAmount)) {
    dom.targetAmountInput.focus();
    showToast('Vui lòng nhập Số tiền cần gom (Z) ở Bước 2.', 'warn');
    return;
  }

  // Tự động hoán đổi nếu kế toán chọn Từ ngày > Đến ngày
  if (state.dateFromIso && state.dateToIso && state.dateFromIso > state.dateToIso) {
    const tmp = state.dateFromIso;
    state.dateFromIso = state.dateToIso;
    state.dateToIso = tmp;
    dom.dateFromInput.value = state.dateFromIso;
    dom.dateToInput.value = state.dateToIso;
    syncDateReadoutsAndMonthHighlight();
    state.ledgerPage = 1;
    updateCandidatePreviewAndLedger();
    showToast('Đã tự động đảo lại thứ tự Từ ngày trước Đến ngày.');
  }

  const safePinnedIds = Array.from(state.pinnedIds).filter((id) => {
    const it = state.itemById.get(id);
    return Boolean(it) && (state.includeSummaryRows || !it.isSummaryRow);
  });

  // Tập hợp danh sách loại trừ = Loại trừ thủ công + Các dòng đã gom ở các đợt trước
  const excludedSet = new Set(state.manualExcludedIds);
  for (const recId of state.reconciledMap.keys()) {
    excludedSet.add(recId);
  }
  for (const pinId of safePinnedIds) {
    excludedSet.delete(pinId);
  }

  const isFullRange = isFullSheetDateRangeSelected();
  const options = {
    items: state.items,
    dateFromIso: isFullRange ? null : state.dateFromIso || null,
    dateToIso: isFullRange ? null : state.dateToIso || null,
    keyword: (state.keyword || '').trim(),
    targetAmount: state.targetAmount,
    tolerance: state.tolerance,
    pinnedIds: safePinnedIds,
    excludedIds: Array.from(excludedSet),
    includeSummaryRows: state.includeSummaryRows,
    minItems: 1,
    maxItems: state.maxItems,
    maxResults: DEFAULT_MAX_RESULTS,
    timeLimitMs: DEFAULT_TIME_LIMIT_MS,
  };

  state.activeRequestId += 1;
  const requestId = state.activeRequestId;
  state.isSearching = true;
  state.pendingSearchOptions = options;
  state.lastSearchContext = {
    dateFromIso: state.dateFromIso || null,
    dateToIso: state.dateToIso || null,
    targetAmount: options.targetAmount,
    tolerance: options.tolerance,
    keyword: options.keyword,
  };

  setSearchingUiState(true);

  const worker = ensureReconcileWorker();
  if (worker) {
    try {
      worker.postMessage({
        type: 'SEARCH',
        requestId,
        options,
      });
    } catch (postErr) {
      console.warn('Gửi dữ liệu qua Worker thất bại, chuyển sang chạy trực tiếp:', postErr);
      state.workerDisabled = true;
      state.worker = null;
      runMainThreadFallback(requestId, options);
    }
  } else {
    runMainThreadFallback(requestId, options);
  }
}

function cancelRunningSearch(silent = false) {
  if (!state.isSearching) return;
  const cancelledReqId = state.activeRequestId;
  state.activeRequestId += 1;
  state.cancelFallbackFlag = true;
  state.isSearching = false;
  state.pendingSearchOptions = null;

  if (state.worker) {
    try {
      state.worker.postMessage({ type: 'CANCEL', requestId: cancelledReqId });
    } catch (e) {
      // Bỏ qua
    }
    // Ngắt ngay lập tức nếu Worker đang bận vòng lặp sâu
    if (state.cancelFallbackTimer) clearTimeout(state.cancelFallbackTimer);
    state.cancelFallbackTimer = setTimeout(() => {
      if (state.worker && !state.isSearching) {
        try {
          state.worker.terminate();
        } catch (e) {
          // Bỏ qua
        }
        state.worker = null;
      }
    }, 120);
  }

  setSearchingUiState(false);
  if (!silent) {
    showToast('Đã dừng phép tìm kiếm theo yêu cầu.');
  }
}

function setSearchingUiState(searching) {
  dom.btnRunSearch.disabled = searching;
  dom.btnRunSearch.textContent = searching
    ? 'ĐANG TÌM TỔ HỢP...'
    : 'TÌM TỔ HỢP KHỚP SỐ TIỀN';
  dom.btnCancelSearch.classList.toggle('hidden', !searching);
  dom.searchProgressStrip.classList.toggle('hidden', !searching);
  if (searching) {
    dom.searchProgressText.textContent = 'Đang kiểm tra các tổ hợp chứng từ phù hợp...';
  }
}

function finishSearchWithResult(result) {
  state.isSearching = false;
  state.pendingSearchOptions = null;
  setSearchingUiState(false);

  state.searchResult = result || { solutions: [], stats: {} };
  state.resultsFilter = 'all';
  state.expandedSolutionIds.clear();

  const solutions = state.searchResult.solutions || [];
  if (solutions.length > 0) {
    // Mặc định mở sẵn Phương án #1 để kế toán nhìn thấy chi tiết ngay
    state.expandedSolutionIds.add(solutions[0].id);
  }

  renderResultsArea();
}

function finishSearchWithError(errorMessage) {
  state.isSearching = false;
  state.pendingSearchOptions = null;
  setSearchingUiState(false);

  state.searchResult = {
    solutions: [],
    stats: {
      candidateCount: 0,
      pinnedCount: state.pinnedIds.size,
      exploredNodes: 0,
      elapsedMs: 0,
      timedOut: false,
      exactCount: 0,
      nearCount: 0,
    },
    error: errorMessage,
  };
  renderResultsArea();
  showToast(errorMessage, 'error');
}

/* ==========================================================================
   BƯỚC 3A: HIỂN THỊ KẾT QUẢ PHƯƠNG ÁN, SAO CHÉP EXCEL & XÁC NHẬN ĐỢT GOM
   ========================================================================== */
function renderResultsArea() {
  if (!state.searchResult || !state.lastSearchContext) {
    dom.resultsFilterBar.classList.add('hidden');
    dom.resultsSummaryBanner.classList.add('hidden');
    dom.solutionsContainer.innerHTML = `
      <div class="empty-state-box">
        <div class="empty-state-title">Chưa thực hiện tìm tổ hợp gom tiền</div>
        <p class="empty-state-desc">
          Vui lòng nạp file Excel ở <strong>Bước 1</strong> (hoặc bấm <strong>Nạp bảng kê mẫu để thử</strong>), nhập <strong>Số tiền cần gom</strong> ở <strong>Bước 2</strong> và bấm nút <strong>TÌM TỔ HỢP KHỚP SỐ TIỀN</strong>.
        </p>
      </div>
    `;
    return;
  }

  const { solutions = [], stats = {}, error = null } = state.searchResult;
  const ctx = state.lastSearchContext;

  if (error) {
    dom.resultsFilterBar.classList.add('hidden');
    dom.resultsSummaryBanner.classList.remove('hidden', 'has-exact');
    dom.resultsSummaryBanner.classList.add('no-match');
    dom.resultsSummaryBanner.innerHTML = `<strong>Không thể tìm tổ hợp phù hợp:</strong> ${escapeHtml(
      error
    )}`;
    dom.solutionsContainer.innerHTML = `
      <div class="empty-state-box">
        <div class="empty-state-title">Vui lòng điều chỉnh lại điều kiện lọc</div>
        <p class="empty-state-desc">${escapeHtml(error)}</p>
      </div>
    `;
    return;
  }

  const totalCount = solutions.length;
  const exactCount = solutions.filter((s) => s.isExact || s.diff === 0).length;
  const nearCount = totalCount - exactCount;
  const elapsedSec = ((stats.elapsedMs || 0) / 1000).toFixed(2).replace('.', ',');

  let periodLabel = 'trên toàn bộ sổ';
  if (ctx.dateFromIso || ctx.dateToIso) {
    const fStr = ctx.dateFromIso ? formatDateVN(ctx.dateFromIso) : 'Đầu sổ';
    const tStr = ctx.dateToIso ? formatDateVN(ctx.dateToIso) : 'Cuối sổ';
    periodLabel = `trong khoảng <strong>${escapeHtml(fStr)} – ${escapeHtml(tStr)}</strong>`;
  }

  const tolFormatted = formatNumberVN(ctx.tolerance);

  // Thanh tóm tắt kết quả
  dom.resultsSummaryBanner.classList.remove('hidden', 'has-exact', 'no-match');
  if (exactCount > 0) {
    dom.resultsSummaryBanner.classList.add('has-exact');
  } else if (totalCount === 0) {
    dom.resultsSummaryBanner.classList.add('no-match');
  }

  const timeoutHtml = stats.timedOut
    ? `<div class="timeout-note">Lưu ý: Phép tìm dừng ở giới hạn ${
        DEFAULT_TIME_LIMIT_MS / 1000
      } giây do số lượng chứng từ lớn và đã hiển thị các phương án khớp nhất tìm được. Bạn có thể thu hẹp khoảng ngày hoặc nhập thêm từ khóa tên khách hàng để kiểm tra sâu hơn.</div>`
    : '';

  dom.resultsSummaryBanner.innerHTML = `
    <div>
      Tìm thấy <strong>${totalCount} phương án</strong> ${periodLabel}
      (<strong>${exactCount}</strong> phương án khớp chính xác 0đ,
      <strong>${nearCount}</strong> phương án gần khớp trong ±${escapeHtml(tolFormatted)}đ)
      — Thời gian: <strong>${elapsedSec} giây</strong>
    </div>
    ${timeoutHtml}
  `;

  // Bộ lọc nhanh kết quả
  if (totalCount > 0) {
    dom.resultsFilterBar.classList.remove('hidden');
    dom.tabResAll.textContent = `Tất cả (${totalCount})`;
    dom.tabResExact.textContent = `Chỉ khớp chính xác = 0 (${exactCount})`;
    dom.tabResNear.textContent = `Gần khớp ±${tolFormatted} (${nearCount})`;

    dom.tabResAll.classList.toggle('active', state.resultsFilter === 'all');
    dom.tabResExact.classList.toggle('active', state.resultsFilter === 'exact');
    dom.tabResNear.classList.toggle('active', state.resultsFilter === 'near');
  } else {
    dom.resultsFilterBar.classList.add('hidden');
  }

  // Trường hợp không có phương án nào
  if (totalCount === 0) {
    dom.solutionsContainer.innerHTML = `
      <div class="empty-state-box">
        <div class="empty-state-title">Không tìm thấy tổ hợp chứng từ nào khớp với số tiền ${escapeHtml(
          formatNumberVN(ctx.targetAmount)
        )} đ (±${escapeHtml(tolFormatted)} đ)</div>
        <ul class="empty-hints-list">
          <li>Kiểm tra lại <strong>Khoảng thời gian (Từ ngày – Đến ngày)</strong> xem có hóa đơn kỳ trước thanh toán gộp sang kỳ này hay không (thử bấm <strong>Toàn bộ ngày</strong>).</li>
          <li>Thử tăng <strong>Độ lệch cho phép (± Sai số)</strong> lên <strong>± 1.000</strong> hoặc <strong>± 5.000</strong> nếu khách hàng trừ phí chuyển khoản ngân hàng hoặc làm tròn lẻ thuế VAT.</li>
          <li>Kiểm tra xem có chứng từ nào đang bị <strong>Loại trừ</strong>, <strong>Ghim bắt buộc</strong> nhầm, hoặc đã được khóa ở đợt đối soát trước hay không.</li>
          <li>Nếu khách hàng thanh toán quá nhiều hóa đơn nhỏ cùng lúc, hãy tăng mục <strong>Giới hạn số dòng gom tối đa</strong> lên 20 hoặc 25 dòng.</li>
        </ul>
      </div>
    `;
    return;
  }

  // Lọc theo tab kết quả
  const filteredSolutions = solutions
    .map((sol, originalIdx) => ({ sol, rank: originalIdx + 1 }))
    .filter(({ sol }) => {
      if (state.resultsFilter === 'exact') return sol.isExact || sol.diff === 0;
      if (state.resultsFilter === 'near') return !(sol.isExact || sol.diff === 0);
      return true;
    });

  if (filteredSolutions.length === 0) {
    dom.solutionsContainer.innerHTML = `
      <div class="empty-state-box">
        <div class="empty-state-title">Không có phương án nào thuộc nhóm lọc đang chọn</div>
        <p class="empty-state-desc">Hãy chuyển sang tab <strong>Tất cả (${totalCount})</strong> để xem các phương án còn lại.</p>
      </div>
    `;
    return;
  }

  dom.solutionsContainer.innerHTML = filteredSolutions
    .map(({ sol, rank }) => buildSolutionCardHtml(sol, rank))
    .join('');
}

function getSolutionReconcileState(sol) {
  const batchesHit = new Set();
  let matchedCount = 0;
  for (const id of sol.itemIds) {
    const b = state.reconciledMap.get(id);
    if (b) {
      matchedCount++;
      batchesHit.add(b);
    }
  }
  if (matchedCount === 0) {
    return { status: 'available', batch: null };
  }
  if (
    matchedCount === sol.itemIds.length &&
    batchesHit.size === 1
  ) {
    const singleBatch = Array.from(batchesHit)[0];
    if (singleBatch.itemIds.length === sol.itemIds.length) {
      return { status: 'confirmed_exact', batch: singleBatch };
    }
  }
  return { status: 'overlapping', batch: Array.from(batchesHit)[0] };
}

function buildSolutionCardHtml(sol, rank) {
  const isExpanded = state.expandedSolutionIds.has(sol.id);
  const isExact = sol.isExact || sol.diff === 0;
  const recState = getSolutionReconcileState(sol);

  const badgeHtml = isExact
    ? '<span class="match-badge match-badge-exact">KHỚP CHÍNH XÁC (Lệch 0)</span>'
    : `<span class="match-badge match-badge-near">GẦN KHỚP (Lệch ${escapeHtml(
        formatNumberVN(sol.diff, { showPlusSign: true })
      )})</span>`;

  let extraStateBadge = '';
  if (recState.status === 'confirmed_exact' && recState.batch) {
    extraStateBadge = `<span class="status-tag status-reconciled">Đã xác nhận (${escapeHtml(
      recState.batch.label
    )})</span>`;
  } else if (recState.status === 'overlapping') {
    extraStateBadge =
      '<span class="match-badge match-badge-stale">Có dòng đã thuộc đợt khác</span>';
  }

  const cardClasses = [
    'solution-card',
    isExact ? 'solution-exact' : 'solution-near',
    isExpanded ? 'solution-expanded' : '',
  ]
    .filter(Boolean)
    .join(' ');

  let bodyHtml = '';
  if (isExpanded) {
    const draftNote = state.solutionDraftNotes.get(sol.id) || '';
    const disableConfirm = recState.status !== 'available';
    const confirmLabel =
      recState.status === 'confirmed_exact' && recState.batch
        ? `Đã lưu vào ${recState.batch.label}`
        : 'Xác nhận Gom đợt này (Đánh dấu Đã đối soát)';

    const rowsHtml = (sol.items || [])
      .map((item, idx) => {
        const dateStr = item.dateDisplay || formatDateVN(item.dateIso) || '-';
        const refStr = item.reference || '-';
        const descStr = getDisplayDescription(item) || '-';
        const isPinned = state.pinnedIds.has(item.id);
        const pinNote = isPinned
          ? ' <span class="status-tag status-pinned">Đã ghim</span>'
          : '';

        return `
          <tr>
            <td class="col-center">${idx + 1}</td>
            <td class="col-excel-row">#${item.excelRowNumber}</td>
            <td class="col-date">${escapeHtml(dateStr)}</td>
            <td class="col-ref">${escapeHtml(refStr)}</td>
            <td class="col-desc">${escapeHtml(descStr)}${pinNote}</td>
            <td class="col-num">${formatAmountHtml(item.amount)}</td>
          </tr>
        `;
      })
      .join('');

    const diffText =
      sol.diff === 0
        ? '<span class="num-val num-zero">0 (Khớp 100%)</span>'
        : `<span class="num-val num-neg">${escapeHtml(
            formatNumberVN(sol.diff, { showPlusSign: true })
          )}</span>`;

    const duplicateHtml = sol.duplicateVariantsNote
      ? `<div class="duplicate-variant-note"><strong>Lưu ý chứng từ cùng mệnh giá:</strong> ${escapeHtml(
          sol.duplicateVariantsNote
        )}</div>`
      : '';

    bodyHtml = `
      <div class="solution-body">
        <div class="solution-actions-bar">
          <div class="solution-note-wrap">
            <label for="solNote_${escapeHtml(sol.id)}">Ghi chú đợt (Số UNC / KH):</label>
            <input
              type="text"
              id="solNote_${escapeHtml(sol.id)}"
              class="form-input solution-note-input"
              data-sol-note-id="${escapeHtml(sol.id)}"
              value="${escapeHtml(draftNote)}"
              placeholder="VD: UNC Vietcombank 15/03, KH Thanh Dat..."
            >
          </div>
          <div class="solution-action-btns">
            <button
              type="button"
              class="btn btn-sm"
              data-sol-action="inspect"
              data-sol-id="${escapeHtml(sol.id)}"
              data-sol-rank="${rank}"
            >
              Đối chiếu trên sổ gốc
            </button>
            <button
              type="button"
              class="btn btn-secondary btn-sm"
              data-sol-action="copy"
              data-sol-id="${escapeHtml(sol.id)}"
            >
              Sao chép sang Excel
            </button>
            <button
              type="button"
              class="btn btn-primary btn-sm"
              data-sol-action="confirm"
              data-sol-id="${escapeHtml(sol.id)}"
              ${disableConfirm ? 'disabled' : ''}
            >
              ${escapeHtml(confirmLabel)}
            </button>
          </div>
        </div>

        <div class="table-scroll-wrap">
          <table class="grid-table">
            <thead>
              <tr>
                <th class="col-center" style="width: 52px;">STT</th>
                <th class="col-excel-row">Dòng Excel</th>
                <th class="col-date">Ngày CT</th>
                <th class="col-ref">Số CT / HĐ</th>
                <th class="col-desc">Diễn giải</th>
                <th class="col-num col-amount">Số tiền (VNĐ)</th>
              </tr>
            </thead>
            <tbody>
              ${rowsHtml}
            </tbody>
            <tfoot>
              <tr>
                <td colspan="5" style="text-align: right;">
                  TỔNG CỘNG PHƯƠNG ÁN #${rank} (${sol.itemCount} dòng) — Số cần gom: ${escapeHtml(
      formatNumberVN(sol.targetAmount)
    )} | Chênh lệch: ${diffText}
                </td>
                <td class="col-num">${formatAmountHtml(sol.totalAmount)}</td>
              </tr>
            </tfoot>
          </table>
        </div>
        ${duplicateHtml}
      </div>
    `;
  }

  return `
    <div class="${cardClasses}" data-solution-card="${escapeHtml(sol.id)}">
      <button
        type="button"
        class="solution-header"
        data-sol-action="toggle"
        data-sol-id="${escapeHtml(sol.id)}"
        aria-expanded="${isExpanded ? 'true' : 'false'}"
      >
        <div class="solution-header-left">
          <span class="solution-rank">Phương án #${rank}</span>
          ${badgeHtml}
          ${extraStateBadge}
          <span class="solution-meta-sep">|</span>
          <span class="solution-meta-item">Gồm <strong>${sol.itemCount} dòng</strong></span>
          <span class="solution-meta-sep">|</span>
          <span class="solution-meta-item">Tổng cộng: <strong>${formatAmountHtml(
            sol.totalAmount
          )}</strong></span>
        </div>
        <div class="solution-header-right">
          <span class="solution-toggle-label">${
            isExpanded ? 'Thu gọn chi tiết' : 'Xem chi tiết các dòng'
          }</span>
        </div>
      </button>
      ${bodyHtml}
    </div>
  `;
}

async function copyTextToClipboard(text) {
  if (navigator.clipboard && typeof navigator.clipboard.writeText === 'function') {
    try {
      await navigator.clipboard.writeText(text);
      return true;
    } catch (e) {
      // Chuyển sang fallback bên dưới
    }
  }
  const textarea = document.createElement('textarea');
  textarea.value = text;
  textarea.setAttribute('readonly', '');
  textarea.style.position = 'fixed';
  textarea.style.left = '-9999px';
  document.body.appendChild(textarea);
  textarea.select();
  let ok = false;
  try {
    ok = document.execCommand('copy');
  } catch (e) {
    ok = false;
  }
  document.body.removeChild(textarea);
  return ok;
}

async function handleCopySolutionToExcel(sol) {
  if (!sol || !Array.isArray(sol.items)) return;

  const lines = [];
  lines.push(['Dòng Excel', 'Ngày CT', 'Số CT/HĐ', 'Diễn giải', 'Số tiền'].join('\t'));

  for (const item of sol.items) {
    const rowNum = `#${item.excelRowNumber}`;
    const dateStr = item.dateDisplay || formatDateVN(item.dateIso) || '';
    const refStr = cleanTsvCell(item.reference);
    const descStr = cleanTsvCell(getDisplayDescription(item));
    const rawNum = Number(item.amount || 0);
    lines.push([rowNum, dateStr, refStr, descStr, String(rawNum)].join('\t'));
  }

  lines.push(
    ['', '', '', `TỔNG CỘNG (${sol.itemCount} dòng)`, String(Number(sol.totalAmount || 0))].join(
      '\t'
    )
  );
  if (sol.diff !== 0) {
    lines.push(
      ['', '', '', 'Số tiền cần gom', String(Number(sol.targetAmount || 0))].join('\t')
    );
    lines.push(['', '', '', 'Chênh lệch', String(Number(sol.diff || 0))].join('\t'));
  }

  const tsvText = lines.join('\r\n');
  const copied = await copyTextToClipboard(tsvText);
  if (copied) {
    showToast(
      `Đã sao chép ${sol.itemCount} dòng vào bộ nhớ tạm. Bạn có thể sang Excel nhấn Ctrl+V để dán ngay.`
    );
  } else {
    showToast(
      'Trình duyệt không cho phép sao chép tự động. Vui lòng thử lại hoặc dùng nút Xuất File Excel.',
      'error'
    );
  }
}

function handleConfirmSolutionAsBatch(sol) {
  if (!sol) return;
  const recState = getSolutionReconcileState(sol);
  if (recState.status !== 'available') {
    showToast(
      'Phương án này có chứng từ đã được đưa vào đợt đối soát trước đó. Hãy bấm Tìm lại để cập nhật tổ hợp mới.',
      'warn'
    );
    return;
  }

  const seq = state.nextBatchSeq++;
  const label = `Đợt #${seq}`;
  const noteText = (state.solutionDraftNotes.get(sol.id) || '').trim();
  const createdAt = formatTimestampNowVN();

  const resolvedItems = sol.itemIds
    .map((id) => state.itemById.get(id))
    .filter(Boolean);

  const batch = {
    id: `batch_${seq}`,
    batchId: label,
    batchNumber: seq,
    code: label,
    label,
    name: label,
    itemIds: [...sol.itemIds],
    items: resolvedItems,
    itemCount: sol.itemIds.length,
    targetAmount: sol.targetAmount,
    totalAmount: sol.totalAmount,
    diff: sol.diff,
    absDiff: sol.absDiff,
    isExact: sol.isExact,
    tolerance: state.lastSearchContext?.tolerance ?? state.tolerance,
    dateFromIso: state.lastSearchContext?.dateFromIso || null,
    dateToIso: state.lastSearchContext?.dateToIso || null,
    keyword: state.lastSearchContext?.keyword || '',
    note: noteText,
    createdAt,
    reconciledAt: createdAt,
  };

  state.reconciledBatches.push(batch);
  for (const itemId of batch.itemIds) {
    state.reconciledMap.set(itemId, batch);
    // Gỡ ghim / bỏ chọn thủ công đối với các dòng đã hoàn tất đối soát
    state.pinnedIds.delete(itemId);
    state.checkedIds.delete(itemId);
  }

  renderBatchesPanel();
  renderResultsArea();
  updateCandidatePreviewAndLedger();

  showToast(
    `Đã xác nhận ${label} (${batch.itemCount} dòng - Tổng ${formatNumberVN(
      batch.totalAmount
    )} đ). Các dòng này đã được đánh dấu Đã đối soát và loại khỏi lần tìm kế tiếp.`
  );
}

/* ==========================================================================
   QUẢN LÝ CÁC ĐỢT ĐÃ ĐỐI SOÁT & XUẤT FILE EXCEL (.XLSX)
   ========================================================================== */
function renderBatchesPanel() {
  const count = state.reconciledBatches.length;
  if (count === 0) {
    dom.batchesPanel.classList.add('hidden');
    return;
  }

  dom.batchesPanel.classList.remove('hidden');
  dom.batchesPanelTitle.textContent = `DANH SÁCH CÁC ĐỢT ĐÃ ĐỐI SOÁT (${count} ĐỢT)`;

  let totalMatchedRows = 0;
  let totalMatchedAmount = 0;
  for (const b of state.reconciledBatches) {
    totalMatchedRows += b.itemCount;
    totalMatchedAmount += Number(b.totalAmount || 0);
  }

  const businessTotal = state.items.filter((it) => !it.isSummaryRow).length;
  const remainingRows = Math.max(0, businessTotal - state.reconciledMap.size);

  dom.batchesSummaryBar.innerHTML = `
    <span>Tổng số đợt đã gom: <strong>${count} đợt</strong></span>
    <span class="statusbar-sep">|</span>
    <span>Tổng chứng từ đã đối soát: <strong>${formatNumberVN(totalMatchedRows)} dòng</strong></span>
    <span class="statusbar-sep">|</span>
    <span>Tổng tiền đã đối soát: <strong>${formatAmountHtml(totalMatchedAmount)} đ</strong></span>
    <span class="statusbar-sep">|</span>
    <span>Còn lại chưa đối soát: <strong>${formatNumberVN(remainingRows)} dòng</strong></span>
  `;

  dom.batchesTableBody.innerHTML = state.reconciledBatches
    .map((b) => {
      const diffHtml =
        b.diff === 0
          ? '<span class="num-val num-zero">0</span>'
          : formatAmountHtml(b.diff, true);

      return `
        <tr>
          <td class="col-center"><span class="status-tag status-reconciled">${escapeHtml(
            b.label
          )}</span></td>
          <td class="col-center num-val">${escapeHtml(b.createdAt)}</td>
          <td class="col-center num-val"><strong>${b.itemCount}</strong></td>
          <td class="col-num">${formatAmountHtml(b.targetAmount)}</td>
          <td class="col-num">${formatAmountHtml(b.totalAmount)}</td>
          <td class="col-num">${diffHtml}</td>
          <td>
            <input
              type="text"
              class="form-input batch-note-edit"
              data-batch-note-id="${escapeHtml(b.id)}"
              value="${escapeHtml(b.note || '')}"
              placeholder="Nhập ghi chú đợt (số UNC, ngày sao kê...)"
            >
          </td>
          <td class="col-center">
            <div class="row-action-group">
              <button
                type="button"
                class="btn btn-xs"
                data-batch-action="inspect"
                data-batch-id="${escapeHtml(b.id)}"
              >
                Xem các dòng
              </button>
              <button
                type="button"
                class="btn btn-xs btn-danger"
                data-batch-action="undo"
                data-batch-id="${escapeHtml(b.id)}"
              >
                Hoàn tác / Hủy đợt
              </button>
            </div>
          </td>
        </tr>
      `;
    })
    .join('');
}

function undoSingleBatch(batchId) {
  const idx = state.reconciledBatches.findIndex((b) => b.id === batchId);
  if (idx === -1) return;
  const batch = state.reconciledBatches[idx];

  state.reconciledBatches.splice(idx, 1);
  for (const itemId of batch.itemIds) {
    if (state.reconciledMap.get(itemId)?.id === batch.id) {
      state.reconciledMap.delete(itemId);
    }
  }

  if (state.ledgerView === 'inspect' && state.inspectLabel === batch.label) {
    state.inspectedIds.clear();
    state.inspectLabel = '';
    state.ledgerView = 'filtered';
  }

  renderBatchesPanel();
  renderResultsArea();
  updateCandidatePreviewAndLedger();
  showToast(`Đã hoàn tác ${batch.label}. Các dòng chứng từ đã mở khóa về Chưa đối soát.`);
}

function handleClearAllBatches() {
  if (state.reconciledBatches.length === 0) return;
  const confirmed = window.confirm(
    `Bạn có chắc chắn muốn hoàn tác toàn bộ ${state.reconciledBatches.length} đợt đã đối soát?`
  );
  if (!confirmed) return;

  state.reconciledBatches = [];
  state.reconciledMap.clear();
  state.nextBatchSeq = 1;
  state.inspectedIds.clear();
  state.inspectLabel = '';
  if (state.ledgerView === 'inspect') state.ledgerView = 'filtered';

  renderBatchesPanel();
  renderResultsArea();
  updateCandidatePreviewAndLedger();
  showToast('Đã hoàn tác toàn bộ các đợt đối soát.');
}

function handleExportReconciledExcel() {
  if (state.items.length === 0) {
    showToast('Chưa có dữ liệu sổ để xuất ra Excel.', 'warn');
    return;
  }

  try {
    const unallocatedOnly = Boolean(dom.chkExportUnallocatedOnly?.checked);
    const bytes = exportReconciledWorkbook({
      sheetName: state.activeSheetName || 'Chi_Tiet_Doi_Soat',
      columns: state.inspection?.columns || [],
      items: state.items,
      reconciledBatches: state.reconciledBatches,
      unallocatedOnly,
    });

    const blob = new Blob([bytes], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });

    const baseName = (state.fileName || 'So_Cong_No')
      .replace(/\.(xlsx|xls|csv)$/i, '')
      .replace(/[^a-zA-Z0-9_-]+/g, '_');
    const now = new Date();
    const dateStamp = `${String(now.getDate()).padStart(2, '0')}${String(
      now.getMonth() + 1
    ).padStart(2, '0')}${now.getFullYear()}`;
    const outFileName = `Ket_Qua_Doi_Soat_${baseName}_${dateStamp}.xlsx`;

    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = outFileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    setTimeout(() => URL.revokeObjectURL(url), 1500);

    showToast(`Đã xuất tệp Excel "${outFileName}" thành công.`);
  } catch (err) {
    console.error('Lỗi xuất file Excel đối soát:', err);
    showToast(
      'Không thể tạo tệp Excel kết quả. Vui lòng thử lại.',
      'error'
    );
  }
}

/* ==========================================================================
   BƯỚC 3B: BẢNG DỮ LIỆU SỔ CHI TIẾT & THANH STATUS BAR TÍNH TỔNG TỨC THỜI
   ========================================================================== */
function inspectIdsInLedger(itemIds, label) {
  state.inspectedIds = new Set(itemIds);
  state.inspectLabel = label || 'Phương án đang chọn';
  state.ledgerView = 'inspect';
  state.ledgerPage = 1;
  renderLedgerSection();

  const ledgerPanel = document.getElementById('ledgerPanel');
  if (ledgerPanel && typeof ledgerPanel.scrollIntoView === 'function') {
    ledgerPanel.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }
}

function getLedgerItemsByView(viewMode) {
  switch (viewMode) {
    case 'all':
      return state.items;
    case 'unreconciled':
      return state.items.filter(
        (it) =>
          !state.reconciledMap.has(it.id) &&
          (state.includeSummaryRows || !it.isSummaryRow)
      );
    case 'reconciled':
      return state.items.filter((it) => state.reconciledMap.has(it.id));
    case 'inspect':
      return state.items.filter((it) => state.inspectedIds.has(it.id));
    case 'filtered':
    default:
      return state.items.filter(
        (it) =>
          (state.pinnedIds.has(it.id) && (state.includeSummaryRows || !it.isSummaryRow)) ||
          doesItemMatchStep2Filter(it)
      );
  }
}

function getCurrentPageLedgerItems(viewItemsInput = null) {
  const currentViewItems =
    viewItemsInput !== null ? viewItemsInput : getLedgerItemsByView(state.ledgerView);
  const totalPages = Math.max(1, Math.ceil(currentViewItems.length / state.ledgerPageSize));
  const safePage = Math.min(Math.max(1, state.ledgerPage), totalPages);
  const startIdx = (safePage - 1) * state.ledgerPageSize;
  return currentViewItems.slice(startIdx, startIdx + state.ledgerPageSize);
}

function syncSelectAllCheckboxState(pageItemsInput = null) {
  if (!dom.chkSelectAllVisible) return;
  const pageItems =
    pageItemsInput !== null ? pageItemsInput : getCurrentPageLedgerItems();
  const selectableItems = pageItems.filter(
    (it) => !it.isSummaryRow || state.includeSummaryRows
  );

  if (selectableItems.length === 0) {
    dom.chkSelectAllVisible.checked = false;
    dom.chkSelectAllVisible.indeterminate = false;
    return;
  }

  let checkedCount = 0;
  for (const it of selectableItems) {
    if (state.checkedIds.has(it.id)) {
      checkedCount++;
    }
  }

  dom.chkSelectAllVisible.checked = checkedCount === selectableItems.length;
  dom.chkSelectAllVisible.indeterminate =
    checkedCount > 0 && checkedCount < selectableItems.length;
}

function renderLedgerSection() {
  const allCount = state.items.length;
  dom.ledgerTotalBadge.textContent = `${formatNumberVN(allCount)} dòng`;

  if (allCount === 0) {
    dom.tabLedgerFiltered.textContent = 'Các dòng trong khoảng ngày đang lọc (0)';
    dom.tabLedgerAll.textContent = 'Toàn bộ file (0)';
    dom.tabLedgerUnreconciled.textContent = 'Chưa đối soát (0)';
    dom.tabLedgerReconciled.textContent = 'Đã đối soát (0)';
    dom.tabLedgerInspect.classList.add('hidden');
    dom.btnClearPins.classList.add('hidden');
    dom.btnClearExcludes.classList.add('hidden');
    dom.ledgerPaginationBar.classList.add('hidden');
    dom.chkSelectAllVisible.checked = false;
    dom.chkSelectAllVisible.indeterminate = false;
    dom.chkSelectAllVisible.disabled = true;

    dom.ledgerTableBody.innerHTML = `
      <tr>
        <td colspan="8" class="empty-grid-cell">
          Chưa có dữ liệu sổ chi tiết. Hãy chọn file Excel hoặc nạp bảng kê mẫu ở Bước 1.
        </td>
      </tr>
    `;
    updateLedgerStatusBar([]);
    return;
  }

  dom.chkSelectAllVisible.disabled = false;

  const filteredList = getLedgerItemsByView('filtered');
  const unreconciledList = getLedgerItemsByView('unreconciled');
  const reconciledList = getLedgerItemsByView('reconciled');

  dom.tabLedgerFiltered.textContent = `Các dòng trong khoảng ngày đang lọc (${formatNumberVN(
    filteredList.length
  )})`;
  dom.tabLedgerAll.textContent = `Toàn bộ file (${formatNumberVN(allCount)})`;
  dom.tabLedgerUnreconciled.textContent = `Chưa đối soát (${formatNumberVN(
    unreconciledList.length
  )})`;
  dom.tabLedgerReconciled.textContent = `Đã đối soát (${formatNumberVN(
    reconciledList.length
  )})`;

  if (state.inspectedIds.size > 0) {
    dom.tabLedgerInspect.classList.remove('hidden');
    dom.tabLedgerInspect.textContent = `Đang đối chiếu: ${state.inspectLabel} (${state.inspectedIds.size})`;
  } else {
    dom.tabLedgerInspect.classList.add('hidden');
    if (state.ledgerView === 'inspect') {
      state.ledgerView = 'filtered';
    }
  }

  const allTabBtns = dom.ledgerFilterTabs.querySelectorAll('.tab-btn[data-ledger-view]');
  allTabBtns.forEach((btn) => {
    btn.classList.toggle(
      'active',
      btn.getAttribute('data-ledger-view') === state.ledgerView
    );
  });

  // Nút gỡ nhanh Ghim / Loại trừ
  if (state.pinnedIds.size > 0) {
    dom.btnClearPins.classList.remove('hidden');
    dom.btnClearPins.textContent = `Bỏ ghim bắt buộc (${state.pinnedIds.size})`;
  } else {
    dom.btnClearPins.classList.add('hidden');
  }

  if (state.manualExcludedIds.size > 0) {
    dom.btnClearExcludes.classList.remove('hidden');
    dom.btnClearExcludes.textContent = `Bỏ loại trừ thủ công (${state.manualExcludedIds.size})`;
  } else {
    dom.btnClearExcludes.classList.add('hidden');
  }

  const currentViewItems = getLedgerItemsByView(state.ledgerView);
  const totalPages = Math.max(1, Math.ceil(currentViewItems.length / state.ledgerPageSize));
  if (state.ledgerPage > totalPages) state.ledgerPage = totalPages;
  if (state.ledgerPage < 1) state.ledgerPage = 1;

  const startIdx = (state.ledgerPage - 1) * state.ledgerPageSize;
  const pageItems = currentViewItems.slice(startIdx, startIdx + state.ledgerPageSize);

  // Phân trang
  if (totalPages > 1) {
    dom.ledgerPaginationBar.classList.remove('hidden');
    dom.ledgerPageInfo.textContent = `Trang ${state.ledgerPage} / ${totalPages} (Hiển thị dòng ${formatNumberVN(
      startIdx + 1
    )} đến ${formatNumberVN(startIdx + pageItems.length)} trên tổng ${formatNumberVN(
      currentViewItems.length
    )} dòng)`;
    dom.btnPrevPage.disabled = state.ledgerPage <= 1;
    dom.btnNextPage.disabled = state.ledgerPage >= totalPages;
  } else {
    dom.ledgerPaginationBar.classList.add('hidden');
  }

  // Render các dòng bảng sổ chi tiết
  if (pageItems.length === 0) {
    syncSelectAllCheckboxState([]);
    dom.ledgerTableBody.innerHTML = `
      <tr>
        <td colspan="8" class="empty-grid-cell">
          Không có dòng dữ liệu nào trong mục lọc đang chọn.
        </td>
      </tr>
    `;
  } else {
    syncSelectAllCheckboxState(pageItems);

    dom.ledgerTableBody.innerHTML = pageItems
      .map((item) => buildLedgerRowHtml(item))
      .join('');
  }

  updateLedgerStatusBar(currentViewItems);
}

function buildLedgerRowHtml(item) {
  const batch = state.reconciledMap.get(item.id);
  const isReconciled = Boolean(batch);
  const isPinned = state.pinnedIds.has(item.id);
  const isManualExcluded = state.manualExcludedIds.has(item.id);
  const isChecked = state.checkedIds.has(item.id);
  const isInspected = state.inspectedIds.has(item.id);
  const isSkippedSummary = item.isSummaryRow && !state.includeSummaryRows;

  const rowClasses = [];
  if (isReconciled) rowClasses.push('row-reconciled');
  else if (isPinned) rowClasses.push('row-pinned');
  else if (isManualExcluded) rowClasses.push('row-excluded');
  else if (isSkippedSummary) rowClasses.push('row-summary');

  if (isChecked) rowClasses.push('row-checked');
  if (isInspected && state.ledgerView !== 'inspect') rowClasses.push('row-inspected');

  let statusHtml = '<span class="status-tag status-unreconciled">Chưa đối soát</span>';
  if (isReconciled) {
    statusHtml = `<span class="status-tag status-reconciled">Đã gom (${escapeHtml(
      batch.label
    )})</span>`;
  } else if (isPinned) {
    statusHtml = '<span class="status-tag status-pinned">Ghim bắt buộc</span>';
  } else if (isManualExcluded) {
    statusHtml = '<span class="status-tag status-excluded">Đã loại trừ</span>';
  } else if (isSkippedSummary) {
    statusHtml = '<span class="status-tag status-summary">Dòng tổng cộng (Bỏ qua)</span>';
  }

  const disablePin = isReconciled || isSkippedSummary;
  const disableExclude = isReconciled || isSkippedSummary;

  const dateStr = item.dateDisplay || formatDateVN(item.dateIso) || '-';
  const refStr = item.reference || '-';
  const descStr = getDisplayDescription(item) || '-';

  return `
    <tr class="${rowClasses.join(' ')}" data-row-id="${escapeHtml(item.id)}">
      <td class="col-check">
        <input
          type="checkbox"
          class="row-check-input"
          data-row-action="check"
          data-item-id="${escapeHtml(item.id)}"
          ${isChecked ? 'checked' : ''}
          aria-label="Tích chọn dòng Excel #${item.excelRowNumber}"
        >
      </td>
      <td class="col-excel-row">#${item.excelRowNumber}</td>
      <td class="col-date">${escapeHtml(dateStr)}</td>
      <td class="col-ref">${escapeHtml(refStr)}</td>
      <td class="col-desc">${escapeHtml(descStr)}</td>
      <td class="col-num col-amount">${formatAmountHtml(item.amount)}</td>
      <td class="col-status">${statusHtml}</td>
      <td class="col-actions">
        <div class="row-action-group">
          <button
            type="button"
            class="row-btn${isPinned ? ' active-pin' : ''}"
            data-row-action="pin"
            data-item-id="${escapeHtml(item.id)}"
            ${disablePin ? 'disabled' : ''}
            title="Bắt buộc mọi phương án gom tiền phải chứa dòng chứng từ này"
          >
            ${isPinned ? 'Đang ghim' : 'Ghim bắt buộc'}
          </button>
          <button
            type="button"
            class="row-btn${isManualExcluded ? ' active-exclude' : ''}"
            data-row-action="exclude"
            data-item-id="${escapeHtml(item.id)}"
            ${disableExclude ? 'disabled' : ''}
            title="Loại dòng chứng từ này khỏi phép tìm kiếm"
          >
            ${isManualExcluded ? 'Bỏ loại' : 'Loại trừ'}
          </button>
        </div>
      </td>
    </tr>
  `;
}

function updateLedgerStatusBar(visibleItemsInput = null) {
  const visibleItems =
    visibleItemsInput !== null
      ? visibleItemsInput
      : getLedgerItemsByView(state.ledgerView);

  let visibleSum = 0;
  for (const it of visibleItems) {
    if (!it.isSummaryRow || state.includeSummaryRows) {
      visibleSum += Number(it.amount || 0);
    }
  }

  dom.statusBarLeft.innerHTML = `
    <span>Số dòng đang xem: <strong>${formatNumberVN(visibleItems.length)}</strong></span>
    <span class="statusbar-sep">|</span>
    <span>Cộng vùng đang xem: <strong>${formatAmountHtml(visibleSum)}</strong></span>
  `;

  // Gộp chung các dòng đang tích chọn hoặc đang ghim bắt buộc (không tính trùng)
  const selectedUnion = new Set([...state.pinnedIds, ...state.checkedIds]);
  const selectedCount = selectedUnion.size;
  let selectedSum = 0;
  for (const id of selectedUnion) {
    const item = state.itemById.get(id);
    if (item) {
      selectedSum += Number(item.amount || 0);
    }
  }

  let diffDisplayHtml = '<strong>-- (Chưa nhập số cần gom)</strong>';
  if (state.targetAmount !== null && Number.isFinite(state.targetAmount)) {
    const diff = Math.round((selectedSum - state.targetAmount) * 100) / 100;
    if (selectedCount === 0) {
      diffDisplayHtml = `<strong class="num-val">${escapeHtml(
        formatNumberVN(-state.targetAmount, { showPlusSign: true })
      )}</strong>`;
    } else if (diff === 0) {
      diffDisplayHtml = '<strong class="num-val num-zero">0 (Khớp chính xác)</strong>';
    } else {
      diffDisplayHtml = `<strong class="num-val num-neg">${escapeHtml(
        formatNumberVN(diff, { showPlusSign: true })
      )}</strong>`;
    }
  }

  dom.ledgerStatusBar.classList.toggle('has-selection', selectedCount > 0);
  dom.btnClearChecked.classList.toggle('hidden', state.checkedIds.size === 0);

  dom.selectionStatusText.innerHTML =
    `Đang ghim/chọn: <strong>${formatNumberVN(selectedCount)} dòng</strong>` +
    ` | Tổng tiền: <strong>${formatAmountHtml(selectedSum)}</strong>` +
    ` | Chênh lệch so với số cần tìm: ${diffDisplayHtml}`;
}

/* ==========================================================================
   GẮN SỰ KIỆN TƯƠNG TÁC (EVENT LISTENERS)
   ========================================================================== */
function bindEvents() {
  // 1. Cỡ chữ Chuẩn / Lớn & Hướng dẫn
  dom.btnFontNormal.addEventListener('click', () => applyFontScale('normal'));
  dom.btnFontLarge.addEventListener('click', () => applyFontScale('large'));

  dom.btnToggleGuide.addEventListener('click', () => {
    const isHidden = dom.quickGuidePanel.classList.toggle('hidden');
    dom.btnToggleGuide.setAttribute('aria-expanded', isHidden ? 'false' : 'true');
  });
  dom.btnCloseGuide.addEventListener('click', () => {
    dom.quickGuidePanel.classList.add('hidden');
    dom.btnToggleGuide.setAttribute('aria-expanded', 'false');
  });

  // 2. Nạp file Excel (Click + Drag & Drop) & Dữ liệu mẫu
  dom.btnPickFile.addEventListener('click', () => {
    dom.excelFileInput.click();
  });

  dom.excelFileInput.addEventListener('change', (e) => {
    const file = e.target.files && e.target.files[0];
    if (file) handleFileSelected(file);
  });

  dom.dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dom.dropZone.classList.add('drag-over');
  });

  dom.dropZone.addEventListener('dragleave', (e) => {
    if (!dom.dropZone.contains(e.relatedTarget)) {
      dom.dropZone.classList.remove('drag-over');
    }
  });

  dom.dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dom.dropZone.classList.remove('drag-over');
    const file = e.dataTransfer?.files?.[0];
    if (file) handleFileSelected(file);
  });

  dom.btnLoadSample.addEventListener('click', handleLoadSampleDataset);

  // Đổi Sheet
  dom.sheetSelect.addEventListener('change', () => {
    const nextSheet = dom.sheetSelect.value;
    if (!nextSheet || nextSheet === state.activeSheetName) return;
    if (!confirmBeforeReplacingData()) {
      dom.sheetSelect.value = state.activeSheetName;
      return;
    }
    resetReconciliationSession();
    state.activeSheetName = nextSheet;
    inspectAndExtractCurrentSheet(null, true);
  });

  // Bật / tắt khung thiết lập cột
  dom.btnToggleMapping.addEventListener('click', () => {
    state.isMappingOpen = !state.isMappingOpen;
    syncMappingPanelVisibility();
  });

  // Đổi dòng tiêu đề trong Excel
  dom.headerRowSelect.addEventListener('change', () => {
    const customRow = Number(dom.headerRowSelect.value);
    if (Number.isInteger(customRow) && customRow >= 0) {
      if (customRow !== state.headerRowIndex) {
        clearBatchesAndResultsOnStructuralChange(
          'Đã làm mới danh sách đợt đối soát do thay đổi dòng tiêu đề Excel.'
        );
      }
      inspectAndExtractCurrentSheet(customRow, false);
      state.isMappingOpen = true;
      syncMappingPanelVisibility();
    }
  });

  // Đổi cấu hình các cột
  const onColumnControlChange = () => {
    const prevMapping = { ...state.mapping };
    readMappingFromControls();
    if (hasAmountMappingChanged(prevMapping, state.mapping)) {
      clearBatchesAndResultsOnStructuralChange(
        'Đã làm mới danh sách đợt đối soát do thay đổi cấu hình cột Số tiền.'
      );
    }
    syncAmountModeVisibility();
    runDataExtraction(false);
  };

  dom.dateColSelect.addEventListener('change', onColumnControlChange);
  dom.amountModeSelect.addEventListener('change', onColumnControlChange);
  dom.amountColSelect.addEventListener('change', onColumnControlChange);
  dom.debitColSelect.addEventListener('change', onColumnControlChange);
  dom.creditColSelect.addEventListener('change', onColumnControlChange);
  dom.descColSelect.addEventListener('change', onColumnControlChange);
  dom.refColSelect.addEventListener('change', onColumnControlChange);
  dom.includeSummaryCheckbox.addEventListener('change', onColumnControlChange);

  dom.btnApplyMapping.addEventListener('click', () => {
    const prevMapping = { ...state.mapping };
    readMappingFromControls();
    const completeness = evaluateMappingCompleteness(state.mapping);
    if (!completeness.hasAmount) {
      showToast('Vui lòng chọn cột Số tiền trước khi đóng thiết lập.', 'warn');
      return;
    }
    if (hasAmountMappingChanged(prevMapping, state.mapping)) {
      clearBatchesAndResultsOnStructuralChange(
        'Đã làm mới danh sách đợt đối soát do thay đổi cấu hình cột Số tiền.'
      );
    }
    state.isMappingOpen = false;
    runDataExtraction(false);
    showToast('Đã áp dụng thiết lập cột dữ liệu.');
  });

  // 3. Bước 2: Presets, Ngày, Tháng nhanh, Số tiền, Sai số, Từ khóa
  dom.presetButtonsList.addEventListener('click', (e) => {
    const btn = e.target.closest('button[data-preset-idx]');
    if (!btn) return;
    applySamplePreset(Number(btn.getAttribute('data-preset-idx')));
  });

  dom.dateFromInput.addEventListener('change', () => {
    state.dateFromIso = dom.dateFromInput.value || '';
    state.activePresetIndex = null;
    renderSamplePresetsBar();
    syncDateReadoutsAndMonthHighlight();
    state.ledgerPage = 1;
    updateCandidatePreviewAndLedger();
  });

  dom.dateToInput.addEventListener('change', () => {
    state.dateToIso = dom.dateToInput.value || '';
    state.activePresetIndex = null;
    renderSamplePresetsBar();
    syncDateReadoutsAndMonthHighlight();
    state.ledgerPage = 1;
    updateCandidatePreviewAndLedger();
  });

  dom.monthQuickBar.addEventListener('click', (e) => {
    const chip = e.target.closest('button[data-month-key]');
    if (!chip) return;
    const key = chip.getAttribute('data-month-key');
    state.activePresetIndex = null;
    renderSamplePresetsBar();

    if (key === 'all') {
      state.dateFromIso = state.minDateIso || '';
      state.dateToIso = state.maxDateIso || '';
    } else {
      const monthObj = state.availableMonths.find((m) => m.key === key);
      if (!monthObj) return;
      state.dateFromIso = monthObj.fromIso;
      state.dateToIso = monthObj.toIso;
    }

    dom.dateFromInput.value = state.dateFromIso;
    dom.dateToInput.value = state.dateToIso;
    syncDateReadoutsAndMonthHighlight();
    state.ledgerPage = 1;
    updateCandidatePreviewAndLedger();
  });

  dom.targetAmountInput.addEventListener('input', () => {
    state.activePresetIndex = null;
    renderSamplePresetsBar();
    handleLiveTargetAmountInput();
  });

  dom.targetAmountInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      executeSearch();
    }
  });

  dom.btnClearAmount.addEventListener('click', () => {
    setTargetAmountProgrammatically(null);
    dom.targetAmountInput.focus();
  });

  dom.toleranceInput.addEventListener('input', handleLiveToleranceInput);
  dom.toleranceInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      executeSearch();
    }
  });

  dom.toleranceQuickGroup.addEventListener('click', (e) => {
    const chip = e.target.closest('button[data-tolerance]');
    if (!chip) return;
    setToleranceProgrammatically(Number(chip.getAttribute('data-tolerance')));
  });

  dom.keywordInput.addEventListener('input', () => {
    state.keyword = dom.keywordInput.value;
    dom.btnClearKeyword.classList.toggle('hidden', !state.keyword);
    state.ledgerPage = 1;
    updateCandidatePreviewAndLedger();
  });

  dom.keywordInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      executeSearch();
    }
  });

  dom.btnClearKeyword.addEventListener('click', () => {
    state.keyword = '';
    dom.keywordInput.value = '';
    dom.btnClearKeyword.classList.add('hidden');
    state.ledgerPage = 1;
    updateCandidatePreviewAndLedger();
    dom.keywordInput.focus();
  });

  dom.maxItemsSelect.addEventListener('change', () => {
    const val = Number(dom.maxItemsSelect.value);
    state.maxItems = Number.isInteger(val) && val > 0 ? val : DEFAULT_MAX_ITEMS;
  });

  dom.btnResetFilters.addEventListener('click', () => {
    state.activePresetIndex = null;
    renderSamplePresetsBar();
    state.dateFromIso = state.minDateIso || '';
    state.dateToIso = state.maxDateIso || '';
    dom.dateFromInput.value = state.dateFromIso;
    dom.dateToInput.value = state.dateToIso;
    setTargetAmountProgrammatically(null);
    setToleranceProgrammatically(DEFAULT_TOLERANCE);
    state.keyword = '';
    dom.keywordInput.value = '';
    dom.btnClearKeyword.classList.add('hidden');
    state.maxItems = DEFAULT_MAX_ITEMS;
    dom.maxItemsSelect.value = String(DEFAULT_MAX_ITEMS);
    syncDateReadoutsAndMonthHighlight();
    state.ledgerPage = 1;
    updateCandidatePreviewAndLedger();
    showToast('Đã đặt lại bộ lọc về mặc định.');
  });

  dom.btnRunSearch.addEventListener('click', executeSearch);
  dom.btnCancelSearch.addEventListener('click', () => cancelRunningSearch(false));

  // 4. Bước 3A: Kết quả phương án
  dom.resultsFilterBar.addEventListener('click', (e) => {
    const btn = e.target.closest('button[data-res-filter]');
    if (!btn) return;
    state.resultsFilter = btn.getAttribute('data-res-filter') || 'all';
    renderResultsArea();
  });

  dom.solutionsContainer.addEventListener('input', (e) => {
    const noteInput = e.target.closest('input[data-sol-note-id]');
    if (!noteInput) return;
    const solId = noteInput.getAttribute('data-sol-note-id');
    state.solutionDraftNotes.set(solId, noteInput.value);
  });

  dom.solutionsContainer.addEventListener('click', (e) => {
    const actionBtn = e.target.closest('[data-sol-action]');
    if (!actionBtn) return;

    const action = actionBtn.getAttribute('data-sol-action');
    const solId = actionBtn.getAttribute('data-sol-id');
    const sol = state.searchResult?.solutions?.find((s) => s.id === solId);
    if (!sol) return;

    if (action === 'toggle') {
      if (state.expandedSolutionIds.has(sol.id)) {
        state.expandedSolutionIds.delete(sol.id);
      } else {
        state.expandedSolutionIds.add(sol.id);
      }
      renderResultsArea();
    } else if (action === 'copy') {
      handleCopySolutionToExcel(sol);
    } else if (action === 'confirm') {
      handleConfirmSolutionAsBatch(sol);
    } else if (action === 'inspect') {
      const rank = actionBtn.getAttribute('data-sol-rank') || '';
      inspectIdsInLedger(sol.itemIds, `Phương án #${rank}`);
    }
  });

  // 5. Quản lý Đợt đối soát & Xuất Excel
  dom.batchesTableBody.addEventListener('input', (e) => {
    const input = e.target.closest('input[data-batch-note-id]');
    if (!input) return;
    const batchId = input.getAttribute('data-batch-note-id');
    const batch = state.reconciledBatches.find((b) => b.id === batchId);
    if (batch) {
      batch.note = input.value;
    }
  });

  dom.batchesTableBody.addEventListener('click', (e) => {
    const btn = e.target.closest('button[data-batch-action]');
    if (!btn) return;
    const action = btn.getAttribute('data-batch-action');
    const batchId = btn.getAttribute('data-batch-id');
    const batch = state.reconciledBatches.find((b) => b.id === batchId);
    if (!batch) return;

    if (action === 'undo') {
      undoSingleBatch(batchId);
    } else if (action === 'inspect') {
      inspectIdsInLedger(batch.itemIds, batch.label);
    }
  });

  dom.btnClearAllBatches.addEventListener('click', handleClearAllBatches);
  dom.btnExportExcelTop.addEventListener('click', handleExportReconciledExcel);
  dom.btnExportExcelLedger.addEventListener('click', handleExportReconciledExcel);

  // 6. Sổ chi tiết dữ liệu gốc & Status bar
  dom.ledgerFilterTabs.addEventListener('click', (e) => {
    const tab = e.target.closest('button[data-ledger-view]');
    if (!tab) return;
    state.ledgerView = tab.getAttribute('data-ledger-view') || 'filtered';
    state.ledgerPage = 1;
    renderLedgerSection();
  });

  dom.ledgerPageSize.addEventListener('change', () => {
    state.ledgerPageSize = Number(dom.ledgerPageSize.value) || 200;
    state.ledgerPage = 1;
    renderLedgerSection();
  });

  dom.btnPrevPage.addEventListener('click', () => {
    if (state.ledgerPage > 1) {
      state.ledgerPage -= 1;
      renderLedgerSection();
      dom.ledgerTableWrap.scrollTop = 0;
    }
  });

  dom.btnNextPage.addEventListener('click', () => {
    state.ledgerPage += 1;
    renderLedgerSection();
    dom.ledgerTableWrap.scrollTop = 0;
  });

  dom.chkSelectAllVisible.addEventListener('change', () => {
    const currentViewItems = getLedgerItemsByView(state.ledgerView);
    const startIdx = (state.ledgerPage - 1) * state.ledgerPageSize;
    const pageItems = currentViewItems.slice(startIdx, startIdx + state.ledgerPageSize);

    if (dom.chkSelectAllVisible.checked) {
      for (const it of pageItems) {
        if (!it.isSummaryRow || state.includeSummaryRows) {
          state.checkedIds.add(it.id);
        }
      }
    } else {
      for (const it of pageItems) {
        state.checkedIds.delete(it.id);
      }
    }
    renderLedgerSection();
  });

  dom.ledgerTableBody.addEventListener('change', (e) => {
    const chk = e.target.closest('input[data-row-action="check"]');
    if (!chk) return;
    const itemId = chk.getAttribute('data-item-id');
    if (!itemId) return;

    if (chk.checked) {
      state.checkedIds.add(itemId);
    } else {
      state.checkedIds.delete(itemId);
    }
    const rowEl = chk.closest('tr');
    if (rowEl) {
      rowEl.classList.toggle('row-checked', chk.checked);
    }
    syncSelectAllCheckboxState();
    updateLedgerStatusBar();
  });

  dom.ledgerTableBody.addEventListener('click', (e) => {
    const btn = e.target.closest('button[data-row-action]');
    if (!btn) return;
    const action = btn.getAttribute('data-row-action');
    const itemId = btn.getAttribute('data-item-id');
    if (!itemId || state.reconciledMap.has(itemId)) return;

    if (action === 'pin') {
      if (state.pinnedIds.has(itemId)) {
        state.pinnedIds.delete(itemId);
      } else {
        state.pinnedIds.add(itemId);
        state.manualExcludedIds.delete(itemId);
      }
      updateCandidatePreviewAndLedger();
    } else if (action === 'exclude') {
      if (state.manualExcludedIds.has(itemId)) {
        state.manualExcludedIds.delete(itemId);
      } else {
        state.manualExcludedIds.add(itemId);
        state.pinnedIds.delete(itemId);
      }
      updateCandidatePreviewAndLedger();
    }
  });

  dom.btnClearPins.addEventListener('click', () => {
    state.pinnedIds.clear();
    updateCandidatePreviewAndLedger();
    showToast('Đã bỏ ghim tất cả các dòng chứng từ.');
  });

  dom.btnClearExcludes.addEventListener('click', () => {
    state.manualExcludedIds.clear();
    updateCandidatePreviewAndLedger();
    showToast('Đã bỏ loại trừ thủ công tất cả các dòng.');
  });

  dom.btnClearChecked.addEventListener('click', () => {
    state.checkedIds.clear();
    syncSelectAllCheckboxState();
    renderLedgerSection();
  });
}

/* ==========================================================================
   KHỞI ĐỘNG ỨNG DỤNG
   ========================================================================== */
function initApp() {
  cacheDomElements();
  initFontScale();
  bindEvents();
  setToleranceProgrammatically(DEFAULT_TOLERANCE);
  updateAmountWordsDisplay();
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', initApp);
} else {
  initApp();
}
