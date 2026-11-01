import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { performance } from 'node:perf_hooks';

import { getSampleDataset } from '../src/core/sampleData.js';
import {
  parseVietnameseAmount,
  formatNumberVN,
  formatInputAmountVN,
  parseVietnameseDate,
  formatDateVN,
  isoToUtcTimestamp,
  readAmountInVietnameseWords,
} from '../src/core/vnFormat.js';
import {
  loadWorkbookFromBuffer,
  getSheetNames,
  inspectSheet,
  inspectRawGrid,
  extractTransactions,
  exportReconciledWorkbook,
} from '../src/core/excelParser.js';
import {
  normalizeVietnameseSearch,
  findCombinations,
} from '../src/core/reconcileEngine.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const SAMPLE_EXCEL_PATH = path.resolve(__dirname, '..', 'public', 'mau-cong-no-ke-toan.xlsx');

function readFileAsArrayBuffer(filePath) {
  const buf = fs.readFileSync(filePath);
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
}

// ============================================================================
// 1. SAMPLE DATA INTEGRITY TESTS (`sampleData.js`)
// ============================================================================
test('sampleData.js: getSampleDataset() structure, headers, summary footer, and immutability', () => {
  const dataset = getSampleDataset();
  assert.equal(dataset.fileName, 'So_Chi_Tiet_Cong_No_Q1_2026.xlsx');
  assert.equal(dataset.sheetName, 'CongNo_PhaiThu_2026');
  assert.ok(Array.isArray(dataset.rawGrid));
  assert.equal(dataset.rawGrid.length, 36); // 4 report headers + 1 table header + 30 tx rows + 1 summary footer

  // Top 4 company report header rows
  assert.equal(dataset.rawGrid[0][0], 'CÔNG TY TNHH THƯƠNG MẠI VÀ DỊCH VỤ PHƯỚC TÂN');
  assert.equal(dataset.rawGrid[1][0], 'Địa chỉ: Quận Hải Châu, TP. Đà Nẵng');
  assert.equal(
    dataset.rawGrid[2][0],
    'SỔ CHI TIẾT ĐỐI SOÁT HÓA ĐƠN & CÔNG NỢ PHẢI THU - QUÝ 1/2026'
  );
  assert.deepEqual(dataset.rawGrid[3], ['', '', '', '', '', '']);

  // Row 5 (index 4) table column headers
  assert.deepEqual(dataset.rawGrid[4], [
    'Ngày chứng từ',
    'Số chứng từ',
    'Khách hàng / Đối tác',
    'Diễn giải nội dung',
    'Số tiền (VNĐ)',
    'Ghi chú',
  ]);

  // Business rows (indices 5..34) & summary footer (index 35)
  const txRows = dataset.rawGrid.slice(5, -1);
  assert.equal(txRows.length, 30);

  const negativeRows = txRows.filter((row) => row[4] < 0);
  assert.ok(negativeRows.length >= 5, 'Should contain realistic negative accounting rows');

  const computedSum = txRows.reduce((acc, row) => acc + row[4], 0);
  const footerRow = dataset.rawGrid[dataset.rawGrid.length - 1];
  assert.equal(footerRow[3], 'TỔNG CỘNG PHÁT SINH');
  assert.equal(footerRow[4], computedSum);
  assert.equal(computedSum, 235100050);

  // Presets check
  assert.equal(dataset.presets.length, 3);
  assert.equal(dataset.presets[0].targetAmount, 18450000);
  assert.equal(dataset.presets[1].targetAmount, 24800000);
  assert.equal(dataset.presets[2].targetAmount, 32150000);
  assert.equal(dataset.presets[2].keyword, 'Minh Phú');

  // Immutability check (fresh copy on each call)
  dataset.rawGrid[0][0] = 'MUTATED';
  dataset.presets[0].targetAmount = 999;
  const freshDataset = getSampleDataset();
  assert.equal(freshDataset.rawGrid[0][0], 'CÔNG TY TNHH THƯƠNG MẠI VÀ DỊCH VỤ PHƯỚC TÂN');
  assert.equal(freshDataset.presets[0].targetAmount, 18450000);
});

// ============================================================================
// 2. VIETNAMESE FORMATTING & PARSING TESTS (`vnFormat.js`)
// ============================================================================
test('vnFormat.js: parseVietnameseAmount handles dots, commas, negatives, parentheses, and decimals', () => {
  assert.equal(parseVietnameseAmount(12500000), 12500000);
  assert.equal(parseVietnameseAmount(0), 0);
  assert.equal(parseVietnameseAmount('12.500.000'), 12500000);
  assert.equal(parseVietnameseAmount('12,500,000'), 12500000);
  assert.equal(parseVietnameseAmount('(500.000)'), -500000);
  assert.equal(parseVietnameseAmount('(1,250,000)'), -1250000);
  assert.equal(parseVietnameseAmount('-1.250.000'), -1250000);
  assert.equal(parseVietnameseAmount('-500.000 đ'), -500000);
  assert.equal(parseVietnameseAmount('24.800.000 VNĐ'), 24800000);
  assert.equal(parseVietnameseAmount('1.250.000,50'), 1250000.5);

  assert.equal(parseVietnameseAmount(''), null);
  assert.equal(parseVietnameseAmount('   '), null);
  assert.equal(parseVietnameseAmount(null), null);
  assert.equal(parseVietnameseAmount(undefined), null);
  assert.equal(parseVietnameseAmount('Không phải số'), null);
});

test('vnFormat.js: formatNumberVN and formatInputAmountVN format Vietnamese thousand separators', () => {
  assert.equal(formatNumberVN(12500000), '12.500.000');
  assert.equal(formatNumberVN(-500000), '-500.000');
  assert.equal(formatNumberVN(0), '0');
  assert.equal(formatNumberVN(250, { showPlusSign: true }), '+250');
  assert.equal(formatNumberVN(-350, { showPlusSign: true }), '-350');

  const liveInput = formatInputAmountVN('18450000');
  assert.equal(liveInput.numeric, 18450000);
  assert.equal(liveInput.formatted, '18.450.000');
});

test('vnFormat.js: parseVietnameseDate strictly defaults XX/YY/ZZZZ to DD/MM/YYYY (Vietnamese standard)', () => {
  // CRITICAL: "05/03/2026" MUST be 5 March 2026, NEVER May 3rd 2026!
  const march5 = parseVietnameseDate('05/03/2026');
  assert.ok(march5, 'Should parse 05/03/2026');
  assert.equal(march5.iso, '2026-03-05');
  assert.equal(march5.display, '05/03/2026');
  assert.equal(march5.timestamp, Date.UTC(2026, 2, 5));

  // Dash, ISO, 2-digit year dot, and Excel serial number
  const dashDate = parseVietnameseDate('02-03-2026');
  assert.equal(dashDate?.iso, '2026-03-02');
  assert.equal(dashDate?.display, '02/03/2026');

  const isoDate = parseVietnameseDate('2026-03-15');
  assert.equal(isoDate?.iso, '2026-03-15');
  assert.equal(isoDate?.display, '15/03/2026');

  const dotShortYear = parseVietnameseDate('15.03.26');
  assert.equal(dotShortYear?.iso, '2026-03-15');

  // Excel serial date 46086 -> 2026-03-05
  const excelSerial = parseVietnameseDate(46086);
  assert.equal(excelSerial?.iso, '2026-03-05');

  // Helpers & invalid dates
  assert.equal(formatDateVN('2026-03-05'), '05/03/2026');
  assert.equal(isoToUtcTimestamp('2026-03-05'), Date.UTC(2026, 2, 5));
  assert.equal(parseVietnameseDate('32/01/2026'), null);
  assert.equal(parseVietnameseDate(''), null);
  assert.equal(parseVietnameseDate('TỔNG CỘNG PHÁT SINH'), null);
});

test('vnFormat.js: readAmountInVietnameseWords reads positive and negative VND amounts accurately', () => {
  const wordsPositive = readAmountInVietnameseWords(12500000).toLowerCase();
  assert.ok(wordsPositive.includes('mười hai triệu'), `Unexpected words: ${wordsPositive}`);
  assert.ok(wordsPositive.includes('năm trăm'), `Unexpected words: ${wordsPositive}`);
  assert.ok(
    wordsPositive.includes('nghìn') || wordsPositive.includes('ngàn'),
    `Should contain nghìn/ngàn: ${wordsPositive}`
  );
  assert.ok(wordsPositive.includes('đồng'), `Unexpected words: ${wordsPositive}`);

  const wordsNegative = readAmountInVietnameseWords(-500000).toLowerCase();
  assert.ok(wordsNegative.startsWith('âm'), `Negative amount should start with "âm": ${wordsNegative}`);
  assert.ok(wordsNegative.includes('năm trăm'), `Unexpected words: ${wordsNegative}`);
  assert.ok(wordsNegative.includes('đồng'), `Unexpected words: ${wordsNegative}`);
});

// ============================================================================
// 3. REAL EXCEL PARSER & ROUNDTRIP EXPORT TESTS (`excelParser.js`)
// ============================================================================
test('excelParser.js: parses real public/mau-cong-no-ke-toan.xlsx, detects header row #5 & summary row, and exports .xlsx', () => {
  assert.ok(fs.existsSync(SAMPLE_EXCEL_PATH), 'public/mau-cong-no-ke-toan.xlsx must exist on disk');
  const arrayBuffer = readFileAsArrayBuffer(SAMPLE_EXCEL_PATH);

  const workbook = loadWorkbookFromBuffer(arrayBuffer);
  const sheetNames = getSheetNames(workbook);
  assert.deepEqual(sheetNames, ['CongNo_PhaiThu_2026']);

  const inspection = inspectSheet(workbook, 'CongNo_PhaiThu_2026');
  assert.equal(
    inspection.detectedHeaderRowIndex,
    4,
    'Must skip 4 company header rows and detect table header at index 4 (Excel row #5)'
  );
  assert.equal(inspection.columns.length, 6);
  assert.equal(inspection.suggestedMapping.dateColIndex, 0);
  assert.equal(inspection.suggestedMapping.amountColIndex, 4);
  assert.equal(inspection.suggestedMapping.amountMode, 'single');
  assert.equal(inspection.suggestedMapping.refColIndex, 1);

  const extracted = extractTransactions(
    inspection.rawGrid,
    inspection.detectedHeaderRowIndex,
    inspection.suggestedMapping,
    inspection.columns
  );

  // Verify summary row detection ("TỔNG CỘNG PHÁT SINH")
  assert.equal(extracted.summaryRowCount, 1, 'Should detect 1 summary row (TỔNG CỘNG PHÁT SINH)');
  const summaryItems = extracted.items.filter((item) => item.isSummaryRow);
  assert.equal(summaryItems.length, 1);
  assert.equal(summaryItems[0].amount, 235100050);

  // Verify 30 valid business items across Jan, Feb, Mar 2026
  const businessItems = extracted.items.filter((item) => !item.isSummaryRow);
  assert.equal(businessItems.length, 30);
  assert.equal(extracted.minDateIso, '2026-01-05');
  assert.equal(extracted.maxDateIso, '2026-03-31');

  const monthKeys = extracted.availableMonths.map((m) => m.key);
  assert.deepEqual(monthKeys, ['2026-01', '2026-02', '2026-03']);
  const monthCounts = Object.fromEntries(extracted.availableMonths.map((m) => [m.key, m.count]));
  assert.deepEqual(monthCounts, {
    '2026-01': 5,
    '2026-02': 9,
    '2026-03': 16,
  });

  // Verify roundtrip Excel reconciliation export
  const sampleBatchItems = businessItems.slice(0, 2);
  const exportedBytes = exportReconciledWorkbook({
    sheetName: inspection.sheetName,
    columns: inspection.columns,
    items: extracted.items,
    reconciledBatches: [
      {
        batchId: 'DOT-01',
        label: 'Đợt 1 - Thu tiền Hòa Bình',
        reconciledAt: '2026-03-31 10:00',
        targetAmount: 15600000,
        totalAmount: 15600150,
        diff: 150,
        note: 'Khách chuyển khoản chẵn',
        itemIds: sampleBatchItems.map((i) => i.id),
        items: sampleBatchItems,
      },
    ],
    unallocatedOnly: false,
  });

  assert.ok(exportedBytes instanceof Uint8Array, 'Exported workbook should be a Uint8Array');
  assert.ok(exportedBytes.byteLength > 1000, 'Exported workbook bytes should be non-trivial');

  const exportedArrayBuffer = exportedBytes.buffer.slice(
    exportedBytes.byteOffset,
    exportedBytes.byteOffset + exportedBytes.byteLength
  );
  const roundtripWb = loadWorkbookFromBuffer(exportedArrayBuffer);
  const roundtripSheets = getSheetNames(roundtripWb);
  assert.ok(roundtripSheets.includes('Chi_Tiet_Doi_Soat'));
  assert.ok(roundtripSheets.includes('Tong_Hop_Cac_Dot_Gom'));
});

test('excelParser.js: supports debit_minus_credit (Phát sinh Nợ - Phát sinh Có) ledger columns', () => {
  const debitCreditGrid = [
    ['Ngày hạch toán', 'Số CT', 'Diễn giải', 'Phát sinh Nợ', 'Phát sinh Có'],
    ['05/03/2026', 'HD01', 'Bán hàng lô A', '10.000.000', '0'],
    ['08/03/2026', 'CK01', 'Chiết khấu thương mại lô A', '0', '500.000'],
    ['', '', 'TỔNG CỘNG', '10.000.000', '500.000'],
  ];

  const inspection = inspectRawGrid(debitCreditGrid, 'SoCai131');
  assert.equal(inspection.detectedHeaderRowIndex, 0);

  const mapping = {
    ...inspection.suggestedMapping,
    dateColIndex: 0,
    refColIndex: 1,
    descColIndex: 2,
    debitColIndex: 3,
    creditColIndex: 4,
    amountMode: 'debit_minus_credit',
  };

  const extracted = extractTransactions(debitCreditGrid, 0, mapping, inspection.columns);
  const validItems = extracted.items.filter((i) => !i.isSummaryRow);
  assert.equal(validItems.length, 2);
  assert.equal(validItems[0].amount, 10000000);
  assert.equal(validItems[1].amount, -500000);
  assert.equal(extracted.summaryRowCount, 1);
});

// ============================================================================
// 4. RECONCILIATION ENGINE & SOLVER TESTS (`reconcileEngine.js`)
// ============================================================================
function loadSampleExtractedItems() {
  const { sheetName, rawGrid, presets } = getSampleDataset();
  const inspection = inspectRawGrid(rawGrid, sheetName);
  const extracted = extractTransactions(
    rawGrid,
    inspection.detectedHeaderRowIndex,
    inspection.suggestedMapping,
    inspection.columns
  );
  return { items: extracted.items, presets };
}

test('reconcileEngine.js: normalizeVietnameseSearch removes accents and handles đ/Đ properly', () => {
  const normalized = normalizeVietnameseSearch('  Công ty TNHH Xây dựng Đạt Minh Phú  ');
  assert.equal(normalized, 'cong ty tnhh xay dung dat minh phu');
});

test('reconcileEngine.js: Preset 1 (01/03..15/03, 18.450.000 ±500) finds exact 0đ with negative discount AND near match -350đ', () => {
  const { items, presets } = loadSampleExtractedItems();
  const preset1 = presets[0];

  const result = findCombinations({
    items,
    dateFromIso: preset1.dateFromIso,
    dateToIso: preset1.dateToIso,
    targetAmount: preset1.targetAmount,
    tolerance: preset1.tolerance,
    keyword: preset1.keyword,
  });

  assert.equal(result.stats.timedOut, false);
  assert.equal(result.stats.exactCount, 1);
  assert.equal(result.stats.nearCount, 1);
  assert.equal(result.solutions.length, 2);

  // First solution MUST be the exact match (0đ diff): 4 positive invoices + 1 negative discount (-550.000)
  const exactSol = result.solutions[0];
  assert.equal(exactSol.isExact, true);
  assert.equal(exactSol.diff, 0);
  assert.equal(exactSol.absDiff, 0);
  assert.equal(exactSol.totalAmount, 18450000);
  assert.equal(exactSol.itemCount, 5);
  const exactRefs = exactSol.items.map((i) => i.reference).sort();
  assert.deepEqual(exactRefs, ['CK0301', 'HD0301', 'HD0303', 'HD0304', 'PT0301']);
  assert.ok(exactSol.items.some((i) => i.amount === -550000));

  // Second solution is the near match (-350đ diff within ±500đ tolerance)
  const nearSol = result.solutions[1];
  assert.equal(nearSol.isExact, false);
  assert.equal(nearSol.totalAmount, 18449650);
  assert.equal(nearSol.diff, -350);
  assert.equal(nearSol.absDiff, 350);
  const nearRefs = nearSol.items.map((i) => i.reference).sort();
  assert.deepEqual(nearRefs, ['HD0302', 'HD0305']);
});

test('reconcileEngine.js: Preset 2 (16/03..31/03, 24.800.000 ±500) finds near match +250đ when customer rounds transfer', () => {
  const { items, presets } = loadSampleExtractedItems();
  const preset2 = presets[1];

  const resultWithTol = findCombinations({
    items,
    dateFromIso: preset2.dateFromIso,
    dateToIso: preset2.dateToIso,
    targetAmount: preset2.targetAmount,
    tolerance: preset2.tolerance,
    keyword: preset2.keyword,
  });

  assert.equal(resultWithTol.solutions.length, 1);
  assert.equal(resultWithTol.stats.exactCount, 0);
  assert.equal(resultWithTol.stats.nearCount, 1);

  const sol = resultWithTol.solutions[0];
  assert.equal(sol.isExact, false);
  assert.equal(sol.totalAmount, 24800250);
  assert.equal(sol.diff, 250);
  assert.equal(sol.absDiff, 250);
  assert.deepEqual(
    sol.items.map((i) => i.reference).sort(),
    ['HD0307', 'HD0309', 'HD0311']
  );

  // With tolerance = 0, no exact match exists
  const resultZeroTol = findCombinations({
    items,
    dateFromIso: preset2.dateFromIso,
    dateToIso: preset2.dateToIso,
    targetAmount: preset2.targetAmount,
    tolerance: 0,
    keyword: preset2.keyword,
  });
  assert.equal(resultZeroTol.solutions.length, 0);
});

test('reconcileEngine.js: Preset 3 (Feb 2026, 32.150.000) filters keyword with & without Vietnamese accents ("minh phu")', () => {
  const { items, presets } = loadSampleExtractedItems();
  const preset3 = presets[2];

  // Without keyword in Feb 2026, there are 2 exact combinations (Minh Phú + cross-customer distractor)
  const withoutKeyword = findCombinations({
    items,
    dateFromIso: preset3.dateFromIso,
    dateToIso: preset3.dateToIso,
    targetAmount: preset3.targetAmount,
    tolerance: preset3.tolerance,
    keyword: '',
  });
  assert.equal(withoutKeyword.solutions.length, 2);

  // With accented keyword "Minh Phú" AND unaccented keyword "minh phu", ONLY Minh Phú's 4-item net debt matches
  for (const kw of ['Minh Phú', 'minh phu', 'MINH PHU']) {
    const filtered = findCombinations({
      items,
      dateFromIso: preset3.dateFromIso,
      dateToIso: preset3.dateToIso,
      targetAmount: preset3.targetAmount,
      tolerance: preset3.tolerance,
      keyword: kw,
    });

    assert.equal(filtered.solutions.length, 1, `Expected 1 solution for keyword="${kw}"`);
    const sol = filtered.solutions[0];
    assert.equal(sol.isExact, true);
    assert.equal(sol.totalAmount, 32150000);
    assert.deepEqual(
      sol.items.map((i) => i.reference).sort(),
      ['CT0201', 'HD0201', 'HD0203', 'HD0205']
    );
  }
});

test('reconcileEngine.js: respects pinnedIds, excludedIds, and skips summary rows by default', () => {
  const { items } = loadSampleExtractedItems();

  const hd0302 = items.find((i) => i.reference === 'HD0302');
  const ck0301 = items.find((i) => i.reference === 'CK0301');
  assert.ok(hd0302 && ck0301);

  // 1) Pinning HD0302 forces every solution to contain HD0302
  const pinnedResult = findCombinations({
    items,
    dateFromIso: '2026-03-01',
    dateToIso: '2026-03-15',
    targetAmount: 18450000,
    tolerance: 500,
    pinnedIds: [hd0302.id],
  });
  assert.equal(pinnedResult.stats.pinnedCount, 1);
  assert.equal(pinnedResult.solutions.length, 1);
  assert.ok(pinnedResult.solutions[0].itemIds.includes(hd0302.id));
  assert.equal(pinnedResult.solutions[0].totalAmount, 18449650);

  // 2) Excluding CK0301 (-550.000) eliminates the 5-item exact combination, leaving only the near match
  const excludedResult = findCombinations({
    items,
    dateFromIso: '2026-03-01',
    dateToIso: '2026-03-15',
    targetAmount: 18450000,
    tolerance: 500,
    excludedIds: [ck0301.id],
  });
  assert.equal(excludedResult.solutions.length, 1);
  assert.ok(!excludedResult.solutions[0].itemIds.includes(ck0301.id));
  assert.equal(excludedResult.solutions[0].diff, -350);

  // 3) Summary row (235.100.050) must NOT be matched as a 1-row solution when includeSummaryRows is default false
  const summaryTest = findCombinations({
    items,
    targetAmount: 235100050,
    tolerance: 0,
    maxItems: 2,
  });
  assert.ok(
    summaryTest.solutions.every((sol) => sol.items.every((it) => !it.isSummaryRow)),
    'Summary row must never appear in solutions when includeSummaryRows is false'
  );
});

test('reconcileEngine.js: handles tricky mixed multiple positive + multiple negative offsets accurately', () => {
  // Target = 10.000.000, reachable ONLY by combining 3 positives (13.700.000) + 2 negatives (-3.700.000)
  // Tests that branch-and-bound does NOT prematurely prune when intermediate sum > targetAmount
  const mixedItems = [
    { id: 'p1', excelRowNumber: 2, dateIso: '2026-03-01', dateDisplay: '01/03/2026', timestamp: Date.UTC(2026, 2, 1), amount: 6400000, description: 'HĐ 01', reference: 'P1', isSummaryRow: false },
    { id: 'p2', excelRowNumber: 3, dateIso: '2026-03-02', dateDisplay: '02/03/2026', timestamp: Date.UTC(2026, 2, 2), amount: 4200000, description: 'HĐ 02', reference: 'P2', isSummaryRow: false },
    { id: 'p3', excelRowNumber: 4, dateIso: '2026-03-03', dateDisplay: '03/03/2026', timestamp: Date.UTC(2026, 2, 3), amount: 3100000, description: 'HĐ 03', reference: 'P3', isSummaryRow: false },
    { id: 'p4', excelRowNumber: 5, dateIso: '2026-03-04', dateDisplay: '04/03/2026', timestamp: Date.UTC(2026, 2, 4), amount: 8900000, description: 'HĐ 04', reference: 'P4', isSummaryRow: false },
    { id: 'n1', excelRowNumber: 6, dateIso: '2026-03-05', dateDisplay: '05/03/2026', timestamp: Date.UTC(2026, 2, 5), amount: -2450000, description: 'Trả hàng 01', reference: 'N1', isSummaryRow: false },
    { id: 'n2', excelRowNumber: 7, dateIso: '2026-03-06', dateDisplay: '06/03/2026', timestamp: Date.UTC(2026, 2, 6), amount: -1250000, description: 'Chiết khấu 02', reference: 'N2', isSummaryRow: false },
    { id: 'n3', excelRowNumber: 8, dateIso: '2026-03-07', dateDisplay: '07/03/2026', timestamp: Date.UTC(2026, 2, 7), amount: -600000, description: 'Giảm giá 03', reference: 'N3', isSummaryRow: false },
  ];

  const result = findCombinations({
    items: mixedItems,
    targetAmount: 10000000,
    tolerance: 0,
    maxItems: 8,
  });

  assert.ok(result.solutions.length >= 1, 'Must find combination requiring overshoot + negative offsets');
  const foundTargetCombo = result.solutions.find((sol) => {
    const ids = [...sol.itemIds].sort().join(',');
    return ids === 'n1,n2,p1,p2,p3';
  });
  assert.ok(foundTargetCombo, 'Should discover exact combination {p1, p2, p3, n1, n2} = 10.000.000');
  assert.equal(foundTargetCombo.totalAmount, 10000000);
  assert.equal(foundTargetCombo.isExact, true);
});

test('reconcileEngine.js: anti-freeze stress test (< 2000ms) on 35 rows with 20 identical 500.000đ rows', () => {
  const stressItems = [];

  // 20 identical rows of 500.000đ (naive solver without symmetry-breaking explores C(20,10) = 184,756 clones!)
  for (let i = 1; i <= 20; i++) {
    const day = String(((i - 1) % 28) + 1).padStart(2, '0');
    stressItems.push({
      id: `dup-${i}`,
      excelRowNumber: i + 1,
      dateIso: `2026-03-${day}`,
      dateDisplay: `${day}/03/2026`,
      timestamp: Date.UTC(2026, 2, (i - 1) % 28 + 1),
      amount: 500000,
      description: `Phí giao hàng định kỳ chuyến #${i}`,
      reference: `PT-${String(i).padStart(2, '0')}`,
      isSummaryRow: false,
    });
  }

  // 15 varied rows (including negative discounts/returns) -> total 35 rows
  const variedAmounts = [
    1200000, 1750000, 2300000, 3100000, 4400000,
    5600000, 6800000, 7250000, 8900000, 9500000,
    11200000, 13400000, -300000, -750000, -1500000,
  ];
  variedAmounts.forEach((amt, idx) => {
    const rowIdx = 21 + idx;
    const day = String((idx % 28) + 1).padStart(2, '0');
    stressItems.push({
      id: `var-${idx + 1}`,
      excelRowNumber: rowIdx + 1,
      dateIso: `2026-03-${day}`,
      dateDisplay: `${day}/03/2026`,
      timestamp: Date.UTC(2026, 2, (idx % 28) + 1),
      amount: amt,
      description: `Nghiệp vụ phát sinh #${idx + 1}`,
      reference: `HD-${String(idx + 1).padStart(2, '0')}`,
      isSummaryRow: false,
    });
  });

  assert.equal(stressItems.length, 35);

  // Case A: Reachable target involving many duplicate 500.000đ rows (target = 5.000.000đ, maxItems = 12)
  const t0 = performance.now();
  const reachableResult = findCombinations({
    items: stressItems,
    targetAmount: 5000000,
    tolerance: 0,
    maxItems: 12,
    maxResults: 50,
    timeLimitMs: 1800,
  });
  const elapsedReachable = performance.now() - t0;

  assert.ok(
    elapsedReachable < 2000,
    `Reachable duplicate-heavy search took too long: ${elapsedReachable.toFixed(1)}ms (must be < 2000ms)`
  );
  assert.ok(reachableResult.solutions.length >= 1, 'Should find valid solutions for 5.000.000đ');

  // Case B: Unreachable target (all items are multiples of 50.000đ, target ends in 12.345đ with tol ±100đ)
  const t1 = performance.now();
  const unreachableResult = findCombinations({
    items: stressItems,
    targetAmount: 6512345,
    tolerance: 100,
    maxItems: 12,
    maxResults: 50,
    timeLimitMs: 1800,
  });
  const elapsedUnreachable = performance.now() - t1;

  assert.ok(
    elapsedUnreachable < 2000,
    `Unreachable duplicate-heavy search took too long: ${elapsedUnreachable.toFixed(1)}ms (must be < 2000ms)`
  );
  assert.equal(unreachableResult.solutions.length, 0);
});
