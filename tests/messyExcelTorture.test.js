import test from 'node:test';
import assert from 'node:assert/strict';
import * as XLSX from '../vendor/xlsx.mjs';

import {
  parseVietnameseAmount,
  formatNumberVN,
  formatInputAmountVN,
  parseVietnameseDate,
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

// ============================================================================
// PHẦN 1: TRA TẤN `parseVietnameseAmount` & `formatInputAmountVN`
// ============================================================================

test('1.1 parseVietnameseAmount: Ngoặc đơn kế toán đủ biến thể (VNĐ, đ, khoảng trắng, dấu âm kép)', () => {
  assert.equal(parseVietnameseAmount('(500.000)'), -500000);
  assert.equal(parseVietnameseAmount('( 1.250.000 đ )'), -1250000);
  assert.equal(parseVietnameseAmount('VNĐ (25,000,000)'), -25000000);
  assert.equal(parseVietnameseAmount('(25,000,000) VNĐ'), -25000000);
  assert.equal(parseVietnameseAmount('(-500.000)'), -500000);
  assert.equal(parseVietnameseAmount('-(500.000)'), -500000);
  assert.equal(parseVietnameseAmount('(1.250.000,50)'), -1250000.5);
});

test('1.2 parseVietnameseAmount: Dấu trừ ở cuối kiểu SAP/ERP & Unicode minus / en-dash / NBSP', () => {
  // SAP / ERP trailing minus
  assert.equal(parseVietnameseAmount('1.250.000-'), -1250000);
  assert.equal(parseVietnameseAmount('500,000-'), -500000);
  assert.equal(parseVietnameseAmount('1.250.000,50-'), -1250000.5);
  assert.equal(parseVietnameseAmount('1250000-'), -1250000);
  assert.equal(parseVietnameseAmount('1.250.000 -'), -1250000);

  // Unicode minus U+2212, en-dash U+2013, NBSP U+00A0, narrow NBSP U+202F
  assert.equal(parseVietnameseAmount('\u22121.500.000'), -1500000);
  assert.equal(parseVietnameseAmount('\u20131.500.000'), -1500000);
  assert.equal(parseVietnameseAmount('1.500.000\u2212'), -1500000);
  assert.equal(parseVietnameseAmount('12\u00A0500\u00A0000'), 12500000);
  assert.equal(parseVietnameseAmount('-12\u00A0500\u00A0000,25'), -12500000.25);
  assert.equal(parseVietnameseAmount('12\u202F500\u202F000'), 12500000);
});

test('1.3 parseVietnameseAmount: Bẫy phân biệt dấu nghìn vs dấu thập phân (chuẩn VN & Quốc tế)', () => {
  assert.equal(parseVietnameseAmount('1.000'), 1000);
  assert.equal(parseVietnameseAmount('500.000'), 500000);
  assert.equal(parseVietnameseAmount('1.000.000'), 1000000);
  assert.equal(parseVietnameseAmount('12,5'), 12.5);
  assert.equal(parseVietnameseAmount('12.500,75'), 12500.75);
  assert.equal(parseVietnameseAmount('12,500.75'), 12500.75);
  assert.equal(parseVietnameseAmount('0,5'), 0.5);
  assert.equal(parseVietnameseAmount('-0,25'), -0.25);

  // Hỏi khó: 1.250 vs 0.500 vs 0,500
  assert.equal(parseVietnameseAmount('1.250'), 1250);
  assert.equal(parseVietnameseAmount('1,250'), 1250);
  assert.equal(parseVietnameseAmount('0.500'), 0.5);
  assert.equal(parseVietnameseAmount('0,500'), 0.5);
  assert.equal(parseVietnameseAmount('-0.500'), -0.5);
  assert.equal(parseVietnameseAmount('-0,500'), -0.5);

  // Đảm bảo không bao giờ trả về -0
  assert.ok(Object.is(parseVietnameseAmount('-0'), 0));
  assert.ok(Object.is(parseVietnameseAmount('(0)'), 0));
  assert.ok(Object.is(parseVietnameseAmount('-0,00'), 0));
});

test('1.4 parseVietnameseAmount: Ký hiệu kế toán biểu diễn rỗng/gạch ngang & giá trị phi số', () => {
  assert.equal(parseVietnameseAmount('-'), null);
  assert.equal(parseVietnameseAmount(' - '), null);
  assert.equal(parseVietnameseAmount('\u2014'), null); // em-dash "—"
  assert.equal(parseVietnameseAmount('\u2013'), null); // en-dash "–"
  assert.equal(parseVietnameseAmount('- đ'), null);
  assert.equal(parseVietnameseAmount(''), null);
  assert.equal(parseVietnameseAmount('   \u00A0  '), null);
  assert.equal(parseVietnameseAmount(null), null);
  assert.equal(parseVietnameseAmount(undefined), null);
  assert.equal(parseVietnameseAmount(NaN), null);
  assert.equal(parseVietnameseAmount(Infinity), null);
  assert.equal(parseVietnameseAmount(-Infinity), null);
  assert.equal(parseVietnameseAmount(true), null);
  assert.equal(parseVietnameseAmount(false), null);
});

test('1.5 parseVietnameseAmount: Chống nhận nhầm chuỗi ngày tháng, mã hóa đơn, mã số thuế', () => {
  assert.equal(parseVietnameseAmount('15/03/2026'), null);
  assert.equal(parseVietnameseAmount('15.03.2026'), null);
  assert.equal(parseVietnameseAmount('15.03.26'), null);
  assert.equal(parseVietnameseAmount('2026-03-15'), null);
  assert.equal(parseVietnameseAmount('HD-00125'), null);
  assert.equal(parseVietnameseAmount('0101234567-001'), null);
  assert.equal(parseVietnameseAmount('1.23.456'), null);
  assert.equal(parseVietnameseAmount('1,23,456'), null);
});

test('1.6 [ADVERSARIAL] parseVietnameseAmount: Mã chứng từ / Mã KH bắt đầu bằng chữ Đ hoặc DONG không được nhận nhầm thành tiền', () => {
  // Trong kế toán VN, "Đ-00125" (Đơn hàng / Điều chỉnh) hoặc "DONG-001" (Mã KH Đông 001) là mã chứng từ/đối tượng, KHÔNG phải số tiền -125 hay -1!
  assert.equal(parseVietnameseAmount('Đ-00125'), null);
  assert.equal(parseVietnameseAmount('DONG-001'), null);
  assert.equal(parseVietnameseAmount('VND-00125'), null);
});

test('1.7 formatInputAmountVN: Paste dữ liệu chuẩn từ Excel/ngân hàng vào ô nhập số tiền', () => {
  assert.deepEqual(formatInputAmountVN('(500.000)'), {
    numeric: -500000,
    formatted: '-500.000',
  });
  assert.deepEqual(formatInputAmountVN('( 1.250.000 đ )'), {
    numeric: -1250000,
    formatted: '-1.250.000',
  });
  assert.deepEqual(formatInputAmountVN('(-500.000)'), {
    numeric: -500000,
    formatted: '-500.000',
  });
  assert.deepEqual(formatInputAmountVN('\u22121.500.000'), {
    numeric: -1500000,
    formatted: '-1.500.000',
  });
  assert.deepEqual(formatInputAmountVN('\u20131.500.000'), {
    numeric: -1500000,
    formatted: '-1.500.000',
  });
  assert.deepEqual(formatInputAmountVN('12\u00A0500\u00A0000'), {
    numeric: 12500000,
    formatted: '12.500.000',
  });
  assert.deepEqual(formatInputAmountVN('12.500,75'), {
    numeric: 12500.75,
    formatted: '12.500,75',
  });
  assert.deepEqual(formatInputAmountVN('12,500.75'), {
    numeric: 12500.75,
    formatted: '12.500,75',
  });
  assert.deepEqual(formatInputAmountVN('-0,25'), {
    numeric: -0.25,
    formatted: '-0,25',
  });
});

test('1.8 [ADVERSARIAL] formatInputAmountVN: Paste số âm kèm tiền tệ ("VNĐ (25,000,000)") hoặc dấu trừ cuối SAP ("1.250.000-") phải giữ nguyên dấu âm', () => {
  // Kế toán copy trực tiếp ô tiền âm từ sao kê / SAP paste vào ô Mục tiêu cần gom: tuyệt đối không được đảo dấu âm thành dương!
  assert.deepEqual(formatInputAmountVN('VNĐ (25,000,000)'), {
    numeric: -25000000,
    formatted: '-25.000.000',
  });
  assert.deepEqual(formatInputAmountVN('(500.000) đ'), {
    numeric: -500000,
    formatted: '-500.000',
  });
  assert.deepEqual(formatInputAmountVN('1.250.000-'), {
    numeric: -1250000,
    formatted: '-1.250.000',
  });
  assert.deepEqual(formatInputAmountVN('500,000-'), {
    numeric: -500000,
    formatted: '-500.000',
  });
});

// ============================================================================
// PHẦN 2: TRA TẤN `readAmountInVietnameseWords` (CHUẨN CHÍNH TẢ KẾ TOÁN VN)
// ============================================================================

test('2.1 readAmountInVietnameseWords: Quy tắc lẻ, mười, mười lăm, mươi mốt, mươi tư, mươi lăm, tỷ, nghìn tỷ, số âm & thập phân', () => {
  assert.equal(readAmountInVietnameseWords(0), 'Không đồng');
  assert.equal(readAmountInVietnameseWords(-0), 'Không đồng');
  assert.equal(readAmountInVietnameseWords(5), 'Năm đồng');
  assert.equal(readAmountInVietnameseWords(10), 'Mười đồng');

  // 15 -> "Mười lăm đồng" (TUYỆT ĐỐI KHÔNG đọc "Mười năm đồng")
  assert.equal(readAmountInVietnameseWords(15), 'Mười lăm đồng');

  // 14 -> "Mười bốn đồng" (KHÔNG đọc "Mười tư đồng")
  assert.equal(readAmountInVietnameseWords(14), 'Mười bốn đồng');

  // 21 -> "Hai mươi mốt đồng" (KHÔNG đọc "Hai mươi một đồng")
  assert.equal(readAmountInVietnameseWords(21), 'Hai mươi mốt đồng');

  // 24 -> "Hai mươi tư đồng"
  assert.equal(readAmountInVietnameseWords(24), 'Hai mươi tư đồng');

  // 25 -> "Hai mươi lăm đồng"
  assert.equal(readAmountInVietnameseWords(25), 'Hai mươi lăm đồng');

  // 101 -> "Một trăm lẻ một đồng" (KHÔNG đọc "lẻ mốt")
  assert.equal(readAmountInVietnameseWords(101), 'Một trăm lẻ một đồng');

  // 104 -> "Một trăm lẻ bốn đồng" (KHÔNG đọc "lẻ tư")
  assert.equal(readAmountInVietnameseWords(104), 'Một trăm lẻ bốn đồng');

  // 105 -> "Một trăm lẻ năm đồng" (TUYỆT ĐỐI KHÔNG đọc "lẻ lăm")
  assert.equal(readAmountInVietnameseWords(105), 'Một trăm lẻ năm đồng');

  // 115 -> "Một trăm mười lăm đồng"
  assert.equal(readAmountInVietnameseWords(115), 'Một trăm mười lăm đồng');

  // Hàng nghìn & triệu có nhóm 0 ở giữa
  assert.equal(
    readAmountInVietnameseWords(1_005),
    'Một nghìn không trăm lẻ năm đồng'
  );
  assert.equal(
    readAmountInVietnameseWords(1_015),
    'Một nghìn không trăm mười lăm đồng'
  );
  assert.equal(
    readAmountInVietnameseWords(2_004),
    'Hai nghìn không trăm lẻ bốn đồng'
  );
  assert.equal(
    readAmountInVietnameseWords(2_024),
    'Hai nghìn không trăm hai mươi tư đồng'
  );
  assert.equal(
    readAmountInVietnameseWords(1_000_005),
    'Một triệu không trăm lẻ năm đồng'
  );
  assert.equal(
    readAmountInVietnameseWords(1_050_000),
    'Một triệu không trăm năm mươi nghìn đồng'
  );
  assert.equal(
    readAmountInVietnameseWords(10_000_001),
    'Mười triệu không trăm lẻ một đồng'
  );

  // Hàng tỷ & nghìn tỷ
  assert.equal(readAmountInVietnameseWords(1_000_000_000), 'Một tỷ đồng');
  assert.equal(
    readAmountInVietnameseWords(1_000_000_001),
    'Một tỷ không trăm lẻ một đồng'
  );
  assert.equal(
    readAmountInVietnameseWords(1_250_000_000_000),
    'Một nghìn hai trăm năm mươi tỷ đồng'
  );

  // Số âm & số thập phân & chuỗi định dạng kế toán
  assert.equal(
    readAmountInVietnameseWords(-18_450_000),
    'Âm mười tám triệu bốn trăm năm mươi nghìn đồng'
  );
  assert.equal(
    readAmountInVietnameseWords('(18.450.000 đ)'),
    'Âm mười tám triệu bốn trăm năm mươi nghìn đồng'
  );
  assert.equal(
    readAmountInVietnameseWords(1250.5),
    'Một nghìn hai trăm năm mươi phẩy năm đồng'
  );
  assert.equal(
    readAmountInVietnameseWords(1250.05),
    'Một nghìn hai trăm năm mươi phẩy không năm đồng'
  );
  assert.equal(
    readAmountInVietnameseWords(1250.25),
    'Một nghìn hai trăm năm mươi phẩy hai mươi lăm đồng'
  );

  // Giá trị rỗng/không hợp lệ
  assert.equal(readAmountInVietnameseWords(null), '');
  assert.equal(readAmountInVietnameseWords(''), '');
  assert.equal(readAmountInVietnameseWords(NaN), '');
});

test('2.2 [ADVERSARIAL] readAmountInVietnameseWords & formatNumberVN: Số nguyên an toàn lớn (500 nghìn tỷ lẻ 1 đồng) không được bị biến dạng thập phân giả', () => {
  const bigSafeInt = 500_000_000_000_001; // Number.isSafeInteger(500_000_000_000_001) === true
  assert.equal(Number.isSafeInteger(bigSafeInt), true);
  assert.equal(formatNumberVN(bigSafeInt), '500.000.000.000.001');
  assert.equal(
    readAmountInVietnameseWords(bigSafeInt),
    'Năm trăm nghìn tỷ không trăm lẻ một đồng'
  );
});

// ============================================================================
// PHẦN 3: TRA TẤN `parseVietnameseDate` VỚI LỊCH THỰC TẾ & SERIAL CÓ GIỜ PHÚT
// ============================================================================

test('3.1 parseVietnameseDate: Kiểm tra nghiêm ngặt năm nhuận (2024, 2000 vs 2026, 1900) & ngày biên tháng', () => {
  // Năm nhuận hợp lệ
  const leap2024 = parseVietnameseDate('29/02/2024');
  assert.ok(leap2024);
  assert.equal(leap2024.iso, '2024-02-29');
  assert.equal(leap2024.display, '29/02/2024');
  assert.equal(leap2024.timestamp, Date.UTC(2024, 1, 29));

  const leap2000 = parseVietnameseDate('29/02/2000');
  assert.equal(leap2000?.iso, '2000-02-29');

  // Năm không nhuận & thế kỷ không nhuận -> PHẢI trả về null (không được tràn sang 01/03)
  assert.equal(parseVietnameseDate('29/02/2026'), null);
  assert.equal(parseVietnameseDate('29/02/1900'), null);

  // Tháng 30 ngày & ngày ngoài biên -> null
  assert.equal(parseVietnameseDate('31/04/2026'), null);
  assert.equal(parseVietnameseDate('31/06/2026'), null);
  assert.equal(parseVietnameseDate('31/09/2026'), null);
  assert.equal(parseVietnameseDate('31/11/2026'), null);
  assert.equal(parseVietnameseDate('00/03/2026'), null);
  assert.equal(parseVietnameseDate('15/00/2026'), null);
  assert.equal(parseVietnameseDate('15/13/2026'), null);
});

test('3.2 parseVietnameseDate: Chuỗi sao kê có giờ phút & Số Serial Excel số thực kèm giờ phút (46086.60789)', () => {
  // Sao kê ngân hàng dạng chuỗi có giờ phút giây
  const strWithTime = parseVietnameseDate('05/03/2026 14:35:22');
  assert.ok(strWithTime);
  assert.equal(strWithTime.iso, '2026-03-05');
  assert.equal(strWithTime.display, '05/03/2026');
  assert.equal(strWithTime.timestamp, Date.UTC(2026, 2, 5));

  // Giờ đứng trước ngày
  assert.equal(parseVietnameseDate('14:35:22 05/03/2026')?.iso, '2026-03-05');
  assert.equal(parseVietnameseDate('02:35:22 CH 05/03/2026')?.iso, '2026-03-05');
  assert.equal(parseVietnameseDate('2026-03-05T14:35:22.000Z')?.iso, '2026-03-05');

  // ĐẶC BIỆT QUAN TRỌNG: Excel Serial dạng số thực kèm giờ phút (VCB / BIDV: 46086.60789 = 05/03/2026 14:35:22)
  const serialFloat = parseVietnameseDate(46086.60789);
  assert.ok(serialFloat, 'Excel float serial 46086.60789 phải parse được ngày hợp lệ');
  assert.equal(serialFloat.iso, '2026-03-05');
  assert.equal(serialFloat.display, '05/03/2026');
  assert.equal(serialFloat.timestamp, Date.UTC(2026, 2, 5));

  // Giao dịch sát nửa đêm 23:59:59 ngày 05/03/2026 (46086 + 86399/86400 = 46086.999988425926)
  const lateNightSerial = parseVietnameseDate(46086 + 86399 / 86400);
  assert.equal(
    lateNightSerial?.iso,
    '2026-03-05',
    'Giao dịch lúc 23:59:59 không được bị làm tròn trượt sang ngày hôm sau 06/03/2026'
  );

  // Chống nhận nhầm số lượng nhỏ / đơn giá (100, 5000) thành Excel date serial
  assert.equal(parseVietnameseDate(100), null);
  assert.equal(parseVietnameseDate(5000), null);
  assert.equal(parseVietnameseDate(32874)?.iso, '1990-01-01');
  assert.equal(parseVietnameseDate(73415)?.iso, '2100-12-31');
  assert.equal(parseVietnameseDate(new Date(2026, 2, 5, 14, 35, 22))?.iso, '2026-03-05');
  assert.equal(parseVietnameseDate(new Date('invalid')), null);
});

test('3.3 parseVietnameseDate: Bẫy định dạng 05/03/2026 vs 03/25/2026 vs 2026/03/05 vs 5.3.26 & văn xuôi', () => {
  assert.equal(parseVietnameseDate('05/03/2026')?.iso, '2026-03-05');
  assert.equal(parseVietnameseDate('03/25/2026')?.iso, '2026-03-25');
  assert.equal(parseVietnameseDate('2026/03/05')?.iso, '2026-03-05');
  assert.equal(parseVietnameseDate('5.3.26')?.iso, '2026-03-05');
  assert.equal(parseVietnameseDate('Ngày 05 tháng 03 năm 2026')?.iso, '2026-03-05');

  // Năm đứng trước nhưng tháng > 12 -> phải từ chối, không được tự ý đảo tháng/ngày
  assert.equal(parseVietnameseDate('2026/25/03'), null);
});

test('3.4 [ADVERSARIAL] parseVietnameseDate: Không được nhận nhầm câu diễn giải bắt đầu bằng ngày tháng thành ô ngày thuần túy', () => {
  // Chuỗi diễn giải kiểu "05/03/2026 Thu tiền khách hàng Minh Phú" không phải là giá trị ngày hợp lệ
  assert.equal(parseVietnameseDate('05/03/2026 Thu tiền khách hàng Minh Phú'), null);
});

// ============================================================================
// PHẦN 4: TRA TẤN `excelParser.js` VỚI EXCEL KẾ TOÁN "BẨN" (TT200/133, GỘP Ô,
//         BẪY TÊN CÔNG TY `isSummaryRow`, `debit_minus_credit`, CSV & BIFF8 `.xls`)
// ============================================================================

test('4.1 excelParser.js: Tiêu đề 2 dòng gộp ô chuẩn TT200 (Chứng từ -> Số hiệu|Ngày tháng; Số phát sinh -> Nợ|Có) + dòng phụ (1)|(2) & A|B|C|1|2', () => {
  const aoa = [
    ['Đơn vị: CÔNG TY CỔ PHẦN CƠ ĐIỆN ĐẠI VIỆT', '', '', '', '', '', 'Mẫu số S03b-DN'],
    ['Địa chỉ: KCN Tân Bình, TP.HCM', '', '', '', '', '', '(Ban hành theo TT 200/2014/TT-BTC)'],
    ['SỔ CÁI TÀI KHOẢN 131 - PHẢI THU CỦA KHÁCH HÀNG', '', '', '', '', '', ''],
    ['Tháng 03 năm 2026', '', '', '', '', '', ''],
    ['', '', '', '', '', '', ''],
    // Row 5 (index 5): Parent merged header
    ['Ngày tháng ghi sổ', 'Chứng từ', '', 'Diễn giải', 'TK đối ứng', 'Số phát sinh', ''],
    // Row 6 (index 6): Child sub-header
    ['', 'Số hiệu', 'Ngày tháng', '', '', 'Nợ', 'Có'],
    // Row 7 (index 7): Column index helper row A | B | C | D | E | 1 | 2
    ['A', 'B', 'C', 'D', 'E', '1', '2'],
    // Data rows
    ['05/03/2026', 'HD-001', '05/03/2026', 'Bán hàng cho Công ty Minh Phát', '5111', 25000000, 0],
    ['10/03/2026', 'UNC-012', '10/03/2026', 'Khách hàng Minh Phát thanh toán đợt 1', '1121', 0, 10000000],
    ['15/03/2026', 'HD-002', '15/03/2026', 'Cung cấp dịch vụ lắp đặt trạm biến áp', '5113', 18500000, ''],
    ['', '', '', 'Cộng phát sinh tháng 03', '', 43500000, 10000000],
    ['', '', '', 'Số dư cuối kỳ', '', 33500000, 0],
  ];

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!merges'] = [
    { s: { r: 5, c: 0 }, e: { r: 6, c: 0 } }, // Ngày tháng ghi sổ
    { s: { r: 5, c: 1 }, e: { r: 5, c: 2 } }, // Chứng từ -> Số hiệu | Ngày tháng
    { s: { r: 5, c: 3 }, e: { r: 6, c: 3 } }, // Diễn giải
    { s: { r: 5, c: 4 }, e: { r: 6, c: 4 } }, // TK đối ứng
    { s: { r: 5, c: 5 }, e: { r: 5, c: 6 } }, // Số phát sinh -> Nợ | Có
  ];
  const wbIn = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wbIn, ws, 'SoCai_131');
  const xlsxBytes = XLSX.write(wbIn, { bookType: 'xlsx', type: 'array' });

  const wb = loadWorkbookFromBuffer(xlsxBytes);
  const inspection = inspectSheet(wb, 'SoCai_131');

  assert.equal(inspection.detectedHeaderRowIndex, 6, 'Phải nhận diện dòng tiêu đề con ở index 6');
  assert.equal(inspection.suggestedMapping.amountMode, 'debit_minus_credit');
  assert.equal(inspection.suggestedMapping.debitColIndex, 5);
  assert.equal(inspection.suggestedMapping.creditColIndex, 6);
  assert.equal(inspection.suggestedMapping.descColIndex, 3);
  assert.equal(inspection.suggestedMapping.refColIndex, 1);

  const extracted = extractTransactions(
    inspection.rawGrid,
    inspection.detectedHeaderRowIndex,
    inspection.suggestedMapping,
    inspection.columns
  );

  // Dòng phụ A|B|C|D|E|1|2 phải bị bỏ qua hoàn toàn
  const businessItems = extracted.items.filter((i) => !i.isSummaryRow);
  assert.equal(businessItems.length, 3);
  assert.deepEqual(
    businessItems.map((i) => i.amount),
    [25000000, -10000000, 18500000]
  );
  assert.equal(extracted.summaryRowCount, 2);

  // Kiểm tra thêm biến thể dòng phụ đánh số (1) | (2) | ... | (7)
  const aoaNumberedHelper = aoa.map((row, idx) =>
    idx === 7 ? ['(1)', '(2)', '(3)', '(4)', '(5)', '(6)', '(7)'] : [...row]
  );
  const insp2 = inspectRawGrid(aoaNumberedHelper, 'SoCai_131_HelperNum');
  const ext2 = extractTransactions(
    insp2.rawGrid,
    insp2.detectedHeaderRowIndex,
    insp2.suggestedMapping,
    insp2.columns
  );
  assert.equal(ext2.items.filter((i) => !i.isSummaryRow).length, 3);
});

test('4.2 [ADVERSARIAL] excelParser.js: Tiêu đề 2 dòng gộp ô kiểu Sổ chi tiết công nợ MISA/TT200 (chỉ tách Số phát sinh -> Nợ | Có) + dòng phụ (1)..(6)', () => {
  // Mẫu cực kỳ phổ biến: các cột Ngày CT, Số CT, Diễn giải, TK đối ứng gộp dọc 2 dòng; chỉ cột Số phát sinh tách ngang thành Nợ | Có ở dòng dưới
  const grid = [
    ['CÔNG TY TNHH SẢN XUẤT THƯƠNG MẠI ĐÔNG Á', '', '', '', '', ''],
    ['SỔ CHI TIẾT CÔNG NỢ PHẢI THU (TK 131)', '', '', '', '', ''],
    ['', '', '', '', '', ''],
    // Row 3: Parent header
    ['Ngày chứng từ', 'Số chứng từ', 'Diễn giải', 'TK đối ứng', 'Số phát sinh', ''],
    // Row 4: Sub-header (only Nợ | Có split)
    ['', '', '', '', 'Nợ', 'Có'],
    // Row 5: Column index helper row
    ['(1)', '(2)', '(3)', '(4)', '(5)', '(6)'],
    ['02/03/2026', 'HD-101', 'Xuất bán lô hàng thiết bị điện đợt 1', '5111', 30000000, ''],
    ['08/03/2026', 'BC-205', 'Khách hàng chuyển khoản thanh toán', '1121', '', 12000000],
    ['', '', 'Tổng cộng', '', 30000000, 12000000],
  ];

  const inspection = inspectRawGrid(grid, 'SoChiTiet_MISA');
  assert.equal(
    inspection.suggestedMapping.amountMode,
    'debit_minus_credit',
    'Phải nhận diện chế độ debit_minus_credit khi tiêu đề 2 dòng tách Số phát sinh -> Nợ | Có'
  );
  assert.equal(inspection.suggestedMapping.debitColIndex, 4);
  assert.equal(inspection.suggestedMapping.creditColIndex, 5);

  const extracted = extractTransactions(
    inspection.rawGrid,
    inspection.detectedHeaderRowIndex,
    inspection.suggestedMapping,
    inspection.columns
  );
  const businessItems = extracted.items.filter((i) => !i.isSummaryRow);
  assert.equal(businessItems.length, 2, 'Không được làm mất dòng phát sinh Có!');
  assert.deepEqual(
    businessItems.map((i) => i.amount),
    [30000000, -12000000]
  );
});

test('4.3 excelParser.js: Bẫy nhận nhầm `isSummaryRow` với tên khách hàng/diễn giải chứa chữ "Tổng", "Cộng" vs dòng tổng thật', () => {
  const grid = [
    ['Ngày chứng từ', 'Số chứng từ', 'Khách hàng / Đối tác', 'Diễn giải', 'Số tiền'],
    // Số dư đầu kỳ -> PHẢI là summary
    ['01/03/2026', '', '', 'Số dư đầu kỳ', 50000000],
    // 4 dòng nghiệp vụ thật có chữ "Tổng" / "Cộng" -> TUYỆT ĐỐI KHÔNG được gắn isSummaryRow = true
    [
      '05/03/2026',
      'HD-001',
      'Tổng công ty Cổ phần Thiết bị Điện',
      'Bán máy biến áp 3 pha',
      45000000,
    ],
    [
      '08/03/2026',
      'HD-002',
      'Công ty TNHH Cộng Đồng Xanh',
      'Cung cấp vật tư cây xanh đô thị',
      18500000,
    ],
    [
      '12/03/2026',
      'HD-003',
      'Cửa hàng Tổng hợp Hòa Bình',
      'Giao hàng thiết bị chiếu sáng',
      12300000,
    ],
    [
      '15/03/2026',
      'PT-004',
      'Công ty TNHH Vận tải Biển Đông',
      'Thu phí cộng thêm ca đêm bốc xếp container',
      4200000,
    ],
    // Các dòng tổng cộng / số dư cuối kỳ -> PHẢI gắn isSummaryRow = true
    ['', '', '', 'Cộng phát sinh tháng 03', 80000000],
    ['', '', '', 'Tổng cộng', 80000000],
    ['31/03/2026', '', '', 'Số dư cuối kỳ', 130000000],
    ['', '', '', 'Grand Total', 130000000],
  ];

  const inspection = inspectRawGrid(grid, 'TrapSummaryTest');
  const extracted = extractTransactions(
    inspection.rawGrid,
    inspection.detectedHeaderRowIndex,
    inspection.suggestedMapping,
    inspection.columns
  );

  const byDescOrPartner = new Map(
    extracted.items.map((it) => [it.description || it.extraCells['Khách hàng / Đối tác'], it])
  );

  // Kiểm tra 4 dòng nghiệp vụ thật KHÔNG bị gắn nhầm
  const realRows = extracted.items.filter((it) =>
    ['HD-001', 'HD-002', 'HD-003', 'PT-004'].includes(it.reference)
  );
  assert.equal(realRows.length, 4);
  for (const row of realRows) {
    assert.equal(
      row.isSummaryRow,
      false,
      `Dòng nghiệp vụ thật "${row.reference} - ${row.description}" bị gắn nhầm isSummaryRow = true!`
    );
  }

  // Kiểm tra 5 dòng tổng / số dư PHẢI là isSummaryRow = true
  assert.equal(byDescOrPartner.get('Số dư đầu kỳ')?.isSummaryRow, true);
  assert.equal(byDescOrPartner.get('Cộng phát sinh tháng 03')?.isSummaryRow, true);
  assert.equal(byDescOrPartner.get('Tổng cộng')?.isSummaryRow, true);
  assert.equal(byDescOrPartner.get('Số dư cuối kỳ')?.isSummaryRow, true);
  assert.equal(byDescOrPartner.get('Grand Total')?.isSummaryRow, true);
  assert.equal(extracted.summaryRowCount, 5);
});

test('4.4 [ADVERSARIAL] excelParser.js: Bẫy khử dấu tiếng Việt làm nhận nhầm "Tổng công trình" / "Tổng công nghiệp" thành "Tổng cộng"', () => {
  const grid = [
    ['Ngày chứng từ', 'Số chứng từ', 'Đối tác', 'Diễn giải', 'Số tiền'],
    [
      '18/03/2026',
      'HD-088',
      'Tổng công nghiệp Xây lắp Miền Trung',
      'Thanh toán đợt 2 Tổng công trình Cầu Thủ Thiêm',
      95000000,
    ],
    // Dòng chi tiết hạch toán gộp chung chứng từ (kế toán để trống ô Ngày ở dòng thứ 2 cùng phiếu)
    [
      '',
      'HD-089',
      'Tổng công ty Cổ phần Thiết bị Điện',
      'Tổng công ty Cổ phần Thiết bị Điện - Phụ kiện đi kèm',
      5000000,
    ],
    ['', '', '', 'Tổng cộng', 100000000],
  ];

  const inspection = inspectRawGrid(grid, 'DiacriticSummaryTrap');
  const extracted = extractTransactions(
    inspection.rawGrid,
    inspection.detectedHeaderRowIndex,
    inspection.suggestedMapping,
    inspection.columns
  );

  const hd088 = extracted.items.find((i) => i.reference === 'HD-088');
  const hd089 = extracted.items.find((i) => i.reference === 'HD-089');
  assert.ok(hd088);
  assert.ok(hd089);
  assert.equal(
    hd088.isSummaryRow,
    false,
    '"Tổng công trình" / "Tổng công nghiệp" (chữ CÔNG, không phải CỘNG) có ngày & số HĐ đầy đủ không được nhận nhầm thành dòng Tổng cộng!'
  );
  assert.equal(
    hd089.isSummaryRow,
    false,
    'Dòng chi tiết của "Tổng công ty Cổ phần Thiết bị Điện" không được bị nhận nhầm là dòng tổng!'
  );
});

test('4.5 excelParser.js: Chế độ `debit_minus_credit` (chỉ có Nợ -> dương, chỉ có Có -> âm, rỗng / gạch ngang / 0-0 trống -> bỏ qua)', () => {
  const grid = [
    ['Ngày giao dịch', 'Số tham chiếu', 'Nội dung giao dịch', 'Phát sinh Nợ', 'Phát sinh Có'],
    [46086.60789, 'FT26064001', 'Thu tiền HĐ 01', '15.000.000', '-'],
    ['06/03/2026 09:15:00', 'FT26065002', 'Hoàn trả chiết khấu thương mại', '-', '3.250.000'],
    ['07/03/2026', 'FT26066003', 'Bù trừ công nợ hai chiều', '10.000.000', '2.500.000'],
    // Các dòng trống hoặc cả Nợ và Có = '-' / '' / 0 không có chứng từ -> bỏ qua
    ['', '', '', '', ''],
    ['', '', '', '-', '-'],
    ['', '', '', 0, 0],
    ['', '', 'Tổng cộng', '25.000.000', '5.750.000'],
  ];

  const inspection = inspectRawGrid(grid, 'BankDebitCredit');
  assert.equal(inspection.suggestedMapping.amountMode, 'debit_minus_credit');
  assert.equal(inspection.suggestedMapping.debitColIndex, 3);
  assert.equal(inspection.suggestedMapping.creditColIndex, 4);

  const extracted = extractTransactions(
    inspection.rawGrid,
    inspection.detectedHeaderRowIndex,
    inspection.suggestedMapping,
    inspection.columns
  );

  const txItems = extracted.items.filter((i) => !i.isSummaryRow);
  assert.equal(txItems.length, 3);
  assert.equal(txItems[0].dateIso, '2026-03-05'); // từ serial float 46086.60789
  assert.equal(txItems[0].amount, 15000000);
  assert.equal(txItems[1].dateIso, '2026-03-06');
  assert.equal(txItems[1].amount, -3250000);
  assert.equal(txItems[2].dateIso, '2026-03-07');
  assert.equal(txItems[2].amount, 7500000);
  assert.equal(extracted.skippedEmptyOrInvalidCount, 3);
  assert.equal(extracted.summaryRowCount, 1);
});

test('4.6 [ADVERSARIAL] excelParser.js: Dòng hóa đơn hủy / dòng tiêu đề nhóm có Nợ = 0 và Có = 0 trong chế độ `debit_minus_credit` nên được bỏ qua nhất quán như dòng Nợ/Có trống', () => {
  const grid = [
    ['Ngày chứng từ', 'Số chứng từ', 'Diễn giải', 'Phát sinh Nợ', 'Phát sinh Có'],
    ['05/03/2026', 'HD-001', 'Bán hàng đợt 1', 12000000, 0],
    // Khi kế toán xuất từ phần mềm, dòng hóa đơn xóa bỏ hoặc dòng tiêu đề nhóm khách hàng thường có Nợ = 0, Có = 0 (thay vì ô trống)
    ['', '', 'I. KHÁCH HÀNG MIỀN BẮC', 0, 0],
    ['06/03/2026', 'HD-002', 'Hóa đơn xóa bỏ (không phát sinh tiền)', 0, 0],
    ['07/03/2026', 'PT-001', 'Thu tiền đợt 1', 0, 5000000],
  ];

  const inspection = inspectRawGrid(grid, 'ZeroDebitCreditRows');
  const extracted = extractTransactions(
    inspection.rawGrid,
    inspection.detectedHeaderRowIndex,
    inspection.suggestedMapping,
    inspection.columns
  );

  const txItems = extracted.items.filter((i) => !i.isSummaryRow);
  assert.deepEqual(
    txItems.map((i) => ({ ref: i.reference, amount: i.amount })),
    [
      { ref: 'HD-001', amount: 12000000 },
      { ref: 'PT-001', amount: -5000000 },
    ],
    'Các dòng Nợ = 0 & Có = 0 (tiêu đề nhóm hoặc hóa đơn hủy không phát sinh tiền) không nên lọt vào danh sách giao dịch gây nhiễu bộ gom tổ hợp'
  );
});

test('4.7 excelParser.js: Đọc CSV tiếng Việt UTF-8 CÓ BOM (\\uFEFF) và file Excel cũ `.xls` (BIFF8)', () => {
  const csvBody = [
    'Ngày chứng từ,Số chứng từ,Đối tác,Diễn giải,Số tiền',
    '05/03/2026,HD-01,"Tổng công ty Cổ phần Thiết bị Điện","Bán dây cáp điện hạ thế","25.000.000"',
    '08/03/2026,CK-02,"Công ty TNHH Cộng Đồng Xanh","Chiết khấu thanh toán sớm","(1.500.000)"',
    ',,,"Tổng cộng","23.500.000"',
  ].join('\r\n');

  // 1) CSV UTF-8 có BOM (\uFEFF)
  const bomBytes = new TextEncoder().encode('\uFEFF' + csvBody);
  const wbCsvBom = loadWorkbookFromBuffer(bomBytes);
  const inspBom = inspectSheet(wbCsvBom, getSheetNames(wbCsvBom)[0]);
  const extBom = extractTransactions(
    inspBom.rawGrid,
    inspBom.detectedHeaderRowIndex,
    inspBom.suggestedMapping,
    inspBom.columns
  );

  const bomBusiness = extBom.items.filter((i) => !i.isSummaryRow);
  assert.equal(bomBusiness.length, 2);
  assert.equal(bomBusiness[0].description, 'Bán dây cáp điện hạ thế');
  assert.equal(bomBusiness[0].amount, 25000000);
  assert.equal(bomBusiness[1].description, 'Chiết khấu thanh toán sớm');
  assert.equal(bomBusiness[1].amount, -1500000);
  assert.equal(extBom.summaryRowCount, 1);

  // 2) File `.xls` (BIFF8) cũ xuất từ hệ thống kế toán / ngân hàng đời cũ
  const xlsAoa = [
    ['CÔNG TY CỔ PHẦN THÉP NAM ĐỊNH', '', '', '', ''],
    ['SỔ THEO DÕI CÔNG NỢ THÁNG 03/2026', '', '', '', ''],
    ['Ngày chứng từ', 'Số chứng từ', 'Khách hàng', 'Diễn giải', 'Số tiền (VNĐ)'],
    [46086.60789, 'HD-901', 'Cửa hàng Tổng hợp Hòa Bình', 'Xuất thép cuộn Hòa Phát', 42500000],
    ['12/03/2026', 'UNC-902', 'Công ty TNHH Cộng Đồng Xanh', 'Giảm trừ hàng lỗi kỹ thuật', -2500000],
    ['', '', '', 'Tổng cộng', 40000000],
  ];
  const wsXls = XLSX.utils.aoa_to_sheet(xlsAoa);
  const wbXlsIn = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wbXlsIn, wsXls, 'CongNo_BIFF8');
  const biff8Bytes = XLSX.write(wbXlsIn, { bookType: 'biff8', type: 'array' });

  const wbBiff8 = loadWorkbookFromBuffer(biff8Bytes);
  const inspBiff8 = inspectSheet(wbBiff8, 'CongNo_BIFF8');
  assert.equal(inspBiff8.detectedHeaderRowIndex, 2);

  const extBiff8 = extractTransactions(
    inspBiff8.rawGrid,
    inspBiff8.detectedHeaderRowIndex,
    inspBiff8.suggestedMapping,
    inspBiff8.columns
  );
  const biff8Business = extBiff8.items.filter((i) => !i.isSummaryRow);
  assert.equal(biff8Business.length, 2);
  assert.equal(biff8Business[0].dateIso, '2026-03-05');
  assert.equal(biff8Business[0].description, 'Xuất thép cuộn Hòa Phát');
  assert.equal(biff8Business[0].amount, 42500000);
  assert.equal(biff8Business[1].amount, -2500000);
  assert.equal(extBiff8.summaryRowCount, 1);
});

test('4.8 [ADVERSARIAL] excelParser.js: Đọc CSV tiếng Việt UTF-8 KHÔNG CÓ BOM không được bị lỗi font (mojibake) và nhận diện đúng cột', () => {
  const csvNoBom = [
    'Ngày chứng từ,Số chứng từ,Đối tác,Diễn giải,Số tiền',
    '05/03/2026,HD-01,"Tổng công ty Cổ phần Thiết bị Điện","Thu tiền bán máy biến áp","25.000.000"',
    '08/03/2026,CK-02,"Công ty TNHH Cộng Đồng Xanh","Thu phí cộng thêm ca đêm","4.200.000"',
    ',,,"Tổng cộng","29.200.000"',
  ].join('\n');

  const rawUtf8Bytes = new TextEncoder().encode(csvNoBom);
  const wbNoBom = loadWorkbookFromBuffer(rawUtf8Bytes);
  const inspNoBom = inspectSheet(wbNoBom, getSheetNames(wbNoBom)[0]);
  const extNoBom = extractTransactions(
    inspNoBom.rawGrid,
    inspNoBom.detectedHeaderRowIndex,
    inspNoBom.suggestedMapping,
    inspNoBom.columns
  );

  // Tiếng Việt có dấu không BOM phải đọc nguyên vẹn, không bị biến thành "NgÃ y chá»©ng tá»«"
  assert.equal(inspNoBom.columns[0].label, 'Ngày chứng từ');
  assert.equal(inspNoBom.columns[3].label, 'Diễn giải');
  assert.equal(inspNoBom.suggestedMapping.dateColIndex, 0);
  assert.equal(inspNoBom.suggestedMapping.amountColIndex, 4);

  const txItems = extNoBom.items.filter((i) => !i.isSummaryRow);
  assert.equal(txItems.length, 2);
  assert.equal(txItems[0].description, 'Thu tiền bán máy biến áp');
  assert.equal(txItems[1].description, 'Thu phí cộng thêm ca đêm');
  assert.equal(extNoBom.summaryRowCount, 1);
});

test('4.9 excelParser.js: CSV dấu chấm phẩy `;` (Windows VN Locale) & Xuất Excel đối soát roundtrip (`unallocatedOnly`)', () => {
  // Trên Windows cấu hình Region Việt Nam (dấu thập phân là `,`), Excel xuất CSV dùng dấu `;` ngăn cách cột!
  const csvSemicolon = [
    'Ngày chứng từ;Số chứng từ;Đối tác;Diễn giải;Số tiền',
    '05/03/2026;HD-01;Tổng công ty Cổ phần Thiết bị Điện;Bán cáp ngầm trung thế;25.000.000,50',
    '09/03/2026;UNC-02;Công ty TNHH Cộng Đồng Xanh;Giảm giá hàng bán;(1.250.000,50)',
    '12/03/2026;HD-03;Cửa hàng Tổng hợp Hòa Bình;Bán tủ điện phân phối;10.000.000',
    ';;;Tổng cộng;33.750.000',
  ].join('\r\n');

  const semiBytes = new TextEncoder().encode('\uFEFF' + csvSemicolon);
  const wbSemi = loadWorkbookFromBuffer(semiBytes);
  const inspSemi = inspectSheet(wbSemi, getSheetNames(wbSemi)[0]);
  const extSemi = extractTransactions(
    inspSemi.rawGrid,
    inspSemi.detectedHeaderRowIndex,
    inspSemi.suggestedMapping,
    inspSemi.columns
  );

  const semiTx = extSemi.items.filter((i) => !i.isSummaryRow);
  assert.equal(semiTx.length, 3);
  assert.deepEqual(
    semiTx.map((i) => i.amount),
    [25000000.5, -1250000.5, 10000000]
  );
  assert.equal(extSemi.summaryRowCount, 1);

  // Kiểm tra xuất workbook đối soát với unallocatedOnly = true (chỉ giữ lại dòng chưa gom, loại cả dòng đã gom & dòng tổng cộng)
  const outBytes = exportReconciledWorkbook({
    sheetName: inspSemi.sheetName,
    columns: inspSemi.columns,
    items: extSemi.items,
    reconciledBatches: [
      {
        code: 'DOT-01',
        targetAmount: 23750000,
        items: [semiTx[0], semiTx[1]],
      },
    ],
    unallocatedOnly: true,
  });

  const wbOut = loadWorkbookFromBuffer(outBytes);
  assert.deepEqual(getSheetNames(wbOut), ['Chi_Tiet_Doi_Soat', 'Tong_Hop_Cac_Dot_Gom']);
  const sheet1Grid = inspectSheet(wbOut, 'Chi_Tiet_Doi_Soat', 0).rawGrid;
  // Header + duy nhất 1 dòng chưa gom (HD-03: 10.000.000)
  assert.equal(sheet1Grid.length, 2);
  assert.equal(sheet1Grid[1][2], 'HD-03');
  assert.equal(sheet1Grid[1][4], 10000000);
  assert.equal(sheet1Grid[1][5], 'Chưa gom');
});

