import { describe, it } from 'node:test';
import assert from 'node:assert/strict';

import {
  loadWorkbookFromBuffer,
  getSheetNames,
  inspectSheet,
  inspectRawGrid,
  extractTransactions,
  exportReconciledWorkbook,
} from '../src/core/excelParser.js';
import {
  findCombinations,
  normalizeVietnameseSearch,
} from '../src/core/reconcileEngine.js';

/**
 * Dựng sổ chi tiết công nợ phải thu thực tế (TK 131 - Quý 1/2026) gồm ĐÚNG 80 dòng trích xuất:
 * - 4 dòng tiêu đề đơn vị & tên sổ ở đầu sheet (row 0..3), dòng header ở row 4
 * - 1 dòng Số dư đầu kỳ (isSummaryRow = true)
 * - Tháng 01/2026: 25 dòng nghiệp vụ thực + 1 dòng Cộng phát sinh tháng 01/2026 (isSummaryRow = true)
 * - Tháng 02/2026: 25 dòng nghiệp vụ thực + 1 dòng Cộng phát sinh tháng 02/2026 (isSummaryRow = true)
 * - Tháng 03/2026: 25 dòng nghiệp vụ thực + 1 dòng Cộng phát sinh tháng 03/2026 (isSummaryRow = true)
 * - 1 dòng Tổng cộng phát sinh cuối kỳ (isSummaryRow = true)
 * Tổng cộng: 75 dòng nghiệp vụ thực + 5 dòng tổng hợp/đầu kỳ/cuối kỳ = 80 dòng.
 *
 * Bất biến mô-đun chống nhiễu tổ hợp ngẫu nhiên:
 * Các hóa đơn nền (filler) đều có phần dư ≡ +1.100 (mod 10.000) để không bao giờ vô tình
 * tạo tổ hợp trùng với các khoản chuyển khoản chẵn chục nghìn của 5 đợt mục tiêu.
 */
function buildRealistic80RowLedger() {
  const bannerRows = [
    ['CÔNG TY CỔ PHẦN VẬT TƯ & THIẾT BỊ CÔNG NGHIỆP ĐẠI VIỆT', '', '', '', '', ''],
    ['Địa chỉ: Khu Công nghiệp Hòa Khánh, Quận Liên Chiểu, TP. Đà Nẵng', '', '', '', '', ''],
    ['SỔ CHI TIẾT CÔNG NỢ PHẢI THU KHÁCH HÀNG (TK 131) - QUÝ 1/2026', '', '', '', '', ''],
    ['', '', '', '', '', ''],
  ];

  const headerRow = [
    'Ngày chứng từ',
    'Số chứng từ',
    'Khách hàng / Đối tác',
    'Diễn giải nội dung',
    'Số tiền phải thu',
    'Ghi chú',
  ];

  const openingRow = [
    '',
    'DK-2026',
    '',
    'Số dư đầu kỳ mang sang từ năm 2025',
    145_000_000,
    'Đầu kỳ',
  ];

  // ==================== THÁNG 01/2026 (25 nghiệp vụ thực) ====================
  // Đợt 1 (DOT-01 - Phú Thành): HD0103 + HD0107 + HD0112 + CK0102
  //   = 11.420.000 + 9.870.000 + 7.200.000 - 850.000 = 27.640.000 (khớp 0đ, span 3 ngày)
  // Đợt 5 phần T01 (An Phát): HD0119 = 15.410.000
  const janTransactions = [
    ['03/01/2026', 'HD0101', 'Công ty TNHH Xây dựng Hoàng Long', 'Xuất bán thép hình chữ H công trình cầu vượt', 16_321_100, 'VAT 8%'],
    ['05/01/2026', 'HD0102', 'Công ty TNHH Cơ khí Minh Quân', 'Cung cấp que hàn chịu lực & đá mài lô MQ-01', 8_451_100, 'VAT 10%'],
    ['08/01/2026', 'HD0103', 'Công ty TNHH Điện máy Phú Thành', 'Xuất bán máy biến áp khô & tủ tụ bù - Giao đợt 1', 11_420_000, 'Đợt 1 Phú Thành'],
    ['09/01/2026', 'HD0104', 'Công ty CP Xây lắp Hải Âu', 'Bán cáp ngầm trung thế CXV/SE 24kV lô HA-01', 19_121_100, 'Đã ký biên bản'],
    ['09/01/2026', 'HD0105', 'Công ty CP Nội thất Đạt Thành', 'Giao ván gỗ MDF chống ẩm phủ Melamine lô ĐT-01', 54_211_100, 'Hợp đồng năm'],
    ['10/01/2026', 'HD0106', 'Công ty TNHH Thương mại An Phát', 'Xuất bán ống thép mạ kẽm Hòa Phát phi 60', 49_101_100, 'Giao kho An Phát'],
    ['10/01/2026', 'HD0107', 'Công ty TNHH Điện máy Phú Thành', 'Xuất bán aptomat công nghiệp MCCB 3P - Giao đợt 2', 9_870_000, 'Đợt 1 Phú Thành'],
    ['11/01/2026', 'HD0108', 'Công ty TNHH Xây dựng Hoàng Long', 'Bổ sung bu lông neo cường độ cao M24', 6_711_100, 'VAT 8%'],
    ['11/01/2026', 'HD0109', 'Công ty TNHH Cơ khí Minh Quân', 'Giao thép tấm cán nóng SS400 dày 12mm', 24_531_100, 'Đã nghiệm thu'],
    ['11/01/2026', 'HD0110', 'Cửa hàng Điện nước Vạn Thịnh', 'Bán sỉ ống nhựa uPVC Bình Minh & phụ kiện ren', 7_831_100, 'Công nợ 30 ngày'],
    ['11/01/2026', 'HD0111', 'Công ty CP Xây lắp Hải Âu', 'Cung cấp thang máng cáp sơn tĩnh điện 300x100', 13_641_100, 'VAT 8%'],
    ['11/01/2026', 'HD0112', 'Công ty TNHH Điện máy Phú Thành', 'Xuất bán đồng hồ đo đa năng & biến dòng - Giao đợt 3', 7_200_000, 'Đợt 1 Phú Thành'],
    ['11/01/2026', 'CK0102', 'Công ty TNHH Điện máy Phú Thành', 'Chiết khấu thương mại đơn hàng thiết bị điện đầu T01', -850_000, 'Giảm trừ Đợt 1'],
    ['14/01/2026', 'TL0101', 'Cửa hàng Điện nước Vạn Thịnh', 'Hàng bán bị trả lại do nứt vỡ đầu nối ống khi bốc dỡ', -1_238_900, 'Biên bản trả hàng'],
    ['16/01/2026', 'HD0113', 'Công ty TNHH Xây dựng Hoàng Long', 'Cung cấp sơn chống rỉ Epoxy hai thành phần', 9_241_100, 'VAT 8%'],
    ['18/01/2026', 'CK0101', 'Công ty TNHH Xây dựng Hoàng Long', 'Chiết khấu thanh toán đúng hạn kỳ trước Hoàng Long', -638_900, 'Phiếu kế toán'],
    ['20/01/2026', 'HD0114', 'Công ty TNHH Điện máy Phú Thành', 'Giao bổ sung quạt công nghiệp thông gió nhà xưởng', 31_411_100, 'Chờ đối soát Q2'],
    ['21/01/2026', 'HD0115', 'Cửa hàng Điện nước Vạn Thịnh', 'Xuất bán đèn pha LED 150W chiếu sáng kho bãi', 5_921_100, 'VAT 8%'],
    ['22/01/2026', 'HD0116', 'Công ty TNHH Cơ khí Minh Quân', 'Cung cấp khí Argon tinh khiết & dây hàn CO2', 11_151_100, 'VAT 10%'],
    ['24/01/2026', 'HD0119', 'Công ty TNHH Thương mại An Phát', 'Xuất bán van cổng mặt bích gang cầu DN100 - Lô AP-01', 15_410_000, 'Đợt 5 An Phát'],
    ['25/01/2026', 'HD0120', 'Công ty CP Nội thất Đạt Thành', 'Cung cấp ray trượt giảm chấn & bản lề inox 304', 47_321_100, 'Lô lớn T01'],
    ['27/01/2026', 'HD0121', 'Công ty CP Xây lắp Hải Âu', 'Xuất bán cọc tiếp địa đồng vàng D16 dài 2.4m', 8_161_100, 'VAT 8%'],
    ['28/01/2026', 'HD0122', 'Cửa hàng Điện nước Vạn Thịnh', 'Giao máy bơm chìm nước thải 1.5kW', 12_471_100, 'Đã bàn giao'],
    ['30/01/2026', 'HD0123', 'Công ty TNHH Xây dựng Hoàng Long', 'Xuất bán lưới thép hàn D6 đổ sàn nhà kho', 17_881_100, 'VAT 8%'],
    ['31/01/2026', 'HD0124', 'Công ty TNHH Cơ khí Minh Quân', 'Phụ phí gia công cắt chấn CNC theo bản vẽ T01', 4_911_100, 'Nghiệm thu T01'],
  ];
  const janTotal = janTransactions.reduce((sum, r) => sum + Number(r[4]), 0);
  const janSummaryRow = ['', '', '', 'Cộng phát sinh tháng 01/2026', janTotal, 'Khóa sổ T01'];

  // ==================== THÁNG 02/2026 (25 nghiệp vụ thực) ====================
  // Đợt 2 (DOT-02 - Hải Âu): HD0204 + HD0208 + HD0211 + TL0201
  //   = 10.510.000 + 9.480.000 + 8.900.000 - 1.250.000 = 27.640.000 (CÙNG SỐ TIỀN Y HỆT ĐỢT 1, span 6 ngày)
  // Đợt 3 (DOT-03 - Đạt Thành): HD0214 + HD0218 + HD0222
  //   = 14.320.280 + 9.680.200 + 8.500.000 = 32.500.480 (khách CK chẵn 32.500.000 -> diff = +480đ)
  // Bẫy nhiễu cho Đợt 3: Hoàng Long có 2 hóa đơn HD0215 (20.000.000) + HD0216 (12.500.000) = 32.500.000 (0đ)
  //   để chứng minh lọc keyword "Đạt Thành" là bắt buộc để loại nhiễu khách hàng khác!
  // Đợt 5 phần T02 (An Phát): HD0224 = 12.830.000
  const febTransactions = [
    ['02/02/2026', 'HD0201', 'Công ty TNHH Xây dựng Hoàng Long', 'Xuất bán cốp pha phủ phim 18mm công trình mới', 21_341_100, 'VAT 8%'],
    ['04/02/2026', 'HD0202', 'Cửa hàng Điện nước Vạn Thịnh', 'Giao ống luồn dây điện chống cháy SP phi 20-25', 6_531_100, 'VAT 8%'],
    ['06/02/2026', 'HD0203', 'Công ty TNHH Điện máy Phú Thành', 'Cung cấp bộ lưu điện công nghiệp UPS 10kVA', 38_621_100, 'Hẹn TT quý 2'],
    ['10/02/2026', 'HD0204', 'Công ty CP Xây lắp Hải Âu', 'Xuất bán tủ điều khiển chiếu sáng ngoài trời - Lô HA-02', 10_510_000, 'Đợt 2 Hải Âu'],
    ['11/02/2026', 'HD0205', 'Công ty TNHH Cơ khí Minh Quân', 'Giao thép tròn đặc S45C gia công trục máy', 15_741_100, 'VAT 10%'],
    ['12/02/2026', 'HD0206', 'Công ty TNHH Thương mại An Phát', 'Cung cấp khớp nối mềm cao su chống rung DN125', 52_111_100, 'Hợp đồng lớn'],
    ['12/02/2026', 'HD0208', 'Công ty CP Xây lắp Hải Âu', 'Xuất bán cáp điều khiển chống nhiễu Altek Kabel 8x1.5', 9_480_000, 'Đợt 2 Hải Âu'],
    ['13/02/2026', 'CK0201', 'Cửa hàng Điện nước Vạn Thịnh', 'Chiết khấu sản lượng đại lý tháng 01 Vạn Thịnh', -738_900, 'Trừ công nợ'],
    ['14/02/2026', 'HD0211', 'Công ty CP Xây lắp Hải Âu', 'Giao sứ đỡ thanh cái & phụ kiện đấu nối trạm biến áp', 8_900_000, 'Đợt 2 Hải Âu'],
    ['16/02/2026', 'TL0201', 'Công ty CP Xây lắp Hải Âu', 'Giảm trừ hàng trả lại do thừa định mức vật tư lô HA-02', -1_250_000, 'Giảm trừ Đợt 2'],
    ['17/02/2026', 'HD0212', 'Công ty TNHH Xây dựng Hoàng Long', 'Xuất bán phụ gia hóa dẻo bê tông Sika Plastiment', 11_921_100, 'VAT 8%'],
    ['18/02/2026', 'HD0214', 'Công ty CP Nội thất Đạt Thành', 'Xuất bán tấm Laminate vân gỗ sồi đơn ĐT-02 (lẻ VAT 8%)', 14_320_280, 'Đợt 3 Đạt Thành'],
    ['19/02/2026', 'HD0215', 'Công ty TNHH Xây dựng Hoàng Long', 'Giao giàn giáo khung mạ kẽm đợt 1 (nhiễu 32.5tr)', 20_000_000, 'Nhiễu khác KH'],
    ['19/02/2026', 'HD0216', 'Công ty TNHH Xây dựng Hoàng Long', 'Giao kích tăng & mâm giàn giáo đợt 2 (nhiễu 32.5tr)', 12_500_000, 'Nhiễu khác KH'],
    ['20/02/2026', 'HD0217', 'Cửa hàng Điện nước Vạn Thịnh', 'Xuất bán bồn nước inox Đại Thành 2000L ngang', 9_841_100, 'VAT 8%'],
    ['21/02/2026', 'HD0218', 'Công ty CP Nội thất Đạt Thành', 'Xuất bán keo dán cạnh nhiệt Jowat & chỉ nhựa PVC (lẻ VAT)', 9_680_200, 'Đợt 3 Đạt Thành'],
    ['22/02/2026', 'HD0219', 'Công ty TNHH Cơ khí Minh Quân', 'Cung cấp vòng bi công nghiệp SKF 6312 chính hãng', 14_251_100, 'VAT 10%'],
    ['23/02/2026', 'TL0202', 'Công ty TNHH Cơ khí Minh Quân', 'Trả lại 02 hộp mũi khoan từ bị sai quy cách', -938_900, 'HĐ điều chỉnh'],
    ['24/02/2026', 'HD0222', 'Công ty CP Nội thất Đạt Thành', 'Phí vận chuyển container & bốc xếp giao xưởng Đạt Thành', 8_500_000, 'Đợt 3 Đạt Thành'],
    ['25/02/2026', 'HD0223', 'Công ty CP Nội thất Đạt Thành', 'Đơn hàng ván HDF siêu chống ẩm lô lớn cuối T02', 61_401_100, 'Chưa đến hạn'],
    ['26/02/2026', 'HD0224', 'Công ty TNHH Thương mại An Phát', 'Xuất bán đồng hồ đo lưu lượng nước điện từ DN80 - Lô AP-02', 12_830_000, 'Đợt 5 An Phát'],
    ['27/02/2026', 'HD0225', 'Công ty CP Xây lắp Hải Âu', 'Cung cấp hộp nối cáp ngầm đổ keo 3M trung thế', 22_461_100, 'VAT 8%'],
    ['27/02/2026', 'HD0226', 'Cửa hàng Điện nước Vạn Thịnh', 'Giao vòi sen tắm nóng lạnh & chậu rửa Inax', 10_371_100, 'VAT 8%'],
    ['28/02/2026', 'HD0227', 'Công ty TNHH Điện máy Phú Thành', 'Bán biến tần điều khiển động cơ Schneider 15kW', 34_181_100, 'Chờ TT tháng 04'],
    ['28/02/2026', 'HD0228', 'Công ty TNHH Xây dựng Hoàng Long', 'Cung cấp bạt chống thấm HDPE lót hồ xử lý nước', 18_591_100, 'Khóa sổ T02'],
  ];
  const febTotal = febTransactions.reduce((sum, r) => sum + Number(r[4]), 0);
  const febSummaryRow = ['', '', '', 'Cộng phát sinh tháng 02/2026', febTotal, 'Khóa sổ T02'];

  // ==================== THÁNG 03/2026 (25 nghiệp vụ thực) ====================
  // Đợt 4 (DOT-04 - Minh Quân, target = 41.300.000):
  //   - Nhiễu không ghim: HD0302 (21.000.000) + HD0306 (20.300.000) = 41.300.000 (2 dòng, không chứa HD0304)
  //   - Hóa đơn gốc kế toán nhớ chắc chắn: HD0304 = 18.300.000 (pinnedIds)
  //   - Hóa đơn đang tranh chấp chất lượng: HD0308 = 23.000.000 (nếu không loại trừ thì HD0304 + HD0308 = 41.300.000 chỉ 2 dòng!)
  //   - Tổ hợp thực tế đúng khi ghim HD0304 + loại trừ HD0308:
  //       HD0304 (18.300.000) + HD0311 (13.450.000) + HD0315 (10.250.000) + CK0303 (-700.000) = 41.300.000
  // Đợt 5 phần T03 (An Phát, gom xuyên 3 tháng T01-T03, target = 37.800.000):
  //   HD0119 (15.410.000) + HD0224 (12.830.000) + HD0319 (11.180.000) + TL0302 (-1.620.000) = 37.800.000
  const marTransactions = [
    ['01/03/2026', 'HD0301', 'Cửa hàng Điện nước Vạn Thịnh', 'Xuất bán dây điện dân dụng Cadivi 2x2.5', 7_411_100, 'VAT 8%'],
    ['03/03/2026', 'HD0302', 'Công ty TNHH Cơ khí Minh Quân', 'Giao máy hàn MIG 350A nguyên chiếc (chưa đến hạn TT)', 21_000_000, 'Nhiễu không ghim'],
    ['04/03/2026', 'HD0303', 'Công ty CP Xây lắp Hải Âu', 'Cung cấp chống sét van trung thế Cooper 24kV', 16_821_100, 'VAT 8%'],
    ['05/03/2026', 'HD0304', 'Công ty TNHH Cơ khí Minh Quân', 'Xuất bán thép tấm chịu mài mòn Hardox 400 (HĐ gốc Đợt 4)', 18_300_000, 'Ghim Đợt 4'],
    ['06/03/2026', 'HD0305', 'Công ty TNHH Xây dựng Hoàng Long', 'Giao xi măng PCB40 đóng bao công trình nhà điều hành', 23_151_100, 'VAT 8%'],
    ['07/03/2026', 'HD0306', 'Công ty TNHH Cơ khí Minh Quân', 'Giao bàn cắt Plasma CNC phụ trợ (chưa đến hạn TT)', 20_300_000, 'Nhiễu không ghim'],
    ['08/03/2026', 'HD0307', 'Công ty CP Nội thất Đạt Thành', 'Xuất bán ốc cam liên kết & chốt gỗ lắp ráp tủ', 53_621_100, 'Lô T03'],
    ['09/03/2026', 'HD0308', 'Công ty TNHH Cơ khí Minh Quân', 'Lô thép hình U200 đang lập biên bản tranh chấp độ phẳng', 23_000_000, 'Đang tranh chấp'],
    ['10/03/2026', 'HD0309', 'Công ty TNHH Điện máy Phú Thành', 'Cung cấp rơ le bảo vệ mất pha & quá dòng Mikro', 29_731_100, 'VAT 8%'],
    ['11/03/2026', 'HD0310', 'Cửa hàng Điện nước Vạn Thịnh', 'Giao bình nước nóng gián tiếp Ariston 30L', 8_641_100, 'VAT 8%'],
    ['12/03/2026', 'HD0311', 'Công ty TNHH Cơ khí Minh Quân', 'Xuất bán lưỡi cưa vòng lưỡng kim cắt thép hợp kim', 13_450_000, 'Đợt 4 Minh Quân'],
    ['14/03/2026', 'TL0301', 'Công ty TNHH Xây dựng Hoàng Long', 'Hàng trả lại do dư khối lượng thép đai phi 8', -1_538_900, 'Phiếu nhập trả'],
    ['15/03/2026', 'HD0314', 'Công ty CP Xây lắp Hải Âu', 'Xuất bán đầu cáp co nhiệt trung thế trong nhà 24kV', 14_961_100, 'VAT 8%'],
    ['16/03/2026', 'HD0315', 'Công ty TNHH Cơ khí Minh Quân', 'Cung cấp dầu cắt gọt kim loại pha nước thùng 200L', 10_250_000, 'Đợt 4 Minh Quân'],
    ['17/03/2026', 'CK0303', 'Công ty TNHH Cơ khí Minh Quân', 'Chiết khấu thanh toán lô vật tư cơ khí Minh Quân T03', -700_000, 'Giảm trừ Đợt 4'],
    ['19/03/2026', 'HD0317', 'Công ty TNHH Thương mại An Phát', 'Giao lô van bướm tay gạt inox vi sinh dự án mới', 57_811_100, 'Hẹn TT tháng 05'],
    ['20/03/2026', 'HD0318', 'Công ty TNHH Xây dựng Hoàng Long', 'Cung cấp vữa rót không co ngót SikaGrout 214-11', 12_271_100, 'VAT 8%'],
    ['22/03/2026', 'HD0319', 'Công ty TNHH Thương mại An Phát', 'Xuất bán mặt bích thép rèn JIS 10K & bộ gioăng EPDM - Lô AP-03', 11_180_000, 'Đợt 5 An Phát'],
    ['24/03/2026', 'TL0302', 'Công ty TNHH Thương mại An Phát', 'Giảm trừ hàng bán bị trả lại do trầy lớp sơn mặt bích Lô AP-03', -1_620_000, 'Giảm trừ Đợt 5'],
    ['25/03/2026', 'HD0320', 'Cửa hàng Điện nước Vạn Thịnh', 'Xuất bán tủ điện nhựa âm tường Sino 12 module', 5_481_100, 'VAT 8%'],
    ['26/03/2026', 'CK0304', 'Công ty CP Xây lắp Hải Âu', 'Chiết khấu doanh số quý 1 cho đối tác Hải Âu', -1_138_900, 'Bù trừ cuối quý'],
    ['27/03/2026', 'HD0321', 'Công ty CP Nội thất Đạt Thành', 'Giao sơn PU bóng mờ 2K hoàn thiện mặt bàn gỗ', 46_911_100, 'VAT 8%'],
    ['29/03/2026', 'HD0322', 'Công ty TNHH Điện máy Phú Thành', 'Xuất bán tụ bù dầu Samwha 3P 30kVAr 440V', 33_521_100, 'VAT 8%'],
    ['30/03/2026', 'HD0323', 'Công ty TNHH Cơ khí Minh Quân', 'Giao đá cắt inox siêu mỏng Hải Dương lô cuối quý', 9_391_100, 'Chốt sổ Q1'],
    ['31/03/2026', 'HD0324', 'Công ty TNHH Xây dựng Hoàng Long', 'Xuất bán nilon lót sàn đổ bê tông & dây kẽm buộc', 6_191_100, 'Chốt sổ Q1'],
  ];
  const marTotal = marTransactions.reduce((sum, r) => sum + Number(r[4]), 0);
  const marSummaryRow = ['', '', '', 'Cộng phát sinh tháng 03/2026', marTotal, 'Khóa sổ T03'];

  const totalPhatSinhTrongKy = janTotal + febTotal + marTotal;
  const grandTotalRow = [
    '',
    '',
    '',
    'TỔNG CỘNG PHÁT SINH TRONG KỲ',
    totalPhatSinhTrongKy,
    'Cuối kỳ Q1/2026',
  ];

  const rawGrid = [
    ...bannerRows,
    headerRow,
    openingRow,
    ...janTransactions,
    janSummaryRow,
    ...febTransactions,
    febSummaryRow,
    ...marTransactions,
    marSummaryRow,
    grandTotalRow,
  ];

  return {
    rawGrid,
    totalPhatSinhTrongKy,
    janTotal,
    febTotal,
    marTotal,
  };
}

/**
 * Helper chuyển toàn bộ chuỗi trong bảng 2 chiều sang dạng Unicode Tổ hợp (NFD).
 */
function toNfdGrid(grid) {
  return grid.map((row) =>
    row.map((cell) => (typeof cell === 'string' ? cell.normalize('NFD') : cell))
  );
}

// ============================================================================
// SUITE 1: LUỒNG NGHIỆP VỤ KẾ TOÁN ĐA ĐỢT TRÊN SỔ CÔNG NỢ 80 DÒNG & XUẤT EXCEL
// ============================================================================
describe('1. Luồng Nghiệp Vụ Thực Chiến: Đối Soát Liên Hoàn 5 Đợt Trên Sổ Công Nợ 80 Dòng & Xuất Sổ Đối Soát', () => {
  it('1.1 Tự động nhận diện cấu trúc sổ công nợ 80 dòng qua 3 tháng, phân loại đúng 75 dòng nghiệp vụ & 5 dòng đầu kỳ/cộng tháng/cuối kỳ', () => {
    const { rawGrid, totalPhatSinhTrongKy } = buildRealistic80RowLedger();

    const inspection = inspectRawGrid(rawGrid, 'CongNo_131_Q1');
    assert.equal(inspection.detectedHeaderRowIndex, 4, 'Phải bỏ qua 4 dòng tiêu đề đơn vị và bắt đúng header ở dòng index 4');
    assert.equal(inspection.suggestedMapping.dateColIndex, 0);
    assert.equal(inspection.suggestedMapping.refColIndex, 1);
    assert.equal(inspection.suggestedMapping.descColIndex, 3);
    assert.equal(inspection.suggestedMapping.amountColIndex, 4);
    assert.equal(inspection.suggestedMapping.amountMode, 'single');

    const extracted = extractTransactions(
      rawGrid,
      inspection.detectedHeaderRowIndex,
      inspection.suggestedMapping,
      inspection.columns
    );

    assert.equal(extracted.items.length, 80, 'Sổ công nợ phải trích xuất đúng 80 dòng');
    assert.equal(extracted.summaryRowCount, 5, 'Phải nhận diện đúng 5 dòng tổng hợp (1 đầu kỳ + 3 cộng tháng + 1 tổng cộng cuối kỳ)');
    assert.equal(extracted.minDateIso, '2026-01-03');
    assert.equal(extracted.maxDateIso, '2026-03-31');
    assert.deepEqual(
      extracted.availableMonths.map((m) => ({ key: m.key, count: m.count })),
      [
        { key: '2026-01', count: 25 },
        { key: '2026-02', count: 25 },
        { key: '2026-03', count: 25 },
      ]
    );

    const realItems = extracted.items.filter((it) => !it.isSummaryRow);
    assert.equal(realItems.length, 75);
    const computedPhatSinh = realItems.reduce((acc, it) => acc + it.amount, 0);
    assert.equal(computedPhatSinh, totalPhatSinhTrongKy);
  });

  it('1.2 Đối soát liên hoàn 5 đợt + Hoàn tác (Undo) Đợt 2 + Kiểm chứng đầy đủ Sheet 1, Sheet 2 & chế độ unallocatedOnly khi xuất Excel', () => {
    const { rawGrid, totalPhatSinhTrongKy } = buildRealistic80RowLedger();
    const inspection = inspectRawGrid(rawGrid, 'CongNo_131_Q1');
    const extracted = extractTransactions(
      rawGrid,
      inspection.detectedHeaderRowIndex,
      inspection.suggestedMapping,
      inspection.columns
    );
    const allItems = extracted.items;
    const itemByRef = new Map(
      allItems.filter((it) => it.reference).map((it) => [it.reference, it])
    );

    const reconciledBatches = [];
    const getLockedIds = () =>
      reconciledBatches.flatMap((batch) => batch.solution.itemIds);

    // ------------------------------------------------------------------------
    // ĐỢT 1: Chuyển khoản 27.640.000đ khớp chính xác 0đ có chứa chiết khấu âm (-850.000đ)
    // ------------------------------------------------------------------------
    const resDot1 = findCombinations({
      items: allItems,
      targetAmount: 27_640_000,
      tolerance: 500,
      excludedIds: getLockedIds(),
      maxItems: 5,
    });

    assert.equal(resDot1.error, null);
    assert.ok(resDot1.solutions.length >= 2, 'Lúc chưa khóa phải thấy cả phương án Đợt 1 (Phú Thành) và Đợt 2 (Hải Âu)');

    const solDot1 = resDot1.solutions[0];
    assert.equal(solDot1.isExact, true);
    assert.equal(solDot1.diff, 0);
    assert.equal(solDot1.totalAmount, 27_640_000);
    assert.deepEqual(
      solDot1.items.map((it) => it.reference),
      ['HD0103', 'HD0107', 'HD0112', 'CK0102'],
      'Đợt 1 phải chọn đúng bộ 3 hóa đơn dương + 1 chiết khấu âm của Phú Thành'
    );
    assert.ok(solDot1.items.some((it) => it.amount < 0), 'Đợt 1 bắt buộc chứa khoản giảm trừ âm');

    reconciledBatches.push({
      code: 'DOT-01',
      targetAmount: 27_640_000,
      totalAmount: solDot1.totalAmount,
      diff: solDot1.diff,
      note: 'UNC Phú Thành thanh toán lô thiết bị điện T01',
      solution: solDot1,
    });
    const dot1LockedSet = new Set(solDot1.itemIds);

    // ------------------------------------------------------------------------
    // ĐỢT 2: Khoản chuyển khoản thứ 2 CÙNG SỐ TIỀN Y HỆT ĐỢT 1 (27.640.000đ)
    // Hệ thống phải tìm ra bộ hóa đơn mới của Hải Âu, TUYỆT ĐỐI không dùng lại dòng Đợt 1!
    // ------------------------------------------------------------------------
    const resDot2 = findCombinations({
      items: allItems,
      targetAmount: 27_640_000,
      tolerance: 500,
      excludedIds: getLockedIds(),
      maxItems: 5,
    });

    assert.equal(resDot2.error, null);
    assert.ok(resDot2.solutions.length >= 1, 'Phải tìm thấy tổ hợp mới cho khoản chuyển khoản thứ 2 cùng mệnh giá');

    for (const sol of resDot2.solutions) {
      for (const id of sol.itemIds) {
        assert.equal(
          dot1LockedSet.has(id),
          false,
          `TUYỆT ĐỐI không được tái sử dụng dòng đã khóa ở Đợt 1 (${id})`
        );
      }
    }

    const solDot2 = resDot2.solutions[0];
    assert.equal(solDot2.isExact, true);
    assert.equal(solDot2.diff, 0);
    assert.equal(solDot2.totalAmount, 27_640_000);
    assert.deepEqual(
      solDot2.items.map((it) => it.reference),
      ['HD0204', 'HD0208', 'HD0211', 'TL0201'],
      'Đợt 2 phải tìm ra đúng bộ 3 hóa đơn + 1 hàng trả lại của Hải Âu'
    );

    reconciledBatches.push({
      code: 'DOT-02',
      targetAmount: 27_640_000,
      totalAmount: solDot2.totalAmount,
      diff: solDot2.diff,
      note: 'UNC Hải Âu thanh toán lô tủ & cáp T02',
      solution: solDot2,
    });

    // ------------------------------------------------------------------------
    // ĐỢT 3: Khách chuyển khoản chẵn 32.500.000đ (lệch lẻ +480đ nằm trong ±500đ) + lọc tên KH "Đạt Thành"
    // ------------------------------------------------------------------------
    const resDot3 = findCombinations({
      items: allItems,
      targetAmount: 32_500_000,
      tolerance: 500,
      keyword: 'Đạt Thành',
      excludedIds: getLockedIds(),
      maxItems: 5,
    });

    assert.equal(resDot3.error, null);
    assert.ok(resDot3.solutions.length >= 1, 'Phải tìm thấy tổ hợp lệch +480đ của Đạt Thành');
    const solDot3 = resDot3.solutions[0];
    assert.equal(solDot3.isExact, false);
    assert.equal(solDot3.totalAmount, 32_500_480);
    assert.equal(solDot3.targetAmount, 32_500_000);
    assert.equal(solDot3.diff, 480);
    assert.equal(solDot3.absDiff, 480);
    assert.deepEqual(
      solDot3.items.map((it) => it.reference),
      ['HD0214', 'HD0218', 'HD0222'],
      'Đợt 3 phải lọc bỏ nhiễu Hoàng Long và lấy đúng 3 chứng từ của Đạt Thành'
    );

    reconciledBatches.push({
      code: 'DOT-03',
      targetAmount: 32_500_000,
      totalAmount: solDot3.totalAmount,
      diff: solDot3.diff,
      note: 'Đạt Thành CK chẵn 32.5tr (thiếu lẻ 480đ thuế VAT)',
      solution: solDot3,
    });

    // ------------------------------------------------------------------------
    // ĐỢT 4: Kế toán ghim 1 hóa đơn gốc (HD0304) + loại trừ thủ công 1 hóa đơn tranh chấp (HD0308)
    // ------------------------------------------------------------------------
    const pinnedRootInvoice = itemByRef.get('HD0304');
    const disputedInvoice = itemByRef.get('HD0308');
    assert.ok(pinnedRootInvoice && disputedInvoice);

    // Kiểm chứng trước khi loại trừ HD0308: cặp 2 dòng HD0304 + HD0308 (41.300.000) sẽ đứng đầu
    const resBeforeDisputeExclude = findCombinations({
      items: allItems,
      targetAmount: 41_300_000,
      tolerance: 500,
      pinnedIds: [pinnedRootInvoice.id],
      excludedIds: getLockedIds(),
      maxItems: 5,
    });
    assert.deepEqual(
      resBeforeDisputeExclude.solutions[0].items.map((it) => it.reference),
      ['HD0304', 'HD0308'],
      'Khi chưa loại trừ hóa đơn tranh chấp HD0308, tổ hợp 2 dòng chứa HD0308 đứng đầu'
    );

    // Sau khi loại trừ thủ công HD0308: tìm ra đúng bộ 4 chứng từ thực thanh toán!
    const resDot4 = findCombinations({
      items: allItems,
      targetAmount: 41_300_000,
      tolerance: 500,
      pinnedIds: [pinnedRootInvoice.id],
      excludedIds: [...getLockedIds(), disputedInvoice.id],
      maxItems: 5,
    });

    assert.equal(resDot4.error, null);
    assert.ok(resDot4.solutions.length >= 1);
    for (const sol of resDot4.solutions) {
      assert.ok(
        sol.itemIds.includes(pinnedRootInvoice.id),
        'Mọi nghiệm đều phải chứa hóa đơn đã ghim HD0304'
      );
      assert.equal(
        sol.itemIds.includes(disputedInvoice.id),
        false,
        'Không nghiệm nào được chứa hóa đơn tranh chấp HD0308'
      );
    }

    const solDot4 = resDot4.solutions[0];
    assert.equal(solDot4.isExact, true);
    assert.equal(solDot4.totalAmount, 41_300_000);
    assert.deepEqual(
      solDot4.items.map((it) => it.reference),
      ['HD0304', 'HD0311', 'HD0315', 'CK0303']
    );

    reconciledBatches.push({
      code: 'DOT-04',
      targetAmount: 41_300_000,
      totalAmount: solDot4.totalAmount,
      diff: solDot4.diff,
      note: 'Minh Quân TT đợt T03 (giữ lại HD0308 chờ xử lý tranh chấp)',
      solution: solDot4,
    });

    // ------------------------------------------------------------------------
    // HOÀN TÁC (UNDO) ĐỢT 2: Mở khóa các dòng Đợt 2 -> Tìm lại phải thấy Đợt 2 xuất hiện trở lại!
    // ------------------------------------------------------------------------
    // Kiểm chứng trước khi Undo: cả Đợt 1 và Đợt 2 đều đã khóa nên tìm 27.640.000đ không còn 2 bộ này
    const resWhileDot2Locked = findCombinations({
      items: allItems,
      targetAmount: 27_640_000,
      tolerance: 500,
      excludedIds: getLockedIds(),
      maxItems: 5,
    });
    assert.equal(
      resWhileDot2Locked.solutions.some((s) =>
        s.itemIds.every((id) => solDot2.itemIds.includes(id)) &&
        s.itemIds.length === solDot2.itemIds.length
      ),
      false,
      'Khi Đợt 2 đang khóa thì không được xuất hiện trong kết quả tìm kiếm'
    );

    // Thực hiện Undo Đợt 2
    const dot2Index = reconciledBatches.findIndex((b) => b.code === 'DOT-02');
    assert.equal(dot2Index, 1);
    const [undoneDot2] = reconciledBatches.splice(dot2Index, 1);

    const resAfterUndoDot2 = findCombinations({
      items: allItems,
      targetAmount: 27_640_000,
      tolerance: 500,
      excludedIds: getLockedIds(),
      maxItems: 5,
    });

    assert.equal(resAfterUndoDot2.error, null);
    assert.ok(resAfterUndoDot2.solutions.length >= 1, 'Sau khi Undo Đợt 2, tìm lại phải thấy kết quả');
    assert.deepEqual(
      resAfterUndoDot2.solutions[0].items.map((it) => it.reference),
      ['HD0204', 'HD0208', 'HD0211', 'TL0201'],
      'Các dòng của Đợt 2 phải xuất hiện trở lại ở vị trí nghiệm tối ưu số 1'
    );
    for (const sol of resAfterUndoDot2.solutions) {
      for (const id of sol.itemIds) {
        assert.equal(
          dot1LockedSet.has(id),
          false,
          'Các dòng của Đợt 1 vẫn đang khóa nên tuyệt đối không được xuất hiện lại khi Undo Đợt 2'
        );
      }
    }

    // Khóa xác nhận lại Đợt 2 sau khi kiểm tra Undo thành công
    reconciledBatches.splice(dot2Index, 0, undoneDot2);

    // ------------------------------------------------------------------------
    // ĐỢT 5: Đối soát khoản chuyển khoản thứ 5 xuyên 3 tháng (T01..T03) của An Phát có hàng trả lại âm
    // ------------------------------------------------------------------------
    const resDot5 = findCombinations({
      items: allItems,
      targetAmount: 37_800_000,
      tolerance: 500,
      keyword: 'An Phát',
      excludedIds: getLockedIds(),
      maxItems: 5,
    });

    assert.equal(resDot5.error, null);
    assert.ok(resDot5.solutions.length >= 1);
    const solDot5 = resDot5.solutions[0];
    assert.equal(solDot5.isExact, true);
    assert.equal(solDot5.totalAmount, 37_800_000);
    assert.deepEqual(
      solDot5.items.map((it) => it.reference),
      ['HD0119', 'HD0224', 'HD0319', 'TL0302']
    );

    reconciledBatches.push({
      code: 'DOT-05',
      targetAmount: 37_800_000,
      totalAmount: solDot5.totalAmount,
      diff: solDot5.diff,
      note: 'An Phát tất toán công nợ gộp 3 tháng Q1 (có trừ trả hàng TL0302)',
      solution: solDot5,
    });

    assert.equal(reconciledBatches.length, 5, 'Hoàn tất đủ 5 đợt gom công nợ');

    // ------------------------------------------------------------------------
    // KIỂM TRA XUẤT FILE EXCEL ĐỐI SOÁT (exportReconciledWorkbook) - CHẾ ĐỘ ĐẦY ĐỦ
    // ------------------------------------------------------------------------
    const fullXlsxBytes = exportReconciledWorkbook({
      sheetName: 'CongNo_131_Q1',
      columns: inspection.columns,
      items: allItems,
      reconciledBatches,
      unallocatedOnly: false,
    });

    assert.ok(fullXlsxBytes instanceof Uint8Array);
    assert.ok(fullXlsxBytes.byteLength > 0);

    const wbFull = loadWorkbookFromBuffer(fullXlsxBytes);
    assert.deepEqual(getSheetNames(wbFull), ['Chi_Tiet_Doi_Soat', 'Tong_Hop_Cac_Dot_Gom']);

    // --- Kiểm chứng Sheet 1 (Chi_Tiet_Doi_Soat - Đầy đủ) ---
    const sheet1FullGrid = inspectSheet(wbFull, 'Chi_Tiet_Doi_Soat').rawGrid;
    assert.equal(
      sheet1FullGrid.length,
      1 + 80,
      'Sheet 1 chế độ đầy đủ phải có đúng 1 dòng header + 80 dòng dữ liệu'
    );

    const s1Header = sheet1FullGrid[0];
    assert.deepEqual(s1Header.slice(0, 9), [
      'Dòng Excel gốc',
      'Ngày',
      'Số chứng từ',
      'Diễn giải',
      'Số tiền',
      'Trạng thái đối soát',
      'Mã đợt gom',
      'Mục tiêu đợt',
      'Chênh lệch đợt',
    ]);
    assert.ok(
      s1Header.includes('Khách hàng / Đối tác') && s1Header.includes('Ghi chú'),
      'Sheet 1 phải bảo toàn đầy đủ các cột phụ gốc từ sổ công nợ'
    );

    const expectedBatchByItemId = new Map();
    for (const b of reconciledBatches) {
      for (const it of b.solution.items) {
        expectedBatchByItemId.set(it.id, b);
      }
    }

    let countDaGom = 0;
    let countChuaGom = 0;
    let countTongCong = 0;

    for (let i = 0; i < allItems.length; i++) {
      const origItem = allItems[i];
      const exportedRow = sheet1FullGrid[i + 1];

      assert.equal(exportedRow[0], origItem.excelRowNumber, 'Dòng Excel gốc phải khớp chính xác');
      assert.equal(exportedRow[4], origItem.amount, 'Số tiền từng dòng trên Sheet 1 phải khớp tuyệt đối');

      const expectedBatch = expectedBatchByItemId.get(origItem.id);
      if (expectedBatch) {
        countDaGom++;
        assert.equal(exportedRow[5], 'Đã gom');
        assert.equal(exportedRow[6], expectedBatch.code);
        assert.equal(exportedRow[7], expectedBatch.targetAmount);
        assert.equal(exportedRow[8], expectedBatch.diff);
      } else if (origItem.isSummaryRow) {
        countTongCong++;
        assert.equal(exportedRow[5], 'Dòng tổng cộng');
        assert.equal(exportedRow[6], '');
        assert.equal(exportedRow[7], '');
        assert.equal(exportedRow[8], '');
      } else {
        countChuaGom++;
        assert.equal(exportedRow[5], 'Chưa gom');
        assert.equal(exportedRow[6], '');
        assert.equal(exportedRow[7], '');
        assert.equal(exportedRow[8], '');
      }
    }

    const expectedReconciledItemCount = reconciledBatches.reduce(
      (s, b) => s + b.solution.items.length,
      0
    );
    assert.equal(expectedReconciledItemCount, 19); // 4 + 4 + 3 + 4 + 4
    assert.equal(countDaGom, 19);
    assert.equal(countTongCong, 5);
    assert.equal(countChuaGom, 80 - 19 - 5); // 56 dòng chưa gom

    // --- Kiểm chứng Sheet 2 (Tong_Hop_Cac_Dot_Gom) ---
    const sheet2Grid = inspectSheet(wbFull, 'Tong_Hop_Cac_Dot_Gom').rawGrid;

    const expectedSumTarget = reconciledBatches.reduce((s, b) => s + b.targetAmount, 0);
    const expectedSumActual = reconciledBatches.reduce((s, b) => s + b.totalAmount, 0);
    const expectedSumDiff = reconciledBatches.reduce((s, b) => s + b.diff, 0);

    assert.equal(expectedSumTarget, 27_640_000 + 27_640_000 + 32_500_000 + 41_300_000 + 37_800_000); // 166.880.000
    assert.equal(expectedSumActual, 166_880_480);
    assert.equal(expectedSumDiff, 480);

    // Dòng tóm tắt nhanh ở đầu Sheet 2 (row index 1)
    assert.equal(sheet2Grid[1][1], 5, 'Số đợt đã gom trên banner Sheet 2 phải là 5');
    assert.equal(sheet2Grid[1][3], 19, 'Tổng số dòng đã gom trên banner Sheet 2 phải là 19');
    assert.equal(sheet2Grid[1][5], expectedSumActual, 'Tổng tiền thực gom trên banner Sheet 2 phải khớp tuyệt đối');

    // Bảng I: Từng đợt gom (row 5..9) + Dòng Tổng cộng (row 10)
    for (let idx = 0; idx < reconciledBatches.length; idx++) {
      const b = reconciledBatches[idx];
      const r = sheet2Grid[5 + idx];
      assert.equal(r[0], idx + 1);
      assert.equal(r[1], b.code);
      assert.equal(r[2], b.solution.items.length);
      assert.equal(r[3], b.targetAmount);
      assert.equal(r[4], b.totalAmount);
      assert.equal(r[5], b.diff);
    }

    const totalBatchSummaryRow = sheet2Grid[5 + reconciledBatches.length];
    assert.equal(totalBatchSummaryRow[1], 'Tổng cộng');
    assert.equal(totalBatchSummaryRow[2], 19);
    assert.equal(totalBatchSummaryRow[3], expectedSumTarget);
    assert.equal(totalBatchSummaryRow[4], expectedSumActual);
    assert.equal(totalBatchSummaryRow[5], expectedSumDiff);

    // Bảng II: Chi tiết các dòng thuộc từng đợt gom (19 dòng chi tiết)
    const detailHeaderIdx = sheet2Grid.findIndex((r) => r[0] === 'Mã đợt gom' && r[1] === 'STT trong đợt');
    assert.ok(detailHeaderIdx > 10, 'Phải có bảng II chi tiết từng dòng của các đợt gom');
    const detailRows = sheet2Grid.slice(detailHeaderIdx + 1);
    assert.equal(detailRows.length, 19, 'Bảng II trên Sheet 2 phải liệt kê đúng 19 dòng đã gom');
    const sumDetailAmounts = detailRows.reduce((s, r) => s + Number(r[6]), 0);
    assert.equal(sumDetailAmounts, expectedSumActual, 'Tổng cột Số tiền ở Bảng II phải bằng đúng Tổng tiền thực gom');

    // ------------------------------------------------------------------------
    // KIỂM TRA CHẾ ĐỘ unallocatedOnly: true (CHỈ XUẤT CÔNG NỢ CÒN PHẢI THU)
    // ------------------------------------------------------------------------
    const unallocatedBytes = exportReconciledWorkbook({
      sheetName: 'CongNo_131_Q1',
      columns: inspection.columns,
      items: allItems,
      reconciledBatches,
      unallocatedOnly: true,
    });

    const wbUnalloc = loadWorkbookFromBuffer(unallocatedBytes);
    const sheet1UnallocGrid = inspectSheet(wbUnalloc, 'Chi_Tiet_Doi_Soat').rawGrid;

    assert.equal(
      sheet1UnallocGrid.length,
      1 + 56,
      'Chế độ unallocatedOnly: true chỉ được chứa đúng 1 header + 56 dòng Chưa gom'
    );

    let remainingReceivableSum = 0;
    for (let r = 1; r < sheet1UnallocGrid.length; r++) {
      const row = sheet1UnallocGrid[r];
      assert.equal(
        row[5],
        'Chưa gom',
        'Mọi dòng trong chế độ unallocatedOnly bắt buộc mang trạng thái "Chưa gom" (không chứa Đã gom hay Dòng tổng cộng)'
      );
      remainingReceivableSum += Number(row[4]);
    }

    assert.equal(
      remainingReceivableSum,
      totalPhatSinhTrongKy - expectedSumActual,
      'Tổng tiền các dòng còn phải thu phải khớp tuyệt đối: Tổng phát sinh ban đầu - Tổng thực gom các đợt'
    );
    assert.equal(
      remainingReceivableSum,
      totalPhatSinhTrongKy - (expectedSumTarget + expectedSumDiff),
      'Bảo toàn phương trình kế toán: Còn phải thu = Tổng phát sinh - (Tổng tiền khách chuyển + Tổng chênh lệch)'
    );
  });

  it('1.3 [Adversarial] exportReconciledWorkbook phải tự tính lại đúng totalAmount và diff khi UI truyền totalAmount: null / diff: null (bẫy Number(null) === 0)', () => {
    const items = [
      {
        id: 'row-2',
        excelRowNumber: 2,
        dateIso: '2026-03-05',
        dateDisplay: '05/03/2026',
        timestamp: Date.UTC(2026, 2, 5),
        amount: 10_000_000,
        description: 'Hóa đơn bán hàng 01',
        reference: 'HD001',
        isSummaryRow: false,
        extraCells: {},
        rawCells: ['05/03/2026', 'HD001', 'Hóa đơn bán hàng 01', 10_000_000],
      },
      {
        id: 'row-3',
        excelRowNumber: 3,
        dateIso: '2026-03-06',
        dateDisplay: '06/03/2026',
        timestamp: Date.UTC(2026, 2, 6),
        amount: 5_000_300,
        description: 'Hóa đơn bán hàng 02',
        reference: 'HD002',
        isSummaryRow: false,
        extraCells: {},
        rawCells: ['06/03/2026', 'HD002', 'Hóa đơn bán hàng 02', 5_000_300],
      },
    ];

    // Trong JS: Number(null) === 0 (Number.isFinite(0) === true)!
    // Nếu UI khởi tạo batch với totalAmount: null, diff: null để nhờ hàm xuất tự tính từ items,
    // tổng tiền thực gom phải là 15.000.300 và chênh lệch phải là +300 chứ KHÔNG ĐƯỢC biến thành 0!
    const bytes = exportReconciledWorkbook({
      sheetName: 'TestNullBatchFields',
      items,
      reconciledBatches: [
        {
          code: 'DOT-NULL',
          itemIds: ['row-2', 'row-3'],
          targetAmount: 15_000_000,
          totalAmount: null,
          diff: null,
        },
      ],
    });

    const wb = loadWorkbookFromBuffer(bytes);
    const s2 = inspectSheet(wb, 'Tong_Hop_Cac_Dot_Gom').rawGrid;
    assert.equal(
      s2[1][5],
      15_000_300,
      'Tổng tiền thực gom trên Sheet 2 phải tính đúng 15.000.300đ thay vì bị ép thành 0đ do Number(null) === 0'
    );
    assert.equal(
      s2[5][5],
      300,
      'Chênh lệch đợt trên Sheet 2 phải tính đúng +300đ thay vì bị ép thành 0đ do Number(null) === 0'
    );
  });

  it('1.4 [Adversarial] exportReconciledWorkbook phải đảm bảo toàn vẹn đối chiếu chéo Sheet 1 === Sheet 2 (khử trùng lặp khi itemId bị truyền trùng trong/giữa các đợt)', () => {
    const items = [
      {
        id: 'row-2',
        excelRowNumber: 2,
        dateIso: '2026-03-05',
        dateDisplay: '05/03/2026',
        timestamp: Date.UTC(2026, 2, 5),
        amount: 10_000_000,
        description: 'Hóa đơn 01',
        reference: 'HD001',
        isSummaryRow: false,
        extraCells: {},
        rawCells: [],
      },
      {
        id: 'row-3',
        excelRowNumber: 3,
        dateIso: '2026-03-06',
        dateDisplay: '06/03/2026',
        timestamp: Date.UTC(2026, 2, 6),
        amount: 5_000_000,
        description: 'Hóa đơn 02',
        reference: 'HD002',
        isSummaryRow: false,
        extraCells: {},
        rawCells: [],
      },
    ];

    // Nếu dữ liệu đợt gom truyền vào vô tình lặp lại cùng 'row-2' (trong cùng 1 đợt hoặc giữa 2 đợt),
    // Sheet 1 và Sheet 2 bắt buộc phải thống nhất số lượng dòng "Đã gom" và tổng tiền thực gom!
    const bytes = exportReconciledWorkbook({
      sheetName: 'DupItemIntegrity',
      items,
      reconciledBatches: [
        {
          code: 'DOT-DUP',
          itemIds: ['row-2', 'row-2', 'row-3'],
          targetAmount: 15_000_000,
        },
      ],
    });

    const wb = loadWorkbookFromBuffer(bytes);
    const s1 = inspectSheet(wb, 'Chi_Tiet_Doi_Soat').rawGrid;
    const s2 = inspectSheet(wb, 'Tong_Hop_Cac_Dot_Gom').rawGrid;

    const s1ReconciledRows = s1.slice(1).filter((r) => r[5] === 'Đã gom');
    const s1ReconciledSum = s1ReconciledRows.reduce((acc, r) => acc + Number(r[4]), 0);

    const s2TotalItemsCount = s2[1][3];
    const s2TotalActualSum = s2[1][5];

    assert.equal(
      s2TotalItemsCount,
      s1ReconciledRows.length,
      'Tổng số dòng đã gom trên Sheet 2 bắt buộc khớp đúng số dòng "Đã gom" trên Sheet 1'
    );
    assert.equal(
      s2TotalActualSum,
      s1ReconciledSum,
      'Tổng tiền thực gom trên Sheet 2 bắt buộc khớp đúng tổng tiền các dòng "Đã gom" trên Sheet 1 (không được nhân đôi hóa đơn)'
    );
  });
});

// ============================================================================
// SUITE 2: BẪY BẢNG MÃ TIẾNG VIỆT DỰNG SẴN (NFC) VS TỔ HỢP (NFD - UNIKEY TỔ HỢP)
// ============================================================================
describe('2. Bẫy Bảng Mã Tiếng Việt Dựng Sẵn (NFC) vs Tổ Hợp (NFD - Unikey Tổ Hợp / Copy PDF)', () => {
  it('2.1 inspectRawGrid nhận diện đúng dòng tiêu đề và các cột Ngày chứng từ, Số tiền, Diễn giải khi toàn bộ bảng ở dạng Unicode Tổ hợp (NFD)', () => {
    const nfcGrid = [
      ['CÔNG TY TNHH ĐẠT THÀNH', '', '', ''],
      ['BẢNG KÊ CÔNG NỢ PHẢI THU THÁNG 03/2026', '', '', ''],
      ['Ngày chứng từ', 'Số chứng từ', 'Diễn giải công nợ', 'Số tiền'],
      ['05/03/2026', 'HD01', 'Bán hàng cho Công ty Đạt Thành đợt 1', 12_500_000],
      ['10/03/2026', 'HD02', 'Chiết khấu thương mại Đạt Thành', -500_000],
      ['', '', 'TỔNG CỘNG PHÁT SINH', 12_000_000],
    ];

    const nfdGrid = toNfdGrid(nfcGrid);
    // Đảm bảo dữ liệu test thực sự ở dạng NFD (độ dài code unit lớn hơn NFC)
    assert.notEqual(nfdGrid[2][0], nfcGrid[2][0]);
    assert.notEqual(nfdGrid[2][2], nfcGrid[2][2]);
    assert.notEqual(nfdGrid[2][3], nfcGrid[2][3]);

    const inspection = inspectRawGrid(nfdGrid, 'CongNo_NFD');
    assert.equal(inspection.detectedHeaderRowIndex, 2);
    assert.equal(inspection.suggestedMapping.dateColIndex, 0);
    assert.equal(inspection.suggestedMapping.refColIndex, 1);
    assert.equal(inspection.suggestedMapping.descColIndex, 2);
    assert.equal(inspection.suggestedMapping.amountColIndex, 3);
    assert.equal(inspection.suggestedMapping.amountMode, 'single');
  });

  it('2.2 [Adversarial] inspectRawGrid phải nhận diện đúng cột Nợ / Có trên Sổ chi tiết công nợ Thông tư 200 khi tiêu đề ở dạng Unicode Tổ hợp (NFD)', () => {
    // Mẫu Sổ chi tiết công nợ TK 131 cực kỳ phổ biến ở Việt Nam (MISA / Fast / Bravo)
    const nfcDebitCreditGrid = [
      ['CÔNG TY TNHH THƯƠNG MẠI ĐẠT THÀNH', '', '', '', ''],
      ['SỔ CHI TIẾT TÀI KHOẢN 131 - PHẢI THU KHÁCH HÀNG', '', '', '', ''],
      ['Ngày chứng từ', 'Số hiệu CT', 'Diễn giải', 'Nợ', 'Có'],
      ['05/03/2026', 'HD001', 'Xuất bán hàng hóa lô 01', 15_000_000, 0],
      ['08/03/2026', 'HD002', 'Xuất bán hàng hóa lô 02', 8_500_000, 0],
      ['12/03/2026', 'CK001', 'Chiết khấu thương mại', 0, 500_000],
    ];

    // Ở dạng NFC: hệ thống nhận diện hoàn hảo debit_minus_credit (Nợ col 3, Có col 4)
    const nfcInspection = inspectRawGrid(nfcDebitCreditGrid, 'TK131_NFC');
    assert.equal(nfcInspection.detectedHeaderRowIndex, 2);
    assert.equal(nfcInspection.suggestedMapping.amountMode, 'debit_minus_credit');
    assert.equal(nfcInspection.suggestedMapping.debitColIndex, 3);
    assert.equal(nfcInspection.suggestedMapping.creditColIndex, 4);

    // Ở dạng NFD (Unikey Unicode tổ hợp): kết quả nhận diện cột PHẢI GIỐNG HỆT NFC!
    const nfdDebitCreditGrid = toNfdGrid(nfcDebitCreditGrid);
    const nfdInspection = inspectRawGrid(nfdDebitCreditGrid, 'TK131_NFD');
    assert.equal(nfdInspection.detectedHeaderRowIndex, 2);
    assert.equal(
      nfdInspection.suggestedMapping.amountMode,
      'debit_minus_credit',
      'Khi header "Nợ" / "Có" ở dạng NFD, hệ thống vẫn phải nhận diện chế độ debit_minus_credit'
    );
    assert.equal(
      nfdInspection.suggestedMapping.debitColIndex,
      3,
      'Không được làm rơi mất cột Nợ (hóa đơn bán hàng) khi chữ "Nợ" ở dạng Unicode Tổ hợp NFD'
    );
    assert.equal(
      nfdInspection.suggestedMapping.creditColIndex,
      4,
      'Không được làm rơi mất cột Có khi chữ "Có" ở dạng Unicode Tổ hợp NFD'
    );
  });

  it('2.3 extractTransactions nhận diện đúng isSummaryRow = true cho "TỔNG CỘNG PHÁT SINH", "SỐ DƯ ĐẦU KỲ", "CỘNG LŨY KẾ" ở dạng NFD', () => {
    const nfdGrid = toNfdGrid([
      ['Ngày chứng từ', 'Số chứng từ', 'Diễn giải', 'Số tiền'],
      ['', 'DK', 'SỐ DƯ ĐẦU KỲ', 50_000_000],
      ['05/03/2026', 'HD01', 'Giao hàng Công ty Đạt Thành', 18_000_000],
      ['09/03/2026', 'CK01', 'Giảm giá hàng bán Đạt Thành', -1_000_000],
      ['', '', 'CỘNG PHÁT SINH TRONG KỲ', 17_000_000],
      ['', '', 'TỔNG CỘNG PHÁT SINH', 17_000_000],
    ]);

    const inspection = inspectRawGrid(nfdGrid, 'Summary_NFD');
    const extracted = extractTransactions(
      nfdGrid,
      inspection.detectedHeaderRowIndex,
      inspection.suggestedMapping,
      inspection.columns
    );

    assert.equal(extracted.items.length, 5);
    assert.equal(extracted.summaryRowCount, 3);
    assert.equal(extracted.items[0].isSummaryRow, true, 'SỐ DƯ ĐẦU KỲ (NFD, có chữ Đ) phải là summary');
    assert.equal(extracted.items[1].isSummaryRow, false);
    assert.equal(extracted.items[2].isSummaryRow, false);
    assert.equal(extracted.items[3].isSummaryRow, true, 'CỘNG PHÁT SINH TRONG KỲ (NFD) phải là summary');
    assert.equal(extracted.items[4].isSummaryRow, true, 'TỔNG CỘNG PHÁT SINH (NFD) phải là summary');
  });

  it('2.4 [Adversarial] extractTransactions phải đọc đúng ô Ngày ("Ngày 05/03/2026") và Số tiền kèm chữ ("12.500.000 đồng") khi chuỗi ở dạng NFD', () => {
    const nfdGrid = toNfdGrid([
      ['Ngày chứng từ', 'Số chứng từ', 'Diễn giải', 'Số tiền'],
      ['Ngày 05/03/2026', 'HD01', 'Thu tiền đợt 1 Công ty Đạt Thành', '12.500.000 đồng'],
      ['Ngày 10 tháng 03 năm 2026', 'HD02', 'Thu tiền đợt 2 Công ty Đạt Thành', '7.500.000 VNĐ'],
    ]);

    const inspection = inspectRawGrid(nfdGrid, 'Cells_NFD');
    const extracted = extractTransactions(
      nfdGrid,
      inspection.detectedHeaderRowIndex,
      inspection.suggestedMapping,
      inspection.columns
    );

    assert.equal(
      extracted.items.length,
      2,
      'Không được làm rơi dòng giao dịch có hậu tố "đồng" ở dạng Unicode Tổ hợp NFD'
    );
    assert.equal(
      extracted.items[0].dateIso,
      '2026-03-05',
      'Phải parse được "Ngày 05/03/2026" ở dạng NFD'
    );
    assert.equal(extracted.items[0].amount, 12_500_000);
    assert.equal(
      extracted.items[1].dateIso,
      '2026-03-10',
      'Phải parse được "Ngày 10 tháng 03 năm 2026" ở dạng NFD'
    );
    assert.equal(extracted.items[1].amount, 7_500_000);
  });

  it('2.5 findCombinations lọc keyword khớp 100% bất chấp bảng NFD vs từ khóa NFC, NFD, không dấu, chữ hoa/thường và chữ Đ/đ kết hợp dấu', () => {
    // Bao gồm:
    // - NFD tiêu chuẩn ("Công ty TNHH Đạt Thành".normalize('NFD'))
    // - Chuỗi lai Unikey tổ hợp (nguyên âm dựng sẵn + dấu thanh tổ hợp rời: "Đa\u0323t Tha\u0300nh")
    // - Chữ D + dấu gạch ngang tổ hợp (U+0335 COMBINING SHORT STROKE OVERLAY: "D\u0335a\u0323t")
    const grid = [
      ['Ngày chứng từ', 'Số chứng từ', 'Khách hàng', 'Diễn giải', 'Số tiền'],
      [
        '05/03/2026',
        'HD-NFD-1',
        'Công ty TNHH Đạt Thành'.normalize('NFD'),
        'Giao thiết bị điện đường Điện Biên Phủ - Đợt 1'.normalize('NFD'),
        8_200_000,
      ],
      [
        '07/03/2026',
        'HD-NFD-2',
        'Công ty TNHH Đa\u0323t Tha\u0300nh', // Unikey tổ hợp bán phần
        'Phụ kiện tủ điện Đồng Nai - Đợt 2'.normalize('NFD'),
        6_800_000,
      ],
      [
        '08/03/2026',
        'CK-NFD-3',
        'Công ty TNHH D\u0335a\u0323t Tha\u0300nh', // D + combining stroke + combining dot below
        'Chiết khấu thanh toán sớm Đạt Thành'.normalize('NFD'),
        -500_000,
      ],
      [
        '09/03/2026',
        'HD-OTHER',
        'Công ty TNHH Đại Thắng'.normalize('NFD'),
        'Hóa đơn nhiễu cùng tổng tiền 14.500.000đ của khách khác',
        14_500_000,
      ],
    ];

    const inspection = inspectRawGrid(grid, 'Keyword_NFD_NFC');
    const { items } = extractTransactions(
      grid,
      inspection.detectedHeaderRowIndex,
      inspection.suggestedMapping,
      inspection.columns
    );

    const keywordVariants = [
      'Đạt Thành'.normalize('NFC'),        // Người dùng gõ NFC dựng sẵn
      'Đạt Thành'.normalize('NFD'),        // Người dùng paste NFD tổ hợp
      'đạt thành',                         // Chữ thường có dấu NFC
      'ĐẠT THÀNH'.normalize('NFD'),        // Chữ IN HOA tổ hợp NFD (Đ + dấu nặng rời)
      'dat thanh',                         // Gõ nhanh không dấu
      'DAT THANH',                         // Gõ IN HOA không dấu
      'D\u0335a\u0323t Tha\u0300nh',       // D + combining stroke overlay
    ];

    for (const kw of keywordVariants) {
      assert.equal(normalizeVietnameseSearch(kw), 'dat thanh');

      const res = findCombinations({
        items,
        targetAmount: 14_500_000,
        tolerance: 0,
        keyword: kw,
      });

      assert.equal(res.error, null);
      assert.equal(
        res.solutions.length,
        1,
        `Phải tìm thấy đúng 1 nghiệm cho biến thể keyword: ${JSON.stringify(kw)}`
      );
      assert.deepEqual(
        res.solutions[0].items.map((it) => it.reference),
        ['HD-NFD-1', 'HD-NFD-2', 'CK-NFD-3'],
        'Phải gom đúng 3 dòng của Đạt Thành (bỏ qua dòng nhiễu Đại Thắng)'
      );
    }

    // Kiểm tra ngược lại: Dữ liệu bảng là NFC, người dùng paste keyword ở dạng NFD
    const nfcItems = items.map((it) => ({
      ...it,
      description: it.description.normalize('NFC'),
      extraCells: Object.fromEntries(
        Object.entries(it.extraCells).map(([k, v]) => [k, String(v).normalize('NFC')])
      ),
    }));
    const reverseRes = findCombinations({
      items: nfcItems,
      targetAmount: 8_200_000,
      tolerance: 0,
      keyword: 'đường Điện Biên Phủ'.normalize('NFD'),
    });
    assert.equal(reverseRes.solutions.length, 1);
    assert.equal(reverseRes.solutions[0].items[0].reference, 'HD-NFD-1');
  });
});

// ============================================================================
// SUITE 3: KIỂM TRA CÁC BIÊN LỌC NGÀY (dateFromIso / dateToIso) & DÒNG TRỐNG NGÀY
// ============================================================================
describe('3. Kiểm Tra Các Biên Lọc Ngày (dateFromIso / dateToIso) & Dòng Trống Ngày (dateIso === null)', () => {
  const rawGrid = [
    ['Ngày chứng từ', 'Số chứng từ', 'Khách hàng', 'Diễn giải', 'Số tiền'],
    ['28/02/2024', 'HD-FEB28', 'Công ty A', 'Giao hàng ngày 28/02 năm nhuận', 5_000_000],
    ['29/02/2024', 'HD-FEB29-1', 'Công ty A', 'Giao hàng sáng ngày nhuận 29/02/2024', 7_200_000],
    ['29/02/2024', 'HD-FEB29-2', 'Công ty A', 'Giao hàng chiều ngày nhuận 29/02/2024', 4_800_000],
    ['29/02/2024', 'CK-FEB29', 'Công ty A', 'Chiết khấu trực tiếp trong ngày 29/02/2024', -500_000],
    ['01/03/2024', 'HD-MAR01', 'Công ty A', 'Giao hàng ngày 01/03 ngay liền sau', 6_500_000],
    ['', 'HD-NODATE-1', 'Công ty A', 'Phụ phí bốc xếp bổ sung quên ghi ngày chứng từ', 3_500_000],
    ['', 'CK-NODATE-2', 'Công ty A', 'Giảm trừ hàng móp méo phiếu viết tay chưa ghi ngày', -400_000],
    ['', '', '', 'TỔNG CỘNG PHÁT SINH', 26_100_000],
  ];

  const inspection = inspectRawGrid(rawGrid, 'DateBounds');
  const { items } = extractTransactions(
    rawGrid,
    inspection.detectedHeaderRowIndex,
    inspection.suggestedMapping,
    inspection.columns
  );

  it('3.1 dateFromIso === dateToIso: Chỉ tìm trong đúng 1 ngày duy nhất (kể cả ngày nhuận 29/02), loại bỏ ngày liền trước, liền sau và dòng không ngày', () => {
    // 7.200.000 + 4.800.000 - 500.000 = 11.500.000 (chỉ riêng ngày 2024-02-29)
    // Lưu ý: HD-FEB28 (5.000.000) + HD-MAR01 (6.500.000) cũng bằng 11.500.000 nhưng nằm ngoài ngày 29/02!
    const singleDayRes = findCombinations({
      items,
      dateFromIso: '2024-02-29',
      dateToIso: '2024-02-29',
      targetAmount: 11_500_000,
      tolerance: 0,
    });

    assert.equal(singleDayRes.error, null);
    assert.equal(singleDayRes.solutions.length, 1);
    assert.equal(singleDayRes.solutions[0].dateSpanDays, 0);
    assert.deepEqual(
      singleDayRes.solutions[0].items.map((it) => it.reference),
      ['HD-FEB29-1', 'HD-FEB29-2', 'CK-FEB29']
    );

    // Tìm 3.500.000 trong ngày 2024-02-29: không được lấy dòng trống ngày HD-NODATE-1 (3.500.000)
    const noDatelessLeak = findCombinations({
      items,
      dateFromIso: '2024-02-29',
      dateToIso: '2024-02-29',
      targetAmount: 3_500_000,
      tolerance: 0,
    });
    assert.equal(noDatelessLeak.solutions.length, 0);
  });

  it('3.2 dateFromIso > dateToIso: Trả về kết quả rỗng an toàn kèm thông báo lỗi tiếng Việt rõ ràng, không crash', () => {
    const invertedRes = findCombinations({
      items,
      dateFromIso: '2024-03-15',
      dateToIso: '2024-02-01',
      targetAmount: 11_500_000,
      tolerance: 500,
    });

    assert.deepEqual(invertedRes.solutions, []);
    assert.equal(invertedRes.stats.exactCount, 0);
    assert.equal(invertedRes.stats.nearCount, 0);
    assert.equal(typeof invertedRes.error, 'string');
    assert.ok(
      invertedRes.error.length > 0,
      'Phải trả về thông báo lỗi cho người dùng khi Từ ngày > Đến ngày'
    );
  });

  it('3.3 Dòng trống cột Ngày (dateIso === null): Bỏ qua khi có lọc ngày (trừ khi được ghim pinnedIds), và tham gia gom bình thường khi KHÔNG lọc ngày', () => {
    const dateless1 = items.find((it) => it.reference === 'HD-NODATE-1');
    const dateless2 = items.find((it) => it.reference === 'CK-NODATE-2');
    assert.ok(dateless1 && dateless2);
    assert.equal(dateless1.dateIso, null);
    assert.equal(dateless2.dateIso, null);

    // Mục tiêu: HD-MAR01 (6.500.000) + HD-NODATE-1 (3.500.000) + CK-NODATE-2 (-400.000) = 9.600.000

    // (a) Có lọc ngày (kể cả chỉ lọc 1 đầu dateFromIso hoặc cả 2 đầu) & KHÔNG ghim -> Bỏ qua dòng không ngày
    const filteredOutRes = findCombinations({
      items,
      dateFromIso: '2024-02-01',
      dateToIso: '2024-03-31',
      targetAmount: 9_600_000,
      tolerance: 0,
    });
    assert.equal(
      filteredOutRes.solutions.length,
      0,
      'Khi đang bật lọc ngày và không ghim, các dòng dateIso === null phải bị loại bỏ'
    );

    const openEndedDateRes = findCombinations({
      items,
      dateFromIso: '2024-02-01',
      dateToIso: null,
      targetAmount: 9_600_000,
      tolerance: 0,
    });
    assert.equal(openEndedDateRes.solutions.length, 0);

    // (b) Có lọc ngày NHƯNG kế toán chủ động ghim (pinnedIds) dòng không ngày -> Bắt buộc tham gia!
    // Mục tiêu: HD-MAR01 (6.500.000) + HD-NODATE-1 (3.500.000, pinned) = 10.000.000
    const pinnedDatelessRes = findCombinations({
      items,
      dateFromIso: '2024-03-01',
      dateToIso: '2024-03-31',
      pinnedIds: [dateless1.id],
      targetAmount: 10_000_000,
      tolerance: 0,
    });
    assert.equal(pinnedDatelessRes.error, null);
    assert.equal(pinnedDatelessRes.solutions.length, 1);
    assert.deepEqual(
      pinnedDatelessRes.solutions[0].items.map((it) => it.reference),
      ['HD-MAR01', 'HD-NODATE-1'],
      'Dòng không ngày được ghim trong pinnedIds vẫn phải tham gia vào nghiệm kể cả khi có lọc ngày'
    );

    // (c) KHÔNG lọc ngày (dateFromIso: null, dateToIso: null) -> Các dòng không ngày tham gia gom bình thường!
    const noDateFilterRes = findCombinations({
      items,
      dateFromIso: null,
      dateToIso: null,
      targetAmount: 9_600_000,
      tolerance: 0,
    });
    assert.equal(noDateFilterRes.error, null);
    assert.equal(noDateFilterRes.solutions.length, 1);
    assert.deepEqual(
      noDateFilterRes.solutions[0].items.map((it) => it.reference),
      ['HD-MAR01', 'HD-NODATE-1', 'CK-NODATE-2'],
      'Khi không lọc ngày, cả dòng có ngày và dòng không ngày đều được gom bình thường'
    );
  });

  it('3.4 [Adversarial] Dòng giao dịch thiếu ngày có diễn giải "Công lắp đặt..." / "Công trình..." (chữ Ô, không phải "Cộng" chữ Ộ) không được bị nhận nhầm là dòng tổng cộng!', () => {
    const constructionGrid = [
      ['Ngày chứng từ', 'Số chứng từ', 'Diễn giải', 'Số tiền'],
      ['10/03/2026', 'HD-TB01', 'Xuất bán tủ điện tổng MSB 400A', 15_000_000],
      ['', 'PT-LD01', 'Công lắp đặt & đấu nối tủ điện tại xưởng khách hàng', 3_500_000],
      ['', 'PT-CT02', 'Công trình nhà máy giai đoạn 2 - phụ phí vận chuyển', 1_500_000],
      ['', '', 'Cộng phát sinh', 20_000_000],
    ];

    const ins = inspectRawGrid(constructionGrid, 'Construction_Dateless');
    const ext = extractTransactions(
      constructionGrid,
      ins.detectedHeaderRowIndex,
      ins.suggestedMapping,
      ins.columns
    );

    const laborRow = ext.items.find((it) => it.reference === 'PT-LD01');
    const projectRow = ext.items.find((it) => it.reference === 'PT-CT02');
    assert.ok(laborRow && projectRow);

    assert.equal(
      laborRow.isSummaryRow,
      false,
      'Khoản thu "Công lắp đặt..." thiếu ngày là nghiệp vụ thực (chữ Công), KHÔNG PHẢI dòng "Cộng" tổng hợp!'
    );
    assert.equal(
      projectRow.isSummaryRow,
      false,
      'Khoản thu "Công trình..." thiếu ngày là nghiệp vụ thực, KHÔNG PHẢI dòng "Cộng" tổng hợp!'
    );

    // Khi không lọc ngày, gom 20.000.000đ phải tìm được bộ 3 dòng thực (15tr + 3.5tr + 1.5tr)
    const searchRes = findCombinations({
      items: ext.items,
      dateFromIso: null,
      dateToIso: null,
      targetAmount: 20_000_000,
      tolerance: 0,
    });
    assert.equal(searchRes.solutions.length, 1);
    assert.deepEqual(
      searchRes.solutions[0].items.map((it) => it.reference),
      ['HD-TB01', 'PT-LD01', 'PT-CT02']
    );
  });

  it('3.5 [Adversarial] Dòng khóa sổ cuối tháng có điền ngày ("31/01/2026" | "Cộng tháng 01/2026") phải được đánh dấu isSummaryRow = true để tránh gom trùng gấp đôi công nợ', () => {
    const misaMonthlyGrid = [
      ['Ngày chứng từ', 'Số chứng từ', 'Diễn giải', 'Số tiền'],
      ['10/01/2026', 'HD01', 'Bán hàng đợt 1', 12_000_000],
      ['20/01/2026', 'HD02', 'Bán hàng đợt 2', 8_000_000],
      ['31/01/2026', '', 'Cộng tháng 01/2026', 20_000_000],
      ['15/02/2026', 'HD03', 'Bán hàng đợt 3', 5_000_000],
      ['', '', 'Tổng cộng', 25_000_000],
    ];

    const ins = inspectRawGrid(misaMonthlyGrid, 'Misa_Subtotal');
    const ext = extractTransactions(
      misaMonthlyGrid,
      ins.detectedHeaderRowIndex,
      ins.suggestedMapping,
      ins.columns
    );

    const janSubtotalRow = ext.items.find((it) => it.description === 'Cộng tháng 01/2026');
    assert.ok(janSubtotalRow);
    assert.equal(
      janSubtotalRow.isSummaryRow,
      true,
      'Dòng "Cộng tháng 01/2026" dù có điền ngày cuối tháng 31/01/2026 vẫn là dòng tổng cộng tháng, tuyệt đối không được coi là hóa đơn thật!'
    );
  });
});
