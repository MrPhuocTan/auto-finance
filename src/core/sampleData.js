/**
 * Realistic Vietnamese Accounting Sample Dataset (Quý 1/2026)
 * Company: CÔNG TY TNHH THƯƠNG MẠI VÀ DỊCH VỤ PHƯỚC TÂN
 *
 * Designed with mathematically verified reconciliation scenarios:
 * - Preset 1 (01/03/2026 - 15/03/2026, target 18.450.000, tol ±500):
 *     * 1 exact match (0đ): 4 positive invoices (HD0301 + HD0303 + HD0304 + PT0301 = 19.000.000)
 *       + 1 commercial discount (CK0301 = -550.000) = 18.450.000
 *     * 1 near match (-350đ): HD0302 (9.119.650) + HD0305 (9.330.000) = 18.449.650
 * - Preset 2 (16/03/2026 - 31/03/2026, target 24.800.000, tol ±500):
 *     * 1 near match (+250đ): HD0307 (10.450.120) + HD0309 (8.620.130) + HD0311 (5.730.000) = 24.800.250
 * - Preset 3 (01/02/2026 - 28/02/2026, target 32.150.000, tol ±500, keyword "Minh Phú"):
 *     * 1 exact match (0đ) for Minh Phú: HD0201 (14.500.000) + HD0203 (11.850.000)
 *       + HD0205 (7.800.000) + CT0201 (-2.000.000 deposit offset) = 32.150.000
 */

const REPORT_HEADER_ROWS = [
  ['CÔNG TY TNHH THƯƠNG MẠI VÀ DỊCH VỤ PHƯỚC TÂN', '', '', '', '', ''],
  ['Địa chỉ: Quận Hải Châu, TP. Đà Nẵng', '', '', '', '', ''],
  ['SỔ CHI TIẾT ĐỐI SOÁT HÓA ĐƠN & CÔNG NỢ PHẢI THU - QUÝ 1/2026', '', '', '', '', ''],
  ['', '', '', '', '', ''],
];

const TABLE_HEADER_ROW = [
  'Ngày chứng từ',
  'Số chứng từ',
  'Khách hàng / Đối tác',
  'Diễn giải nội dung',
  'Số tiền (VNĐ)',
  'Ghi chú',
];

const TRANSACTION_ROWS = [
  // ================= THÁNG 01/2026 (5 nghiệp vụ) =================
  [
    '05/01/2026',
    'HD0101',
    'Công ty TNHH Thương mại Hòa Bình',
    'Xuất bán sơn chống thấm ngoại thất Dulux - Đợt 1',
    8450150,
    'VAT 8%',
  ],
  [
    '09/01/2026',
    'HD0102',
    'Công ty TNHH Thương mại Hòa Bình',
    'Xuất bán bột trét tường & rulo lăn sơn công trình Hòa Bình - Đợt 2',
    7150000,
    'VAT 8%',
  ],
  [
    '15/01/2026',
    'HD0103',
    'Cửa hàng VLXD Đại Lộc',
    'Bán sỉ ống nhựa PVC Bình Minh phi 27-90 theo đơn ĐL-01',
    12300000,
    'Đã giao đủ',
  ],
  [
    '21/01/2026',
    'TL0101',
    'Cửa hàng VLXD Đại Lộc',
    'Giảm trừ hàng trả lại do móp méo khi vận chuyển đơn ĐL-01',
    -850000,
    'Biên bản trả hàng',
  ],
  [
    '28/01/2026',
    'HD0104',
    'Công ty CP Nội thất Hoàng Gia',
    'Cung cấp ván MDF chống ẩm phủ Melamine lô HG-01',
    16900000,
    'Hẹn TT tháng 02',
  ],

  // ================= THÁNG 02/2026 (9 nghiệp vụ) =================
  [
    '03/02/2026',
    'HD0201',
    'Công ty TNHH Xây dựng Minh Phú',
    'Xuất bán thép hộp mạ kẽm công trình Minh Phú - Đợt 1',
    14500000,
    'Giao tại công trình',
  ],
  [
    '06/02/2026',
    'HD0202',
    'Công ty CP Nội thất Hoàng Gia',
    'Xuất bán phụ kiện bản lề giảm chấn Hafele lô HG-02',
    9650000,
    'VAT 8%',
  ],
  [
    '10/02/2026',
    'HD0203',
    'Công ty TNHH Xây dựng Minh Phú',
    'Xuất bán xi măng & phụ gia chống thấm công trình Minh Phú - Đợt 2',
    11850000,
    'Đã ký nhận',
  ],
  [
    '13/02/2026',
    'HD0204',
    'Công ty TNHH Cơ khí Hải Vân',
    'Bán đá cắt sắt, que hàn chịu lực & vật tư tiêu hao xưởng Hải Vân',
    9020000,
    'VAT 10%',
  ],
  [
    '18/02/2026',
    'HD0205',
    'Công ty TNHH Xây dựng Minh Phú',
    'Xuất bán vật tư hoàn thiện công trình Minh Phú - Đợt 3',
    7800000,
    'Đã đối chiếu',
  ],
  [
    '20/02/2026',
    'CT0201',
    'Công ty TNHH Xây dựng Minh Phú',
    'Cấn trừ tiền đặt cọc hợp đồng nguyên tắc Minh Phú số 02/HĐNT',
    -2000000,
    'Cấn trừ đặt cọc',
  ],
  [
    '23/02/2026',
    'CK0201',
    'Công ty CP Nội thất Hoàng Gia',
    'Chiết khấu doanh số tháng 01-02 cho đối tác Hoàng Gia',
    -750000,
    'Phiếu điều chỉnh',
  ],
  [
    '25/02/2026',
    'HD0206',
    'Cửa hàng VLXD Đại Lộc',
    'Giao hàng thiết bị vệ sinh Inax & vòi sen nóng lạnh đợt 1',
    13480000,
    'VAT 8%',
  ],
  [
    '27/02/2026',
    'HD0207',
    'Công ty TNHH Xây dựng Minh Phú',
    'Giao bổ sung gạch ốp lát công trình Minh Phú (kỳ TT tháng 03)',
    9400000,
    'Chưa đến hạn TT',
  ],

  // ================= THÁNG 03/2026 - KỲ ĐẦU THÁNG 01/03..15/03 (9 nghiệp vụ) =================
  [
    '02/03/2026',
    'HD0301',
    'Công ty TNHH Điện nước Thành Đạt',
    'Xuất bán cáp điện Cadivi CVV 3x2.5 - Giao hàng đợt 1',
    6250000,
    'VAT 8%',
  ],
  [
    '04/03/2026',
    'HD0302',
    'Công ty CP Thương mại An Khang',
    'Xuất bán tấm thạch cao Vĩnh Tường & khung xương chìm',
    9119650,
    'Lẻ thuế VAT 8%',
  ],
  [
    '05/03/2026',
    'HD0303',
    'Công ty TNHH Điện nước Thành Đạt',
    'Xuất bán tủ điện âm tường & aptomat Schneider - Giao đợt 2',
    4850000,
    'VAT 8%',
  ],
  [
    '07/03/2026',
    'TL0301',
    'Cửa hàng VLXD Đại Lộc',
    'Hàng bán bị trả lại do giao nhầm mẫu gạch lát nền 60x60',
    -1200000,
    'Hóa đơn điều chỉnh',
  ],
  [
    '09/03/2026',
    'HD0304',
    'Công ty TNHH Điện nước Thành Đạt',
    'Cung cấp đèn LED âm trần Rạng Đông 9W-12W theo báo giá',
    5100000,
    'VAT 8%',
  ],
  [
    '11/03/2026',
    'HD0305',
    'Công ty CP Thương mại An Khang',
    'Giao hàng sơn lót kháng kiềm & vật tư phụ hoàn thiện trần',
    9330000,
    'Đã nghiệm thu',
  ],
  [
    '12/03/2026',
    'PT0301',
    'Công ty TNHH Điện nước Thành Đạt',
    'Phụ phí vận chuyển xe tải & bốc xếp giao tận kho Thành Đạt',
    2800000,
    'Chứng từ vận chuyển',
  ],
  [
    '13/03/2026',
    'CK0301',
    'Công ty TNHH Điện nước Thành Đạt',
    'Chiết khấu thương mại đơn hàng thiết bị điện đầu tháng 03',
    -550000,
    'Giảm trừ trực tiếp',
  ],
  [
    '15/03/2026',
    'HD0306',
    'Cửa hàng VLXD Đại Lộc',
    'Xuất đổi lại lô gạch lát nền Granite 60x60 vân đá trắng',
    15920000,
    'Đã giao bù',
  ],

  // ================= THÁNG 03/2026 - KỲ CUỐI THÁNG 16/03..31/03 (7 nghiệp vụ) =================
  [
    '18/03/2026',
    'HD0307',
    'Công ty TNHH Kỹ thuật Nam Việt',
    'Xuất bán ống nhựa chịu nhiệt PPR Bình Minh - Đợt 1',
    10450120,
    'Lẻ thuế VAT 8%',
  ],
  [
    '19/03/2026',
    'HD0308',
    'Công ty TNHH Cơ khí Hải Vân',
    'Cung cấp thép tấm cán nóng & bu lông cường độ cao M16',
    17650000,
    'VAT 10%',
  ],
  [
    '23/03/2026',
    'HD0309',
    'Công ty TNHH Kỹ thuật Nam Việt',
    'Xuất bán van bi đồng, co nối ren & keo dán ống - Đợt 2',
    8620130,
    'Lẻ thuế VAT 8%',
  ],
  [
    '25/03/2026',
    'HD0310',
    'Công ty CP Nội thất Hoàng Gia',
    'Xuất bán sàn gỗ công nghiệp chống nước 12mm lô HG-03',
    13200000,
    'Đã giao kho',
  ],
  [
    '27/03/2026',
    'HD0311',
    'Công ty TNHH Kỹ thuật Nam Việt',
    'Giao bổ sung máy bơm tăng áp Panasonic & phụ phí lắp đặt',
    5730000,
    'Biên bản bàn giao',
  ],
  [
    '28/03/2026',
    'CK0302',
    'Công ty TNHH Cơ khí Hải Vân',
    'Chiết khấu thanh toán sớm trước hạn hợp đồng tháng 03',
    -450000,
    'Trừ công nợ',
  ],
  [
    '31/03/2026',
    'HD0312',
    'Công ty TNHH Thương mại Hòa Bình',
    'Xuất bán vật tư chống thấm sân thượng Sika TopSeal 107',
    11380000,
    'Chốt sổ cuối quý',
  ],
];

const TOTAL_AMOUNT = TRANSACTION_ROWS.reduce((sum, row) => sum + Number(row[4] || 0), 0);

const SUMMARY_FOOTER_ROW = ['', '', '', 'TỔNG CỘNG PHÁT SINH', TOTAL_AMOUNT, ''];

const SAMPLE_PRESETS = [
  {
    label: 'Đầu tháng 03: Gom 18.450.000đ (có trừ chiết khấu)',
    dateFromIso: '2026-03-01',
    dateToIso: '2026-03-15',
    targetAmount: 18450000,
    tolerance: 500,
    keyword: '',
  },
  {
    label: 'Cuối tháng 03: Khách CK chẵn 24.800.000đ (lệch lẻ 250đ)',
    dateFromIso: '2026-03-16',
    dateToIso: '2026-03-31',
    targetAmount: 24800000,
    tolerance: 500,
    keyword: '',
  },
  {
    label: 'Tháng 02: Công nợ Minh Phú 32.150.000đ (cấn trừ cọc)',
    dateFromIso: '2026-02-01',
    dateToIso: '2026-02-28',
    targetAmount: 32150000,
    tolerance: 500,
    keyword: 'Minh Phú',
  },
];

/**
 * Returns a fresh copy of the sample Vietnamese accounting dataset & presets.
 * @returns {{
 *   fileName: string,
 *   sheetName: string,
 *   rawGrid: any[][],
 *   presets: Array<{
 *     label: string,
 *     dateFromIso: string,
 *     dateToIso: string,
 *     targetAmount: number,
 *     tolerance: number,
 *     keyword: string
 *   }>
 * }}
 */
export function getSampleDataset() {
  const rawGrid = [
    ...REPORT_HEADER_ROWS.map((row) => [...row]),
    [...TABLE_HEADER_ROW],
    ...TRANSACTION_ROWS.map((row) => [...row]),
    [...SUMMARY_FOOTER_ROW],
  ];

  const presets = SAMPLE_PRESETS.map((preset) => ({ ...preset }));

  return {
    fileName: 'So_Chi_Tiet_Cong_No_Q1_2026.xlsx',
    sheetName: 'CongNo_PhaiThu_2026',
    rawGrid,
    presets,
  };
}
