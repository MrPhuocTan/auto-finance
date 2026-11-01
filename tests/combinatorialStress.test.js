import test from 'node:test';
import assert from 'node:assert/strict';
import { performance } from 'node:perf_hooks';
import { Worker as NodeWorker } from 'node:worker_threads';
import { pathToFileURL } from 'node:url';
import path from 'node:path';

import { findCombinations } from '../src/core/reconcileEngine.js';
import { handleWorkerMessage } from '../src/core/reconcileWorker.js';

const WORKER_MODULE_URL = pathToFileURL(
  path.resolve(process.cwd(), 'src/core/reconcileWorker.js')
).href;

/**
 * Helper tạo dòng giao dịch chuẩn xác theo hợp đồng dữ liệu TransactionItem.
 */
function makeTx({
  id,
  excelRowNumber,
  dateIso = '2026-03-10',
  amount,
  description = '',
  reference = '',
  isSummaryRow = false,
}) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(dateIso);
  const timestamp = m ? Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : null;
  const dateDisplay = m ? `${m[3]}/${m[2]}/${m[1]}` : '';
  return {
    id,
    excelRowNumber,
    dateIso,
    dateDisplay,
    timestamp,
    amount,
    description: description || `Giao dịch ${id}`,
    reference: reference || id.toUpperCase(),
    isSummaryRow,
    extraCells: [],
    rawCells: [dateDisplay, reference || id, amount, description],
  };
}

// ============================================================================
// PHẦN 1: BẪY TỔ HỢP NP-HARD, VÔ NGHIỆM (MODULAR TRAP) & SUÝT TRÚNG DÀY ĐẶC
// ============================================================================

test('1.1 [NP-Hard Parity/Modular Trap] 32 dòng số chẵn nghìn tìm Z = 15.000.600đ (±500đ) vô nghiệm, không treo và tôn trọng timeLimitMs', () => {
  // Ca A: 32 hóa đơn toàn bội số của 2.000.000 (2.000.000, 4.000.000, ..., -6.000.000, ...)
  // Mọi tổng con đều chia hết cho 2.000.000 -> gần 15.000.600 nhất là 14.000.000 (-1.000.600) hoặc 16.000.000 (+999.400)
  const evenMillionItems = [];
  for (let i = 1; i <= 24; i++) {
    const day = String(((i - 1) % 28) + 1).padStart(2, '0');
    evenMillionItems.push(
      makeTx({
        id: `even2m-pos-${i}`,
        excelRowNumber: i + 1,
        dateIso: `2026-03-${day}`,
        amount: i * 2_000_000,
        description: `Hóa đơn chẵn triệu #${i}`,
      })
    );
  }
  for (let j = 1; j <= 8; j++) {
    const day = String(((j - 1) % 28) + 1).padStart(2, '0');
    evenMillionItems.push(
      makeTx({
        id: `even2m-neg-${j}`,
        excelRowNumber: 25 + j,
        dateIso: `2026-03-${day}`,
        amount: -j * 2_000_000,
        description: `Giảm trừ chẵn triệu #${j}`,
      })
    );
  }
  assert.equal(evenMillionItems.length, 32);

  const t0 = performance.now();
  const resMillion = findCombinations({
    items: evenMillionItems,
    targetAmount: 15_000_600,
    tolerance: 500,
    minItems: 1,
    maxItems: 12,
    timeLimitMs: 400,
  });
  const elapsedMillion = performance.now() - t0;

  assert.equal(resMillion.error, null);
  assert.equal(resMillion.solutions.length, 0);
  assert.ok(
    elapsedMillion < 600,
    `Ca 32 bội số 2.000.000đ phải thoát nhanh, thực tế mất ${elapsedMillion.toFixed(1)}ms`
  );

  // Ca B (Worst-case NP-Hard Parity Trap): 32 dòng phân biệt đều là bội số chẵn của 2.000đ (2.000 * m) cụm sát nhau
  // Z = 15.000.600 mod 2.000 = 600 -> gần nhất là 15.000.000 (-600) và 15.002.000 (+1.400) -> hoàn toàn vô nghiệm ngoài ±500!
  // Do các số xấp xỉ nhau, cây nhánh-cận mở rộng hàng triệu nút C(32, <=12) -> kiểm chứng ngắt đúng hạn timeLimitMs!
  const clusteredEvenThousandItems = [];
  for (let i = 1; i <= 20; i++) {
    clusteredEvenThousandItems.push(
      makeTx({
        id: `even2k-pos-${i}`,
        excelRowNumber: i + 1,
        dateIso: '2026-03-10',
        amount: 2_000_000 + i * 2_000,
      })
    );
  }
  for (let j = 1; j <= 12; j++) {
    clusteredEvenThousandItems.push(
      makeTx({
        id: `even2k-neg-${j}`,
        excelRowNumber: 21 + j,
        dateIso: '2026-03-10',
        amount: -(2_000_000 + j * 2_000),
      })
    );
  }
  assert.equal(clusteredEvenThousandItems.length, 32);

  const strictLimitMs = 250;
  const t1 = performance.now();
  const resClustered = findCombinations({
    items: clusteredEvenThousandItems,
    targetAmount: 15_000_600,
    tolerance: 500,
    minItems: 1,
    maxItems: 12,
    timeLimitMs: strictLimitMs,
  });
  const elapsedClustered = performance.now() - t1;

  assert.equal(resClustered.error, null);
  assert.equal(resClustered.solutions.length, 0);
  assert.equal(resClustered.stats.timedOut, true, 'Cây vô nghiệm bùng nổ phải bật cờ stats.timedOut = true');
  assert.ok(
    elapsedClustered < strictLimitMs + 150,
    `Bộ giải phải ngắt đúng hạn timeLimitMs=${strictLimitMs}ms, thực tế mất ${elapsedClustered.toFixed(1)}ms`
  );
});

test('1.2 [Dense Near-Misses Trap] 40 dòng có >8.000 tổ hợp rơi vào ±504..±510đ (sát mép ngoài ±500đ), tìm đúng DUY NHẤT 1 nghiệm 6 dòng ở diff = 0 và diff = +120', () => {
  const buildDenseNearMissDataset = (goldenDelta) => {
    const items = [];
    let rowNo = 2;

    // 34 dòng phân biệt tạo >8.000 tổ hợp 6 dòng có tổng = 60.000.000 + (504..510), nằm sát mép ngoài ±500
    for (let i = 1; i <= 17; i++) {
      const day = String(((i - 1) % 28) + 1).padStart(2, '0');
      const remA = i % 2 === 0 ? 84 : 85;
      items.push(
        makeTx({
          id: `nm-a-${i}`,
          excelRowNumber: rowNo++,
          dateIso: `2026-02-${day}`,
          amount: 10_000_000 + 2_000 * i + remA,
        })
      );
      items.push(
        makeTx({
          id: `nm-b-${i}`,
          excelRowNumber: rowNo++,
          dateIso: `2026-02-${day}`,
          amount: 10_000_000 - 2_000 * i + 84,
        })
      );
    }

    // 6 dòng "kim đáy bể" duy nhất có tổng = 60.000.000 + goldenDelta
    const goldenAmounts = [
      6_285_000,
      8_285_000,
      10_285_000,
      12_285_000,
      14_285_000,
      8_575_000 + goldenDelta,
    ];
    for (let g = 0; g < goldenAmounts.length; g++) {
      items.push(
        makeTx({
          id: `gold-${g + 1}`,
          excelRowNumber: rowNo++,
          dateIso: '2026-03-15',
          amount: goldenAmounts[g],
        })
      );
    }

    assert.equal(items.length, 40);
    return items;
  };

  // Ca A: Duy nhất 1 tổ hợp 6 dòng khớp tuyệt đối diff = 0
  const exactItems = buildDenseNearMissDataset(0);
  const resExact = findCombinations({
    items: exactItems,
    targetAmount: 60_000_000,
    tolerance: 500,
    minItems: 1,
    maxItems: 6,
    timeLimitMs: 4000,
  });

  assert.equal(resExact.error, null);
  assert.equal(resExact.solutions.length, 1, 'Phải bắt trúng duy nhất 1 nghiệm 6 dòng khớp diff = 0');
  assert.equal(resExact.solutions[0].isExact, true);
  assert.equal(resExact.solutions[0].diff, 0);
  assert.equal(resExact.solutions[0].itemCount, 6);
  assert.deepEqual(
    [...resExact.solutions[0].itemIds].sort(),
    ['gold-1', 'gold-2', 'gold-3', 'gold-4', 'gold-5', 'gold-6']
  );

  // Ca B: Duy nhất 1 tổ hợp 6 dòng khớp lệch nhẹ diff = +120
  const nearItems = buildDenseNearMissDataset(120);
  const resNear = findCombinations({
    items: nearItems,
    targetAmount: 60_000_000,
    tolerance: 500,
    minItems: 1,
    maxItems: 6,
    timeLimitMs: 4000,
  });

  assert.equal(resNear.error, null);
  assert.equal(resNear.solutions.length, 1, 'Phải bắt trúng duy nhất 1 nghiệm 6 dòng khớp diff = +120');
  assert.equal(resNear.solutions[0].isExact, false);
  assert.equal(resNear.solutions[0].diff, 120);
  assert.equal(resNear.solutions[0].absDiff, 120);
  assert.equal(resNear.solutions[0].itemCount, 6);
  assert.deepEqual(
    [...resNear.solutions[0].itemIds].sort(),
    ['gold-1', 'gold-2', 'gold-3', 'gold-4', 'gold-5', 'gold-6']
  );
});

test('1.3 [Dynamic Tightening vs DateSpan Priority] Khi có >200 nghiệm cùng mức k=6 lệch ngày xa (88 ngày), siết biên động KHÔNG được đánh mất nghiệm cùng ngày (dateSpanDays = 0)', () => {
  // Kiểm chứng cho cả diff = +120 (bị mất do levelIntTolerance = worstKept.intAbsDiff - 1)
  // và diff = 0 (bị mất do stopCurrentLevel = true khi worstKept.intAbsDiff === 0)
  for (const targetDelta of [120, 0]) {
    const items = [];
    let rowNo = 2;

    const perItemDelta = targetDelta / 6; // 20đ/dòng khi targetDelta=120 -> mọi tập con 6 dòng đều ≡ 120 (mod 1000)

    for (let i = 1; i <= 6; i++) {
      items.push(makeTx({ id: `p1a-${i}`, excelRowNumber: rowNo++, dateIso: '2026-01-01', amount: 9_000_000 + i * 1_000 + perItemDelta }));
      items.push(makeTx({ id: `p1b-${i}`, excelRowNumber: rowNo++, dateIso: '2026-03-30', amount: 8_000_000 - i * 1_000 + perItemDelta }));

      items.push(makeTx({ id: `p2a-${i}`, excelRowNumber: rowNo++, dateIso: '2026-01-01', amount: 8_500_000 + i * 10_000 + perItemDelta }));
      items.push(makeTx({ id: `p2b-${i}`, excelRowNumber: rowNo++, dateIso: '2026-03-30', amount: 6_500_000 - i * 10_000 + perItemDelta }));

      items.push(makeTx({ id: `p3a-${i}`, excelRowNumber: rowNo++, dateIso: '2026-01-01', amount: 7_200_000 + i * 100_000 + perItemDelta }));
      items.push(makeTx({ id: `p3b-${i}`, excelRowNumber: rowNo++, dateIso: '2026-03-30', amount: 5_800_000 - i * 100_000 + perItemDelta }));
    }

    const sameDayAmounts = [7_501_000, 7_502_000, 7_503_000, 7_497_000, 7_498_000, 7_499_000].map(
      (a) => a + perItemDelta
    );
    for (let s = 0; s < sameDayAmounts.length; s++) {
      items.push(
        makeTx({
          id: `sameday-${s + 1}`,
          excelRowNumber: rowNo++,
          dateIso: '2026-03-15',
          amount: sameDayAmounts[s],
        })
      );
    }

    const res = findCombinations({
      items,
      targetAmount: 45_000_000,
      tolerance: 500,
      minItems: 6,
      maxItems: 6,
      maxResults: 10,
      timeLimitMs: 4000,
    });

    assert.equal(res.error, null);
    assert.ok(res.solutions.length >= 1);
    assert.equal(
      res.solutions[0].dateSpanDays,
      0,
      `[diff=${targetDelta}] Nghiệm cùng ngày (dateSpanDays = 0) phải xếp hạng #1 thay vì bị cắt tỉa mất bởi compactLevelIfNeeded (hiện nhận dateSpanDays = ${res.solutions[0].dateSpanDays})`
    );
  }
});

// ============================================================================
// PHẦN 2: CA KHÓ VỀ NHIỀU DÒNG TRÙNG MỆNH GIÁ + PINNED + EXCLUDED + DATES
// ============================================================================

test('2.1 [Duplicate + Pinned + Excluded] 15 dòng +500.000đ & 10 dòng -150.000đ: ghim 2 dòng, loại 3 dòng, tìm 1.700.000đ', () => {
  const items = [];
  let rowNo = 2;

  for (let i = 1; i <= 15; i++) {
    const day = String(i).padStart(2, '0');
    items.push(
      makeTx({
        id: `pos500-${i}`,
        excelRowNumber: rowNo++,
        dateIso: `2026-03-${day}`,
        amount: 500_000,
        description: `Phí dịch vụ lô #${i}`,
      })
    );
  }

  for (let j = 1; j <= 10; j++) {
    const day = String(j * 2).padStart(2, '0');
    items.push(
      makeTx({
        id: `neg150-${j}`,
        excelRowNumber: rowNo++,
        dateIso: `2026-03-${day}`,
        amount: -150_000,
        description: `Chiết khấu thương mại #${j}`,
      })
    );
  }

  const pinnedIds = ['pos500-3', 'pos500-9'];
  const excludedIds = ['pos500-1', 'pos500-5', 'pos500-12'];

  const res = findCombinations({
    items,
    targetAmount: 1_700_000, // = 4 * 500.000 + 2 * (-150.000)
    tolerance: 0,
    pinnedIds,
    excludedIds,
    minItems: 1,
    maxItems: 8,
    maxResults: 50,
  });

  assert.equal(res.error, null);
  assert.equal(res.stats.pinnedCount, 2);
  assert.equal(res.stats.candidateCount, 20);
  assert.ok(res.solutions.length >= 1, 'Phải tìm thấy nghiệm cho mục tiêu 1.700.000đ');

  for (const sol of res.solutions) {
    assert.equal(sol.totalAmount, 1_700_000);
    assert.equal(sol.isExact, true);

    // 1) Bắt buộc chứa đủ 2 ID đã ghim
    for (const pid of pinnedIds) {
      assert.ok(sol.itemIds.includes(pid), `Nghiệm ${sol.id} thiếu dòng đã ghim ${pid}`);
    }

    // 2) Tuyệt đối không chứa bất kỳ ID nào đã loại trừ
    for (const eid of excludedIds) {
      assert.ok(!sol.itemIds.includes(eid), `Nghiệm ${sol.id} chứa dòng đã loại trừ ${eid}`);
    }

    // 3) Không trùng lặp cùng 1 itemId 2 lần trong cùng 1 nghiệm
    const uniqueSet = new Set(sol.itemIds);
    assert.equal(uniqueSet.size, sol.itemIds.length, `Nghiệm ${sol.id} bị lặp itemId`);
    assert.equal(sol.items.length, sol.itemIds.length);
  }

  // Nghiệm 6 dòng: C(10,2) * C(10,2) = 2.025 biến thể, xuất top 3 -> còn 2.022 cách chọn
  const sixItemSolutions = res.solutions.filter((s) => s.itemCount === 6);
  assert.equal(sixItemSolutions.length, 3);
  assert.ok(
    sixItemSolutions[0].duplicateVariantsNote?.includes('2.022'),
    `Ghi chú trùng mệnh giá phải phản ánh 2.022 biến thể còn lại, nhận được: ${sixItemSolutions[0].duplicateVariantsNote}`
  );
});

test('2.2 [Multi-Group Duplicate Date Alignment] 15 dòng +500.000đ & 10 dòng -150.000đ có cụm 6 dòng CÙNG NGÀY 20/03/2026 phải đạt dateSpanDays = 0', () => {
  const items = [];
  let rowNo = 2;

  const posDates = [
    '2026-03-01', '2026-03-02', '2026-03-03', // excluded
    '2026-03-20', '2026-03-20',               // pinned (20/03/2026)
    '2026-03-01', '2026-03-02', '2026-03-03', '2026-03-04',
    '2026-03-05', '2026-03-06', '2026-03-07', '2026-03-08',
    '2026-03-20', '2026-03-20',               // candidate idx 8, 9 (20/03/2026)
  ];
  for (let i = 0; i < 15; i++) {
    items.push(
      makeTx({
        id: `dpos-${i + 1}`,
        excelRowNumber: rowNo++,
        dateIso: posDates[i],
        amount: 500_000,
      })
    );
  }

  const negDates = [
    '2026-03-01', '2026-03-02', '2026-03-03', '2026-03-04', '2026-03-05',
    '2026-03-20', '2026-03-20', // candidate idx 5, 6 (20/03/2026)
    '2026-03-28', '2026-03-29', '2026-03-30',
  ];
  for (let j = 0; j < 10; j++) {
    items.push(
      makeTx({
        id: `dneg-${j + 1}`,
        excelRowNumber: rowNo++,
        dateIso: negDates[j],
        amount: -150_000,
      })
    );
  }

  const res = findCombinations({
    items,
    targetAmount: 1_700_000,
    tolerance: 0,
    pinnedIds: ['dpos-4', 'dpos-5'],
    excludedIds: ['dpos-1', 'dpos-2', 'dpos-3'],
    minItems: 6,
    maxItems: 6,
  });

  assert.equal(res.error, null);
  assert.ok(res.solutions.length >= 1);
  assert.equal(
    res.solutions[0].dateSpanDays,
    0,
    `Phải ưu tiên ghép được bộ chứng từ trùng mệnh giá cùng ngày 20/03/2026 (dateSpanDays = 0), thực tế nhận dateSpanDays = ${res.solutions[0].dateSpanDays}`
  );
});

test('2.3 [Duplicate 3 vs 4 Rows of 1.000.000đ] Chọn 2 dòng từ 3 dòng (trả đủ 3, note = null) và từ 4 dòng (trả top 3 theo ngày, note báo còn 3 cách)', () => {
  // Trường hợp A: Đúng 3 dòng 1.000.000đ -> C(3, 2) = 3 cách chọn
  const threeItems = [
    makeTx({ id: 'm1', excelRowNumber: 2, dateIso: '2026-03-01', amount: 1_000_000 }),
    makeTx({ id: 'm2', excelRowNumber: 3, dateIso: '2026-03-10', amount: 1_000_000 }),
    makeTx({ id: 'm3', excelRowNumber: 4, dateIso: '2026-03-11', amount: 1_000_000 }),
  ];

  const res3 = findCombinations({
    items: threeItems,
    targetAmount: 2_000_000,
    tolerance: 0,
    maxResults: 50,
  });

  assert.equal(res3.solutions.length, 3, 'C(3, 2) = 3 phải trả đủ cả 3 tổ hợp');
  assert.deepEqual(res3.solutions[0].itemIds, ['m2', 'm3']);
  assert.equal(res3.solutions[0].dateSpanDays, 1);
  for (const sol of res3.solutions) {
    assert.equal(sol.duplicateVariantsNote, null);
  }

  // Trường hợp B: Đúng 4 dòng 1.000.000đ -> C(4, 2) = 6 cách chọn (> 3) -> trả top 3 & báo còn 3 cách chọn
  const fourItems = [
    makeTx({ id: 'q1', excelRowNumber: 2, dateIso: '2026-03-01', amount: 1_000_000 }),
    makeTx({ id: 'q2', excelRowNumber: 3, dateIso: '2026-03-05', amount: 1_000_000 }),
    makeTx({ id: 'q3', excelRowNumber: 4, dateIso: '2026-03-10', amount: 1_000_000 }),
    makeTx({ id: 'q4', excelRowNumber: 5, dateIso: '2026-03-10', amount: 1_000_000 }),
  ];

  const res4 = findCombinations({
    items: fourItems,
    targetAmount: 2_000_000,
    tolerance: 0,
    maxResults: 50,
  });

  assert.equal(res4.solutions.length, 3);
  assert.deepEqual(res4.solutions[0].itemIds, ['q3', 'q4']);
  assert.equal(res4.solutions[0].dateSpanDays, 0);
  for (const sol of res4.solutions) {
    assert.equal(
      sol.duplicateVariantsNote,
      'Còn 3 cách chọn tương đương từ các dòng cùng số tiền (1.000.000) có thể thay thế'
    );
  }
});

test('2.4 [duplicateVariantsNote vs maxResults Truncation] Khi maxResults = 1 cắt nghiệm của 4 dòng 1.000.000đ (tổng 6 cách), ghi chú phải báo còn 5 cách', () => {
  const fourItems = [
    makeTx({ id: 'q1', excelRowNumber: 2, dateIso: '2026-03-01', amount: 1_000_000 }),
    makeTx({ id: 'q2', excelRowNumber: 3, dateIso: '2026-03-05', amount: 1_000_000 }),
    makeTx({ id: 'q3', excelRowNumber: 4, dateIso: '2026-03-10', amount: 1_000_000 }),
    makeTx({ id: 'q4', excelRowNumber: 5, dateIso: '2026-03-10', amount: 1_000_000 }),
  ];

  const res4Limited = findCombinations({
    items: fourItems,
    targetAmount: 2_000_000,
    tolerance: 0,
    maxResults: 1,
  });

  assert.equal(res4Limited.solutions.length, 1);
  assert.equal(
    res4Limited.solutions[0].duplicateVariantsNote,
    'Còn 5 cách chọn tương đương từ các dòng cùng số tiền (1.000.000) có thể thay thế',
    'Tổng C(4,2) = 6 cách, chỉ trả cho người dùng 1 cách thì phải báo còn 5 cách thay vì 3'
  );
});

test('2.5 [Extreme Symmetry Breaking] 100 dòng trùng mệnh giá 100.000đ tìm tổng 1.200.000đ (maxItems = 12) chạy < 500ms', () => {
  const hundredItems = [];
  for (let i = 1; i <= 100; i++) {
    const dateIso = `2026-03-${String(((i - 1) % 28) + 1).padStart(2, '0')}`;
    hundredItems.push(
      makeTx({
        id: `dup100-${i}`,
        excelRowNumber: i + 1,
        dateIso,
        amount: 100_000,
      })
    );
  }

  const t0 = performance.now();
  const res = findCombinations({
    items: hundredItems,
    targetAmount: 1_200_000,
    tolerance: 0,
    minItems: 1,
    maxItems: 12,
    maxResults: 50,
  });
  const elapsedMs = performance.now() - t0;

  assert.equal(res.error, null);
  assert.ok(
    elapsedMs < 500,
    `Tìm trên 100 dòng trùng mệnh giá (C(100,12) ~ 1,05 * 10^15) phải < 500ms, thực tế mất ${elapsedMs.toFixed(2)}ms`
  );
  assert.equal(res.solutions.length, 3);
  assert.equal(res.solutions[0].itemCount, 12);
  assert.equal(res.solutions[0].totalAmount, 1_200_000);
  assert.equal(new Set(res.solutions[0].itemIds).size, 12);
  assert.ok(
    res.solutions[0].duplicateVariantsNote?.includes('nhiều cách chọn tương đương'),
    `C(100,12) vượt ngưỡng 1.000.000 phải hiển thị ghi chú tổng quát, nhận: ${res.solutions[0].duplicateVariantsNote}`
  );
});

test('2.6 [100 Identical Rows Late-Month Cluster] 100 dòng trùng 100.000đ có cụm 12 dòng cuối tháng (offset 88) cùng ngày 25/03/2026 phải đạt dateSpanDays = 0', () => {
  const hundredItems = [];
  // 88 dòng đầu rải đều mỗi ngày tối đa 2 dòng trong tháng 1 & 2 (mọi tập 12 dòng đầu đều có dateSpanDays >= 5)
  for (let i = 1; i <= 88; i++) {
    const dayIndex = Math.floor((i - 1) / 2); // 0..43 ngày tính từ 2026-01-01
    const dt = new Date(Date.UTC(2026, 0, 1 + dayIndex));
    const dateIso = dt.toISOString().slice(0, 10);
    hundredItems.push(
      makeTx({
        id: `late100-${i}`,
        excelRowNumber: i + 1,
        dateIso,
        amount: 100_000,
      })
    );
  }
  // 12 dòng cuối (i = 89..100, offset cửa sổ trượt = 88 > 48) phát sinh CÙNG NGÀY 2026-03-25 (dateSpanDays = 0)
  for (let i = 89; i <= 100; i++) {
    hundredItems.push(
      makeTx({
        id: `late100-${i}`,
        excelRowNumber: i + 1,
        dateIso: '2026-03-25',
        amount: 100_000,
      })
    );
  }

  const res = findCombinations({
    items: hundredItems,
    targetAmount: 1_200_000,
    tolerance: 0,
    minItems: 12,
    maxItems: 12,
  });

  assert.equal(res.error, null);
  assert.ok(res.solutions.length >= 1);
  assert.equal(
    res.solutions[0].dateSpanDays,
    0,
    `Cửa sổ trượt trên 100 dòng trùng mệnh giá phải bắt được cụm 12 hóa đơn cùng ngày ở cuối danh sách (dateSpanDays = 0), thực tế nhận dateSpanDays = ${res.solutions[0].dateSpanDays}`
  );
});

// ============================================================================
// PHẦN 3: CA BIÊN NGẶT NGHÈO CỦA PINNED, EXCLUDED, MIN/MAX ITEMS & DÒNG 0Đ
// ============================================================================

test('3.1 [Pinned Already Matches Target (r = 0)] Kiểm tra khi minItems = 1 so với khi minItems > pinnedIds.length', () => {
  const items = [
    makeTx({ id: 'pin-1', excelRowNumber: 2, dateIso: '2026-03-05', amount: 7_000_000 }),
    makeTx({ id: 'pin-2', excelRowNumber: 3, dateIso: '2026-03-06', amount: 3_000_000 }),
    makeTx({ id: 'cand-pos', excelRowNumber: 4, dateIso: '2026-03-07', amount: 500_000 }),
    makeTx({ id: 'cand-neg', excelRowNumber: 5, dateIso: '2026-03-08', amount: -500_000 }),
    makeTx({ id: 'cand-other', excelRowNumber: 6, dateIso: '2026-03-09', amount: 2_000_000 }),
  ];

  // Ca A: minItems = 1 -> Bản thân 2 dòng ghim (7M + 3M = 10M) đã khớp ngay ở r = 0 (2 dòng), xếp trước nghiệm 4 dòng
  const resMin1 = findCombinations({
    items,
    targetAmount: 10_000_000,
    tolerance: 0,
    pinnedIds: ['pin-1', 'pin-2'],
    minItems: 1,
    maxItems: 6,
  });

  assert.equal(resMin1.error, null);
  assert.ok(resMin1.solutions.length >= 2);
  assert.equal(resMin1.solutions[0].itemCount, 2, 'Nghiệm r = 0 (chỉ gồm 2 dòng ghim) phải đứng đầu');
  assert.deepEqual(resMin1.solutions[0].itemIds, ['pin-1', 'pin-2']);
  assert.equal(resMin1.solutions[0].totalAmount, 10_000_000);

  // Ca B: minItems = 3 (> pinnedIds.length = 2) -> nghiệm r = 0 bị chặn, tìm nghiệm 4 dòng (thêm +500k & -500k)
  const resMin3 = findCombinations({
    items,
    targetAmount: 10_000_000,
    tolerance: 0,
    pinnedIds: ['pin-1', 'pin-2'],
    minItems: 3,
    maxItems: 6,
  });

  assert.equal(resMin3.error, null);
  assert.ok(
    resMin3.solutions.every((s) => s.itemCount >= 3),
    'Khi minItems = 3, nghiệm r = 0 (2 dòng) không được phép xuất hiện'
  );
  assert.equal(resMin3.solutions.length, 1);
  assert.equal(resMin3.solutions[0].itemCount, 4);
  assert.deepEqual(
    [...resMin3.solutions[0].itemIds].sort(),
    ['cand-neg', 'cand-pos', 'pin-1', 'pin-2']
  );

  // Ca C: minItems = 3 & maxItems = 3 -> không tồn tại nghiệm đúng 3 dòng
  const resExact3 = findCombinations({
    items,
    targetAmount: 10_000_000,
    tolerance: 0,
    pinnedIds: ['pin-1', 'pin-2'],
    minItems: 3,
    maxItems: 3,
  });
  assert.equal(resExact3.error, null);
  assert.equal(resExact3.solutions.length, 0);
});

test('3.2 [Pinned Sum > Target & Negative Target] Ghim dòng 20.000.000đ tìm Z = 18.500.000đ (kéo giảm bằng số âm) & tìm mục tiêu âm Z = -1.500.000đ', () => {
  const items = [
    makeTx({ id: 'inv-big', excelRowNumber: 2, dateIso: '2026-03-02', amount: 20_000_000 }),
    makeTx({ id: 'pos-extra', excelRowNumber: 3, dateIso: '2026-03-03', amount: 500_000 }),
    makeTx({ id: 'credit-1500', excelRowNumber: 4, dateIso: '2026-03-04', amount: -1_500_000 }),
    makeTx({ id: 'credit-900', excelRowNumber: 5, dateIso: '2026-03-05', amount: -900_000 }),
    makeTx({ id: 'credit-600', excelRowNumber: 6, dateIso: '2026-03-06', amount: -600_000 }),
    makeTx({ id: 'credit-2000', excelRowNumber: 7, dateIso: '2026-03-07', amount: -2_000_000 }),
  ];

  const res = findCombinations({
    items,
    targetAmount: 18_500_000,
    tolerance: 0,
    pinnedIds: ['inv-big'],
    minItems: 1,
    maxItems: 5,
  });

  assert.equal(res.error, null);
  assert.equal(res.solutions.length, 3, 'Phải tìm đủ cả 3 phương án cấn trừ số âm về đúng 18.500.000đ');
  assert.deepEqual(res.solutions[0].itemIds, ['inv-big', 'credit-1500']);
  assert.equal(res.solutions[0].totalAmount, 18_500_000);

  const comboKeys = res.solutions.map((s) => [...s.itemIds].sort().join(',')).sort();
  assert.deepEqual(comboKeys, [
    'credit-1500,inv-big',
    'credit-2000,inv-big,pos-extra',
    'credit-600,credit-900,inv-big',
  ]);

  // Kiểm tra thêm mục tiêu Z âm (Z = -1.500.000đ: hoàn tiền/trả hàng ròng)
  const negTargetRes = findCombinations({
    items,
    targetAmount: -1_500_000,
    tolerance: 0,
    minItems: 1,
    maxItems: 4,
  });
  assert.equal(negTargetRes.error, null);
  assert.equal(negTargetRes.solutions.length, 3); // {-1.5M}, {-900k, -600k}, {+500k, -2M}
  assert.deepEqual(negTargetRes.solutions[0].itemIds, ['credit-1500']);
});

test('3.3 [Pinned Constraints & Conflict Safety] pinnedIds.length > maxItems, xung đột pinnedIds ∩ excludedIds, trùng ID trong pinnedIds và biên cấu hình lỗi', () => {
  const items = [
    makeTx({ id: 'a1', excelRowNumber: 2, dateIso: '2026-03-01', amount: 1_000_000 }),
    makeTx({ id: 'a2', excelRowNumber: 3, dateIso: '2026-03-02', amount: 2_000_000 }),
    makeTx({ id: 'a3', excelRowNumber: 4, dateIso: '2026-03-03', amount: 3_000_000 }),
    makeTx({ id: 'a4', excelRowNumber: 5, dateIso: '2026-03-04', amount: 4_000_000 }),
  ];

  // 1) pinnedIds.length (4) > maxItems (3) -> trả về 0 nghiệm kèm thông báo lỗi rõ ràng, không crash
  const overMaxRes = findCombinations({
    items,
    targetAmount: 10_000_000,
    pinnedIds: ['a1', 'a2', 'a3', 'a4'],
    maxItems: 3,
  });
  assert.equal(overMaxRes.solutions.length, 0);
  assert.equal(overMaxRes.stats.pinnedCount, 4);
  assert.ok(
    typeof overMaxRes.error === 'string' && overMaxRes.error.length > 0,
    'Khi số dòng ghim vượt maxItems phải trả về chuỗi báo lỗi rõ ràng'
  );

  // 2) Truyền trùng lặp ID trong mảng pinnedIds (['a1', 'a1', 'a2', 'a2']) khi maxItems = 2 -> tự khử trùng thành 2 dòng
  const dupPinInputRes = findCombinations({
    items,
    targetAmount: 3_000_000,
    tolerance: 0,
    pinnedIds: ['a1', 'a1', 'a2', 'a2'],
    maxItems: 2,
  });
  assert.equal(dupPinInputRes.error, null);
  assert.equal(dupPinInputRes.stats.pinnedCount, 2);
  assert.equal(dupPinInputRes.solutions.length, 1);
  assert.deepEqual(dupPinInputRes.solutions[0].itemIds, ['a1', 'a2']);

  // 3) Cùng một ID ('a2') vừa nằm trong pinnedIds vừa nằm trong excludedIds -> không crash, trả 0 nghiệm nhất quán
  const conflictRes = findCombinations({
    items,
    targetAmount: 3_000_000,
    tolerance: 0,
    pinnedIds: ['a1', 'a2'],
    excludedIds: ['a2'],
    maxItems: 4,
  });
  assert.equal(conflictRes.solutions.length, 0, 'Không thể có nghiệm khi dòng bắt buộc ghim đồng thời bị loại trừ');

  // 4) Ghim một ID không tồn tại -> không crash, trả 0 nghiệm
  const missingPinRes = findCombinations({
    items,
    targetAmount: 3_000_000,
    tolerance: 0,
    pinnedIds: ['a1', 'id-khong-ton-tai'],
  });
  assert.equal(missingPinRes.solutions.length, 0);

  // 5) minItems > maxItems & khoảng ngày ngược -> trả error an toàn, không crash
  const badMinMax = findCombinations({ items, targetAmount: 1_000_000, minItems: 5, maxItems: 3 });
  assert.equal(badMinMax.solutions.length, 0);
  assert.ok(typeof badMinMax.error === 'string');

  const badDates = findCombinations({
    items,
    targetAmount: 1_000_000,
    dateFromIso: '2026-03-20',
    dateToIso: '2026-03-05',
  });
  assert.equal(badDates.solutions.length, 0);
  assert.ok(typeof badDates.error === 'string');
});

test('3.4 [Zero-Amount Rows] Dòng amount = 0: không ghim thì không sinh nghiệm rác lặp lại; cố tình ghim thì bắt buộc giữ trong nghiệm', () => {
  const items = [
    makeTx({ id: 'z1', excelRowNumber: 2, dateIso: '2026-03-01', amount: 0, description: 'Bút toán điều chỉnh 0đ #1' }),
    makeTx({ id: 'z2', excelRowNumber: 3, dateIso: '2026-03-02', amount: 0, description: 'Bút toán điều chỉnh 0đ #2' }),
    makeTx({ id: 'p1', excelRowNumber: 4, dateIso: '2026-03-03', amount: 3_000_000 }),
    makeTx({ id: 'p2', excelRowNumber: 5, dateIso: '2026-03-04', amount: 2_000_000 }),
    makeTx({ id: 'n1', excelRowNumber: 6, dateIso: '2026-03-05', amount: -2_000_000 }),
  ];

  // Ca A: Không ghim dòng 0đ -> tìm 5.000.000đ chỉ trả đúng 1 nghiệm {p1, p2}
  const unpinnedRes = findCombinations({
    items,
    targetAmount: 5_000_000,
    tolerance: 0,
    maxItems: 6,
  });
  assert.equal(unpinnedRes.error, null);
  assert.equal(unpinnedRes.solutions.length, 1, 'Dòng 0đ không ghim tuyệt đối không được nhân bản nghiệm rác');
  assert.deepEqual(unpinnedRes.solutions[0].itemIds, ['p1', 'p2']);

  // Kể cả khi targetAmount = 0 và không ghim -> chỉ nhận cặp đối ứng thực {p2, n1}
  const zeroTargetUnpinned = findCombinations({
    items,
    targetAmount: 0,
    tolerance: 0,
    maxItems: 6,
  });
  assert.equal(zeroTargetUnpinned.solutions.length, 1);
  assert.deepEqual(zeroTargetUnpinned.solutions[0].itemIds, ['p2', 'n1']);

  // Ca B: Người dùng CỐ TÌNH ghim dòng z1 (amount = 0) khi tìm 5.000.000đ -> z1 phải có mặt trong nghiệm!
  const pinnedZeroRes = findCombinations({
    items,
    targetAmount: 5_000_000,
    tolerance: 0,
    pinnedIds: ['z1'],
    maxItems: 6,
  });
  assert.equal(pinnedZeroRes.error, null);
  assert.equal(pinnedZeroRes.solutions.length, 1);
  assert.deepEqual(pinnedZeroRes.solutions[0].itemIds, ['z1', 'p1', 'p2']);
  assert.ok(!pinnedZeroRes.solutions[0].itemIds.includes('z2'));

  // Ca C: Người dùng cố tình ghim z1 và tìm targetAmount = 0 -> nghiệm r = 0 gồm [z1] hợp lệ!
  const pinnedZeroAtZeroTarget = findCombinations({
    items,
    targetAmount: 0,
    tolerance: 0,
    pinnedIds: ['z1'],
    minItems: 1,
    maxItems: 6,
  });
  assert.equal(pinnedZeroAtZeroTarget.error, null);
  assert.ok(pinnedZeroAtZeroTarget.solutions.length >= 1);
  assert.deepEqual(pinnedZeroAtZeroTarget.solutions[0].itemIds, ['z1']);
  assert.equal(pinnedZeroAtZeroTarget.solutions[0].totalAmount, 0);
});

test('3.5 [Excel Formula Unrounded Float Precision] 3 dòng chia đều =1.000.000/3 từ công thức Excel phải gom khớp đúng 1.000.000đ (tolerance = 0)', () => {
  const third = 1_000_000 / 3;
  const items = [
    makeTx({ id: 'split-1', excelRowNumber: 2, dateIso: '2026-03-10', amount: third }),
    makeTx({ id: 'split-2', excelRowNumber: 3, dateIso: '2026-03-10', amount: third }),
    makeTx({ id: 'split-3', excelRowNumber: 4, dateIso: '2026-03-10', amount: third }),
  ];

  const res = findCombinations({
    items,
    targetAmount: 1_000_000,
    tolerance: 0,
    minItems: 1,
    maxItems: 5,
  });

  assert.equal(res.error, null);
  assert.equal(
    res.solutions.length,
    1,
    '3 khoản phân bổ =1.000.000/3 có tổng đúng 1.000.000đ phải khớp chính xác ngay cả khi tolerance = 0'
  );
  assert.equal(res.solutions[0].isExact, true);
  assert.equal(res.solutions[0].totalAmount, 1_000_000);
  assert.equal(res.solutions[0].diff, 0);
});

// ============================================================================
// PHẦN 4: KIỂM THỬ MÔ PHỎNG & THỰC THI GIAO THỨC `reconcileWorker.js`
// ============================================================================

test('4.1 [Worker Simulated `self` Protocol] SEARCH liên tiếp, tiền hủy CANCEL, hủy giữa chừng qua PROGRESS, tràn >256 ID và chống crash với dữ liệu lỗi', () => {
  const hadSelf = Object.prototype.hasOwnProperty.call(globalThis, 'self');
  const prevSelf = globalThis.self;

  const outbox = [];
  let onPostMessageHook = null;

  globalThis.self = {
    postMessage(msg) {
      outbox.push(msg);
      if (onPostMessageHook) {
        onPostMessageHook(msg);
      }
    },
  };

  try {
    const sampleItems = [
      makeTx({ id: 'w1', excelRowNumber: 2, dateIso: '2026-03-01', amount: 1_500_000 }),
      makeTx({ id: 'w2', excelRowNumber: 3, dateIso: '2026-03-02', amount: 2_500_000 }),
      makeTx({ id: 'w3', excelRowNumber: 4, dateIso: '2026-03-03', amount: 4_000_000 }),
    ];

    // 1) Gửi các message rác / dị dạng -> không crash hay phát message lạ
    handleWorkerMessage(null);
    handleWorkerMessage(undefined);
    handleWorkerMessage('chuoi-rac');
    handleWorkerMessage(12345);
    handleWorkerMessage({ data: null });
    handleWorkerMessage({ data: 'khong-phai-object' });
    handleWorkerMessage({ data: { type: 'UNKNOWN_CMD', requestId: 999 } });
    handleWorkerMessage({ data: { type: 'CANCEL' } });
    assert.equal(outbox.length, 0);

    // 2) Gửi liên tiếp nhiều SEARCH với requestId tăng dần (bao gồm cả requestId = 0)
    for (const reqId of [0, 1, 2, 3]) {
      handleWorkerMessage({
        data: {
          type: 'SEARCH',
          requestId: reqId,
          options: {
            items: sampleItems,
            targetAmount: 4_000_000,
            tolerance: 0,
          },
        },
      });
    }

    const resultMsgs = outbox.filter((m) => m.type === 'RESULT');
    assert.deepEqual(
      resultMsgs.map((m) => m.requestId),
      [0, 1, 2, 3],
      'Phải xử lý chính xác cả requestId = 0 và chuỗi requestId tăng dần'
    );
    for (const rm of resultMsgs) {
      assert.equal(rm.result.solutions.length, 2);
    }

    // 3) Gửi CANCEL trước rồi mới gửi SEARCH cùng requestId -> bỏ qua không chạy
    outbox.length = 0;
    handleWorkerMessage({ data: { type: 'CANCEL', requestId: 50 } });
    handleWorkerMessage({
      data: {
        type: 'SEARCH',
        requestId: 50,
        options: { items: sampleItems, targetAmount: 4_000_000, tolerance: 0 },
      },
    });
    assert.equal(outbox.length, 0, 'Request đã bị CANCEL trước không được phát sinh PROGRESS hay RESULT');

    // 4) Hủy giữa chừng (Mid-flight CANCEL) khi nhận sự kiện PROGRESS (không truyền requestId -> hủy theo activeRequestId)
    const heavyItems = [];
    for (let i = 1; i <= 24; i++) {
      heavyItems.push(
        makeTx({
          id: `hv-pos-${i}`,
          excelRowNumber: i + 1,
          dateIso: '2026-03-10',
          amount: 2_000_000 + i * 2_000,
        })
      );
    }
    for (let j = 1; j <= 8; j++) {
      heavyItems.push(
        makeTx({
          id: `hv-neg-${j}`,
          excelRowNumber: 25 + j,
          dateIso: '2026-03-10',
          amount: -(2_000_000 + j * 2_000),
        })
      );
    }

    outbox.length = 0;
    onPostMessageHook = (msg) => {
      if (msg.type === 'PROGRESS' && msg.requestId === 60) {
        handleWorkerMessage({ data: { type: 'CANCEL' } });
      }
    };

    handleWorkerMessage({
      data: {
        type: 'SEARCH',
        requestId: 60,
        options: {
          items: heavyItems,
          targetAmount: 15_000_600,
          tolerance: 500,
          maxItems: 12,
          timeLimitMs: 3000,
        },
      },
    });
    onPostMessageHook = null;

    assert.ok(
      outbox.some((m) => m.type === 'PROGRESS' && m.requestId === 60),
      'Phải phát ít nhất 1 bản tin PROGRESS trước khi bị CANCEL'
    );
    assert.ok(
      !outbox.some((m) => m.type === 'RESULT' && m.requestId === 60),
      'Sau khi bị CANCEL giữa chừng, tuyệt đối không được phát RESULT'
    );

    // 5) Kiểm tra giới hạn bộ nhớ cancelledRequestIds (> 256 ID không rò rỉ hay crash)
    for (let id = 1000; id < 1350; id++) {
      handleWorkerMessage({ data: { type: 'CANCEL', requestId: id } });
    }

    // 6) Input lỗi cấu hình hợp lệ (trả RESULT kèm result.error) vs Exception bất ngờ (trả ERROR message)
    outbox.length = 0;
    handleWorkerMessage({
      data: {
        type: 'SEARCH',
        requestId: 70,
        options: { items: 'khong-phai-mang', targetAmount: NaN },
      },
    });
    assert.equal(outbox.length, 1);
    assert.equal(outbox[0].type, 'RESULT');
    assert.equal(outbox[0].requestId, 70);
    assert.ok(typeof outbox[0].result.error === 'string');

    outbox.length = 0;
    const throwingOptions = {
      items: sampleItems,
      get targetAmount() {
        throw new Error('Lỗi đọc bộ nhớ giả lập');
      },
    };
    handleWorkerMessage({
      data: {
        type: 'SEARCH',
        requestId: 71,
        options: throwingOptions,
      },
    });
    assert.equal(outbox.length, 1);
    assert.equal(outbox[0].type, 'ERROR');
    assert.equal(outbox[0].requestId, 71);
    assert.match(outbox[0].message, /Lỗi đọc bộ nhớ giả lập/);

    // Đảm bảo sau khi văng lỗi ở request 71, Worker vẫn phục hồi hoàn toàn cho request 72 (gồm PROGRESS cuối + RESULT)
    outbox.length = 0;
    handleWorkerMessage({
      data: {
        type: 'SEARCH',
        requestId: 72,
        options: { items: sampleItems, targetAmount: 4_000_000, tolerance: 0 },
      },
    });
    const res72 = outbox.filter((m) => m.type === 'RESULT' && m.requestId === 72);
    assert.equal(res72.length, 1);
    assert.equal(res72[0].result.solutions.length, 2);
  } finally {
    if (hadSelf) {
      globalThis.self = prevSelf;
    } else {
      delete globalThis.self;
    }
  }
});

test('4.2 [Real Node `worker_threads` End-to-End] Khởi chạy reconcileWorker.js trên luồng thực tế: chuỗi SEARCH tăng dần, tiền hủy CANCEL và xử lý input lỗi', async () => {
  const bootstrapCode = `
    import { parentPort } from 'node:worker_threads';
    globalThis.self = {
      postMessage(msg) {
        parentPort.postMessage(msg);
      },
      onmessage: null,
    };
    await import(${JSON.stringify(WORKER_MODULE_URL)});
    parentPort.on('message', (data) => {
      if (typeof globalThis.self.onmessage === 'function') {
        globalThis.self.onmessage({ data });
      }
    });
    parentPort.postMessage({ type: 'READY' });
  `;

  const worker = new NodeWorker(bootstrapCode, { eval: true, inputType: 'module' });
  const received = [];

  try {
    await new Promise((resolve, reject) => {
      const onMsg = (msg) => {
        if (msg && msg.type === 'READY') {
          worker.off('message', onMsg);
          resolve();
        }
      };
      worker.on('message', onMsg);
      worker.once('error', reject);
    });

    worker.on('message', (msg) => {
      received.push(msg);
    });

    const items = [
      makeTx({ id: 't1', excelRowNumber: 2, dateIso: '2026-03-10', amount: 3_200_000 }),
      makeTx({ id: 't2', excelRowNumber: 3, dateIso: '2026-03-10', amount: 1_800_000 }),
      makeTx({ id: 't3', excelRowNumber: 4, dateIso: '2026-03-11', amount: -500_000 }),
    ];

    const donePromise = new Promise((resolve, reject) => {
      const timer = setTimeout(() => reject(new Error('Timeout waiting for requestId=104')), 4000);
      worker.on('message', (msg) => {
        if (msg && msg.type === 'RESULT' && msg.requestId === 104) {
          clearTimeout(timer);
          resolve(msg);
        }
      });
      worker.once('error', (err) => {
        clearTimeout(timer);
        reject(err);
      });
    });

    // Tiền hủy requestId = 102 trước khi gửi SEARCH 102
    worker.postMessage({ type: 'CANCEL', requestId: 102 });
    worker.postMessage({ type: 'SEARCH', requestId: 101, options: { items, targetAmount: 5_000_000, tolerance: 0 } });
    worker.postMessage({ type: 'SEARCH', requestId: 102, options: { items, targetAmount: 5_000_000, tolerance: 0 } });
    worker.postMessage({ type: 'SEARCH', requestId: 103, options: null }); // Input lỗi
    worker.postMessage({ type: 'SEARCH', requestId: 104, options: { items, targetAmount: 4_500_000, tolerance: 0 } });

    await donePromise;

    const resultsById = new Map(
      received.filter((m) => m.type === 'RESULT').map((m) => [m.requestId, m.result])
    );

    assert.ok(resultsById.has(101), 'Phải có kết quả cho request 101');
    assert.equal(resultsById.get(101).solutions.length, 1);
    assert.deepEqual(resultsById.get(101).solutions[0].itemIds, ['t1', 't2']);

    assert.ok(!resultsById.has(102), 'Request 102 đã bị CANCEL trước không được trả RESULT');

    assert.ok(resultsById.has(103), 'Request 103 input null phải trả RESULT báo lỗi an toàn');
    assert.ok(typeof resultsById.get(103).error === 'string');

    assert.ok(resultsById.has(104), 'Worker phải tiếp tục xử lý bình thường cho request 104');
    assert.deepEqual(resultsById.get(104).solutions[0].itemIds, ['t1', 't2', 't3']);
  } finally {
    await worker.terminate();
  }
});

test('4.3 [Worker Rapid Burst Queue Cancellation] Khi dồn hàng đợi SEARCH #1 -> CANCEL #1 -> SEARCH #2 -> CANCEL #2 -> SEARCH #3 lúc Worker đang bận, không được chạy oan SEARCH #2 đã hủy', async () => {
  const bootstrapCode = `
    import { parentPort } from 'node:worker_threads';
    globalThis.self = {
      postMessage(msg) {
        parentPort.postMessage(msg);
      },
      onmessage: null,
    };
    await import(${JSON.stringify(WORKER_MODULE_URL)});
    parentPort.on('message', (data) => {
      if (typeof globalThis.self.onmessage === 'function') {
        globalThis.self.onmessage({ data });
      }
    });
    parentPort.postMessage({ type: 'READY' });
  `;

  const worker = new NodeWorker(bootstrapCode, { eval: true, inputType: 'module' });
  const received = [];

  try {
    await new Promise((resolve, reject) => {
      const onMsg = (msg) => {
        if (msg && msg.type === 'READY') {
          worker.off('message', onMsg);
          resolve();
        }
      };
      worker.on('message', onMsg);
      worker.once('error', reject);
    });

    worker.on('message', (msg) => {
      received.push(msg);
    });

    const slowUnreachableItems = [];
    for (let i = 1; i <= 20; i++) {
      slowUnreachableItems.push(
        makeTx({ id: `wk-p-${i}`, excelRowNumber: i + 1, dateIso: '2026-03-10', amount: 2_000_000 + i * 2_000 })
      );
    }
    for (let j = 1; j <= 10; j++) {
      slowUnreachableItems.push(
        makeTx({ id: `wk-n-${j}`, excelRowNumber: 21 + j, dateIso: '2026-03-10', amount: -(2_000_000 + j * 2_000) })
      );
    }

    const fastItems = [
      makeTx({ id: 'wk-fast-1', excelRowNumber: 2, dateIso: '2026-03-15', amount: 3_200_000 }),
      makeTx({ id: 'wk-fast-2', excelRowNumber: 3, dateIso: '2026-03-15', amount: 1_800_000 }),
    ];

    const finalResultPromise = new Promise((resolve, reject) => {
      const timer = setTimeout(() => reject(new Error('Worker timed out waiting for requestId=3')), 4000);
      worker.on('message', (msg) => {
        if (msg && msg.type === 'RESULT' && msg.requestId === 3) {
          clearTimeout(timer);
          resolve(msg);
        }
      });
      worker.once('error', (err) => {
        clearTimeout(timer);
        reject(err);
      });
    });

    worker.postMessage({
      type: 'SEARCH',
      requestId: 1,
      options: { items: slowUnreachableItems, targetAmount: 15_000_600, tolerance: 500, timeLimitMs: 120 },
    });
    worker.postMessage({ type: 'CANCEL', requestId: 1 });

    worker.postMessage({
      type: 'SEARCH',
      requestId: 2,
      options: { items: slowUnreachableItems, targetAmount: 15_000_600, tolerance: 500, timeLimitMs: 150 },
    });
    worker.postMessage({ type: 'CANCEL', requestId: 2 });

    worker.postMessage({
      type: 'SEARCH',
      requestId: 3,
      options: { items: fastItems, targetAmount: 5_000_000, tolerance: 0, timeLimitMs: 1000 },
    });

    const msg3 = await finalResultPromise;
    assert.equal(msg3.result.error, null);
    assert.equal(msg3.result.solutions.length, 1);

    const emittedResultForCancelledReq2 = received.some(
      (m) => m.type === 'RESULT' && m.requestId === 2
    );
    assert.equal(
      emittedResultForCancelledReq2,
      false,
      'Request #2 đã có lệnh CANCEL nằm sẵn trong hàng đợi trước khi chạy thì không được thực thi đồng bộ gây nghẽn hàng đợi và phát RESULT #2'
    );
  } finally {
    await worker.terminate();
  }
});
