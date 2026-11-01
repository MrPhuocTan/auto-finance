import test from 'node:test';
import assert from 'node:assert/strict';

import { findCombinations, normalizeVietnameseSearch } from '../src/core/reconcileEngine.js';
import {
  DEFAULT_TOLERANCE,
  DEFAULT_MAX_ITEMS,
  DEFAULT_MAX_RESULTS,
} from '../src/core/contracts.js';

const MS_PER_DAY = 86_400_000;
const ORACLE_SCALE_DECIMALS = 6; // 10^6 exact fixed-point scale in BigInt (supports up to 6 decimals & unlimited magnitude)
const ORACLE_SCALE_BIGINT = 1_000_000n;

/**
 * Helper tạo dòng giao dịch chuẩn theo hợp đồng dữ liệu `TransactionItem` trong `contracts.js`.
 */
function makeTx(
  id,
  amount,
  {
    excelRowNumber = 2,
    dateIso = '2026-03-01',
    description = `Nghiệp vụ ${id}`,
    reference = `CT-${id}`,
    isSummaryRow = false,
    extraCells = {},
  } = {}
) {
  let timestamp = null;
  let dateDisplay = '';
  if (typeof dateIso === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(dateIso)) {
    const [y, m, d] = dateIso.split('-').map(Number);
    timestamp = Date.UTC(y, m - 1, d);
    dateDisplay = `${String(d).padStart(2, '0')}/${String(m).padStart(2, '0')}/${y}`;
  }
  return {
    id: String(id),
    excelRowNumber,
    dateIso,
    dateDisplay,
    timestamp,
    amount,
    description,
    reference,
    isSummaryRow,
    extraCells,
    rawCells: [dateDisplay, reference, description, amount],
  };
}

/**
 * Bộ sinh số giả ngẫu nhiên có seed cố định (Mulberry32) đảm bảo tái lập 100% mọi kịch bản fuzzing.
 */
function createSeededRng(seed) {
  let state = seed >>> 0;
  return {
    nextFloat() {
      state = (state + 0x6d2b79f5) | 0;
      let t = Math.imul(state ^ (state >>> 15), 1 | state);
      t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    },
    nextInt(minInclusive, maxInclusive) {
      const span = maxInclusive - minInclusive + 1;
      return minInclusive + Math.floor(this.nextFloat() * span);
    },
    pick(arr) {
      return arr[this.nextInt(0, arr.length - 1)];
    },
    shuffle(arr) {
      const copy = arr.slice();
      for (let i = copy.length - 1; i > 0; i--) {
        const j = this.nextInt(0, i);
        const tmp = copy[i];
        copy[i] = copy[j];
        copy[j] = tmp;
      }
      return copy;
    },
  };
}

/**
 * Chuyển đổi chính xác một số thực JS sang BigInt đã nhân 10^scaleDecimals bằng phân tích chuỗi thập phân
 * (dựa trên biểu diễn ngắn nhất chuẩn Ryū của ECMAScript, KHÔNG dùng phép nhân float `num * 10^k`).
 */
function numberToScaledBigInt(num, scaleDecimals = ORACLE_SCALE_DECIMALS) {
  if (!Number.isFinite(num) || num === 0) return 0n;
  const raw = String(num).trim().toLowerCase();
  let sign = 1n;
  let body = raw;
  if (body.startsWith('-')) {
    sign = -1n;
    body = body.slice(1);
  } else if (body.startsWith('+')) {
    body = body.slice(1);
  }

  let mantissa = body;
  let exp = 0;
  if (body.includes('e')) {
    const [m, e] = body.split('e');
    mantissa = m;
    exp = parseInt(e, 10);
  }

  const [intStr, fracStr = ''] = mantissa.split('.');
  const digits = (intStr + fracStr).replace(/^0+/, '') || '0';
  if (digits === '0') return 0n;

  const totalShift = exp - fracStr.length + scaleDecimals;
  if (totalShift >= 0) {
    return sign * BigInt(digits + '0'.repeat(totalShift));
  }

  // Làm tròn half-away-from-zero nếu vượt quá scaleDecimals
  const cut = -totalShift;
  if (cut > digits.length) {
    return 0n;
  }
  const head = digits.slice(0, digits.length - cut) || '0';
  const nextDigit = Number(digits[digits.length - cut] || '0');
  let val = BigInt(head);
  if (nextDigit >= 5) val += 1n;
  return sign * val;
}

function scaledBigIntToNumber(bigVal, scaleDecimals = ORACLE_SCALE_DECIMALS) {
  if (bigVal === 0n) return 0;
  const neg = bigVal < 0n;
  const absStr = (neg ? -bigVal : bigVal).toString();
  let outStr;
  if (scaleDecimals === 0) {
    outStr = absStr;
  } else if (absStr.length <= scaleDecimals) {
    const padded = absStr.padStart(scaleDecimals, '0').replace(/0+$/, '');
    outStr = padded ? `0.${padded}` : '0';
  } else {
    const intPart = absStr.slice(0, absStr.length - scaleDecimals);
    const fracPart = absStr.slice(absStr.length - scaleDecimals).replace(/0+$/, '');
    outStr = fracPart ? `${intPart}.${fracPart}` : intPart;
  }
  if (outStr === '0') return 0;
  return Number(neg ? `-${outStr}` : outStr);
}

/**
 * Hàm Vét Cạn Độc Lập dùng `BigInt` (`oracleBruteForce`) để làm chuẩn đối chiếu toán học (Ground-Truth Oracle).
 * Duyệt toàn bộ 2^M tập con của tập ứng viên với số học nguyên lớn chính xác tuyệt đối.
 */
function oracleBruteForce(options) {
  const {
    items = [],
    dateFromIso = null,
    dateToIso = null,
    keyword = '',
    targetAmount,
    tolerance = DEFAULT_TOLERANCE,
    pinnedIds = [],
    excludedIds = [],
    includeSummaryRows = false,
    minItems = 1,
    maxItems = DEFAULT_MAX_ITEMS,
    maxResults = DEFAULT_MAX_RESULTS,
  } = options;

  const safeTolerance = Number.isFinite(tolerance) ? Math.max(0, tolerance) : DEFAULT_TOLERANCE;
  const safeMinItems = Number.isFinite(minItems) ? Math.max(1, Math.floor(minItems)) : 1;
  const safeMaxItems = Number.isFinite(maxItems) ? Math.floor(maxItems) : DEFAULT_MAX_ITEMS;
  const safeMaxResults = Number.isFinite(maxResults) ? Math.max(1, Math.floor(maxResults)) : DEFAULT_MAX_RESULTS;

  if (safeMaxItems < 1 || safeMinItems > safeMaxItems) {
    return { allValidSolutions: [], solutions: [] };
  }

  const fromIso = typeof dateFromIso === 'string' && dateFromIso.trim() ? dateFromIso.trim() : null;
  const toIso = typeof dateToIso === 'string' && dateToIso.trim() ? dateToIso.trim() : null;
  if (fromIso && toIso && fromIso > toIso) {
    return { allValidSolutions: [], solutions: [] };
  }

  const normKw = normalizeVietnameseSearch(keyword);
  const excludedSet = new Set((excludedIds || []).filter((x) => x !== null && x !== undefined && x !== ''));
  const uniquePinned = [...new Set((pinnedIds || []).filter((x) => x !== null && x !== undefined && x !== ''))];
  const pinnedSet = new Set(uniquePinned);

  const pinnedItems = [];
  const seenPinned = new Set();
  const candidates = [];

  const getHaystack = (it) => {
    const parts = [];
    if (it.description) parts.push(String(it.description));
    if (it.reference) parts.push(String(it.reference));
    if (Array.isArray(it.extraCells)) {
      for (const e of it.extraCells) {
        if (e && typeof e === 'object') {
          for (const v of Object.values(e)) {
            if (v !== null && v !== undefined && v !== '') parts.push(String(v));
          }
        } else if (e !== null && e !== undefined && e !== '') {
          parts.push(String(e));
        }
      }
    } else if (it.extraCells && typeof it.extraCells === 'object') {
      for (const v of Object.values(it.extraCells)) {
        if (v !== null && v !== undefined && v !== '') parts.push(String(v));
      }
    }
    return normalizeVietnameseSearch(parts.join('\n'));
  };

  for (const item of items) {
    if (!item || typeof item.amount !== 'number' || !Number.isFinite(item.amount)) continue;
    if (excludedSet.has(item.id)) continue;

    if (pinnedSet.has(item.id)) {
      if (!seenPinned.has(item.id)) {
        seenPinned.add(item.id);
        pinnedItems.push(item);
      }
      continue;
    }

    if (item.isSummaryRow === true && !includeSummaryRows) continue;
    if (fromIso && (!item.dateIso || item.dateIso < fromIso)) continue;
    if (toIso && (!item.dateIso || item.dateIso > toIso)) continue;
    if (normKw && !getHaystack(item).includes(normKw)) continue;
    if (item.amount === 0) continue;

    candidates.push(item);
  }

  if (pinnedItems.length < uniquePinned.length || pinnedItems.length > safeMaxItems) {
    return { allValidSolutions: [], solutions: [] };
  }

  const bigTarget = numberToScaledBigInt(targetAmount);
  const bigTol = numberToScaledBigInt(safeTolerance);
  const pinnedBigSum = pinnedItems.reduce((acc, it) => acc + numberToScaledBigInt(it.amount), 0n);
  const candBigAmounts = candidates.map((it) => numberToScaledBigInt(it.amount));

  const getDayTs = (it) => {
    if (Number.isFinite(it.timestamp)) return it.timestamp;
    if (typeof it.dateIso === 'string') {
      const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(it.dateIso.trim());
      if (m) return Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    }
    return null;
  };

  const sortItemsChrono = (a, b) => {
    const ta = getDayTs(a);
    const tb = getDayTs(b);
    if (ta !== null && tb !== null && ta !== tb) return ta - tb;
    if ((ta !== null) !== (tb !== null)) return ta !== null ? -1 : 1;
    const ra = Number.isFinite(a.excelRowNumber) ? a.excelRowNumber : Number.MAX_SAFE_INTEGER;
    const rb = Number.isFinite(b.excelRowNumber) ? b.excelRowNumber : Number.MAX_SAFE_INTEGER;
    if (ra !== rb) return ra - rb;
    return String(a.id ?? '').localeCompare(String(b.id ?? ''));
  };

  const calcSpanDays = (arr) => {
    let minT = Infinity;
    let maxT = -Infinity;
    let cnt = 0;
    for (const it of arr) {
      const t = getDayTs(it);
      if (t !== null) {
        if (t < minT) minT = t;
        if (t > maxT) maxT = t;
        cnt++;
      }
    }
    return cnt <= 1 ? 0 : Math.max(0, Math.round((maxT - minT) / MS_PER_DAY));
  };

  const m = candidates.length;
  const totalMasks = 1 << m;
  const validSolutions = [];

  for (let mask = 0; mask < totalMasks; mask++) {
    // popcount nhanh
    let bits = mask;
    let chosenCount = 0;
    while (bits) {
      bits &= bits - 1;
      chosenCount++;
    }

    const totalItemCount = pinnedItems.length + chosenCount;
    if (totalItemCount < safeMinItems || totalItemCount > safeMaxItems) continue;

    let bigSum = pinnedBigSum;
    for (let i = 0; i < m; i++) {
      if ((mask & (1 << i)) !== 0) {
        bigSum += candBigAmounts[i];
      }
    }

    const bigDiff = bigSum - bigTarget;
    const bigAbsDiff = bigDiff < 0n ? -bigDiff : bigDiff;
    if (bigAbsDiff > bigTol) continue;

    const chosenItems = pinnedItems.slice();
    for (let i = 0; i < m; i++) {
      if ((mask & (1 << i)) !== 0) {
        chosenItems.push(candidates[i]);
      }
    }

    const sortedItems = chosenItems.sort(sortItemsChrono);
    const itemIds = sortedItems.map((it) => it.id);

    validSolutions.push({
      id: `sol:${itemIds.join(',')}`,
      canonicalKey: itemIds.slice().sort().join('\u0001'),
      items: sortedItems,
      itemIds,
      itemCount: sortedItems.length,
      bigSum,
      bigDiff,
      bigAbsDiff,
      totalAmount: scaledBigIntToNumber(bigSum),
      targetAmount,
      diff: scaledBigIntToNumber(bigDiff),
      absDiff: scaledBigIntToNumber(bigAbsDiff),
      isExact: bigDiff === 0n,
      dateSpanDays: calcSpanDays(sortedItems),
    });
  }

  validSolutions.sort((a, b) => {
    if (a.bigAbsDiff !== b.bigAbsDiff) return a.bigAbsDiff < b.bigAbsDiff ? -1 : 1;
    if (a.itemCount !== b.itemCount) return a.itemCount - b.itemCount;
    if (a.dateSpanDays !== b.dateSpanDays) return a.dateSpanDays - b.dateSpanDays;
    const len = Math.min(a.items.length, b.items.length);
    for (let i = 0; i < len; i++) {
      const ra = Number.isFinite(a.items[i].excelRowNumber) ? a.items[i].excelRowNumber : Number.MAX_SAFE_INTEGER;
      const rb = Number.isFinite(b.items[i].excelRowNumber) ? b.items[i].excelRowNumber : Number.MAX_SAFE_INTEGER;
      if (ra !== rb) return ra - rb;
    }
    return a.id.localeCompare(b.id);
  });

  return {
    allValidSolutions: validSolutions,
    solutions: validSolutions.slice(0, safeMaxResults),
  };
}

// ============================================================================
// 1. BẪY DẤU PHẨY ĐỘNG IEEE-754 & SỐ LẺ THẬP PHÂN (VAT 8%, 10%, USD/EUR)
// ============================================================================

test('1.1 IEEE-754 classic 2-decimal traps: 0.1+0.2=0.3, 19.99+0.01=20, 33.33+33.33+33.34=100, 123456.78+876543.22=1000000', () => {
  // Ca 1: 0.1 + 0.2 = 0.3 (trong JS thuần: 0.1 + 0.2 === 0.30000000000000004 !== 0.3)
  assert.notEqual(0.1 + 0.2, 0.3, 'Sanity check: raw IEEE-754 0.1 + 0.2 !== 0.3');
  const res01 = findCombinations({
    items: [
      makeTx('a', 0.1, { excelRowNumber: 2 }),
      makeTx('b', 0.2, { excelRowNumber: 3 }),
      makeTx('noise', 0.15, { excelRowNumber: 4 }),
    ],
    targetAmount: 0.3,
    tolerance: 0,
  });
  assert.equal(res01.solutions.length, 1);
  assert.equal(res01.solutions[0].isExact, true);
  assert.ok(Object.is(res01.solutions[0].diff, 0), 'diff must be +0 (not -0 or 4e-17)');
  assert.ok(Object.is(res01.solutions[0].absDiff, 0));
  assert.equal(res01.solutions[0].totalAmount, 0.3);

  // Ca 2: 0.1 + 0.2 + 0.3 = 0.6 (trong JS thuần: 0.6000000000000001)
  const res06 = findCombinations({
    items: [makeTx('x1', 0.1), makeTx('x2', 0.2), makeTx('x3', 0.3)],
    targetAmount: 0.6,
    tolerance: 0,
  });
  assert.equal(res06.solutions.length, 1);
  assert.equal(res06.solutions[0].isExact, true);
  assert.equal(res06.solutions[0].diff, 0);

  // Ca 3: 19.99 + 0.01 = 20.00
  const res20 = findCombinations({
    items: [makeTx('u1', 19.99), makeTx('u2', 0.01), makeTx('u3', 0.02)],
    targetAmount: 20.0,
    tolerance: 0,
  });
  assert.equal(res20.solutions.length, 1);
  assert.deepEqual(res20.solutions[0].itemIds, ['u1', 'u2']);
  assert.equal(res20.solutions[0].isExact, true);
  assert.equal(res20.solutions[0].diff, 0);
  assert.equal(res20.solutions[0].totalAmount, 20);

  // Ca 4: Chia 3 hóa đơn lẻ xu 33.33 + 33.33 + 33.34 = 100
  const res100 = findCombinations({
    items: [
      makeTx('s1', 33.33, { excelRowNumber: 2 }),
      makeTx('s2', 33.33, { excelRowNumber: 3 }),
      makeTx('s3', 33.34, { excelRowNumber: 4 }),
      makeTx('s4', 33.32, { excelRowNumber: 5 }),
    ],
    targetAmount: 100,
    tolerance: 0,
  });
  assert.equal(res100.solutions.length, 1);
  assert.deepEqual(res100.solutions[0].itemIds, ['s1', 's2', 's3']);
  assert.equal(res100.solutions[0].isExact, true);
  assert.equal(res100.solutions[0].diff, 0);
  assert.equal(res100.solutions[0].totalAmount, 100);

  // Ca 5: Hóa đơn ngoại tệ / VAT lớn lẻ xu 123456.78 + 876543.22 = 1000000
  const res1M = findCombinations({
    items: [
      makeTx('fx1', 123456.78, { excelRowNumber: 2 }),
      makeTx('fx2', 876543.22, { excelRowNumber: 3 }),
      makeTx('fx3', -500.55, { excelRowNumber: 4 }),
    ],
    targetAmount: 1000000,
    tolerance: 0,
  });
  assert.equal(res1M.solutions.length, 1);
  assert.deepEqual(res1M.solutions[0].itemIds, ['fx1', 'fx2']);
  assert.equal(res1M.solutions[0].isExact, true);
  assert.equal(res1M.solutions[0].diff, 0);
  assert.equal(res1M.solutions[0].totalAmount, 1000000);

  // Ca 6: Bù trừ thập phân âm-dương bị drift trong IEEE-754 (1000.1 + 2000.2 === 3000.2999999999997)
  const resDrift = findCombinations({
    items: [
      makeTx('d1', 1000.1, { excelRowNumber: 2 }),
      makeTx('d2', 2000.2, { excelRowNumber: 3 }),
      makeTx('d3', -500.3, { excelRowNumber: 4 }),
    ],
    targetAmount: 2500,
    tolerance: 0,
  });
  assert.equal(resDrift.solutions.length, 1);
  assert.equal(resDrift.solutions[0].isExact, true);
  assert.equal(resDrift.solutions[0].diff, 0);
  assert.equal(resDrift.solutions[0].totalAmount, 2500);
});

test('1.2 VAT 8%/10% & FX 3-4 decimal places: 10.125 + 20.375 = 30.5, VAT 8% 3-decimal sum, 4-decimal FX, +/-10.125 cancellation, and sub-cent items', async (t) => {
  await t.test('1.2a Exact 3-decimal sum: 10.125 + 20.375 = 30.5 (tolerance = 0)', () => {
    const res3Dec = findCombinations({
      items: [
        makeTx('v1', 10.125, { excelRowNumber: 2 }),
        makeTx('v2', 20.375, { excelRowNumber: 3 }),
      ],
      targetAmount: 30.5,
      tolerance: 0,
    });
    assert.equal(
      res3Dec.solutions.length,
      1,
      '10.125 + 20.375 === 30.5 (exact in both decimal and binary IEEE-754) MUST be found at tolerance=0'
    );
    assert.equal(res3Dec.solutions[0].isExact, true);
    assert.equal(res3Dec.solutions[0].diff, 0);
    assert.equal(res3Dec.solutions[0].totalAmount, 30.5);
  });

  await t.test('1.2b VAT 8% 3-decimal invoices: (100.05 + 100.05 + 299.90) * 1.08 = 540.00', () => {
    const resVat8 = findCombinations({
      items: [
        makeTx('vat1', 108.054, { excelRowNumber: 2 }),
        makeTx('vat2', 108.054, { excelRowNumber: 3 }),
        makeTx('vat3', 323.892, { excelRowNumber: 4 }),
      ],
      targetAmount: 540,
      tolerance: 0,
    });
    assert.equal(
      resVat8.solutions.length,
      1,
      'VAT 8% invoices 108.054 + 108.054 + 323.892 === 540.00 MUST be found at tolerance=0'
    );
    assert.equal(resVat8.solutions[0].isExact, true);
    assert.equal(resVat8.solutions[0].diff, 0);
  });

  await t.test('1.2c 4-decimal FX / unit-price sum: 12.3456 + 12.3456 + 75.3088 = 100', () => {
    const res4Dec = findCombinations({
      items: [
        makeTx('fx4_1', 12.3456, { excelRowNumber: 2 }),
        makeTx('fx4_2', 12.3456, { excelRowNumber: 3 }),
        makeTx('fx4_3', 75.3088, { excelRowNumber: 4 }),
      ],
      targetAmount: 100,
      tolerance: 0,
    });
    assert.equal(
      res4Dec.solutions.length,
      1,
      '4-decimal FX sum 12.3456 + 12.3456 + 75.3088 === 100 MUST be found at tolerance=0'
    );
    assert.equal(res4Dec.solutions[0].isExact, true);
    assert.equal(res4Dec.solutions[0].diff, 0);
  });

  await t.test('1.2d Asymmetric Math.round trap: +10.125 and -10.125 exact zero cancellation', () => {
    const resCancel3Dec = findCombinations({
      items: [
        makeTx('inv', 10.125, { excelRowNumber: 2 }),
        makeTx('adj', -10.125, { excelRowNumber: 3 }),
      ],
      targetAmount: 0,
      tolerance: 0,
    });
    assert.equal(
      resCancel3Dec.solutions.length,
      1,
      '+10.125 and -10.125 MUST cancel out to 0 with tolerance=0 (Math.round(1012.5)=1013 vs Math.round(-1012.5)=-1012 bug!)'
    );
    assert.equal(resCancel3Dec.solutions[0].isExact, true);
    assert.equal(resCancel3Dec.solutions[0].diff, 0);
  });

  await t.test('1.2e Sub-cent FX rounding adjustment (< 0.005) must NOT be silently dropped as intAmount===0', () => {
    const resSubCent = findCombinations({
      items: [
        makeTx('main', 19.992, { excelRowNumber: 2 }),
        makeTx('adj1', 0.004, { excelRowNumber: 3 }),
        makeTx('adj2', 0.004, { excelRowNumber: 4 }),
      ],
      targetAmount: 20,
      tolerance: 0,
    });
    assert.equal(
      resSubCent.solutions.length,
      1,
      '19.992 + 0.004 + 0.004 === 20 MUST NOT drop 0.004 items via Math.round(0.004 * 100) === 0'
    );
    assert.deepEqual(resSubCent.solutions[0].itemIds, ['main', 'adj1', 'adj2']);
  });
});

test('1.3 Decimal tolerance & isExact boundary: tolerance=0.05 accepts diff=-0.05, 0, +0.05 and rejects -0.06, +0.06', () => {
  // Kiểm chứng bẫy IEEE-754 ở biên tolerance thập phân:
  // Trong JS thuần: (100.02 + 150.03) - 250 === 0.05000000000001137 > 0.05 (so sánh float ngây thơ sẽ loại nhầm biên +0.05!)
  assert.ok(
    100.02 + 150.03 - 250 > 0.05,
    'Sanity check: raw IEEE-754 (100.02 + 150.03) - 250 overshoots 0.05'
  );

  // Dùng các cặp có cơ sở cách biệt lũy thừa 3 (1000, 3000, 9000, 27000, 81000) để chỉ ghép đúng từng cặp 2 dòng
  const items = [
    // Cặp 0 (IEEE-754 boundary +0.05): 100.02 + 150.03 = 250.05 -> diff = +0.05 (PHẢI NHẬN)
    makeTx('ieee_p05_a', 100.02, { excelRowNumber: 2 }),
    makeTx('ieee_p05_b', 150.03, { excelRowNumber: 3 }),
    // Cặp 1: diff = 0 (3100.10 + (-2850.10) = 250.00 -> Exact)
    makeTx('ex_a', 3100.1, { excelRowNumber: 4 }),
    makeTx('ex_b', -2850.1, { excelRowNumber: 5 }),
    // Cặp 2: diff = -0.05 (9100.10 + (-8850.15) = 249.95 -> PHẢI NHẬN)
    makeTx('m05_a', 9100.1, { excelRowNumber: 6 }),
    makeTx('m05_b', -8850.15, { excelRowNumber: 7 }),
    // Cặp 3: diff = +0.06 (27100.10 + (-26850.04) = 250.06 -> PHẢI BỊ LOẠI)
    makeTx('p06_a', 27100.1, { excelRowNumber: 8 }),
    makeTx('p06_b', -26850.04, { excelRowNumber: 9 }),
    // Cặp 4: diff = -0.06 (81100.10 + (-80850.16) = 249.94 -> PHẢI BỊ LOẠI)
    makeTx('m06_a', 81100.1, { excelRowNumber: 10 }),
    makeTx('m06_b', -80850.16, { excelRowNumber: 11 }),
  ];

  const res = findCombinations({
    items,
    targetAmount: 250.0,
    tolerance: 0.05,
    maxItems: 2,
  });

  assert.equal(res.solutions.length, 3, 'Must include ONLY diff = 0, +0.05, -0.05');
  assert.equal(res.stats.exactCount, 1);
  assert.equal(res.stats.nearCount, 2);

  // Nghiệm đầu tiên phải là isExact = true, diff = 0
  assert.equal(res.solutions[0].isExact, true);
  assert.equal(res.solutions[0].diff, 0);
  assert.equal(res.solutions[0].absDiff, 0);

  // Hai nghiệm kế tiếp có absDiff = 0.05, isExact = false
  const nearDiffs = res.solutions.slice(1).map((s) => s.diff).sort((a, b) => a - b);
  assert.deepEqual(nearDiffs, [-0.05, 0.05]);
  for (const s of res.solutions.slice(1)) {
    assert.equal(s.isExact, false);
    assert.equal(s.absDiff, 0.05);
  }
});

// ============================================================================
// 2. CON SỐ HÀNG TỶ & HÀNG TRĂM TỶ VNĐ (VƯỢT GIỚI HẠN 32-BIT `2.147.483.647`)
// ============================================================================

test('2.1 B2B / Construction 450 Billion VNĐ contract (±500đ) with 15.85B, 128.45B, -24.30B and 32-bit overflow decoys', () => {
  const TWO_POW_32 = 4_294_967_296; // Nếu có phép toán `| 0` hoặc `>>> 0`, bội số 2^32 sẽ bị cắt cụt về 0!

  const items = [
    // Tổ hợp nghiệm đúng chính xác 450.000.000.000 VNĐ (6 dòng):
    // 15.850.000.000 + 128.450.750.000 - 24.300.250.000 + 195.600.500.000 + 168.900.000.000 - 34.501.000.000 = 450.000.000.000
    makeTx('b2b_1', 15_850_000_000, { excelRowNumber: 2, dateIso: '2026-01-10', description: 'Tạm ứng đợt 1' }),
    makeTx('b2b_2', 128_450_750_000, { excelRowNumber: 3, dateIso: '2026-01-25', description: 'Nghiệm thu giai đoạn 1' }),
    makeTx('b2b_3', -24_300_250_000, { excelRowNumber: 4, dateIso: '2026-02-02', description: 'Cấn trừ tạm ứng vật tư' }),
    makeTx('b2b_4', 195_600_500_000, { excelRowNumber: 5, dateIso: '2026-02-18', description: 'Nghiệm thu giai đoạn 2' }),
    makeTx('b2b_5', 168_900_000_000, { excelRowNumber: 6, dateIso: '2026-03-05', description: 'Nghiệm thu quyết toán' }),
    makeTx('b2b_6', -34_501_000_000, { excelRowNumber: 7, dateIso: '2026-03-10', description: 'Giữ lại bảo hành 5%' }),

    // Tổ hợp cận biên +500đ (2 dòng): 300 tỷ + 150.000.000.500 = 450.000.000.500
    makeTx('near_p500_a', 300_000_000_000, { excelRowNumber: 8, dateIso: '2026-03-12' }),
    makeTx('near_p500_b', 150_000_000_500, { excelRowNumber: 9, dateIso: '2026-03-12' }),

    // Tổ hợp ngoài biên +501đ (2 dòng): 250 tỷ + 200.000.000.501 = 450.000.000.501 -> PHẢI BỊ LOẠI!
    makeTx('out_p501_a', 250_000_000_000, { excelRowNumber: 10, dateIso: '2026-03-14' }),
    makeTx('out_p501_b', 200_000_000_501, { excelRowNumber: 11, dateIso: '2026-03-14' }),

    // Bẫy tràn 32-bit: 450.000.000.000 + 2^32 (đồng dư với 450 tỷ theo modulo 2^32!) -> PHẢI BỊ LOẠI!
    makeTx('bit32_trap', 450_000_000_000 + TWO_POW_32, { excelRowNumber: 12, dateIso: '2026-03-15' }),
  ];

  const res = findCombinations({
    items,
    targetAmount: 450_000_000_000,
    tolerance: 500,
    maxItems: 8,
  });

  // Kiểm tra tìm đúng nghiệm chính xác 450 tỷ và nghiệm cận biên +500đ, tuyệt đối không nhận +501đ hay bẫy 2^32
  const exactSol = res.solutions.find((s) => s.isExact);
  assert.ok(exactSol, 'Must find exact 6-item 450.000.000.000 VNĐ B2B combination');
  assert.equal(exactSol.totalAmount, 450_000_000_000);
  assert.equal(exactSol.diff, 0);
  assert.deepEqual(
    exactSol.itemIds.slice().sort(),
    ['b2b_1', 'b2b_2', 'b2b_3', 'b2b_4', 'b2b_5', 'b2b_6']
  );

  const near500Sol = res.solutions.find((s) => s.itemIds.includes('near_p500_a') && s.itemIds.includes('near_p500_b'));
  assert.ok(near500Sol, 'Must find +500đ boundary solution at 450.000.000.500 VNĐ');
  assert.equal(near500Sol.totalAmount, 450_000_000_500);
  assert.equal(near500Sol.diff, 500);
  assert.equal(near500Sol.isExact, false);

  for (const sol of res.solutions) {
    assert.ok(!sol.itemIds.includes('bit32_trap'), '32-bit modulo trap must never match');
    assert.ok(
      !(sol.itemIds.includes('out_p501_a') && sol.itemIds.includes('out_p501_b')),
      '+501đ out-of-bounds pair at 450B scale must never match'
    );
    assert.ok(Math.abs(sol.totalAmount - 450_000_000_000) <= 500);
  }
});

test('2.2 Single-dong (1 VNĐ) precision at 450 Billion & 900 Trillion VNĐ scale (tolerance = 0)', () => {
  // Kiểm tra sai lệch đúng 1 đồng ở quy mô 450 tỷ VNĐ và 900 nghìn tỷ VNĐ (9e14 < MAX_SAFE_INTEGER)
  const items450B = [
    makeTx('exact_1', 321_549_250_000, { excelRowNumber: 2 }),
    makeTx('exact_2', 128_450_750_000, { excelRowNumber: 3 }),
    makeTx('off_plus1', 128_450_750_001, { excelRowNumber: 4 }),
    makeTx('off_minus1', 128_450_749_999, { excelRowNumber: 5 }),
  ];

  const res450B = findCombinations({
    items: items450B,
    targetAmount: 450_000_000_000,
    tolerance: 0,
    maxItems: 4,
  });

  assert.equal(res450B.solutions.length, 1, 'Only the exact 1-dong match may pass when tolerance=0');
  assert.deepEqual(res450B.solutions[0].itemIds, ['exact_1', 'exact_2']);
  assert.equal(res450B.solutions[0].totalAmount, 450_000_000_000);
  assert.equal(res450B.solutions[0].diff, 0);

  // Quy mô 900 nghìn tỷ VNĐ (900.000.000.000.000) nhạy đến từng 1 đồng
  const items900T = [
    makeTx('t1', 512_345_678_901_234, { excelRowNumber: 2 }),
    makeTx('t2', 387_654_321_098_766, { excelRowNumber: 3 }),
    makeTx('t2_off1', 387_654_321_098_767, { excelRowNumber: 4 }),
  ];
  const res900T = findCombinations({
    items: items900T,
    targetAmount: 900_000_000_000_000,
    tolerance: 0,
    maxItems: 3,
  });
  assert.equal(res900T.solutions.length, 1);
  assert.deepEqual(res900T.solutions[0].itemIds, ['t1', 't2']);
});

test('2.3 Safe-integer precision (> 90.07 Trillion) must NOT overflow MAX_SAFE_INTEGER when a decimal row coexists in items', () => {
  // Cả 3 con số lớn dưới đây đều là Safe Integer hợp lệ tuyệt đối trong JS (`Number.isSafeInteger(...) === true`)
  const a = 400_000_000_000_001;
  const b = 100_000_000_000_002;
  const target = 500_000_000_000_003;
  assert.ok(Number.isSafeInteger(a) && Number.isSafeInteger(b) && Number.isSafeInteger(target));
  assert.equal(a + b, target);

  // Khi có dù chỉ 1 dòng lẻ thập phân không liên quan (`0.5`), việc nhân toàn cục `scale = 100` bằng kiểu Number
  // sẽ đẩy `4e14 * 100 = 4e16 > Number.MAX_SAFE_INTEGER (9.007e15)` làm sai số bậc 8 và đánh mất nghiệm đúng!
  const res = findCombinations({
    items: [
      makeTx('safe_a', a, { excelRowNumber: 2 }),
      makeTx('safe_b', b, { excelRowNumber: 3 }),
      makeTx('unrelated_dec', 0.5, { excelRowNumber: 4 }),
    ],
    targetAmount: target,
    tolerance: 0,
    maxItems: 3,
  });

  assert.equal(
    res.solutions.length,
    1,
    'Exact safe-integer match (a + b === target) MUST NOT be corrupted by scale=100 overflow from an unrelated 0.5 row'
  );
  assert.deepEqual(res.solutions[0].itemIds, ['safe_a', 'safe_b']);
  assert.equal(res.solutions[0].isExact, true);
  assert.equal(res.solutions[0].diff, 0);
});

// ============================================================================
// 3. BẪY BÙ TRỪ ÂM - DƯƠNG "TÀU LƯỢN", MỤC TIÊU ÂM & MỤC TIÊU BẰNG 0
// ============================================================================

test('3.1 Huge positive (+85.000.000) + huge negative (-84.500.000) netting to small target Z = 500.000', () => {
  const items = [
    makeTx('big_inv', 85_000_000, { excelRowNumber: 2, dateIso: '2026-03-02', description: 'Hóa đơn tổng lô hàng' }),
    makeTx('big_ret', -84_500_000, { excelRowNumber: 3, dateIso: '2026-03-05', description: 'Cấn trừ đặt cọc & trả hàng' }),
    makeTx('d1', 180_000, { excelRowNumber: 4 }),
    makeTx('d2', 210_000, { excelRowNumber: 5 }),
    makeTx('d3', 190_000, { excelRowNumber: 6 }),
    makeTx('d4', -70_000, { excelRowNumber: 7 }),
  ];

  const res = findCombinations({
    items,
    targetAmount: 500_000,
    tolerance: 0,
    maxItems: 6,
  });

  const foundBigNet = res.solutions.find(
    (s) => s.itemIds.length === 2 && s.itemIds.includes('big_inv') && s.itemIds.includes('big_ret')
  );
  assert.ok(foundBigNet, 'Must NOT prune +85.000.000 when target is only 500.000');
  assert.equal(foundBigNet.totalAmount, 500_000);
  assert.equal(foundBigNet.isExact, true);
  assert.equal(foundBigNet.diff, 0);
});

test('3.2 Rollercoaster 4 positives + 4 negatives swinging wildly (+300M down to -90M) before landing on Z = 1.250.000', () => {
  // 4 số dương: +95.000.000, +142.000.000, +68.500.000, +41.750.000 (Tổng dương = +347.250.000)
  // 4 số âm:   -118.000.000, -135.400.000, -62.300.000, -30.300.000 (Tổng âm   = -346.000.000)
  // Tổng đại số 8 dòng = +1.250.000!
  const items = [
    makeTx('p1', 95_000_000, { excelRowNumber: 2, dateIso: '2026-03-01' }),
    makeTx('n1', -118_000_000, { excelRowNumber: 3, dateIso: '2026-03-02' }),
    makeTx('p2', 142_000_000, { excelRowNumber: 4, dateIso: '2026-03-03' }),
    makeTx('n2', -135_400_000, { excelRowNumber: 5, dateIso: '2026-03-04' }),
    makeTx('p3', 68_500_000, { excelRowNumber: 6, dateIso: '2026-03-05' }),
    makeTx('n3', -62_300_000, { excelRowNumber: 7, dateIso: '2026-03-06' }),
    makeTx('p4', 41_750_000, { excelRowNumber: 8, dateIso: '2026-03-07' }),
    makeTx('n4', -30_300_000, { excelRowNumber: 9, dateIso: '2026-03-08' }),
    // Các dòng nhiễu lẻ nguyên tố không tạo nghiệm giả
    makeTx('noise1', 7_111_113, { excelRowNumber: 10, dateIso: '2026-03-09' }),
    makeTx('noise2', -5_333_337, { excelRowNumber: 11, dateIso: '2026-03-10' }),
    makeTx('noise3', 19_777_779, { excelRowNumber: 12, dateIso: '2026-03-11' }),
  ];

  const res = findCombinations({
    items,
    targetAmount: 1_250_000,
    tolerance: 0,
    maxItems: 10,
  });

  const expectedIds = ['n1', 'n2', 'n3', 'n4', 'p1', 'p2', 'p3', 'p4'];
  const rollercoasterSol = res.solutions.find(
    (s) => JSON.stringify(s.itemIds.slice().sort()) === JSON.stringify(expectedIds)
  );
  assert.ok(rollercoasterSol, 'Must discover the 8-item rollercoaster combination');
  assert.equal(rollercoasterSol.totalAmount, 1_250_000);
  assert.equal(rollercoasterSol.diff, 0);
  assert.equal(rollercoasterSol.isExact, true);
});

test('3.3 Negative targetAmount (Z = -3.250.000 ± 500): reconciling customer refund / credit balance', () => {
  const items = [
    // Tổ hợp 1 (Exact = -3.250.000): Hóa đơn +4.500.000 + Trả hàng -5.000.000 + Chiết khấu quý -2.750.000
    makeTx('inv1', 4_500_000, { excelRowNumber: 2, dateIso: '2026-03-01' }),
    makeTx('ret1', -5_000_000, { excelRowNumber: 3, dateIso: '2026-03-02' }),
    makeTx('reb1', -2_750_000, { excelRowNumber: 4, dateIso: '2026-03-03' }),

    // Tổ hợp 2 (Near diff = -350 -> total = -3.250.350): +50.000.000 + (-53.250.350)
    makeTx('n350_a', 50_000_000, { excelRowNumber: 5, dateIso: '2026-03-04' }),
    makeTx('n350_b', -53_250_350, { excelRowNumber: 6, dateIso: '2026-03-04' }),

    // Tổ hợp 3 (Near diff = +500 -> total = -3.249.500): +80.000.000 + (-83.249.500)
    makeTx('p500_a', 80_000_000, { excelRowNumber: 7, dateIso: '2026-03-05' }),
    makeTx('p500_b', -83_249_500, { excelRowNumber: 8, dateIso: '2026-03-05' }),

    // Tổ hợp 4 (Out of bounds diff = -501 -> total = -3.250.501 -> PHẢI BỊ LOẠI)
    makeTx('out_m501_a', 120_000_000, { excelRowNumber: 9, dateIso: '2026-03-06' }),
    makeTx('out_m501_b', -123_250_501, { excelRowNumber: 10, dateIso: '2026-03-06' }),

    // Tổ hợp 5 (Out of bounds diff = +501 -> total = -3.249.499 -> PHẢI BỊ LOẠI)
    makeTx('out_p501_a', 160_000_000, { excelRowNumber: 11, dateIso: '2026-03-07' }),
    makeTx('out_p501_b', -163_249_499, { excelRowNumber: 12, dateIso: '2026-03-07' }),
  ];

  const res = findCombinations({
    items,
    targetAmount: -3_250_000,
    tolerance: 500,
    maxItems: 3,
  });

  assert.equal(res.solutions.length, 3);
  assert.equal(res.solutions[0].isExact, true);
  assert.equal(res.solutions[0].totalAmount, -3_250_000);
  assert.equal(res.solutions[0].diff, 0);

  assert.equal(res.solutions[1].totalAmount, -3_250_350);
  assert.equal(res.solutions[1].diff, -350);
  assert.equal(res.solutions[1].absDiff, 350);

  assert.equal(res.solutions[2].totalAmount, -3_249_500);
  assert.equal(res.solutions[2].diff, 500);
  assert.equal(res.solutions[2].absDiff, 500);
});

test('3.4 Zero targetAmount (Z = 0, tolerance = 0 & tolerance = 500): self-cancelling invoice & credit note pairs', () => {
  // Dùng cơ sở lũy thừa 3 (100M, 300M, 900M) để các cặp 2 dòng không vô tình cộng chéo thành bộ 3 dòng (như 30M + 60M = 90M)
  const items = [
    // Cặp tự triệt tiêu hoàn toàn: +12.500.000 và -12.500.000 = 0
    makeTx('inv_125', 12_500_000, { excelRowNumber: 2, dateIso: '2026-03-01', description: 'Hóa đơn xuất hàng' }),
    makeTx('adj_125', -12_500_000, { excelRowNumber: 3, dateIso: '2026-03-02', description: 'Phiếu điều chỉnh giảm 100%' }),

    // Tổ hợp 3 dòng tự triệt tiêu: +8.000.000 + +4.500.000 + (-12.500.000) = 0
    makeTx('inv_80', 8_000_000, { excelRowNumber: 4, dateIso: '2026-03-03' }),
    makeTx('inv_45', 4_500_000, { excelRowNumber: 5, dateIso: '2026-03-03' }),

    // Cặp lệch +450đ quanh 0: +100.000.450 + (-100.000.000) = +450
    makeTx('near_z_a', 100_000_450, { excelRowNumber: 6, dateIso: '2026-03-04' }),
    makeTx('near_z_b', -100_000_000, { excelRowNumber: 7, dateIso: '2026-03-04' }),

    // Cặp lệch -500đ quanh 0: +300.000.000 + (-300.000.500) = -500
    makeTx('near_z2_a', 300_000_000, { excelRowNumber: 8, dateIso: '2026-03-05' }),
    makeTx('near_z2_b', -300_000_500, { excelRowNumber: 9, dateIso: '2026-03-05' }),

    // Cặp lệch +501đ quanh 0 (ngoài biên): +900.000.501 + (-900.000.000) = +501
    makeTx('out_z_a', 900_000_501, { excelRowNumber: 10, dateIso: '2026-03-06' }),
    makeTx('out_z_b', -900_000_000, { excelRowNumber: 11, dateIso: '2026-03-06' }),

    // Dòng chứng từ hủy 0đ không ghim -> không được tự biến thành nghiệm 1 dòng giả tạo
    makeTx('void_zero', 0, { excelRowNumber: 12, dateIso: '2026-03-06', description: 'Hóa đơn hủy 0đ' }),
  ];

  // Khi tolerance = 0: chỉ trả về đúng 2 tổ hợp tổng = 0 (2 dòng đứng trước 3 dòng)
  const resZeroExact = findCombinations({
    items,
    targetAmount: 0,
    tolerance: 0,
    maxItems: 3,
  });
  assert.equal(resZeroExact.solutions.length, 2);
  assert.deepEqual(resZeroExact.solutions[0].itemIds, ['inv_125', 'adj_125']);
  assert.equal(resZeroExact.solutions[0].totalAmount, 0);
  assert.equal(resZeroExact.solutions[0].diff, 0);
  assert.equal(resZeroExact.solutions[0].isExact, true);

  assert.deepEqual(resZeroExact.solutions[1].itemIds, ['adj_125', 'inv_80', 'inv_45']);
  assert.equal(resZeroExact.solutions[1].totalAmount, 0);
  assert.equal(resZeroExact.solutions[1].isExact, true);

  // Khi tolerance = 500: trả về 4 nghiệm (2 exact + 1 lệch +450 + 1 lệch -500), loại bỏ +501
  const resZeroTol500 = findCombinations({
    items,
    targetAmount: 0,
    tolerance: 500,
    maxItems: 3,
  });
  assert.equal(resZeroTol500.solutions.length, 4);
  assert.deepEqual(
    resZeroTol500.solutions.map((s) => s.diff),
    [0, 0, 450, -500]
  );
  for (const s of resZeroTol500.solutions) {
    assert.ok(!s.itemIds.includes('void_zero'), 'Unpinned 0-amount void row must not pollute solutions');
  }
});

// ============================================================================
// 4. KIỂM TRA BIÊN ĐỘ LỆCH TUYỆT ĐỐI (± 500) & THỨ TỰ XẾP HẠNG NGHIÊM NGẶT
// ============================================================================

test('4.1 Exact ±500 boundary inclusion/exclusion: diff = -500, -499, 0, +499, +500 IN; -501, +501 OUT', () => {
  const target = 10_000_000;
  // Thiết kế 7 cặp tách biệt bậc 100 triệu để không có bất kỳ tổ hợp chéo nào rơi gần 10 triệu
  const items = [
    // -501 (OUT)
    makeTx('m501_a', 110_000_000, { excelRowNumber: 2 }),
    makeTx('m501_b', -100_000_501, { excelRowNumber: 3 }),
    // -500 (IN)
    makeTx('m500_a', 210_000_000, { excelRowNumber: 4 }),
    makeTx('m500_b', -200_000_500, { excelRowNumber: 5 }),
    // -499 (IN)
    makeTx('m499_a', 310_000_000, { excelRowNumber: 6 }),
    makeTx('m499_b', -300_000_499, { excelRowNumber: 7 }),
    // 0 (IN - Exact)
    makeTx('z0_a', 410_000_000, { excelRowNumber: 8 }),
    makeTx('z0_b', -400_000_000, { excelRowNumber: 9 }),
    // +499 (IN)
    makeTx('p499_a', 510_000_000, { excelRowNumber: 10 }),
    makeTx('p499_b', -499_999_501, { excelRowNumber: 11 }),
    // +500 (IN)
    makeTx('p500_a', 610_000_000, { excelRowNumber: 12 }),
    makeTx('p500_b', -599_999_500, { excelRowNumber: 13 }),
    // +501 (OUT)
    makeTx('p501_a', 710_000_000, { excelRowNumber: 14 }),
    makeTx('p501_b', -699_999_499, { excelRowNumber: 15 }),
  ];

  const res = findCombinations({
    items,
    targetAmount: target,
    tolerance: 500,
    maxItems: 2,
  });

  assert.equal(res.solutions.length, 5, 'Exactly 5 combinations (-500, -499, 0, +499, +500) must appear');
  assert.equal(res.stats.exactCount, 1);
  assert.equal(res.stats.nearCount, 4);

  const diffs = res.solutions.map((s) => s.diff);
  assert.ok(!diffs.includes(-501), 'diff = -501 must NEVER appear');
  assert.ok(!diffs.includes(501), 'diff = +501 must NEVER appear');
  assert.deepEqual(
    diffs.slice().sort((a, b) => a - b),
    [-500, -499, 0, 499, 500]
  );

  // Kiểm tra xếp hạng: absDiff = 0 -> absDiff = 499 (2 nghiệm) -> absDiff = 500 (2 nghiệm)
  assert.deepEqual(
    res.solutions.map((s) => s.absDiff),
    [0, 499, 499, 500, 500]
  );
});

test('4.2 Strict 4-tier ranking hierarchy: absDiff -> itemCount -> dateSpanDays -> excelRowNumber', () => {
  const target = 20_000_000;
  // Dùng cơ sở lũy thừa 4 (100M, 400M, 1.6B, 6.4B, 25.6B) để không xảy ra cộng chéo giữa các nhóm 3 phần tử
  const items = [
    // Hạng 1 kỳ vọng: absDiff = 0, itemCount = 3, dateSpanDays = 20 (dù nhiều dòng & xa ngày hơn vẫn thắng absDiff = 1!)
    makeTx('r1_a', 100_000_000, { excelRowNumber: 20, dateIso: '2026-03-01' }),
    makeTx('r1_b', 10_000_000, { excelRowNumber: 21, dateIso: '2026-03-10' }),
    makeTx('r1_c', -90_000_000, { excelRowNumber: 22, dateIso: '2026-03-21' }),

    // Hạng 2 kỳ vọng: absDiff = 1 (+1đ), itemCount = 2, dateSpanDays = 0, excelRowNumber = [10, 11]
    makeTx('r2_a', 420_000_001, { excelRowNumber: 10, dateIso: '2026-03-05' }),
    makeTx('r2_b', -400_000_000, { excelRowNumber: 11, dateIso: '2026-03-05' }),

    // Hạng 3 kỳ vọng: absDiff = 1 (-1đ), itemCount = 2, dateSpanDays = 0, excelRowNumber = [30, 31] (cùng span=0 nhưng dòng Excel lớn hơn)
    makeTx('r3_a', 1_620_000_000, { excelRowNumber: 30, dateIso: '2026-03-05' }),
    makeTx('r3_b', -1_600_000_001, { excelRowNumber: 31, dateIso: '2026-03-05' }),

    // Hạng 4 kỳ vọng: absDiff = 1 (+1đ), itemCount = 2, dateSpanDays = 7 (cùng 2 dòng nhưng khoảng cách ngày xa hơn)
    makeTx('r4_a', 6_420_000_001, { excelRowNumber: 2, dateIso: '2026-03-01' }),
    makeTx('r4_b', -6_400_000_000, { excelRowNumber: 3, dateIso: '2026-03-08' }),

    // Hạng 5 kỳ vọng: absDiff = 1 (+1đ), itemCount = 3, dateSpanDays = 0 (cùng absDiff=1, dù cùng ngày nhưng 3 dòng > 2 dòng)
    // Lưu ý: Không đặt r5_a + r5_c = 0 để tránh ghép chéo với r6_single tạo thêm bộ 3 dòng phụ
    makeTx('r5_a', 20_000_000_000, { excelRowNumber: 4, dateIso: '2026-03-05' }),
    makeTx('r5_b', 5_620_000_001, { excelRowNumber: 5, dateIso: '2026-03-05' }),
    makeTx('r5_c', -25_600_000_000, { excelRowNumber: 6, dateIso: '2026-03-05' }),

    // Hạng 6 kỳ vọng: absDiff = 500 (-500đ), itemCount = 1 (dù chỉ 1 dòng duy nhất, absDiff=500 vẫn phải xếp sau absDiff=1!)
    makeTx('r6_single', 19_999_500, { excelRowNumber: 1, dateIso: '2026-03-05' }),
  ];

  const res = findCombinations({
    items,
    targetAmount: target,
    tolerance: 500,
    maxItems: 3,
  });

  assert.equal(res.solutions.length, 6);
  const rankedFirstIds = res.solutions.map((s) => s.itemIds.slice().sort().join('+'));
  assert.deepEqual(rankedFirstIds, [
    'r1_a+r1_b+r1_c', // absDiff=0, k=3, span=20
    'r2_a+r2_b',      // absDiff=1, k=2, span=0, rows=[10,11]
    'r3_a+r3_b',      // absDiff=1, k=2, span=0, rows=[30,31]
    'r4_a+r4_b',      // absDiff=1, k=2, span=7
    'r5_a+r5_b+r5_c', // absDiff=1, k=3, span=0
    'r6_single',      // absDiff=500, k=1, span=0
  ]);
});

test('4.3 Intra-level dateSpanDays optimality when >200 exact solutions exist at same itemCount k', () => {
  // Kiểm tra lỗi cắt sớm trong cùng tầng k (`compactLevelIfNeeded` gán `stopCurrentLevel = true`
  // khi thu thập đủ 200 nghiệm exact ở nhánh số tiền lớn có dateSpanDays = 27 ngày, làm mất nghiệm exact
  // cùng ngày `dateSpanDays = 0` ở nhánh số tiền nhỏ hơn!).
  const items = [];
  let rowNo = 2;

  // Tạo 15 cặp bù trừ ở mức mệnh giá lớn (100M..114M) ở ngày 01/03 và 15 số ở ngày 28/03
  // sao cho có 15 * 14 = 210 bộ 3 dòng (k = 3) khớp đúng target = 30.000.000 nhưng dateSpanDays = 27 ngày:
  // A_i (01/03) + B_j (28/03) + C_{i,j}? Dễ hơn: dùng 22 số nhóm X (01/03) và 22 số nhóm Y (28/03):
  // Với k = 2: chỉ có 22 cặp. Với k = 3: chọn i < j từ 22 số:
  // Đặt X_i = 100_000_000 + i * 1_000 (i = 1..22, ngày 01/03)
  // Đặt Y_j = 100_000_000 + j * 23_000 (j = 1..12, ngày 28/03)
  // Đặt Z_{i,j} = 30_000_000 - X_i - Y_j = -170_000_000 - i*1_000 - j*23_000 (264 số âm phân biệt, ngày 28/03)
  // Mỗi bộ (X_i, Y_j, Z_{i,j}) có tổng đúng 30.000.000 (k = 3, absDiff = 0), nhưng dateSpanDays = 27 ngày!
  for (let i = 1; i <= 18; i++) {
    items.push(makeTx(`X_${i}`, 100_000_000 + i * 1_000, { excelRowNumber: rowNo++, dateIso: '2026-03-01' }));
  }
  for (let j = 1; j <= 13; j++) {
    items.push(makeTx(`Y_${j}`, 100_000_000 + j * 29_000, { excelRowNumber: rowNo++, dateIso: '2026-03-28' }));
  }
  for (let i = 1; i <= 18; i++) {
    for (let j = 1; j <= 13; j++) {
      const zAmt = 30_000_000 - (100_000_000 + i * 1_000) - (100_000_000 + j * 29_000);
      items.push(makeTx(`Z_${i}_${j}`, zAmt, { excelRowNumber: rowNo++, dateIso: '2026-03-28' }));
    }
  }

  // Thêm 1 bộ 3 dòng mệnh giá nhỏ hơn (10M + 10.5M + 9.5M = 30M) CÙNG NGÀY 15/03/2026 (dateSpanDays = 0)!
  // Theo quy tắc kế toán (compareSolutions): cùng absDiff=0, cùng k=3 -> dateSpanDays=0 PHẢI ĐỨNG ĐẦU TIÊN (#1)!
  items.push(makeTx('best_1', 10_500_000, { excelRowNumber: 2, dateIso: '2026-03-15' }));
  items.push(makeTx('best_2', 10_000_000, { excelRowNumber: 3, dateIso: '2026-03-15' }));
  items.push(makeTx('best_3', 9_500_000, { excelRowNumber: 4, dateIso: '2026-03-15' }));

  const res = findCombinations({
    items,
    targetAmount: 30_000_000,
    tolerance: 0,
    minItems: 3,
    maxItems: 3,
    maxResults: 10,
    timeLimitMs: 5000,
  });

  assert.ok(res.solutions.length >= 1);
  assert.equal(
    res.solutions[0].dateSpanDays,
    0,
    'Rank #1 solution at k=3 MUST be the same-day (dateSpanDays=0) exact match, NOT pruned by wide-span earlier branches'
  );
  assert.deepEqual(res.solutions[0].itemIds.slice().sort(), ['best_1', 'best_2', 'best_3']);
});

test('4.4 Intra-level near-match (absDiff > 0) dateSpanDays optimality when >200 near solutions exist at same itemCount k', () => {
  // Kiểm tra nhánh thứ 2 của `compactLevelIfNeeded`: `levelIntTolerance = worstKept.intAbsDiff - 1`
  // Khi không có nghiệm exact nào mà có >200 nghiệm lệch +100đ (dateSpanDays = 27 ngày) ở nhánh lớn,
  // việc siết `levelIntTolerance = 99` trong cùng tầng k=3 sẽ cắt nhầm nghiệm lệch +100đ có `dateSpanDays = 0`!
  const items = [];
  let rowNo = 2;

  for (let i = 1; i <= 18; i++) {
    items.push(makeTx(`NX_${i}`, 100_000_000 + i * 1_000, { excelRowNumber: rowNo++, dateIso: '2026-03-01' }));
  }
  for (let j = 1; j <= 13; j++) {
    items.push(makeTx(`NY_${j}`, 100_000_000 + j * 29_000, { excelRowNumber: rowNo++, dateIso: '2026-03-28' }));
  }
  for (let i = 1; i <= 18; i++) {
    for (let j = 1; j <= 13; j++) {
      // Tổng luôn bằng 30.000.100 (absDiff = 100, dateSpanDays = 27)
      const zAmt = 30_000_100 - (100_000_000 + i * 1_000) - (100_000_000 + j * 29_000);
      items.push(makeTx(`NZ_${i}_${j}`, zAmt, { excelRowNumber: rowNo++, dateIso: '2026-03-28' }));
    }
  }

  // Bộ 3 dòng mệnh giá nhỏ hơn cũng có absDiff = 100 nhưng CÙNG NGÀY 15/03/2026 (dateSpanDays = 0) -> PHẢI ĐỨNG HẠNG #1!
  items.push(makeTx('near_best_1', 10_500_100, { excelRowNumber: 2, dateIso: '2026-03-15' }));
  items.push(makeTx('near_best_2', 10_000_000, { excelRowNumber: 3, dateIso: '2026-03-15' }));
  items.push(makeTx('near_best_3', 9_500_000, { excelRowNumber: 4, dateIso: '2026-03-15' }));

  const res = findCombinations({
    items,
    targetAmount: 30_000_000,
    tolerance: 500,
    minItems: 3,
    maxItems: 3,
    maxResults: 10,
    timeLimitMs: 5000,
  });

  assert.ok(res.solutions.length >= 1);
  assert.equal(res.solutions[0].absDiff, 100);
  assert.equal(
    res.solutions[0].dateSpanDays,
    0,
    'Rank #1 near-match at k=3 MUST be dateSpanDays=0, NOT pruned by `worstKept.intAbsDiff - 1` intra-level tightening'
  );
  assert.deepEqual(res.solutions[0].itemIds.slice().sort(), ['near_best_1', 'near_best_2', 'near_best_3']);
});

// ============================================================================
// 5. DIFFERENTIAL FUZZING VỚI BIGINT BRUTE-FORCE ORACLE (>= 300 KỊCH BẢN SEEDED)
// ============================================================================

test('5.1 Differential Fuzzing (320 seeded scenarios, N=10..15): integer VNĐ, +/-/0, billions, pinned, excluded, dates, tolerance', () => {
  const rng = createSeededRng(20261101);
  const NUM_SCENARIOS = 320;

  for (let scenario = 1; scenario <= NUM_SCENARIOS; scenario++) {
    const n = rng.nextInt(10, 15);
    const useBillionScale = scenario % 4 === 0; // 25% kịch bản ở quy mô hàng tỷ - hàng trăm tỷ VNĐ
    const items = [];
    const amountFreq = new Map();

    for (let i = 0; i < n; i++) {
      let amount;
      const roll = rng.nextFloat();
      if (roll < 0.08) {
        // 8% dòng mệnh giá 0đ
        amount = 0;
      } else if (useBillionScale) {
        // Khoản tiền từ -150 tỷ đến +250 tỷ VNĐ, có lẻ đến từng đồng!
        const sign = rng.nextFloat() < 0.35 ? -1 : 1;
        const billions = rng.nextInt(1, 200) * 1_000_000_000;
        const oddDong = rng.nextInt(0, 999_999);
        amount = sign * (billions + oddDong);
      } else {
        // Khoản tiền từ -50 triệu đến +80 triệu VNĐ, có lẻ đến từng đồng
        const sign = rng.nextFloat() < 0.35 ? -1 : 1;
        amount = sign * rng.nextInt(1, 80_000_000);
      }

      // Giới hạn tối đa 3 dòng trùng hệt mệnh giá khác 0 để đối chiếu 1:1 tất cả biến thể dòng cụ thể (MAX_VARIANTS_PER_MULTISET = 3)
      if (amount !== 0) {
        while ((amountFreq.get(amount) || 0) >= 3) {
          amount += amount > 0 ? 1 : -1;
        }
        amountFreq.set(amount, (amountFreq.get(amount) || 0) + 1);
      }

      // Tạo khoảng 15% xác suất trùng mệnh giá có kiểm soát (<= 3 dòng) để test cả nhánh expandMultisetSolutions
      if (amount !== 0 && i > 0 && rng.nextFloat() < 0.15) {
        const prevAmt = items[rng.nextInt(0, i - 1)].amount;
        if (prevAmt !== 0 && (amountFreq.get(prevAmt) || 0) < 3) {
          amountFreq.set(amount, (amountFreq.get(amount) || 1) - 1);
          amount = prevAmt;
          amountFreq.set(amount, (amountFreq.get(amount) || 0) + 1);
        }
      }

      const day = rng.nextInt(1, 28);
      const month = rng.nextInt(1, 3);
      const dateIso = `2026-0${month}-${String(day).padStart(2, '0')}`;
      const isSummaryRow = rng.nextFloat() < 0.06;
      const tag = rng.pick(['Khách hàng Alpha', 'Công ty Minh Phú', 'Đối tác Đông Á', 'Phụ lục B2B']);

      items.push(
        makeTx(`s${scenario}_r${i + 1}`, amount, {
          excelRowNumber: i + 2,
          dateIso,
          description: `${isSummaryRow ? 'Tổng cộng ' : 'HĐ '}${tag} #${i + 1}`,
          reference: `REF-${i + 1}`,
          isSummaryRow,
        })
      );
    }

    // Sinh ngẫu nhiên pinnedIds (0..2 dòng), excludedIds (0..2 dòng)
    const shuffledIds = rng.shuffle(items.map((it) => it.id));
    const pinCount = rng.nextFloat() < 0.35 ? rng.nextInt(1, 2) : 0;
    const exclCount = rng.nextFloat() < 0.35 ? rng.nextInt(1, 2) : 0;
    const pinnedIds = shuffledIds.slice(0, pinCount);
    // Đảm bảo hầu hết kịch bản không trùng pin & exclude, nhưng cho phép ~5% kịch bản xung đột để test cả nhánh rỗng
    const excludedIds =
      rng.nextFloat() < 0.05
        ? shuffledIds.slice(0, exclCount)
        : shuffledIds.slice(pinCount, pinCount + exclCount);

    const useDateFilter = rng.nextFloat() < 0.3;
    const dateFromIso = useDateFilter ? '2026-01-10' : null;
    const dateToIso = useDateFilter ? '2026-03-20' : null;
    const keyword = rng.nextFloat() < 0.2 ? rng.pick(['minh phu', 'alpha', 'dong a']) : '';
    const includeSummaryRows = rng.nextFloat() < 0.2;
    const minItems = rng.nextInt(1, 3);
    const maxItems = rng.nextInt(Math.max(minItems, 3), Math.min(10, n));

    // Chọn targetAmount:
    // - 55% tạo từ tổng thực tế của một tập con ngẫu nhiên (kèm nhiễu nhỏ) để chắc chắn có nghiệm thực tế (cả dương, âm, 0)
    // - 15% target = 0
    // - 15% target âm ngẫu nhiên
    // - 15% target ngẫu nhiên bất kỳ (test cả trường hợp vô nghiệm)
    let targetAmount;
    const targetMode = rng.nextFloat();
    if (targetMode < 0.55) {
      const subsetSize = rng.nextInt(1, Math.min(5, n));
      const picked = rng.shuffle(items).slice(0, subsetSize);
      const rawBig = picked.reduce((acc, it) => acc + BigInt(it.amount), 0n);
      const jitter = BigInt(rng.nextInt(-600, 600));
      targetAmount = Number(rawBig + jitter);
    } else if (targetMode < 0.7) {
      targetAmount = 0;
    } else if (targetMode < 0.85) {
      targetAmount = useBillionScale
        ? -rng.nextInt(1, 80) * 1_000_000_000
        : -rng.nextInt(100_000, 25_000_000);
    } else {
      targetAmount = useBillionScale
        ? rng.nextInt(1, 300) * 1_000_000_000
        : rng.nextInt(100_000, 100_000_000);
    }

    const tolerance = rng.pick([0, 0, 1, 100, 500, 500, 1500]);

    const queryOptions = {
      items,
      dateFromIso,
      dateToIso,
      keyword,
      targetAmount,
      tolerance,
      pinnedIds,
      excludedIds,
      includeSummaryRows,
      minItems,
      maxItems,
      maxResults: 150, // Đặt cao để so khớp toàn diện tập nghiệm
      timeLimitMs: 5000,
    };

    const oracle = oracleBruteForce(queryOptions);
    const actual = findCombinations(queryOptions);

    assert.equal(actual.stats.timedOut, false, `Scenario #${scenario} timed out unexpectedly`);

    // 1. Kiểm tra KHÔNG CÓ FALSE POSITIVE: Mọi nghiệm từ findCombinations phải hợp lệ 100% trong Oracle
    const oracleMap = new Map(oracle.allValidSolutions.map((s) => [s.canonicalKey, s]));
    for (const sol of actual.solutions) {
      const key = sol.itemIds.slice().sort().join('\u0001');
      const expectedSol = oracleMap.get(key);
      assert.ok(
        expectedSol,
        `Scenario #${scenario}: FALSE POSITIVE detected! Solution ${sol.id} (total=${sol.totalAmount}, diff=${sol.diff}) is invalid according to BigInt Oracle`
      );
      assert.equal(sol.totalAmount, expectedSol.totalAmount, `Scenario #${scenario}: totalAmount mismatch`);
      assert.equal(sol.diff, expectedSol.diff, `Scenario #${scenario}: diff mismatch`);
      assert.equal(sol.absDiff, expectedSol.absDiff, `Scenario #${scenario}: absDiff mismatch`);
      assert.equal(sol.isExact, expectedSol.isExact, `Scenario #${scenario}: isExact mismatch`);
      assert.equal(sol.dateSpanDays, expectedSol.dateSpanDays, `Scenario #${scenario}: dateSpanDays mismatch`);
    }

    // 2. Kiểm tra KHÔNG CÓ FALSE NEGATIVE & ĐÚNG THỨ TỰ XẾP HẠNG 1:1
    assert.equal(
      actual.solutions.length,
      oracle.solutions.length,
      `Scenario #${scenario}: Solution count mismatch (actual=${actual.solutions.length} vs oracle=${oracle.solutions.length}, target=${targetAmount}, tol=${tolerance})`
    );

    for (let idx = 0; idx < oracle.solutions.length; idx++) {
      assert.equal(
        actual.solutions[idx].id,
        oracle.solutions[idx].id,
        `Scenario #${scenario} rank #${idx}: Expected ${oracle.solutions[idx].id} (absDiff=${oracle.solutions[idx].absDiff}, k=${oracle.solutions[idx].itemCount}, span=${oracle.solutions[idx].dateSpanDays}) but got ${actual.solutions[idx].id}`
      );
    }
  }
});

test('5.2 Differential Fuzzing (60 seeded 2-decimal scenarios, N=10..14): USD/EUR cents & VAT rounded to 2 decimals vs BigInt Oracle', () => {
  const rng = createSeededRng(88882026);
  const NUM_SCENARIOS = 60;

  for (let scenario = 1; scenario <= NUM_SCENARIOS; scenario++) {
    const n = rng.nextInt(10, 14);
    const items = [];
    const centsFreq = new Map();

    for (let i = 0; i < n; i++) {
      let cents = rng.nextFloat() < 0.08 ? 0 : (rng.nextFloat() < 0.35 ? -1 : 1) * rng.nextInt(1, 5_000_000);
      if (cents !== 0) {
        while ((centsFreq.get(cents) || 0) >= 3) {
          cents += cents > 0 ? 1 : -1;
        }
        centsFreq.set(cents, (centsFreq.get(cents) || 0) + 1);
      }
      const amount = cents / 100;
      const day = rng.nextInt(1, 28);
      items.push(
        makeTx(`dec2_${scenario}_${i + 1}`, amount, {
          excelRowNumber: i + 2,
          dateIso: `2026-02-${String(day).padStart(2, '0')}`,
        })
      );
    }

    const picked = rng.shuffle(items).slice(0, rng.nextInt(2, 5));
    const sumBig = picked.reduce((acc, it) => acc + numberToScaledBigInt(it.amount), 0n);
    const jitterCents = BigInt(rng.nextInt(-5, 5)) * 10_000n; // ±0.05
    const targetAmount =
      scenario % 5 === 0 ? 0 : scaledBigIntToNumber(sumBig + jitterCents);
    const tolerance = rng.pick([0, 0.01, 0.05, 0.25]);

    const queryOptions = {
      items,
      targetAmount,
      tolerance,
      minItems: 1,
      maxItems: 8,
      maxResults: 120,
    };

    const oracle = oracleBruteForce(queryOptions);
    const actual = findCombinations(queryOptions);

    assert.equal(
      actual.solutions.length,
      oracle.solutions.length,
      `2-Decimal Scenario #${scenario}: count mismatch (${actual.solutions.length} vs ${oracle.solutions.length})`
    );
    for (let idx = 0; idx < oracle.solutions.length; idx++) {
      assert.equal(actual.solutions[idx].id, oracle.solutions[idx].id, `2-Decimal Scenario #${scenario} rank #${idx}`);
      assert.equal(actual.solutions[idx].diff, oracle.solutions[idx].diff);
      assert.equal(actual.solutions[idx].isExact, oracle.solutions[idx].isExact);
    }
  }
});

test('5.3 Differential Fuzzing (40 seeded 3-4 decimal scenarios, N=10..12): VAT 8%/10% & FX 3-4 decimals vs BigInt Oracle', () => {
  const rng = createSeededRng(99993456);
  const NUM_SCENARIOS = 40;
  let falseNegativeCount = 0;
  let falsePositiveCount = 0;
  let firstMismatchDetail = '';

  for (let scenario = 1; scenario <= NUM_SCENARIOS; scenario++) {
    const n = rng.nextInt(10, 12);
    const items = [];
    const milliFreq = new Map();

    for (let i = 0; i < n; i++) {
      // Sinh số thập phân 3-4 chữ số (đơn vị 10^-4, ví dụ 10.1250, 20.3750, 108.0540, -15.6250)
      const sign = rng.nextFloat() < 0.3 ? -1 : 1;
      let units1e4 = sign * rng.nextInt(50, 500_000);
      while ((milliFreq.get(units1e4) || 0) >= 3) {
        units1e4 += sign;
      }
      milliFreq.set(units1e4, (milliFreq.get(units1e4) || 0) + 1);

      items.push(
        makeTx(`dec4_${scenario}_${i + 1}`, units1e4 / 10_000, {
          excelRowNumber: i + 2,
          dateIso: `2026-03-${String(rng.nextInt(1, 28)).padStart(2, '0')}`,
        })
      );
    }

    // Luôn chọn 1 tập con thực tế làm targetAmount với tolerance = 0 (hoặc 0.005) để kiểm chứng độ chính xác tuyệt đối
    const picked = rng.shuffle(items).slice(0, rng.nextInt(2, 4));
    const exactTargetBig = picked.reduce((acc, it) => acc + numberToScaledBigInt(it.amount), 0n);
    const targetAmount = scaledBigIntToNumber(exactTargetBig);
    const tolerance = rng.pick([0, 0, 0.005]);

    const queryOptions = {
      items,
      targetAmount,
      tolerance,
      minItems: 1,
      maxItems: 6,
      maxResults: 100,
    };

    const oracle = oracleBruteForce(queryOptions);
    const actual = findCombinations(queryOptions);

    const oracleKeys = new Set(oracle.solutions.map((s) => s.canonicalKey));
    const actualKeys = new Set(actual.solutions.map((s) => s.canonicalKey));

    for (const k of oracleKeys) {
      if (!actualKeys.has(k)) {
        falseNegativeCount++;
        if (!firstMismatchDetail) {
          const missed = oracle.solutions.find((s) => s.canonicalKey === k);
          firstMismatchDetail = `Scenario #${scenario} (target=${targetAmount}, tol=${tolerance}): Missed valid solution [${missed.items.map((x) => x.amount).join(' + ')}] = ${missed.totalAmount}`;
        }
      }
    }
    for (const k of actualKeys) {
      if (!oracleKeys.has(k)) {
        falsePositiveCount++;
        if (!firstMismatchDetail) {
          const bogus = actual.solutions.find((s) => s.canonicalKey === k);
          firstMismatchDetail = `Scenario #${scenario} (target=${targetAmount}, tol=${tolerance}): False positive [${bogus.items.map((x) => x.amount).join(' + ')}] reported total=${bogus.totalAmount}`;
        }
      }
    }
  }

  assert.equal(
    falseNegativeCount + falsePositiveCount,
    0,
    `3-4 Decimal Fuzzing detected ${falseNegativeCount} false negatives and ${falsePositiveCount} false positives due to fixed scale=100 rounding! First example: ${firstMismatchDetail}`
  );
});

test('5.4 Differential Fuzzing (40 seeded heavy-duplicate & pinned-zero scenarios, N=12..15): multiset completeness & zero false positives', () => {
  const rng = createSeededRng(77772026);
  const NUM_SCENARIOS = 40;

  for (let scenario = 1; scenario <= NUM_SCENARIOS; scenario++) {
    const n = rng.nextInt(12, 15);
    const palette = [
      500_000,
      1_200_000,
      -500_000,
      -300_000,
      2_500_000,
      0,
    ];
    const items = [];
    for (let i = 0; i < n; i++) {
      const amount = rng.pick(palette);
      items.push(
        makeTx(`dup_${scenario}_${i + 1}`, amount, {
          excelRowNumber: i + 2,
          dateIso: `2026-03-${String(rng.nextInt(1, 28)).padStart(2, '0')}`,
        })
      );
    }

    // Ghim cả dòng 0đ (nếu có) ở một nửa số kịch bản để kiểm chứng xử lý dòng ghim 0đ
    const zeroItem = items.find((x) => x.amount === 0);
    const pinnedIds = zeroItem && scenario % 2 === 0 ? [zeroItem.id] : [];
    const targetAmount = rng.pick([0, 500_000, 1_500_000, -500_000, 2_200_000]);
    const tolerance = rng.pick([0, 200, 500]);

    const queryOptions = {
      items,
      targetAmount,
      tolerance,
      pinnedIds,
      minItems: 1,
      maxItems: 6,
      maxResults: 200,
    };

    const oracle = oracleBruteForce(queryOptions);
    const actual = findCombinations(queryOptions);

    // 1. Mọi nghiệm trả về đều phải hợp lệ trong Oracle (Zero False Positives)
    const oracleByKey = new Map(oracle.allValidSolutions.map((s) => [s.canonicalKey, s]));
    for (const sol of actual.solutions) {
      const ref = oracleByKey.get(sol.itemIds.slice().sort().join('\u0001'));
      assert.ok(ref, `Heavy-dup Scenario #${scenario}: False positive ${sol.id}`);
      assert.equal(sol.totalAmount, ref.totalAmount);
      assert.equal(sol.diff, ref.diff);
      if (pinnedIds.length > 0) {
        for (const pid of pinnedIds) {
          assert.ok(sol.itemIds.includes(pid), `Pinned zero-amount item ${pid} must appear in every solution`);
        }
      }
    }

    // 2. Tập hợp các multiset mệnh giá hợp lệ giữa Oracle và Engine phải khớp 100% (Zero Missed Multisets)
    const toMultisetSignature = (solItems) =>
      solItems
        .map((it) => it.amount)
        .sort((a, b) => a - b)
        .join('|');

    const oracleMultisets = new Set(oracle.allValidSolutions.map((s) => toMultisetSignature(s.items)));
    const actualMultisets = new Set(actual.solutions.map((s) => toMultisetSignature(s.items)));
    assert.deepEqual(
      [...actualMultisets].sort(),
      [...oracleMultisets].sort(),
      `Heavy-dup Scenario #${scenario}: Multiset coverage mismatch`
    );
  }
});
