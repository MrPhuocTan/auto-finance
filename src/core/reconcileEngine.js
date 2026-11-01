import {
  DEFAULT_TOLERANCE,
  DEFAULT_MAX_ITEMS,
  DEFAULT_MAX_RESULTS,
  DEFAULT_TIME_LIMIT_MS,
} from './contracts.js';

const MS_PER_DAY = 86_400_000;
const MAX_VARIANTS_PER_MULTISET = 3;
const MAX_EXACT_DECIMALS = 6;
const PROGRESS_INTERVAL_MS = 60;
const NODE_CHECK_MASK = 2047; // Kiểm tra đồng hồ & cờ hủy mỗi 2048 nút

/**
 * Chuẩn hóa chuỗi tiếng Việt để tìm kiếm không phân biệt hoa thường & dấu:
 * - Tách tổ hợp dấu Unicode (NFD) và xóa combining marks
 * - Chuyển đ/Đ -> d
 * - Chuyển chữ thường, gom khoảng trắng thừa và trim
 *
 * @param {any} str
 * @returns {string}
 */
export function normalizeVietnameseSearch(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[đĐ]/g, 'd')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Trích xuất văn bản phục vụ tìm kiếm keyword từ description, reference và extraCells.
 *
 * @param {object} item
 * @returns {string}
 */
function buildSearchHaystack(item) {
  const parts = [];
  if (item.description !== null && item.description !== undefined && item.description !== '') {
    parts.push(String(item.description));
  }
  if (item.reference !== null && item.reference !== undefined && item.reference !== '') {
    parts.push(String(item.reference));
  }

  const extra = item.extraCells;
  if (Array.isArray(extra)) {
    for (let i = 0; i < extra.length; i++) {
      const entry = extra[i];
      if (entry === null || entry === undefined) continue;
      if (typeof entry === 'object') {
        const vals = Object.values(entry);
        for (let j = 0; j < vals.length; j++) {
          if (vals[j] !== null && vals[j] !== undefined && vals[j] !== '') {
            parts.push(String(vals[j]));
          }
        }
      } else if (entry !== '') {
        parts.push(String(entry));
      }
    }
  } else if (extra && typeof extra === 'object') {
    const vals = Object.values(extra);
    for (let i = 0; i < vals.length; i++) {
      const v = vals[i];
      if (v !== null && v !== undefined && v !== '') {
        parts.push(String(v));
      }
    }
  }

  return normalizeVietnameseSearch(parts.join('\n'));
}

/**
 * Lấy UTC day timestamp của một dòng giao dịch (từ item.timestamp hoặc fallback sang item.dateIso).
 *
 * @param {object} item
 * @returns {number | null}
 */
function getItemDayTimestamp(item) {
  if (Number.isFinite(item.timestamp)) {
    return item.timestamp;
  }
  if (typeof item.dateIso === 'string') {
    const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(item.dateIso.trim());
    if (m) {
      return Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    }
  }
  return null;
}

/**
 * Đếm số chữ số thập phân hữu hiệu của một số thực JS dựa trên biểu diễn chuỗi ngắn nhất chuẩn ECMAScript.
 *
 * @param {number} num
 * @returns {number}
 */
function getDecimalDigitsCount(num) {
  if (!Number.isFinite(num) || num === 0) return 0;
  const raw = String(num).trim().toLowerCase();
  let body = raw[0] === '-' || raw[0] === '+' ? raw.slice(1) : raw;
  let exp = 0;
  const eIdx = body.indexOf('e');
  if (eIdx !== -1) {
    exp = parseInt(body.slice(eIdx + 1), 10);
    body = body.slice(0, eIdx);
  }
  const dotIdx = body.indexOf('.');
  const fracLen = dotIdx === -1 ? 0 : body.length - dotIdx - 1;
  return Math.max(0, fracLen - exp);
}

/**
 * Chuyển đổi chính xác tuyệt đối số thực JS sang BigInt đã nhân 10^scaleDecimals bằng phân tích chuỗi chữ số:
 * - Không bao giờ dùng phép nhân dấu phẩy động `num * 10^k` (tránh sai lệch IEEE-754 và tránh tràn MAX_SAFE_INTEGER).
 * - Đối xứng dấu hoàn hảo (half-away-from-zero) nếu số chữ số lẻ vượt quá scaleDecimals.
 *
 * @param {number} num
 * @param {number} scaleDecimals
 * @returns {bigint}
 */
function numberToScaledBigInt(num, scaleDecimals) {
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
  const eIdx = body.indexOf('e');
  if (eIdx !== -1) {
    mantissa = body.slice(0, eIdx);
    exp = parseInt(body.slice(eIdx + 1), 10);
  }

  const dotIdx = mantissa.indexOf('.');
  const intStr = dotIdx === -1 ? mantissa : mantissa.slice(0, dotIdx);
  const fracStr = dotIdx === -1 ? '' : mantissa.slice(dotIdx + 1);
  const digits = (intStr + fracStr).replace(/^0+/, '') || '0';
  if (digits === '0') return 0n;

  const totalShift = exp - fracStr.length + scaleDecimals;
  if (totalShift >= 0) {
    return sign * BigInt(digits + '0'.repeat(totalShift));
  }

  const cut = -totalShift;
  if (cut > digits.length) {
    return 0n;
  }
  const head = digits.slice(0, digits.length - cut) || '0';
  const nextDigit = Number(digits[digits.length - cut] || '0');
  let val = BigInt(head);
  if (nextDigit >= 5) val += 1n;
  return val === 0n ? 0n : sign * val;
}

/**
 * Quy đổi ngược từ BigInt đã nhân 10^scaleDecimals về số thực JS chuẩn xác, đảm bảo 0 luôn là +0.
 *
 * @param {bigint} bigVal
 * @param {number} scaleDecimals
 * @returns {number}
 */
function scaledBigIntToNumber(bigVal, scaleDecimals) {
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
  const num = Number(neg ? `-${outStr}` : outStr);
  return num === 0 ? 0 : num;
}

/**
 * Định dạng số nguyên có dấu chấm phân cách hàng nghìn theo chuẩn kế toán VN.
 *
 * @param {number} n
 * @returns {string}
 */
function formatIntVN(n) {
  const sign = n < 0 ? '-' : '';
  const str = String(Math.round(Math.abs(n))).replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  return sign + str;
}

/**
 * Định dạng mệnh giá BigInt đã scale sang chuỗi tiền tệ chuẩn VN phục vụ ghi chú dòng trùng mệnh giá.
 *
 * @param {bigint} bigVal
 * @param {number} scaleDecimals
 * @returns {string}
 */
function formatScaledBigIntVN(bigVal, scaleDecimals) {
  if (bigVal === 0n) return '0';
  const sign = bigVal < 0n ? '-' : '';
  const absStr = (bigVal < 0n ? -bigVal : bigVal).toString();

  if (scaleDecimals === 0) {
    return sign + absStr.replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  }

  let intPart = '0';
  let fracPart = '';
  if (absStr.length <= scaleDecimals) {
    fracPart = absStr.padStart(scaleDecimals, '0').replace(/0+$/, '');
  } else {
    intPart = absStr.slice(0, absStr.length - scaleDecimals);
    fracPart = absStr.slice(absStr.length - scaleDecimals).replace(/0+$/, '');
  }

  const formattedInt = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  return fracPart ? `${sign}${formattedInt},${fracPart}` : `${sign}${formattedInt}`;
}

/**
 * Tính chính xác tổ hợp C(n, k) bằng BigInt với ngưỡng cắt trần `cap` (không sai số chia thực).
 *
 * @param {number} n
 * @param {number} k
 * @param {number} cap
 * @returns {{ value: number, capped: boolean }}
 */
function calcBinomialCapped(n, k, cap = 1_000_000) {
  if (k < 0 || k > n) return { value: 0, capped: false };
  if (k === 0 || k === n) return { value: 1, capped: false };
  const kk = Math.min(k, n - k);
  const bigCap = BigInt(cap);
  let res = 1n;
  for (let i = 1; i <= kk; i++) {
    res = (res * BigInt(n - kk + i)) / BigInt(i);
    if (res >= bigCap) {
      return { value: cap, capped: true };
    }
  }
  return { value: Number(res), capped: false };
}

/**
 * Sắp xếp các phần tử bên trong một nghiệm tăng dần theo ngày (timestamp) rồi theo dòng Excel.
 *
 * @param {object} a
 * @param {object} b
 * @returns {number}
 */
function compareItemsChronological(a, b) {
  const ta = getItemDayTimestamp(a);
  const tb = getItemDayTimestamp(b);
  const hasTa = ta !== null;
  const hasTb = tb !== null;
  if (hasTa && hasTb && ta !== tb) return ta - tb;
  if (hasTa !== hasTb) return hasTa ? -1 : 1;

  const ra = Number.isFinite(a.excelRowNumber) ? a.excelRowNumber : Number.MAX_SAFE_INTEGER;
  const rb = Number.isFinite(b.excelRowNumber) ? b.excelRowNumber : Number.MAX_SAFE_INTEGER;
  if (ra !== rb) return ra - rb;

  return String(a.id ?? '').localeCompare(String(b.id ?? ''));
}

/**
 * Tính khoảng cách ngày (dateSpanDays) của một tập giao dịch.
 *
 * @param {object[]} items
 * @returns {number}
 */
function computeDateSpanDays(items) {
  let minTs = Infinity;
  let maxTs = -Infinity;
  let count = 0;

  for (let i = 0; i < items.length; i++) {
    const ts = getItemDayTimestamp(items[i]);
    if (ts !== null) {
      if (ts < minTs) minTs = ts;
      if (ts > maxTs) maxTs = ts;
      count++;
    }
  }

  if (count <= 1) return 0;
  return Math.max(0, Math.round((maxTs - minTs) / MS_PER_DAY));
}

/**
 * So sánh 2 nghiệm theo thứ tự ưu tiên nghiệp vụ kế toán:
 * 1. absDiff tăng dần (khớp chính xác 0 đứng đầu)
 * 2. itemCount tăng dần (ít hóa đơn hơn đứng trước)
 * 3. dateSpanDays tăng dần (các hóa đơn gần ngày nhau hơn đứng trước)
 * 4. Thứ tự dòng Excel tăng dần
 *
 * @param {object} a
 * @param {object} b
 * @returns {number}
 */
function compareSolutions(a, b) {
  if (a.bigAbsDiff !== b.bigAbsDiff) {
    return a.bigAbsDiff < b.bigAbsDiff ? -1 : 1;
  }
  if (a.itemCount !== b.itemCount) return a.itemCount - b.itemCount;
  if (a.dateSpanDays !== b.dateSpanDays) return a.dateSpanDays - b.dateSpanDays;

  const len = Math.min(a.items.length, b.items.length);
  for (let i = 0; i < len; i++) {
    const ra = Number.isFinite(a.items[i].excelRowNumber) ? a.items[i].excelRowNumber : Number.MAX_SAFE_INTEGER;
    const rb = Number.isFinite(b.items[i].excelRowNumber) ? b.items[i].excelRowNumber : Number.MAX_SAFE_INTEGER;
    if (ra !== rb) return ra - rb;
  }

  return a.id.localeCompare(b.id);
}

/**
 * Tiền tính toán bảng cận trên và cận dưới CHÍNH XÁC theo số lượng phần tử chọn bằng BigInt:
 * - maxExact[idx * stride + rem]: tổng đại số lớn nhất khi chọn ĐÚNG `rem` phần tử từ `idx..N-1`
 * - minExact[idx * stride + rem]: tổng đại số nhỏ nhất khi chọn ĐÚNG `rem` phần tử từ `idx..N-1`
 *
 * Độ phức tạp: O(N * maxRem), không bao giờ tràn số dù quy mô hàng triệu tỷ kết hợp 6 số lẻ thập phân!
 *
 * @param {bigint[]} candAmounts
 * @param {number} maxRem
 * @returns {{ maxExact: bigint[], minExact: bigint[], stride: number }}
 */
function buildExactCardinalityBounds(candAmounts, maxRem) {
  const n = candAmounts.length;
  const stride = maxRem + 1;
  const maxExact = new Array(n * stride).fill(0n);
  const minExact = new Array(n * stride).fill(0n);

  if (n === 0 || maxRem <= 0) {
    return { maxExact, minExact, stride };
  }

  const topLargest = new Array(maxRem).fill(0n);
  const topSmallest = new Array(maxRem).fill(0n);
  let count = 0;

  for (let idx = n - 1; idx >= 0; idx--) {
    const val = candAmounts[idx];
    const nextCount = count < maxRem ? count + 1 : maxRem;

    if (count < maxRem || val > topLargest[count - 1]) {
      let pos = nextCount - 1;
      while (pos > 0 && topLargest[pos - 1] < val) {
        topLargest[pos] = topLargest[pos - 1];
        pos--;
      }
      topLargest[pos] = val;
    }

    if (count < maxRem || val < topSmallest[count - 1]) {
      let pos = nextCount - 1;
      while (pos > 0 && topSmallest[pos - 1] > val) {
        topSmallest[pos] = topSmallest[pos - 1];
        pos--;
      }
      topSmallest[pos] = val;
    }

    count = nextCount;

    const base = idx * stride;
    let sumMax = 0n;
    let sumMin = 0n;
    for (let rem = 1; rem <= count; rem++) {
      sumMax += topLargest[rem - 1];
      sumMin += topSmallest[rem - 1];
      maxExact[base + rem] = sumMax;
      minExact[base + rem] = sumMin;
    }
  }

  return { maxExact, minExact, stride };
}

/**
 * Quét TOÀN BỘ `m - c + 1` cửa sổ trượt liên tiếp trên mảng nhóm trùng mệnh giá (đã sắp theo ngày & dòng Excel),
 * tính khoảng cách ngày kết hợp với `pinnedItems` và sắp xếp cửa sổ tối ưu nhất lên đầu.
 *
 * @param {object} grp
 * @param {number} c
 * @param {number} pinnedMinTs
 * @param {number} pinnedMaxTs
 * @param {boolean} hasPinnedTs
 * @returns {Array<{ offset: number, hasTs: boolean, wMinTs: number, wMaxTs: number, ownSpan: number, combinedSpan: number, firstRow: number }>}
 */
function buildSortedGroupWindows(grp, c, pinnedMinTs, pinnedMaxTs, hasPinnedTs) {
  const m = grp.count;
  const items = grp.items;
  const tsArr = grp.dayTimestamps;
  const numWin = m - c + 1;
  const windows = new Array(numWin);

  for (let offset = 0; offset < numWin; offset++) {
    const firstTs = tsArr[offset];
    let hasTs = false;
    let wMinTs = 0;
    let wMaxTs = 0;

    if (firstTs !== null) {
      hasTs = true;
      wMinTs = firstTs;
      const lastTs = tsArr[offset + c - 1];
      if (lastTs !== null) {
        wMaxTs = lastTs;
      } else {
        let p = offset + c - 2;
        while (p > offset && tsArr[p] === null) p--;
        wMaxTs = tsArr[p] !== null ? tsArr[p] : firstTs;
      }
    }

    const ownSpan = hasTs ? wMaxTs - wMinTs : 0;
    let combinedSpan = ownSpan;
    if (hasPinnedTs && hasTs) {
      const mergedMin = pinnedMinTs < wMinTs ? pinnedMinTs : wMinTs;
      const mergedMax = pinnedMaxTs > wMaxTs ? pinnedMaxTs : wMaxTs;
      combinedSpan = mergedMax - mergedMin;
    } else if (hasPinnedTs) {
      combinedSpan = pinnedMaxTs - pinnedMinTs;
    }

    const firstRow = Number.isFinite(items[offset].excelRowNumber)
      ? items[offset].excelRowNumber
      : Number.MAX_SAFE_INTEGER;

    windows[offset] = {
      offset,
      hasTs,
      wMinTs,
      wMaxTs,
      ownSpan,
      combinedSpan,
      firstRow,
    };
  }

  windows.sort((a, b) => {
    if (a.combinedSpan !== b.combinedSpan) return a.combinedSpan - b.combinedSpan;
    if (a.ownSpan !== b.ownSpan) return a.ownSpan - b.ownSpan;
    if (a.firstRow !== b.firstRow) return a.firstRow - b.firstRow;
    return a.offset - b.offset;
  });

  return windows;
}

/**
 * Sinh tập con kích thước `c` từ `grp` dựa trên danh sách cửa sổ tối ưu kết hợp thứ tự từ điển bổ sung.
 *
 * @param {object} grp
 * @param {number} c
 * @param {Array} sortedWindows
 * @param {number} limit
 * @param {boolean} exactAll
 * @returns {object[][]}
 */
function buildGroupSubsets(grp, c, sortedWindows, limit, exactAll = false) {
  const groupItems = grp.items;
  const m = groupItems.length;
  if (c === 0) return [[]];
  if (c === m) return [groupItems.slice()];

  const results = [];

  if (exactAll) {
    const indices = new Int32Array(c);
    for (let i = 0; i < c; i++) indices[i] = i;

    while (results.length < limit) {
      const subset = new Array(c);
      for (let i = 0; i < c; i++) subset[i] = groupItems[indices[i]];
      results.push(subset);

      let pos = c - 1;
      while (pos >= 0 && indices[pos] === m - c + pos) {
        pos--;
      }
      if (pos < 0) break;
      indices[pos]++;
      for (let j = pos + 1; j < c; j++) {
        indices[j] = indices[j - 1] + 1;
      }
    }
    return results;
  }

  const seenKeys = new Set();
  const addByIndices = (idxArr) => {
    const key = idxArr.join(',');
    if (seenKeys.has(key)) return;
    seenKeys.add(key);
    const subset = new Array(c);
    for (let i = 0; i < c; i++) subset[i] = groupItems[idxArr[i]];
    results.push(subset);
  };

  // 1. Lấy các cửa sổ trượt tốt nhất (đã quét toàn bộ m - c + 1 vị trí trên toàn danh sách)
  for (let w = 0; w < sortedWindows.length && results.length < limit; w++) {
    const off = sortedWindows[w].offset;
    const idxArr = new Array(c);
    for (let i = 0; i < c; i++) idxArr[i] = off + i;
    addByIndices(idxArr);
  }

  // 2. Bổ sung thêm các tổ hợp không liên tiếp theo thứ tự từ điển nếu còn chỗ
  if (results.length < limit) {
    const indices = new Array(c);
    for (let i = 0; i < c; i++) indices[i] = i;
    while (results.length < limit) {
      addByIndices(indices);
      let pos = c - 1;
      while (pos >= 0 && indices[pos] === m - c + pos) {
        pos--;
      }
      if (pos < 0) break;
      indices[pos]++;
      for (let j = pos + 1; j < c; j++) {
        indices[j] = indices[j - 1] + 1;
      }
    }
  }

  return results;
}

/**
 * Tạo đối tượng kết quả mặc định khi có lỗi cấu hình hoặc dữ liệu rỗng.
 */
function makeEmptyResult({
  candidateCount = 0,
  pinnedCount = 0,
  exploredNodes = 0,
  elapsedMs = 0,
  timedOut = false,
  error = null,
} = {}) {
  return {
    solutions: [],
    stats: {
      candidateCount,
      pinnedCount,
      exploredNodes,
      elapsedMs,
      timedOut,
      exactCount: 0,
      nearCount: 0,
    },
    error,
  };
}

/**
 * Tìm tổ hợp các khoản tiền khớp mục tiêu trong biên độ sai số ± tolerance.
 * Hỗ trợ đầy đủ cả số dương (phải thu/hóa đơn) và số âm (chiết khấu/trả hàng/cấn trừ),
 * số nguyên lớn hàng trăm nghìn tỷ VNĐ, số thập phân 1..6 chữ số & công thức chia Excel không làm tròn,
 * ghim dòng bắt buộc (pinnedIds), loại trừ dòng (excludedIds), lọc khoảng ngày & từ khóa,
 * tìm kiếm sâu dần theo số lượng phần tử (Iterative Deepening) kèm cắt tỉa cận đúng BigInt và chống bùng nổ trùng mệnh giá.
 *
 * @param {object} options
 * @param {((progress: { exploredNodes: number, foundCount: number, currentSize: number }) => void) | null} onProgress
 * @param {(() => boolean) | null} shouldCancel
 * @returns {object} ReconcileSearchResult
 */
export function findCombinations(options, onProgress = null, shouldCancel = null) {
  const startTime = Date.now();

  if (!options || typeof options !== 'object') {
    return makeEmptyResult({ error: 'Tham số tìm kiếm không hợp lệ.' });
  }

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
    timeLimitMs = DEFAULT_TIME_LIMIT_MS,
  } = options;

  if (!Array.isArray(items)) {
    return makeEmptyResult({ error: 'Danh sách giao dịch không hợp lệ.' });
  }

  if (typeof targetAmount !== 'number' || !Number.isFinite(targetAmount)) {
    return makeEmptyResult({ error: 'Số tiền mục tiêu không hợp lệ.' });
  }

  const safeTolerance = Number.isFinite(tolerance) ? Math.max(0, tolerance) : DEFAULT_TOLERANCE;
  const safeMinItems = Number.isFinite(minItems) ? Math.max(1, Math.floor(minItems)) : 1;
  const rawMaxItems = Number.isFinite(maxItems) ? Math.floor(maxItems) : DEFAULT_MAX_ITEMS;
  if (rawMaxItems < 1 || safeMinItems > rawMaxItems) {
    return makeEmptyResult({
      error: 'Giới hạn số lượng dòng trong tổ hợp không hợp lệ (minItems > maxItems).',
    });
  }
  const safeMaxItems = rawMaxItems;
  const safeMaxResults = Number.isFinite(maxResults) ? Math.max(1, Math.floor(maxResults)) : DEFAULT_MAX_RESULTS;
  const safeTimeLimitMs = Number.isFinite(timeLimitMs) && timeLimitMs > 0 ? timeLimitMs : DEFAULT_TIME_LIMIT_MS;

  const fromIso = typeof dateFromIso === 'string' && dateFromIso.trim() ? dateFromIso.trim() : null;
  const toIso = typeof dateToIso === 'string' && dateToIso.trim() ? dateToIso.trim() : null;
  if (fromIso && toIso && fromIso > toIso) {
    return makeEmptyResult({
      error: 'Khoảng ngày lọc không hợp lệ (Từ ngày lớn hơn Đến ngày).',
    });
  }

  const normKeyword = normalizeVietnameseSearch(keyword);
  const excludedSet = new Set(
    Array.isArray(excludedIds)
      ? excludedIds.filter((id) => id !== null && id !== undefined && id !== '')
      : []
  );
  const uniquePinnedIds = Array.isArray(pinnedIds)
    ? [...new Set(pinnedIds.filter((id) => id !== null && id !== undefined && id !== ''))]
    : [];
  const pinnedSet = new Set(uniquePinnedIds);

  // 1. Phân tách tập dòng ghim (pinned) và tập ứng viên (candidates)
  const pinnedItems = [];
  const seenPinnedIds = new Set();
  const rawCandidates = [];

  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    if (!item || typeof item.amount !== 'number' || !Number.isFinite(item.amount)) {
      continue;
    }

    if (excludedSet.has(item.id)) {
      continue;
    }

    const isPinned = pinnedSet.has(item.id);
    if (isPinned) {
      if (!seenPinnedIds.has(item.id)) {
        seenPinnedIds.add(item.id);
        pinnedItems.push(item);
      }
      continue;
    }

    if (item.isSummaryRow === true && !includeSummaryRows) {
      continue;
    }

    if (fromIso && (!item.dateIso || item.dateIso < fromIso)) {
      continue;
    }
    if (toIso && (!item.dateIso || item.dateIso > toIso)) {
      continue;
    }

    if (normKeyword && !buildSearchHaystack(item).includes(normKeyword)) {
      continue;
    }

    if (item.amount === 0) {
      continue;
    }

    rawCandidates.push(item);
  }

  if (pinnedItems.length < uniquePinnedIds.length) {
    return makeEmptyResult({
      candidateCount: rawCandidates.length,
      pinnedCount: pinnedItems.length,
      elapsedMs: Date.now() - startTime,
    });
  }

  if (pinnedItems.length > safeMaxItems) {
    return makeEmptyResult({
      candidateCount: rawCandidates.length,
      pinnedCount: pinnedItems.length,
      elapsedMs: Date.now() - startTime,
      error: `Số dòng đang ghim (${pinnedItems.length}) vượt quá số dòng tối đa cho phép (${safeMaxItems}).`,
    });
  }

  // 2. Phân tích bậc thập phân chính xác của tập dữ liệu (0..6 chữ số thập phân + phát hiện số thực lặp vô hạn)
  let maxDetectedDecimals = Math.max(
    getDecimalDigitsCount(targetAmount),
    getDecimalDigitsCount(safeTolerance)
  );
  for (let i = 0; i < pinnedItems.length; i++) {
    const d = getDecimalDigitsCount(pinnedItems[i].amount);
    if (d > maxDetectedDecimals) maxDetectedDecimals = d;
  }
  for (let i = 0; i < rawCandidates.length; i++) {
    const d = getDecimalDigitsCount(rawCandidates[i].amount);
    if (d > maxDetectedDecimals) maxDetectedDecimals = d;
  }

  const hasRepeatingFloat = maxDetectedDecimals > MAX_EXACT_DECIMALS;
  const decimals = Math.min(MAX_EXACT_DECIMALS, maxDetectedDecimals);

  const bigTarget = numberToScaledBigInt(targetAmount, decimals);
  const initialBigTolerance = numberToScaledBigInt(safeTolerance, decimals);

  let pinnedBigSum = 0n;
  let pinnedMinTs = Infinity;
  let pinnedMaxTs = -Infinity;
  let hasPinnedTs = false;

  for (let i = 0; i < pinnedItems.length; i++) {
    pinnedBigSum += numberToScaledBigInt(pinnedItems[i].amount, decimals);
    const ts = getItemDayTimestamp(pinnedItems[i]);
    if (ts !== null) {
      if (ts < pinnedMinTs) pinnedMinTs = ts;
      if (ts > pinnedMaxTs) pinnedMaxTs = ts;
      hasPinnedTs = true;
    }
  }

  const candidatesWithBig = [];
  for (let i = 0; i < rawCandidates.length; i++) {
    const item = rawCandidates[i];
    const bigAmount = numberToScaledBigInt(item.amount, decimals);
    if (bigAmount === 0n && !hasRepeatingFloat) continue;
    const absBigAmount = bigAmount < 0n ? -bigAmount : bigAmount;
    candidatesWithBig.push({
      item,
      bigAmount,
      absBigAmount,
      dayTs: getItemDayTimestamp(item),
    });
  }

  // 3. Sắp xếp ứng viên theo |amount| giảm dần, dương trước âm, rồi theo ngày & dòng Excel tăng dần
  candidatesWithBig.sort((a, b) => {
    if (a.absBigAmount !== b.absBigAmount) {
      return a.absBigAmount > b.absBigAmount ? -1 : 1;
    }
    if (a.bigAmount !== b.bigAmount) {
      return a.bigAmount > b.bigAmount ? -1 : 1;
    }
    const hasTa = a.dayTs !== null;
    const hasTb = b.dayTs !== null;
    if (hasTa && hasTb && a.dayTs !== b.dayTs) return a.dayTs - b.dayTs;
    if (hasTa !== hasTb) return hasTa ? -1 : 1;
    const ra = Number.isFinite(a.item.excelRowNumber) ? a.item.excelRowNumber : Number.MAX_SAFE_INTEGER;
    const rb = Number.isFinite(b.item.excelRowNumber) ? b.item.excelRowNumber : Number.MAX_SAFE_INTEGER;
    if (ra !== rb) return ra - rb;
    return String(a.item.id ?? '').localeCompare(String(b.item.id ?? ''));
  });

  const candidateCount = candidatesWithBig.length;
  const pinnedCount = pinnedItems.length;

  const candBigAmounts = new Array(candidateCount);
  for (let i = 0; i < candidateCount; i++) {
    candBigAmounts[i] = candidatesWithBig[i].bigAmount;
  }

  // Gom các ứng viên trùng hệt mệnh giá có dấu thành từng Value Group liên tiếp
  const groups = [];
  for (let i = 0; i < candidateCount; ) {
    const val = candidatesWithBig[i].bigAmount;
    const groupItems = [candidatesWithBig[i].item];
    const dayTimestamps = [candidatesWithBig[i].dayTs];
    let j = i + 1;
    while (j < candidateCount && candidatesWithBig[j].bigAmount === val) {
      groupItems.push(candidatesWithBig[j].item);
      dayTimestamps.push(candidatesWithBig[j].dayTs);
      j++;
    }
    groups.push({
      bigAmount: val,
      absBigAmount: val < 0n ? -val : val,
      formattedLabel: formatScaledBigIntVN(val, decimals),
      startIdx: i,
      count: j - i,
      items: groupItems,
      dayTimestamps,
    });
    i = j;
  }

  const numGroups = groups.length;
  const maxRem = Math.min(candidateCount, safeMaxItems - pinnedCount);
  const { maxExact, minExact, stride } = buildExactCardinalityBounds(candBigAmounts, maxRem);

  // 4. Trạng thái tìm kiếm, kiểm soát thời gian & bộ nhớ thu thập nghiệm
  let exploredNodes = 0;
  let timedOut = false;
  let stopped = false;
  let lastProgressTime = startTime;
  let currentSize = pinnedCount;
  let multisetCounter = 0;

  const collectedSolutions = [];
  const seenSolutionIds = new Set();

  const buildSolutionObject = (rawItems, bigTotal, multisetMeta = null) => {
    const sortedItems = rawItems.slice().sort(compareItemsChronological);
    const itemIds = sortedItems.map((it) => it.id);
    const solKey = itemIds.slice().sort().join('\u0001');
    if (seenSolutionIds.has(solKey)) return null;

    let totalAmount;
    let diff;
    let absDiff;
    let isExact;
    let bigAbsDiff;

    if (!hasRepeatingFloat) {
      const bigDiff = bigTotal - bigTarget;
      bigAbsDiff = bigDiff < 0n ? -bigDiff : bigDiff;
      if (bigAbsDiff > initialBigTolerance) return null;

      totalAmount = scaledBigIntToNumber(bigTotal, decimals);
      diff = scaledBigIntToNumber(bigDiff, decimals);
      absDiff = scaledBigIntToNumber(bigAbsDiff, decimals);
      isExact = bigDiff === 0n;
    } else {
      // Kiểm định lại bằng tổng dấu phẩy động thực tế cho các công thức chia lặp vô hạn (như =1000000/3)
      let rawSum = 0;
      let sumAbsTerms = Math.abs(targetAmount);
      for (let i = 0; i < sortedItems.length; i++) {
        rawSum += sortedItems[i].amount;
        sumAbsTerms += Math.abs(sortedItems[i].amount);
      }
      const floatEps = Math.max(1e-9, Number.EPSILON * sumAbsTerms * 8);
      const rawDiff = rawSum - targetAmount;
      const rawAbsDiff = Math.abs(rawDiff);

      if (rawAbsDiff > safeTolerance + floatEps) {
        return null;
      }

      if (rawAbsDiff <= floatEps) {
        totalAmount = targetAmount === 0 ? 0 : targetAmount;
        diff = 0;
        absDiff = 0;
        isExact = true;
        bigAbsDiff = 0n;
      } else {
        totalAmount = rawSum === 0 ? 0 : rawSum;
        diff = rawDiff === 0 ? 0 : rawDiff;
        absDiff = rawAbsDiff;
        isExact = false;
        bigAbsDiff = numberToScaledBigInt(rawAbsDiff, decimals);
      }
    }

    seenSolutionIds.add(solKey);

    return {
      id: `sol:${itemIds.join(',')}`,
      canonicalKey: solKey,
      items: sortedItems,
      itemIds,
      itemCount: sortedItems.length,
      totalAmount,
      targetAmount: targetAmount === 0 ? 0 : targetAmount,
      diff,
      absDiff,
      isExact,
      dateSpanDays: computeDateSpanDays(sortedItems),
      duplicateVariantsNote: null,
      bigAbsDiff,
      multisetMeta,
    };
  };

  const checkRuntimeLimits = () => {
    const now = Date.now();
    if (shouldCancel && shouldCancel()) {
      stopped = true;
      return;
    }
    if (now - startTime > safeTimeLimitMs) {
      timedOut = true;
      stopped = true;
      return;
    }
    if (onProgress && now - lastProgressTime >= PROGRESS_INTERVAL_MS) {
      lastProgressTime = now;
      onProgress({
        exploredNodes,
        foundCount: collectedSolutions.length,
        currentSize,
      });
    }
  };

  // Kiểm tra trường hợp riêng các dòng ghim (r = 0) đã đủ điều kiện nghiệm
  if (shouldCancel && shouldCancel()) {
    stopped = true;
  } else if (pinnedCount >= Math.max(1, safeMinItems) && pinnedCount <= safeMaxItems) {
    exploredNodes++;
    const repeatSlack0 = hasRepeatingFloat ? BigInt(pinnedCount + 2) : 0n;
    const pinDiff = pinnedBigSum - bigTarget;
    const pinAbsDiff = pinDiff < 0n ? -pinDiff : pinDiff;
    if (pinAbsDiff <= initialBigTolerance + repeatSlack0) {
      const pinnedSol = buildSolutionObject(pinnedItems, pinnedBigSum, null);
      if (pinnedSol) {
        collectedSolutions.push(pinnedSol);
      }
    }
  }

  // Stack lưu cấu trúc multiset đang chọn: (groupIndex, count)
  const chosenGroupIdx = new Int32Array(Math.max(1, maxRem));
  const chosenGroupCnt = new Int32Array(Math.max(1, maxRem));
  let chosenDepth = 0;

  /**
   * Từ một multiset mệnh giá hợp lệ, sinh tối đa MAX_VARIANTS_PER_MULTISET (3) phương án dòng cụ thể
   * tối ưu nhất về khoảng cách ngày (kể cả căn chỉnh đồng bộ ngày giữa nhiều nhóm trùng) và gắn metadata multiset.
   */
  const expandMultisetSolutions = (bigTotal, outBuffer) => {
    let totalVariants = 1;
    let isCapped = false;
    const dupGroupLabels = [];
    const multiChoiceDepthIndices = [];

    for (let d = 0; d < chosenDepth; d++) {
      const grp = groups[chosenGroupIdx[d]];
      const c = chosenGroupCnt[d];
      if (grp.count > c) {
        multiChoiceDepthIndices.push(d);
        const ways = calcBinomialCapped(grp.count, c, 1_000_000);
        if (ways.capped || isCapped) {
          isCapped = true;
          totalVariants = 1_000_000;
        } else {
          const prod = totalVariants * ways.value;
          if (prod >= 1_000_000) {
            isCapped = true;
            totalVariants = 1_000_000;
          } else {
            totalVariants = prod;
          }
        }
        dupGroupLabels.push(grp.formattedLabel);
      }
    }

    // Trường hợp duy nhất 1 tổ hợp dòng cụ thể
    if (!isCapped && totalVariants === 1) {
      const rawItems = pinnedItems.slice();
      for (let d = 0; d < chosenDepth; d++) {
        const grpItems = groups[chosenGroupIdx[d]].items;
        for (let k = 0; k < grpItems.length; k++) {
          rawItems.push(grpItems[k]);
        }
      }
      const sol = buildSolutionObject(rawItems, bigTotal, null);
      if (sol) outBuffer.push(sol);
      return;
    }

    const multisetMeta = {
      multisetId: ++multisetCounter,
      totalVariants,
      isCapped,
      amountsLabel: dupGroupLabels.join('; '),
    };

    const exactEnum = !isCapped && totalVariants <= 36;
    const allGroupWindows = new Array(chosenDepth);
    const groupSubsetsList = new Array(chosenDepth);
    const perGroupLimit = exactEnum ? 36 : multiChoiceDepthIndices.length === 1 ? 48 : 10;

    for (let d = 0; d < chosenDepth; d++) {
      const grp = groups[chosenGroupIdx[d]];
      const c = chosenGroupCnt[d];
      const sortedWin =
        !exactEnum && grp.count > c
          ? buildSortedGroupWindows(grp, c, pinnedMinTs, pinnedMaxTs, hasPinnedTs)
          : [];
      allGroupWindows[d] = sortedWin;
      groupSubsetsList[d] = buildGroupSubsets(grp, c, sortedWin, perGroupLimit, exactEnum);
    }

    const tempSolutions = [];
    const localSeen = new Set();
    const maxComboPool = 64;

    const pushCandidateCombo = (rawItems) => {
      const sortedItems = rawItems.slice().sort(compareItemsChronological);
      const key = sortedItems.map((it) => it.id).sort().join('\u0001');
      if (localSeen.has(key)) return;
      localSeen.add(key);
      tempSolutions.push({
        sortedItems,
        dateSpanDays: computeDateSpanDays(sortedItems),
      });
    };

    // 1. Khi có >= 2 nhóm trùng mệnh giá và không vét cạn, thực hiện căn chỉnh mốc ngày gần nhất giữa các nhóm
    if (!exactEnum && multiChoiceDepthIndices.length >= 2) {
      const anchorDepth = multiChoiceDepthIndices[0];
      const anchorWindows = allGroupWindows[anchorDepth];
      const maxAnchors = Math.min(anchorWindows.length, 24);

      for (let a = 0; a < maxAnchors && tempSolutions.length < maxComboPool; a++) {
        const anchorWin = anchorWindows[a];
        let curHasTs = hasPinnedTs;
        let curMinTs = hasPinnedTs ? pinnedMinTs : Infinity;
        let curMaxTs = hasPinnedTs ? pinnedMaxTs : -Infinity;

        if (anchorWin.hasTs) {
          if (!curHasTs) {
            curHasTs = true;
            curMinTs = anchorWin.wMinTs;
            curMaxTs = anchorWin.wMaxTs;
          } else {
            if (anchorWin.wMinTs < curMinTs) curMinTs = anchorWin.wMinTs;
            if (anchorWin.wMaxTs > curMaxTs) curMaxTs = anchorWin.wMaxTs;
          }
        }

        const alignedItems = pinnedItems.slice();

        for (let d = 0; d < chosenDepth; d++) {
          const grp = groups[chosenGroupIdx[d]];
          const c = chosenGroupCnt[d];

          if (grp.count === c) {
            for (let k = 0; k < grp.items.length; k++) alignedItems.push(grp.items[k]);
            continue;
          }

          let chosenWin = anchorWin;
          if (d !== anchorDepth) {
            const dWindows = allGroupWindows[d];
            let bestWin = dWindows[0];
            let bestSpan = Infinity;

            for (let w = 0; w < dWindows.length; w++) {
              const candW = dWindows[w];
              let span = candW.ownSpan;
              if (curHasTs && candW.hasTs) {
                const mMin = curMinTs < candW.wMinTs ? curMinTs : candW.wMinTs;
                const mMax = curMaxTs > candW.wMaxTs ? curMaxTs : candW.wMaxTs;
                span = mMax - mMin;
              } else if (curHasTs) {
                span = curMaxTs - curMinTs;
              }

              if (
                span < bestSpan ||
                (span === bestSpan &&
                  (candW.ownSpan < bestWin.ownSpan ||
                    (candW.ownSpan === bestWin.ownSpan && candW.firstRow < bestWin.firstRow)))
              ) {
                bestSpan = span;
                bestWin = candW;
                if (span === 0 && candW.ownSpan === 0) break;
              }
            }

            chosenWin = bestWin;
            if (chosenWin.hasTs) {
              if (!curHasTs) {
                curHasTs = true;
                curMinTs = chosenWin.wMinTs;
                curMaxTs = chosenWin.wMaxTs;
              } else {
                if (chosenWin.wMinTs < curMinTs) curMinTs = chosenWin.wMinTs;
                if (chosenWin.wMaxTs > curMaxTs) curMaxTs = chosenWin.wMaxTs;
              }
            }
          }

          for (let k = 0; k < c; k++) {
            alignedItems.push(grp.items[chosenWin.offset + k]);
          }
        }

        pushCandidateCombo(alignedItems);
      }
    }

    // 2. Ghép tổ hợp chéo từ các tập con đầu bảng (khử trùng trực tiếp ngay khi đẩy vào)
    const combineGroups = (depthIdx, accItems) => {
      if (tempSolutions.length >= maxComboPool) return;
      if (depthIdx === chosenDepth) {
        pushCandidateCombo(accItems);
        return;
      }
      const subsets = groupSubsetsList[depthIdx];
      for (let s = 0; s < subsets.length; s++) {
        const sub = subsets[s];
        const prevLen = accItems.length;
        for (let j = 0; j < sub.length; j++) accItems.push(sub[j]);
        combineGroups(depthIdx + 1, accItems);
        accItems.length = prevLen;
        if (tempSolutions.length >= maxComboPool) return;
      }
    };

    combineGroups(0, pinnedItems.slice());

    tempSolutions.sort((a, b) => {
      if (a.dateSpanDays !== b.dateSpanDays) return a.dateSpanDays - b.dateSpanDays;
      const len = Math.min(a.sortedItems.length, b.sortedItems.length);
      for (let i = 0; i < len; i++) {
        const ra = Number.isFinite(a.sortedItems[i].excelRowNumber)
          ? a.sortedItems[i].excelRowNumber
          : Number.MAX_SAFE_INTEGER;
        const rb = Number.isFinite(b.sortedItems[i].excelRowNumber)
          ? b.sortedItems[i].excelRowNumber
          : Number.MAX_SAFE_INTEGER;
        if (ra !== rb) return ra - rb;
      }
      return 0;
    });

    const emitCount = Math.min(tempSolutions.length, MAX_VARIANTS_PER_MULTISET);
    for (let i = 0; i < emitCount; i++) {
      const sol = buildSolutionObject(tempSolutions[i].sortedItems, bigTotal, multisetMeta);
      if (sol) outBuffer.push(sol);
    }
  };

  // 5. Iterative Deepening theo tổng số phần tử k = max(1, minItems, pinnedCount + 1) .. maxItems
  const startK = Math.max(1, safeMinItems, pinnedCount + 1);
  const endK = Math.min(safeMaxItems, pinnedCount + candidateCount);
  let globalActiveBigTolerance = initialBigTolerance;

  const keepPerLevel = Math.max(100, safeMaxResults * 3);
  const compactThreshold = keepPerLevel * 2;

  for (let k = startK; k <= endK && !stopped; k++) {
    // Chỉ khi chuyển sang mức k mới (itemCount lớn hơn nghiêm ngặt) mới được phép dừng sớm hoặc trừ 1 đơn vị sai số
    if (!hasRepeatingFloat && collectedSolutions.length >= safeMaxResults) {
      collectedSolutions.sort(compareSolutions);
      if (collectedSolutions.length > keepPerLevel * 2) {
        collectedSolutions.length = keepPerLevel * 2;
      }
      const kthBest = collectedSolutions[safeMaxResults - 1];
      if (kthBest.bigAbsDiff === 0n) {
        break;
      }
      if (kthBest.bigAbsDiff - 1n < globalActiveBigTolerance) {
        globalActiveBigTolerance = kthBest.bigAbsDiff - 1n;
      }
    }

    currentSize = k;
    checkRuntimeLimits();
    if (stopped) break;

    const r = k - pinnedCount;
    const repeatSlack = hasRepeatingFloat ? BigInt(k + 2) : 0n;
    let levelBigTolerance = globalActiveBigTolerance;
    let levelLowBound = bigTarget - levelBigTolerance - repeatSlack;
    let levelHighBound = bigTarget + levelBigTolerance + repeatSlack;
    const levelSolutions = [];

    // Trong cùng tầng k: KHÔNG dừng sớm và KHÔNG trừ 1 đơn vị sai số để bảo toàn tuyệt đối thứ tự ưu tiên dateSpanDays & excelRowNumber!
    const compactLevelIfNeeded = () => {
      if (levelSolutions.length < compactThreshold) return;
      levelSolutions.sort(compareSolutions);
      levelSolutions.length = keepPerLevel;
      const worstKept = levelSolutions[keepPerLevel - 1];
      if (!hasRepeatingFloat && worstKept.bigAbsDiff < levelBigTolerance) {
        levelBigTolerance = worstKept.bigAbsDiff;
        levelLowBound = bigTarget - levelBigTolerance;
        levelHighBound = bigTarget + levelBigTolerance;
      }
    };

    const dfs = (groupIdx, rem, currentSum) => {
      exploredNodes++;
      if ((exploredNodes & NODE_CHECK_MASK) === 0) {
        checkRuntimeLimits();
      }
      if (stopped || groupIdx >= numGroups) return;

      const grp = groups[groupIdx];
      const startIdx = grp.startIdx;
      const avail = candidateCount - startIdx;
      if (avail < rem) return;

      const boundIdx = startIdx * stride + rem;
      if (currentSum + maxExact[boundIdx] < levelLowBound) return;
      if (currentSum + minExact[boundIdx] > levelHighBound) return;

      // Nhánh lá tối ưu khi chỉ cần chọn thêm đúng 1 phần tử (rem === 1)
      if (rem === 1) {
        let needLow = levelLowBound - currentSum;
        let needHigh = levelHighBound - currentSum;
        let minReqAbs = needLow > 0n ? needLow : needHigh < 0n ? -needHigh : 0n;

        for (let j = groupIdx; j < numGroups; j++) {
          if (j > groupIdx) {
            exploredNodes++;
            if ((exploredNodes & NODE_CHECK_MASK) === 0) {
              checkRuntimeLimits();
            }
            if (stopped) return;
          }

          const g = groups[j];
          if (g.absBigAmount < minReqAbs) break;

          if (g.bigAmount >= needLow && g.bigAmount <= needHigh) {
            chosenGroupIdx[chosenDepth] = j;
            chosenGroupCnt[chosenDepth] = 1;
            chosenDepth++;
            expandMultisetSolutions(currentSum + g.bigAmount, levelSolutions);
            chosenDepth--;

            if (levelSolutions.length >= compactThreshold) {
              compactLevelIfNeeded();
              needLow = levelLowBound - currentSum;
              needHigh = levelHighBound - currentSum;
              minReqAbs = needLow > 0n ? needLow : needHigh < 0n ? -needHigh : 0n;
              if (g.absBigAmount < minReqAbs) {
                break;
              }
            }
          }
        }
        return;
      }

      // Trường hợp tổng quát rem >= 2: duyệt số lượng c phần tử lấy từ nhóm `groupIdx`
      const nextGroupIdx = groupIdx + 1;
      const nextStartIdx = startIdx + grp.count;
      const availAfterGroup = candidateCount - nextStartIdx;
      const maxC = grp.count < rem ? grp.count : rem;
      const minC = rem - availAfterGroup > 0 ? rem - availAfterGroup : 0;

      if (minC > maxC) return;

      for (let c = maxC; c >= minC; c--) {
        if (stopped) return;

        const nextRem = rem - c;
        const nextSum = currentSum + BigInt(c) * grp.bigAmount;

        if (nextRem === 0) {
          if (nextSum >= levelLowBound && nextSum <= levelHighBound) {
            chosenGroupIdx[chosenDepth] = groupIdx;
            chosenGroupCnt[chosenDepth] = c;
            chosenDepth++;
            expandMultisetSolutions(nextSum, levelSolutions);
            chosenDepth--;
            if (levelSolutions.length >= compactThreshold) {
              compactLevelIfNeeded();
            }
          }
          continue;
        }

        if (nextGroupIdx >= numGroups) continue;

        const nextBoundIdx = nextStartIdx * stride + nextRem;
        if (nextSum + maxExact[nextBoundIdx] < levelLowBound) continue;
        if (nextSum + minExact[nextBoundIdx] > levelHighBound) continue;

        if (c > 0) {
          chosenGroupIdx[chosenDepth] = groupIdx;
          chosenGroupCnt[chosenDepth] = c;
          chosenDepth++;
          dfs(nextGroupIdx, nextRem, nextSum);
          chosenDepth--;
        } else {
          dfs(nextGroupIdx, nextRem, nextSum);
        }
      }
    };

    dfs(0, r, pinnedBigSum);

    for (let i = 0; i < levelSolutions.length; i++) {
      collectedSolutions.push(levelSolutions[i]);
    }

    if (onProgress && !stopped) {
      const now = Date.now();
      if (now - lastProgressTime >= PROGRESS_INTERVAL_MS) {
        lastProgressTime = now;
        onProgress({
          exploredNodes,
          foundCount: collectedSolutions.length,
          currentSize,
        });
      }
    }
  }

  // 6. Sắp xếp nghiệm cuối cùng, cắt top `maxResults`, rồi mới tính chính xác số biến thể trùng bị ẩn thực tế
  collectedSolutions.sort(compareSolutions);
  const topSolutions = collectedSolutions.slice(0, safeMaxResults);

  const shownCountByMultiset = new Map();
  for (let i = 0; i < topSolutions.length; i++) {
    const meta = topSolutions[i].multisetMeta;
    if (meta) {
      shownCountByMultiset.set(
        meta.multisetId,
        (shownCountByMultiset.get(meta.multisetId) || 0) + 1
      );
    }
  }

  let exactCount = 0;
  let nearCount = 0;

  const finalSolutions = topSolutions.map((sol) => {
    if (sol.isExact) {
      exactCount++;
    } else {
      nearCount++;
    }

    let duplicateVariantsNote = null;
    const meta = sol.multisetMeta;
    if (meta) {
      const shownCount = shownCountByMultiset.get(meta.multisetId) || 1;
      if (meta.isCapped) {
        duplicateVariantsNote = `Còn nhiều cách chọn tương đương từ các dòng cùng số tiền (${meta.amountsLabel}) có thể thay thế`;
      } else if (meta.totalVariants > shownCount) {
        const omitted = meta.totalVariants - shownCount;
        duplicateVariantsNote = `Còn ${formatIntVN(omitted)} cách chọn tương đương từ các dòng cùng số tiền (${meta.amountsLabel}) có thể thay thế`;
      }
    }

    const outSol = {
      id: sol.id,
      items: sol.items,
      itemIds: sol.itemIds,
      itemCount: sol.itemCount,
      totalAmount: sol.totalAmount,
      targetAmount: sol.targetAmount,
      diff: sol.diff,
      absDiff: sol.absDiff,
      isExact: sol.isExact,
      dateSpanDays: sol.dateSpanDays,
      duplicateVariantsNote,
    };
    Object.defineProperty(outSol, 'canonicalKey', {
      value: sol.canonicalKey,
      enumerable: false,
      configurable: true,
    });
    return outSol;
  });

  const elapsedMs = Math.max(0, Date.now() - startTime);

  if (onProgress) {
    onProgress({
      exploredNodes,
      foundCount: finalSolutions.length,
      currentSize,
    });
  }

  return {
    solutions: finalSolutions,
    stats: {
      candidateCount,
      pinnedCount,
      exploredNodes,
      elapsedMs,
      timedOut,
      exactCount,
      nearCount,
    },
    error: null,
  };
}
