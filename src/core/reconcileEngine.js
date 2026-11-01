import {
  DEFAULT_TOLERANCE,
  DEFAULT_MAX_ITEMS,
  DEFAULT_MAX_RESULTS,
  DEFAULT_TIME_LIMIT_MS,
} from './contracts.js';

const MS_PER_DAY = 86_400_000;
const MAX_VARIANTS_PER_MULTISET = 3;
const PROGRESS_INTERVAL_MS = 60;
const NODE_CHECK_MASK = 2047; // Check clock & cancellation every 2048 nodes

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
 * Kiểm tra một số thực có phần thập phân đáng kể hay không.
 *
 * @param {number} x
 * @returns {boolean}
 */
function hasDecimalFraction(x) {
  return Math.abs(x - Math.round(x)) > 1e-6;
}

/**
 * Quy đổi ngược từ đơn vị nguyên (đã nhân scale) về số thực chuẩn xác.
 *
 * @param {number} intVal
 * @param {number} scale
 * @returns {number}
 */
function unscaleAmount(intVal, scale) {
  if (scale === 1) return intVal === 0 ? 0 : intVal;
  const val = Number((intVal / scale).toFixed(2));
  return val === 0 ? 0 : val;
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
 * Định dạng số tiền VN phục vụ ghi chú dòng trùng mệnh giá.
 *
 * @param {number} value
 * @returns {string}
 */
function formatAmountVN(value) {
  if (!Number.isFinite(value) || value === 0) return '0';
  const sign = value < 0 ? '-' : '';
  const abs = Math.abs(value);
  const rounded = Math.round(abs * 100) / 100;
  if (Math.abs(rounded - Math.round(rounded)) < 1e-9) {
    return sign + String(Math.round(rounded)).replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  }
  const [intPart, decPart] = rounded.toFixed(2).replace(/0+$/, '').split('.');
  const formattedInt = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  return `${sign}${formattedInt},${decPart}`;
}

/**
 * Tính tổ hợp C(n, k) có giới hạn trần để không tràn số nguyên an toàn.
 *
 * @param {number} n
 * @param {number} k
 * @param {number} cap
 * @returns {number}
 */
function safeBinomial(n, k, cap = 1_000_000) {
  if (k < 0 || k > n) return 0;
  if (k === 0 || k === n) return 1;
  const kk = Math.min(k, n - k);
  let res = 1;
  for (let i = 1; i <= kk; i++) {
    res = Math.round((res * (n - kk + i)) / i);
    if (res >= cap) return cap;
  }
  return res;
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
  if (a.intAbsDiff !== b.intAbsDiff) return a.intAbsDiff - b.intAbsDiff;
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
 * Tiền tính toán bảng cận trên và cận dưới CHÍNH XÁC theo số lượng phần tử chọn (Exact Cardinality Bounds):
 * - maxExact[idx * stride + rem]: tổng đại số lớn nhất có thể khi chọn ĐÚNG `rem` phần tử từ `idx..N-1`
 * - minExact[idx * stride + rem]: tổng đại số nhỏ nhất có thể khi chọn ĐÚNG `rem` phần tử từ `idx..N-1`
 *
 * Độ phức tạp: O(N * maxRem) về thời gian và bộ nhớ — chạy < 1ms cho hàng nghìn dòng,
 * đúng tuyệt đối 100% cho cả số dương, số âm và hỗn hợp dương/âm!
 *
 * @param {Float64Array} candAmounts
 * @param {number} maxRem
 * @returns {{ maxExact: Float64Array, minExact: Float64Array, stride: number }}
 */
function buildExactCardinalityBounds(candAmounts, maxRem) {
  const n = candAmounts.length;
  const stride = maxRem + 1;
  const maxExact = new Float64Array(n * stride);
  const minExact = new Float64Array(n * stride);

  if (n === 0 || maxRem <= 0) {
    return { maxExact, minExact, stride };
  }

  const topLargest = new Float64Array(maxRem);  // giảm dần: topLargest[0] >= topLargest[1] >= ...
  const topSmallest = new Float64Array(maxRem); // tăng dần: topSmallest[0] <= topSmallest[1] <= ...
  let count = 0;

  for (let idx = n - 1; idx >= 0; idx--) {
    const val = candAmounts[idx];
    const nextCount = count < maxRem ? count + 1 : maxRem;

    // Cập nhật mảng `maxRem` phần tử lớn nhất về mặt đại số
    if (count < maxRem || val > topLargest[count - 1]) {
      let pos = nextCount - 1;
      while (pos > 0 && topLargest[pos - 1] < val) {
        topLargest[pos] = topLargest[pos - 1];
        pos--;
      }
      topLargest[pos] = val;
    }

    // Cập nhật mảng `maxRem` phần tử nhỏ nhất về mặt đại số
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
    let sumMax = 0;
    let sumMin = 0;
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
 * Sinh tối đa `limit` tập con kích thước `c` từ `groupItems` (đã sắp xếp tăng dần theo ngày & dòng Excel).
 * Ưu tiên các cửa sổ trượt liên tiếp (tối ưu khoảng cách ngày nhỏ nhất) kết hợp thứ tự từ điển đầu tiên.
 *
 * @param {object[]} groupItems
 * @param {number} c
 * @param {number} limit
 * @param {boolean} exactAll
 * @returns {object[][]}
 */
function sampleGroupSubsets(groupItems, c, limit, exactAll = false) {
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

  // 1. Quét cửa sổ trượt liên tiếp theo thời gian để bắt đúng cụm hóa đơn cùng ngày / gần ngày nhất
  const maxSlide = Math.min(m - c, Math.max(limit, 48));
  for (let offset = 0; offset <= maxSlide && results.length < Math.max(limit, 48); offset++) {
    const idxArr = new Array(c);
    for (let i = 0; i < c; i++) idxArr[i] = offset + i;
    addByIndices(idxArr);
  }

  // 2. Bổ sung thêm các tổ hợp đầu theo thứ tự từ điển nếu còn chỗ
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
 * ghim dòng bắt buộc (pinnedIds), loại trừ dòng (excludedIds), lọc khoảng ngày & từ khóa,
 * tìm kiếm sâu dần theo số lượng phần tử (Iterative Deepening) kèm cắt tỉa cận đúng và chống bùng nổ trùng mệnh giá.
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

  // Nếu có pinnedId bị loại trừ hoặc không tồn tại trong dữ liệu thì không thể có nghiệm chứa đủ pinnedIds
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

  // 2. Xác định hệ số nhân nguyên (scale = 1 cho số nguyên VNĐ, scale = 100 nếu có phần thập phân)
  let useDecimalScale = hasDecimalFraction(targetAmount) || hasDecimalFraction(safeTolerance);
  if (!useDecimalScale) {
    for (let i = 0; i < pinnedItems.length; i++) {
      if (hasDecimalFraction(pinnedItems[i].amount)) {
        useDecimalScale = true;
        break;
      }
    }
  }
  if (!useDecimalScale) {
    for (let i = 0; i < rawCandidates.length; i++) {
      if (hasDecimalFraction(rawCandidates[i].amount)) {
        useDecimalScale = true;
        break;
      }
    }
  }

  const scale = useDecimalScale ? 100 : 1;
  const intTarget = Math.round(targetAmount * scale);
  const initialIntTolerance = Math.round(safeTolerance * scale);

  let pinnedIntSum = 0;
  for (let i = 0; i < pinnedItems.length; i++) {
    pinnedIntSum += Math.round(pinnedItems[i].amount * scale);
  }

  // Lọc bỏ ứng viên có intAmount === 0 sau khi quy đổi và gắn metadata số nguyên
  const candidatesWithInt = [];
  for (let i = 0; i < rawCandidates.length; i++) {
    const item = rawCandidates[i];
    const intAmount = Math.round(item.amount * scale);
    if (intAmount === 0) continue;
    candidatesWithInt.push({
      item,
      intAmount,
      absIntAmount: Math.abs(intAmount),
      dayTs: getItemDayTimestamp(item),
    });
  }

  // 3. Sắp xếp ứng viên theo |amount| giảm dần, dương trước âm, rồi theo ngày & dòng Excel tăng dần
  candidatesWithInt.sort((a, b) => {
    if (a.absIntAmount !== b.absIntAmount) return b.absIntAmount - a.absIntAmount;
    if (a.intAmount !== b.intAmount) return b.intAmount - a.intAmount;
    const hasTa = a.dayTs !== null;
    const hasTb = b.dayTs !== null;
    if (hasTa && hasTb && a.dayTs !== b.dayTs) return a.dayTs - b.dayTs;
    if (hasTa !== hasTb) return hasTa ? -1 : 1;
    const ra = Number.isFinite(a.item.excelRowNumber) ? a.item.excelRowNumber : Number.MAX_SAFE_INTEGER;
    const rb = Number.isFinite(b.item.excelRowNumber) ? b.item.excelRowNumber : Number.MAX_SAFE_INTEGER;
    if (ra !== rb) return ra - rb;
    return String(a.item.id ?? '').localeCompare(String(b.item.id ?? ''));
  });

  const candidateCount = candidatesWithInt.length;
  const pinnedCount = pinnedItems.length;

  // Mảng phẳng phục vụ bảng cận Exact Cardinality Bounds
  const candAmounts = new Float64Array(candidateCount);
  for (let i = 0; i < candidateCount; i++) {
    candAmounts[i] = candidatesWithInt[i].intAmount;
  }

  // Gom các ứng viên trùng hệt mệnh giá có dấu thành từng Value Group liên tiếp để chống bùng nổ tổ hợp
  const groups = [];
  for (let i = 0; i < candidateCount; ) {
    const val = candidatesWithInt[i].intAmount;
    const groupItems = [candidatesWithInt[i].item];
    let j = i + 1;
    while (j < candidateCount && candidatesWithInt[j].intAmount === val) {
      groupItems.push(candidatesWithInt[j].item);
      j++;
    }
    groups.push({
      intAmount: val,
      absIntAmount: Math.abs(val),
      displayAmount: unscaleAmount(val, scale),
      startIdx: i,
      count: j - i,
      items: groupItems,
    });
    i = j;
  }

  const numGroups = groups.length;
  const maxRem = Math.min(candidateCount, safeMaxItems - pinnedCount);
  const { maxExact, minExact, stride } = buildExactCardinalityBounds(candAmounts, maxRem);

  // 4. Trạng thái tìm kiếm, kiểm soát thời gian & bộ nhớ thu thập nghiệm
  let exploredNodes = 0;
  let timedOut = false;
  let stopped = false;
  let lastProgressTime = startTime;
  let currentSize = pinnedCount;

  const collectedSolutions = [];
  const seenSolutionIds = new Set();

  const buildSolutionObject = (rawItems, intTotal, duplicateVariantsNote = null) => {
    const sortedItems = rawItems.slice().sort(compareItemsChronological);
    const itemIds = sortedItems.map((it) => it.id);
    const solKey = itemIds.slice().sort().join('\u0001');
    if (seenSolutionIds.has(solKey)) return null;
    seenSolutionIds.add(solKey);

    const intDiff = intTotal - intTarget;
    const intAbsDiff = Math.abs(intDiff);

    return {
      id: `sol:${itemIds.join(',')}`,
      items: sortedItems,
      itemIds,
      itemCount: sortedItems.length,
      totalAmount: unscaleAmount(intTotal, scale),
      targetAmount,
      diff: unscaleAmount(intDiff, scale),
      absDiff: unscaleAmount(intAbsDiff, scale),
      isExact: intDiff === 0,
      dateSpanDays: computeDateSpanDays(sortedItems),
      duplicateVariantsNote,
      intAbsDiff,
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
    if (Math.abs(pinnedIntSum - intTarget) <= initialIntTolerance) {
      const pinnedSol = buildSolutionObject(pinnedItems, pinnedIntSum, null);
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
   * tốt nhất (khoảng cách ngày ngắn nhất, dòng Excel sớm nhất) và đính kèm ghi chú nếu rút gọn trùng lặp.
   */
  const expandMultisetSolutions = (intTotal, outBuffer) => {
    let totalVariants = 1;
    const dupGroupLabels = [];
    let multiChoiceGroupCount = 0;

    for (let d = 0; d < chosenDepth; d++) {
      const grp = groups[chosenGroupIdx[d]];
      const c = chosenGroupCnt[d];
      if (grp.count > c) {
        multiChoiceGroupCount++;
        const ways = safeBinomial(grp.count, c, 1_000_000);
        totalVariants = Math.min(1_000_000, totalVariants * ways);
        dupGroupLabels.push(formatAmountVN(grp.displayAmount));
      }
    }

    // Trường hợp duy nhất 1 tổ hợp dòng cụ thể
    if (totalVariants === 1) {
      const rawItems = pinnedItems.slice();
      for (let d = 0; d < chosenDepth; d++) {
        const grpItems = groups[chosenGroupIdx[d]].items;
        for (let k = 0; k < grpItems.length; k++) {
          rawItems.push(grpItems[k]);
        }
      }
      const sol = buildSolutionObject(rawItems, intTotal, null);
      if (sol) outBuffer.push(sol);
      return;
    }

    // Sinh các tập con ứng viên cho từng nhóm mệnh giá
    const exactEnum = totalVariants <= 36;
    const perGroupLimit = exactEnum ? 36 : multiChoiceGroupCount === 1 ? 48 : 8;
    const groupSubsetsList = new Array(chosenDepth);

    for (let d = 0; d < chosenDepth; d++) {
      const grp = groups[chosenGroupIdx[d]];
      const c = chosenGroupCnt[d];
      groupSubsetsList[d] = sampleGroupSubsets(grp.items, c, perGroupLimit, exactEnum);
    }

    // Kết hợp thành danh sách tổ hợp dòng ứng viên (tối đa 48 biến thể để chọn top 3 tốt nhất)
    const candidateCombos = [];
    const maxComboPool = 48;

    const combineGroups = (depthIdx, accItems) => {
      if (candidateCombos.length >= maxComboPool) return;
      if (depthIdx === chosenDepth) {
        candidateCombos.push(accItems.slice());
        return;
      }
      const subsets = groupSubsetsList[depthIdx];
      for (let s = 0; s < subsets.length; s++) {
        const sub = subsets[s];
        const prevLen = accItems.length;
        for (let j = 0; j < sub.length; j++) accItems.push(sub[j]);
        combineGroups(depthIdx + 1, accItems);
        accItems.length = prevLen;
        if (candidateCombos.length >= maxComboPool) return;
      }
    };

    // Nếu có nhiều nhóm trùng đồng thời và không vét cạn, ưu tiên ghép đồng bộ theo chỉ số thời gian trước
    if (!exactEnum && multiChoiceGroupCount > 1) {
      const maxAligned = 16;
      for (let alignIdx = 0; alignIdx < maxAligned && candidateCombos.length < maxComboPool; alignIdx++) {
        const aligned = pinnedItems.slice();
        for (let d = 0; d < chosenDepth; d++) {
          const subsets = groupSubsetsList[d];
          const chosenSub = subsets[Math.min(alignIdx, subsets.length - 1)];
          for (let j = 0; j < chosenSub.length; j++) aligned.push(chosenSub[j]);
        }
        candidateCombos.push(aligned);
      }
    }

    combineGroups(0, pinnedItems.slice());

    // Xây dựng các nghiệm tạm, sắp xếp theo khoảng cách ngày ngắn nhất & dòng Excel sớm nhất
    const tempSolutions = [];
    const localSeen = new Set();

    for (let i = 0; i < candidateCombos.length; i++) {
      const sortedItems = candidateCombos[i].slice().sort(compareItemsChronological);
      const key = sortedItems.map((it) => it.id).sort().join('\u0001');
      if (localSeen.has(key)) continue;
      localSeen.add(key);

      tempSolutions.push({
        sortedItems,
        dateSpanDays: computeDateSpanDays(sortedItems),
      });
    }

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
    let note = null;

    if (totalVariants > emitCount) {
      const omitted = totalVariants - emitCount;
      const amountsLabel = dupGroupLabels.join('; ');
      if (totalVariants < 1_000_000) {
        note = `Còn ${formatIntVN(omitted)} cách chọn tương đương từ các dòng cùng số tiền (${amountsLabel}) có thể thay thế`;
      } else {
        note = `Còn nhiều cách chọn tương đương từ các dòng cùng số tiền (${amountsLabel}) có thể thay thế`;
      }
    }

    for (let i = 0; i < emitCount; i++) {
      const sol = buildSolutionObject(tempSolutions[i].sortedItems, intTotal, note);
      if (sol) outBuffer.push(sol);
    }
  };

  // 5. Iterative Deepening theo tổng số phần tử k = max(1, minItems, pinnedCount + 1) .. maxItems
  const startK = Math.max(1, safeMinItems, pinnedCount + 1);
  const endK = Math.min(safeMaxItems, pinnedCount + candidateCount);
  let globalActiveIntTolerance = initialIntTolerance;

  const keepPerLevel = Math.max(100, safeMaxResults * 2);
  const compactThreshold = keepPerLevel * 2;

  for (let k = startK; k <= endK && !stopped; k++) {
    // Trước khi sang mức k lớn hơn, kiểm tra xem tập nghiệm ở các mức < k đã tối ưu tuyệt đối chưa
    if (collectedSolutions.length >= safeMaxResults) {
      collectedSolutions.sort(compareSolutions);
      if (collectedSolutions.length > safeMaxResults * 4) {
        collectedSolutions.length = safeMaxResults * 4;
      }
      const kthBest = collectedSolutions[safeMaxResults - 1];
      if (kthBest.intAbsDiff === 0) {
        // Đã có đủ >= maxResults nghiệm khớp chính xác 100% với ít dòng hơn -> không nghiệm nào ở mức >= k có thể vượt qua!
        break;
      }
      if (kthBest.intAbsDiff - 1 < globalActiveIntTolerance) {
        globalActiveIntTolerance = kthBest.intAbsDiff - 1;
      }
    }

    currentSize = k;
    checkRuntimeLimits();
    if (stopped) break;

    const r = k - pinnedCount;
    let levelIntTolerance = globalActiveIntTolerance;
    let levelLowBound = intTarget - levelIntTolerance;
    let levelHighBound = intTarget + levelIntTolerance;
    let stopCurrentLevel = false;
    const levelSolutions = [];

    const compactLevelIfNeeded = () => {
      if (levelSolutions.length < compactThreshold) return;
      levelSolutions.sort(compareSolutions);
      levelSolutions.length = keepPerLevel;
      const worstKept = levelSolutions[keepPerLevel - 1];
      if (worstKept.intAbsDiff === 0) {
        // Đã thu thập rất nhiều nghiệm khớp tuyệt đối ở mức k hiện tại
        stopCurrentLevel = true;
      } else if (worstKept.intAbsDiff - 1 < levelIntTolerance) {
        levelIntTolerance = worstKept.intAbsDiff - 1;
        levelLowBound = intTarget - levelIntTolerance;
        levelHighBound = intTarget + levelIntTolerance;
      }
    };

    const dfs = (groupIdx, rem, currentSum) => {
      exploredNodes++;
      if ((exploredNodes & NODE_CHECK_MASK) === 0) {
        checkRuntimeLimits();
      }
      if (stopped || stopCurrentLevel || groupIdx >= numGroups) return;

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
        let minReqAbs = needLow > 0 ? needLow : needHigh < 0 ? -needHigh : 0;

        for (let j = groupIdx; j < numGroups; j++) {
          if (j > groupIdx) {
            exploredNodes++;
            if ((exploredNodes & NODE_CHECK_MASK) === 0) {
              checkRuntimeLimits();
            }
            if (stopped || stopCurrentLevel) return;
          }

          const g = groups[j];
          // Do các nhóm đã sắp xếp giảm dần theo |intAmount|, có thể cắt sớm ngay khi |val| < minReqAbs
          if (g.absIntAmount < minReqAbs) break;

          if (g.intAmount >= needLow && g.intAmount <= needHigh) {
            chosenGroupIdx[chosenDepth] = j;
            chosenGroupCnt[chosenDepth] = 1;
            chosenDepth++;
            expandMultisetSolutions(currentSum + g.intAmount, levelSolutions);
            chosenDepth--;

            if (levelSolutions.length >= compactThreshold) {
              compactLevelIfNeeded();
              if (stopCurrentLevel) return;
              // Cập nhật lại biên tìm kiếm nếu levelIntTolerance vừa được siết chặt
              needLow = levelLowBound - currentSum;
              needHigh = levelHighBound - currentSum;
              minReqAbs = needLow > 0 ? needLow : needHigh < 0 ? -needHigh : 0;
              if (g.absIntAmount < minReqAbs) {
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
        if (stopped || stopCurrentLevel) return;

        const nextRem = rem - c;
        const nextSum = currentSum + c * grp.intAmount;

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

    dfs(0, r, pinnedIntSum);

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

  // 6. Sắp xếp nghiệm cuối cùng, cắt top `maxResults` và gỡ trường nội bộ `intAbsDiff`
  collectedSolutions.sort(compareSolutions);
  const topSolutions = collectedSolutions.slice(0, safeMaxResults);

  let exactCount = 0;
  let nearCount = 0;

  const finalSolutions = topSolutions.map((sol) => {
    if (sol.isExact) {
      exactCount++;
    } else {
      nearCount++;
    }
    return {
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
      duplicateVariantsNote: sol.duplicateVariantsNote,
    };
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
