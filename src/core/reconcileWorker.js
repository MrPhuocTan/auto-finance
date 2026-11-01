import { findCombinations } from './reconcileEngine.js';

const getWorkerScope = () => (typeof self !== 'undefined' ? self : globalThis);
const cancelledRequestIds = new Set();
const pendingSearchQueue = [];
let activeRequestId = null;
let drainTimer = null;
let isExecutingSearch = false;

function markCancelled(id) {
  if (id === undefined || id === null) return;
  cancelledRequestIds.add(id);
  if (cancelledRequestIds.size > 256) {
    const oldest = cancelledRequestIds.values().next().value;
    cancelledRequestIds.delete(oldest);
  }
  for (let i = pendingSearchQueue.length - 1; i >= 0; i--) {
    if (pendingSearchQueue[i].requestId === id) {
      pendingSearchQueue.splice(i, 1);
    }
  }
}

function executeSearchNow(requestId, options, scope) {
  if (requestId !== undefined && requestId !== null && cancelledRequestIds.has(requestId)) {
    return;
  }

  activeRequestId = requestId ?? null;

  try {
    const isCancelled = () =>
      requestId !== undefined && requestId !== null && cancelledRequestIds.has(requestId);

    const result = findCombinations(
      options,
      ({ exploredNodes, foundCount, currentSize }) => {
        if (isCancelled()) return;
        if (scope && typeof scope.postMessage === 'function') {
          scope.postMessage({
            type: 'PROGRESS',
            requestId,
            exploredNodes,
            foundCount,
            currentSize,
          });
        }
      },
      isCancelled
    );

    if (isCancelled()) return;

    if (scope && typeof scope.postMessage === 'function') {
      scope.postMessage({
        type: 'RESULT',
        requestId,
        result,
      });
    }
  } catch (err) {
    if (scope && typeof scope.postMessage === 'function') {
      scope.postMessage({
        type: 'ERROR',
        requestId,
        message: err instanceof Error ? err.message : String(err || 'Lỗi xử lý tìm kiếm tổ hợp.'),
      });
    }
  } finally {
    if (activeRequestId === requestId) {
      activeRequestId = null;
    }
  }
}

/**
 * Xử lý trực tiếp thông điệp Worker (hỗ trợ cả gọi đồng bộ trong unit test giả lập `self`).
 *
 * @param {any} event
 * @param {object | null} customScope
 */
export function handleWorkerMessage(event, customScope = null) {
  const scope = customScope || getWorkerScope();
  const payload = event && typeof event === 'object' && 'data' in event ? event.data : event;
  if (!payload || typeof payload !== 'object') return;

  const { type, requestId, options } = payload;

  if (type === 'CANCEL') {
    if (requestId !== undefined && requestId !== null) {
      markCancelled(requestId);
    } else if (activeRequestId !== null) {
      markCancelled(activeRequestId);
    }
    return;
  }

  if (type === 'SEARCH') {
    executeSearchNow(requestId, options, scope);
  }
}

function scheduleNextQueuedSearch() {
  if (drainTimer !== null || isExecutingSearch || pendingSearchQueue.length === 0) {
    return;
  }
  drainTimer = setTimeout(drainOneQueuedSearch, 0);
}

function drainOneQueuedSearch() {
  drainTimer = null;

  while (pendingSearchQueue.length > 0) {
    const task = pendingSearchQueue.shift();
    const { requestId, options, scope } = task;
    if (requestId !== undefined && requestId !== null && cancelledRequestIds.has(requestId)) {
      continue;
    }

    isExecutingSearch = true;
    try {
      executeSearchNow(requestId, options, scope);
    } finally {
      isExecutingSearch = false;
    }
    // Nhường lại vòng lặp sự kiện sau mỗi 1 yêu cầu để nhận hết các lệnh CANCEL mới trong hàng đợi
    break;
  }

  if (pendingSearchQueue.length > 0) {
    scheduleNextQueuedSearch();
  }
}

function handleQueuedWorkerMessage(event) {
  const scope = getWorkerScope();
  const payload = event && typeof event === 'object' && 'data' in event ? event.data : event;
  if (!payload || typeof payload !== 'object') return;

  const { type, requestId, options } = payload;

  if (type === 'CANCEL') {
    if (requestId !== undefined && requestId !== null) {
      markCancelled(requestId);
    } else if (activeRequestId !== null) {
      markCancelled(activeRequestId);
    } else if (pendingSearchQueue.length > 0) {
      const lastQueued = pendingSearchQueue.pop();
      if (lastQueued && lastQueued.requestId !== undefined && lastQueued.requestId !== null) {
        markCancelled(lastQueued.requestId);
      }
    }
    return;
  }

  if (type === 'SEARCH') {
    if (requestId !== undefined && requestId !== null && cancelledRequestIds.has(requestId)) {
      return;
    }
    pendingSearchQueue.push({ requestId, options, scope });
    scheduleNextQueuedSearch();
  }
}

// Chỉ tự động gắn listener khi thực sự đang chạy trong môi trường Dedicated Worker (không ghi đè globalThis trên Node main thread)
const isWorkerRuntime =
  typeof self !== 'undefined' &&
  self !== null &&
  typeof window === 'undefined' &&
  typeof self.postMessage === 'function';

if (isWorkerRuntime) {
  self.onmessage = handleQueuedWorkerMessage;
}
