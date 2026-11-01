import { findCombinations } from './reconcileEngine.js';

const getWorkerScope = () => (typeof self !== 'undefined' ? self : globalThis);
const cancelledRequestIds = new Set();
let activeRequestId = null;

function markCancelled(id) {
  if (id === undefined || id === null) return;
  cancelledRequestIds.add(id);
  if (cancelledRequestIds.size > 256) {
    const oldest = cancelledRequestIds.values().next().value;
    cancelledRequestIds.delete(oldest);
  }
}

export function handleWorkerMessage(event) {
  const scope = getWorkerScope();
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
          scope.postMessage({
            type: 'PROGRESS',
            requestId,
            exploredNodes,
            foundCount,
            currentSize,
          });
        },
        isCancelled
      );

      if (isCancelled()) return;

      scope.postMessage({
        type: 'RESULT',
        requestId,
        result,
      });
    } catch (err) {
      scope.postMessage({
        type: 'ERROR',
        requestId,
        message: err instanceof Error ? err.message : String(err || 'Lỗi xử lý tìm kiếm tổ hợp.'),
      });
    } finally {
      if (activeRequestId === requestId) {
        activeRequestId = null;
      }
    }
  }
}

const initialScope = getWorkerScope();
if (initialScope) {
  initialScope.onmessage = handleWorkerMessage;
}
