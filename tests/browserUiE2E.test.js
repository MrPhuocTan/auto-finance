import { describe, test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import http from 'node:http';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import * as XLSX from '../vendor/xlsx.mjs';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const PROJECT_ROOT = path.resolve(__dirname, '..');
const CHROME_EXECUTABLE = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome';
const SUBPATH_PREFIX = '/auto-finance';

const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.ico': 'image/x-icon',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.xls': 'application/vnd.ms-excel',
  '.csv': 'text/csv; charset=utf-8',
};

/**
 * Khởi tạo HTTP Server cục bộ phục vụ ứng dụng tại đúng đường dẫn `/auto-finance/`
 * (mô phỏng môi trường GitHub Pages subpath).
 */
function startLocalSubpathServer() {
  const httpErrors = [];

  const server = http.createServer((req, res) => {
    try {
      const method = req.method || 'GET';
      if (method !== 'GET' && method !== 'HEAD') {
        res.writeHead(405, { 'Content-Type': 'text/plain; charset=utf-8' });
        res.end('Method Not Allowed');
        return;
      }

      const urlObj = new URL(req.url || '/', 'http://127.0.0.1');
      const pathname = decodeURIComponent(urlObj.pathname);

      if (pathname === '/favicon.ico') {
        res.writeHead(204);
        res.end();
        return;
      }

      if (pathname === SUBPATH_PREFIX) {
        res.writeHead(302, { Location: `${SUBPATH_PREFIX}/` });
        res.end();
        return;
      }

      if (!pathname.startsWith(`${SUBPATH_PREFIX}/`)) {
        httpErrors.push({ status: 404, url: pathname });
        res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
        res.end('Not Found');
        return;
      }

      const relPath = pathname.slice(`${SUBPATH_PREFIX}/`.length) || 'index.html';
      const absPath = path.resolve(PROJECT_ROOT, relPath);

      if (absPath !== PROJECT_ROOT && !absPath.startsWith(PROJECT_ROOT + path.sep)) {
        httpErrors.push({ status: 403, url: pathname });
        res.writeHead(403, { 'Content-Type': 'text/plain; charset=utf-8' });
        res.end('Forbidden');
        return;
      }

      let targetFile = absPath;
      if (fs.existsSync(targetFile) && fs.statSync(targetFile).isDirectory()) {
        targetFile = path.join(targetFile, 'index.html');
      }

      if (!fs.existsSync(targetFile) || !fs.statSync(targetFile).isFile()) {
        httpErrors.push({ status: 404, url: pathname });
        res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
        res.end('File Not Found');
        return;
      }

      const ext = path.extname(targetFile).toLowerCase();
      const contentType = MIME_TYPES[ext] || 'application/octet-stream';
      const body = fs.readFileSync(targetFile);

      res.writeHead(200, {
        'Content-Type': contentType,
        'Content-Length': body.byteLength,
        'Cache-Control': 'no-store',
      });

      if (method === 'HEAD') {
        res.end();
      } else {
        res.end(body);
      }
    } catch (err) {
      httpErrors.push({ status: 500, url: req.url, message: err.message });
      res.writeHead(500, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('Internal Server Error');
    }
  });

  return new Promise((resolve, reject) => {
    server.once('error', reject);
    server.listen(0, '127.0.0.1', () => {
      const addr = server.address();
      const port = typeof addr === 'object' && addr ? addr.port : 0;
      resolve({
        server,
        port,
        appUrl: `http://127.0.0.1:${port}${SUBPATH_PREFIX}/`,
        httpErrors,
      });
    });
  });
}

/**
 * Khởi chạy Google Chrome Headless với `--remote-debugging-port=0` và thư mục profile tạm.
 */
function launchHeadlessChrome(userDataDir) {
  return new Promise((resolve, reject) => {
    const args = [
      '--headless=new',
      '--disable-gpu',
      '--no-first-run',
      '--no-default-browser-check',
      '--disable-extensions',
      '--disable-background-networking',
      '--disable-sync',
      '--disable-popup-blocking',
      '--window-size=1440,960',
      '--remote-debugging-port=0',
      `--user-data-dir=${userDataDir}`,
      'about:blank',
    ];

    const proc = spawn(CHROME_EXECUTABLE, args, {
      stdio: ['ignore', 'pipe', 'pipe'],
    });

    let stderrBuf = '';
    let resolved = false;

    const timer = setTimeout(() => {
      if (!resolved) {
        resolved = true;
        proc.kill('SIGKILL');
        reject(new Error(`Quá thời gian chờ Chrome khởi động CDP. stderr:\n${stderrBuf}`));
      }
    }, 15000);

    proc.once('error', (err) => {
      if (!resolved) {
        resolved = true;
        clearTimeout(timer);
        reject(err);
      }
    });

    proc.once('exit', (code, signal) => {
      if (!resolved) {
        resolved = true;
        clearTimeout(timer);
        reject(
          new Error(
            `Chrome thoát sớm (code=${code}, signal=${signal}) trước khi mở CDP. stderr:\n${stderrBuf}`
          )
        );
      }
    });

    proc.stderr.setEncoding('utf8');
    proc.stderr.on('data', (chunk) => {
      stderrBuf += chunk;
      const match = stderrBuf.match(/DevTools listening on (ws:\/\/127\.0\.0\.1:(\d+)\/[^\s]+)/);
      if (match && !resolved) {
        resolved = true;
        clearTimeout(timer);
        resolve({
          proc,
          browserWsUrl: match[1],
          debugPort: Number(match[2]),
        });
      }
    });
  });
}

/**
 * Kết nối CDP trực tiếp tới tab trang qua `WebSocket` gốc của Node v22.
 */
async function createCdpPageClient(debugPort) {
  for (let attempt = 0; attempt < 25; attempt++) {
    const resp = await fetch(`http://127.0.0.1:${debugPort}/json/list`);
    const targets = await resp.json();
    const pageTarget = targets.find((t) => t.type === 'page' && t.webSocketDebuggerUrl);
    if (pageTarget) {
      return connectWebSocketCdp(pageTarget.webSocketDebuggerUrl);
    }
    await new Promise((r) => setTimeout(r, 100));
  }
  throw new Error('Không tìm thấy page target nào từ Chrome DevTools JSON endpoint.');
}

function connectWebSocketCdp(wsUrl) {
  return new Promise((resolve, reject) => {
    const ws = new WebSocket(wsUrl);
    let seq = 0;
    const pending = new Map();
    const listeners = new Map();

    const send = (method, params = {}, timeoutMs = 15000) =>
      new Promise((res, rej) => {
        const id = ++seq;
        const timer = setTimeout(() => {
          pending.delete(id);
          rej(new Error(`CDP timeout khi gọi ${method}`));
        }, timeoutMs);

        pending.set(id, { res, rej, timer, method });
        ws.send(JSON.stringify({ id, method, params }));
      });

    const on = (eventMethod, fn) => {
      if (!listeners.has(eventMethod)) {
        listeners.set(eventMethod, new Set());
      }
      listeners.get(eventMethod).add(fn);
      return () => listeners.get(eventMethod)?.delete(fn);
    };

    ws.addEventListener('message', (event) => {
      const rawText = typeof event.data === 'string' ? event.data : String(event.data);
      const msg = JSON.parse(rawText);

      if (msg.id !== undefined) {
        const entry = pending.get(msg.id);
        if (!entry) return;
        pending.delete(msg.id);
        clearTimeout(entry.timer);
        if (msg.error) {
          entry.rej(new Error(`CDP error (${entry.method}): ${msg.error.message}`));
        } else {
          entry.res(msg.result);
        }
        return;
      }

      if (msg.method && listeners.has(msg.method)) {
        for (const fn of listeners.get(msg.method)) {
          fn(msg.params);
        }
      }
    });

    ws.addEventListener('open', () => {
      resolve({
        ws,
        send,
        on,
        close: () =>
          new Promise((done) => {
            for (const [, entry] of pending) {
              clearTimeout(entry.timer);
            }
            pending.clear();
            if (ws.readyState === WebSocket.CLOSED) {
              done();
              return;
            }
            ws.addEventListener('close', () => done(), { once: true });
            ws.close();
          }),
      });
    });

    ws.addEventListener('error', (err) => {
      reject(err);
    });
  });
}

/**
 * Tạo động 1 file `.xlsx` 2 sheets thực tế trong bộ nhớ bằng `vendor/xlsx.mjs`:
 * - Sheet 1 (`Bia_Ho_So`): Trang bìa & thuyết minh (ít dữ liệu, không có cột Số tiền)
 * - Sheet 2 (`So_Chi_Tiet_131`): Sổ chi tiết công nợ chuẩn Thông tư 200 gồm 2 cột
 *   `Phát sinh Nợ` / `Phát sinh Có`, dòng đánh mã cột phụ `A | B | C | 1 | 2 | D`,
 *   các nghiệp vụ dương/âm/hỗn hợp và 1 dòng `Tổng cộng phát sinh trong kỳ` cuối bảng.
 */
function buildDynamicTwoSheetWorkbookBase64() {
  const wb = XLSX.utils.book_new();

  const sheet1Aoa = [
    ['CÔNG TY TNHH KIỂM TOÁN & KẾ TOÁN ĐẠI VIỆT'],
    ['HỒ SƠ ĐỐI SOÁT CÔNG NỢ PHẢI THU KHÁCH HÀNG - THÁNG 03/2026'],
    ['Ghi chú hồ sơ', 'Số liệu sổ chi tiết nằm ở trang tính So_Chi_Tiet_131'],
    ['Người lập biểu', 'Nguyễn Thị Mai - Kế toán công nợ'],
  ];
  const wsCover = XLSX.utils.aoa_to_sheet(sheet1Aoa);
  XLSX.utils.book_append_sheet(wb, wsCover, 'Bia_Ho_So');

  const sheet2Aoa = [
    ['CÔNG TY TNHH THƯƠNG MẠI ĐẠI VIỆT', '', '', '', '', ''],
    ['SỔ CHI TIẾT TÀI KHOẢN 131 - PHẢI THU CỦA KHÁCH HÀNG', '', '', '', '', ''],
    ['', '', '', '', '', ''],
    [
      'Ngày chứng từ',
      'Số chứng từ',
      'Diễn giải',
      'Phát sinh Nợ',
      'Phát sinh Có',
      'Ghi chú đối chiếu',
    ],
    ['A', 'B', 'C', '1', '2', 'D'],
    // Row 6 (id: row-6): +10.000.000
    ['03/03/2026', 'HD001', 'Xuất bán thiết bị mạng đợt 1', 10000000, 0, 'Ghi chú lô M1'],
    // Row 7 (id: row-7): +6.500.000
    ['05/03/2026', 'HD002', 'Xuất bán bộ lưu điện UPS đợt 2', 6500000, '', 'Ghi chú lô M2'],
    // Row 8 (id: row-8): -1.500.000 (Phát sinh Có -> giảm trừ công nợ)
    ['08/03/2026', 'CK001', 'Chiết khấu thương mại theo hợp đồng quý', '', 1500000, 'Giảm trừ CK quý'],
    // Row 9 (id: row-9): +5.000.000
    ['10/03/2026', 'HD003', 'Cung cấp dây cáp quang nội bộ', 5000000, 0, 'Ghi chú lô M3'],
    // Row 10 (id: row-10): 8.800.000 - 300.000 = +8.500.000
    ['12/03/2026', 'HD004', 'Lắp đặt tủ rack trung tâm (cấn trừ phí)', 8800000, 300000, 'Ghi chú lô M4'],
    // Row 11 (id: row-11): -500.000 (Hàng bán bị trả lại)
    ['15/03/2026', 'TL001', 'Hàng bán bị trả lại do lỗi kỹ thuật', 0, 500000, 'Biên bản trả hàng'],
    // Row 12 (id: row-12): +8.000.000
    ['18/03/2026', 'HD005', 'Bảo trì hệ thống mạng định kỳ', 8000000, 0, 'Ghi chú lô M5'],
    // Row 13 (id: row-13): Dòng Tổng cộng cuối kỳ (Nợ 38.300.000 - Có 2.300.000 = 36.000.000)
    ['', '', 'Tổng cộng phát sinh trong kỳ', 38300000, 2300000, 'Cộng cuối kỳ'],
  ];
  const wsLedger = XLSX.utils.aoa_to_sheet(sheet2Aoa);
  XLSX.utils.book_append_sheet(wb, wsLedger, 'So_Chi_Tiet_131');

  const outBytes = XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' });
  return Buffer.from(outBytes).toString('base64');
}

describe('Auto Finance Browser UI E2E (Chrome Headless CDP)', () => {
  let serverInfo = null;
  let chromeInfo = null;
  let cdp = null;
  let tempProfileDir = null;

  const consoleErrors = [];
  const runtimeExceptions = [];
  const browserLogErrors = [];
  const openedDialogs = [];

  async function evalPage(expression) {
    const res = await cdp.send('Runtime.evaluate', {
      expression,
      awaitPromise: true,
      returnByValue: true,
    });
    if (res.exceptionDetails) {
      const detail =
        res.exceptionDetails.exception?.description ||
        res.exceptionDetails.text ||
        'Unknown Runtime.evaluate exception';
      throw new Error(`Lỗi thực thi JS trên trang: ${detail}`);
    }
    return res.result ? res.result.value : undefined;
  }

  async function waitForPageCondition(jsBooleanExpr, timeoutMs = 8000) {
    const deadline = Date.now() + timeoutMs;
    while (Date.now() < deadline) {
      const ok = await evalPage(`Boolean(${jsBooleanExpr})`);
      if (ok) return;
      await new Promise((r) => setTimeout(r, 40));
    }
    throw new Error(`Hết thời gian chờ điều kiện UI (${timeoutMs}ms): ${jsBooleanExpr}`);
  }

  async function waitForSearchIdle(timeoutMs = 10000) {
    await waitForPageCondition(
      `!document.getElementById('btnRunSearch').disabled && document.getElementById('searchProgressStrip').classList.contains('hidden')`,
      timeoutMs
    );
  }

  async function navigateFreshApp() {
    await cdp.send('Page.navigate', { url: serverInfo.appUrl });
    await waitForPageCondition(
      `document.readyState === 'complete' && document.getElementById('targetAmountInput') && document.getElementById('toleranceInput').value === '500'`
    );
    await evalPage(`window.__xssTriggered = false;`);
  }

  async function uploadDynamicTwoSheetWorkbookAndSelectSheet2() {
    const xlsxBase64 = buildDynamicTwoSheetWorkbookBase64();
    await evalPage(`(() => {
      const b64 = ${JSON.stringify(xlsxBase64)};
      const binStr = atob(b64);
      const bytes = new Uint8Array(binStr.length);
      for (let i = 0; i < binStr.length; i++) {
        bytes[i] = binStr.charCodeAt(i);
      }
      const file = new File([bytes], 'So_Cong_No_2_Sheets.xlsx', {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      const dt = new DataTransfer();
      dt.items.add(file);
      const fileInput = document.getElementById('excelFileInput');
      fileInput.files = dt.files;
      fileInput.dispatchEvent(new Event('change', { bubbles: true }));
    })()`);

    await waitForPageCondition(
      `document.getElementById('fileNameDisplay').textContent === 'So_Cong_No_2_Sheets.xlsx'`
    );
  }

  before(async () => {
    assert.ok(
      fs.existsSync(CHROME_EXECUTABLE),
      `Không tìm thấy Google Chrome tại ${CHROME_EXECUTABLE}`
    );

    serverInfo = await startLocalSubpathServer();
    tempProfileDir = fs.mkdtempSync(path.join(os.tmpdir(), 'auto-finance-chrome-e2e-'));
    chromeInfo = await launchHeadlessChrome(tempProfileDir);
    cdp = await createCdpPageClient(chromeInfo.debugPort);

    cdp.on('Runtime.consoleAPICalled', (params) => {
      if (params.type === 'error') {
        const text = (params.args || [])
          .map((a) => a.description || a.value || a.unserializableValue || JSON.stringify(a))
          .join(' ');
        consoleErrors.push(text);
      }
    });

    cdp.on('Runtime.exceptionThrown', (params) => {
      const desc =
        params.exceptionDetails?.exception?.description ||
        params.exceptionDetails?.text ||
        JSON.stringify(params.exceptionDetails);
      runtimeExceptions.push(desc);
    });

    cdp.on('Log.entryAdded', (params) => {
      if (params.entry?.level === 'error') {
        browserLogErrors.push(
          `${params.entry.source}: ${params.entry.text} (${params.entry.url || ''})`
        );
      }
    });

    cdp.on('Page.javascriptDialogOpening', (params) => {
      openedDialogs.push({ type: params.type, message: params.message });
      cdp.send('Page.handleJavaScriptDialog', { accept: true }).catch(() => {});
    });

    await cdp.send('Runtime.enable');
    await cdp.send('Log.enable');
    await cdp.send('Page.enable');

    await navigateFreshApp();
  });

  after(async () => {
    if (cdp) {
      await cdp.close().catch(() => {});
    }

    if (chromeInfo?.proc) {
      await new Promise((resolve) => {
        const proc = chromeInfo.proc;
        if (proc.exitCode !== null || proc.killed) {
          resolve();
          return;
        }
        const killTimer = setTimeout(() => {
          try {
            proc.kill('SIGKILL');
          } catch {}
          resolve();
        }, 3000);
        proc.once('exit', () => {
          clearTimeout(killTimer);
          resolve();
        });
        try {
          proc.kill('SIGTERM');
        } catch {
          clearTimeout(killTimer);
          resolve();
        }
      });
    }

    if (serverInfo?.server) {
      await new Promise((resolve) => serverInfo.server.close(() => resolve()));
    }

    if (tempProfileDir && fs.existsSync(tempProfileDir)) {
      fs.rmSync(tempProfileDir, { recursive: true, force: true });
    }
  });

  // ==========================================================================
  // CA 1: Ô NHẬP SỐ TIỀN (#targetAmountInput), Ô SAI SỐ (#toleranceInput) & VALIDATION
  // ==========================================================================
  test('1.1 #targetAmountInput & #toleranceInput xử lý chính xác các chuỗi dán sao kê, số âm, ngoặc đơn, số tỷ, 0 và xóa trắng', async () => {
    await navigateFreshApp();

    const targetCases = [
      {
        raw: '+ 18,450,000 VND',
        expectedValue: '18.450.000',
        expectedWords: 'Bằng chữ: Mười tám triệu bốn trăm năm mươi nghìn đồng',
        expectedNegClass: false,
      },
      {
        raw: '12.500.000 đ',
        expectedValue: '12.500.000',
        expectedWords: 'Bằng chữ: Mười hai triệu năm trăm nghìn đồng',
        expectedNegClass: false,
      },
      {
        raw: '-2.450.000',
        expectedValue: '-2.450.000',
        expectedWords: 'Bằng chữ: Âm hai triệu bốn trăm năm mươi nghìn đồng',
        expectedNegClass: true,
      },
      {
        raw: '(1.500.000)',
        expectedValue: '-1.500.000',
        expectedWords: 'Bằng chữ: Âm một triệu năm trăm nghìn đồng',
        expectedNegClass: true,
      },
      {
        raw: '125.600.000.000',
        expectedValue: '125.600.000.000',
        expectedWords: 'Bằng chữ: Một trăm hai mươi lăm tỷ sáu trăm triệu đồng',
        expectedNegClass: false,
      },
      {
        raw: '0',
        expectedValue: '0',
        expectedWords: 'Bằng chữ: Không đồng',
        expectedNegClass: false,
      },
      {
        raw: '',
        expectedValue: '',
        expectedWords:
          'Nhập số tiền khách thanh toán trên sao kê ngân hàng để kiểm tra bằng chữ',
        expectedNegClass: false,
      },
    ];

    for (const tc of targetCases) {
      const actual = await evalPage(`(() => {
        const el = document.getElementById('targetAmountInput');
        el.focus();
        el.value = ${JSON.stringify(tc.raw)};
        el.dispatchEvent(new Event('input', { bubbles: true }));
        return {
          value: el.value,
          words: document.getElementById('amountWordsDisplay').textContent.trim(),
          hasNegClass: el.classList.contains('is-negative'),
        };
      })()`);

      assert.equal(
        actual.value,
        tc.expectedValue,
        `Sai giá trị format cho input "${tc.raw}"`
      );
      assert.equal(
        actual.words,
        tc.expectedWords,
        `Sai dòng đọc bằng chữ tiếng Việt cho input "${tc.raw}"`
      );
      assert.equal(
        actual.hasNegClass,
        tc.expectedNegClass,
        `Sai class .is-negative cho input "${tc.raw}"`
      );
    }

    // Kiểm tra nút Xóa số (#btnClearAmount)
    const afterClearBtn = await evalPage(`(() => {
      const el = document.getElementById('targetAmountInput');
      el.value = '9.876.543';
      el.dispatchEvent(new Event('input', { bubbles: true }));
      document.getElementById('btnClearAmount').click();
      return {
        value: el.value,
        words: document.getElementById('amountWordsDisplay').textContent.trim(),
      };
    })()`);
    assert.equal(afterClearBtn.value, '');
    assert.match(afterClearBtn.words, /Nhập số tiền khách thanh toán/);

    // Kiểm tra ô Sai số (#toleranceInput) với chuỗi dán phức tạp & quick chips
    const tolCheck = await evalPage(`(() => {
      const tol = document.getElementById('toleranceInput');

      tol.value = '+ 1,000 VND';
      tol.dispatchEvent(new Event('input', { bubbles: true }));
      const afterPaste1000 = {
        value: tol.value,
        chip1000Active: document.querySelector('[data-tolerance="1000"]').classList.contains('active'),
      };

      tol.value = '(5.000)';
      tol.dispatchEvent(new Event('input', { bubbles: true }));
      const afterPasteNeg5000 = {
        value: tol.value,
        chip5000Active: document.querySelector('[data-tolerance="5000"]').classList.contains('active'),
      };

      document.querySelector('[data-tolerance="0"]').click();
      const afterClickChip0 = {
        value: tol.value,
        chip0Active: document.querySelector('[data-tolerance="0"]').classList.contains('active'),
      };

      document.querySelector('[data-tolerance="500"]').click();
      return { afterPaste1000, afterPasteNeg5000, afterClickChip0, restored: tol.value };
    })()`);

    assert.equal(tolCheck.afterPaste1000.value, '1.000');
    assert.equal(tolCheck.afterPaste1000.chip1000Active, true);
    assert.equal(tolCheck.afterPasteNeg5000.value, '5.000');
    assert.equal(tolCheck.afterPasteNeg5000.chip5000Active, true);
    assert.equal(tolCheck.afterClickChip0.value, '0');
    assert.equal(tolCheck.afterClickChip0.chip0Active, true);
    assert.equal(tolCheck.restored, '500');
  });

  test('1.2 Cảnh báo tiếng Việt thân thiện khi chưa nạp sổ, chưa nhập số tiền, hoặc chọn Từ ngày > Đến ngày', async () => {
    await navigateFreshApp();

    // 1) Bấm tìm khi chưa nạp file Excel
    const noDataToast = await evalPage(`(() => {
      document.getElementById('btnRunSearch').click();
      const toast = document.getElementById('statusToast');
      return {
        hidden: toast.classList.contains('hidden'),
        text: toast.textContent.trim(),
      };
    })()`);
    assert.equal(noDataToast.hidden, false);
    assert.match(noDataToast.text, /Vui lòng chọn file Excel.*Nạp bảng kê mẫu/i);

    // 2) Nạp bảng kê mẫu nhưng để trống #targetAmountInput rồi bấm tìm
    const emptyAmountToast = await evalPage(`(() => {
      document.getElementById('btnLoadSample').click();
      const amountEl = document.getElementById('targetAmountInput');
      amountEl.value = '';
      amountEl.dispatchEvent(new Event('input', { bubbles: true }));
      document.getElementById('btnRunSearch').click();
      const toast = document.getElementById('statusToast');
      return {
        hidden: toast.classList.contains('hidden'),
        text: toast.textContent.trim(),
        isFocused: document.activeElement === amountEl,
      };
    })()`);
    assert.equal(emptyAmountToast.hidden, false);
    assert.match(emptyAmountToast.text, /Vui lòng nhập Số tiền cần gom/i);
    assert.equal(emptyAmountToast.isFocused, true);

    // 3) Chọn dateFromInput > dateToInput rồi bấm tìm -> tự đảo lại ngày, hiện thông báo tiếng Việt, không lỗi JS
    const invertedSearchSetup = await evalPage(`(() => {
      const fromEl = document.getElementById('dateFromInput');
      const toEl = document.getElementById('dateToInput');
      fromEl.value = '2026-03-15';
      fromEl.dispatchEvent(new Event('change', { bubbles: true }));
      toEl.value = '2026-03-01';
      toEl.dispatchEvent(new Event('change', { bubbles: true }));

      const amountEl = document.getElementById('targetAmountInput');
      amountEl.value = '18.450.000';
      amountEl.dispatchEvent(new Event('input', { bubbles: true }));

      document.getElementById('btnRunSearch').click();
      const toast = document.getElementById('statusToast');
      return {
        swappedFrom: fromEl.value,
        swappedTo: toEl.value,
        fromVN: document.getElementById('dateFromVN').textContent.trim(),
        toVN: document.getElementById('dateToVN').textContent.trim(),
        toastText: toast.textContent.trim(),
      };
    })()`);

    assert.equal(invertedSearchSetup.swappedFrom, '2026-03-01');
    assert.equal(invertedSearchSetup.swappedTo, '2026-03-15');
    assert.equal(invertedSearchSetup.fromVN, 'Định dạng VN: 01/03/2026');
    assert.equal(invertedSearchSetup.toVN, 'Định dạng VN: 15/03/2026');
    assert.match(invertedSearchSetup.toastText, /tự động đảo lại thứ tự Từ ngày trước Đến ngày/i);

    await waitForSearchIdle();
    const bannerAfterSwap = await evalPage(
      `document.getElementById('resultsSummaryBanner').textContent.replace(/\\s+/g, ' ').trim()`
    );
    assert.match(bannerAfterSwap, /Tìm thấy 2 phương án trong khoảng 01\/03\/2026 – 15\/03\/2026/);
  });

  // ==========================================================================
  // CA 2: UPLOAD FILE .XLSX THẬT 2 SHEETS & ĐỔI CẤU HÌNH CỘT (#mappingPanel)
  // ==========================================================================
  test('2.1 Upload file .xlsx 2 sheets qua DataTransfer, chuyển sang Sheet 2 tự nhận diện debit_minus_credit và tùy biến #mappingPanel', async () => {
    await navigateFreshApp();
    await uploadDynamicTwoSheetWorkbookAndSelectSheet2();

    // Kiểm tra trạng thái ngay khi mở Sheet 1 (Bia_Ho_So - trang bìa thiếu cột Số tiền)
    const sheet1State = await evalPage(`(() => {
      const sheetSel = document.getElementById('sheetSelect');
      return {
        sheets: Array.from(sheetSel.options).map((o) => o.value),
        activeSheet: sheetSel.value,
        disabled: sheetSel.disabled,
        mappingOpenOnCover: !document.getElementById('mappingPanel').classList.contains('hidden'),
        warningVisible: !document.getElementById('mappingWarning').classList.contains('hidden'),
      };
    })()`);

    assert.deepEqual(sheet1State.sheets, ['Bia_Ho_So', 'So_Chi_Tiet_131']);
    assert.equal(sheet1State.activeSheet, 'Bia_Ho_So');
    assert.equal(sheet1State.disabled, false);
    assert.equal(
      sheet1State.mappingOpenOnCover,
      true,
      'Sheet bìa không có cột Số tiền nên #mappingPanel phải tự động mở'
    );
    assert.equal(sheet1State.warningVisible, true);

    // Chuyển #sheetSelect sang Sheet 2 (So_Chi_Tiet_131)
    const sheet2State = await evalPage(`(() => {
      const sheetSel = document.getElementById('sheetSelect');
      sheetSel.value = 'So_Chi_Tiet_131';
      sheetSel.dispatchEvent(new Event('change', { bubbles: true }));

      const rows = Array.from(document.querySelectorAll('#ledgerTableBody tr[data-row-id]')).map((tr) => ({
        id: tr.getAttribute('data-row-id'),
        ref: tr.querySelector('.col-ref')?.textContent.trim(),
        desc: tr.querySelector('.col-desc')?.textContent.trim(),
        amountText: tr.querySelector('.col-amount')?.textContent.trim(),
        isNegStyle: Boolean(tr.querySelector('.col-amount .num-neg')),
        status: tr.querySelector('.col-status')?.textContent.trim(),
      }));

      return {
        headerRowValue: document.getElementById('headerRowSelect').value,
        amountMode: document.getElementById('amountModeSelect').value,
        debitCol: document.getElementById('debitColSelect').value,
        creditCol: document.getElementById('creditColSelect').value,
        debitWrapHidden: document.getElementById('debitColWrap').classList.contains('hidden'),
        singleWrapHidden: document.getElementById('singleAmountWrap').classList.contains('hidden'),
        mappingAutoCollapsed: document.getElementById('mappingPanel').classList.contains('hidden'),
        extractSummary: document.getElementById('extractSummaryText').textContent.trim(),
        mappedSummary: document.getElementById('mappedColsSummary').textContent.trim(),
        rows,
      };
    })()`);

    // Dòng tiêu đề là dòng 4 trong Excel (0-based index = 3), bỏ qua dòng phụ A|B|C|1|2|D ở dòng 5
    assert.equal(sheet2State.headerRowValue, '3');
    assert.equal(sheet2State.amountMode, 'debit_minus_credit');
    assert.equal(sheet2State.debitCol, '3');
    assert.equal(sheet2State.creditCol, '4');
    assert.equal(sheet2State.debitWrapHidden, false);
    assert.equal(sheet2State.singleWrapHidden, true);
    assert.equal(sheet2State.mappingAutoCollapsed, true);
    assert.equal(
      sheet2State.extractSummary,
      'Đã đọc 7 dòng nghiệp vụ (tự động bỏ qua 1 dòng tổng cộng)'
    );
    assert.match(sheet2State.mappedSummary, /Nợ \[D\] Phát sinh Nợ - Có \[E\] Phát sinh Có/);

    // Kiểm tra chính xác 7 dòng nghiệp vụ dương & âm (Nợ trừ Có)
    assert.equal(sheet2State.rows.length, 7);
    const byRef = Object.fromEntries(sheet2State.rows.map((r) => [r.ref, r]));
    assert.equal(byRef.HD001.amountText, '10.000.000');
    assert.equal(byRef.HD001.isNegStyle, false);
    assert.equal(byRef.HD002.amountText, '6.500.000');
    assert.equal(byRef.CK001.amountText, '-1.500.000');
    assert.equal(byRef.CK001.isNegStyle, true);
    assert.equal(byRef.HD004.amountText, '8.500.000'); // 8.800.000 - 300.000
    assert.equal(byRef.TL001.amountText, '-500.000');
    assert.equal(byRef.TL001.isNegStyle, true);

    // Mở #mappingPanel, đổi cột Diễn giải sang cột Ghi chú (index 5), bật/tắt #includeSummaryCheckbox và bấm #btnApplyMapping
    const customMappingCheck = await evalPage(`(() => {
      document.getElementById('tabLedgerAll').click();
      const summaryStatusBefore = document
        .querySelector('#ledgerTableBody tr[data-row-id="row-13"] .col-status')
        ?.textContent.trim();

      document.getElementById('btnToggleMapping').click();
      const panelOpen = !document.getElementById('mappingPanel').classList.contains('hidden');

      const descSel = document.getElementById('descColSelect');
      descSel.value = '5';
      descSel.dispatchEvent(new Event('change', { bubbles: true }));

      const sumChk = document.getElementById('includeSummaryCheckbox');
      sumChk.checked = true;
      sumChk.dispatchEvent(new Event('change', { bubbles: true }));

      const firstDesc = document
        .querySelector('#ledgerTableBody tr[data-row-id="row-6"] .col-desc')
        ?.textContent.trim();
      const summaryStatusWhenIncluded = document
        .querySelector('#ledgerTableBody tr[data-row-id="row-13"] .col-status')
        ?.textContent.trim();

      // Khôi phục lại cột Diễn giải gốc (index 2), tắt #includeSummaryCheckbox, chuyển lại tab Filtered và bấm #btnApplyMapping
      descSel.value = '2';
      descSel.dispatchEvent(new Event('change', { bubbles: true }));
      sumChk.checked = false;
      sumChk.dispatchEvent(new Event('change', { bubbles: true }));
      document.getElementById('tabLedgerFiltered').click();
      document.getElementById('btnApplyMapping').click();

      return {
        panelOpen,
        summaryStatusBefore,
        firstDesc,
        summaryStatusWhenIncluded,
        panelClosedAfterApply: document.getElementById('mappingPanel').classList.contains('hidden'),
        toastText: document.getElementById('statusToast').textContent.trim(),
      };
    })()`);

    assert.equal(customMappingCheck.panelOpen, true);
    assert.equal(customMappingCheck.summaryStatusBefore, 'Dòng tổng cộng (Bỏ qua)');
    assert.equal(customMappingCheck.firstDesc, 'Ghi chú lô M1');
    assert.equal(customMappingCheck.summaryStatusWhenIncluded, 'Chưa đối soát');
    assert.equal(customMappingCheck.panelClosedAfterApply, true);
    assert.match(customMappingCheck.toastText, /Đã áp dụng thiết lập cột dữ liệu/);
  });

  // ==========================================================================
  // CA 3: TƯƠNG TÁC BẢNG SỔ CHI TIẾT, GHIM, LOẠI TRỪ & STATUS BAR KIỂU EXCEL
  // ==========================================================================
  test('3.1 Tích chọn thủ công 3 dòng (gồm dương và âm) -> #selectionStatusText tính đúng tổng và độ lệch so với #targetAmountInput', async () => {
    // Trên Sheet 2 đang mở: chọn row-6 (+10.000.000), row-7 (+6.500.000), row-8 (-1.500.000) => Tổng = 15.000.000
    const statusResult = await evalPage(`(() => {
      const targetEl = document.getElementById('targetAmountInput');
      targetEl.value = '14.800.000';
      targetEl.dispatchEvent(new Event('input', { bubbles: true }));

      const idsToCheck = ['row-6', 'row-7', 'row-8'];
      for (const id of idsToCheck) {
        const chk = document.querySelector(\`#ledgerTableBody input[data-row-action="check"][data-item-id="\${id}"]\`);
        chk.checked = true;
        chk.dispatchEvent(new Event('change', { bubbles: true }));
      }
      const whenTarget14_8M = document.getElementById('selectionStatusText').textContent.replace(/\\s+/g, ' ').trim();

      targetEl.value = '15.500.000';
      targetEl.dispatchEvent(new Event('input', { bubbles: true }));
      const whenTarget15_5M = document.getElementById('selectionStatusText').textContent.replace(/\\s+/g, ' ').trim();

      targetEl.value = '15.000.000';
      targetEl.dispatchEvent(new Event('input', { bubbles: true }));
      const whenTargetExact15M = document.getElementById('selectionStatusText').textContent.replace(/\\s+/g, ' ').trim();

      document.getElementById('btnClearChecked').click();
      const afterClearChecked = document.getElementById('selectionStatusText').textContent.replace(/\\s+/g, ' ').trim();

      return { whenTarget14_8M, whenTarget15_5M, whenTargetExact15M, afterClearChecked };
    })()`);

    assert.equal(
      statusResult.whenTarget14_8M,
      'Đang ghim/chọn: 3 dòng | Tổng tiền: 15.000.000 | Chênh lệch so với số cần tìm: +200.000'
    );
    assert.equal(
      statusResult.whenTarget15_5M,
      'Đang ghim/chọn: 3 dòng | Tổng tiền: 15.000.000 | Chênh lệch so với số cần tìm: -500.000'
    );
    assert.equal(
      statusResult.whenTargetExact15M,
      'Đang ghim/chọn: 3 dòng | Tổng tiền: 15.000.000 | Chênh lệch so với số cần tìm: 0 (Khớp chính xác)'
    );
    assert.equal(
      statusResult.afterClearChecked,
      'Đang ghim/chọn: 0 dòng | Tổng tiền: 0 | Chênh lệch so với số cần tìm: -15.000.000'
    );
  });

  test('3.2 Ghim bắt buộc 1 dòng, Loại trừ 1 dòng khi tìm kiếm & chuyển đổi trực tiếp Ghim <-> Loại trừ không xung đột', async () => {
    // Với mục tiêu 15.000.000 (±0) trên Sheet 2:
    // Nếu không ghim/loại trừ thì có 4 nghiệm: (row-6+row-9), (row-7+row-10), (row-6+row-7+row-8), (row-8+row-10+row-12).
    // Khi Ghim row-8 (-1.5M, CK001) và Loại trừ row-12 (+8.0M, HD005), chỉ còn đúng nghiệm chứa row-8 và không chứa row-12: (row-6 + row-7 + row-8).
    await evalPage(`(() => {
      document.querySelector('[data-tolerance="0"]').click();
      const targetEl = document.getElementById('targetAmountInput');
      targetEl.value = '15.000.000';
      targetEl.dispatchEvent(new Event('input', { bubbles: true }));

      document.querySelector('#ledgerTableBody button[data-row-action="pin"][data-item-id="row-8"]').click();
      document.querySelector('#ledgerTableBody button[data-row-action="exclude"][data-item-id="row-12"]').click();
      document.getElementById('btnRunSearch').click();
    })()`);

    await waitForSearchIdle();

    const searchWithPinAndExclude = await evalPage(`(() => {
      const cards = Array.from(document.querySelectorAll('[data-solution-card]'));
      for (const card of cards) {
        const headerBtn = card.querySelector('button[data-sol-action="toggle"]');
        if (headerBtn && headerBtn.getAttribute('aria-expanded') === 'false') {
          headerBtn.click();
        }
      }

      const expandedCards = Array.from(document.querySelectorAll('[data-solution-card]'));
      return expandedCards.map((card) => {
        const excelRows = Array.from(card.querySelectorAll('tbody tr .col-excel-row')).map((td) =>
          td.textContent.trim()
        );
        const refs = Array.from(card.querySelectorAll('tbody tr .col-ref')).map((td) =>
          td.textContent.trim()
        );
        const pinnedBadges = card.querySelectorAll('tbody tr .status-pinned').length;
        return { excelRows, refs, pinnedBadges };
      });
    })()`);

    assert.ok(searchWithPinAndExclude.length >= 1, 'Phải tìm thấy ít nhất 1 phương án hợp lệ');
    for (const sol of searchWithPinAndExclude) {
      assert.ok(
        sol.excelRows.includes('#8') && sol.refs.includes('CK001'),
        `Mọi phương án đều phải chứa dòng đã ghim #8 (CK001), thực tế: ${sol.excelRows.join(', ')}`
      );
      assert.ok(
        !sol.excelRows.includes('#12') && !sol.refs.includes('HD005'),
        `Không phương án nào được phép chứa dòng bị loại trừ #12 (HD005), thực tế: ${sol.excelRows.join(', ')}`
      );
      assert.equal(sol.pinnedBadges, 1, 'Dòng đã ghim trong phương án phải gắn nhãn "Đã ghim"');
    }

    // Thử bấm trực tiếp "Ghim bắt buộc" trên dòng #12 đang bị "Loại trừ" và ngược lại "Loại trừ" trên dòng đang "Ghim"
    const toggleMutualExclusion = await evalPage(`(() => {
      document.querySelector('#ledgerTableBody button[data-row-action="pin"][data-item-id="row-12"]').click();
      const row12AfterPin = document.querySelector('#ledgerTableBody tr[data-row-id="row-12"]');
      const step1 = {
        isRowPinned: row12AfterPin.classList.contains('row-pinned'),
        isRowExcluded: row12AfterPin.classList.contains('row-excluded'),
        statusText: row12AfterPin.querySelector('.col-status').textContent.trim(),
        pinBtnText: row12AfterPin.querySelector('[data-row-action="pin"]').textContent.trim(),
        excludeBtnText: row12AfterPin.querySelector('[data-row-action="exclude"]').textContent.trim(),
        pinBtnActive: row12AfterPin.querySelector('[data-row-action="pin"]').classList.contains('active-pin'),
        exclBtnActive: row12AfterPin.querySelector('[data-row-action="exclude"]').classList.contains('active-exclude'),
        clearExcludesHidden: document.getElementById('btnClearExcludes').classList.contains('hidden'),
        clearPinsText: document.getElementById('btnClearPins').textContent.trim(),
      };

      document.querySelector('#ledgerTableBody button[data-row-action="exclude"][data-item-id="row-12"]').click();
      const row12AfterExcl = document.querySelector('#ledgerTableBody tr[data-row-id="row-12"]');
      const step2 = {
        isRowPinned: row12AfterExcl.classList.contains('row-pinned'),
        isRowExcluded: row12AfterExcl.classList.contains('row-excluded'),
        statusText: row12AfterExcl.querySelector('.col-status').textContent.trim(),
        pinBtnText: row12AfterExcl.querySelector('[data-row-action="pin"]').textContent.trim(),
        excludeBtnText: row12AfterExcl.querySelector('[data-row-action="exclude"]').textContent.trim(),
        pinBtnActive: row12AfterExcl.querySelector('[data-row-action="pin"]').classList.contains('active-pin'),
        exclBtnActive: row12AfterExcl.querySelector('[data-row-action="exclude"]').classList.contains('active-exclude'),
      };

      document.getElementById('btnClearPins').click();
      document.getElementById('btnClearExcludes').click();
      return { step1, step2 };
    })()`);

    assert.equal(toggleMutualExclusion.step1.isRowPinned, true);
    assert.equal(toggleMutualExclusion.step1.isRowExcluded, false);
    assert.equal(toggleMutualExclusion.step1.statusText, 'Ghim bắt buộc');
    assert.equal(toggleMutualExclusion.step1.pinBtnText, 'Đang ghim');
    assert.equal(toggleMutualExclusion.step1.excludeBtnText, 'Loại trừ');
    assert.equal(toggleMutualExclusion.step1.pinBtnActive, true);
    assert.equal(toggleMutualExclusion.step1.exclBtnActive, false);
    assert.equal(toggleMutualExclusion.step1.clearExcludesHidden, true);
    assert.equal(toggleMutualExclusion.step1.clearPinsText, 'Bỏ ghim bắt buộc (2)');

    assert.equal(toggleMutualExclusion.step2.isRowPinned, false);
    assert.equal(toggleMutualExclusion.step2.isRowExcluded, true);
    assert.equal(toggleMutualExclusion.step2.statusText, 'Đã loại trừ');
    assert.equal(toggleMutualExclusion.step2.pinBtnText, 'Ghim bắt buộc');
    assert.equal(toggleMutualExclusion.step2.excludeBtnText, 'Bỏ loại');
    assert.equal(toggleMutualExclusion.step2.pinBtnActive, false);
    assert.equal(toggleMutualExclusion.step2.exclBtnActive, true);
  });

  // ==========================================================================
  // CA 4: LUỒNG ĐỐI SOÁT NHIỀU ĐỢT, GHI CHÚ UNC (CHỐNG XSS), HOÀN TÁC & CỠ CHỮ LỚN
  // ==========================================================================
  test('4.1 Nạp dữ liệu mẫu -> Preset 1 -> Ghi chú chống XSS -> Xác nhận Đợt #1 -> Tìm lại không trùng -> Hoàn tác -> Cỡ chữ Lớn lưu localStorage', async () => {
    await navigateFreshApp();

    // 1) Nạp bảng kê mẫu & chạy Preset 1 (18.450.000đ, 01/03/2026 - 15/03/2026, ±500)
    await evalPage(`(() => {
      document.getElementById('btnLoadSample').click();
      document.querySelector('button[data-preset-idx="0"]').click();
    })()`);

    await waitForSearchIdle();

    const preset1Initial = await evalPage(`(() => {
      const cards = Array.from(document.querySelectorAll('[data-solution-card]'));
      const firstId = cards[0]?.getAttribute('data-solution-card');
      const firstRowRefs = Array.from(cards[0]?.querySelectorAll('tbody tr .col-ref') || []).map(
        (el) => el.textContent.trim()
      );
      return {
        solutionCount: cards.length,
        firstId,
        firstRowRefs,
      };
    })()`);

    assert.equal(preset1Initial.solutionCount, 2);
    assert.deepEqual(preset1Initial.firstRowRefs, ['HD0301', 'HD0303', 'HD0304', 'PT0301', 'CK0301']);

    // 2) Nhập ghi chú có ký tự đặc biệt HTML / XSS payload vào ô ghi chú phương án #1 rồi bấm Xác nhận Gom đợt này
    const xssNotePayload = '<script>alert(1)</script> UNC #01';
    const afterConfirmBatch1 = await evalPage(`(() => {
      const noteInput = document.querySelector('[data-solution-card] input.solution-note-input');
      noteInput.value = ${JSON.stringify(xssNotePayload)};
      noteInput.dispatchEvent(new Event('input', { bubbles: true }));

      document.querySelector('[data-solution-card] button[data-sol-action="confirm"]').click();

      // Chuyển sổ gốc sang tab Đã đối soát để kiểm tra trạng thái dòng
      document.getElementById('tabLedgerReconciled').click();
      const reconciledRows = Array.from(document.querySelectorAll('#ledgerTableBody tr[data-row-id]')).map((tr) => ({
        ref: tr.querySelector('.col-ref')?.textContent.trim(),
        status: tr.querySelector('.col-status')?.textContent.trim(),
        pinDisabled: tr.querySelector('[data-row-action="pin"]')?.disabled,
        exclDisabled: tr.querySelector('[data-row-action="exclude"]')?.disabled,
      }));

      return {
        xssTriggered: Boolean(window.__xssTriggered),
        injectedScriptsInBatches: document.querySelectorAll('#batchesPanel script').length,
        injectedScriptsInResults: document.querySelectorAll('#solutionsContainer script').length,
        batchesHidden: document.getElementById('batchesPanel').classList.contains('hidden'),
        batchesTitle: document.getElementById('batchesPanelTitle').textContent.trim(),
        firstBatchLabel: document.querySelector('#batchesTableBody .status-reconciled')?.textContent.trim(),
        firstBatchNoteVal: document.querySelector('#batchesTableBody input.batch-note-edit')?.value,
        reconciledTabLabel: document.getElementById('tabLedgerReconciled').textContent.replace(/\\s+/g, ' ').trim(),
        reconciledRows,
      };
    })()`);

    assert.equal(afterConfirmBatch1.xssTriggered, false, 'Tuyệt đối không được thực thi mã XSS');
    assert.equal(openedDialogs.length, 0, 'Không được bật bất kỳ hộp thoại alert() nào từ XSS');
    assert.equal(afterConfirmBatch1.injectedScriptsInBatches, 0);
    assert.equal(afterConfirmBatch1.injectedScriptsInResults, 0);
    assert.equal(afterConfirmBatch1.batchesHidden, false);
    assert.equal(afterConfirmBatch1.batchesTitle, 'DANH SÁCH CÁC ĐỢT ĐÃ ĐỐI SOÁT (1 ĐỢT)');
    assert.equal(afterConfirmBatch1.firstBatchLabel, 'Đợt #1');
    assert.equal(afterConfirmBatch1.firstBatchNoteVal, xssNotePayload);
    assert.equal(afterConfirmBatch1.reconciledTabLabel, 'Đã đối soát (5)');
    assert.equal(afterConfirmBatch1.reconciledRows.length, 5);
    for (const r of afterConfirmBatch1.reconciledRows) {
      assert.equal(r.status, 'Đã gom (Đợt #1)');
      assert.equal(r.pinDisabled, true);
      assert.equal(r.exclDisabled, true);
    }

    // 3) Bấm tìm lại cùng số tiền 18.450.000 -> không còn trùng phương án 5 dòng cũ
    await evalPage(`document.getElementById('btnRunSearch').click()`);
    await waitForSearchIdle();

    const searchAfterBatch1 = await evalPage(`(() => {
      const cards = Array.from(document.querySelectorAll('[data-solution-card]'));
      const refs = Array.from(cards[0]?.querySelectorAll('tbody tr .col-ref') || []).map((el) =>
        el.textContent.trim()
      );
      return {
        solutionCount: cards.length,
        refs,
      };
    })()`);

    assert.equal(
      searchAfterBatch1.solutionCount,
      1,
      'Sau khi khóa Đợt #1, tìm lại cùng số tiền chỉ được còn 1 phương án gần khớp khác'
    );
    assert.deepEqual(searchAfterBatch1.refs, ['HD0302', 'HD0305']);

    // 4) Bấm "Hoàn tác / Hủy đợt" -> các dòng mở khóa lại bình thường, tìm lại thấy đủ 2 phương án
    const afterUndo = await evalPage(`(() => {
      document.querySelector('#batchesTableBody button[data-batch-action="undo"]').click();
      document.getElementById('tabLedgerFiltered').click();

      const hd0301Row = Array.from(document.querySelectorAll('#ledgerTableBody tr[data-row-id]')).find(
        (tr) => tr.querySelector('.col-ref')?.textContent.trim() === 'HD0301'
      );

      return {
        batchesHidden: document.getElementById('batchesPanel').classList.contains('hidden'),
        reconciledTabLabel: document.getElementById('tabLedgerReconciled').textContent.replace(/\\s+/g, ' ').trim(),
        hd0301Status: hd0301Row?.querySelector('.col-status')?.textContent.trim(),
        hd0301PinDisabled: hd0301Row?.querySelector('[data-row-action="pin"]')?.disabled,
      };
    })()`);

    assert.equal(afterUndo.batchesHidden, true);
    assert.equal(afterUndo.reconciledTabLabel, 'Đã đối soát (0)');
    assert.equal(afterUndo.hd0301Status, 'Chưa đối soát');
    assert.equal(afterUndo.hd0301PinDisabled, false);

    await evalPage(`document.getElementById('btnRunSearch').click()`);
    await waitForSearchIdle();
    const restoredCount = await evalPage(
      `document.querySelectorAll('[data-solution-card]').length`
    );
    assert.equal(restoredCount, 2, 'Sau khi hoàn tác Đợt #1, tìm lại phải trả về đủ 2 phương án');

    // 5) Kiểm tra nút "Cỡ chữ: Lớn" đặt data-font-size="large" trên <html>, lưu localStorage và bền vững sau khi tải lại trang
    const fontLargeState = await evalPage(`(() => {
      document.getElementById('btnFontLarge').click();
      return {
        htmlAttr: document.documentElement.getAttribute('data-font-size'),
        stored: localStorage.getItem('af_font_scale'),
        largeBtnActive: document.getElementById('btnFontLarge').classList.contains('active'),
        normalBtnActive: document.getElementById('btnFontNormal').classList.contains('active'),
      };
    })()`);

    assert.equal(fontLargeState.htmlAttr, 'large');
    assert.equal(fontLargeState.stored, 'large');
    assert.equal(fontLargeState.largeBtnActive, true);
    assert.equal(fontLargeState.normalBtnActive, false);

    // Tải lại trang để kiểm chứng inline script trên <head> đọc lại localStorage
    await cdp.send('Page.reload', { ignoreCache: true });
    await waitForPageCondition(
      `document.readyState === 'complete' && document.getElementById('btnFontLarge')`
    );

    const fontPersistedAfterReload = await evalPage(`(() => ({
      htmlAttr: document.documentElement.getAttribute('data-font-size'),
      stored: localStorage.getItem('af_font_scale'),
      largeBtnActive: document.getElementById('btnFontLarge').classList.contains('active'),
    }))()`);

    assert.equal(fontPersistedAfterReload.htmlAttr, 'large');
    assert.equal(fontPersistedAfterReload.stored, 'large');
    assert.equal(fontPersistedAfterReload.largeBtnActive, true);

    // Đưa về cỡ chữ chuẩn
    const fontNormalRestored = await evalPage(`(() => {
      document.getElementById('btnFontNormal').click();
      return {
        htmlAttr: document.documentElement.getAttribute('data-font-size'),
        stored: localStorage.getItem('af_font_scale'),
      };
    })()`);
    assert.equal(fontNormalRestored.htmlAttr, 'normal');
    assert.equal(fontNormalRestored.stored, 'normal');
  });

  // ==========================================================================
  // CA 5: CÁC CA KIỂM THỬ ĐỐI KHÁNG (ADVERSARIAL BUG REPRODUCTIONS)
  // Giữ nguyên assertion đúng kỳ vọng nghiệp vụ kế toán để tái hiện lỗi rõ ràng.
  // ==========================================================================
  test('5.1 [Adversarial] Dán số âm kèm ký hiệu tiền tệ sao kê "(1.500.000) VND" hoặc "VND -2,450,000" vào #targetAmountInput phải giữ nguyên dấu âm', async () => {
    await navigateFreshApp();

    const pastedNegWithCurrency = await evalPage(`(() => {
      const el = document.getElementById('targetAmountInput');

      el.value = '(1.500.000) VND';
      el.dispatchEvent(new Event('input', { bubbles: true }));
      const caseParenVnd = {
        value: el.value,
        words: document.getElementById('amountWordsDisplay').textContent.trim(),
      };

      el.value = 'VND -2,450,000';
      el.dispatchEvent(new Event('input', { bubbles: true }));
      const casePrefixVnd = {
        value: el.value,
        words: document.getElementById('amountWordsDisplay').textContent.trim(),
      };

      return { caseParenVnd, casePrefixVnd };
    })()`);

    assert.equal(
      pastedNegWithCurrency.caseParenVnd.value,
      '-1.500.000',
      'Chuỗi "(1.500.000) VND" là số âm kế toán, không được mất dấu trừ thành số dương'
    );
    assert.equal(
      pastedNegWithCurrency.caseParenVnd.words,
      'Bằng chữ: Âm một triệu năm trăm nghìn đồng'
    );
    assert.equal(
      pastedNegWithCurrency.casePrefixVnd.value,
      '-2.450.000',
      'Chuỗi "VND -2,450,000" là số âm sao kê, không được mất dấu trừ thành số dương'
    );
    assert.equal(
      pastedNegWithCurrency.casePrefixVnd.words,
      'Bằng chữ: Âm hai triệu bốn trăm năm mươi nghìn đồng'
    );
  });

  test('5.2 [Adversarial] Sau khi tự động đảo Từ ngày > Đến ngày lúc bấm Tìm, bộ đếm ứng viên & bảng Sổ chi tiết phải đồng bộ lại theo khoảng ngày mới', async () => {
    await navigateFreshApp();

    await evalPage(`(() => {
      document.getElementById('btnLoadSample').click();
      const fromEl = document.getElementById('dateFromInput');
      const toEl = document.getElementById('dateToInput');
      fromEl.value = '2026-03-15';
      fromEl.dispatchEvent(new Event('change', { bubbles: true }));
      toEl.value = '2026-03-01';
      toEl.dispatchEvent(new Event('change', { bubbles: true }));

      const amountEl = document.getElementById('targetAmountInput');
      amountEl.value = '18.450.000';
      amountEl.dispatchEvent(new Event('input', { bubbles: true }));

      document.getElementById('btnRunSearch').click();
    })()`);

    await waitForSearchIdle();

    const uiSyncAfterSwap = await evalPage(`(() => ({
      candidatePreview: document.getElementById('candidatePreviewInfo').textContent.trim(),
      filteredTabText: document.getElementById('tabLedgerFiltered').textContent.replace(/\\s+/g, ' ').trim(),
      ledgerRowCount: document.querySelectorAll('#ledgerTableBody tr[data-row-id]').length,
    }))()`);

    assert.equal(
      uiSyncAfterSwap.candidatePreview,
      'Đủ điều kiện xét gom: 9 dòng',
      '#candidatePreviewInfo không được kẹt ở "0 dòng" sau khi ngày đã tự động đảo về 01/03/2026 – 15/03/2026'
    );
    assert.equal(
      uiSyncAfterSwap.filteredTabText,
      'Các dòng trong khoảng ngày đang lọc (9)',
      'Tab sổ chi tiết đang lọc phải cập nhật lên (9) dòng thay vì kẹt ở (0)'
    );
    assert.equal(uiSyncAfterSwap.ledgerRowCount, 9);
  });

  test('5.3 [Adversarial] Bật #includeSummaryCheckbox khi đang chọn "Toàn bộ ngày" phải đưa dòng Tổng cộng cuối kỳ (không có ngày) vào tập xét gom', async () => {
    await navigateFreshApp();
    await uploadDynamicTwoSheetWorkbookAndSelectSheet2();

    await evalPage(`(() => {
      const sheetSel = document.getElementById('sheetSelect');
      sheetSel.value = 'So_Chi_Tiet_131';
      sheetSel.dispatchEvent(new Event('change', { bubbles: true }));

      // Đảm bảo đang chọn chế độ "Toàn bộ ngày"
      document.querySelector('#monthQuickBar button[data-month-key="all"]').click();

      // Bật checkbox "Cho phép xét cả các dòng Tổng cộng / Cộng cuối kỳ khi tìm tổ hợp"
      document.getElementById('btnToggleMapping').click();
      const sumChk = document.getElementById('includeSummaryCheckbox');
      sumChk.checked = true;
      sumChk.dispatchEvent(new Event('change', { bubbles: true }));
      document.getElementById('btnApplyMapping').click();
    })()`);

    const summaryIncludedCheck = await evalPage(`(() => ({
      candidatePreview: document.getElementById('candidatePreviewInfo').textContent.trim(),
      filteredTabText: document.getElementById('tabLedgerFiltered').textContent.replace(/\\s+/g, ' ').trim(),
    }))()`);

    assert.equal(
      summaryIncludedCheck.candidatePreview,
      'Đủ điều kiện xét gom: 8 dòng',
      'Khi bật #includeSummaryCheckbox ở chế độ "Toàn bộ ngày", dòng Tổng cộng cuối kỳ (không có ngày) phải được tính vào 8 dòng ứng viên'
    );
  });

  test('5.4 [Adversarial] Khi tắt #includeSummaryCheckbox thì dòng Tổng cộng đang ghim trước đó phải tự động gỡ ghim và không được lọt vào kết quả tìm kiếm', async () => {
    await navigateFreshApp();
    await uploadDynamicTwoSheetWorkbookAndSelectSheet2();

    await evalPage(`(() => {
      const sheetSel = document.getElementById('sheetSelect');
      sheetSel.value = 'So_Chi_Tiet_131';
      sheetSel.dispatchEvent(new Event('change', { bubbles: true }));

      document.getElementById('tabLedgerAll').click();
      document.getElementById('btnToggleMapping').click();

      const sumChk = document.getElementById('includeSummaryCheckbox');
      sumChk.checked = true;
      sumChk.dispatchEvent(new Event('change', { bubbles: true }));

      // Ghim dòng Tổng cộng (row-13 = 36.000.000) khi đang cho phép dòng Tổng cộng
      document.querySelector('#ledgerTableBody button[data-row-action="pin"][data-item-id="row-13"]').click();

      // Tắt lại #includeSummaryCheckbox ("mặc định tự động loại bỏ để tránh tính trùng")
      sumChk.checked = false;
      sumChk.dispatchEvent(new Event('change', { bubbles: true }));
      document.getElementById('btnApplyMapping').click();

      const targetEl = document.getElementById('targetAmountInput');
      targetEl.value = '36.000.000';
      targetEl.dispatchEvent(new Event('input', { bubbles: true }));
      document.querySelector('[data-tolerance="0"]').click();
      document.getElementById('btnRunSearch').click();
    })()`);

    await waitForSearchIdle();

    const summaryLeakCheck = await evalPage(`(() => {
      const row13 = document.querySelector('#ledgerTableBody tr[data-row-id="row-13"]');
      const row13Status = row13 ? row13.querySelector('.col-status')?.textContent.trim() : '';
      const hasRow13InSolutions = Array.from(
        document.querySelectorAll('#solutionsContainer tbody tr .col-excel-row')
      ).some((td) => td.textContent.trim() === '#13');

      return { row13Status, hasRow13InSolutions };
    })()`);

    assert.equal(
      summaryLeakCheck.row13Status,
      'Dòng tổng cộng (Bỏ qua)',
      'Khi tắt #includeSummaryCheckbox, dòng Tổng cộng không được giữ trạng thái Ghim bắt buộc bị khóa cứng'
    );
    assert.equal(
      summaryLeakCheck.hasRow13InSolutions,
      false,
      'Dòng Tổng cộng #13 không được phép lọt vào kết quả khi #includeSummaryCheckbox đang tắt'
    );
  });

  test('5.5 [Adversarial] Khi bỏ tích 1 dòng lẻ sau khi bấm Chọn tất cả (#chkSelectAllVisible), ô checkbox tiêu đề phải tự bỏ trạng thái checked', async () => {
    await navigateFreshApp();
    await uploadDynamicTwoSheetWorkbookAndSelectSheet2();

    const selectAllSync = await evalPage(`(() => {
      const sheetSel = document.getElementById('sheetSelect');
      sheetSel.value = 'So_Chi_Tiet_131';
      sheetSel.dispatchEvent(new Event('change', { bubbles: true }));

      const selectAll = document.getElementById('chkSelectAllVisible');
      selectAll.checked = true;
      selectAll.dispatchEvent(new Event('change', { bubbles: true }));

      const firstRowChk = document.querySelector('#ledgerTableBody input[data-row-action="check"][data-item-id="row-6"]');
      firstRowChk.checked = false;
      firstRowChk.dispatchEvent(new Event('change', { bubbles: true }));

      return {
        selectAllStillChecked: selectAll.checked,
      };
    })()`);

    assert.equal(
      selectAllSync.selectAllStillChecked,
      false,
      'Khi bỏ tích 1 dòng lẻ, #chkSelectAllVisible phải bỏ trạng thái checked=true để tránh ngược hành vi ở lần bấm kế tiếp'
    );
  });

  // ==========================================================================
  // CA 6: KIỂM TRA SỨC KHỎE RUNTIME TRÌNH DUYỆT (0 CONSOLE.ERROR & 0 EXCEPTION)
  // ==========================================================================
  test('6.1 Toàn bộ phiên E2E đạt chuẩn 0 console.error, 0 Runtime.exceptionThrown và 0 lỗi tải tài nguyên HTTP', () => {
    assert.deepEqual(
      consoleErrors,
      [],
      `Phát hiện console.error trên trình duyệt: ${JSON.stringify(consoleErrors)}`
    );
    assert.deepEqual(
      runtimeExceptions,
      [],
      `Phát hiện Runtime.exceptionThrown chưa bắt: ${JSON.stringify(runtimeExceptions)}`
    );
    assert.deepEqual(
      browserLogErrors,
      [],
      `Phát hiện lỗi Log.entryAdded: ${JSON.stringify(browserLogErrors)}`
    );
    assert.deepEqual(
      serverInfo.httpErrors,
      [],
      `Phát hiện lỗi HTTP 4xx/5xx khi tải tài nguyên dưới /auto-finance/: ${JSON.stringify(
        serverInfo.httpErrors
      )}`
    );
  });
});
