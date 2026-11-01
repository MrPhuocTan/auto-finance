/**
 * Shared Data Contracts & Module Interfaces for Auto Finance (Đối Soát & Gom Khoản Tiền Excel)
 *
 * CRITICAL HOSTING NOTE (GitHub Pages Subpath `/auto-finance/`):
 * - NEVER use root-absolute paths like `/src/...` or `/public/...`.
 * - ALWAYS use relative paths (`./src/...`, `./public/...`) or `new URL('../core/reconcileWorker.js', import.meta.url)`.
 *
 * ============================================================================
 * 1. `src/core/vnFormat.js` EXPORTS:
 * ============================================================================
 * - `parseVietnameseAmount(raw): number | null`
 *    Parses numbers & VN/accounting formatted strings:
 *    12500000, "12.500.000", "12,500,000", "(500.000)" -> -500000,
 *    "-500.000 đ" -> -500000, "1.250.000,50" -> 1250000.5. Returns null if invalid/empty.
 * - `formatNumberVN(value, options = { showPlusSign: false }): string`
 *    Formats number with dot thousand separator & comma decimal: 12500000 -> "12.500.000", -500000 -> "-500.000".
 * - `formatInputAmountVN(rawText): { numeric: number | null, formatted: string }`
 *    Live input helper for amount input box.
 * - `parseVietnameseDate(raw): { iso: string, display: string, timestamp: number } | null`
 *    Supports Date, Excel serial date, "DD/MM/YYYY", "DD-MM-YYYY", "YYYY-MM-DD", "DD.MM.YY", datetime strings.
 *    Ambiguous XX/YY/ZZZZ MUST default to DD/MM/YYYY (Vietnamese standard).
 *    `timestamp` is `Date.UTC(year, month - 1, day)`.
 * - `formatDateVN(isoOrDate): string` -> "DD/MM/YYYY" or "".
 * - `isoToUtcTimestamp(iso): number | null`
 * - `readAmountInVietnameseWords(amount): string`
 *    E.g., 12500000 -> "Mười hai triệu năm trăm nghìn đồng", -500000 -> "Âm năm trăm nghìn đồng".
 *
 * ============================================================================
 * 2. `src/core/excelParser.js` EXPORTS:
 * ============================================================================
 * - `loadWorkbookFromBuffer(arrayBuffer): WorkBook`
 * - `getSheetNames(workbook): string[]`
 * - `inspectSheet(workbook, sheetName, customHeaderRowIndex = null): SheetInspection`
 * - `inspectRawGrid(rawGrid, sheetName = 'Sheet1', customHeaderRowIndex = null): SheetInspection`
 *    Returns:
 *    {
 *      sheetName: string,
 *      rawRowsCount: number,
 *      detectedHeaderRowIndex: number, // 0-based
 *      previewTopRows: Array<{ rowIndex: number, excelRowNumber: number, cells: string[] }>,
 *      columns: Array<{ index: number, excelCol: string, label: string, sampleValues: string[] }>,
 *      suggestedMapping: {
 *        dateColIndex: number | null,
 *        amountColIndex: number | null,
 *        debitColIndex: number | null,
 *        creditColIndex: number | null,
 *        amountMode: 'single' | 'debit_minus_credit',
 *        descColIndex: number | null,
 *        refColIndex: number | null
 *      },
 *      rawGrid: any[][]
 *    }
 * - `extractTransactions(rawGrid, headerRowIndex, mapping, columns = null): ExtractedSheetData`
 *    Returns:
 *    {
 *      items: TransactionItem[], // Each item has { id, excelRowNumber, dateIso, dateDisplay, timestamp, amount, description, reference, isSummaryRow, extraCells, rawCells }
 *      summaryRowCount: number,
 *      skippedEmptyOrInvalidCount: number,
 *      minDateIso: string | null,
 *      maxDateIso: string | null,
 *      availableMonths: Array<{ key: string, label: string, fromIso: string, toIso: string, count: number }>
 *    }
 * - `exportReconciledWorkbook({ sheetName, columns, items, reconciledBatches, unallocatedOnly = false }): Uint8Array`
 *    Generates `.xlsx` bytes with Sheet 1 (`Chi_Tiet_Doi_Soat`) and Sheet 2 (`Tong_Hop_Cac_Dot_Gom`).
 *
 * ============================================================================
 * 3. `src/core/reconcileEngine.js` & `src/core/reconcileWorker.js` EXPORTS:
 * ============================================================================
 * - `normalizeVietnameseSearch(str): string`
 * - `findCombinations(options, onProgress = null, shouldCancel = null): ReconcileSearchResult`
 *    `options`:
 *    {
 *      items: TransactionItem[],
 *      dateFromIso?: string | null,
 *      dateToIso?: string | null,
 *      keyword?: string,
 *      targetAmount: number,
 *      tolerance?: number,         // default 500
 *      pinnedIds?: string[],       // items forced into every combination
 *      excludedIds?: string[],     // items excluded (manual or already reconciled)
 *      includeSummaryRows?: boolean, // default false (skips item.isSummaryRow)
 *      minItems?: number,          // default 1
 *      maxItems?: number,          // default 12
 *      maxResults?: number,        // default 50
 *      timeLimitMs?: number        // default 5000
 *    }
 *    Returns:
 *    {
 *      solutions: Array<{
 *        id: string,
 *        items: TransactionItem[],
 *        itemIds: string[],
 *        itemCount: number,
 *        totalAmount: number,
 *        targetAmount: number,
 *        diff: number,
 *        absDiff: number,
 *        isExact: boolean,
 *        dateSpanDays: number,
 *        duplicateVariantsNote: string | null
 *      }>,
 *      stats: {
 *        candidateCount: number,
 *        pinnedCount: number,
 *        exploredNodes: number,
 *        elapsedMs: number,
 *        timedOut: boolean,
 *        exactCount: number,
 *        nearCount: number
 *      },
 *      error?: string | null
 *    }
 * - `src/core/reconcileWorker.js`:
 *    Handles messages:
 *      Incoming: `{ type: 'SEARCH', requestId, options }` | `{ type: 'CANCEL', requestId }`
 *      Outgoing: `{ type: 'PROGRESS', requestId, exploredNodes, foundCount, currentSize }`
 *              | `{ type: 'RESULT', requestId, result }`
 *              | `{ type: 'ERROR', requestId, message }`
 *
 * ============================================================================
 * 4. `src/core/sampleData.js` EXPORTS:
 * ============================================================================
 * - `getSampleDataset(): { fileName: string, sheetName: string, rawGrid: any[][], presets: Array<{ label: string, dateFromIso: string, dateToIso: string, targetAmount: number, tolerance: number, keyword: string }> }`
 */

export const DEFAULT_TOLERANCE = 500;
export const DEFAULT_MAX_ITEMS = 12;
export const DEFAULT_MAX_RESULTS = 50;
export const DEFAULT_TIME_LIMIT_MS = 5000;
