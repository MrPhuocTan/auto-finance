/**
 * Vietnamese Accounting Excel / CSV Parser, Auto-Inspector & Reconciled Exporter
 * Pure ES Module compatible with Browser (GitHub Pages) and Node.js ESM.
 */

import * as XLSX from '../../vendor/xlsx.mjs';
import * as cptable from '../../vendor/cpexcel.full.mjs';
import {
  parseVietnameseAmount,
  parseVietnameseDate,
  formatDateVN,
} from './vnFormat.js';

if (typeof XLSX.set_cptable === 'function') {
  XLSX.set_cptable(cptable);
}

/**
 * Normalizes Vietnamese text for accent-insensitive header & summary keyword matching.
 * @param {unknown} input
 * @returns {string}
 */
function normalizeText(input) {
  if (input === null || input === undefined) return '';
  return String(input)
    .toLowerCase()
    .replace(/đ/g, 'd')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

/**
 * Formats any raw cell value into a clean string for display/preview.
 * @param {unknown} cell
 * @returns {string}
 */
function cellToCleanString(cell) {
  if (cell === null || cell === undefined) return '';
  if (cell instanceof Date) return formatDateVN(cell);
  if (typeof cell === 'boolean') return cell ? 'TRUE' : 'FALSE';
  return String(cell)
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Checks whether a row right under the header is a Thông tư 200/133 column-index helper row
 * such as `| A | B | C | D | 1 | 2 |` or `| (1) | (2) | (3) | (4) |`.
 *
 * @param {any[]} row
 * @returns {boolean}
 */
function isColumnIndexHelperRow(row) {
  if (!Array.isArray(row)) return false;
  const nonEmpty = [];
  for (let i = 0; i < row.length; i++) {
    const s = cellToCleanString(row[i]);
    if (s !== '') nonEmpty.push(s);
  }
  if (nonEmpty.length < 3) return false;

  let shortTokenCount = 0;
  for (const token of nonEmpty) {
    // Reject if any cell is a real date string
    if (typeof token === 'string' && token.length >= 6 && parseVietnameseDate(token)) {
      return false;
    }
    if (/^\(?([a-zA-Z]{1,2}|\d{1,2})\)?$/.test(token)) {
      shortTokenCount++;
    }
  }

  return shortTokenCount / nonEmpty.length >= 0.75;
}

/**
 * Checks if a single cell text contains accounting header keywords and classifies its category.
 * @param {string} rawText
 * @returns {{ score: number, category: string | null }}
 */
function scoreHeaderCellKeyword(rawText) {
  const lower = rawText.toLowerCase().trim();
  const norm = normalizeText(rawText);
  if (!norm || norm.length > 80) return { score: 0, category: null };

  // Date keywords
  if (
    /\b(ngay|date|thoi gian|timestamp)\b/.test(norm) ||
    /^(ngay ct|ngay hd|ngay gd|posting date|value date|transaction date)$/.test(norm)
  ) {
    return { score: 14, category: 'date' };
  }

  // Amount / Debit / Credit keywords
  if (
    /\b(so tien|thanh tien|phat sinh|con lai|phai thu|phai tra|thuc thu|thuc chi|tong thanh toan|gia tri|ghi no|ghi co|ps no|ps co|tien vao|tien ra|amount|debit|credit|balance|withdrawal|deposit)\b/.test(
      norm
    ) ||
    lower === 'nợ' ||
    lower === 'có'
  ) {
    return { score: 15, category: 'amount' };
  }

  // Description keywords
  if (
    /\b(dien giai|noi dung|chi tiet|mo ta|trich yeu|ly do|ghi chu|description|particulars|narrative|memo|remarks)\b/.test(
      norm
    )
  ) {
    return { score: 14, category: 'desc' };
  }

  // Reference / Partner / Index keywords
  if (
    /\b(chung tu|so ct|so hieu|hoa don|so hd|ky hieu|tham chieu|ma gd|ma kh|khach hang|doi tuong|nha cung cap|tai khoan|tk no|tk co|reference|ref|invoice|voucher|doc no|stt|so thu tu)\b/.test(
      norm
    )
  ) {
    return { score: 12, category: 'ref' };
  }

  return { score: 0, category: null };
}

/**
 * Evaluates whether a row looks like a pure text parent header row (above a 2-row merged header).
 * @param {any[]} row
 * @returns {boolean}
 */
function isLikelyParentHeaderRow(row) {
  if (!Array.isArray(row)) return false;
  let nonEmptyCount = 0;
  let textCount = 0;
  let keywordHits = 0;

  for (const cell of row) {
    const s = cellToCleanString(cell);
    if (!s) continue;
    nonEmptyCount++;
    if (typeof cell === 'number') return false;
    if (parseVietnameseDate(s) || parseVietnameseAmount(s) !== null) return false;
    textCount++;
    if (scoreHeaderCellKeyword(s).score > 0) keywordHits++;
  }

  return nonEmptyCount >= 3 && textCount === nonEmptyCount && keywordHits >= 2;
}

/**
 * Auto-detects the 0-based header row index within the first 25 rows of a sheet grid.
 * @param {any[][]} grid
 * @returns {number}
 */
function detectHeaderRowIndex(grid) {
  if (!Array.isArray(grid) || grid.length === 0) return 0;

  const scanLimit = Math.min(25, grid.length);
  let bestRowIndex = 0;
  let bestScore = -Infinity;

  for (let r = 0; r < scanLimit; r++) {
    const row = grid[r];
    if (!Array.isArray(row)) continue;

    let nonEmptyCount = 0;
    let textCount = 0;
    let dataLikeCells = 0;
    let keywordScore = 0;
    const categories = new Set();
    const seenTexts = new Set();
    let longParagraphCells = 0;

    for (let c = 0; c < row.length; c++) {
      const rawCell = row[c];
      const str = cellToCleanString(rawCell);
      if (!str) continue;
      nonEmptyCount++;

      if (str.length > 90) {
        longParagraphCells++;
      }

      if (typeof rawCell === 'number') {
        dataLikeCells++;
        continue;
      }

      if (parseVietnameseDate(str) !== null || parseVietnameseAmount(str) !== null) {
        dataLikeCells++;
        continue;
      }

      textCount++;
      seenTexts.add(str.toLowerCase());

      const kw = scoreHeaderCellKeyword(str);
      if (kw.score > 0) {
        keywordScore += Math.min(kw.score, 15);
        if (kw.category) categories.add(kw.category);
      }
    }

    if (nonEmptyCount === 0) continue;

    let score = 0;

    // Penalize single-cell title/banner rows (company name, report title, date range subtitle)
    if (nonEmptyCount === 1) {
      score -= 45;
    } else if (nonEmptyCount === 2) {
      score -= 8;
    }

    if (isColumnIndexHelperRow(row)) {
      score -= 70;
    }

    // Reward distinct text headers & keyword coverage
    score += Math.min(seenTexts.size, 10) * 3.5;
    if (textCount > 0) {
      score += (seenTexts.size / textCount) * 8;
    }
    score += Math.min(keywordScore, 75);
    score += categories.size * 12;

    // Penalize rows that already contain dates/numbers (likely data rows) or long notes
    score -= dataLikeCells * 10;
    score -= longParagraphCells * 15;

    // 2-row merged header awareness: inherit parent header categories from r - 1
    if (r > 0 && textCount >= 2 && dataLikeCells === 0 && isLikelyParentHeaderRow(grid[r - 1])) {
      const parentRow = grid[r - 1];
      for (let c = 0; c < parentRow.length; c++) {
        const childStr = cellToCleanString(row[c]);
        const parentStr = cellToCleanString(parentRow[c]);
        if (!childStr && parentStr) {
          const kw = scoreHeaderCellKeyword(parentStr);
          if (kw.score > 0) {
            score += kw.score * 0.8;
            if (kw.category) categories.add(kw.category);
          }
        }
      }
      // Slight preference for bottom sub-header row (immediately above data)
      score += 8;
    }

    // Lookahead: inspect up to 5 non-empty rows below `r` for real dates & amounts
    let inspectedBelow = 0;
    let validDataRowsBelow = 0;
    for (let nextR = r + 1; nextR < grid.length && inspectedBelow < 5; nextR++) {
      const nextRow = grid[nextR];
      if (!Array.isArray(nextRow)) continue;
      const hasAny = nextRow.some((cell) => cellToCleanString(cell) !== '');
      if (!hasAny) continue;

      if (nextR <= r + 2 && isColumnIndexHelperRow(nextRow)) {
        continue;
      }

      inspectedBelow++;
      let rowHasAmount = false;
      let rowHasDate = false;
      for (const cell of nextRow) {
        if (!rowHasDate && cell !== '' && cell !== null && cell !== undefined) {
          // For strings, check date directly; for numbers, check typical serial window
          if (typeof cell === 'string' && parseVietnameseDate(cell) !== null) {
            rowHasDate = true;
          } else if (typeof cell === 'number' && cell >= 32874 && cell <= 55000 && cell % 100 !== 0) {
            rowHasDate = true;
          }
        }
        if (!rowHasAmount && parseVietnameseAmount(cell) !== null) {
          // Exclude strings that are dates
          if (typeof cell === 'number' || parseVietnameseDate(cell) === null) {
            rowHasAmount = true;
          }
        }
      }
      if (rowHasAmount && rowHasDate) {
        validDataRowsBelow += 1.5;
      } else if (rowHasAmount) {
        validDataRowsBelow += 0.8;
      }
    }

    score += Math.min(validDataRowsBelow * 6, 30);

    if (score > bestScore) {
      bestScore = score;
      bestRowIndex = r;
    }
  }

  return bestRowIndex;
}

/**
 * Resolves column metadata (`index`, `excelCol`, `label`, `sampleValues`) from the grid.
 * Supports 2-row Vietnamese accounting headers where parent group labels sit at `headerRowIndex - 1`.
 *
 * @param {any[][]} grid
 * @param {number} headerRowIndex
 * @returns {Array<{ index: number, excelCol: string, label: string, sampleValues: string[] }>}
 */
function buildColumnsMetadata(grid, headerRowIndex) {
  let maxCols = 0;
  for (let r = 0; r < grid.length; r++) {
    const row = grid[r];
    if (!Array.isArray(row)) continue;
    for (let c = row.length - 1; c >= maxCols; c--) {
      if (cellToCleanString(row[c]) !== '') {
        maxCols = c + 1;
        break;
      }
    }
  }

  if (maxCols === 0) return [];

  const headerRow = Array.isArray(grid[headerRowIndex]) ? grid[headerRowIndex] : [];
  const hasParentHeader =
    headerRowIndex > 0 && isLikelyParentHeaderRow(grid[headerRowIndex - 1]);
  const parentRow = hasParentHeader ? grid[headerRowIndex - 1] : null;

  // Propagate horizontal group labels in parent header (e.g. "Số phát sinh" spanning Nợ | Có)
  const parentGroupLabels = [];
  if (parentRow) {
    let activeGroup = '';
    for (let c = 0; c < maxCols; c++) {
      const pText = cellToCleanString(parentRow[c]);
      const cText = cellToCleanString(headerRow[c]);
      if (pText) {
        parentGroupLabels[c] = pText;
        activeGroup = cText ? pText : '';
      } else if (cText && activeGroup) {
        parentGroupLabels[c] = activeGroup;
      } else {
        activeGroup = '';
        parentGroupLabels[c] = '';
      }
    }
  }

  const columns = [];
  for (let c = 0; c < maxCols; c++) {
    const excelCol = XLSX.utils.encode_col(c);
    const childLabel = cellToCleanString(headerRow[c]);
    const parentLabel = parentRow ? parentGroupLabels[c] || '' : '';

    let label = childLabel;
    if (!label && parentLabel) {
      label = parentLabel;
    } else if (
      label &&
      parentLabel &&
      parentLabel.toLowerCase() !== label.toLowerCase() &&
      !label.toLowerCase().includes(parentLabel.toLowerCase())
    ) {
      // Prefix short ambiguous subheaders like "Nợ", "Có", "Số hiệu", "Ngày tháng" with parent group
      const childNorm = normalizeText(label);
      if (
        label.toLowerCase() === 'nợ' ||
        label.toLowerCase() === 'có' ||
        /^(so hieu|so|ngay thang|ngay|vnd|nguyen te)$/.test(childNorm)
      ) {
        label = `${parentLabel} - ${label}`;
      }
    }

    if (!label) {
      label = `Cột ${excelCol}`;
    }

    const sampleValues = [];
    for (let r = headerRowIndex + 1; r < grid.length && sampleValues.length < 4; r++) {
      const row = grid[r];
      if (!Array.isArray(row)) continue;
      if (r <= headerRowIndex + 2 && isColumnIndexHelperRow(row)) continue;
      const rawVal = row[c];
      const strVal = cellToCleanString(rawVal);
      if (!strVal) continue;
      if (detectSummaryRow(row, null, null)) continue;
      sampleValues.push(strVal);
    }

    columns.push({
      index: c,
      excelCol,
      label,
      sampleValues,
    });
  }

  return columns;
}

/**
 * Infers the best column mapping (`dateColIndex`, `amountColIndex`, `debitColIndex`, `creditColIndex`,
 * `amountMode`, `descColIndex`, `refColIndex`) using both Vietnamese/English headers and data profiles.
 *
 * @param {any[][]} grid
 * @param {number} headerRowIndex
 * @param {Array<{ index: number, excelCol: string, label: string, sampleValues: string[] }>} columns
 * @returns {{
 *   dateColIndex: number | null,
 *   amountColIndex: number | null,
 *   debitColIndex: number | null,
 *   creditColIndex: number | null,
 *   amountMode: 'single' | 'debit_minus_credit',
 *   descColIndex: number | null,
 *   refColIndex: number | null
 * }}
 */
function inferSuggestedMapping(grid, headerRowIndex, columns) {
  const emptyMapping = {
    dateColIndex: null,
    amountColIndex: null,
    debitColIndex: null,
    creditColIndex: null,
    amountMode: 'single',
    descColIndex: null,
    refColIndex: null,
  };

  if (columns.length === 0) return emptyMapping;

  // Collect up to 50 non-summary data rows for statistical profiling
  const sampleRows = [];
  for (let r = headerRowIndex + 1; r < grid.length && sampleRows.length < 50; r++) {
    const row = grid[r];
    if (!Array.isArray(row)) continue;
    if (!row.some((cell) => cellToCleanString(cell) !== '')) continue;
    if (r <= headerRowIndex + 2 && isColumnIndexHelperRow(row)) continue;
    if (detectSummaryRow(row, null, null)) continue;
    sampleRows.push(row);
  }

  const profiles = columns.map((col) => {
    const lower = col.label.toLowerCase().trim();
    const norm = normalizeText(col.label);

    let nonEmpty = 0;
    let stringDateCount = 0;
    let serialDateCount = 0;
    let numericCount = 0;
    let roundThousandCount = 0;
    let textCount = 0;
    let totalTextLen = 0;
    let sequentialIndexMatches = 0;

    for (let i = 0; i < sampleRows.length; i++) {
      const val = sampleRows[i][col.index];
      const str = cellToCleanString(val);
      if (!str) continue;
      nonEmpty++;

      if (typeof val === 'string' && parseVietnameseDate(val) !== null) {
        stringDateCount++;
        continue;
      }

      if (typeof val === 'number' && Number.isFinite(val)) {
        if (val === i + 1) sequentialIndexMatches++;
        if (val >= 32874 && val <= 73415 && Number.isInteger(Math.floor(val + 1e-9))) {
          serialDateCount++;
        }
        if (Math.abs(val) >= 1000 && val % 1000 === 0) {
          roundThousandCount++;
        }
        numericCount++;
        continue;
      }

      const parsedAmt = parseVietnameseAmount(val);
      if (parsedAmt !== null) {
        numericCount++;
        if (Math.abs(parsedAmt) >= 1000 && parsedAmt % 1000 === 0) {
          roundThousandCount++;
        }
        if (parsedAmt === i + 1) sequentialIndexMatches++;
      } else {
        textCount++;
        totalTextLen += str.length;
      }
    }

    const denom = Math.max(nonEmpty, 1);
    const stringDateRatio = stringDateCount / denom;
    const serialDateRatio = serialDateCount / denom;
    const numericRatio = numericCount / denom;
    const roundThousandRatio = numericCount > 0 ? roundThousandCount / numericCount : 0;
    const textRatio = textCount / denom;
    const avgTextLen = textCount > 0 ? totalTextLen / textCount : 0;
    const isLikelySttData = nonEmpty >= 3 && sequentialIndexMatches / denom >= 0.8;

    const isSttHeader = /^(stt|so thu tu|thu tu|tt|no)$/.test(norm);
    const isAccountCodeHeader =
      /\b(tk|tai khoan|dinh khoan|tk doi ung|ma tk|debit account|credit account)\b/.test(norm) &&
      !/\b(so tien|phat sinh|so du)\b/.test(norm);

    // 1. Date score
    let dateScore = -10;
    if (isSttHeader || isAccountCodeHeader) {
      dateScore -= 150;
    } else if (
      /\b(ngay chung tu|ngay ct|ngay giao dich|ngay gd|ngay hoa don|ngay hd|ngay hach toan|ngay ht|ngay ghi so|ngay thanh toan|ngay thu|ngay chi|ngay thang|transaction date|posting date|invoice date|value date|doc date)\b/.test(
        norm
      ) ||
      norm === 'ngay' ||
      norm === 'date'
    ) {
      dateScore += 65;
    } else if (/\b(ngay|date)\b/.test(norm)) {
      dateScore += 40;
      if (/\b(dao han|han thanh toan|due date|ngay sinh|ngay tao|ngay cap)\b/.test(norm)) {
        dateScore -= 25;
      }
    } else if (
      /\b(so tien|thanh tien|phat sinh|dien giai|noi dung|so ct|hoa don|khach hang)\b/.test(norm)
    ) {
      dateScore -= 100;
    }

    if (nonEmpty > 0) {
      if (stringDateRatio >= 0.5) {
        dateScore += 55 * stringDateRatio;
      } else if (serialDateRatio >= 0.5 && (roundThousandRatio < 0.35 || /\b(ngay|date)\b/.test(norm))) {
        dateScore += 42 * serialDateRatio;
      } else if (stringDateCount === 0 && serialDateCount === 0) {
        dateScore -= 35;
      }
    }

    // 2. Debit & Credit scores
    const isBalanceHeader = /\b(so du|du cuoi|du dau|cuoi ky|dau ky|luy ke|balance|closing|opening)\b/.test(
      norm
    );
    const isPureDebitHeader =
      !isAccountCodeHeader &&
      (/\b(phat sinh no|ps no|so phat sinh no|ghi no|so tien ghi no|tien ra|chi ra|thuc chi|debit|withdrawal)\b/.test(
        norm
      ) ||
        lower === 'nợ' ||
        /(?:^|[\s-])nợ$/.test(lower));

    const isPureCreditHeader =
      !isAccountCodeHeader &&
      (/\b(phat sinh co|ps co|so phat sinh co|ghi co|so tien ghi co|tien vao|thu vao|thuc thu|credit|deposit)\b/.test(
        norm
      ) ||
        lower === 'có' ||
        /(?:^|[\s-])có$/.test(lower));

    let debitScore = -20;
    if (isPureDebitHeader) {
      debitScore = isBalanceHeader ? 18 : 75;
      if (nonEmpty > 0 && numericRatio > 0) debitScore += 15;
    }

    let creditScore = -20;
    if (isPureCreditHeader) {
      creditScore = isBalanceHeader ? 18 : 75;
      if (nonEmpty > 0 && numericRatio > 0) creditScore += 15;
    }

    // 3. Single Amount score
    let singleAmountScore = -10;
    if (
      isSttHeader ||
      isLikelySttData ||
      isAccountCodeHeader ||
      isPureDebitHeader ||
      isPureCreditHeader ||
      /\b(ngay|date|dien giai|noi dung|so chung tu|so ct|so hieu|so hoa don|so hd|ma kh|khach hang|so luong|don gia|ty gia|thue suat)\b/.test(
        norm
      )
    ) {
      singleAmountScore = -150;
    } else if (
      /\b(so tien|con lai|con phai thu|con phai tra|tong thanh toan|tong tien thanh toan|tong cong tien|gia tri thanh toan|so tien gd|so tien giao dich|amount|net amount|total amount|remaining)\b/.test(
        norm
      )
    ) {
      singleAmountScore = 72;
    } else if (/\b(thanh tien)\b/.test(norm)) {
      singleAmountScore = /\b(truoc thue|chua thue)\b/.test(norm) ? 48 : 68;
    } else if (/\b(phat sinh|so phat sinh|tong tien|phai thu|phai tra|gia tri)\b/.test(norm)) {
      singleAmountScore = 45;
    } else if (/\b(thue gtgt|tien thue|vat|tax)\b/.test(norm)) {
      singleAmountScore = 15;
    } else if (isBalanceHeader) {
      singleAmountScore = 8;
    }

    if (sampleRows.length > 0) {
      if (numericRatio >= 0.4 && stringDateRatio < 0.3) {
        singleAmountScore += 28 * numericRatio;
      } else if (nonEmpty === 0 || numericCount === 0) {
        singleAmountScore -= 35;
      }
    }

    // 4. Description score
    let descScore = -10;
    if (isSttHeader || isAccountCodeHeader) {
      descScore = -120;
    } else if (
      /\b(dien giai|noi dung|chi tiet|mo ta|trich yeu|ly do|description|particulars|narrative|memo|remarks)\b/.test(
        norm
      )
    ) {
      descScore = 75;
    } else if (/\b(ghi chu|note)\b/.test(norm)) {
      descScore = 42;
    } else if (
      /\b(khach hang|ten khach hang|doi tuong|ten doi tuong|nha cung cap|nguoi nop|nguoi nhan|partner|customer|vendor)\b/.test(
        norm
      ) &&
      !/^ma\b/.test(norm)
    ) {
      descScore = 45;
    }

    if (nonEmpty > 0 && textRatio >= 0.5) {
      descScore += 15 * textRatio + Math.min(20, avgTextLen * 0.6);
    }

    // 5. Reference score (Số CT / Số HĐ / Mã KH)
    let refScore = -10;
    if (
      isSttHeader ||
      isLikelySttData ||
      isAccountCodeHeader ||
      /\b(so tien|thanh tien|so du|so luong|phat sinh|ngay|date|dien giai|noi dung)\b/.test(norm)
    ) {
      refScore = -120;
    } else if (
      /\b(so chung tu|so ct|so hieu|so hoa don|so hd|ma chung tu|so tham chieu|ma giao dich|ma gd|but toan|chung tu|hoa don|voucher|invoice|doc no|document no|reference|ref)\b/.test(
        norm
      )
    ) {
      refScore = 72;
    } else if (/\b(ma kh|ma khach hang|ma doi tuong|ma ncc|ky hieu|seri|code)\b/.test(norm)) {
      refScore = 52;
    }

    if (nonEmpty > 0 && avgTextLen >= 2 && avgTextLen <= 28 && stringDateRatio < 0.2) {
      refScore += 10;
    }

    return {
      index: col.index,
      dateScore,
      singleAmountScore,
      debitScore,
      creditScore,
      descScore,
      refScore,
    };
  });

  const usedCols = new Set();
  const pickBest = (scoreKey, minThreshold) => {
    let bestIdx = null;
    let bestVal = minThreshold;
    for (const p of profiles) {
      if (usedCols.has(p.index)) continue;
      if (p[scoreKey] > bestVal) {
        bestVal = p[scoreKey];
        bestIdx = p.index;
      }
    }
    return { index: bestIdx, score: bestVal };
  };

  // 1. Pick Date column
  const bestDate = pickBest('dateScore', 15);
  const dateColIndex = bestDate.index;
  if (dateColIndex !== null) usedCols.add(dateColIndex);

  // 2. Decide Amount mode: single vs debit_minus_credit
  const bestDebit = pickBest('debitScore', 40);
  if (bestDebit.index !== null) usedCols.add(bestDebit.index);
  const bestCredit = pickBest('creditScore', 40);
  if (bestDebit.index !== null) usedCols.delete(bestDebit.index); // release temporarily until mode chosen

  const bestSingle = pickBest('singleAmountScore', 12);

  let amountMode = 'single';
  let amountColIndex = null;
  let debitColIndex = null;
  let creditColIndex = null;

  const hasSplitDebitCredit = bestDebit.index !== null && bestCredit.index !== null;
  const hasStrongSingleAmount = bestSingle.index !== null && bestSingle.score >= 58;

  if (hasSplitDebitCredit && !hasStrongSingleAmount) {
    amountMode = 'debit_minus_credit';
    debitColIndex = bestDebit.index;
    creditColIndex = bestCredit.index;
    amountColIndex = null;
    usedCols.add(debitColIndex);
    usedCols.add(creditColIndex);
  } else if (bestSingle.index !== null) {
    amountMode = 'single';
    amountColIndex = bestSingle.index;
    usedCols.add(amountColIndex);
    if (hasSplitDebitCredit) {
      debitColIndex = bestDebit.index;
      creditColIndex = bestCredit.index;
    }
  } else if (bestDebit.index !== null || bestCredit.index !== null) {
    amountMode = 'debit_minus_credit';
    debitColIndex = bestDebit.index;
    creditColIndex = bestCredit.index;
    if (debitColIndex !== null) usedCols.add(debitColIndex);
    if (creditColIndex !== null) usedCols.add(creditColIndex);
  }

  // 3. Pick Description column
  const bestDesc = pickBest('descScore', 5);
  const descColIndex = bestDesc.index;
  if (descColIndex !== null) usedCols.add(descColIndex);

  // 4. Pick Reference column
  const bestRef = pickBest('refScore', 15);
  const refColIndex = bestRef.index;

  return {
    dateColIndex,
    amountColIndex,
    debitColIndex,
    creditColIndex,
    amountMode,
    descColIndex,
    refColIndex,
  };
}

/**
 * Detects whether a row is an accounting summary/total/balance row ("Tổng cộng", "Cộng phát sinh",
 * "Số dư cuối kỳ", "Số dư đầu kỳ", "Subtotal", "Total", or blank date + "Cộng").
 * Avoids false positives on normal transactions like "Tổng công ty" or "Cộng tác viên".
 *
 * @param {any[]} row
 * @param {number | null} dateColIndex
 * @param {number | null} descColIndex
 * @returns {{ isSummary: boolean, matchedLabel: string }}
 */
function detectSummaryRowInternal(row, dateColIndex, descColIndex) {
  if (!Array.isArray(row)) return { isSummary: false, matchedLabel: '' };

  const hasValidDate =
    dateColIndex !== null &&
    dateColIndex !== undefined &&
    parseVietnameseDate(row[dateColIndex]) !== null;

  for (let c = 0; c < row.length; c++) {
    const raw = row[c];
    if (typeof raw !== 'string') continue;
    const clean = cellToCleanString(raw);
    if (!clean) continue;
    const norm = normalizeText(clean);
    if (!norm) continue;

    // Unconditional accounting summary phrases (note negative lookahead excluding "tong cong ty")
    if (
      /\btong cong(?!\s+ty\b)\b/.test(norm) ||
      /\b(cong phat sinh|tong phat sinh|cong ps|tong ps|so du cuoi ky|so du dau ky|du cuoi ky|du dau ky|so du mang sang|cong chuyen sang|cong luy ke|grand total|subtotal|sub total)\b/.test(
        norm
      ) ||
      /^(tong|cong|tong so|tong ket|total|totals)$/.test(norm)
    ) {
      return { isSummary: true, matchedLabel: clean };
    }

    // Conditional summary rule when date column is blank/invalid
    if (!hasValidDate) {
      if (
        /^cong(?!\s+(?:ty|tac|dong|hoa)\b)\b/.test(norm) ||
        /^tong(?!\s+(?:cong\s+ty|giam\s+doc|dai)\b)\b/.test(norm) ||
        /^total\b/.test(norm) ||
        /^so du\b/.test(norm)
      ) {
        return { isSummary: true, matchedLabel: clean };
      }
    }
  }

  // Extra check specifically on descColIndex if date is empty and description says "Cộng"
  if (!hasValidDate && descColIndex !== null && descColIndex !== undefined) {
    const descClean = cellToCleanString(row[descColIndex]);
    const descNorm = normalizeText(descClean);
    if (/^(cong|tong)\b/.test(descNorm) && !/^cong\s+(?:ty|tac|dong|hoa)\b/.test(descNorm)) {
      return { isSummary: true, matchedLabel: descClean };
    }
  }

  return { isSummary: false, matchedLabel: '' };
}

/**
 * Boolean wrapper used during inspection sampling.
 * @param {any[]} row
 * @param {number | null} dateColIndex
 * @param {number | null} descColIndex
 * @returns {boolean}
 */
function detectSummaryRow(row, dateColIndex, descColIndex) {
  return detectSummaryRowInternal(row, dateColIndex, descColIndex).isSummary;
}

/**
 * Loads an Excel (`.xlsx`, `.xls`) or `.csv` workbook from an ArrayBuffer or Uint8Array.
 *
 * @param {ArrayBuffer | Uint8Array} arrayBuffer
 * @returns {import('../../vendor/xlsx.mjs').WorkBook}
 */
export function loadWorkbookFromBuffer(arrayBuffer) {
  if (!arrayBuffer) {
    throw new Error('Không có dữ liệu file để đọc.');
  }
  const bytes =
    arrayBuffer instanceof Uint8Array ? arrayBuffer : new Uint8Array(arrayBuffer);
  if (bytes.byteLength === 0) {
    throw new Error('File tải lên bị rỗng (0 byte).');
  }

  return XLSX.read(bytes, {
    type: 'array',
    cellDates: false,
    raw: true,
    codepage: 65001,
  });
}

/**
 * Returns the list of sheet names in a workbook.
 *
 * @param {import('../../vendor/xlsx.mjs').WorkBook} workbook
 * @returns {string[]}
 */
export function getSheetNames(workbook) {
  if (!workbook || !Array.isArray(workbook.SheetNames)) {
    return [];
  }
  return [...workbook.SheetNames];
}

/**
 * Extracts a 0-based 2D `rawGrid` from a SheetJS worksheet while preserving exact
 * 1-based Excel row numbers and column indices (even if top rows or Column A are blank),
 * without allocating millions of empty rows when Excel formatted whole columns.
 *
 * @param {any} worksheet
 * @returns {any[][]}
 */
function worksheetToExactGrid(worksheet) {
  if (!worksheet || !worksheet['!ref']) return [];

  let maxRow = -1;
  let maxCol = -1;

  if (Array.isArray(worksheet['!data'])) {
    const dense = worksheet['!data'];
    for (let r = 0; r < dense.length; r++) {
      const row = dense[r];
      if (!Array.isArray(row)) continue;
      for (let c = 0; c < row.length; c++) {
        const cell = row[c];
        if (cell && cell.v !== undefined && cell.v !== '') {
          if (r > maxRow) maxRow = r;
          if (c > maxCol) maxCol = c;
        }
      }
    }
  } else {
    for (const key of Object.keys(worksheet)) {
      if (key.charCodeAt(0) === 33) continue; // skip '!' metadata keys
      const cell = worksheet[key];
      if (!cell || (cell.v === undefined && cell.w === undefined) || cell.v === '') continue;
      const addr = XLSX.utils.decode_cell(key);
      if (addr.r > maxRow) maxRow = addr.r;
      if (addr.c > maxCol) maxCol = addr.c;
    }
  }

  if (maxRow < 0 || maxCol < 0) return [];

  const boundedRange = {
    s: { r: 0, c: 0 },
    e: { r: maxRow, c: maxCol },
  };

  return XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    raw: true,
    defval: '',
    blankrows: true,
    range: boundedRange,
  });
}

/**
 * Inspects a specific sheet in a SheetJS workbook and returns header, column & mapping metadata.
 *
 * @param {any} workbook
 * @param {string} sheetName
 * @param {number | null} [customHeaderRowIndex]
 */
export function inspectSheet(workbook, sheetName, customHeaderRowIndex = null) {
  const names = getSheetNames(workbook);
  const resolvedName =
    sheetName && workbook?.Sheets?.[sheetName]
      ? sheetName
      : names[0] || sheetName || 'Sheet1';
  const worksheet = workbook?.Sheets?.[resolvedName];
  const rawGrid = worksheet ? worksheetToExactGrid(worksheet) : [];
  return inspectRawGrid(rawGrid, resolvedName, customHeaderRowIndex);
}

/**
 * Inspects a 2D `rawGrid` directly, auto-detects the accounting table header row,
 * extracts column samples, and recommends field mappings.
 *
 * @param {any[][]} rawGrid
 * @param {string} [sheetName]
 * @param {number | null} [customHeaderRowIndex]
 */
export function inspectRawGrid(rawGrid, sheetName = 'Sheet1', customHeaderRowIndex = null) {
  const grid = Array.isArray(rawGrid) ? rawGrid : [];
  const rawRowsCount = grid.length;

  let detectedHeaderRowIndex = 0;
  if (
    customHeaderRowIndex !== null &&
    customHeaderRowIndex !== undefined &&
    Number.isInteger(Number(customHeaderRowIndex)) &&
    Number(customHeaderRowIndex) >= 0 &&
    Number(customHeaderRowIndex) < Math.max(rawRowsCount, 1)
  ) {
    detectedHeaderRowIndex = Number(customHeaderRowIndex);
  } else {
    detectedHeaderRowIndex = detectHeaderRowIndex(grid);
  }

  const columns = buildColumnsMetadata(grid, detectedHeaderRowIndex);
  const colCount = columns.length;

  const previewLimit = Math.min(Math.max(25, detectedHeaderRowIndex + 6), rawRowsCount);
  const previewTopRows = [];
  for (let r = 0; r < previewLimit; r++) {
    const row = Array.isArray(grid[r]) ? grid[r] : [];
    const cells = [];
    for (let c = 0; c < colCount; c++) {
      cells.push(cellToCleanString(row[c]));
    }
    previewTopRows.push({
      rowIndex: r,
      excelRowNumber: r + 1,
      cells,
    });
  }

  const suggestedMapping = inferSuggestedMapping(grid, detectedHeaderRowIndex, columns);

  return {
    sheetName: sheetName || 'Sheet1',
    rawRowsCount,
    detectedHeaderRowIndex,
    previewTopRows,
    columns,
    suggestedMapping,
    rawGrid: grid,
  };
}

/**
 * Extracts normalized transaction items, summary row flags, date bounds, and monthly buckets
 * starting from `headerRowIndex + 1`.
 *
 * @param {any[][]} rawGrid
 * @param {number} headerRowIndex
 * @param {{
 *   dateColIndex?: number | null,
 *   amountColIndex?: number | null,
 *   debitColIndex?: number | null,
 *   creditColIndex?: number | null,
 *   amountMode?: 'single' | 'debit_minus_credit',
 *   descColIndex?: number | null,
 *   refColIndex?: number | null
 * }} mapping
 * @param {Array<{ index: number, excelCol?: string, label: string }> | null} [columns]
 */
export function extractTransactions(rawGrid, headerRowIndex, mapping, columns = null) {
  const grid = Array.isArray(rawGrid) ? rawGrid : [];
  const startRow = Math.max(0, Number(headerRowIndex || 0) + 1);

  const dateColIndex = mapping?.dateColIndex ?? null;
  const amountColIndex = mapping?.amountColIndex ?? null;
  const debitColIndex = mapping?.debitColIndex ?? null;
  const creditColIndex = mapping?.creditColIndex ?? null;
  const descColIndex = mapping?.descColIndex ?? null;
  const refColIndex = mapping?.refColIndex ?? null;
  const useDebitCredit =
    mapping?.amountMode === 'debit_minus_credit' ||
    (amountColIndex === null && (debitColIndex !== null || creditColIndex !== null));

  // Precompute extra column descriptors (columns not directly mapped)
  const mappedIndices = new Set();
  if (dateColIndex !== null) mappedIndices.add(dateColIndex);
  if (useDebitCredit) {
    if (debitColIndex !== null) mappedIndices.add(debitColIndex);
    if (creditColIndex !== null) mappedIndices.add(creditColIndex);
  } else if (amountColIndex !== null) {
    mappedIndices.add(amountColIndex);
  }
  if (descColIndex !== null) mappedIndices.add(descColIndex);
  if (refColIndex !== null) mappedIndices.add(refColIndex);

  const resolvedCols = Array.isArray(columns)
    ? columns
    : buildColumnsMetadata(grid, Math.max(0, Number(headerRowIndex || 0)));

  const extraColDefs = [];
  const seenLabels = new Set();
  for (const col of resolvedCols) {
    if (mappedIndices.has(col.index)) continue;
    let key = col.label || `Cột ${col.excelCol || XLSX.utils.encode_col(col.index)}`;
    if (seenLabels.has(key)) {
      key = `${key} (${col.excelCol || XLSX.utils.encode_col(col.index)})`;
    }
    seenLabels.add(key);
    extraColDefs.push({ index: col.index, label: key });
  }

  const items = [];
  let summaryRowCount = 0;
  let skippedEmptyOrInvalidCount = 0;
  let minDateIso = null;
  let maxDateIso = null;
  const monthBuckets = new Map();

  for (let rowIndex = startRow; rowIndex < grid.length; rowIndex++) {
    const row = grid[rowIndex];
    if (!Array.isArray(row)) {
      skippedEmptyOrInvalidCount++;
      continue;
    }

    // Check completely blank row
    let hasAnyNonEmpty = false;
    for (let c = 0; c < row.length; c++) {
      if (cellToCleanString(row[c]) !== '') {
        hasAnyNonEmpty = true;
        break;
      }
    }
    if (!hasAnyNonEmpty) {
      skippedEmptyOrInvalidCount++;
      continue;
    }

    // Skip accounting column-numbering row (e.g. A | B | C | 1 | 2) right under header
    if (rowIndex <= startRow + 1 && isColumnIndexHelperRow(row)) {
      skippedEmptyOrInvalidCount++;
      continue;
    }

    // Parse transaction amount
    let amount = null;
    if (useDebitCredit) {
      const debitVal = debitColIndex !== null ? parseVietnameseAmount(row[debitColIndex]) : null;
      const creditVal = creditColIndex !== null ? parseVietnameseAmount(row[creditColIndex]) : null;
      if (debitVal === null && creditVal === null) {
        skippedEmptyOrInvalidCount++;
        continue;
      }
      const rawDiff = (debitVal ?? 0) - (creditVal ?? 0);
      amount = Math.round((rawDiff + Number.EPSILON) * 100) / 100;
    } else {
      if (amountColIndex === null) {
        skippedEmptyOrInvalidCount++;
        continue;
      }
      amount = parseVietnameseAmount(row[amountColIndex]);
      if (amount === null) {
        skippedEmptyOrInvalidCount++;
        continue;
      }
    }

    if (amount === 0) amount = 0; // normalize -0

    const parsedDate = dateColIndex !== null ? parseVietnameseDate(row[dateColIndex]) : null;
    let description = descColIndex !== null ? cellToCleanString(row[descColIndex]) : '';
    const reference = refColIndex !== null ? cellToCleanString(row[refColIndex]) : '';

    const summaryInfo = detectSummaryRowInternal(row, dateColIndex, descColIndex);
    const isSummaryRow = summaryInfo.isSummary;
    if (isSummaryRow && !description && summaryInfo.matchedLabel) {
      description = summaryInfo.matchedLabel;
    }

    // Skip zero-amount formula placeholder rows that have no date, description, or reference
    if (!isSummaryRow && amount === 0 && !parsedDate && !description && !reference) {
      skippedEmptyOrInvalidCount++;
      continue;
    }

    if (isSummaryRow) {
      summaryRowCount++;
    }

    const extraCells = {};
    for (const def of extraColDefs) {
      extraCells[def.label] = cellToCleanString(row[def.index]);
    }

    const excelRowNumber = rowIndex + 1;
    const dateIso = parsedDate ? parsedDate.iso : null;
    const dateDisplay = parsedDate ? parsedDate.display : '';
    const timestamp = parsedDate ? parsedDate.timestamp : null;

    items.push({
      id: `row-${excelRowNumber}`,
      excelRowNumber,
      dateIso,
      dateDisplay,
      timestamp,
      amount,
      description,
      reference,
      isSummaryRow,
      extraCells,
      rawCells: Array.from(row),
    });

    if (!isSummaryRow && dateIso) {
      if (minDateIso === null || dateIso < minDateIso) minDateIso = dateIso;
      if (maxDateIso === null || dateIso > maxDateIso) maxDateIso = dateIso;

      const monthKey = dateIso.slice(0, 7); // "YYYY-MM"
      let bucket = monthBuckets.get(monthKey);
      if (!bucket) {
        const [yyyyStr, mmStr] = monthKey.split('-');
        const yearNum = Number(yyyyStr);
        const monthNum = Number(mmStr);
        const lastDay = new Date(Date.UTC(yearNum, monthNum, 0)).getUTCDate();
        bucket = {
          key: monthKey,
          label: `Tháng ${mmStr}/${yyyyStr}`,
          fromIso: `${monthKey}-01`,
          toIso: `${monthKey}-${String(lastDay).padStart(2, '0')}`,
          count: 0,
        };
        monthBuckets.set(monthKey, bucket);
      }
      bucket.count++;
    }
  }

  const availableMonths = Array.from(monthBuckets.values()).sort((a, b) =>
    a.key.localeCompare(b.key)
  );

  return {
    items,
    summaryRowCount,
    skippedEmptyOrInvalidCount,
    minDateIso,
    maxDateIso,
    availableMonths,
  };
}

/**
 * Computes comfortable Excel column widths (`!cols`) from an AoA table.
 * @param {any[][]} aoa
 * @param {number[]} [minWidths]
 * @param {number} [maxWidth]
 * @returns {Array<{ wch: number }>}
 */
function computeAutoColumnWidths(aoa, minWidths = [], maxWidth = 52) {
  let colCount = minWidths.length;
  for (const row of aoa) {
    if (Array.isArray(row) && row.length > colCount) colCount = row.length;
  }

  const widths = new Array(colCount).fill(10);
  for (let c = 0; c < minWidths.length; c++) {
    if (minWidths[c] > widths[c]) widths[c] = minWidths[c];
  }

  for (let r = 0; r < aoa.length; r++) {
    const row = aoa[r];
    if (!Array.isArray(row)) continue;
    for (let c = 0; c < row.length; c++) {
      const val = row[c];
      if (val === null || val === undefined || val === '') continue;
      let len = 0;
      if (typeof val === 'number') {
        // Approximate formatted length with thousand separators
        len = Math.abs(val).toLocaleString('en-US').length + (val < 0 ? 2 : 1);
      } else {
        len = String(val).length;
      }
      const padded = Math.min(maxWidth, len + 3);
      if (padded > widths[c]) widths[c] = padded;
    }
  }

  return widths.map((wch) => ({ wch }));
}

/**
 * Applies Excel thousand-separator number formatting (`#,##0` or `#,##0.##`) to numeric cells
 * in specified column indices of a SheetJS worksheet.
 *
 * @param {any} ws
 * @param {number} startRowIdx
 * @param {number} endRowIdx
 * @param {number[]} numericColIndices
 */
function applyNumericFormatToColumns(ws, startRowIdx, endRowIdx, numericColIndices) {
  for (let r = startRowIdx; r <= endRowIdx; r++) {
    for (const c of numericColIndices) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      if (cell && cell.t === 'n' && Number.isFinite(cell.v)) {
        cell.z = Number.isInteger(cell.v) ? '#,##0' : '#,##0.##';
      }
    }
  }
}

/**
 * Normalizes a batch entry from `reconciledBatches` regardless of minor UI property variations.
 *
 * @param {any} batch
 * @param {number} idx
 * @param {Map<string, any>} itemsById
 */
function normalizeBatchEntry(batch, idx, itemsById) {
  const defaultCode = `DOT-${String(idx + 1).padStart(2, '0')}`;
  const rawCode =
    batch?.code ||
    batch?.batchCode ||
    batch?.name ||
    batch?.label ||
    (typeof batch?.id === 'string' && batch.id.length <= 20 ? batch.id : defaultCode);
  const code = String(rawCode).trim() || defaultCode;

  let resolvedItems = [];
  if (Array.isArray(batch?.items) && batch.items.length > 0) {
    resolvedItems = batch.items.map((it) => (typeof it === 'string' ? itemsById.get(it) : it)).filter(Boolean);
  } else if (Array.isArray(batch?.solution?.items) && batch.solution.items.length > 0) {
    resolvedItems = batch.solution.items
      .map((it) => (typeof it === 'string' ? itemsById.get(it) : it))
      .filter(Boolean);
  } else {
    const rawIds = batch?.itemIds || batch?.solution?.itemIds || [];
    if (Array.isArray(rawIds)) {
      resolvedItems = rawIds.map((id) => itemsById.get(id)).filter(Boolean);
    }
  }

  const computedTotal =
    Math.round(
      (resolvedItems.reduce((acc, it) => acc + (Number(it?.amount) || 0), 0) + Number.EPSILON) * 100
    ) / 100;

  const totalAmount = Number.isFinite(Number(batch?.totalAmount))
    ? Number(batch.totalAmount)
    : Number.isFinite(Number(batch?.solution?.totalAmount))
      ? Number(batch.solution.totalAmount)
      : computedTotal;

  const targetAmount = Number.isFinite(Number(batch?.targetAmount))
    ? Number(batch.targetAmount)
    : Number.isFinite(Number(batch?.target))
      ? Number(batch.target)
      : Number.isFinite(Number(batch?.solution?.targetAmount))
        ? Number(batch.solution.targetAmount)
        : totalAmount;

  const diff = Number.isFinite(Number(batch?.diff))
    ? Number(batch.diff)
    : Number.isFinite(Number(batch?.solution?.diff))
      ? Number(batch.solution.diff)
      : Math.round((totalAmount - targetAmount + Number.EPSILON) * 100) / 100;

  const note = cellToCleanString(batch?.note ?? batch?.notes ?? batch?.keyword ?? '');

  // Date range summary of items in batch
  let minIso = null;
  let maxIso = null;
  for (const it of resolvedItems) {
    if (it?.dateIso) {
      if (!minIso || it.dateIso < minIso) minIso = it.dateIso;
      if (!maxIso || it.dateIso > maxIso) maxIso = it.dateIso;
    }
  }
  const dateRangeDisplay =
    minIso && maxIso
      ? minIso === maxIso
        ? formatDateVN(minIso)
        : `${formatDateVN(minIso)} - ${formatDateVN(maxIso)}`
      : '';

  return {
    code,
    items: resolvedItems,
    itemCount: resolvedItems.length,
    targetAmount,
    totalAmount,
    diff,
    note,
    dateRangeDisplay,
  };
}

/**
 * Exports a clean, accountant-ready `.xlsx` workbook (as `Uint8Array`) containing:
 * - Sheet 1 (`Chi_Tiet_Doi_Soat`): Full or unallocated-only transaction lines with reconciliation status & original extra columns.
 * - Sheet 2 (`Tong_Hop_Cac_Dot_Gom`): Batch summary table + itemized lines per reconciled batch.
 *
 * @param {{
 *   sheetName?: string,
 *   columns?: Array<{ index: number, excelCol?: string, label: string }>,
 *   items?: any[],
 *   reconciledBatches?: any[],
 *   unallocatedOnly?: boolean
 * }} params
 * @returns {Uint8Array}
 */
export function exportReconciledWorkbook({
  sheetName = 'Sheet1',
  columns = [],
  items = [],
  reconciledBatches = [],
  unallocatedOnly = false,
} = {}) {
  const safeItems = Array.isArray(items) ? items : [];
  const safeBatches = Array.isArray(reconciledBatches) ? reconciledBatches : [];

  const itemsById = new Map();
  for (const item of safeItems) {
    if (item && item.id !== undefined) {
      itemsById.set(item.id, item);
    }
  }

  const normalizedBatches = safeBatches.map((b, idx) => normalizeBatchEntry(b, idx, itemsById));

  // Map each reconciled item ID -> its batch metadata
  const batchByItemId = new Map();
  for (const b of normalizedBatches) {
    for (const it of b.items) {
      if (it && it.id !== undefined && !batchByItemId.has(it.id)) {
        batchByItemId.set(it.id, b);
      }
    }
  }

  // Collect extra original column keys in stable order
  const extraColKeys = [];
  const seenExtraKeys = new Set();
  for (const item of safeItems) {
    if (item?.extraCells && typeof item.extraCells === 'object' && !Array.isArray(item.extraCells)) {
      for (const key of Object.keys(item.extraCells)) {
        if (!seenExtraKeys.has(key)) {
          seenExtraKeys.add(key);
          extraColKeys.push(key);
        }
      }
    }
  }

  // Fallback to `columns` + `rawCells` if synthetic items didn't populate `extraCells`
  const useRawColumnsFallback =
    extraColKeys.length === 0 && Array.isArray(columns) && columns.length > 0;
  const fallbackCols = useRawColumnsFallback ? columns : [];

  // ==========================================================================
  // SHEET 1: Chi_Tiet_Doi_Soat
  // ==========================================================================
  const sheet1Header = [
    'Dòng Excel gốc',
    'Ngày',
    'Số chứng từ',
    'Diễn giải',
    'Số tiền',
    'Trạng thái đối soát',
    'Mã đợt gom',
    'Mục tiêu đợt',
    'Chênh lệch đợt',
    ...(useRawColumnsFallback
      ? fallbackCols.map((c) => c.label || `Cột ${c.excelCol || c.index + 1}`)
      : extraColKeys),
  ];

  const sheet1Aoa = [sheet1Header];

  for (const item of safeItems) {
    if (!item) continue;
    const batchInfo = batchByItemId.get(item.id) || null;

    if (unallocatedOnly) {
      if (batchInfo || item.isSummaryRow) continue;
    }

    const statusLabel = batchInfo
      ? 'Đã gom'
      : item.isSummaryRow
        ? 'Dòng tổng cộng'
        : 'Chưa gom';

    const dateText = item.dateDisplay || (item.dateIso ? formatDateVN(item.dateIso) : '');

    const extraValues = useRawColumnsFallback
      ? fallbackCols.map((c) =>
          Array.isArray(item.rawCells) ? cellToCleanString(item.rawCells[c.index]) : ''
        )
      : extraColKeys.map((k) =>
          item.extraCells && item.extraCells[k] !== undefined
            ? cellToCleanString(item.extraCells[k])
            : ''
        );

    sheet1Aoa.push([
      item.excelRowNumber ?? '',
      dateText,
      item.reference ?? '',
      item.description ?? '',
      Number.isFinite(Number(item.amount)) ? Number(item.amount) : 0,
      statusLabel,
      batchInfo ? batchInfo.code : '',
      batchInfo ? batchInfo.targetAmount : '',
      batchInfo ? batchInfo.diff : '',
      ...extraValues,
    ]);
  }

  const ws1 = XLSX.utils.aoa_to_sheet(sheet1Aoa);
  applyNumericFormatToColumns(ws1, 1, sheet1Aoa.length - 1, [4, 7, 8]);
  ws1['!cols'] = computeAutoColumnWidths(sheet1Aoa, [14, 13, 16, 36, 18, 18, 14, 18, 16]);

  // ==========================================================================
  // SHEET 2: Tong_Hop_Cac_Dot_Gom
  // ==========================================================================
  const totalReconciledItems = normalizedBatches.reduce((s, b) => s + b.itemCount, 0);
  const sumTarget =
    Math.round(
      (normalizedBatches.reduce((s, b) => s + b.targetAmount, 0) + Number.EPSILON) * 100
    ) / 100;
  const sumActual =
    Math.round(
      (normalizedBatches.reduce((s, b) => s + b.totalAmount, 0) + Number.EPSILON) * 100
    ) / 100;
  const sumDiff =
    Math.round((normalizedBatches.reduce((s, b) => s + b.diff, 0) + Number.EPSILON) * 100) / 100;

  const sheet2Aoa = [
    [`BÁO CÁO TỔNG HỢP CÁC ĐỢT GOM ĐỐI SOÁT - SHEET: ${sheetName || 'Sheet1'}`],
    [
      'Số đợt đã gom:',
      normalizedBatches.length,
      'Tổng số dòng đã gom:',
      totalReconciledItems,
      'Tổng tiền thực gom:',
      sumActual,
    ],
    [],
    ['I. BẢNG TỔNG HỢP TỪNG ĐỢT GOM'],
    [
      'STT',
      'Mã đợt gom',
      'Số lượng dòng',
      'Mục tiêu cần gom',
      'Tổng tiền thực gom',
      'Chênh lệch',
      'Khoảng ngày chứng từ',
      'Ghi chú',
    ],
  ];

  const summaryStartRow = sheet2Aoa.length;
  if (normalizedBatches.length === 0) {
    sheet2Aoa.push(['', 'Chưa có đợt gom nào', 0, 0, 0, 0, '', '']);
  } else {
    normalizedBatches.forEach((b, idx) => {
      sheet2Aoa.push([
        idx + 1,
        b.code,
        b.itemCount,
        b.targetAmount,
        b.totalAmount,
        b.diff,
        b.dateRangeDisplay,
        b.note,
      ]);
    });
    sheet2Aoa.push([
      '',
      'Tổng cộng',
      totalReconciledItems,
      sumTarget,
      sumActual,
      sumDiff,
      '',
      '',
    ]);
  }
  const summaryEndRow = sheet2Aoa.length - 1;

  sheet2Aoa.push([]);
  sheet2Aoa.push(['II. DANH SÁCH CHI TIẾT CÁC DÒNG THUỘC TỪNG ĐỢT GOM']);
  sheet2Aoa.push([
    'Mã đợt gom',
    'STT trong đợt',
    'Dòng Excel gốc',
    'Ngày',
    'Số chứng từ',
    'Diễn giải',
    'Số tiền',
    'Mục tiêu đợt',
    'Tổng tiền đợt',
    'Chênh lệch đợt',
  ]);

  const detailStartRow = sheet2Aoa.length;
  for (const b of normalizedBatches) {
    b.items.forEach((it, itemIdx) => {
      const dateText = it.dateDisplay || (it.dateIso ? formatDateVN(it.dateIso) : '');
      sheet2Aoa.push([
        b.code,
        itemIdx + 1,
        it.excelRowNumber ?? '',
        dateText,
        it.reference ?? '',
        it.description ?? '',
        Number.isFinite(Number(it.amount)) ? Number(it.amount) : 0,
        itemIdx === 0 ? b.targetAmount : '',
        itemIdx === 0 ? b.totalAmount : '',
        itemIdx === 0 ? b.diff : '',
      ]);
    });
  }
  const detailEndRow = sheet2Aoa.length - 1;

  const ws2 = XLSX.utils.aoa_to_sheet(sheet2Aoa);
  applyNumericFormatToColumns(ws2, 1, 1, [5]);
  applyNumericFormatToColumns(ws2, summaryStartRow, summaryEndRow, [2, 3, 4, 5]);
  if (detailEndRow >= detailStartRow) {
    applyNumericFormatToColumns(ws2, detailStartRow, detailEndRow, [6, 7, 8, 9]);
  }
  ws2['!cols'] = computeAutoColumnWidths(sheet2Aoa, [14, 16, 15, 18, 18, 36, 22, 18, 18, 16]);

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws1, 'Chi_Tiet_Doi_Soat');
  XLSX.utils.book_append_sheet(wb, ws2, 'Tong_Hop_Cac_Dot_Gom');

  const rawOut = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  return rawOut instanceof Uint8Array ? rawOut : new Uint8Array(rawOut);
}
