/**
 * Vietnamese Accounting Number, Currency, Date & Words Formatting Utilities
 * Pure ES Module compatible with Browser (GitHub Pages) and Node.js ESM.
 */

const EXCEL_EPOCH_UTC_MS = Date.UTC(1899, 11, 30); // 1899-12-30T00:00:00.000Z
const MS_PER_DAY = 86_400_000;
// Valid Excel serial range: 1990-01-01 (32874) .. 2100-12-31 (73415)
const MIN_EXCEL_DATE_SERIAL = 32874;
const MAX_EXCEL_DATE_SERIAL = 73415;

const VN_DIGITS = ['không', 'một', 'hai', 'ba', 'bốn', 'năm', 'sáu', 'bảy', 'tám', 'chín'];
const VN_SCALE_UNITS = ['', 'nghìn', 'triệu'];

const LEADING_CURRENCY_RE =
  /^(?:(?:vnđ|vnd|đồng|dong|usd|eur|jpy|gbp|aud|sgd|cny|krw)\.?(?=\s|[+\-−–([\d]|$)|[₫đĐ$€¥£₩]\.?)\s*/i;
const TRAILING_CURRENCY_RE =
  /\s*(?:(?<=[\s\d).\-−–])(?:vnđ|vnd|đồng|dong|usd|eur|jpy|gbp|aud|sgd|cny|krw)\.?|[₫đĐ$€¥£₩]\.?)$/i;

// Strict optional time/timezone suffix after a date (rejects arbitrary prose after a date)
const OPTIONAL_TIME_SUFFIX_SOURCE =
  '(?:(?:[T\\s,]+|\\s+-\\s+)\\d{1,2}:\\d{2}(?::\\d{2}(?:[.,]\\d{1,6})?)?(?:\\s*(?:AM|PM|SA|CH))?(?:\\s*(?:Z|UTC|GMT|[+\\-]\\d{2}:?\\d{2}))?)?';

const YMD_DATE_RE = new RegExp(
  `^(\\d{4})([/.-])(\\d{1,2})\\2(\\d{1,2})${OPTIONAL_TIME_SUFFIX_SOURCE}$`,
  'i'
);
const DMY_DATE_RE = new RegExp(
  `^(\\d{1,2})([/.-])(\\d{1,2})\\2(\\d{2}|\\d{4})${OPTIONAL_TIME_SUFFIX_SOURCE}$`,
  'i'
);

/**
 * Normalizes Unicode (including Unikey NFD & combining stroke D/d) and exotic spaces.
 * @param {string} str
 * @returns {string}
 */
function normalizeUnicodeInput(str) {
  return str
    .replace(/[dD]\u0335/g, (m) => (m[0] === 'D' ? 'Đ' : 'đ'))
    .normalize('NFC')
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/[\s\u00A0\u202F\u2009]+/g, ' ')
    .trim();
}

/**
 * Strip leading/trailing currency symbols or ISO codes without touching normal words.
 * @param {string} text
 * @returns {string}
 */
function stripCurrencyAffixes(text) {
  let s = text.trim();
  for (let i = 0; i < 2; i++) {
    const prev = s;
    s = s.replace(LEADING_CURRENCY_RE, '').replace(TRAILING_CURRENCY_RE, '').trim();
    if (s === prev) break;
  }
  return s;
}

/**
 * Parses numbers and Vietnamese/international accounting formatted amount strings.
 * Examples:
 * - 12500000 -> 12500000
 * - "12.500.000" -> 12500000
 * - "12,500,000" -> 12500000
 * - "(500.000)" -> -500000
 * - "- 1.250.000 đ" -> -1250000
 * - "1.250.000,50" -> 1250000.5
 * - "1,250,000.50" -> 1250000.5
 *
 * @param {unknown} raw
 * @returns {number | null}
 */
export function parseVietnameseAmount(raw) {
  if (raw === null || raw === undefined || typeof raw === 'boolean') {
    return null;
  }

  if (typeof raw === 'number') {
    if (!Number.isFinite(raw)) return null;
    return raw === 0 ? 0 : raw;
  }

  if (typeof raw !== 'string') {
    return null;
  }

  let s = normalizeUnicodeInput(raw);
  if (!s || !/\d/.test(s)) {
    return null;
  }

  // Guard against voucher / partner codes starting with currency-like letters glued directly to a hyphen
  // e.g. "Đ-00125", "DONG-001", "VND-00125"
  if (/^(?:vnđ|vnd|đồng|dong|usd|eur|jpy|gbp|aud|sgd|cny|krw|[đĐ])[\-−–]\d/i.test(s)) {
    return null;
  }

  s = stripCurrencyAffixes(s);
  if (!s || !/\d/.test(s)) {
    return null;
  }

  // Scientific notation support (e.g. "1.25e7" from CSV export)
  if (/^[+\-]?\d+(?:\.\d+)?[eE][+\-]?\d+$/.test(s)) {
    const sciVal = Number(s);
    if (!Number.isFinite(sciVal)) return null;
    return sciVal === 0 ? 0 : sciVal;
  }

  let isNegative = false;
  let hasExplicitSign = false;

  // Leading sign (+ / - / unicode minus / en-dash)
  const leadSignMatch = s.match(/^([+\-−–])\s*/);
  if (leadSignMatch) {
    hasExplicitSign = true;
    if (leadSignMatch[1] !== '+') isNegative = true;
    s = s.slice(leadSignMatch[0].length).trim();
  }

  // Trailing minus sign (e.g. "500.000-" in SAP/ERP exports)
  if (!hasExplicitSign && /\s*[\-−–]$/.test(s) && !/^\d+[\-−–]\d+/.test(s)) {
    hasExplicitSign = true;
    isNegative = true;
    s = s.replace(/\s*[\-−–]$/, '').trim();
  }

  s = stripCurrencyAffixes(s);

  // Accounting parentheses: "(500.000)" or "( 1.250.000 đ )"
  if (s.startsWith('(') && s.endsWith(')')) {
    isNegative = true;
    s = s.slice(1, -1).trim();
    s = stripCurrencyAffixes(s);

    // Optional inner sign inside parens like "(-500.000)" or "(+500.000)"
    const innerSign = s.match(/^([+\-−–])\s*/);
    if (innerSign) {
      s = s.slice(innerSign[0].length).trim();
      s = stripCurrencyAffixes(s);
    }
  } else if (!hasExplicitSign) {
    // Handle case where sign appeared after leading currency, e.g. "VND -500.000" or "1.250.000- VND"
    const postCurrLeadSign = s.match(/^([+\-−–])\s*/);
    if (postCurrLeadSign) {
      if (postCurrLeadSign[1] !== '+') isNegative = true;
      s = s.slice(postCurrLeadSign[0].length).trim();
      s = stripCurrencyAffixes(s);
    } else if (/\s*[\-−–]$/.test(s) && !/^\d+[\-−–]\d+/.test(s)) {
      isNegative = true;
      s = s.replace(/\s*[\-−–]$/, '').trim();
      s = stripCurrencyAffixes(s);
    }
  }

  if (!s || !/\d/.test(s)) {
    return null;
  }

  // Reject any remaining non-numeric characters (letters, slashes in dates, inner hyphens, extra parens)
  if (/[^0-9., ]/.test(s)) {
    return null;
  }

  // Space-grouped thousands (e.g. "12 500 000" or "12 500 000,50")
  if (s.includes(' ')) {
    const [spaceMain, spaceDec] = s.split(',');
    if (
      s.split(',').length <= 2 &&
      !s.includes('.') &&
      /^\d{1,3}(?: \d{3})+$/.test(spaceMain) &&
      spaceMain[0] !== '0' &&
      (spaceDec === undefined || /^\d+$/.test(spaceDec))
    ) {
      s = spaceDec !== undefined ? `${spaceMain.replace(/ /g, '')},${spaceDec}` : spaceMain.replace(/ /g, '');
    } else {
      return null;
    }
  }

  const hasDot = s.includes('.');
  const hasComma = s.includes(',');
  let normalized = '';

  if (hasDot && hasComma) {
    const lastDot = s.lastIndexOf('.');
    const lastComma = s.lastIndexOf(',');

    if (lastComma > lastDot) {
      // Vietnamese / European: "1.250.000,50" -> dot is thousand, comma is decimal
      const commaParts = s.split(',');
      if (commaParts.length !== 2) return null;
      const [intPart, fracPart] = commaParts;
      if (!intPart || !fracPart || !/^\d+$/.test(fracPart)) return null;
      if (/^0\d+$/.test(intPart)) return null;
      if (!/^\d{1,3}(?:\.\d{3})+$/.test(intPart) && !/^\d+$/.test(intPart)) return null;
      normalized = `${intPart.replace(/\./g, '')}.${fracPart}`;
    } else {
      // US / International: "1,250,000.50" -> comma is thousand, dot is decimal
      const dotParts = s.split('.');
      if (dotParts.length !== 2) return null;
      const [intPart, fracPart] = dotParts;
      if (!intPart || !fracPart || !/^\d+$/.test(fracPart)) return null;
      if (/^0\d+$/.test(intPart)) return null;
      if (!/^\d{1,3}(?:,\d{3})+$/.test(intPart) && !/^\d+$/.test(intPart)) return null;
      normalized = `${intPart.replace(/,/g, '')}.${fracPart}`;
    }
  } else if (hasDot) {
    const parts = s.split('.');
    if (parts.length > 2) {
      // Multiple dots MUST be valid thousand groupings (rejects dates like "15.03.2026")
      if (!/^\d{1,3}(?:\.\d{3})+$/.test(s) || parts[0][0] === '0') return null;
      normalized = parts.join('');
    } else {
      const [intPart, fracPart] = parts;
      if (fracPart === '') return null;
      if (/^0\d+$/.test(intPart)) return null;
      if (intPart === '') {
        normalized = `0.${fracPart}`;
      } else if (fracPart.length === 3 && intPart.length >= 1 && intPart.length <= 3 && intPart[0] !== '0') {
        // Exactly 3 digits after dot with 1..3 non-zero-leading digits before dot -> thousand separator ("500.000" -> 500000)
        normalized = `${intPart}${fracPart}`;
      } else {
        // Decimal separator ("12.50" -> 12.5, "0.500" -> 0.5)
        normalized = `${intPart}.${fracPart}`;
      }
    }
  } else if (hasComma) {
    const parts = s.split(',');
    if (parts.length > 2) {
      if (!/^\d{1,3}(?:,\d{3})+$/.test(s) || parts[0][0] === '0') return null;
      normalized = parts.join('');
    } else {
      const [intPart, fracPart] = parts;
      if (fracPart === '') return null;
      if (/^0\d+$/.test(intPart)) return null;
      if (intPart === '') {
        normalized = `0.${fracPart}`;
      } else if (fracPart.length === 3 && intPart.length >= 1 && intPart.length <= 3 && intPart[0] !== '0') {
        // International single-thousand e.g. "500,000" -> 500000
        normalized = `${intPart}${fracPart}`;
      } else {
        // Vietnamese decimal comma e.g. "12,50" -> 12.5
        normalized = `${intPart}.${fracPart}`;
      }
    }
  } else {
    // Reject zero-padded multi-digit codes (e.g. "-00125", "001", tax IDs "0101234567")
    if (!/^\d+$/.test(s) || /^0\d+$/.test(s)) return null;
    normalized = s;
  }

  const num = Number(normalized);
  if (!Number.isFinite(num)) return null;
  if (num === 0) return 0;
  return isNegative ? -num : num;
}

/**
 * Formats a numeric value using Vietnamese accounting conventions:
 * dot (`.`) thousand separator and comma (`,`) decimal separator (up to 2 decimal places).
 *
 * @param {number | string | null | undefined} value
 * @param {{ showPlusSign?: boolean }} [options]
 * @returns {string}
 */
export function formatNumberVN(value, options = { showPlusSign: false }) {
  const num = typeof value === 'number' ? value : parseVietnameseAmount(value);
  if (num === null || !Number.isFinite(num)) {
    return '';
  }

  if (num === 0) {
    return '0';
  }

  // Exact fast-path for safe integers (avoids IEEE-754 * 100 precision overflow near 2^53)
  if (Number.isSafeInteger(num)) {
    const formattedInt = String(Math.abs(num)).replace(/\B(?=(\d{3})+(?!\d))/g, '.');
    const sign =
      num < 0
        ? '-'
        : options && options.showPlusSign && num > 0
          ? '+'
          : '';
    return `${sign}${formattedInt}`;
  }

  const absVal = Math.abs(num);
  const absRounded =
    absVal <= 9e13 ? Math.round((absVal + Number.EPSILON) * 100) / 100 : absVal;
  if (absRounded === 0) {
    return '0';
  }

  const [intPart, rawFrac] = absRounded.toFixed(2).split('.');
  const formattedInt = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  const cleanFrac = rawFrac.replace(/0+$/, '');

  const sign =
    num < 0
      ? '-'
      : options && options.showPlusSign && num > 0
        ? '+'
        : '';

  return cleanFrac ? `${sign}${formattedInt},${cleanFrac}` : `${sign}${formattedInt}`;
}

/**
 * Live formatting helper for web amount `<input>` fields (typing and pasting).
 * Supports leading `-`, trailing SAP `-`, or accounting `(...)` (with or without currency affixes)
 * for negative numbers, strips noise, and preserves smooth live typing with `.` thousand separators.
 *
 * @param {string | number | null | undefined} rawText
 * @returns {{ numeric: number | null, formatted: string }}
 */
export function formatInputAmountVN(rawText) {
  if (rawText === null || rawText === undefined) {
    return { numeric: null, formatted: '' };
  }

  if (typeof rawText === 'number') {
    if (!Number.isFinite(rawText)) return { numeric: null, formatted: '' };
    const cleanNum = rawText === 0 ? 0 : rawText;
    return { numeric: cleanNum, formatted: formatNumberVN(cleanNum) };
  }

  const normalizedInput = normalizeUnicodeInput(String(rawText));
  if (!normalizedInput) {
    return { numeric: null, formatted: '' };
  }

  let isNegative = false;
  let core = stripCurrencyAffixes(normalizedInput);

  // Leading sign before or after currency (e.g. "-1.250.000", "VND -2,450,000")
  if (/^[\-−–]/.test(normalizedInput) || /^[\-−–]/.test(core)) {
    isNegative = true;
    core = core.replace(/^[\-−–]\s*/, '').trim();
    core = stripCurrencyAffixes(core);
  } else if (/^\+/.test(core)) {
    core = core.replace(/^\+\s*/, '').trim();
    core = stripCurrencyAffixes(core);
  }

  // Trailing minus sign SAP/ERP style (e.g. "1.250.000-", "500,000-", "1.250.000- VND")
  if (/[\-−–]$/.test(core)) {
    isNegative = true;
    core = core.replace(/\s*[\-−–]$/, '').trim();
    core = stripCurrencyAffixes(core);
  }

  // Accounting parentheses (e.g. "(500.000)", "VNĐ (25,000,000)", "(500.000) đ", "(")
  if ((core.startsWith('(') && core.endsWith(')')) || core === '(') {
    isNegative = true;
    if (core.length > 1) {
      core = core.slice(1, -1).trim();
      core = stripCurrencyAffixes(core);
    }
  }

  // Keep only digits, dots, and commas to inspect numeric intent
  const cleaned = core.replace(/[^0-9.,]/g, '');

  if (!/\d/.test(cleaned)) {
    return isNegative ? { numeric: null, formatted: '-' } : { numeric: null, formatted: '' };
  }

  const hasDot = cleaned.includes('.');
  const hasComma = cleaned.includes(',');

  let intDigits = '';
  let decDigits = '';
  let hasTrailingComma = false;

  if (hasDot && hasComma) {
    if (cleaned.lastIndexOf(',') > cleaned.lastIndexOf('.')) {
      // VN format with comma decimal: "1.250.000,50" or live typing "1.250,"
      const lastComma = cleaned.lastIndexOf(',');
      intDigits = cleaned.slice(0, lastComma).replace(/[.,]/g, '');
      decDigits = cleaned.slice(lastComma + 1).replace(/[.,]/g, '').slice(0, 2);
      hasTrailingComma = decDigits.length === 0 && cleaned.endsWith(',');
    } else {
      // Pasted US format: "1,250,000.50"
      const lastDot = cleaned.lastIndexOf('.');
      intDigits = cleaned.slice(0, lastDot).replace(/[.,]/g, '');
      decDigits = cleaned.slice(lastDot + 1).replace(/[.,]/g, '').slice(0, 2).replace(/0+$/, '');
    }
  } else if (hasComma) {
    const parts = cleaned.split(',');
    if (
      parts.length > 2 ||
      (parts.length === 2 &&
        parts[1].length === 3 &&
        parts[0].length >= 1 &&
        parts[0].length <= 3 &&
        parts[0][0] !== '0')
    ) {
      // Pasted comma-thousand format: "12,500,000" or "500,000"
      intDigits = cleaned.replace(/,/g, '');
    } else {
      // VN decimal comma during typing/paste: "12500," or "12500,5"
      intDigits = parts[0];
      decDigits = parts.slice(1).join('').slice(0, 2);
      hasTrailingComma = decDigits.length === 0 && cleaned.endsWith(',');
    }
  } else if (hasDot) {
    const parts = cleaned.split('.');
    if (parts.length === 2 && parts[0].length >= 4 && parts[1].length >= 1 && parts[1].length <= 2) {
      // Pasted raw float with ungrouped integer >= 4 digits, e.g. "1250000.50"
      intDigits = parts[0];
      decDigits = parts[1].replace(/0+$/, '');
    } else {
      // VN thousand dots during live typing (e.g. "1.2500" -> "12.500", "12.50" -> "1.250")
      intDigits = cleaned.replace(/\./g, '');
    }
  } else {
    intDigits = cleaned;
  }

  // Strip leading zeros on integer part (keep single "0" if all zeros)
  intDigits = intDigits.replace(/^0+(?=\d)/, '');
  if (!intDigits) intDigits = '0';

  const numericStr = decDigits ? `${intDigits}.${decDigits}` : intDigits;
  const absNum = Number(numericStr);
  if (!Number.isFinite(absNum)) {
    return { numeric: null, formatted: '' };
  }

  const signedNum = absNum === 0 ? 0 : isNegative ? -absNum : absNum;
  const formattedInt = intDigits.replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  const signPrefix = isNegative && (absNum > 0 || hasTrailingComma || decDigits.length > 0) ? '-' : '';

  let formatted = `${signPrefix}${formattedInt}`;
  if (decDigits.length > 0) {
    formatted += `,${decDigits}`;
  } else if (hasTrailingComma) {
    formatted += ',';
  }

  return { numeric: signedNum, formatted };
}

/**
 * Validates calendar year, month (1..12), day (1..31) and constructs standard UTC result.
 * @param {number} year
 * @param {number} month
 * @param {number} day
 * @returns {{ iso: string, display: string, timestamp: number } | null}
 */
function buildValidDateResult(year, month, day) {
  if (
    !Number.isInteger(year) ||
    !Number.isInteger(month) ||
    !Number.isInteger(day) ||
    year < 1900 ||
    year > 2200 ||
    month < 1 ||
    month > 12 ||
    day < 1 ||
    day > 31
  ) {
    return null;
  }

  const timestamp = Date.UTC(year, month - 1, day);
  const check = new Date(timestamp);
  if (
    check.getUTCFullYear() !== year ||
    check.getUTCMonth() !== month - 1 ||
    check.getUTCDate() !== day
  ) {
    return null;
  }

  const yyyy = String(year).padStart(4, '0');
  const mm = String(month).padStart(2, '0');
  const dd = String(day).padStart(2, '0');

  return {
    iso: `${yyyy}-${mm}-${dd}`,
    display: `${dd}/${mm}/${yyyy}`,
    timestamp,
  };
}

/**
 * Resolves 2-digit or 4-digit year string (YY < 70 -> 20YY, >= 70 -> 19YY).
 * @param {string} yearStr
 * @returns {number}
 */
function resolveYearToken(yearStr) {
  const num = Number(yearStr);
  if (yearStr.length === 2) {
    return num < 70 ? 2000 + num : 1900 + num;
  }
  return num;
}

/**
 * Parses Vietnamese dates, ISO dates, Date objects, and Excel serial dates.
 * Ambiguous XX/YY/ZZZZ defaults strictly to DD/MM/YYYY (Vietnamese standard),
 * switching to MM/DD/YYYY only when first <= 12 and second > 12.
 *
 * @param {unknown} raw
 * @returns {{ iso: string, display: string, timestamp: number } | null}
 */
export function parseVietnameseDate(raw) {
  if (raw === null || raw === undefined || typeof raw === 'boolean') {
    return null;
  }

  if (raw instanceof Date) {
    const time = raw.getTime();
    if (Number.isNaN(time)) return null;
    if (
      raw.getUTCHours() === 0 &&
      raw.getUTCMinutes() === 0 &&
      raw.getUTCSeconds() === 0 &&
      raw.getUTCMilliseconds() === 0
    ) {
      return buildValidDateResult(raw.getUTCFullYear(), raw.getUTCMonth() + 1, raw.getUTCDate());
    }
    return buildValidDateResult(raw.getFullYear(), raw.getMonth() + 1, raw.getDate());
  }

  if (typeof raw === 'number') {
    if (!Number.isFinite(raw)) return null;
    const serialDays = Math.floor(raw + 1e-9);
    if (serialDays < MIN_EXCEL_DATE_SERIAL || serialDays > MAX_EXCEL_DATE_SERIAL) {
      return null;
    }
    const utcMs = EXCEL_EPOCH_UTC_MS + serialDays * MS_PER_DAY;
    const dt = new Date(utcMs);
    const year = dt.getUTCFullYear();
    if (year < 1990 || year > 2100) return null;
    return buildValidDateResult(year, dt.getUTCMonth() + 1, dt.getUTCDate());
  }

  if (typeof raw !== 'string') {
    return null;
  }

  let s = normalizeUnicodeInput(raw);
  if (!s) return null;

  // Vietnamese prose date: "Ngày 05 tháng 03 năm 2026"
  const vnProseMatch = s.match(/^(?:ngày\s*)?(\d{1,2})\s*tháng\s*(\d{1,2})\s*năm\s*(\d{2}|\d{4})$/i);
  if (vnProseMatch) {
    const day = Number(vnProseMatch[1]);
    const month = Number(vnProseMatch[2]);
    const year = resolveYearToken(vnProseMatch[3]);
    return buildValidDateResult(year, month, day);
  }

  // Strip optional leading "Ngày " prefix
  s = s.replace(/^ngày\s+/i, '');

  // 1. Year-first format: YYYY-MM-DD or YYYY/MM/DD or YYYY.MM.DD (with optional strict time suffix)
  const ymdMatch = s.match(YMD_DATE_RE);
  if (ymdMatch) {
    const year = Number(ymdMatch[1]);
    const month = Number(ymdMatch[3]);
    const day = Number(ymdMatch[4]);
    return buildValidDateResult(year, month, day);
  }

  // 2. Day-first (VN default) format: DD/MM/YYYY, D/M/YYYY, DD-MM-YYYY, DD.MM.YYYY, DD/MM/YY (with optional strict time suffix)
  const dmyMatch = s.match(DMY_DATE_RE);
  if (dmyMatch) {
    const first = Number(dmyMatch[1]);
    const second = Number(dmyMatch[3]);
    const year = resolveYearToken(dmyMatch[4]);
    const day = first <= 12 && second > 12 ? second : first;
    const month = first <= 12 && second > 12 ? first : second;
    return buildValidDateResult(year, month, day);
  }

  // 3. Leading time followed by DD/MM/YYYY (e.g. bank export "08:30:15 15/03/2026" or "02:35:22 CH 05/03/2026")
  const timeFirstMatch = s.match(
    /^\d{1,2}:\d{2}(?::\d{2}(?:[.,]\d{1,6})?)?(?:\s*(?:AM|PM|SA|CH))?\s+(\d{1,2})([/.-])(\d{1,2})\2(\d{2}|\d{4})$/i
  );
  if (timeFirstMatch) {
    const first = Number(timeFirstMatch[1]);
    const second = Number(timeFirstMatch[3]);
    const year = resolveYearToken(timeFirstMatch[4]);
    const day = first <= 12 && second > 12 ? second : first;
    const month = first <= 12 && second > 12 ? first : second;
    return buildValidDateResult(year, month, day);
  }

  return null;
}

/**
 * Formats an ISO string, Date, or Vietnamese date value to "DD/MM/YYYY" (or "" if invalid).
 * @param {unknown} isoOrDate
 * @returns {string}
 */
export function formatDateVN(isoOrDate) {
  const parsed = parseVietnameseDate(isoOrDate);
  return parsed ? parsed.display : '';
}

/**
 * Converts an ISO date string ("YYYY-MM-DD") or parseable date to UTC midnight timestamp.
 * @param {string | Date | number | null | undefined} iso
 * @returns {number | null}
 */
export function isoToUtcTimestamp(iso) {
  if (iso === null || iso === undefined || iso === '') {
    return null;
  }
  if (typeof iso === 'string') {
    const fastIso = iso.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (fastIso) {
      const res = buildValidDateResult(Number(fastIso[1]), Number(fastIso[2]), Number(fastIso[3]));
      return res ? res.timestamp : null;
    }
  }
  const parsed = parseVietnameseDate(iso);
  return parsed ? parsed.timestamp : null;
}

/**
 * Reads a 3-digit block (0..999) into Vietnamese words.
 * @param {number} block
 * @param {boolean} readFullHundred
 * @returns {string[]}
 */
function readThreeDigitBlock(block, readFullHundred) {
  const hundreds = Math.floor(block / 100);
  const tens = Math.floor((block % 100) / 10);
  const units = block % 10;
  const words = [];

  if (hundreds > 0 || readFullHundred) {
    words.push(VN_DIGITS[hundreds], 'trăm');
  }

  if (tens === 0) {
    if (units > 0) {
      if (hundreds > 0 || readFullHundred) {
        words.push('lẻ');
      }
      words.push(VN_DIGITS[units]);
    }
  } else if (tens === 1) {
    words.push('mười');
    if (units === 1) {
      words.push('một');
    } else if (units === 5) {
      words.push('lăm');
    } else if (units > 0) {
      words.push(VN_DIGITS[units]);
    }
  } else {
    words.push(VN_DIGITS[tens], 'mươi');
    if (units === 1) {
      words.push('mốt');
    } else if (units === 4) {
      words.push('tư');
    } else if (units === 5) {
      words.push('lăm');
    } else if (units > 0) {
      words.push(VN_DIGITS[units]);
    }
  }

  return words;
}

/**
 * Reads a non-negative integer string into Vietnamese words (supports > 1,000 tỷ).
 * @param {string} intStr
 * @returns {string[]}
 */
function readIntegerDigitsVN(intStr) {
  const clean = intStr.replace(/^0+/, '');
  if (!clean) return ['không'];

  const groups = [];
  for (let end = clean.length; end > 0; end -= 3) {
    const start = Math.max(0, end - 3);
    groups.push(Number(clean.slice(start, end)));
  }

  const words = [];
  let hasHigherNonZero = false;

  for (let pos = groups.length - 1; pos >= 0; pos--) {
    const val = groups[pos];
    const scaleMod = pos % 3;

    if (val > 0) {
      words.push(...readThreeDigitBlock(val, hasHigherNonZero));
      if (scaleMod > 0) {
        words.push(VN_SCALE_UNITS[scaleMod]);
      }
      hasHigherNonZero = true;
    }

    // Emit "tỷ" at every 10^9 boundary if this 9-digit billion tier had any non-zero block
    if (pos > 0 && scaleMod === 0) {
      const tierHasNonZero =
        val > 0 ||
        (pos + 1 < groups.length && groups[pos + 1] > 0) ||
        (pos + 2 < groups.length && groups[pos + 2] > 0);

      if (tierHasNonZero) {
        const billionCount = pos / 3;
        for (let b = 0; b < billionCount; b++) {
          words.push('tỷ');
        }
      }
    }
  }

  return words;
}

/**
 * Reads an amount in standard Vietnamese accounting words.
 * Examples:
 * - 0 -> "Không đồng"
 * - 12500000 -> "Mười hai triệu năm trăm nghìn đồng"
 * - -500000 -> "Âm năm trăm nghìn đồng"
 *
 * @param {number | string | null | undefined} amount
 * @returns {string}
 */
export function readAmountInVietnameseWords(amount) {
  const num = typeof amount === 'number' ? amount : parseVietnameseAmount(amount);
  if (num === null || !Number.isFinite(num)) {
    return '';
  }

  if (num === 0) {
    return 'Không đồng';
  }

  let intStr = '';
  let cleanFrac = '';

  // Exact fast-path for safe integers (prevents IEEE-754 * 100 precision loss on large integers)
  if (Number.isSafeInteger(num)) {
    intStr = String(Math.abs(num));
  } else {
    const absVal = Math.abs(num);
    const absRounded =
      absVal <= 9e13 ? Math.round((absVal + Number.EPSILON) * 100) / 100 : absVal;
    if (absRounded === 0) {
      return 'Không đồng';
    }
    const parts = absRounded.toFixed(2).split('.');
    intStr = parts[0];
    cleanFrac = parts[1].replace(/0+$/, '');
  }

  const words = [];
  if (num < 0) {
    words.push('âm');
  }

  words.push(...readIntegerDigitsVN(intStr));

  if (cleanFrac.length > 0) {
    words.push('phẩy');
    if (cleanFrac.length === 1) {
      words.push(VN_DIGITS[Number(cleanFrac)]);
    } else if (cleanFrac[0] === '0') {
      words.push('không', VN_DIGITS[Number(cleanFrac[1])]);
    } else {
      words.push(...readThreeDigitBlock(Number(cleanFrac), false));
    }
  }

  words.push('đồng');

  const sentence = words.join(' ');
  return sentence.charAt(0).toUpperCase() + sentence.slice(1);
}
