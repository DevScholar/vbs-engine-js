import { nlsNumberFormat, type NlsNumberFormat } from './nls-number-table.generated.ts';

/**
 * Writes numbers the way vbscript.dll writes them, from the Windows NLS values in
 * nls-number-table.generated.ts.
 *
 * Why this is not Intl.NumberFormat with a few characters swapped:
 *
 *   * Windows keeps two sets of marks, one for plain numbers and one for money, and they are not always
 *     the same. de-AT groups a plain number with U+00A0 and money with a full stop, so a rule that
 *     rewrote marks across the whole string would have to know which of the two it was looking at.
 *   * The 16 ways Windows writes a negative amount, and the 4 ways it places a currency symbol, are
 *     NLS values (LOCALE_INEGCURR, LOCALE_ICURRENCY), not a pattern string. Intl's nearest equivalents
 *     (signDisplay, currencySign, accounting) cover a different, smaller set.
 *   * Some locales have no ICU data at all. 1073 (Tsonga) is the only such LCID in this table: Intl
 *     resolves 'ts' to the host language, so its marks belong to another language entirely.
 *
 * Rounding is the one part Intl still does. It is asked for a plain grouped-less string in en-US, whose
 * digits are the same in every locale, and the answer is then re-marked from the NLS values. That keeps
 * ICU's correct decimal rounding (0.125 to 1,005 all round the way vbscript.dll rounds them, where
 * JavaScript's own toFixed does not) without taking any of its marks or patterns.
 */

/** Which of VBScript's three writers is asking. */
export type NlsNumberStyle = 'number' | 'currency' | 'percent';

/** VBScript's own arguments, before the locale fills in what the caller left out. */
export interface NlsNumberOptions {
  /** numDigitsAfterDecimal: a count, or -1 for the locale's own. */
  digits?: number;
  /** includeLeadingDigit: -1 keep, 0 drop, anything else for the locale's own. */
  leadingDigit?: number;
  /** useParensForNegativeNumbers: -1 parens, 0 a minus sign, anything else for the locale's own. */
  parens?: number;
  /** groupDigits: -1 group, 0 do not, anything else for the locale's own. */
  group?: number;
}

/** The neutral row used for the one LCID Windows has no name for (1084). */
const ASCII_FALLBACK: NlsNumberFormat = {
  decimal: '.',
  thousand: ',',
  grouping: '3;0',
  digits: 2,
  leadingZero: 1,
  negativeNumber: 1,
  negativeSign: '-',
  positiveSign: '',
  currencySymbol: '',
  currencyCode: '',
  currencyDigits: 2,
  currencyPosition: 0,
  negativeCurrency: 0,
  monDecimal: '.',
  monThousand: ',',
  monGrouping: '3;0',
  nativeDigits: '0123456789',
};

const roundedMagnitude = (value: number, digits: number): string =>
  new Intl.NumberFormat('en-US', {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits,
    useGrouping: false,
  }).format(Math.abs(value));

/**
 * The group sizes a Windows number uses, e.g. "3;0" -> [3, 3].
 *
 * Windows' LOCALE_SGROUPING can name two sizes, and the Indian locales do ("3;2;0", the lakh grouping).
 * VBScript's three number writers never use the second one: measured under 1081 (hi-IN) and 1102 (mr-IN),
 * FormatNumber(1234567.891) is "1,234,567.89" and not "12,34,567.89", and FormatCurrency groups the same
 * way. The size is therefore repeated rather than read, and the pattern only says whether to group at all.
 */
function groupSizes(pattern: string): number[] {
  const sizes = pattern
    .split(';')
    .map(part => Number(part.trim()))
    .filter(size => Number.isFinite(size));
  if (sizes.length === 0 || sizes[0] <= 0) return [];
  return [sizes[0], sizes[0]];
}

/** Inserts the group mark between the digits of a whole-number part. */
function groupInteger(integer: string, mark: string, pattern: string): string {
  const [primary, secondary] = groupSizes(pattern);
  if (!mark || primary === undefined) return integer;
  let grouped = '';
  let rest = integer;
  let width = primary;
  while (rest.length > width) {
    grouped = mark + rest.slice(-width) + grouped;
    rest = rest.slice(0, -width);
    width = secondary ?? width;
  }
  return rest + grouped;
}

/** The unsigned number as the locale writes it: grouped whole part, fraction mark, fraction. */
function writeMagnitude(
  value: number,
  digits: number,
  facts: NlsNumberFormat,
  useMoney: boolean,
  group: boolean
): string {
  const magnitude = roundedMagnitude(value, digits);
  const dot = magnitude.indexOf('.');
  let integer = dot < 0 ? magnitude : magnitude.slice(0, dot);
  const fraction = dot < 0 ? '' : magnitude.slice(dot + 1);
  if (integer === '0' && facts.leadingZero === 0) integer = '';
  const decimalMark = useMoney ? facts.monDecimal : facts.decimal;
  const groupMark = useMoney ? facts.monThousand : facts.thousand;
  const grouped = group
    ? groupInteger(integer, groupMark, useMoney ? facts.monGrouping : facts.grouping)
    : integer;
  return digits > 0 ? grouped + decimalMark + fraction : grouped;
}

/** The negative-sign placement LOCALE_INEGNUMBER picks: 0 "(1)", 1 "-1", 2 "- 1", 3 "1-", 4 "1 -". */
function negativeNumber(text: string, sign: string, pattern: number): string {
  switch (pattern) {
    case 0:
      return '(' + text + ')';
    case 2:
      return sign + ' ' + text;
    case 3:
      return text + sign;
    case 4:
      return text + ' ' + sign;
    default:
      return sign + text;
  }
}

/**
 * The sixteen LOCALE_INEGCURR shapes. 'n' is the amount, 's' the symbol, '-' the negative sign; the
 * spaces are the ones Windows writes.
 */
const NEGATIVE_CURRENCY: readonly string[] = [
  '(sn)',
  '-sn',
  's-n',
  'sn-',
  '(ns)',
  '-ns',
  'n-s',
  'ns-',
  '-n s',
  '-s n',
  's n-',
  'n s-',
  's -n',
  'n- s',
  '(s n)',
  '(n s)',
];

/** The four LOCALE_ICURRENCY placements: 0 "$1", 1 "1$", 2 "$ 1", 3 "1 $". */
const POSITIVE_CURRENCY: readonly string[] = ['sn', 'ns', 's n', 'n s'];

/** Fills one of the currency shapes above with the amount and the symbol. */
function fillCurrency(shape: string, text: string, symbol: string, sign: string): string {
  let written = '';
  for (const shapeCharacter of shape) {
    if (shapeCharacter === 'n') written += text;
    else if (shapeCharacter === 's') written += symbol;
    else if (shapeCharacter === '-') written += sign;
    else written += shapeCharacter;
  }
  return written;
}

/** One of the currency shapes above with its sign replaced by parentheses, or as it stands. */
function currencyShapeWithParens(shape: string): string {
  const withoutSign = shape.replace('-', '');
  return withoutSign.includes('(') ? withoutSign : '(' + withoutSign + ')';
}

/** One of the currency shapes above with parentheses replaced by a minus sign, or as it stands. */
function currencyShapeWithMinus(shape: string): string {
  const withoutParens = shape.replace(/[()]/g, '');
  return withoutParens.includes('-') ? withoutParens : '-' + withoutParens;
}

/** What a caller left out is taken from the locale, which is what -1 and -2 mean in VBScript. */
function localeDigits(style: NlsNumberStyle, facts: NlsNumberFormat, asked?: number): number {
  if (asked !== undefined && asked >= 0) return asked;
  if (style === 'currency') return facts.currencyDigits >= 0 ? facts.currencyDigits : facts.digits;
  return facts.digits >= 0 ? facts.digits : 2;
}

/**
 * Writes a number as vbscript.dll writes it under an LCID.
 *
 * @param value - The number to write.
 * @param lcid - The LCID whose Windows number format applies.
 * @param style - FormatNumber, FormatCurrency or FormatPercent.
 * @param options - VBScript's own arguments, with the sentinel values it uses for "the locale's".
 */
export function formatNlsNumber(
  value: number,
  lcid: number,
  style: NlsNumberStyle,
  options: NlsNumberOptions = {}
): string {
  const facts = nlsNumberFormat(lcid) ?? ASCII_FALLBACK;
  const money = style === 'currency';
  const scaled = style === 'percent' ? value * 100 : value;
  const digits = localeDigits(style, facts, options.digits);

  // The locale's leading-zero setting, unless the caller asked for one or the other.
  const withFacts: NlsNumberFormat =
    options.leadingDigit === -1
      ? { ...facts, leadingZero: 1 }
      : options.leadingDigit === 0
        ? { ...facts, leadingZero: 0 }
        : facts;

  const group = options.group === 0 ? false : options.group === -1 ? true : true;
  const text = writeMagnitude(scaled, digits, withFacts, money, group);
  const negative = scaled < 0 || Object.is(scaled, -0);
  const sign = facts.negativeSign || '-';

  if (money) {
    const symbol = facts.currencySymbol;
    if (!negative) {
      return fillCurrency(POSITIVE_CURRENCY[facts.currencyPosition] ?? 'sn', text, symbol, sign);
    }
    const own = NEGATIVE_CURRENCY[facts.negativeCurrency] ?? '-sn';
    // useParensForNegativeNumbers changes only the sign of the locale's own shape, measured under 1031 and
    // 1033: True puts the amount in parentheses ("(1.234,50 €)" and "(50,00%)"), False writes a minus
    // ("-$1,234.50" where the locale's own shape is "($1,234.50)"), and anything else leaves it alone.
    const shape =
      options.parens === -1
        ? currencyShapeWithParens(own)
        : options.parens === 0
          ? currencyShapeWithMinus(own)
          : own;
    return fillCurrency(shape, text, symbol, sign);
  }

  if (style === 'percent') {
    if (!negative) return text + '%';
    if (options.parens === -1) return '(' + text + '%)';
    if (options.parens === 0) return sign + text + '%';
    return negativeNumber(text, sign, facts.negativeNumber) + '%';
  }

  if (!negative) return text;
  const pattern = options.parens === -1 ? 0 : options.parens === 0 ? 1 : facts.negativeNumber;
  return negativeNumber(text, sign, pattern);
}

/**
 * Writes a number the way CStr writes one: the locale's fraction mark, no grouping, and no trailing
 * zeros beyond what the value itself needs.
 *
 * Each VBScript type carries its own precision, and the three were measured against a fresh cscript:
 *
 *   * A Double keeps fifteen significant digits (CStr(1/3) is "0,333333333333333", and CStr(0.1+0.2) is
 *     "0,3" where the value itself is 0.30000000000000004).
 *   * A Single keeps seven (CStr(CSng(1/3)) is "0,3333333", and CStr(CSng(1234567.89)) is "1234568").
 *   * A Currency keeps four decimal places rather than a count of significant digits, and it keeps the
 *     sign through a rounding to zero: CStr(CCur(0.00005)) is "0,0001" and CStr(CCur(-0.00005)) is
 *     "-0,0001", while CStr(CCur(0.00004)) is "0".
 *
 * Outside a window of exponents the value is written with one, and the exponent carries a sign and at
 * least two digits: a Double writes "1E+15" and "1E-16" plainly either side, and a Single writes
 * "1,234568E+07" and "4E-08" either side of its own narrower window. The window is the count of
 * significant digits on each side of the point, so it is fifteen for a Double and seven for a Single.
 */

/** How many significant digits each type keeps. */
const SIGNIFICANT_DIGITS: Record<'single' | 'double', number> = { single: 7, double: 15 };

/** How many decimal places a Currency keeps, which is what its scaled integer counts. */
export const CURRENCY_DIGITS = 4;

/** Writes a value with a fixed count of significant digits, in an exponent only outside the plain window. */
function significantText(value: number, significantDigits: number): string {
  if (!Number.isFinite(value)) return String(value);
  const magnitude = Math.abs(value);
  if (magnitude === 0) return '0';
  const precise = magnitude.toPrecision(significantDigits);
  const [rawMantissa, rawExponent] = precise.split('e');
  let digits = rawMantissa.replace('.', '');
  let leadingPower =
    rawExponent === undefined
      ? (rawMantissa.includes('.') ? rawMantissa.indexOf('.') : rawMantissa.length) - 1
      : Number(rawExponent);
  // A mantissa below one starts with zeros that are not significant: drop them and move the power with
  // them, so the first digit written is the first significant one.
  const leadingZeros = digits.length - digits.replace(/^0+/, '').length;
  digits = digits.slice(leadingZeros).replace(/0+$/, '');
  leadingPower -= leadingZeros;
  const sign = value < 0 ? '-' : '';
  if (leadingPower >= -significantDigits && leadingPower <= significantDigits - 1) {
    if (leadingPower < 0) return sign + '0.' + '0'.repeat(-leadingPower - 1) + digits;
    if (digits.length <= leadingPower + 1)
      return sign + digits + '0'.repeat(leadingPower + 1 - digits.length);
    return sign + digits.slice(0, leadingPower + 1) + '.' + digits.slice(leadingPower + 1);
  }
  const mantissa = digits.length === 1 ? digits : digits[0] + '.' + digits.slice(1);
  const exponent = (leadingPower < 0 ? '-' : '+') + String(Math.abs(leadingPower)).padStart(2, '0');
  return sign + mantissa + 'E' + exponent;
}

/**
 * Writes a number the way CStr writes one.
 *
 * @param value - The number to write.
 * @param lcid - The LCID whose fraction mark applies.
 * @param kind - Which floating point type is being written, since the two carry different precision.
 */
export function formatVbNumberText(
  value: number,
  lcid: number,
  kind: 'single' | 'double' = 'double'
): string {
  const facts = nlsNumberFormat(lcid) ?? ASCII_FALLBACK;
  const written = significantText(value, SIGNIFICANT_DIGITS[kind]);
  return facts.decimal === '.' ? written : written.replace('.', facts.decimal);
}

/**
 * Writes a Currency from the scaled integer it is held as, which is the only way to keep every digit.
 *
 * The amount is a count of ten-thousandths, so the text comes out of the integer itself: a Currency reaches
 * eighteen digits and a JavaScript number counts with sixteen, which is why the value is not a number.
 *
 * @param scaled - The amount as a count of ten-thousandths.
 * @param lcid - The LCID whose fraction mark applies.
 */
export function formatVbCurrencyText(scaled: bigint, lcid: number): string {
  const facts = nlsNumberFormat(lcid) ?? ASCII_FALLBACK;
  const digits = (scaled < 0n ? -scaled : scaled).toString().padStart(5, '0');
  const wholePart = digits.slice(0, -CURRENCY_DIGITS);
  const fractionPart = digits.slice(-CURRENCY_DIGITS).replace(/0+$/, '');
  const written = fractionPart === '' ? wholePart : wholePart + '.' + fractionPart;
  const signed = scaled < 0n ? '-' + written : written;
  return facts.decimal === '.' ? signed : signed.replace('.', facts.decimal);
}
