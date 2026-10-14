// VBScript dates are OLE Automation dates: a day count from 30 December 1899, with the time of day as
// the fraction. A Date here carries the same instant as a local wall-clock time.

const MS_PER_DAY = 86400000;
const OLE_EPOCH_UTC = Date.UTC(1899, 11, 30);

/** The OLE Automation number of a date: whole days since 30 December 1899, time of day as the fraction. */
export function dateToSerial(date: Date): number {
  const wallClock = Date.UTC(
    date.getFullYear(),
    date.getMonth(),
    date.getDate(),
    date.getHours(),
    date.getMinutes(),
    date.getSeconds(),
    date.getMilliseconds()
  );
  return (wallClock - OLE_EPOCH_UTC) / MS_PER_DAY;
}

/** The date an OLE Automation number stands for, as local wall-clock time. */
export function serialToDate(serial: number): Date {
  const wallClock = new Date(OLE_EPOCH_UTC + Math.round(serial * MS_PER_DAY));
  return new Date(
    wallClock.getUTCFullYear(),
    wallClock.getUTCMonth(),
    wallClock.getUTCDate(),
    wallClock.getUTCHours(),
    wallClock.getUTCMinutes(),
    wallClock.getUTCSeconds(),
    wallClock.getUTCMilliseconds()
  );
}

function isDayZero(date: Date): boolean {
  return date.getFullYear() === 1899 && date.getMonth() === 11 && date.getDate() === 30;
}

function isMidnight(date: Date): boolean {
  return date.getHours() === 0 && date.getMinutes() === 0 && date.getSeconds() === 0;
}

// Windows' default short date keeps day and month at two digits in nearly every locale; US English is the
// common exception (M/d/yyyy).
function unpaddedDayMonth(locale: string): boolean {
  return locale.toLowerCase() === 'en-us' || locale.toLowerCase() === 'en';
}

/**
 * CStr of a date, as VBScript writes it: the short date and the long time of the locale, without the date
 * when it is day 0 (30 December 1899) and without the time at midnight.
 */
export function formatVbDate(date: Date, locale: string): string {
  const padding = unpaddedDayMonth(locale) ? 'numeric' : '2-digit';
  const datePart = new Intl.DateTimeFormat(locale, { year: 'numeric', month: padding, day: padding }).format(date);
  const hour12 = new Intl.DateTimeFormat(locale, { hour: 'numeric' }).resolvedOptions().hour12 === true;
  const timePart = new Intl.DateTimeFormat(locale, {
    hour: hour12 ? 'numeric' : '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12,
  }).format(date);

  if (isDayZero(date)) return timePart;
  if (isMidnight(date)) return datePart;
  return `${datePart} ${timePart}`;
}

// ---------------------------------------------------------------------------
// Month names
// ---------------------------------------------------------------------------
// vbscript.dll reads month names from the locale, with two measured exceptions. A Chinese, Japanese or
// Korean locale's own names ("一月", "1月", "1월") are never matched and only the English names are read
// there; Thai reads both. Every other locale reads only its own names, which is why "1 January 2026" is a
// type mismatch under 1031 while "1 Januar 2026" is not.
type MonthNamePolicy = 'own' | 'english' | 'own+english';

const MONTH_NAME_POLICY: Record<string, MonthNamePolicy> = {
  zh: 'english',
  ja: 'english',
  ko: 'english',
  th: 'own+english',
};

const monthTableCache: Map<string, { month: number; name: string }[]> = new Map();

function bareMonthNames(locale: string): { month: number; name: string }[] {
  const long = new Intl.DateTimeFormat(locale, { month: 'long', timeZone: 'UTC' });
  const short = new Intl.DateTimeFormat(locale, { month: 'short', timeZone: 'UTC' });
  // A month name has a stand-alone and a format form, and they differ where the language inflects the
  // month ("январь" and "января" in Russian); vbscript.dll reads both, so both are collected here.
  const longInDateFormat = new Intl.DateTimeFormat(locale, { year: 'numeric', month: 'long', day: 'numeric', timeZone: 'UTC' });
  const shortInDateFormat = new Intl.DateTimeFormat(locale, { year: 'numeric', month: 'short', day: 'numeric', timeZone: 'UTC' });
  const named = (formatter: Intl.DateTimeFormat, day: Date): string =>
    formatter.formatToParts(day).find(part => part.type === 'month')?.value ?? '';

  const names: { month: number; name: string }[] = [];
  for (let month = 1; month <= 12; month++) {
    const day = new Date(Date.UTC(2026, month - 1, 15));
    names.push(
      { month, name: long.format(day) },
      { month, name: short.format(day) },
      { month, name: named(longInDateFormat, day) },
      { month, name: named(shortInDateFormat, day) }
    );
  }
  return names.filter(entry => entry.name !== '');
}

function monthTable(locale: string): { month: number; name: string }[] {
  const cached = monthTableCache.get(locale);
  if (cached) return cached;
  const policy = MONTH_NAME_POLICY[locale.toLowerCase().split(/[-_]/)[0]] ?? 'own';
  const names =
    policy === 'english'
      ? bareMonthNames('en-US')
      : policy === 'own+english'
        ? [...bareMonthNames(locale), ...bareMonthNames('en-US')]
        : bareMonthNames(locale);
  monthTableCache.set(locale, names);
  return names;
}

/**
 * The month a word names, or null. vbscript.dll matches three letters or more as a case-insensitive
 * prefix of one of the locale's month names ("Ja" is no month, "Janu" and "Januar" are January in
 * English), so the longest name that starts with the word decides.
 */
function monthOfWord(word: string, locale: string): number | null {
  const bare = word.toLowerCase().replace(/\.+$/, '');
  if (bare.length < 3) return null;
  let best: { month: number; length: number } | null = null;
  for (const entry of monthTable(locale)) {
    const name = entry.name.toLowerCase().replace(/\.+$/, '');
    if (!name.startsWith(bare)) continue;
    if (best === null || name.length > best.length) best = { month: entry.month, length: name.length };
  }
  return best === null ? null : best.month;
}

// ---------------------------------------------------------------------------
// Reading a date or a time
// ---------------------------------------------------------------------------

type DateOrder = 'DMY' | 'MDY' | 'YMD';

function dateOrderOf(locale: string): { order: DateOrder; separator: string } {
  const parts = new Intl.DateTimeFormat(locale, { year: 'numeric', month: '2-digit', day: '2-digit' }).formatToParts(
    new Date(2026, 8, 29)
  );
  const order = parts
    .filter(part => part.type === 'day' || part.type === 'month' || part.type === 'year')
    .map(part => part.type[0].toUpperCase())
    .join('') as DateOrder;
  const separator = parts.find(part => part.type === 'literal')?.value.trim() || '/';
  return { order: order === 'MDY' || order === 'YMD' ? order : 'DMY', separator };
}

function toFullYear(year: number, digits: number): number {
  if (digits > 2) return year;
  return year < 30 ? 2000 + year : 1900 + year;
}

/** A year field: two digits are a year of the near past or present, as vbscript.dll reads them. */
function yearOf(field: string): number {
  return field.length > 2 ? Number(field) : toFullYear(Number(field), field.length);
}

function validDate(year: number, month: number, day: number): Date | null {
  // OLE Automation dates run from 1 January 100 to 31 December 9999; a year outside that is no date.
  if (year < 100 || year > 9999) return null;
  const date = new Date(year, month - 1, day);
  date.setFullYear(year);
  return date.getMonth() === month - 1 && date.getDate() === day ? date : null;
}

/** A date written with a month name: "1 Jan 2026", "Jan 1, 2026", "März 1 2026", "Jan 2026", "1 Jan". */
function parseMonthNameDate(text: string, locale: string): Date | null {
  const words = text.split(/[\s,./-]+/).filter(word => word !== '');
  if (words.length < 2 || words.length > 3) return null;

  const months = words.map(word => monthOfWord(word, locale));
  const named = months.filter(month => month !== null).length;
  if (named !== 1) return null;
  const month = months[months.findIndex(entry => entry !== null)] as number;

  const numbers = words.filter((_, index) => months[index] === null);
  if (!/^\d+$/.test(numbers.join(''))) return null;

  // "Jan 2026" is January 1st of that year; "1 Jan" is the 1st of January this year.
  if (numbers.length === 1) {
    const [only] = numbers;
    if (only.length === 4 || Number(only) > 31) return validDate(yearOf(only), month, 1);
    return validDate(new Date().getFullYear(), month, Number(only));
  }

  // A four-digit number names the year wherever it stands ("1 Jan 2026"). With two short numbers the
  // locale's own order decides: a year-first locale reads "1 Jan 26" as 2001-01-26, the others as
  // 1 January 2026.
  const [first, second] = numbers;
  const fourDigit = numbers.find(number => number.length === 4);
  if (fourDigit !== undefined) {
    const other = numbers.find(number => number !== fourDigit) as string;
    return validDate(Number(fourDigit), month, Number(other));
  }
  const { order } = dateOrderOf(locale);
  const [year, day] = order === 'YMD' ? [first, second] : [second, first];
  return validDate(yearOf(year), month, Number(day));
}

/**
 * A date written with the locale's own markers instead of separators: "2026年9月29日" in Chinese and
 * Japanese, "2026년 9월 29일" in Korean. The layout comes from the locale itself, so only the markers it
 * actually uses are accepted, and a part that is not digits-then-marker bails out and leaves the text to
 * the other readers.
 */
function parseLocalizedDate(text: string, locale: string): Date | null {
  const parts = new Intl.DateTimeFormat(locale, {
    year: 'numeric',
    month: 'long',
    day: 'numeric',
    timeZone: 'UTC',
  }).formatToParts(new Date(Date.UTC(2026, 8, 29)));
  // Nothing to add when the locale spells its dates with "/", "-" or "." alone.
  if (!parts.some(part => part.type === 'literal' && /[^\s./-]/.test(part.value))) return null;

  const escape = (value: string): string => value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  let pattern = '^';
  const order: ('year' | 'month' | 'day')[] = [];
  for (const part of parts) {
    if (part.type === 'literal') {
      pattern += escape(part.value);
      continue;
    }
    if (part.type !== 'year' && part.type !== 'month' && part.type !== 'day') return null;
    // A month can carry its own marker ("9월"); everything after the digits stays a literal.
    const field = /^(\d+)(.*)$/s.exec(part.value);
    if (!field) return null;
    pattern += (part.type === 'year' ? '(\\d{1,4})' : '(\\d{1,2})') + escape(field[2]);
    order.push(part.type);
  }

  const read = (expression: RegExp, fields: readonly string[]): Date | null => {
    const found = expression.exec(text);
    if (!found) return null;
    let year = 0;
    let month = 0;
    let day = 0;
    fields.forEach((field, index) => {
      const value = Number(found[index + 1]);
      if (field === 'year') year = value;
      else if (field === 'month') month = value;
      else day = value;
    });
    return validDate(year, month, day);
  };

  const own = read(new RegExp(`${pattern}$`), order);
  if (own) return own;
  // The ideographic form as well as the locale's own markers: vbscript.dll reads "2026年9月29日" under
  // Korean too, even though Korean writes its own dates as "2026년 9월 29일". Measured on 2052, 1028,
  // 1041 and 1042; en-US and de-DE reject both forms.
  return read(/^(\d{1,4})年(\d{1,2})月(\d{1,2})日$/, ['year', 'month', 'day']);
}

function parseDatePart(text: string, locale: string): Date | null {
  const named = parseMonthNameDate(text, locale);
  if (named) return named;

  const localized = parseLocalizedDate(text, locale);
  if (localized) return localized;

  const iso = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(text);
  if (iso) return validDate(Number(iso[1]), Number(iso[2]), Number(iso[3]));

  const { order, separator } = dateOrderOf(locale);
  // "/" and "-" are read everywhere; the locale's own separator is read too, and it holds the text to the
  // locale's own order instead of letting the other order rescue it ("29.9.26" is no date in Hungarian).
  const own = separator === '/' || separator === '-' ? null : separator;
  const separatorAllowed = (mark: string): boolean => mark === '/' || mark === '-' || mark === own;

  const two = /^(\d{1,4})([./-])(\d{1,4})$/.exec(text);
  if (two && separatorAllowed(two[2])) {
    // "1/2" is the 1st of February in German and 2 January in English, in the current year.
    const year = new Date().getFullYear();
    const [day, month] = order === 'DMY' ? [Number(two[1]), Number(two[3])] : [Number(two[3]), Number(two[1])];
    return validDate(year, month, day) ?? validDate(year, day, month);
  }

  const match = /^(\d{1,4})([./-])(\d{1,2})\2(\d{1,4})$/.exec(text);
  if (!match || !separatorAllowed(match[2])) return null;

  const [first, second, third] = [match[1], match[3], match[4]];
  const strict = own !== null && match[2] === own;
  const candidates: [number, number, number][] = [];
  if (first.length === 4) {
    // "2026/09/29" is year first wherever it is written.
    candidates.push([Number(first), Number(second), Number(third)]);
  } else if (order === 'YMD') {
    candidates.push([Number(first), Number(second), Number(third)]);
    if (!strict) {
      candidates.push([yearOf(third), Number(first), Number(second)], [yearOf(third), Number(second), Number(first)]);
    }
  } else if (order === 'DMY') {
    candidates.push([yearOf(third), Number(second), Number(first)]);
    if (!strict) candidates.push([yearOf(third), Number(first), Number(second)]);
  } else {
    candidates.push([yearOf(third), Number(first), Number(second)]);
    if (!strict) candidates.push([yearOf(third), Number(second), Number(first)]);
  }
  for (const [year, month, day] of candidates) {
    const date = validDate(year, month, day);
    if (date) return date;
  }
  return null;
}

function parseTimePart(text: string): [number, number, number] | null {
  const match = /^(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?\s*([AaPp][Mm])?$/.exec(text);
  if (!match) return null;
  let hours = Number(match[1]);
  const minutes = Number(match[2]);
  const seconds = Number(match[3] ?? 0);
  const meridiem = match[4]?.toUpperCase();
  if (hours > 23 || minutes > 59 || seconds > 59) return null;
  // An hour above 12 wraps, as it does in vbscript.dll: "13:30 PM" is half past one in the afternoon.
  if (meridiem) hours = (hours % 12) + (meridiem === 'PM' ? 12 : 0);
  return [hours, minutes, seconds];
}

/**
 * A date or a time as CDate and IsDate read it: the locale's date order and separator (or ISO), a month
 * name, a 24-hour or AM/PM time, or a date and a time. A time alone is on day 0; a date alone is at
 * midnight. `null` when it is neither.
 *
 * Only a space separates a date from a time ("2026-09-29 10:30"); "2026-09-29T10:30:00" is a type
 * mismatch in vbscript.dll, and so is a weekday name ("Tue Sep 29 2026").
 */
export function parseVbDate(text: string, locale: string): Date | null {
  const trimmed = text.trim();
  if (trimmed === '') return null;

  const time = parseTimePart(trimmed);
  if (time) return new Date(1899, 11, 30, ...time);

  const words = trimmed.split(/\s+/);
  for (let split = words.length - 1; split >= 1; split--) {
    const date = parseDatePart(words.slice(0, split).join(' '), locale);
    if (!date) continue;
    const timeOfDay = parseTimePart(words.slice(split).join(' '));
    if (!timeOfDay) continue;
    date.setHours(...timeOfDay);
    return date;
  }
  return parseDatePart(trimmed, locale);
}
