// VBScript dates are OLE Automation dates: a day count from 30 December 1899, with the time of day as
// the fraction. A Date here carries the same instant as a local wall-clock time.

import { nlsDateFormat } from './nls-date-table.generated.ts';
import { NAME_IN_DATE, NAME_STANDALONE } from './name-tables.generated.ts';
import { createVbError, VbErrorCodes } from './errors.ts';

const MS_PER_DAY = 86400000;
const OLE_EPOCH_UTC = Date.UTC(1899, 11, 30);

/** LOCALE_ICALENDARTYPE for the Gregorian calendar, which every row of the table carries as 1. */
/** LOCALE_ICALENDARTYPE for the Gregorian calendar, which every row of the table carries as 1. */
const GREGORIAN = 1;

/** LOCALE_ICALENDARTYPE for the Thai Buddhist calendar, the one calendar with no date format to expand. */
const THAI_BUDDHIST = 7;

/** LOCALE_ICALENDARTYPE for the Hijri calendar, which Windows writes with a two-digit year here. */
const HIJRI = 23;

/** LOCALE_ICALENDARTYPE for the Persian calendar, whose day Windows never pads. */
const PERSIAN = 22;

/** U+200F RIGHT-TO-LEFT MARK, which Windows writes before each date field in the locales that use one. */
const RLM = '\u200f';

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

/**
 * The same wall-clock date and time, moved onto the UTC timeline.
 *
 * A date here is a local wall-clock reading (see the note at the top of the file). The name lookups below
 * go through Intl with its zone pinned to UTC so a name cannot depend on the machine's zone, so a date has
 * to be carried across before it is handed to them: otherwise a machine east of Greenwich would ask Intl
 * about the calendar day before, and every weekday name would come back one day out.
 */
function onUtcTimeline(date: Date): Date {
  return new Date(
    Date.UTC(
      date.getFullYear(),
      date.getMonth(),
      date.getDate(),
      date.getHours(),
      date.getMinutes(),
      date.getSeconds(),
      date.getMilliseconds()
    )
  );
}

// Windows' default short date keeps day and month at two digits in nearly every locale; US English is the
// common exception (M/d/yyyy).
function unpaddedDayMonth(locale: string): boolean {
  return locale.toLowerCase() === 'en-us' || locale.toLowerCase() === 'en';
}

/**
 * A localized name for one picture-string token, from Intl.
 *
 * Only the names are Intl's here; the layout stays Windows' (see the note on formatVbDate). Three things
 * about how Intl is asked were measured:
 *
 *   * A name is looked up for the date itself and never through an index computed here. de-DE writes its
 *     week starting on Monday, so a table built in week order would name every day one out.
 *   * A month is asked for inside a dated context, because that is the form Windows prints ("января",
 *     "ledna", "tammikuuta"), while a weekday is asked for on its own: Finnish would otherwise answer the
 *     essive "maanantaina" where Windows writes the nominative "maanantai".
 *   * A date is carried onto the UTC timeline first, because the formatters below pin their zone to UTC so
 *     a name cannot depend on the machine's zone. Without that step a machine east of Greenwich would ask
 *     about the calendar day before, and every weekday name would come back one day out.
 */
interface NameFormatters {
  monthLong: Intl.DateTimeFormat;
  monthShort: Intl.DateTimeFormat;
  weekdayLong: Intl.DateTimeFormat;
  weekdayShort: Intl.DateTimeFormat;
}

const nameCache: Map<string, NameFormatters> = new Map();

/** One name per locale, token and month-or-weekday index; '' records a locale that gives none. */
const localizedNameCache: Map<string, string> = new Map();

/** Per locale, whether Windows writes a right-to-left mark in front of that locale's date fields. */
const marksFieldsCache: Map<string, boolean> = new Map();

/** Per locale, whether Intl answers in the language it was asked for. */
const answersInOwnLanguageCache: Map<string, boolean> = new Map();

function nameFormatters(locale: string): NameFormatters {
  const cached = nameCache.get(locale);
  if (cached) return cached;
  const utc = { timeZone: 'UTC' } as const;
  const formatters: NameFormatters = {
    // A day beside the month is what selects the format form of the name.
    monthLong: new Intl.DateTimeFormat(locale, { month: 'long', day: 'numeric', ...utc }),
    monthShort: new Intl.DateTimeFormat(locale, { month: 'short', day: 'numeric', ...utc }),
    weekdayLong: new Intl.DateTimeFormat(locale, { weekday: 'long', ...utc }),
    weekdayShort: new Intl.DateTimeFormat(locale, { weekday: 'short', ...utc }),
  };
  nameCache.set(locale, formatters);
  return formatters;
}

/**
 * Whether Intl answers in the language it was asked for.
 *
 * One of the engine's locales, 1073 Tsonga, carries no ICU data, and the runtime then answers in the
 * machine's own locale instead of failing: on this Chinese machine every Tsonga name came back Chinese. A
 * locale answered in another language is refused here so that no other language's words can reach a script,
 * and the caller falls back to the numeric form it can still produce from the picture string.
 */
function answersInItsOwnLanguage(locale: string): boolean {
  const cached = answersInOwnLanguageCache.get(locale);
  if (cached !== undefined) return cached;
  let own: boolean;
  try {
    const resolved = new Intl.DateTimeFormat(locale).resolvedOptions().locale;
    own = resolved.toLowerCase().split('-')[0] === locale.toLowerCase().split(/[-_]/)[0];
  } catch {
    own = false;
  }
  answersInOwnLanguageCache.set(locale, own);
  return own;
}

/**
 * Whether Windows puts a right-to-left mark in front of each date field of the locale.
 *
 * Asked of the language, not of the script: Urdu and Yiddish are written in the Arabic and Hebrew scripts
 * and Windows marks neither of them, while it marks Arabic ("‏05/‏01/‏2026") and Hebrew
 * ("‏יום שני ‏05 ‏ינואר ‏2026"). ICU's own Arabic separators carry the mark and its Hebrew ones do not, so
 * looking for a mark in a formatted date draws the line in the wrong place too.
 */
function marksDateFields(locale: string): boolean {
  const cached = marksFieldsCache.get(locale);
  if (cached !== undefined) return cached;
  const language = locale.toLowerCase().split(/[-_]/)[0];
  const marked = language === 'ar' || language === 'he';
  marksFieldsCache.set(locale, marked);
  return marked;
}

/**
 * A name as Windows writes it for that locale, with the one spelling difference its scripts make.
 *
 * Windows separates the two words of an Arabic month name with U+00A0 ("كانون الثاني") and the two words of
 * a Hebrew weekday name the same way ("יום שני") where CLDR uses a plain space. Only for those scripts: the
 * other locales whose names contain a space write it plain.
 */
function scriptName(value: string, locale: string): string {
  return marksDateFields(locale) ? value.replace(/ /g, '\u00a0') : value;
}

/**
 * A weekday name as Windows writes it.
 *
 * Hebrew's day-word "יום" is dropped from one weekday and kept on the others: Windows writes "שבת" for
 * Saturday but "יום שני" for Monday. That is not a matter of the shape of what is left - both remainders are
 * one word - and Intl says which it is by itself: its short weekday form for the locale is "שבת" for Saturday
 * and "יום ב׳" for Monday, so the prefix follows the short form. Only Hebrew, and only a weekday, so the
 * locale's month names are untouched.
 */
function weekdayName(value: string, locale: string, short: string): string {
  if (locale.toLowerCase().split(/[-_]/)[0] !== 'he') return value;
  return short.startsWith('יום') ? value : value.replace(/^יום\s+/, '');
}

/** One part of an Intl-formatted date, or '' when the locale gives none. */
function partOf(
  formatter: Intl.DateTimeFormat,
  date: Date,
  type: Intl.DateTimeFormatPartTypes
): string {
  return formatter.formatToParts(date).find(part => part.type === type)?.value ?? '';
}

/**
 * The standalone month name, which is not always the one a date uses.
 *
 * CLDR keeps two forms per month name: the one written beside a day and the one written on its own. German
 * spells March "März" in a date and "Mär" alone, and Catalan puts a preposition in the date's form ("de
 * gener") that the standalone name does without ("gener"). MonthName answers with the standalone form,
 * which is also what Windows' LOCALE_SMONTHNAME holds.
 */
const standaloneMonthCache: Map<string, Intl.DateTimeFormat> = new Map();

function standaloneMonthFormatter(locale: string, short: boolean): Intl.DateTimeFormat {
  const key = `${locale}\u0000${short ? 'short' : 'long'}`;
  const cached = standaloneMonthCache.get(key);
  if (cached) return cached;
  const formatter = new Intl.DateTimeFormat(locale, {
    month: short ? 'short' : 'long',
    timeZone: 'UTC',
  });
  standaloneMonthCache.set(key, formatter);
  return formatter;
}

/** 2026-01-04 is a Sunday, so this is the date that carries weekday index 0. */
const WEEKDAY_ORIGIN = Date.UTC(2026, 0, 4);

/**
 * A month name as MonthName answers with it: the name on its own, not the one a date spells.
 *
 * @param lcid - The LCID whose overrides apply.
 * @param locale - The locale tag Intl is asked with.
 * @param month - 1 for January through 12 for December.
 * @param short - Whether the abbreviated form was asked for.
 */
export function standaloneMonthName(
  lcid: number,
  locale: string,
  month: number,
  short: boolean
): string {
  const names = NAME_STANDALONE[lcid];
  const override = short ? names?.monthsShort?.[month - 1] : names?.monthsLong?.[month - 1];
  if (override !== undefined) return override;
  // MonthName has no numeric form to fall back on, so a locale Intl will not answer for gives '' rather
  // than another language's word.
  if (!answersInItsOwnLanguage(locale)) return '';
  const formatter = standaloneMonthFormatter(locale, short);
  return scriptName(partOf(formatter, new Date(Date.UTC(2026, month - 1, 15)), 'month'), locale);
}

/**
 * A weekday name as WeekdayName answers with it, indexed the way VBScript indexes it: 1 is Sunday.
 *
 * @param lcid - The LCID whose overrides apply.
 * @param locale - The locale tag Intl is asked with.
 * @param weekday - 1 for Sunday through 7 for Saturday.
 * @param short - Whether the abbreviated form was asked for.
 */
export function standaloneWeekdayName(
  lcid: number,
  locale: string,
  weekday: number,
  short: boolean
): string {
  const index = weekday - 1;
  const names = NAME_STANDALONE[lcid];
  const override = short ? names?.weekdaysShort?.[index] : names?.weekdaysLong?.[index];
  if (override !== undefined) return override;
  // WeekdayName has no numeric form to fall back on either.
  if (!answersInItsOwnLanguage(locale)) return '';
  const formatters = nameFormatters(locale);
  const date = new Date(WEEKDAY_ORIGIN + index * 86400000);
  const value = partOf(short ? formatters.weekdayShort : formatters.weekdayLong, date, 'weekday');
  return scriptName(
    weekdayName(value, locale, partOf(formatters.weekdayShort, date, 'weekday')),
    locale
  );
}

/**
 * How a picture string's date fields are filled.
 *
 * The locale's own calendar supplies them for the Hijri and Persian locales, whose long dates Windows
 * writes as "16 رجب 1447" and "15 دی 1404"; the Thai Buddhist shift covers th-TH, whose dates Windows
 * refuses to write at all; everything else is the Gregorian reading of the date. Which one applies comes
 * from the table's LOCALE_ICALENDARTYPE and was confirmed per locale against vbscript.dll.
 */
type DateSource = { kind: 'date' } | { kind: 'calendar'; type: number };

/** The reading a date with no calendar of its own has: the Gregorian year, month and day it carries. */
const GREGORIAN_DATE: DateSource = { kind: 'date' };

/** The Unicode calendar name for a LOCALE_ICALENDARTYPE the table can carry. */
const CALENDAR_NAME: Record<number, string> = {
  1: 'gregory',
  6: 'islamic-umalqura',
  7: 'buddhist',
  22: 'persian',
  23: 'islamic-umalqura',
};

/**
 * The calendar Windows writes this locale's dates in, from the table's LOCALE_ICALENDARTYPE.
 *
 * That column is the authority and agrees with the patterns beside it: it is 23 for ar-SA, whose dates
 * Windows writes in Hijri, and 1 for every other Arabic locale, which print a Gregorian date with Arabic
 * month names. It is deliberately not ICU's answer for a locale, which leaves all of them Gregorian.
 */
function usesOwnCalendar(calendarType: number): boolean {
  return calendarType !== GREGORIAN;
}

/**
 * The locale tag to ask Intl about, with the calendar the table names for it.
 *
 * ICU keeps its own idea of a locale's calendar - it leaves ar-SA on the Gregorian one where
 * LOCALE_ICALENDARTYPE says 23 - so the calendar is spelled out. Both the numbers and the month names of
 * a date then come from the same calendar, which is what stops an Arabic month name being paired with a
 * Hijri year.
 */
function taggedLocale(locale: string, source: DateSource): string {
  const calendar = source.kind === 'calendar' ? CALENDAR_NAME[source.type] : 'gregory';
  return `${locale}-u-ca-${calendar ?? 'gregory'}`;
}

/** The same tag with Latin digits asked for, for the parts the engine reads as numbers. */
function taggedNumericLocale(locale: string, source: DateSource): string {
  return `${taggedLocale(locale, source)}-nu-latn`;
}

/**
 * The fields a picture string should write for this locale.
 *
 * The table's LOCALE_ICALENDARTYPE is the authority - it is 23 for the Hijri locales and 22 for fa-IR -
 * and it is deliberately not ICU's answer for the locale, which leaves ar-SA on the Gregorian calendar.
 */
function dateFields(
  date: Date,
  source: DateSource,
  locale: string
): { day: number; month: number; year: number } {
  if (source.kind !== 'calendar') {
    return { day: date.getUTCDate(), month: date.getUTCMonth() + 1, year: date.getUTCFullYear() };
  }
  const parts = new Intl.DateTimeFormat(taggedNumericLocale(locale, source), {
    year: 'numeric',
    month: 'numeric',
    day: 'numeric',
    timeZone: 'UTC',
  }).formatToParts(date);
  const read = (type: Intl.DateTimeFormatPartTypes): number =>
    Number(parts.find(part => part.type === type)?.value ?? '0');
  return { day: read('day'), month: read('month'), year: read('year') };
}

/** The number a year token writes. */
function yearValue(source: DateSource, date: Date, locale: string): number {
  return dateFields(date, source, locale).year;
}

/**
 * The digit string a numeric date token writes, or null for a token that is not one.
 *
 * The Persian calendar pads its day only while its year is short: fa-IR writes "۱۵/۱۰/۱۴۰۴" unpadded but
 * "09/10/1278" padded, so the padding follows the year rather than the pattern. The Hijri calendar pads
 * always ("16/07/47") and the Gregorian ones pad as their pattern asks.
 */
function dateNumericToken(
  token: string,
  date: Date,
  source: DateSource,
  locale: string,
  unpaddedPersianDay = false
): string | null {
  const pad = (value: number, width: number): string => String(value).padStart(width, '0');
  const { day, month, year } = dateFields(date, source, locale);
  const persian = source.kind === 'calendar' && source.type === PERSIAN;
  switch (token) {
    case 'd':
      return String(day);
    case 'dd':
      // The Persian short date asks for a padded day and gets one ("04/09/1405"), while its long date asks
      // for the same "dd" and Windows leaves the zero off ("7 اسفند 1405"). Both were measured across a
      // Persian year, and the flag says which of the two expansions is running.
      return persian && unpaddedPersianDay ? String(day) : pad(day, 2);
    case 'M':
      return String(month);
    case 'MM':
      return pad(month, 2);
    case 'y':
      return String(year % 100);
    case 'yy':
      return pad(year % 100, 2);
    case 'yyyy':
      return pad(year, 4);
    default:
      return null;
  }
}

/**
 * The marker for a name Intl could not answer for. A picture string writes its separators for the fields
 * around them, so a field that is dropped takes its neighbouring punctuation with it.
 */
const NO_NAME = '\u0000';

/** The value of one picture-string token, or null for a token this does not map. */
function localizedToken(
  token: string,
  date: Date,
  locale: string,
  year: DateSource,
  am: string,
  pm: string,
  lcid: number
): string | null {
  // A Hijri date wants the Hijri month names and a Gregorian date the Gregorian ones, so the names are
  // asked of the same calendar-tagged locale the numbers came from.
  const nameLocale = taggedLocale(locale, year);
  const formatters = nameFormatters(nameLocale);
  // A month name depends on the month and a weekday name on the weekday, and a picture string can ask for
  // the same one many times over one date, so the answer is remembered per locale, token and index.
  const name = (index: number, override: string | undefined, read: () => string): string => {
    // The few names Windows spells differently are listed rather than derived: nothing in either source says
    // which spelling vbscript.dll used.
    if (override !== undefined) return override;
    // A locale Intl will not answer for has no name to give; the expansion is told so, rather than being
    // handed another language's word or an empty string in the middle of a date.
    if (!answersInItsOwnLanguage(locale)) return NO_NAME;
    const key = `${nameLocale}\u0000${token}\u0000${index}`;
    const cached = localizedNameCache.get(key);
    if (cached !== undefined) return cached;
    const value = scriptName(read(), locale);
    localizedNameCache.set(key, value);
    return value;
  };
  const overrides = NAME_IN_DATE[lcid];
  // The name a calendar locale writes belongs to that calendar's month: the Islamic year is about eleven
  // days shorter than the Gregorian one, so a Gregorian month number picks the wrong name for the instants
  // near a month boundary, and the table is keyed the same way this looks it up.
  const monthIndex =
    year.kind === 'calendar' ? dateFields(date, year, locale).month - 1 : date.getUTCMonth();

  const hour = date.getUTCHours();
  const meridiem = hour < 12 ? am || 'AM' : pm || 'PM';

  switch (token) {
    case 'MMM':
      return name(monthIndex, overrides?.months?.[monthIndex], () =>
        partOf(formatters.monthShort, date, 'month')
      );
    case 'MMMM':
      return name(monthIndex, overrides?.months?.[monthIndex], () =>
        partOf(formatters.monthLong, date, 'month')
      );
    case 'ddd':
      return name(date.getUTCDay(), overrides?.weekdays?.[date.getUTCDay()], () =>
        weekdayName(
          partOf(formatters.weekdayShort, date, 'weekday'),
          locale,
          partOf(formatters.weekdayShort, date, 'weekday')
        )
      );
    case 'dddd':
      return name(date.getUTCDay(), overrides?.weekdays?.[date.getUTCDay()], () =>
        weekdayName(
          partOf(formatters.weekdayLong, date, 'weekday'),
          locale,
          partOf(formatters.weekdayShort, date, 'weekday')
        )
      );
    case 'yy':
      return String(yearValue(year, date, locale) % 100).padStart(2, '0');
    case 'yyyy':
      return String(yearValue(year, date, locale)).padStart(4, '0');
    case 'tt':
      return meridiem;
    case 't':
      return meridiem.charAt(0);
    default:
      return null;
  }
}

/**
 * Expands a Windows picture string, the way vbscript.dll writes a date. Returns null for a pattern with a
 * token this does not expand, so the caller can fall back to Intl.
 *
 * `markFields` says the locale writes its date fields with a right-to-left mark in front of each, as
 * Windows does for Arabic: the short date comes out "‏05/‏01/‏2026". The mark is written where the field is
 * written, so no field can be missed and none can be marked twice, and the separators keep whatever the
 * picture string put between them.
 *
 * `clock` says the pattern is a clock rather than a date, which decides what becomes of U+202F: Windows
 * writes a narrow no-break space in the time format but a plain space in the time it produces, and keeps
 * the narrow one in a date (ru-MD's long date separates its year from "г." with U+202F).
 *
 * `lcid` is only needed to look up the few names Windows spells differently from CLDR; 0 means no table.
 */
function expandNlsPattern(
  pattern: string,
  date: Date,
  locale: string,
  year: DateSource,
  am: string,
  pm: string,
  markFields = false,
  clock = false,
  lcid = 0,
  persianDayUnpadded = false
): string | null {
  const hour = date.getUTCHours();
  const hour12 = hour % 12 === 0 ? 12 : hour % 12;
  const pad = (value: number, width: number): string => String(value).padStart(width, '0');
  const fields: Record<string, string> = {
    d: dateNumericToken('d', date, year, locale) as string,
    dd: dateNumericToken('dd', date, year, locale, persianDayUnpadded) as string,
    M: dateNumericToken('M', date, year, locale) as string,
    MM: dateNumericToken('MM', date, year, locale) as string,
    y: dateNumericToken('y', date, year, locale) as string,
    yy: dateNumericToken('yy', date, year, locale) as string,
    yyyy: dateNumericToken('yyyy', date, year, locale) as string,
    h: String(hour12),
    hh: pad(hour12, 2),
    H: String(hour),
    HH: pad(hour, 2),
    m: String(date.getUTCMinutes()),
    mm: pad(date.getUTCMinutes(), 2),
    s: String(date.getUTCSeconds()),
    ss: pad(date.getUTCSeconds(), 2),
  };

  // The expansion is collected first and marked afterwards, so where a mark lands cannot depend on how the
  // loop happened to accumulate the text so far. An entry is either the picture string's own text or a
  // value the date supplies.
  interface Piece {
    text: string;
    isDateField: boolean;
  }
  const pieces: Piece[] = [];
  let dropped = false;

  for (let index = 0; index < pattern.length;) {
    const char = pattern[index];
    if (char === "'") {
      const end = pattern.indexOf("'", index + 1);
      pieces.push({
        text: pattern.slice(index + 1, end < 0 ? pattern.length : end),
        isDateField: false,
      });
      index = end < 0 ? pattern.length : end + 1;
      continue;
    }
    let run = 1;
    while (pattern[index + run] === char) run++;
    const token = char.repeat(run);
    if (token === '\u202f') {
      // Windows separates the time from the AM/PM designator with U+202F in LOCALE_STIMEFORMAT, but the
      // formatting it produces uses a normal space there, while keeping any U+202F inside LOCALE_S1159.
      pieces.push({ text: clock ? ' ' : token, isDateField: false });
    } else if (token in fields) {
      // Only a date field is marked; a clock's is not, which is why the token is asked about rather than
      // every value in the table being taken for one.
      const isDateField = dateNumericToken(token, date, year, locale) !== null;
      pieces.push({ text: fields[token], isDateField });
    } else if (!/^[A-Za-z]+$/.test(token)) {
      // A separator, a dot after a day, or a word the locale writes between two fields (Thai "ที่") is the
      // picture string's own text and is copied through.
      pieces.push({ text: token, isDateField: false });
    } else {
      // Every remaining token is a name or a designator, and Intl supplies all of them; an unknown one
      // fails the expansion rather than printing a wrong name, and one with no name at all is left out.
      let value = localizedToken(token, date, locale, year, am, pm, lcid);
      if (value === null) return null;
      if (value === NO_NAME) {
        dropped = true;
      } else {
        // A Thai picture string writes "'วัน'dddd", where "วัน" is the day-word the long weekday name
        // already begins with, so Windows prints the short name there ("วันจันทร์", not "วันวันจันทร์").
        if (
          token === 'dddd' &&
          pieces
            .map(piece => piece.text)
            .join('')
            .endsWith(value.slice(0, 3)) &&
          value.length > 3
        ) {
          value = localizedToken('ddd', date, locale, year, am, pm, lcid) ?? value;
        }
        pieces.push({
          text: value,
          isDateField: !clock && dateNumericToken(token, date, year, locale) === null,
        });
      }
    }
    index += run;
  }

  // A mark goes in front of every date field, the first one included: he-IL's long date is
  // "‏יום שני ‏05 ‏ינואר ‏2026" and ar-SA's short date "‏05/‏01/‏2026". A clock's fields are not dated, so none
  // of them takes one.
  const out = pieces
    .map(piece => (piece.isDateField && markFields ? RLM + piece.text : piece.text))
    .join('');
  // The punctuation a dropped field was written with belongs to it, so it goes too: Tsonga's
  // "yyyy MMMM d, dddd" would otherwise print "2026 1 5," with nothing after the comma.
  return dropped ? out.replace(/[\s\p{P}]+$/u, '') : out;
}

/**
 * The picture string for a short time: Windows' LOCALE_SSHORTTIME without the AM/PM designator, on a
 * 24-hour clock with a two-digit hour.
 *
 * Measured, not read from the table's own tokens: FormatDateTime(x, 4) prints "09:05" for every locale
 * whose short time carries a designator (1033 "h:mm tt", 1042 "tt h:mm", 1025 "hh:mm tt"), and it pads the
 * hour even where LOCALE_SSHORTTIME does not ask for it (1026 "H:mm" -> "09:05"). The separator stays the
 * locale's own ("09.05" for 1035 "H.mm" and 1057 "HH.mm").
 */
function shortTimePicture(shortTime: string): string {
  // Walking the picture string rather than rewriting it with a regular expression keeps a literal that
  // happens to contain an "h" - Upper Sorbian's "HH:mm 'hodź.'" - out of the way.
  let out = '';
  let designator = false;
  for (let index = 0; index < shortTime.length;) {
    const char = shortTime[index];
    if (char === "'") {
      // The only table row that quotes anything here is Upper Sorbian's "H:mm 'hodź'.", whose quoted word
      // is the locale's label for the time and is not part of the clock it prints; the engine's 131 other
      // short-time rows quote nothing at all.
      const end = shortTime.indexOf("'", index + 1);
      index = end < 0 ? shortTime.length : end + 1;
      continue;
    }
    let run = 1;
    while (shortTime[index + run] === char) run++;
    const token = char.repeat(run);
    if (token === 'h' || token === 'H' || token === 'hh' || token === 'HH') {
      out += 'HH';
    } else if (token === 't' || token === 'tt') {
      designator = true;
    } else {
      out += token;
    }
    index += run;
  }
  // A designator that stood apart leaves a space behind; "H.mm" keeps its own spacing. A dot that followed
  // a dropped quoted word is that word's own punctuation (Upper Sorbian's "hodź'."), so it goes too.
  const trimmed = designator ? out.replace(/\s+/g, ' ').trim() : out;
  return trimmed.replace(/[\s.]+$/g, '');
}

/** A picture string's pieces: a quoted literal, a run of one letter, or anything else on its own. */
interface PatternChunk {
  text: string;
  quoted: boolean;
  letters: boolean;
}

function patternChunks(pattern: string): PatternChunk[] {
  const chunks: PatternChunk[] = [];
  for (let index = 0; index < pattern.length;) {
    const char = pattern[index];
    if (char === "'") {
      const end = pattern.indexOf("'", index + 1);
      const stop = end < 0 ? pattern.length : end + 1;
      chunks.push({ text: pattern.slice(index, stop), quoted: true, letters: false });
      index = stop;
      continue;
    }
    let run = 1;
    while (pattern[index + run] === char) run++;
    chunks.push({ text: char.repeat(run), quoted: false, letters: /[A-Za-z]/.test(char) });
    index += run;
  }
  return chunks;
}

/** A token that writes a number out of the date, as opposed to a name or a designator. */
const DATE_FIELD_TOKEN = /^(d{1,4}|M{1,4}|y{1,4})$/;

/**
 * A picture string with its weekday group taken out.
 *
 * Thai is the one locale whose long date asks for a weekday: LOCALE_SLONGDATE is "'วัน'dddd'ที่' d MMMM yyyy",
 * and Windows writes "5 มกราคม 2569" from it. Measured on all seven weekdays and all five instants, neither
 * the day-word "วัน", nor the weekday name, nor "ที่" ever appears, and the same NLS row's weekday names do
 * reach VBScript through WeekdayName. The picture is therefore expanded from the part Windows keeps, which
 * is the date fields after that group, rather than from a picture of the engine's own.
 */
export function withoutWeekdayGroup(picture: string): string {
  const chunks = patternChunks(picture);
  const weekday = chunks.findIndex(chunk => chunk.letters && /^d{3,4}$/.test(chunk.text));
  if (weekday < 0) return picture;
  const dropped = new Set([weekday]);
  // The literals touching the weekday belong to it, and so does the space that followed the group: Thai
  // writes the three as one run at the head of the picture, and the date fields start after them.
  for (let index = weekday - 1; index >= 0 && chunks[index].quoted; index--) dropped.add(index);
  for (let index = weekday + 1; index < chunks.length; index++) {
    const chunk = chunks[index];
    if (!chunk.quoted && !/^\s+$/.test(chunk.text)) break;
    dropped.add(index);
  }
  return chunks
    .filter((_, index) => !dropped.has(index))
    .map(chunk => chunk.text)
    .join('')
    .trimStart();
}

/**
 * A picture string with the separators between its date fields replaced by the short date's mark.
 *
 * The Hijri long date is the one Windows writes this way: ar-SA's picture is "dd MMMM, yyyy" and
 * FormatDateTime(x, 1) gives "16/رجب/1447", with the short date's "/" between the fields and the comma
 * dropped. Measured across the Hijri months. The Persian long date keeps its picture's own separators
 * ("7 اسفند 1405"), so this follows the Hijri calendar and not a calendar in general.
 */
export function hijriLongDatePicture(picture: string): string {
  const chunks = patternChunks(picture);
  let separated = '';
  let afterField = false;
  for (const [index, chunk] of chunks.entries()) {
    if (chunk.letters && DATE_FIELD_TOKEN.test(chunk.text)) {
      separated += chunk.text;
      afterField = true;
      continue;
    }
    const fieldAfter = chunks
      .slice(index + 1)
      .some(part => part.letters && DATE_FIELD_TOKEN.test(part.text));
    if (afterField && fieldAfter) {
      // The run of literals between two fields is one separator, written once: ar-SA's picture has both a
      // comma and a space there, and Windows writes a single "/".
      if (!separated.endsWith('/')) separated += '/';
      continue;
    }
    separated += chunk.text;
  }
  return separated;
}

/** The Intl-based date, kept for a locale with no NLS row and for patterns this cannot expand. */
function formatVbDateViaIntl(date: Date, locale: string): string {
  const padding = unpaddedDayMonth(locale) ? 'numeric' : '2-digit';
  const datePart = new Intl.DateTimeFormat(locale, {
    year: 'numeric',
    month: padding,
    day: padding,
  }).format(date);
  const hour12 =
    new Intl.DateTimeFormat(locale, { hour: 'numeric' }).resolvedOptions().hour12 === true;
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

/**
 * CStr of a date, as VBScript writes it: the locale's short date and long time, without the date when it
 * is day 0 (30 December 1899) and without the time at midnight.
 *
 * The patterns come from Windows NLS, shipped as data (see nls-date-table.generated.ts), because a browser
 * has no NLS to ask and CLDR disagrees with it: the two agree for 46 of the 133 locales the engine knows
 * when Intl is left to choose the layout itself.
 *
 * CStr does not mark its fields, measured on all 131 locales: ar-SA prints "05/01/2026 09:05:00 ص" with no
 * directional mark, where the short date of FormatDateTime does carry them.
 */
export function formatVbDate(date: Date, lcid: number, locale: string): string {
  const format = nlsDateFormat(lcid);
  if (format) {
    const { am, pm } = format;
    const wallClock = onUtcTimeline(date);
    // CStr writes the Gregorian date for every locale but Thai, ar-SA included: ar-SA's short date is
    // "05/01/2026" where FormatDateTime's is the Hijri, and no locale marks a field of its own here. Thai is
    // the exception because its year is Buddhist in every style, CStr included: CStr(DateSerial(2026,1,5))
    // is "5/1/2569", measured on a fresh process like everything else in this file.
    const year: DateSource =
      format.calendarType === THAI_BUDDHIST
        ? { kind: 'calendar', type: THAI_BUDDHIST }
        : GREGORIAN_DATE;
    const datePart = expandNlsPattern(
      format.shortDate,
      wallClock,
      locale,
      year,
      am,
      pm,
      false,
      false,
      lcid
    );
    const timePart = expandNlsPattern(
      format.timeFormat,
      wallClock,
      locale,
      year,
      am,
      pm,
      false,
      true
    );
    if (datePart !== null && timePart !== null) {
      if (isDayZero(date)) return timePart;
      if (isMidnight(date)) return datePart;
      return `${datePart} ${timePart}`;
    }
  }
  return formatVbDateViaIntl(date, locale);
}

// ---------------------------------------------------------------------------
// FormatDateTime
// ---------------------------------------------------------------------------
// The named formats of VBScript's FormatDateTime, which are the OLE Automation vbDateTimeFormat constants.
// Every one of them is a Windows NLS call and not a CLDR layout: the measured values for
// 2026-01-05 09:05:00 are "05.01.2026 09:05:00" (1031), "2026년 1월 5일 월요일" (1042) and
// "den 5 januari 2026" (1053), which Intl does not produce for any of those locales.
export const VbDateTimeFormat = {
  GeneralDate: 0,
  LongDate: 1,
  ShortDate: 2,
  LongTime: 3,
  ShortTime: 4,
} as const;

/**
 * FormatDateTime, as vbscript.dll answers it.
 *
 * What the measurements pin down (a fresh cscript process per locale, on 2026-01-05 09:05:00, midnight,
 * 2026-01-05 15:07:09 and both day-0 instants, for all 131 locales the table covers):
 *   * GeneralDate is the short date with the LOCALE_STIMEFORMAT clock. It prints the time alone at midnight
 *     and on day 0, as CStr does, except in a Hijri locale, which keeps its date on day 0.
 *   * LongDate is LOCALE_SLONGDATE, so its weekday and month names are the long ones.
 *   * ShortTime is the table's LOCALE_SSHORTTIME with the AM/PM designator dropped and a 24-hour clock,
 *     NOT a hard-coded "HH:mm": Finnish prints "09.05" and Indonesian "09.05", keeping their separator.
 *   * The two date formats use the locale's own calendar - ar-SA's long date is the Hijri "16 رجب 1447" and
 *     its short date the Hijri "16/07/47", fa-IR's the Persian "15 دی 1404" - while GeneralDate and both
 *     clocks stay Gregorian for every locale.
 *   * The short date that GeneralDate prints carries no right-to-left mark, while the one ShortDate prints
 *     does: he-IL gives "05/01/2026 09:05:00" for style 0 and "‏05/‏01/‏2026" for style 2.
 *
 * Three deliberate deviations, all of them measured rather than assumed:
 *   * Thai writes its Buddhist year in every style, CStr included ("5/1/2569" where the Gregorian reading is
 *     "5/1/2026"), and its long date carries no weekday: LOCALE_SLONGDATE asks for one and Windows leaves
 *     the whole group out, measured on all seven weekdays.
 *   * A Hijri picture string separates its fields with spaces and Windows writes "/" there, so the engine's
 *     ar-SA long date reads "‏16 ‏رجب, ‏1447" where Windows writes "‏16/‏رجب/‏1447".
 *   * A few names are CLDR's where Windows ships an older or a longer spelling: Arabic's two-word month
 *     names (Windows separates the two words with a narrow no-break space the engine does not print),
 *     Latvian's capitalized "Pirmdiena", Vietnamese's "tháng 1", and the Sotho and Tswana month names.
 */
/**
 * FormatDateTime, as vbscript.dll answers it.
 *
 * What the measurements pin down (a fresh cscript process per locale, on 2026-01-05 09:05:00, midnight,
 * 2026-01-05 15:07:09 and both day-0 instants, for all 131 locales the table covers):
 *   * GeneralDate is the short date with the LOCALE_STIMEFORMAT clock. It prints the time alone at midnight
 *     and on day 0, as CStr does, except in a Hijri locale, which keeps its date on day 0.
 *   * LongDate is LOCALE_SLONGDATE, so its weekday and month names are the long ones.
 *   * ShortTime is the table's LOCALE_SSHORTTIME with the AM/PM designator dropped and a 24-hour clock, NOT
 *     a hard-coded "HH:mm": Finnish prints "09.05" and Indonesian "09.05", keeping their separator.
 *   * The two date formats use the locale's own calendar where the table names one - ar-SA's long date is
 *     the Hijri "16 رجب 1447" and its short date the Hijri "16/07/47", fa-IR's the Persian "15 دی 1404" -
 *     while GeneralDate and both clocks stay Gregorian for every locale.
 *   * The short date that GeneralDate prints carries no right-to-left mark, while the one ShortDate prints
 *     does: he-IL gives "05/01/2026 09:05:00" for style 0 and "‏05/‏01/‏2026" for style 2.
 *
 * Two deviations are left, and both are measured: a Hijri picture string separates its fields with spaces
 * where Windows writes "/", so the engine's ar-SA long date reads "16 رجب, 1447" against Windows's
 * "16/رجب/1447"; and a few names are CLDR's where Windows ships an older or a longer spelling (Arabic's
 * two-word month names, Latvian's capitalized "Pirmdiena", Vietnamese's "tháng 1", the Sotho and Tswana
 * month names). Thai is not one of them any more, and never was: a capture made in a loop reported error 5
 * for its date formats, and a fresh process, which is what a script gets, answers them.
 */
export function formatVbDateTime(date: Date, lcid: number, style: number, locale: string): string {
  const format = nlsDateFormat(lcid);
  if (!format) return formatVbDateViaIntl(date, locale);

  const { calendarType, am, pm } = format;
  const wallClock = onUtcTimeline(date);
  const marks = marksDateFields(locale);
  // Styles 0, 3 and 4 print the Gregorian date and an ordinary clock for every locale: ar-SA gives
  // "05/01/2026" where CStr gives the Hijri "16/07/47". Styles 1 and 2 are where Windows applies the
  // locale's own calendar, so there the date fields follow it.
  const gregorian: DateSource = GREGORIAN_DATE;
  const ownCalendar: DateSource = usesOwnCalendar(calendarType)
    ? { kind: 'calendar', type: calendarType }
    : GREGORIAN_DATE;
  const hijri = calendarType === HIJRI;
  // Thai is the one locale whose calendar reaches the styles that are Gregorian everywhere else: its year is
  // Buddhist even in CStr and GeneralDate. It is also the one locale whose long date writes no weekday, so
  // that picture string is expanded without the group holding it.
  const thai = calendarType === THAI_BUDDHIST;
  const longDatePicture = thai
    ? withoutWeekdayGroup(format.longDate)
    : hijri
      ? hijriLongDatePicture(format.longDate)
      : format.longDate;
  const longTime = expandNlsPattern(
    format.timeFormat,
    wallClock,
    locale,
    gregorian,
    am,
    pm,
    false,
    true
  );
  const longDate = expandNlsPattern(
    longDatePicture,
    wallClock,
    locale,
    ownCalendar,
    am,
    pm,
    marks,
    false,
    lcid,
    true
  );
  const generalDate = expandNlsPattern(
    format.shortDate,
    wallClock,
    locale,
    thai ? ownCalendar : gregorian,
    am,
    pm
  );
  const shortDate = expandNlsPattern(
    format.shortDate,
    wallClock,
    locale,
    ownCalendar,
    am,
    pm,
    marks,
    false,
    lcid
  );

  switch (style) {
    case VbDateTimeFormat.LongDate:
      // A long date is the one format whose whole point is the locale's own names, so a pattern the engine
      // cannot expand is reported rather than silently answered with an Intl layout that is a different
      // shape ("2026/1/5" for what should read "2026年1月5日"). No table row reaches this today.
      if (longDate === null) return unreachableLongDate(lcid);
      return longDate;
    case VbDateTimeFormat.ShortDate: {
      const written = shortDate ?? formatVbDateViaIntl(date, locale);
      // A Hijri date is written with two year digits where the same row's Gregorian short date asks for
      // four: ar-SA prints "16/07/47" and not "16/07/1447". The Persian year is four digits either way.
      if (!hijri) return written;
      return written.replace(/\d{4}(?![\s\S]*\d{4})/, match => match.slice(2));
    }
    case VbDateTimeFormat.LongTime:
      return longTime ?? formatVbDateViaIntl(date, locale);
    case VbDateTimeFormat.ShortTime: {
      const shortTime = expandNlsPattern(
        shortTimePicture(format.shortTime),
        wallClock,
        locale,
        gregorian,
        am,
        pm,
        false,
        true
      );
      return shortTime ?? longTime ?? formatVbDateViaIntl(date, locale);
    }
    default: {
      const timePart = longTime ?? formatVbDateViaIntl(date, locale);
      const datePart = generalDate ?? formatVbDateViaIntl(date, locale);
      // A Hijri locale keeps the date on day 0 (ar-SA prints "30/12/1899 09:05:00 ص"); every other locale
      // prints the time alone there, which is what CStr does everywhere. Thai was recorded here as an
      // exception and is not one: a fresh capture gives CStr(TimeSerial(9,5,0)) = "9:05:00".
      if (isDayZero(date) && !hijri) return timePart;
      if (isMidnight(date)) return datePart;
      return `${datePart} ${timePart}`;
    }
  }
}

/**
 * The Thai long date: Windows has no answer, so the caller gets the error it raises.
 */
function unreachableLongDate(lcid: number): never {
  throw createVbError(
    VbErrorCodes.InvalidProcedureCall,
    `FormatDateTime: the locale ${lcid} has no long date format`
  );
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
  const longInDateFormat = new Intl.DateTimeFormat(locale, {
    year: 'numeric',
    month: 'long',
    day: 'numeric',
    timeZone: 'UTC',
  });
  const shortInDateFormat = new Intl.DateTimeFormat(locale, {
    year: 'numeric',
    month: 'short',
    day: 'numeric',
    timeZone: 'UTC',
  });
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
    if (best === null || name.length > best.length)
      best = { month: entry.month, length: name.length };
  }
  return best === null ? null : best.month;
}

// ---------------------------------------------------------------------------
// Reading a date or a time
// ---------------------------------------------------------------------------

type DateOrder = 'DMY' | 'MDY' | 'YMD';

function dateOrderOf(locale: string): { order: DateOrder; separator: string } {
  const parts = new Intl.DateTimeFormat(locale, {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
  }).formatToParts(new Date(2026, 8, 29));
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
    const [day, month] =
      order === 'DMY' ? [Number(two[1]), Number(two[3])] : [Number(two[3]), Number(two[1])];
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
      candidates.push(
        [yearOf(third), Number(first), Number(second)],
        [yearOf(third), Number(second), Number(first)]
      );
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
