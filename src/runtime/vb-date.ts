// VBScript dates are OLE Automation dates: a day count from 30 December 1899, with the time of day as
// the fraction. A Date here carries the same instant as a local wall-clock time.

import { nlsDateFormat } from './nls-date-table.generated.ts';
import { NLS_MONTH_NAME_READS } from './nls-month-name-reads.generated.ts';
import { ICU_MONTH_NAME_READS } from './icu-month-name-reads.generated.ts';
import { NLS_COMMA_BEFORE_TIME } from './nls-comma-time.generated.ts';
import {
  HIJRI_FIRST_YEAR,
  HIJRI_YEAR_MONTH_MASKS,
  HIJRI_YEAR_START_SERIALS,
} from './hijri-table.generated.ts';
import { parseVbNumericString } from './nls-number.ts';
import { NAME_IN_DATE, NAME_STANDALONE } from './name-tables.generated.ts';
import { createVbError, VbErrorCodes } from './errors.ts';
import { currentLcid, getCurrentBCP47Locale } from '../builtins/locale.ts';

const MS_PER_DAY = 86400000;
const OLE_EPOCH_UTC = Date.UTC(1899, 11, 30);

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
/**
 * The OLE Automation number of a date: whole days since 30 December 1899, time of day as the fraction.
 *
 * A date before that epoch counts its time of day away from zero rather than toward it, which a fresh cscript
 * shows: CDbl(CDate("12/29/1899 12:00:00")) is -1.5 rather than -0.5, and the day alone is -1.
 */
export function dateToSerial(date: Date): number {
  const dayCount = Math.round(
    (Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()) - OLE_EPOCH_UTC) / MS_PER_DAY
  );
  const timeOfDay =
    (date.getHours() * 3600 + date.getMinutes() * 60 + date.getSeconds()) / 86400 +
    date.getMilliseconds() / MS_PER_DAY;
  return dayCount < 0 && timeOfDay > 0 ? dayCount - timeOfDay : dayCount + timeOfDay;
}

/**
 * The date an OLE Automation number stands for, as local wall-clock time.
 *
 * The whole days are truncated toward zero and the fraction is read as time after them, which is what a
 * negative number means: CDate(-1.5) is 29 December 1899 at noon and CDate(-0.5) is 30 December 1899 at noon,
 * both measured on a fresh cscript.
 */
export function serialToDate(serial: number): Date {
  const wholeDays = Math.trunc(serial);
  const timeOfDay = Math.abs(serial - wholeDays);
  const wallClock = new Date(OLE_EPOCH_UTC + Math.round((wholeDays + timeOfDay) * MS_PER_DAY));
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
  if (source.type === THAI_BUDDHIST) {
    // The Buddhist year is the Gregorian one plus 543, and the months and days are the Gregorian ones. ICU
    // is not asked, because its Buddhist calendar counts in the Julian calendar before 1582: it writes the
    // Gregorian 5 January 1483 as 27 December 2025, where Windows writes 5 January 2026.
    return {
      day: date.getUTCDate(),
      month: date.getUTCMonth() + 1,
      year: date.getUTCFullYear() + BUDDHIST_YEAR_OFFSET,
    };
  }
  if (source.type === HIJRI) {
    // The Umm al-Qura table, not ICU: measured over the years 1900 to 2100, ICU's islamic-umalqura calendar
    // is a day behind Windows' for about a fifth of the days (15 June 1929 is 8 Muharram 1348 to Windows and
    // 7 Muharram 1348 to ICU), while .NET's UmAlQuraCalendar agrees with Windows day for day. A day the table
    // does not cover is clamped to the nearest one it has; the date styles refuse such a day before this.
    return hijriFieldsOf(clampedHijriDay(wallClockDayNumber(date))) as {
      day: number;
      month: number;
      year: number;
    };
  }
  if (source.type === PERSIAN) {
    return icuCalendarFields(persianFieldDate(date, source, locale), source, locale);
  }
  return icuCalendarFields(date, source, locale);
}

/** The fields ICU answers for a date under a calendar, which the Hijri calendar does not use. */
function icuCalendarFields(
  date: Date,
  source: DateSource,
  locale: string
): { day: number; month: number; year: number } {
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

/**
 * The day whose ICU fields are a Persian date's own.
 *
 * ICU's Persian calendar runs one day ahead of Windows' for the years before 980. Measured on 1065:
 * 1601-01-01 is 11 Dey 979 to Windows and 12 Dey 979 to ICU, and every day of that tail is one ahead, while
 * the two agree again from 1601-03-21, which is 1 Farvardin 980. Windows writes no date before 1601-01-01,
 * so the last months of 979 are the only place the two can be seen apart.
 */
function persianFieldDate(date: Date, source: DateSource, locale: string): Date {
  if (icuCalendarFields(date, source, locale).year <= 979) {
    return new Date(date.getTime() - MS_PER_DAY);
  }
  return date;
}

/** A Gregorian day number brought inside the days the Umm al-Qura table covers. */
function clampedHijriDay(dayNumber: number): number {
  return Math.min(Math.max(dayNumber, HIJRI_YEAR_START_SERIALS[0]!), hijriLastDayNumber());
}

/** The first day of a Hijri month, as a date on the UTC timeline, or null for a year the table has not. */
function hijriMonthStartDate(year: number, month: number): Date | null {
  const index = year - HIJRI_FIRST_YEAR;
  if (index < 0 || index >= HIJRI_YEAR_START_SERIALS.length) return null;
  let day = HIJRI_YEAR_START_SERIALS[index]!;
  const mask = HIJRI_YEAR_MONTH_MASKS[index]!;
  for (let earlier = 0; earlier < month - 1; earlier++) {
    day += ((mask >> earlier) & 1) === 1 ? 30 : 29;
  }
  return new Date(OLE_EPOCH_UTC + day * MS_PER_DAY);
}

/**
 * The day a month name should be read from, which is not always the day itself.
 *
 * The engine's own fields for a Hijri date come from the Umm al-Qura table and for a Persian date before 980
 * from the day before, so asking ICU about the date itself can name a month the picture is not writing:
 * 30 December 1899 is before the Hijri table, and ICU falls back there to a calendar that calls it Sha'ban
 * while the fields say Muharram. The name of a Hijri month is read from that month's first day instead.
 */
function calendarNameDate(date: Date, source: DateSource, locale: string): Date {
  if (source.kind !== 'calendar') return date;
  if (source.type === HIJRI) {
    const fields = hijriFieldsOf(clampedHijriDay(wallClockDayNumber(date)));
    if (!fields) return date;
    return hijriMonthStartDate(fields.year, fields.month) ?? date;
  }
  if (source.type === PERSIAN) return persianFieldDate(date, source, locale);
  return date;
}

/** The number a year token writes. */
function yearValue(source: DateSource, date: Date, locale: string): number {
  return dateFields(date, source, locale).year;
}

/**
 * The whole days from the OLE epoch to the calendar day a wall-clock date names.
 *
 * A date here is a wall-clock reading carried onto the UTC timeline, so its UTC fields are the day the script
 * means, which is what the date writers count from.
 */
function wallClockDayNumber(date: Date): number {
  return Math.round(
    (Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()) - OLE_EPOCH_UTC) /
      MS_PER_DAY
  );
}

/** The last day the Umm al-Qura table covers, as a Gregorian day number. */
function hijriLastDayNumber(): number {
  const lastIndex = HIJRI_YEAR_START_SERIALS.length - 1;
  const mask = HIJRI_YEAR_MONTH_MASKS[lastIndex]!;
  let last = HIJRI_YEAR_START_SERIALS[lastIndex]! - 1;
  for (let month = 0; month < 12; month++) last += ((mask >> month) & 1) === 1 ? 30 : 29;
  return last;
}

/**
 * The Umm al-Qura fields of a Gregorian day number, or null for a day the table does not cover.
 *
 * The table is a start day and twelve month lengths a year, so the fields of any day inside it come from
 * walking the months of the one year whose start is the last at or before that day.
 */
function hijriFieldsOf(dayNumber: number): { day: number; month: number; year: number } | null {
  if (dayNumber < HIJRI_YEAR_START_SERIALS[0]! || dayNumber > hijriLastDayNumber()) return null;

  let low = 0;
  let high = HIJRI_YEAR_START_SERIALS.length - 1;
  while (low < high) {
    const middle = Math.ceil((low + high) / 2);
    if (HIJRI_YEAR_START_SERIALS[middle]! <= dayNumber) low = middle;
    else high = middle - 1;
  }

  let remaining = dayNumber - HIJRI_YEAR_START_SERIALS[low]!;
  const mask = HIJRI_YEAR_MONTH_MASKS[low]!;
  for (let month = 0; month < 12; month++) {
    const length = ((mask >> month) & 1) === 1 ? 30 : 29;
    if (remaining < length) {
      return { day: remaining + 1, month: month + 1, year: HIJRI_FIRST_YEAR + low };
    }
    remaining -= length;
  }
  return null;
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
      // Windows writes the digits a year has, so year 100 is "100" and not "0100": under 1031,
      // CDate("1/1/100") is written "01.01.100" by a fresh cscript.
      return String(year);
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
  // The month's name is read from the day the fields themselves name, which for a Hijri date is a day of
  // that month in the Umm al-Qura table and for a Persian date before 980 is the day before.
  const nameDate = calendarNameDate(date, year, locale);

  const hour = date.getUTCHours();
  const meridiem = hour < 12 ? am || 'AM' : pm || 'PM';

  switch (token) {
    case 'MMM':
      return name(monthIndex, overrides?.months?.[monthIndex], () =>
        partOf(formatters.monthShort, nameDate, 'month')
      );
    case 'MMMM':
      return name(monthIndex, overrides?.months?.[monthIndex], () =>
        partOf(formatters.monthLong, nameDate, 'month')
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
      // Windows writes the digits a year has, so year 100 is "100" and not "0100": under 1031,
      // CDate("1/1/100") is written "01.01.100" by a fresh cscript.
      return String(yearValue(year, date, locale));
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
  persianDayUnpadded = false,
  widenTwoDigitYear = false
): string | null {
  // Windows writes a date to the second, rounding the fraction of a second rather than dropping it: the
  // serial 46028.2345 is 5:37:41 and not the 5:37:40 a truncation gives, measured on a fresh cscript.
  const milliseconds = date.getUTCMilliseconds();
  if (milliseconds !== 0) {
    date = new Date(date.getTime() + (milliseconds >= 500 ? 1000 - milliseconds : -milliseconds));
  }
  const hour = date.getUTCHours();
  const hour12 = hour % 12 === 0 ? 12 : hour % 12;
  const pad = (value: number, width: number): string => String(value).padStart(width, '0');
  // A two digit year is written with four digits outside 1950 to 2049, which GeneralDate and CStr do while
  // the short date does not: measured on 1059 and 2060, whose short dates are the only ones in the table
  // that ask for two digits, GeneralDate gives "15.06.1949" and "15.06.2050" but "15.06.50" for 1950, and
  // the short date gives "15.06.49" for 1949. The window is the machine's own: its Gregorian calendar has
  // a two digit year maximum of 2049, and Windows writes two digits for that year and the 99 before it.
  const calendarYear = yearValue(year, date, locale);
  const wideYear = calendarYear < 1950 || calendarYear > 2049;
  const yearShort = String(calendarYear % 100);
  const yearTwoDigits =
    widenTwoDigitYear && wideYear ? String(calendarYear) : pad(calendarYear % 100, 2);
  const fields: Record<string, string> = {
    d: dateNumericToken('d', date, year, locale) as string,
    dd: dateNumericToken('dd', date, year, locale, persianDayUnpadded) as string,
    M: dateNumericToken('M', date, year, locale) as string,
    MM: dateNumericToken('MM', date, year, locale) as string,
    y: widenTwoDigitYear && wideYear ? String(calendarYear) : yearShort,
    yy: yearTwoDigits,
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
 * CStr does not mark its fields, measured on all 132 locales: ar-SA prints "05/01/2026 09:05:00 ص" with no
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
      lcid,
      false,
      true
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
 * 2026-01-05 15:07:09, a time that is only a time, dates from 1000 to 4000 and both day-0 instants, for all
 * 132 locales the table covers):
 *   * GeneralDate is the short date with the LOCALE_STIMEFORMAT clock. A value that is only a time prints
 *     the time alone in every locale, ar-SA included, and one whose clock is midnight prints the date alone.
 *   * LongDate is LOCALE_SLONGDATE, so its weekday and month names are the long ones.
 *   * ShortTime is the table's LOCALE_SSHORTTIME with the AM/PM designator dropped and a 24-hour clock,
 *     NOT a hard-coded "HH:mm": Finnish prints "09.05" and Indonesian "09.05", keeping their separator.
 *   * The two date formats use the locale's own calendar - ar-SA's long date is the Hijri "16/رجب/1447" and
 *     its short date the Hijri "16/07/47", fa-IR's the Persian "15 دی 1404" - while GeneralDate and both
 *     clocks stay Gregorian for every locale.
 *   * The short date that GeneralDate prints carries no right-to-left mark, while the one ShortDate prints
 *     does: he-IL gives "05/01/2026 09:05:00" for style 0 and "‏05/‏01/‏2026" for style 2.
 *   * Thai is not a special case: it writes its Buddhist year in every style, CStr included ("5/1/2569"),
 *     and its long date carries no weekday, which is what Windows answers.
 *   * The two date formats are error 5 for a day Windows' own date writers cannot write: anything before
 *     1601-01-01 in every locale, and for ar-SA anything outside the Hijri range 1318-01-01 to 1500-12-30,
 *     which the calendar itself covers and GeneralDate does not care about.
 *
 * Every one of the 141 measured cells agrees with vbscript.dll for each of the 132 locales, as do the 19
 * long-date samples a locale. The two readings that still differ are in the parser, not the writer, and are
 * written down beside parseVbDate.
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
  // A date Windows' own date writers cannot write is error 5, measured on a fresh cscript over every locale:
  // a day before 1601 (1033 gives error 5 for the long and short dates of 1483 but a plain date for its
  // GeneralDate), and for ar-SA a day outside the Hijri range the calendar covers, 1318-01-01 to 1500-12-30,
  // which is 30 April 1900 to 16 November 2077.
  if (style === VbDateTimeFormat.LongDate || style === VbDateTimeFormat.ShortDate) {
    assertDateStyleWritable(wallClock, locale, hijri);
  }
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
    pm,
    false,
    false,
    0,
    false,
    true
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
      // A value that is only a time prints the time alone, in every locale, ar-SA included: a fresh cscript
      // answers FormatDateTime(CDate(0)) with "12:00:00 ص" for ar-SA and with "12:00:00 AM" for 1033, and
      // FormatDateTime(TimeSerial(13, 5, 9)) with the clock alone in both. A whole day at midnight prints the
      // date alone instead.
      if (isDayZero(date)) return timePart;
      if (isMidnight(date)) return datePart;
      return `${datePart} ${timePart}`;
    }
  }
}

/** The first day Windows' own date writers accept, which is 1601-01-01 for every locale. */
const FIRST_WRITABLE_DAY = Date.UTC(1601, 0, 1);

/**
 * Refuses a date the long or the short date cannot write, with the error Windows raises.
 *
 * Measured on a fresh cscript over all 132 locales: the long and the short date of 15 June 1483 are error 5
 * under 1033 while its GeneralDate is "6/15/1483", and the same split holds for the Persian and Thai years,
 * whose own calendars start at the same 1601-01-01. The Hijri range is its own calendar's and was measured
 * day by day: 30 April 1900 is "01/01/1318" and 29 April 1900 is error 5, 16 November 2077 is "30/12/1500"
 * and 17 November 2077 is error 5.
 *
 * The Hebrew locale is the one that answers a different number. It is 32812, "the date is not in the
 * calendar of the current locale", where every other locale answers 5, measured for 1400, 1500 and 1600
 * alike; its GeneralDate and CStr answer normally for all of them.
 */
function assertDateStyleWritable(date: Date, locale: string, hijri: boolean): void {
  const refused = (): never => {
    const number = locale.toLowerCase().startsWith('he')
      ? 32812
      : VbErrorCodes.InvalidProcedureCall;
    throw createVbError(number, 'This date has no long or short form in this locale');
  };
  if (date.getTime() < FIRST_WRITABLE_DAY) refused();
  if (!hijri) return;
  if (hijriFieldsOf(wallClockDayNumber(date)) === null) refused();
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

function bareMonthNames(locale: string): { month: number; name: string; short: boolean }[] {
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

  const names: { month: number; name: string; short: boolean }[] = [];
  for (let month = 1; month <= 12; month++) {
    const day = new Date(Date.UTC(2026, month - 1, 15));
    names.push(
      { month, name: long.format(day), short: false },
      { month, name: short.format(day), short: true },
      { month, name: named(longInDateFormat, day), short: false },
      { month, name: named(shortInDateFormat, day), short: true }
    );
  }
  return names.filter(entry => entry.name !== '');
}

function monthTable(locale: string): { month: number; name: string }[] {
  const lcid = currentLcid();
  const cacheKey = `${lcid}|${locale}`;
  const cached = monthTableCache.get(cacheKey);
  if (cached) return cached;
  const policy = MONTH_NAME_POLICY[locale.toLowerCase().split(/[-_]/)[0]] ?? 'own';
  const english = policy === 'own' ? [] : bareMonthNames('en-US');
  const ownNames = policy === 'english' ? [] : bareMonthNames(locale);
  const windowsNames = policy === 'english' ? [] : standaloneNamesFromWindows(lcid);
  // Windows writes its own spelling where CLDR has changed one, and its parser reads that spelling back:
  // under 1031 "5. Mrz 2026" is March, while CLDR's short name for March is "Mär". A locale that reads only
  // English names keeps none of its own, which is the measured rule for Chinese, Japanese and Korean: their
  // "一月", "1月" and "1월" are refused by a fresh cscript, and the standalone names Windows ships are their
  // own, so adding them here would read a text Windows does not.
  //
  // Which of a locale's own names are read back is measured twice, because the two spellings differ and
  // Windows reads one where it refuses the other. Yiddish and Tamil read no name at all, Hindi only May,
  // Tsonga everything but December, the Vietnamese long names ("tháng 1" from ICU, "Tháng Giêng" from
  // Windows) are refused where its short ones are read, and Greek reads the names Windows writes and not
  // ICU's long ones.
  //
  // A name that is only digits is left out: Intl answers a bare number for the short form of a few locales
  // (Slovak's is "9" for September), and a fresh cscript reads no such text as a name, so "29.9.26" is error
  // 13 under 1051 rather than 29 September. The names that carry a marker do stay, since "1月" is a name.
  const readable = (masks: number | undefined) => (entry: { month: number; short: boolean }) => {
    if (masks === undefined) return true;
    const bits = entry.short ? masks & 0xfff : (masks >> 12) & 0xfff;
    return (bits & (1 << (entry.month - 1))) !== 0;
  };
  const usable = [
    ...ownNames.filter(readable(ICU_MONTH_NAME_READS.get(lcid))),
    ...windowsNames.filter(readable(NLS_MONTH_NAME_READS.get(lcid))),
    ...english,
  ].filter(entry => /\p{L}/u.test(entry.name));
  monthTableCache.set(cacheKey, usable);
  return usable;
}

/** The names Windows writes for a month on its own, which its parser reads as well as CLDR's. */
function standaloneNamesFromWindows(
  lcid: number
): { month: number; name: string; short: boolean }[] {
  const row = NAME_STANDALONE[lcid];
  if (!row) return [];
  const names: { month: number; name: string; short: boolean }[] = [];
  for (const [monthKey, name] of Object.entries(row.monthsLong ?? {})) {
    names.push({ month: Number(monthKey) + 1, name, short: false });
  }
  for (const [monthKey, name] of Object.entries(row.monthsShort ?? {})) {
    names.push({ month: Number(monthKey) + 1, name, short: true });
  }
  return names;
}

/**
 * The month a word names, or null. vbscript.dll matches three letters or more as a case-insensitive prefix
 * of one of the locale's month names ("Ja" is no month, "Janu" and "Januar" are January in English), so the
 * longest name that starts with the word decides. A word that *is* a name matches whatever its length, which
 * a fresh cscript shows for the short names shorter than three letters: Catalan "15 ag. 2026" is August,
 * Arabic "15 آب 2026" is August, Persian "15 مه 2026" is May, and Czech "15 červen 2026" is June and not the
 * July the longer "červenec" would give by the prefix rule alone.
 */
function monthOfWord(word: string, locale: string): number | null {
  const bare = word.toLowerCase().replace(/\.+$/, '');
  if (bare === '') return null;
  let best: { month: number; length: number } | null = null;
  for (const entry of monthTable(locale)) {
    const name = entry.name.toLowerCase().replace(/\.+$/, '');
    if (name === bare) return entry.month;
    if (bare.length < 3 || !name.startsWith(bare)) continue;
    if (best === null || name.length > best.length)
      best = { month: entry.month, length: name.length };
  }
  return best === null ? null : best.month;
}

/**
 * The first complete month name a text holds, with where it sits, or null.
 *
 * A name a locale writes with a dot or a space in it cannot survive the split into words below: Thai writes
 * "ม.ค." and Vietnamese "thg 1", and the text "15 ม.ค. 2026" becomes four words and is refused, where a
 * fresh cscript reads both. Only such a name is looked for here, so a name of one word is left to the rules
 * above and a text Windows refuses is not read on the strength of this. A name counts only as a whole word,
 * which is what keeps "1 Janu 2026" to the prefix rule rather than reading the "Jan" inside "Janu". The
 * longest name the text holds wins, any dots right after it belong to it, and the match ends at a separator
 * or at the end of the text.
 */
function findCompleteMonthName(
  text: string,
  locale: string
): { month: number; start: number; end: number } | null {
  const lower = text.toLowerCase();
  const isSeparator = (char: string | undefined): boolean =>
    char === undefined || /[\s,./-]/.test(char);
  let best: { month: number; start: number; end: number; length: number } | null = null;

  for (const entry of monthTable(locale)) {
    if (!/[\s.]/.test(entry.name)) continue;
    for (const candidate of [entry.name, entry.name.replace(/\.+$/, '')]) {
      if (candidate === '') continue;
      const needle = candidate.toLowerCase();
      for (let at = lower.indexOf(needle); at >= 0; at = lower.indexOf(needle, at + 1)) {
        if (!isSeparator(text[at - 1])) continue;
        let end = at + needle.length;
        while (text[end] === '.') end++;
        if (!isSeparator(text[end])) continue;
        if (best === null || end - at > best.length) {
          best = { month: entry.month, start: at, end, length: end - at };
        }
      }
    }
  }
  return best === null ? null : { month: best.month, start: best.start, end: best.end };
}

// ---------------------------------------------------------------------------
// Reading a date or a time
// ---------------------------------------------------------------------------

type DateOrder = 'DMY' | 'MDY' | 'YMD';

/** A date that was read, with whether the text named a year of its own. */
interface WrittenDate {
  date: Date;
  yearWritten: boolean;
}

/** A date the reader found, tagged with whether the text named a year. */
function written(date: Date | null, yearWritten: boolean): WrittenDate | null {
  return date === null ? null : { date, yearWritten };
}

/**
 * The field order, the separator and the dotted-date habit a locale's own short date uses, all read from the
 * Windows picture string.
 *
 * The order decides which of two numbers in a written date is the month, and Windows reads a two field date
 * as day and month where its picture is day first and as month and day everywhere else. Intl answers the
 * order differently from Windows in five of the 132 locales: 1065 (fa-IR) and 4100 (zh-SG) are day first for
 * Windows and year first for Intl, 2077 (sv-FI) is year first for Windows and day first for Intl, 9225
 * (en-CB) is day first for Windows and month first for Intl, and 1077 (zu-ZA) is month first for Windows and
 * year first for Intl. Asking Intl therefore reads CDate("1,5") as the other date under all five.
 *
 * The picture is walked rather than searched, so a quoted literal is never mistaken for a field, and a
 * picture that separates its fields with a word (1028 writes "yyyy/M/d") keeps that word out of the reading.
 *
 * A dot only counts as a field mark where the picture writes it against the field. 1038 (hu-HU), 1051
 * (sk-SK) and 1060 (sl-SI) write ". " instead, and Windows reads no dotted date under them: "1.5" is the
 * time 1:05 there, "1.5.26" is 1:05:26, "29.9" is the day count 299 under 1060 and a type mismatch under the
 * other two. Only their own spaced form is a date ("2026. 05. 01." under 1038, "1. 5. 2026" under 1051).
 */
function dateOrderOf(): {
  order: DateOrder;
  separator: string;
  dottedDates: boolean;
  trailingDot: boolean;
} {
  const picture = nlsDateFormat(currentLcid())?.shortDate ?? '';
  const fields: string[] = [];
  let separator = '/';
  let separatorSeen = false;
  let dottedDates = false;
  for (let index = 0; index < picture.length;) {
    const character = picture[index]!;
    if (character === "'") {
      const end = picture.indexOf("'", index + 1);
      index = end < 0 ? picture.length : end + 1;
      continue;
    }
    let run = 1;
    while (picture[index + run] === character) run++;
    const token = character.repeat(run);
    if (token === 'd' || token === 'dd') fields.push('D');
    else if (token === 'M' || token === 'MM') fields.push('M');
    else if (/^y+$/.test(token)) fields.push('Y');
    else if (!separatorSeen && fields.length > 0 && !/[A-Za-z]/.test(token)) {
      separator = token;
      separatorSeen = true;
      dottedDates = token === '.' && picture[index + run] !== ' ';
    }
    index += run;
  }
  const order = fields.join('');
  // A picture that ends with a dot writes that dot into a date it prints, and Windows reads it back only
  // there: 1038 writes "yyyy. MM. dd." and reads CDate("2026. 05. 01."), while 1031, 1051 and 1060 write no
  // trailing dot and answer a type mismatch for the same text.
  return {
    order: order === 'MDY' || order === 'YMD' ? order : 'DMY',
    separator,
    dottedDates,
    trailingDot: /\.\s*$/.test(picture),
  };
}

/**
 * A year under a hundred, read as the year it stands for.
 *
 * The window runs from 76 years before the current year to 23 years after it, which a fresh cscript shows
 * under 1031 and 1033 in 2026: 00 to 49 are 2000 to 2049 and 50 to 99 are 1950 to 1999. Under a calendar that
 * counts from another epoch the same window sits around that calendar's own current year, measured under 1054
 * in 2026, whose calendar is the Buddhist one and whose year is 2569: "5. 1. 00" is 2500, "5. 1. 92" is 2592
 * and "5. 1. 93" is 2493.
 *
 * It reads the VALUE and not the digits written, so CDate("0001-01-01") is 1 January 2001 while
 * CDate("1/1/100") is year 100.
 */
function toFullYear(year: number): number {
  // A value outside the one and two digit range is no year field, and keeps the reading it always had rather
  // than being folded into the window: 65535 wraps to -1, which a fresh cscript under 1031 reads as 1999 in
  // CDate("65535.5"), while the wrapped -31072 of CDate("100000.5") stays out of range so that reading fails
  // and the marked number 1000005, the day count Windows reads there, gets its turn.
  if (year < 0 || year > 99) return year < 50 ? 2000 + year : 1900 + year;
  const base = currentYearInItsOwnCalendar() - 76;
  return base + ((((year - base) % 100) + 100) % 100);
}

/** The current year as the locale's own calendar counts it, which a Buddhist calendar counts as 2569 where
 * the Gregorian one counts 2026. */
function currentYearInItsOwnCalendar(): number {
  const parts = new Intl.DateTimeFormat(getCurrentBCP47Locale(), { year: 'numeric' }).formatToParts(
    new Date()
  );
  const ownYear = Number(parts.find(part => part.type === 'year')?.value);
  return Number.isFinite(ownYear) && ownYear > 0 ? ownYear : new Date().getFullYear();
}

/** A year field: a value under a hundred is a year of the near past or present, as vbscript.dll reads it. */
function yearOf(field: string): number {
  return yearOfNumber(Number(field));
}

/** The same reading for a year that was already turned into a number. */
function yearOfNumber(value: number): number {
  return value < 100 ? toFullYear(value) : value;
}

/**
 * A number read as a date field, which the original reader holds in sixteen signed bits.
 *
 * This is a limit of the reader rather than a rule about dates, and a fresh cscript under 1031 shows it
 * plainly: CDate("2958465.5") is 1 May 9345, because 2958465 wraps to 9345, and CDate("65535.5") is 1 May
 * 1999, because 65535 wraps to -1 and a year under a hundred is read as 1999. 65536.5, 131073.5 and 196609.5
 * are 1 May 2000, 2001 and 2001 for the same reason, 70000.5 is 1 May 4464, and CDate("2958465.25") is a
 * type mismatch because the wrapped month is 25.
 */
function wrapToDateField(value: number): number {
  const wrapped = ((value % 65536) + 65536) % 65536;
  return wrapped >= 32768 ? wrapped - 65536 : wrapped;
}

function validDate(year: number, month: number, day: number): Date | null {
  // OLE Automation dates run from 1 January 100 to 31 December 9999; a year outside that is no date.
  if (year < 100 || year > 9999) return null;
  const date = new Date(year, month - 1, day);
  date.setFullYear(year);
  return date.getMonth() === month - 1 && date.getDate() === day ? date : null;
}

/** A date written with a month name: "1 Jan 2026", "Jan 1, 2026", "März 1 2026", "Jan 2026", "1 Jan". */
function parseMonthNameDate(text: string, locale: string): WrittenDate | null {
  // A month name is set off from the rest by a space, a comma, a slash, a hyphen or the locale's own date
  // mark. A dot counts only where the locale writes its dates with a dot, which a fresh cscript shows: under
  // 1033 "5. Januar 2026" is a type mismatch while "Januar. 5, 2026" reads, because the dot there belongs to
  // the name, and under 1031 both read. 1060 reads the dotted form too, and 1042 refuses it because its own
  // dates are written with hyphens. 1054 reads it as well, which is measured and comes with the rest of that
  // locale's reading of dotted dates.
  const ownMark = dateOrderOf().separator;
  const buddhistDates = nlsDateFormat(currentLcid())?.calendarType === THAI_BUDDHIST;
  const separators = ownMark === '.' || buddhistDates ? /[\s,./-]+/ : /[\s,/-]+/;
  // A name with a dot or a space in it is taken out of the text first, so the split cannot break it in two.
  const complete = findCompleteMonthName(text, locale);
  const rest =
    complete === null ? text : text.slice(0, complete.start) + ' ' + text.slice(complete.end);
  const words = rest.split(separators).filter(word => word !== '');

  let month: number;
  let numbers: string[];
  if (complete !== null) {
    if (words.length < 1 || words.length > 2) return null;
    if (!words.every(word => /^\d+$/.test(word))) return null;
    month = complete.month;
    numbers = words;
  } else {
    if (words.length < 2 || words.length > 3) return null;
    const months = words.map(word => monthOfWord(word, locale));
    const named = months.filter(entry => entry !== null).length;
    if (named !== 1) return null;
    month = months[months.findIndex(entry => entry !== null)] as number;
    numbers = words.filter((_, index) => months[index] === null);
    if (!/^\d+$/.test(numbers.join(''))) return null;
  }

  // "Jan 2026" is January 1st of that year; "1 Jan" is the 1st of January this year.
  if (numbers.length === 1) {
    const [only] = numbers;
    if (only.length === 4 || Number(only) > 31)
      return written(validDate(yearOf(only), month, 1), true);
    return written(validDate(new Date().getFullYear(), month, Number(only)), false);
  }

  // A four-digit number names the year wherever it stands ("1 Jan 2026"). With two short numbers the
  // locale's own order decides: a year-first locale reads "1 Jan 26" as 2001-01-26, the others as
  // 1 January 2026.
  const [first, second] = numbers;
  const fourDigit = numbers.find(number => number.length === 4);
  if (fourDigit !== undefined) {
    const other = numbers.find(number => number !== fourDigit) as string;
    return written(validDate(Number(fourDigit), month, Number(other)), true);
  }
  const { order } = dateOrderOf();
  const [year, day] = order === 'YMD' ? [first, second] : [second, first];
  return written(validDate(yearOf(year), month, Number(day)), true);
}

/**
 * A date written with the locale's own markers instead of separators: "2026年9月29日" in Chinese and
 * Japanese, "2026년 9월 29일" in Korean. The layout comes from the locale itself, so only the markers it
 * actually uses are accepted, and a part that is not digits-then-marker bails out and leaves the text to
 * the other readers.
 */
function parseLocalizedDate(text: string, locale: string): WrittenDate | null {
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
  if (own) return written(own, true);
  // The ideographic form as well as the locale's own markers: vbscript.dll reads "2026年9月29日" under
  // Korean too, even though Korean writes its own dates as "2026년 9월 29일". Measured on 2052, 1028,
  // 1041 and 1042; en-US and de-DE reject both forms.
  return written(read(/^(\d{1,4})年(\d{1,2})月(\d{1,2})日$/, ['year', 'month', 'day']), true);
}

function parseDatePart(text: string, locale: string): WrittenDate | null {
  const named = parseMonthNameDate(text, locale);
  if (named) return named;

  const localized = parseLocalizedDate(text, locale);
  if (localized) return localized;

  const iso = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(text);
  // The year goes through the same reading as any other year field, so a written year under a hundred is a
  // year of the near past or present: CDate("0001-01-01") is 1 January 2001, as a fresh cscript answers.
  if (iso) return written(validDate(yearOf(iso[1]), Number(iso[2]), Number(iso[3])), true);

  const { order, separator, dottedDates, trailingDot } = dateOrderOf();
  // "/" and "-" are read everywhere, a comma is read too, and the locale's own field mark is read as well.
  // The locale's own mark holds the text to the locale's own order instead of letting the other order
  // rescue it ("29.9.26" is no date in Hungarian). A dot is that mark only where the locale's picture puts
  // it against the field, which is what keeps a dotted text out of the date reader under 1038, 1051 and 1060.
  const ownMark = separator === '/' || separator === '-' ? null : separator;
  const own = ownMark === '.' && !dottedDates ? null : ownMark;
  const separatorAllowed = (mark: string): boolean =>
    mark === '/' || mark === '-' || mark === ',' || mark === own;

  // A dotted date written with the spaces its own locale puts in the picture, which Windows reads where the
  // locale writes its dates with a dot, and under 1054 as well: measured, CDate("2026. 05. 01.") is 1 May
  // 2026 under 1038 and CDate("1. 5. 2026") is a date under 1031, 1038, 1051, 1054 and 1060, while 1033,
  // 1041, 1042, 2052 and 2057 answer a type mismatch or read it as a time. The trailing dot is read only
  // under 1038, whose own picture ends with one, and under 1054. Why Thai takes both readings is not known,
  // only that it does.
  const buddhistDates = nlsDateFormat(currentLcid())?.calendarType === THAI_BUDDHIST;
  if (ownMark === '.' || buddhistDates) {
    const spaced = /^(\d{1,4})\.\s+(\d{1,2})\.\s+(\d{1,4})(\.?)$/.exec(text);
    if (spaced && (spaced[4] === '' || trailingDot || buddhistDates)) {
      const read = threeFieldDate(order, spaced[1]!, spaced[2]!, spaced[3]!);
      if (read) return read;
    }
    // 1054 reads a two field spaced form as a date as well, which its own picture does not write: measured,
    // CDate("1. 5.") is 1 May of the current year there and CDate("2026. 05.") is May of the Buddhist year
    // 2026, where 1033 reads the same texts as the times 1:05 and 20:26.
    if (buddhistDates) {
      const spacedTwo = /^(\d{1,9})\.\s+(\d{1,9})\.?$/.exec(text);
      if (spacedTwo) {
        const read = twoFieldDate(order, spacedTwo[1]!, spacedTwo[2]!);
        if (read) return read;
      }
    }
  }

  // A date written with spaces between its numbers is read the same way as one written with the locale's own
  // mark, which a fresh cscript shows under 1033 and 1025 alike: "15 1 26" is 2015-1-26, "26 9 29" is
  // 2026-9-29 and "15 1 2026" is 2026-1-15, exactly as the slashed forms of the same numbers are.
  const spacedNumbers = /^(\d{1,9})[\s,]+(\d{1,9})[\s,]+(\d{1,9})$/.exec(text);
  if (spacedNumbers) {
    return threeFieldDate(order, spacedNumbers[1]!, spacedNumbers[2]!, spacedNumbers[3]!);
  }
  const spacedPair = /^(\d{1,9})[\s,]+(\d{1,9})$/.exec(text);
  if (spacedPair) {
    return twoFieldDate(order, spacedPair[1]!, spacedPair[2]!);
  }

  // A field may be written with more digits than a date can hold, and the reader wraps it: see
  // wrapToDateField. The time and the marked number are read before and after this, so a text that this
  // does not make a date still has its other readings.
  const two = /^(\d{1,9})([.,/-])(\d{1,9})$/.exec(text);
  if (two && separatorAllowed(two[2])) {
    return twoFieldDate(order, two[1]!, two[3]!);
  }

  const match = /^(\d{1,4})([.,/-])(\d{1,2})\2(\d{1,4})$/.exec(text);
  if (!match || !separatorAllowed(match[2])) return null;

  const [first, second, third] = [match[1], match[3], match[4]];
  return threeFieldDate(order, first, second, third);
}

/**
 * The date a written two field date names, read as a day and a month first, the locale's own way round, and
 * then as a month and a year, and then as a year and a month.
 *
 * Measured against a fresh cscript under 1031, where the date separator is a dot: "5.1" is 5 January and
 * "31.12" is 31 December, "12.31" is 31 December too because the first field cannot be a day there, "2.30"
 * is February 2030 and "0.5" is May 2000.
 */
function twoFieldDate(order: DateOrder, firstText: string, secondText: string): WrittenDate | null {
  const year = new Date().getFullYear();
  const first = Number(firstText);
  const second = Number(secondText);
  // The last field of each candidate says whether the reading used a year, which a Buddhist calendar has to
  // know: measured under 1054, CDate("2026. 05.") is May of the Buddhist year 2026, while CDate("1. 5.") is
  // 1 May of the current year.
  const byDayAndMonth: [number, number, number, boolean][] =
    order === 'DMY'
      ? [
          [year, second, first, false],
          [year, first, second, false],
        ]
      : [
          [year, first, second, false],
          [year, second, first, false],
        ];
  // Only a year is wrapped: a day or a month written with more digits than it can hold is simply not a day or
  // a month, which is why "5.2958465" is May 9345 rather than a type mismatch.
  const byMonthAndYear: [number, number, number, boolean][] = [
    [yearOfNumber(wrapToDateField(second)), first, 1, true],
    [yearOfNumber(wrapToDateField(first)), second, 1, true],
  ];
  for (const [candidateYear, month, day, yearWritten] of [...byDayAndMonth, ...byMonthAndYear]) {
    const date = validDate(candidateYear, month, day);
    if (date) return written(date, yearWritten);
  }
  return null;
}

/**
 * The date a written three field date names, in the locale's own order and then in the other one.
 *
 * Shared by the unspaced form and the spaced dotted form, which read the same way once their fields are
 * separated: measured, "1. 5. 2026" under 1038 is 5 January 2026, the same date "1-5-2026" gives there.
 *
 * A field with more digits than a date can hold wraps, and only a year is wrapped: a day or a month written
 * that way is simply not a day or a month, which is why "5.2958465" is May 9345 rather than a type mismatch.
 */
function threeFieldDate(
  order: DateOrder,
  first: string,
  second: string,
  third: string
): WrittenDate | null {
  const candidates: [number, number, number][] = [];
  if (first.length === 4) {
    // "2026/09/29" is year first wherever it is written.
    candidates.push([Number(first), Number(second), Number(third)]);
  } else if (order === 'YMD') {
    // A year written with fewer than four digits goes through the same window as any other year field, which
    // a fresh cscript shows under 1038, 1041, 1042 and 2052: CDate("1-5-26") is 26 May 2001.
    candidates.push([yearOf(first), Number(second), Number(third)]);
    candidates.push(
      [yearOf(third), Number(first), Number(second)],
      [yearOf(third), Number(second), Number(first)]
    );
  } else if (order === 'DMY') {
    candidates.push([yearOf(third), Number(second), Number(first)]);
    // A first field that is no day and no month is a year: a fresh cscript reads "31/12/26" and "31 12 26"
    // as 2031-12-26 under 1033, not as the 2026-12-31 the day-first reading would give.
    candidates.push([yearOf(first), Number(second), Number(third)]);
    candidates.push([yearOf(third), Number(first), Number(second)]);
  } else {
    candidates.push([yearOf(third), Number(first), Number(second)]);
    candidates.push([yearOf(first), Number(second), Number(third)]);
    candidates.push([yearOf(third), Number(second), Number(first)]);
  }
  for (const [year, month, day] of candidates) {
    const date = validDate(year, month, day);
    if (date) return written(date, true);
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
 * A time whose fields are marked with a dot, which Windows reads where a dot is not the date separator.
 *
 * Its fields are held in sixteen signed bits like a date field, so under 1033 CDate("65536.5") is five past
 * midnight because 65536 wraps to 0, and CDate("65535.5") is the day count 65535.5 instead, because -1 is no
 * hour. Both were measured on a fresh cscript.
 */
function parseDottedTimePart(text: string): [number, number, number] | null {
  // The spaces after the marks are read as well as the marks themselves, which a fresh cscript shows under
  // 1033: CDate("1. 5. 26") is 1:05:26 in the morning there, and CDate("1. 5 PM") is five past one in the
  // afternoon. A dot at the very end is not part of a time ("1. 5." is a type mismatch).
  const match = /^(\d{1,9})\.\s*(\d{1,9})(?:\.\s*(\d{1,9}))?\s*([AaPp][Mm])?$/.exec(text);
  if (!match) return null;
  const hours = wrapToDateField(Number(match[1]));
  const minutes = wrapToDateField(Number(match[2]));
  const seconds = match[3] === undefined ? 0 : wrapToDateField(Number(match[3]));
  const meridiem = match[4]?.toUpperCase();
  if (hours < 0 || hours > 23 || minutes < 0 || minutes > 59 || seconds < 0 || seconds > 59)
    return null;
  const hourOfDay = meridiem ? (hours % 12) + (meridiem === 'PM' ? 12 : 0) : hours;
  return [hourOfDay, minutes, seconds];
}

/**
 * A date or a time as CDate and IsDate read it: the locale's date order and separator (or ISO), a month
 * name, a 24-hour or AM/PM time, or a date and a time. A time alone is on day 0; a date alone is at
 * midnight. `null` when it is neither.
 *
 * Only a space separates a date from a time ("2026-09-29 10:30"); "2026-09-29T10:30:00" is a type
 * mismatch in vbscript.dll, and so is a weekday name ("Tue Sep 29 2026").
 *
 * One reading is wider than Windows, and it is measured and deliberate: English month names are read under
 * the three CJK locales, whose own names Windows refuses and the engine refuses with it.
 *
 * Which name a locale answers to is measured for every locale, month and spelling, because Windows reads one
 * spelling where it refuses another: Greek reads "Ιανουάριος", the name Windows writes, and refuses the
 * inflected "Ιανουαρίου" that ICU gives in a date, and Vietnamese reads its own "Thg1" and not ICU's "Tháng
 * 1". A name in another language is refused, and a dot is a field mark only where the locale's own picture
 * writes it against the field, so "5. Januar 2026" is refused under 1042, whose dates are written with
 * hyphens.
 *
 * A dotted date written with the spaces its locale's own picture uses is read too, which is measured: Windows
 * reads CDate("2026. 05. 01.") as 1 May 2026 under 1038 and CDate("1. 5. 2026") as a date under 1031, 1038,
 * 1051, 1054 and 1060, while its unspaced form is refused under all of those but 1031. Two smaller readings
 * come with it, both measured under 1054 alone: a German month name with a dot, and the spaced form of two
 * fields.
 */
export function parseVbDate(text: string, locale: string): Date | null {
  const read = parseVbDateInItsOwnCalendar(text, locale);
  if (read === null) return null;
  return read.yearWritten ? fromBuddhistYear(read.date, locale) : read.date;
}

/**
 * A date read under a locale whose calendar counts its years from the Buddhist era, which Thai does.
 *
 * Measured under 1054: CDate("5 มกราคม 2026") is the Buddhist year 2026, which CStr writes back as
 * "5/1/2026", while CDate("5 มกราคม 2569") is the Gregorian year 2026 and writes as "5/1/2569". The other
 * calendars do not shift a year that was written: under ar-SA CDate("5 يناير 2026") prints "05/01/2026",
 * which is the Gregorian year 2026 there.
 *
 * Only a year the text itself names is shifted. A date written without one is filled in with the current
 * Gregorian year, and a time alone is on a Gregorian day 0, so shifting those answers a date that Windows
 * never answers: under 1054 CDate("1,5") is 1 May 2026 and CDate("1:05") is 1:05 on 30 December 1899.
 */
function fromBuddhistYear(date: Date, locale: string): Date {
  if (new Intl.DateTimeFormat(locale).resolvedOptions().calendar !== 'buddhist') return date;
  const shifted = new Date(date.getTime());
  shifted.setFullYear(shifted.getFullYear() - BUDDHIST_YEAR_OFFSET);
  return shifted;
}

/** The years between the Buddhist and the Gregorian count. */
const BUDDHIST_YEAR_OFFSET = 543;

function parseVbDateInItsOwnCalendar(text: string, locale: string): WrittenDate | null {
  const trimmed = text.trim();
  if (trimmed === '') return null;

  const time = parseTimePart(trimmed);
  if (time) return written(new Date(1899, 11, 30, ...time), false);

  const words = trimmed.split(/\s+/);
  for (let split = words.length - 1; split >= 1; split--) {
    // A comma may stand between the fields of a date, but not between a date and its time: a fresh cscript
    // reads "January 15, 2026 13:05" and refuses "15 January 2026, 13:05" under almost every locale. Five
    // locales accept the comma there as well, and only when the date part names its month.
    const dateText = words.slice(0, split).join(' ');
    if (words[split - 1]!.endsWith(',') && !commaMayPrecedeTime(dateText)) continue;
    const read = parseDatePart(dateText, locale);
    if (!read) continue;
    const timeOfDay = parseTimePart(words.slice(split).join(' '));
    if (!timeOfDay) continue;
    read.date.setHours(...timeOfDay);
    return read;
  }
  const date = parseDatePart(trimmed, locale);
  if (date) return date;
  // A time whose fields are marked with a dot, which Windows reads where its dates are not written that way:
  // under 1033 CDate("2.5") is 2:05 in the morning and CDate("2.5 PM") is 2:05 in the afternoon, while under
  // 1031 a dot is the date separator and the same text is 2 May. 1054 is the one locale that takes neither
  // reading: its calendar is the Buddhist one, and Windows reads "2.5" there as the day count 2.5.
  const { dottedDates } = dateOrderOf();
  if (!dottedDates && nlsDateFormat(currentLcid())?.calendarType !== THAI_BUDDHIST) {
    const dottedTime = parseDottedTimePart(trimmed);
    if (dottedTime) return written(new Date(1899, 11, 30, ...dottedTime), false);
  }
  return null;
}

/**
 * Whether a comma may stand between a date and the time that follows it.
 *
 * Measured over all 132 locales: five of them read it when the date part names its month ("15 január 2026,
 * 13:05:09" under 1038), and they refuse it when the date part is numbers ("15 1 2026, 13:05:09"), which the
 * other 127 refuse either way.
 */
function commaMayPrecedeTime(dateText: string): boolean {
  return /\p{L}/u.test(dateText) && NLS_COMMA_BEFORE_TIME.has(currentLcid());
}

/** The day counts a date holds: 1 January 100 is the earliest and 31 December 9999 the latest. */
const EARLIEST_SERIAL = -657434;
const LATEST_SERIAL = 2958465;

/**
 * A date as CDate reads one from a string, which is a wider set than IsDate and DateValue read.
 *
 * The extra shape is a plain number, read as an OLE day count: CDate("2026") is 18 July 1905. Only CDate
 * reads it, since IsDate("2026") is False and DateValue("2026") is a type mismatch, all measured on a fresh
 * cscript. The number may carry a sign, spaces around it, leading zeros or a VB literal prefix, so
 * CDate("&H10") is the 16th day and CDate(" +7") the 7th, and it has to land inside the range a date holds:
 * CDate("2958465") is 31 December 9999 while CDate("2958466") and CDate("-657435") are a type mismatch.
 */
export function parseVbDateAsCdate(text: string, locale: string): Date | null {
  const literal = readVbNumberLiteral(text);
  if (literal !== null) return literal;
  const date = parseVbDate(text, locale);
  if (date !== null) return date;
  return readMarkedNumber(text);
}

/**
 * A number written with the marks a locale uses between digits, read as a day count.
 *
 * Measured under 1031 and 1033, and only CDate reads these: IsDate("13.45") is False under 1031, and
 * IsDate("2.60"), IsDate("123.456") and IsDate("60.5") are False under 1033.
 *
 *   * The number is read with the marks the locale's own row in the Windows number table names, the same
 *     reading CDbl uses, so under 1033 "123.456" is the day count 123.456 while "13,45" is 1345 because a
 *     comma groups there, and under 1031 it is the other way round ("13,45" is 13.45 and "13.45" is 1345).
 *   * A mark the locale does not name is not a mark at all. The earlier rule took whichever of the full
 *     stop and the comma was not the fraction mark out of the number, which reads "29.9" as the day count
 *     299 under every locale that groups with a space. Windows answers a type mismatch there, as a fresh
 *     cscript shows under 1036, 1053, 1063, 1065, 1072, 1073, 1078, 1091, 2070, 2077, 2115, 3084, 5130
 *     and 7177.
 *
 * Only digits and the two marks between them are read here, so a currency symbol, a pair of parentheses and
 * an exponent are left to the other readers. The whole days have to land inside the range a date holds,
 * which is why CDate("2958466") refuses while CDate("2958465.5") is 31 December 9999 at noon.
 */
function readMarkedNumber(text: string): Date | null {
  const trimmed = text.trim();
  // A space may stand where the locale groups its numbers, which the number reader itself decides: measured,
  // CDate("2026 15 1") is the day count 2026151 under 1026, where a space is the group mark, and a type
  // mismatch under 1033, where it is not.
  if (!/^[+-]?[\d.,\s]+$/.test(trimmed)) return null;
  const serial = parseVbNumericString(trimmed);
  if (serial === null || !Number.isFinite(serial)) return null;
  const wholeDays = Math.trunc(serial);
  if (wholeDays < EARLIEST_SERIAL || wholeDays > LATEST_SERIAL) return null;
  return serialToDate(serial);
}

/** A whole number written the way VB writes one, as a day count, or null when the text is something else. */
function readVbNumberLiteral(text: string): Date | null {
  const match = /^([+-]?)(?:&[Hh]([0-9A-Fa-f]+)|&[Oo]([0-7]+)|(\d+))$/.exec(text.trim());
  if (!match) return null;
  const magnitude =
    match[2] !== undefined
      ? parseInt(match[2], 16)
      : match[3] !== undefined
        ? parseInt(match[3], 8)
        : Number(match[4]);
  const serial = match[1] === '-' ? -magnitude : magnitude;
  if (serial < EARLIEST_SERIAL || serial > LATEST_SERIAL) return null;
  return serialToDate(serial);
}
