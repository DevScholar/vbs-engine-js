import type { VbValue } from '../runtime/index.ts';
import {
  createVbError,
  rejectObjectArgument,
  roundHalfToEven,
  serialToDate,
  standaloneMonthName,
  standaloneWeekdayName,
  toNumber,
  toString,
  toVbDate,
  VbErrorCodes,
  VbNull,
} from '../runtime/index.ts';
import { NLS_FIRST_DAY_OF_WEEK } from '../runtime/nls-first-day.generated.ts';
import { currentLcid, getCurrentBCP47Locale } from './locale.ts';

/**
 * The month or weekday index a name function was handed, which is a whole number from 1 to the last one.
 *
 * Measured on a fresh cscript over all 132 locales: the index is rounded to a whole number with a tie going
 * to the even neighbour (MonthName(3.6) is April and MonthName(2.5) is February), and anything outside the
 * range is error 5 rather than an empty name, MonthName(0), MonthName(13), MonthName(-1), MonthName(Empty)
 * and MonthName(True) alike. A Nothing is error 91 and a Null is 94.
 */
function readNameIndex(value: VbValue, last: number, functionName: string): number {
  rejectObjectArgument(value, functionName);
  const index = roundHalfToEven(toNumber(value));
  if (index < 1 || index > last) {
    throw createVbError(
      VbErrorCodes.InvalidProcedureCall,
      `Invalid procedure call or argument: ${functionName}`
    );
  }
  return index;
}

/** Whether the name was asked for in its short form, which any number other than zero asks for. */
function readAbbreviate(value: VbValue | undefined): boolean {
  if (!value) return false;
  return toNumber(value) !== 0;
}

/**
 * The date argument of DateValue and TimeValue, which take a date or a string only: a number, a Boolean
 * or Empty is a type mismatch there even though CDate reads them, and Null is error 94.
 */
function toVbDateArgument(value: VbValue): Date {
  if (value.type === 'Null')
    throw createVbError(VbErrorCodes.InvalidUseOfNull, 'Invalid use of Null');
  if (value.type !== 'String' && value.type !== 'Date') {
    throw createVbError(
      VbErrorCodes.TypeMismatch,
      `Type mismatch: ${value.type} cannot be converted to Date`
    );
  }
  return toVbDate(value);
}

/** The name of the error a Null in an argument that is no date raises, which is not what a Null date does. */
function nullArgument(functionName: string): never {
  throw createVbError(VbErrorCodes.InvalidUseOfNull, `Invalid use of Null: ${functionName}`);
}

/**
 * The date a date function was handed, or null for a Null, which the caller answers with Null.
 *
 * Measured on a fresh cscript: a Null date answers Null for Year, Month, Day, Hour, Minute, Second, Weekday,
 * DatePart, DateAdd and DateDiff alike, while a Null anywhere else in those calls is the invalid use of Null
 * (DatePart(Null, Date), DateAdd("d", Null, Date), DateDiff(Null, Date, Date) and Weekday(Date, Null) are all
 * error 94), and DateSerial and TimeSerial refuse a Null in every position.
 */
function readDateArgument(value: VbValue): Date | null {
  if (value.type === 'Null') return null;
  return value.type === 'Date' ? (value.value as Date) : serialToDate(toNumber(value));
}

export const dateFunctions = {
  Now: (): VbValue => {
    return { type: 'Date', value: new Date() };
  },

  Date: (): VbValue => {
    const now = new Date();
    return { type: 'Date', value: new Date(now.getFullYear(), now.getMonth(), now.getDate()) };
  },

  Time: (): VbValue => {
    const now = new Date();
    return {
      type: 'Date',
      value: new Date(1899, 11, 30, now.getHours(), now.getMinutes(), now.getSeconds()),
    };
  },

  Year: (date: VbValue): VbValue => {
    const d = readDateArgument(date);
    if (d === null) return VbNull;
    return { type: 'Integer', value: d.getFullYear() };
  },

  Month: (date: VbValue): VbValue => {
    const d = readDateArgument(date);
    if (d === null) return VbNull;
    return { type: 'Integer', value: d.getMonth() + 1 };
  },

  Day: (date: VbValue): VbValue => {
    const d = readDateArgument(date);
    if (d === null) return VbNull;
    return { type: 'Integer', value: d.getDate() };
  },

  Weekday: (date: VbValue, firstDayOfWeek?: VbValue): VbValue => {
    const d = readDateArgument(date);
    if (d === null) return VbNull;
    if (firstDayOfWeek && firstDayOfWeek.type === 'Null') nullArgument('Weekday');
    const firstDay = firstDayOfWeek ? Math.floor(toNumber(firstDayOfWeek)) : 1;
    let day = d.getDay() + 1;
    day = day - firstDay + 1;
    if (day < 1) day += 7;
    return { type: 'Integer', value: day };
  },

  Hour: (time: VbValue): VbValue => {
    const d = readDateArgument(time);
    if (d === null) return VbNull;
    return { type: 'Integer', value: d.getHours() };
  },

  Minute: (time: VbValue): VbValue => {
    const d = readDateArgument(time);
    if (d === null) return VbNull;
    return { type: 'Integer', value: d.getMinutes() };
  },

  Second: (time: VbValue): VbValue => {
    const d = readDateArgument(time);
    if (d === null) return VbNull;
    return { type: 'Integer', value: d.getSeconds() };
  },

  DateAdd: (interval: VbValue, number: VbValue, date: VbValue): VbValue => {
    if (interval.type === 'Null' || number.type === 'Null') nullArgument('DateAdd');
    const d = readDateArgument(date);
    if (d === null) return VbNull;
    const intv = toString(interval).toLowerCase();
    const num = toNumber(number);
    const shifted = new Date(d);

    switch (intv) {
      case 'yyyy':
        shifted.setFullYear(shifted.getFullYear() + num);
        break;
      case 'q':
        shifted.setMonth(shifted.getMonth() + num * 3);
        break;
      case 'm':
        shifted.setMonth(shifted.getMonth() + num);
        break;
      case 'y':
      case 'd':
      case 'w':
        shifted.setDate(shifted.getDate() + num);
        break;
      case 'ww':
        shifted.setDate(shifted.getDate() + num * 7);
        break;
      case 'h':
        shifted.setHours(shifted.getHours() + num);
        break;
      case 'n':
        shifted.setMinutes(shifted.getMinutes() + num);
        break;
      case 's':
        shifted.setSeconds(shifted.getSeconds() + num);
        break;
    }

    return { type: 'Date', value: shifted };
  },

  DateDiff: (
    interval: VbValue,
    date1: VbValue,
    date2: VbValue,
    _firstDayOfWeek?: VbValue,
    _firstWeekOfYear?: VbValue
  ): VbValue => {
    void _firstDayOfWeek; // Intentionally unused - matches VBScript signature
    void _firstWeekOfYear; // Intentionally unused - matches VBScript signature
    if (interval.type === 'Null') nullArgument('DateDiff');
    const d1 = readDateArgument(date1);
    const d2 = readDateArgument(date2);
    if (d1 === null || d2 === null) return VbNull;
    const intv = toString(interval).toLowerCase();
    const diff = d2.getTime() - d1.getTime();

    let result: number;
    switch (intv) {
      case 'yyyy':
        result = d2.getFullYear() - d1.getFullYear();
        break;
      case 'q':
        result = Math.floor(
          (d2.getFullYear() - d1.getFullYear()) * 4 + (d2.getMonth() - d1.getMonth()) / 3
        );
        break;
      case 'm':
        result = (d2.getFullYear() - d1.getFullYear()) * 12 + (d2.getMonth() - d1.getMonth());
        break;
      case 'y':
      case 'd':
        result = Math.floor(diff / 86400000);
        break;
      case 'w':
        result = Math.floor(diff / 86400000);
        break;
      case 'ww':
        result = Math.floor(diff / (86400000 * 7));
        break;
      case 'h':
        result = Math.floor(diff / 3600000);
        break;
      case 'n':
        result = Math.floor(diff / 60000);
        break;
      case 's':
        result = Math.floor(diff / 1000);
        break;
      default:
        result = 0;
    }

    return { type: 'Long', value: result };
  },

  DatePart: (
    interval: VbValue,
    date: VbValue,
    _firstDayOfWeek?: VbValue,
    _firstWeekOfYear?: VbValue
  ): VbValue => {
    void _firstDayOfWeek; // Intentionally unused - matches VBScript signature
    void _firstWeekOfYear; // Intentionally unused - matches VBScript signature
    if (interval.type === 'Null') nullArgument('DatePart');
    const d = readDateArgument(date);
    if (d === null) return VbNull;
    const intv = toString(interval).toLowerCase();

    let result: number;
    switch (intv) {
      case 'yyyy':
        result = d.getFullYear();
        break;
      case 'q':
        result = Math.floor(d.getMonth() / 3) + 1;
        break;
      case 'm':
        result = d.getMonth() + 1;
        break;
      case 'y': {
        const startOfYear = new Date(d.getFullYear(), 0, 1);
        result = Math.floor((d.getTime() - startOfYear.getTime()) / 86400000) + 1;
        break;
      }
      case 'd':
        result = d.getDate();
        break;
      case 'w':
        result = d.getDay() + 1;
        break;
      case 'ww': {
        const startOfYear2 = new Date(d.getFullYear(), 0, 1);
        result = Math.floor((d.getTime() - startOfYear2.getTime()) / (86400000 * 7)) + 1;
        break;
      }
      case 'h':
        result = d.getHours();
        break;
      case 'n':
        result = d.getMinutes();
        break;
      case 's':
        result = d.getSeconds();
        break;
      default:
        result = 0;
    }

    return { type: 'Integer', value: result };
  },

  DateSerial: (year: VbValue, month: VbValue, day: VbValue): VbValue => {
    const y = Math.floor(toNumber(year));
    const m = Math.floor(toNumber(month));
    const d = Math.floor(toNumber(day));
    return { type: 'Date', value: new Date(y, m - 1, d) };
  },

  TimeSerial: (hour: VbValue, minute: VbValue, second: VbValue): VbValue => {
    const h = Math.floor(toNumber(hour));
    const m = Math.floor(toNumber(minute));
    const s = Math.floor(toNumber(second));
    return { type: 'Date', value: new Date(1899, 11, 30, h, m, s) };
  },

  DateValue: (date: VbValue): VbValue => {
    const d = toVbDateArgument(date);
    return { type: 'Date', value: new Date(d.getFullYear(), d.getMonth(), d.getDate()) };
  },

  TimeValue: (time: VbValue): VbValue => {
    const d = toVbDateArgument(time);
    return {
      type: 'Date',
      value: new Date(1899, 11, 30, d.getHours(), d.getMinutes(), d.getSeconds()),
    };
  },

  MonthName: (month: VbValue, abbreviate?: VbValue): VbValue => {
    const m = readNameIndex(month, 12, 'MonthName');
    return {
      type: 'String',
      value: standaloneMonthName(
        currentLcid(),
        getCurrentBCP47Locale(),
        m,
        readAbbreviate(abbreviate)
      ),
    };
  },

  // WeekdayName(weekday[, abbreviate[, firstdayofweek]]), measured: the third argument counts the week from
  // a day of its own, so WeekdayName(1, False, 2) is the name of Monday and WeekdayName(1, False, 7) that of
  // Saturday, and 0 counts from the locale's own first day, which is Monday for most of Europe and Saturday
  // for the Arabic locales. Anything outside 0 to 7 is error 5, and leaving the argument out counts from
  // Sunday, whatever the locale's own day is.
  WeekdayName: (weekday: VbValue, abbreviate?: VbValue, firstDayOfWeek?: VbValue): VbValue => {
    const w = readNameIndex(weekday, 7, 'WeekdayName');
    let first = 1;
    if (firstDayOfWeek !== undefined) {
      rejectObjectArgument(firstDayOfWeek, 'WeekdayName');
      const given = roundHalfToEven(toNumber(firstDayOfWeek));
      if (given < 0 || given > 7) {
        throw createVbError(
          VbErrorCodes.InvalidProcedureCall,
          'Invalid procedure call or argument: WeekdayName'
        );
      }
      first = given === 0 ? (NLS_FIRST_DAY_OF_WEEK.get(currentLcid()) ?? 1) : given;
    }
    const day = ((first - 1 + (w - 1)) % 7) + 1;
    return {
      type: 'String',
      value: standaloneWeekdayName(
        currentLcid(),
        getCurrentBCP47Locale(),
        day,
        readAbbreviate(abbreviate)
      ),
    };
  },

  Timer: (): VbValue => {
    const now = new Date();
    return {
      type: 'Single',
      value:
        now.getHours() * 3600 +
        now.getMinutes() * 60 +
        now.getSeconds() +
        now.getMilliseconds() / 1000,
    };
  },
};
