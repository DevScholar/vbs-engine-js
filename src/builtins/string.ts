import type { VbValue } from '../runtime/index.ts';
import {
  toNumber,
  toString,
  createVbArrayFromValues,
  createVbError,
  VbErrorCodes,
  VbNull,
} from '../runtime/index.ts';
import { getCurrentBCP47Locale } from './locale.ts';

/** The invalid use of Null, which the string functions raise where the others answer it back. */
function nullArgument(functionName: string): never {
  throw createVbError(VbErrorCodes.InvalidUseOfNull, `Invalid use of Null: ${functionName}`);
}

/** The invalid procedure call, which an argument out of its range raises. */
function invalidArgument(functionName: string, what: string): never {
  throw createVbError(
    VbErrorCodes.InvalidProcedureCall,
    `Invalid procedure call: ${functionName} ${what}`
  );
}

/** Whether any of the arguments handed to a string function is a Null. */
function anyNullArgument(...values: (VbValue | undefined)[]): boolean {
  return values.some(value => value !== undefined && value.type === 'Null');
}

/** A compare mode, which VBScript allows to be 0 (binary) or 1 (text) and nothing else. */
function readCompareMode(value: VbValue | undefined, functionName: string): number {
  if (value === undefined) return 0;
  const mode = Math.floor(toNumber(value));
  if (mode !== 0 && mode !== 1) invalidArgument(functionName, 'compare');
  return mode;
}

/**
 * Text cut at a mark, which is what Split does.
 *
 * The mark is matched whole and, with a text compare, whatever its case, and a limit of one or more caps the
 * number of parts by leaving the rest of the text in the last one.
 */
function splitText(text: string, mark: string, limit: number, compareMode: number): string[] {
  if (mark === '') return [text];
  const haystack = compareMode === 0 ? text : text.toLowerCase();
  const needle = compareMode === 0 ? mark : mark.toLowerCase();
  const parts: string[] = [];
  let from = 0;
  while (limit <= 0 || parts.length < limit - 1) {
    const at = haystack.indexOf(needle, from);
    if (at < 0) break;
    parts.push(text.slice(from, at));
    from = at + needle.length;
  }
  parts.push(text.slice(from));
  return parts;
}

function formatDate(value: Date, format: string): string {
  const result: string[] = [];
  let i = 0;

  while (i < format.length) {
    const char = format[i]!.toLowerCase();
    let count = 1;
    while (i + count < format.length && format[i + count]!.toLowerCase() === char) {
      count++;
    }

    switch (char) {
      case 'y':
        if (count >= 4) {
          result.push(value.getFullYear().toString());
        } else if (count === 3) {
          result.push(value.getFullYear().toString());
        } else if (count === 2) {
          result.push(value.getFullYear().toString().slice(-2));
        } else {
          result.push(value.getFullYear().toString().slice(-2));
        }
        break;
      case 'm':
        if (count >= 4) {
          const months: string[] = [
            'January',
            'February',
            'March',
            'April',
            'May',
            'June',
            'July',
            'August',
            'September',
            'October',
            'November',
            'December',
          ];
          result.push(months[value.getMonth()]!);
        } else if (count === 3) {
          const months: string[] = [
            'Jan',
            'Feb',
            'Mar',
            'Apr',
            'May',
            'Jun',
            'Jul',
            'Aug',
            'Sep',
            'Oct',
            'Nov',
            'Dec',
          ];
          result.push(months[value.getMonth()]!);
        } else if (count === 2) {
          result.push((value.getMonth() + 1).toString().padStart(2, '0'));
        } else {
          result.push((value.getMonth() + 1).toString());
        }
        break;
      case 'd':
        if (count >= 4) {
          const days: string[] = [
            'Sunday',
            'Monday',
            'Tuesday',
            'Wednesday',
            'Thursday',
            'Friday',
            'Saturday',
          ];
          result.push(days[value.getDay()]!);
        } else if (count === 3) {
          const days: string[] = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
          result.push(days[value.getDay()]!);
        } else if (count === 2) {
          result.push(value.getDate().toString().padStart(2, '0'));
        } else {
          result.push(value.getDate().toString());
        }
        break;
      case 'h':
        if (count >= 2) {
          result.push((value.getHours() % 12 || 12).toString().padStart(2, '0'));
        } else {
          result.push((value.getHours() % 12 || 12).toString());
        }
        break;
      case 'n':
        if (count >= 2) {
          result.push(value.getMinutes().toString().padStart(2, '0'));
        } else {
          result.push(value.getMinutes().toString());
        }
        break;
      case 's':
        if (count >= 2) {
          result.push(value.getSeconds().toString().padStart(2, '0'));
        } else {
          result.push(value.getSeconds().toString());
        }
        break;
      case 'q':
        result.push(Math.floor(value.getMonth() / 3 + 1).toString());
        break;
      case 'w':
        result.push((value.getDay() + 1).toString());
        break;
      case 'a':
        if (count >= 2) {
          result.push(value.getHours() >= 12 ? 'PM' : 'AM');
        } else {
          result.push(value.getHours() >= 12 ? 'P' : 'A');
        }
        break;
      default:
        result.push(format.substring(i, i + count));
    }
    i += count;
  }

  return result.join('');
}

function formatNumber(value: number, format: string): string {
  let positiveFormat = format;
  let negativeFormat = '';
  let zeroFormat = '';

  const semicolonIndex = format.indexOf(';');
  if (semicolonIndex !== -1) {
    positiveFormat = format.substring(0, semicolonIndex);
    const remaining = format.substring(semicolonIndex + 1);
    const secondSemicolon = remaining.indexOf(';');
    if (secondSemicolon !== -1) {
      negativeFormat = remaining.substring(0, secondSemicolon);
      zeroFormat = remaining.substring(secondSemicolon + 1);
    } else {
      negativeFormat = remaining;
    }
  }

  if (value === 0 && zeroFormat) {
    return formatNumberInternal(0, zeroFormat);
  }

  if (value < 0 && negativeFormat) {
    return formatNumberInternal(Math.abs(value), negativeFormat);
  }

  if (value < 0) {
    return '-' + formatNumberInternal(Math.abs(value), positiveFormat);
  }

  return formatNumberInternal(value, positiveFormat);
}

function formatNumberInternal(value: number, format: string): string {
  const hasPercent = format.includes('%');
  if (hasPercent) {
    value *= 100;
  }

  const formatLower = format.toLowerCase();
  const decimalPos = formatLower.indexOf('.');
  let decimalDigits = 0;

  if (decimalPos !== -1) {
    let i = decimalPos + 1;
    while (i < format.length && (format[i] === '0' || format[i] === '#')) {
      decimalDigits++;
      i++;
    }
  }

  let intPart = Math.floor(Math.abs(value));
  let decPart =
    decimalDigits > 0 ? Math.round((Math.abs(value) - intPart) * Math.pow(10, decimalDigits)) : 0;

  if (decPart >= Math.pow(10, decimalDigits)) {
    intPart++;
    decPart = 0;
  }

  let intStr = intPart.toString();

  const commaPos = formatLower.indexOf(',');
  if (commaPos !== -1) {
    intStr = intStr.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  }

  let result = intStr;

  if (decimalDigits > 0) {
    const decStr = decPart.toString().padStart(decimalDigits, '0');
    result += '.' + decStr;
  }

  if (hasPercent) {
    result += '%';
  }

  return result;
}

function formatString(value: string, format: string): string {
  if (format === '>' || format.toLowerCase() === '>') {
    return value.toUpperCase();
  }
  if (format === '<' || format.toLowerCase() === '<') {
    return value.toLowerCase();
  }
  return value;
}

export const stringFunctions = {
  // The functions that take text and give text back answer Null for a Null argument, which a fresh cscript
  // shows for Len, Left, Mid, UCase and Trim alike; the ones that build something from their arguments
  // (Replace, String, Join, Filter and Split) raise the invalid use of Null instead.
  Len: (str: VbValue): VbValue => {
    if (anyNullArgument(str)) return VbNull;
    const s = toString(str);
    return { type: 'Long', value: s.length };
  },

  Left: (str: VbValue, length: VbValue): VbValue => {
    if (anyNullArgument(str, length)) return VbNull;
    const s = toString(str);
    const len = Math.floor(toNumber(length));
    if (len < 0) invalidArgument('Left', 'length');
    return { type: 'String', value: s.substring(0, len) };
  },

  Right: (str: VbValue, length: VbValue): VbValue => {
    if (anyNullArgument(str, length)) return VbNull;
    const s = toString(str);
    const len = Math.floor(toNumber(length));
    if (len < 0) invalidArgument('Right', 'length');
    return { type: 'String', value: s.substring(s.length - len) };
  },

  Mid: (str: VbValue, start: VbValue, length?: VbValue): VbValue => {
    if (anyNullArgument(str, start, length)) return VbNull;
    const s = toString(str);
    // The start counts from 1 and a start below it is error 5, which a fresh cscript shows for Mid("abc", 0)
    // and Mid("abc", -1) alike, and so is a length below zero.
    const startPos = Math.floor(toNumber(start)) - 1;
    if (startPos < 0) invalidArgument('Mid', 'start');
    if (length) {
      const len = Math.floor(toNumber(length));
      if (len < 0) invalidArgument('Mid', 'length');
      return { type: 'String', value: s.substring(startPos, startPos + len) };
    }
    return { type: 'String', value: s.substring(startPos) };
  },

  InStr: (
    startOrStr: VbValue,
    str1OrStr2?: VbValue,
    str2OrCompare?: VbValue,
    compare?: VbValue
  ): VbValue => {
    // InStr([start,] string1, string2[, compare]), measured against a fresh cscript:
    //   * With two strings and no more, a Null on either side answers Null, a text to search that is empty
    //     answers 0, and an empty string to look for answers 1.
    //   * With a start, a Null is the invalid use of Null, a start below 1 is error 5, a start past the end
    //     of the text is 0 (an empty string to look for included), and a compare other than 0 or 1 is error 5.
    let start = 1;
    let str1: string;
    let str2: string;
    let compareMode = 0;

    if (str1OrStr2 === undefined) {
      str1 = '';
      str2 = toString(startOrStr);
    } else if (str2OrCompare === undefined) {
      if (startOrStr.type === 'Null' || str1OrStr2.type === 'Null') return VbNull;
      str1 = toString(startOrStr);
      str2 = toString(str1OrStr2);
      if (str2 === '') return { type: 'Long', value: str1 === '' ? 0 : 1 };
      if (str1 === '') return { type: 'Long', value: 0 };
    } else {
      start = Math.floor(toNumber(startOrStr));
      str1 = toString(str1OrStr2);
      str2 = toString(str2OrCompare);
      compareMode = readCompareMode(compare, 'InStr');
      if (start < 1) invalidArgument('InStr', 'start');
      if (str2 === '') {
        return { type: 'Long', value: start <= str1.length + 1 ? start : 0 };
      }
      if (start > str1.length) return { type: 'Long', value: 0 };
    }

    let index: number;
    if (compareMode === 0) {
      index = str1.indexOf(str2, start - 1);
    } else {
      index = str1.toLowerCase().indexOf(str2.toLowerCase(), start - 1);
    }
    return { type: 'Long', value: index + 1 };
  },

  InStrRev: (str1: VbValue, str2: VbValue, start?: VbValue, compare?: VbValue): VbValue => {
    // InStrRev(string1, string2[, start[, compare]]), measured:
    //   * A Null anywhere is the invalid use of Null, and a compare other than 0 or 1 is error 5.
    //   * The start is the position to search back from and defaults to -1, which means the end of the text.
    //     Any other negative start is error 5, and so is 0. A start past the end of the text answers 0.
    //   * An empty string to look for answers with the start itself.
    if (anyNullArgument(str1, str2, start, compare)) nullArgument('InStrRev');
    const s1 = toString(str1);
    const s2 = toString(str2);
    const compareMode = readCompareMode(compare, 'InStrRev');

    let from: number;
    if (start === undefined) {
      from = s1.length;
    } else {
      const startPos = Math.floor(toNumber(start));
      if (startPos === -1) from = s1.length;
      else if (startPos < 1) invalidArgument('InStrRev', 'start');
      else if (startPos > s1.length) return { type: 'Long', value: 0 };
      else from = startPos;
    }
    if (s2 === '') return { type: 'Long', value: from };

    const index =
      compareMode === 0
        ? s1.lastIndexOf(s2, from - 1)
        : s1.toLowerCase().lastIndexOf(s2.toLowerCase(), from - 1);
    return { type: 'Long', value: index + 1 };
  },

  LCase: (str: VbValue): VbValue => {
    if (anyNullArgument(str)) return VbNull;
    return { type: 'String', value: toString(str).toLowerCase() };
  },

  UCase: (str: VbValue): VbValue => {
    if (anyNullArgument(str)) return VbNull;
    return { type: 'String', value: toString(str).toUpperCase() };
  },

  LTrim: (str: VbValue): VbValue => {
    if (anyNullArgument(str)) return VbNull;
    // Only a space comes off, which a fresh cscript shows: Trim leaves a tab, a line break and a
    // non-breaking space alone, and the length of Len(Trim(Chr(9) & "a" & Chr(9))) is 3.
    return { type: 'String', value: toString(str).replace(/^ +/, '') };
  },

  RTrim: (str: VbValue): VbValue => {
    if (anyNullArgument(str)) return VbNull;
    return { type: 'String', value: toString(str).replace(/ +$/, '') };
  },

  Trim: (str: VbValue): VbValue => {
    if (anyNullArgument(str)) return VbNull;
    return { type: 'String', value: toString(str).replace(/^ +/, '').replace(/ +$/, '') };
  },

  Replace: (
    str: VbValue,
    find: VbValue,
    replace: VbValue,
    start?: VbValue,
    count?: VbValue,
    compare?: VbValue
  ): VbValue => {
    if (anyNullArgument(str, find, replace, start, count, compare)) nullArgument('Replace');
    const s = toString(str);
    const findStr = toString(find);
    const replaceStr = toString(replace);
    const startPos = start ? Math.max(1, Math.floor(toNumber(start))) - 1 : 0;
    const maxCount = count ? Math.floor(toNumber(count)) : -1;
    const compareMode = compare ? Math.floor(toNumber(compare)) : 0;
    // An empty string to look for leaves the text as it stands, which a fresh cscript shows for
    // Replace("abc", "", "X"): there is nothing to step over, so the search would never end.
    if (findStr === '') return { type: 'String', value: s.substring(startPos) };

    let result = '';
    let replaced = 0;
    let i = startPos;
    const searchStr = compareMode === 0 ? s : s.toLowerCase();
    const searchFind = compareMode === 0 ? findStr : findStr.toLowerCase();

    while (i < s.length) {
      if (maxCount !== -1 && replaced >= maxCount) {
        result += s.substring(i);
        break;
      }
      if (searchStr.substring(i, i + findStr.length) === searchFind) {
        result += replaceStr;
        i += findStr.length;
        replaced++;
      } else {
        result += s[i];
        i++;
      }
    }

    return { type: 'String', value: result };
  },

  StrReverse: (str: VbValue): VbValue => {
    return { type: 'String', value: toString(str).split('').reverse().join('') };
  },

  Space: (number: VbValue): VbValue => {
    if (anyNullArgument(number)) nullArgument('Space');
    const count = Math.floor(toNumber(number));
    if (count < 0) invalidArgument('Space', 'number');
    return { type: 'String', value: ' '.repeat(count) };
  },

  String: (number: VbValue, character: VbValue): VbValue => {
    if (anyNullArgument(number, character)) nullArgument('String');
    const count = Math.floor(toNumber(number));
    if (count < 0) invalidArgument('String', 'number');
    // A number is a character code and a string is its first character, which a fresh cscript shows:
    // String(3, 65) is "AAA" where String(3, "ab") is "aaa".
    const char =
      character.type === 'String'
        ? (toString(character)[0] ?? ' ')
        : String.fromCharCode(toNumber(character));
    return { type: 'String', value: char.repeat(count) };
  },

  Asc: (str: VbValue): VbValue => {
    const s = toString(str);
    return { type: 'Integer', value: s.length > 0 ? s.charCodeAt(0) : 0 };
  },

  AscW: (str: VbValue): VbValue => {
    const s = toString(str);
    return { type: 'Integer', value: s.length > 0 ? s.charCodeAt(0) : 0 };
  },

  Chr: (charCode: VbValue): VbValue => {
    const code = Math.floor(toNumber(charCode));
    return { type: 'String', value: String.fromCharCode(code) };
  },

  ChrW: (charCode: VbValue): VbValue => {
    const code = Math.floor(toNumber(charCode));
    return { type: 'String', value: String.fromCharCode(code) };
  },

  StrComp: (str1: VbValue, str2: VbValue, compare?: VbValue): VbValue => {
    const s1 = toString(str1);
    const s2 = toString(str2);
    const compareMode = compare ? Math.floor(toNumber(compare)) : 0;

    let result: number;
    if (compareMode === 1) {
      result = s1.toLowerCase().localeCompare(s2.toLowerCase());
    } else if (compareMode === 2) {
      const locale = getCurrentBCP47Locale();
      result = s1.localeCompare(s2, locale, { sensitivity: 'base' });
    } else {
      result = s1.localeCompare(s2);
    }
    return { type: 'Integer', value: result };
  },

  Split: (str: VbValue, delimiter?: VbValue, count?: VbValue, compare?: VbValue): VbValue => {
    // Split(expression[, delimiter[, limit[, compare]]]), measured:
    //   * A Null anywhere is the invalid use of Null, and a compare other than 0 or 1 is error 5.
    //   * The delimiter defaults to a space, and an empty delimiter leaves the whole text as one part.
    //   * The limit is how many parts to make and defaults to -1, which is as many as there are. A limit
    //     below -1 is error 5, a limit of 0 makes no parts at all, and 1 leaves the text whole.
    //   * A text that is empty makes no parts either, whatever the delimiter, and a delimiter of one
    //     character that is the whole text makes two empty parts.
    if (anyNullArgument(str, delimiter, count, compare)) nullArgument('Split');
    const compareMode = readCompareMode(compare, 'Split');
    const text = toString(str);
    const mark = delimiter === undefined ? ' ' : toString(delimiter);
    const limit = count === undefined ? -1 : Math.floor(toNumber(count));
    if (limit < -1) invalidArgument('Split', 'limit');

    const parts = limit === 0 || text === '' ? [] : splitText(text, mark, limit, compareMode);
    const vbArray = createVbArrayFromValues(
      parts.map(part => ({ type: 'String', value: part }) as VbValue)
    );
    return { type: 'Array', value: vbArray };
  },

  Join: (list: VbValue, delimiter?: VbValue): VbValue => {
    if (anyNullArgument(list, delimiter)) nullArgument('Join');
    const delim = delimiter ? toString(delimiter) : ' ';
    if (list.type === 'Array') {
      const arr = list.value;
      if (typeof arr === 'object' && arr !== null && 'toArray' in arr) {
        const vbArray = arr as { toArray: () => VbValue[] };
        return {
          type: 'String',
          value: vbArray
            .toArray()
            .map(v => toString(v))
            .join(delim),
        };
      }
      const plainArr = arr as VbValue[];
      return { type: 'String', value: plainArr.map(v => toString(v)).join(delim) };
    }
    return { type: 'String', value: '' };
  },

  Format: (expression: VbValue, format?: VbValue): VbValue => {
    if (expression.type === 'Empty' || expression.type === 'Null') {
      return { type: 'String', value: '' };
    }

    const formatStr = format ? toString(format) : '';

    if (!formatStr) {
      return { type: 'String', value: toString(expression) };
    }

    if (expression.type === 'Date' && expression.value instanceof Date) {
      return { type: 'String', value: formatDate(expression.value, formatStr) };
    }

    if (expression.type === 'Boolean') {
      return { type: 'String', value: expression.value ? 'True' : 'False' };
    }

    if (['Integer', 'Long', 'Single', 'Double', 'Currency', 'Byte'].includes(expression.type)) {
      return { type: 'String', value: formatNumber(toNumber(expression), formatStr) };
    }

    if (expression.type === 'String') {
      return { type: 'String', value: formatString(toString(expression), formatStr) };
    }

    return { type: 'String', value: toString(expression) };
  },

  LSet: (string: VbValue, length: VbValue): VbValue => {
    const str = toString(string);
    const len = Math.max(0, Math.floor(toNumber(length)));

    if (str.length >= len) {
      return { type: 'String', value: str.substring(0, len) };
    }

    return { type: 'String', value: str.padEnd(len, ' ') };
  },

  RSet: (string: VbValue, length: VbValue): VbValue => {
    const str = toString(string);
    const len = Math.max(0, Math.floor(toNumber(length)));

    if (str.length >= len) {
      return { type: 'String', value: str.substring(0, len) };
    }

    return { type: 'String', value: str.padStart(len, ' ') };
  },
};
