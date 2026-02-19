import type { VbValue } from '../runtime/index.ts';
import { toNumber, toString, createVbArrayFromValues } from '../runtime/index.ts';
import { getCurrentBCP47Locale } from './locale.ts';

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
          const months: string[] = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
          result.push(months[value.getMonth()]!);
        } else if (count === 3) {
          const months: string[] = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
          result.push(months[value.getMonth()]!);
        } else if (count === 2) {
          result.push((value.getMonth() + 1).toString().padStart(2, '0'));
        } else {
          result.push((value.getMonth() + 1).toString());
        }
        break;
      case 'd':
        if (count >= 4) {
          const days: string[] = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
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
  let hasPercent = format.includes('%');
  if (hasPercent) {
    value *= 100;
  }
  
  const formatLower = format.toLowerCase();
  let decimalPos = formatLower.indexOf('.');
  let decimalDigits = 0;
  
  if (decimalPos !== -1) {
    let i = decimalPos + 1;
    while (i < format.length && (format[i] === '0' || format[i] === '#')) {
      decimalDigits++;
      i++;
    }
  }
  
  let intPart = Math.floor(Math.abs(value));
  let decPart = decimalDigits > 0 ? Math.round((Math.abs(value) - intPart) * Math.pow(10, decimalDigits)) : 0;
  
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
  Len: (str: VbValue): VbValue => {
    const s = toString(str);
    return { type: 'Long', value: s.length };
  },

  Left: (str: VbValue, length: VbValue): VbValue => {
    const s = toString(str);
    const len = Math.max(0, Math.floor(toNumber(length)));
    return { type: 'String', value: s.substring(0, len) };
  },

  Right: (str: VbValue, length: VbValue): VbValue => {
    const s = toString(str);
    const len = Math.max(0, Math.floor(toNumber(length)));
    return { type: 'String', value: s.substring(s.length - len) };
  },

  Mid: (str: VbValue, start: VbValue, length?: VbValue): VbValue => {
    const s = toString(str);
    const startPos = Math.max(1, Math.floor(toNumber(start))) - 1;
    if (length) {
      const len = Math.floor(toNumber(length));
      return { type: 'String', value: s.substring(startPos, startPos + len) };
    }
    return { type: 'String', value: s.substring(startPos) };
  },

  InStr: (startOrStr: VbValue, str1OrStr2?: VbValue, str2OrCompare?: VbValue, compare?: VbValue): VbValue => {
    let start = 1;
    let str1: string;
    let str2: string;
    let compareMode = 0;
    
    if (str1OrStr2 === undefined) {
      str1 = '';
      str2 = toString(startOrStr);
    } else if (str2OrCompare === undefined) {
      str1 = toString(startOrStr);
      str2 = toString(str1OrStr2);
    } else {
      start = Math.max(1, Math.floor(toNumber(startOrStr)));
      str1 = toString(str1OrStr2);
      str2 = toString(str2OrCompare);
      if (compare) {
        compareMode = Math.floor(toNumber(compare));
      }
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
    const s1 = toString(str1);
    const s2 = toString(str2);
    const startPos = start ? Math.floor(toNumber(start)) : s1.length;
    const compareMode = compare ? Math.floor(toNumber(compare)) : 0;
    
    let index: number;
    if (compareMode === 0) {
      index = s1.lastIndexOf(s2, startPos - 1);
    } else {
      index = s1.toLowerCase().lastIndexOf(s2.toLowerCase(), startPos - 1);
    }
    return { type: 'Long', value: index + 1 };
  },

  LCase: (str: VbValue): VbValue => {
    return { type: 'String', value: toString(str).toLowerCase() };
  },

  UCase: (str: VbValue): VbValue => {
    return { type: 'String', value: toString(str).toUpperCase() };
  },

  LTrim: (str: VbValue): VbValue => {
    return { type: 'String', value: toString(str).replace(/^\s+/, '') };
  },

  RTrim: (str: VbValue): VbValue => {
    return { type: 'String', value: toString(str).replace(/\s+$/, '') };
  },

  Trim: (str: VbValue): VbValue => {
    return { type: 'String', value: toString(str).trim() };
  },

  Replace: (str: VbValue, find: VbValue, replace: VbValue, start?: VbValue, count?: VbValue, compare?: VbValue): VbValue => {
    const s = toString(str);
    const findStr = toString(find);
    const replaceStr = toString(replace);
    const startPos = start ? Math.max(1, Math.floor(toNumber(start))) - 1 : 0;
    const maxCount = count ? Math.floor(toNumber(count)) : -1;
    const compareMode = compare ? Math.floor(toNumber(compare)) : 0;
    
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
    const count = Math.max(0, Math.floor(toNumber(number)));
    return { type: 'String', value: ' '.repeat(count) };
  },

  String: (number: VbValue, character: VbValue): VbValue => {
    const count = Math.max(0, Math.floor(toNumber(number)));
    const char = toString(character);
    const charCode = char.length > 0 ? char[0]! : ' ';
    return { type: 'String', value: charCode.repeat(count) };
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

  Split: (str: VbValue, delimiter?: VbValue, count?: VbValue, _compare?: VbValue): VbValue => {
    const s = toString(str);
    const delim = delimiter ? toString(delimiter) : ' ';
    const maxCount = count ? Math.floor(toNumber(count)) : -1;
    
    const parts = s.split(delim);
    let resultParts: string[];
    if (maxCount > 0 && parts.length > maxCount) {
      resultParts = parts.slice(0, maxCount - 1);
      resultParts.push(parts.slice(maxCount - 1).join(delim));
    } else {
      resultParts = parts;
    }
    
    const vbArray = createVbArrayFromValues(resultParts.map(p => ({ type: 'String', value: p } as VbValue)));
    return { type: 'Array', value: vbArray };
  },

  Join: (list: VbValue, delimiter?: VbValue): VbValue => {
    const delim = delimiter ? toString(delimiter) : ' ';
    if (list.type === 'Array') {
      const arr = list.value;
      if (typeof arr === 'object' && arr !== null && 'toArray' in arr) {
        const vbArray = arr as { toArray: () => VbValue[] };
        return { type: 'String', value: vbArray.toArray().map(v => toString(v)).join(delim) };
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
