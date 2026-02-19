import type { VbValue } from '../runtime/index.ts';
import { createVbValue, toNumber, toString, toBoolean, VbEmpty, VbNull, VbNothing, isNumeric, isEmpty, isNull, createVbArrayFromValues } from '../runtime/index.ts';

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
    }
    
    const index = str1.toLowerCase().indexOf(str2.toLowerCase(), start - 1);
    return { type: 'Long', value: index + 1 };
  },

  InStrRev: (str1: VbValue, str2: VbValue, start?: VbValue, compare?: VbValue): VbValue => {
    const s1 = toString(str1);
    const s2 = toString(str2);
    const startPos = start ? Math.floor(toNumber(start)) : s1.length;
    const index = s1.toLowerCase().lastIndexOf(s2.toLowerCase(), startPos - 1);
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
    
    let result = '';
    let replaced = 0;
    let i = startPos;
    const lowerS = s.toLowerCase();
    const lowerFind = findStr.toLowerCase();
    
    while (i < s.length) {
      if (maxCount !== -1 && replaced >= maxCount) {
        result += s.substring(i);
        break;
      }
      if (lowerS.substring(i, i + findStr.length) === lowerFind) {
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
    const charCode = char.length > 0 ? char[0] : ' ';
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
    } else {
      result = s1.localeCompare(s2);
    }
    return { type: 'Integer', value: result };
  },

  Split: (str: VbValue, delimiter?: VbValue, count?: VbValue, compare?: VbValue): VbValue => {
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
};
