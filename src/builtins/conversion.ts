import type { VbValue } from '../runtime/index.ts';
import {
  checkLongLongRange,
  createVbError,
  roundHalfToEven,
  toCurrencyScaled,
  toLongLong,
  toNumber,
  toSinglePrecision,
  toString,
  toBoolean,
  toVbDateForCdate,
  rejectObjectArgument,
  VbErrorCodes,
  isEmpty,
  isNull,
  isNumeric,
  parseVbDate,
  formatVbDateTime,
  VbNull,
} from '../runtime/index.ts';
import { currentLcid, getCurrentBCP47Locale } from './locale.ts';
import { formatNlsNumber } from '../runtime/nls-number.ts';

/**
 * The digits Hex and Oct write, with the two's complement width the value's type asks for.
 *
 * Measured against a fresh cscript. A negative value is written as its two's complement, and the width
 * follows the type as well as the size: an Integer, a Boolean or a Byte fills sixteen bits while its
 * magnitude fits a signed sixteen bit value (Hex(-1) is "FFFF", Oct(-1) is "177777"), and everything else
 * fills thirty-two (Hex(CLng(-1)), Hex(CDbl(-1)) and Hex(CCur(-1)) are all "FFFFFFFF", and Oct(CCur(-1)) is
 * "37777777777"). A value of -32768 is written in thirty-two bits even as an Integer, because its magnitude
 * no longer fits: Hex(-32768) is "FFFF8000" and Oct(-32768) is "37777700000".
 *
 * The value is rounded to a whole number first, a tie going to the even neighbour (Hex(0.5) is "0", Hex(1.5)
 * is "2", Hex(2.5) is "2"), and a value outside the Long range is Overflow: Hex(4294967295) and
 * Hex(CCur(2147483648)) both refuse. A Null answers Null, which a fresh cscript shows for Hex(Null) and
 * Oct(Null) alike.
 */
function toBaseDigits(number: VbValue, radix: 8 | 16, functionName: string): VbValue {
  if (number.type === 'Null') return VbNull;
  const value = roundHalfToEven(toNumber(number));
  if (value < -2147483648 || value > 2147483647) {
    throw createVbError(VbErrorCodes.Overflow, `Overflow: ${functionName}`);
  }
  if (value >= 0) return { type: 'String', value: value.toString(radix).toUpperCase() };
  const narrowType =
    number.type === 'Integer' || number.type === 'Boolean' || number.type === 'Byte';
  const width = narrowType && value >= -32767 ? 16 : 32;
  return { type: 'String', value: (2 ** width + value).toString(radix).toUpperCase() };
}

/**
 * The number a Format function was handed.
 *
 * A Null is a type mismatch here rather than the invalid use of Null most functions raise, which a fresh
 * cscript shows: FormatNumber(Null) is error 13 and not 94. A Nothing is the object variable that is not
 * set, error 91, and any other object is 438.
 */
function toFormattedNumber(expression: VbValue, functionName: string): number {
  rejectObjectArgument(expression, functionName);
  if (expression.type === 'Null') {
    throw createVbError(VbErrorCodes.TypeMismatch, `Type mismatch: ${functionName} Null`);
  }
  return toNumber(expression);
}

/**
 * The fraction count a Format function was handed, or -1 for the locale's own.
 *
 * The three settings are the tristate ones, so -1 and anything above it are a count to use, and a number
 * below -1 is not a setting at all: a fresh cscript answers FormatNumber(1.25, -2) with error 5. The value
 * is rounded to a whole number first, a tie going to the even neighbour, which the measurements show:
 * FormatNumber(1234.5, 1.5) keeps two digits and FormatNumber(1234.5, -0.5) keeps none.
 */
function readFractionCount(value: VbValue | undefined, functionName: string): number {
  if (!value) return -1;
  rejectObjectArgument(value, functionName);
  const digits = roundHalfToEven(toNumber(value));
  if (digits < -1) {
    throw createVbError(
      VbErrorCodes.InvalidProcedureCall,
      `Invalid procedure call: ${functionName}`
    );
  }
  return digits;
}

/**
 * One of the three tristate flags, which a fresh cscript accepts as -1 (True), 0 (False) or -2 (the
 * locale's own) and refuses anything else with error 5.
 *
 * The value is rounded to a whole number first, a tie going to the even neighbour: FormatNumber(0.5, 2, 1)
 * and FormatNumber(0.5, 2, -3) are error 5, while -2.5 is the locale's own because it rounds to -2 and
 * -0.5 is False because it rounds to 0. A flag the script leaves out is the locale's own as well.
 */
function readTriState(value: VbValue | undefined, functionName: string): number {
  if (value === undefined) return -2;
  rejectObjectArgument(value, functionName);
  const mode = roundHalfToEven(toNumber(value));
  if (mode !== -2 && mode !== -1 && mode !== 0) {
    throw createVbError(
      VbErrorCodes.InvalidProcedureCall,
      `Invalid procedure call: ${functionName}`
    );
  }
  return mode;
}

/**
 * The named format FormatDateTime was handed, which is 0 to 4 and nothing else.
 *
 * A fresh cscript rounds it to a whole number, a tie going to the even neighbour, so 2.7 is the long time
 * and 4.5 is the short time, and refuses anything outside 0 to 4 with error 5: FormatDateTime(x, 5) and
 * FormatDateTime(x, -1) are both error 5. A Null is the invalid use of Null here, error 94.
 */
function readNamedFormat(value: VbValue | undefined, functionName: string): number {
  if (value === undefined) return 0;
  rejectObjectArgument(value, functionName);
  const style = roundHalfToEven(toNumber(value));
  if (style < 0 || style > 4) {
    throw createVbError(
      VbErrorCodes.InvalidProcedureCall,
      `Invalid procedure call: ${functionName}`
    );
  }
  return style;
}

export const conversionFunctions = {
  CBool: (expression: VbValue): VbValue => {
    return { type: 'Boolean', value: toBoolean(expression) };
  },

  CByte: (expression: VbValue): VbValue => {
    const num = roundHalfToEven(toNumber(expression));
    if (num < 0 || num > 255) {
      throw createVbError(VbErrorCodes.Overflow, 'Overflow: CByte');
    }
    return { type: 'Byte', value: num };
  },

  CCur: (expression: VbValue): VbValue => {
    return { type: 'Currency', value: toCurrencyScaled(toNumber(expression)) };
  },

  CDate: (expression: VbValue): VbValue => {
    return { type: 'Date', value: toVbDateForCdate(expression) };
  },

  CDbl: (expression: VbValue): VbValue => {
    return { type: 'Double', value: toNumber(expression) };
  },

  CInt: (expression: VbValue): VbValue => {
    const num = roundHalfToEven(toNumber(expression));
    if (num < -32768 || num > 32767) {
      throw createVbError(VbErrorCodes.Overflow, 'Overflow: CInt');
    }
    return { type: 'Integer', value: num };
  },

  CLng: (expression: VbValue): VbValue => {
    const num = roundHalfToEven(toNumber(expression));
    if (num < -2147483648 || num > 2147483647) {
      throw createVbError(VbErrorCodes.Overflow, 'Overflow: CLng');
    }
    return { type: 'Long', value: num };
  },

  CLngLng: (expression: VbValue): VbValue => {
    if (expression.type === 'LongLong') return expression;
    if (expression.type === 'String') {
      // A string that is a whole number is read as one, so a value past what a double counts with is not
      // rounded on the way in. Anything else, such as "1.5", goes through a number.
      const written = expression.value.trim();
      try {
        return { type: 'LongLong', value: checkLongLongRange(BigInt(written)) };
      } catch {
        const num = parseFloat(written);
        if (isNaN(num)) throw createVbError(VbErrorCodes.TypeMismatch, 'Type mismatch: CLngLng');
        return { type: 'LongLong', value: toLongLong(num) };
      }
    }
    return { type: 'LongLong', value: toLongLong(toNumber(expression)) };
  },

  CSng: (expression: VbValue): VbValue => {
    return { type: 'Single', value: toSinglePrecision(toNumber(expression)) };
  },

  CStr: (expression: VbValue): VbValue => {
    return { type: 'String', value: toString(expression) };
  },

  CVar: (expression: VbValue): VbValue => {
    return expression;
  },

  CVErr: (errorNumber: VbValue): VbValue => {
    return { type: 'Error', value: toNumber(errorNumber) };
  },

  Val: (string: VbValue): VbValue => {
    const str = toString(string).trim();
    const match = str.match(/^[-+]?[0-9]*\.?[0-9]+([eE][-+]?[0-9]+)?/);
    if (match) {
      return { type: 'Double', value: parseFloat(match[0]) };
    }
    return { type: 'Double', value: 0 };
  },

  Str: (number: VbValue): VbValue => {
    const num = toNumber(number);
    const str = String(num);
    if (num >= 0) {
      return { type: 'String', value: ' ' + str };
    }
    return { type: 'String', value: str };
  },

  FormatNumber: (
    expression: VbValue,
    numDigitsAfterDecimal?: VbValue,
    includeLeadingDigit?: VbValue,
    useParensForNegativeNumbers?: VbValue,
    groupDigits?: VbValue
  ): VbValue => {
    const num = toFormattedNumber(expression, 'FormatNumber');
    const digits = readFractionCount(numDigitsAfterDecimal, 'FormatNumber');
    const leadingDigit = readTriState(includeLeadingDigit, 'FormatNumber');
    const parens = readTriState(useParensForNegativeNumbers, 'FormatNumber');
    const group = readTriState(groupDigits, 'FormatNumber');

    return {
      type: 'String',
      value: formatNlsNumber(num, currentLcid(), 'number', { digits, leadingDigit, parens, group }),
    };
  },

  FormatCurrency: (
    expression: VbValue,
    numDigitsAfterDecimal?: VbValue,
    includeLeadingDigit?: VbValue,
    useParensForNegativeNumbers?: VbValue,
    groupDigits?: VbValue
  ): VbValue => {
    const num = toFormattedNumber(expression, 'FormatCurrency');
    const digits = readFractionCount(numDigitsAfterDecimal, 'FormatCurrency');
    const leadingDigit = readTriState(includeLeadingDigit, 'FormatCurrency');
    const parens = readTriState(useParensForNegativeNumbers, 'FormatCurrency');
    const group = readTriState(groupDigits, 'FormatCurrency');

    return {
      type: 'String',
      value: formatNlsNumber(num, currentLcid(), 'currency', {
        digits,
        leadingDigit,
        parens,
        group,
      }),
    };
  },

  FormatPercent: (
    expression: VbValue,
    numDigitsAfterDecimal?: VbValue,
    includeLeadingDigit?: VbValue,
    useParensForNegativeNumbers?: VbValue,
    groupDigits?: VbValue
  ): VbValue => {
    const num = toFormattedNumber(expression, 'FormatPercent');
    const digits = readFractionCount(numDigitsAfterDecimal, 'FormatPercent');
    const leadingDigit = readTriState(includeLeadingDigit, 'FormatPercent');
    const parens = readTriState(useParensForNegativeNumbers, 'FormatPercent');
    const group = readTriState(groupDigits, 'FormatPercent');

    return {
      type: 'String',
      value: formatNlsNumber(num, currentLcid(), 'percent', {
        digits,
        leadingDigit,
        parens,
        group,
      }),
    };
  },

  FormatDateTime: (date: VbValue, namedFormat?: VbValue): VbValue => {
    // The date is read the way CDate reads one, and not by the JavaScript Date parser, which a fresh cscript
    // shows from every side: FormatDateTime(45000) is 15 March 2023 and not the year 45000, FormatDateTime(-1)
    // is 29 December 1899, and a text is read with the locale's own date order and calendar, so under ar-SA
    // "1/5/2026" is 1 May and under th-TH "2026-01-05" is the Buddhist year 2026, which is Gregorian 1483.
    rejectObjectArgument(date, 'FormatDateTime');
    if (date.type === 'Null') {
      throw createVbError(VbErrorCodes.TypeMismatch, 'Type mismatch: FormatDateTime Null');
    }
    const parsed = date.type === 'Date' ? (date.value as Date) : toVbDateForCdate(date);
    const format = readNamedFormat(namedFormat, 'FormatDateTime');
    return {
      type: 'String',
      value: formatVbDateTime(parsed, currentLcid(), format, getCurrentBCP47Locale()),
    };
  },

  Hex: (number: VbValue): VbValue => toBaseDigits(number, 16, 'Hex'),

  Oct: (number: VbValue): VbValue => toBaseDigits(number, 8, 'Oct'),
};

export const inspectionFunctions = {
  IsArray: (expression: VbValue): VbValue => {
    return { type: 'Boolean', value: expression.type === 'Array' };
  },

  IsDate: (expression: VbValue): VbValue => {
    if (expression.type === 'Date') {
      return { type: 'Boolean', value: true };
    }
    if (expression.type === 'String') {
      return {
        type: 'Boolean',
        value: parseVbDate(expression.value, getCurrentBCP47Locale()) !== null,
      };
    }
    return { type: 'Boolean', value: false };
  },

  IsEmpty: (expression: VbValue): VbValue => {
    return { type: 'Boolean', value: isEmpty(expression) };
  },

  IsNull: (expression: VbValue): VbValue => {
    return { type: 'Boolean', value: isNull(expression) };
  },

  IsNumeric: (expression: VbValue): VbValue => {
    return { type: 'Boolean', value: isNumeric(expression) };
  },

  IsObject: (expression: VbValue): VbValue => {
    return { type: 'Boolean', value: expression.type === 'Object' };
  },

  VarType: (expression: VbValue): VbValue => {
    const typeMap: Record<string, number> = {
      Empty: 0,
      Null: 1,
      Integer: 2,
      Long: 3,
      Single: 4,
      Double: 5,
      Currency: 6,
      Date: 7,
      String: 8,
      Object: 9,
      Error: 10,
      Boolean: 11,
      Variant: 12,
      Byte: 17,
      LongLong: 20,
      // An array of variants is VT_ARRAY with VT_VARIANT, which a fresh cscript answers with 8204.
      Array: 8204,
    };
    return { type: 'Integer', value: typeMap[expression.type] ?? 12 };
  },

  TypeName: (expression: VbValue): VbValue => {
    if (expression.type === 'Object' && expression.value === null)
      return { type: 'String', value: 'Nothing' };
    // An object the engine's own classes made answers with its class name, as vbscript.dll does for a
    // Collection or a RegExp.
    if (expression.type === 'Object' && expression.value && typeof expression.value === 'object') {
      const obj = expression.value as { classInfo?: { name: string } };
      if (obj.classInfo && obj.classInfo.name) {
        return { type: 'String', value: obj.classInfo.name };
      }
    }
    const typeNames: Record<string, string> = {
      Empty: 'Empty',
      Null: 'Null',
      Integer: 'Integer',
      Long: 'Long',
      LongLong: 'LongLong',
      Single: 'Single',
      Double: 'Double',
      Currency: 'Currency',
      Date: 'Date',
      String: 'String',
      Object: 'Object',
      Error: 'Error',
      Boolean: 'Boolean',
      Variant: 'Variant',
      Byte: 'Byte',
      Array: 'Variant()',
    };
    return { type: 'String', value: typeNames[expression.type] ?? 'Variant' };
  },
};
