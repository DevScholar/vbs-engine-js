import type { VbValue } from '../runtime/index.ts';
import { createVbValue, toNumber, toString, VbEmpty } from '../runtime/index.ts';

export const mathFunctions = {
  Abs: (number: VbValue): VbValue => {
    const num = toNumber(number);
    return createVbValue(Math.abs(num));
  },

  Sgn: (number: VbValue): VbValue => {
    const num = toNumber(number);
    return { type: 'Integer', value: num > 0 ? 1 : num < 0 ? -1 : 0 };
  },

  Sqr: (number: VbValue): VbValue => {
    const num = toNumber(number);
    if (num < 0) {
      throw new Error('Invalid procedure call or argument: Sqr');
    }
    return { type: 'Double', value: Math.sqrt(num) };
  },

  Int: (number: VbValue): VbValue => {
    const num = toNumber(number);
    return { type: 'Long', value: Math.floor(num) };
  },

  Fix: (number: VbValue): VbValue => {
    const num = toNumber(number);
    return { type: 'Long', value: num >= 0 ? Math.floor(num) : Math.ceil(num) };
  },

  Round: (number: VbValue, numDecimalPlaces?: VbValue): VbValue => {
    const num = toNumber(number);
    const decimals = numDecimalPlaces ? Math.floor(toNumber(numDecimalPlaces)) : 0;
    const factor = Math.pow(10, decimals);
    return { type: 'Double', value: Math.round(num * factor) / factor };
  },

  Atn: (number: VbValue): VbValue => {
    const num = toNumber(number);
    return { type: 'Double', value: Math.atan(num) };
  },

  Cos: (number: VbValue): VbValue => {
    const num = toNumber(number);
    return { type: 'Double', value: Math.cos(num) };
  },

  Sin: (number: VbValue): VbValue => {
    const num = toNumber(number);
    return { type: 'Double', value: Math.sin(num) };
  },

  Tan: (number: VbValue): VbValue => {
    const num = toNumber(number);
    return { type: 'Double', value: Math.tan(num) };
  },

  Exp: (number: VbValue): VbValue => {
    const num = toNumber(number);
    return { type: 'Double', value: Math.exp(num) };
  },

  Log: (number: VbValue): VbValue => {
    const num = toNumber(number);
    if (num <= 0) {
      throw new Error('Invalid procedure call or argument: Log');
    }
    return { type: 'Double', value: Math.log(num) };
  },

  Rnd: (number?: VbValue): VbValue => {
    return { type: 'Single', value: Math.random() };
  },

  Randomize: (number?: VbValue): VbValue => {
    return VbEmpty;
  },

  Oct: (number: VbValue): VbValue => {
    const num = Math.floor(toNumber(number));
    return { type: 'String', value: num.toString(8) };
  },

  Hex: (number: VbValue): VbValue => {
    const num = Math.floor(toNumber(number));
    return { type: 'String', value: num.toString(16).toUpperCase() };
  },
};

export const constants: Record<string, VbValue> = {
  vbCr: { type: 'String', value: '\r' },
  vbCrLf: { type: 'String', value: '\r\n' },
  vbFormFeed: { type: 'String', value: '\f' },
  vbLf: { type: 'String', value: '\n' },
  vbNewLine: { type: 'String', value: '\r\n' },
  vbNullChar: { type: 'String', value: '\0' },
  vbTab: { type: 'String', value: '\t' },
  vbVerticalTab: { type: 'String', value: '\v' },
  vbNullString: { type: 'String', value: '' },
  vbObjectError: { type: 'Long', value: -2147221504 },
  
  vbSunday: { type: 'Integer', value: 1 },
  vbMonday: { type: 'Integer', value: 2 },
  vbTuesday: { type: 'Integer', value: 3 },
  vbWednesday: { type: 'Integer', value: 4 },
  vbThursday: { type: 'Integer', value: 5 },
  vbFriday: { type: 'Integer', value: 6 },
  vbSaturday: { type: 'Integer', value: 7 },
  
  vbUseSystemDayOfWeek: { type: 'Integer', value: 0 },
  vbFirstJan1: { type: 'Integer', value: 1 },
  vbFirstFourDays: { type: 'Integer', value: 2 },
  vbFirstFullWeek: { type: 'Integer', value: 3 },
  
  vbBinaryCompare: { type: 'Integer', value: 0 },
  vbTextCompare: { type: 'Integer', value: 1 },
  vbDatabaseCompare: { type: 'Integer', value: 2 },
  
  vbGeneralDate: { type: 'Integer', value: 0 },
  vbLongDate: { type: 'Integer', value: 1 },
  vbShortDate: { type: 'Integer', value: 2 },
  vbLongTime: { type: 'Integer', value: 3 },
  vbShortTime: { type: 'Integer', value: 4 },
  
  vbEmpty: { type: 'Integer', value: 0 },
  vbNull: { type: 'Integer', value: 1 },
  vbInteger: { type: 'Integer', value: 2 },
  vbLong: { type: 'Integer', value: 3 },
  vbSingle: { type: 'Integer', value: 4 },
  vbDouble: { type: 'Integer', value: 5 },
  vbCurrency: { type: 'Integer', value: 6 },
  vbDate: { type: 'Integer', value: 7 },
  vbString: { type: 'Integer', value: 8 },
  vbObject: { type: 'Integer', value: 9 },
  vbError: { type: 'Integer', value: 10 },
  vbBoolean: { type: 'Integer', value: 11 },
  vbVariant: { type: 'Integer', value: 12 },
  vbDataObject: { type: 'Integer', value: 13 },
  vbDecimal: { type: 'Integer', value: 14 },
  vbByte: { type: 'Integer', value: 17 },
  vbArray: { type: 'Integer', value: 8192 },

  vbOKOnly: { type: 'Integer', value: 0 },
  vbOKCancel: { type: 'Integer', value: 1 },
  vbAbortRetryIgnore: { type: 'Integer', value: 2 },
  vbYesNoCancel: { type: 'Integer', value: 3 },
  vbYesNo: { type: 'Integer', value: 4 },
  vbRetryCancel: { type: 'Integer', value: 5 },

  vbCritical: { type: 'Integer', value: 16 },
  vbQuestion: { type: 'Integer', value: 32 },
  vbExclamation: { type: 'Integer', value: 48 },
  vbInformation: { type: 'Integer', value: 64 },

  vbDefaultButton1: { type: 'Integer', value: 0 },
  vbDefaultButton2: { type: 'Integer', value: 256 },
  vbDefaultButton3: { type: 'Integer', value: 512 },

  vbOK: { type: 'Integer', value: 1 },
  vbCancel: { type: 'Integer', value: 2 },
  vbAbort: { type: 'Integer', value: 3 },
  vbRetry: { type: 'Integer', value: 4 },
  vbIgnore: { type: 'Integer', value: 5 },
  vbYes: { type: 'Integer', value: 6 },
  vbNo: { type: 'Integer', value: 7 },
};
