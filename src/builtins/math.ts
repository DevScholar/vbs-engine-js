import type { VbValue } from '../runtime/index.ts';
import {
  absoluteScaled,
  CURRENCY_SCALE,
  createVbValue,
  floorScaled,
  roundScaledToMultiple,
  toCurrencyScaled,
  toNumber,
  toSinglePrecision,
  truncateScaled,
  VbEmpty,
} from '../runtime/index.ts';

/**
 * The result of a function that keeps the type of the operand it was given.
 *
 * Measured against a fresh cscript: Abs, Int, Fix and Round answer with a Single for a Single and a
 * Currency for a Currency, while Sgn answers with an Integer whatever it is handed.
 */
function sameNumericType(operand: VbValue, value: number): VbValue {
  if (operand.type === 'Single') return { type: 'Single', value: toSinglePrecision(value) };
  if (operand.type === 'Currency') return { type: 'Currency', value: toCurrencyScaled(value) };
  if (operand.type === 'Double') return { type: 'Double', value };
  return createVbValue(value);
}

export const mathFunctions = {
  Abs: (number: VbValue): VbValue => {
    if (number.type === 'Currency')
      return { type: 'Currency', value: absoluteScaled(number.value) };
    const num = toNumber(number);
    return sameNumericType(number, Math.abs(num));
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
    if (number.type === 'Currency') return { type: 'Currency', value: floorScaled(number.value) };
    const num = toNumber(number);
    if (number.type === 'Single')
      return { type: 'Single', value: toSinglePrecision(Math.floor(num)) };
    return { type: 'Long', value: Math.floor(num) };
  },

  Fix: (number: VbValue): VbValue => {
    if (number.type === 'Currency')
      return { type: 'Currency', value: truncateScaled(number.value) };
    const num = toNumber(number);
    const fixed = num >= 0 ? Math.floor(num) : Math.ceil(num);
    if (number.type === 'Single') return { type: 'Single', value: toSinglePrecision(fixed) };
    return { type: 'Long', value: fixed };
  },

  Round: (number: VbValue, numDecimalPlaces?: VbValue): VbValue => {
    const decimals = numDecimalPlaces ? Math.floor(toNumber(numDecimalPlaces)) : 0;
    // A Currency rounds on its own scale, so a tie goes to the even neighbour with no double in between:
    // Round(CCur(1.005), 2) is the Currency 1, and Round(CCur(1234.5678), 2) is 1234.57.
    if (number.type === 'Currency') {
      const unit =
        decimals >= 0
          ? CURRENCY_SCALE / 10n ** BigInt(decimals)
          : CURRENCY_SCALE * 10n ** BigInt(-decimals);
      return { type: 'Currency', value: roundScaledToMultiple(number.value, unit) };
    }
    const num = toNumber(number);
    const factor = Math.pow(10, decimals);
    const scaled = num * factor;

    // A tie goes to the EVEN neighbour, not upwards: 0.5 -> 0, 1.5 -> 2, 2.5 -> 2.
    //
    // Only a tie. The scaled value is deliberately left as the double it is, with no attempt to
    // recover the decimal the literal was written as: VBScript does not do that either, and
    // Round(1.015, 2) is the case that says so -- 1.015 * 100 is 101.49999999999999, just under
    // the tie, and the real engine answers 1.01 rather than the 1.02 a decimal reading gives.
    const below = Math.floor(scaled);
    const fraction = scaled - below;
    let rounded: number;
    if (fraction > 0.5) {
      rounded = below + 1;
    } else if (fraction < 0.5) {
      rounded = below;
    } else {
      // -3 % 2 is -1 in JavaScript, so this reads as odd and steps up to -2, which is right.
      rounded = below % 2 === 0 ? below : below + 1;
    }

    return sameNumericType(number, rounded / factor);
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

  Rnd: (_number?: VbValue): VbValue => {
    void _number; // Intentionally unused - matches VBScript signature
    return { type: 'Single', value: Math.random() };
  },

  Randomize: (_number?: VbValue): VbValue => {
    void _number; // Intentionally unused - matches VBScript signature
    return VbEmpty;
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
