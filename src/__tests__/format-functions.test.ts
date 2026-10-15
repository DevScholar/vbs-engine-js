import { describe, it, expect } from 'vitest';
import { VbsEngine } from '../index.ts';

// FormatNumber, FormatCurrency and FormatPercent with their optional arguments, measured against a fresh
// cscript under 1031 and 1033.
//
// The three arguments after the fraction count are the tristate ones: -1 yes, 0 no, and anything else the
// locale's own setting. Measured:
//
//   * The fraction count takes any value from -1 up, where -1 means the locale's own count, and a value
//     below that is error 5 rather than a count.
//   * includeLeadingDigit False drops the zero before the fraction mark, so "0,50" becomes ",50".
//   * useParensForNegativeNumbers changes only the sign of the shape the locale already uses: True puts the
//     amount in parentheses and drops the minus, so a German amount becomes "(1.234,50 €)" and a percentage
//     "(50,00%)", while False writes a minus, which turns the American "($1,234.50)" into "-$1,234.50".
//   * groupDigits False takes the group marks out ("1234,57") and True puts them in.
//   * A Null is a type mismatch here, error 13, and not the invalid use of Null that most functions raise.
//   * The rounding is away from zero, so FormatNumber(2.5, 0) is 3 where CInt(2.5) is 2.
const CELLS: ReadonlyArray<readonly [number, string, string]> = [
  [1031, 'FormatNumber(1234.567)', '1.234,57'],
  [1031, 'FormatNumber(1234.567, 0)', '1.235'],
  [1031, 'FormatNumber(1234.567, 1)', '1.234,6'],
  [1031, 'FormatNumber(1234.567, 3)', '1.234,567'],
  [1031, 'FormatNumber(1234.567, 5)', '1.234,56700'],
  [1031, 'FormatNumber(1234.567, 20)', '1.234,56700000000000000000'],
  [1031, 'FormatNumber(1234.567, -1)', '1.234,57'],
  [1031, 'FormatNumber(1234.567, -2)', 'ERR 5'],
  [1031, 'FormatNumber(2.5, 0)', '3'],
  [1031, 'FormatNumber(1.005, 2)', '1,01'],
  [1031, 'FormatNumber(0.5, 2, -1)', '0,50'],
  [1031, 'FormatNumber(0.5, 2, 0)', ',50'],
  [1031, 'FormatNumber(0.5, 2, -2)', '0,50'],
  [1031, 'FormatNumber(0.5, 2, True)', '0,50'],
  [1031, 'FormatNumber(-1234.5, 2, -1, -1)', '(1.234,50)'],
  [1031, 'FormatNumber(-1234.5, 2, -1, 0)', '-1.234,50'],
  [1031, 'FormatNumber(-1234.5, 2, -1, -2)', '-1.234,50'],
  [1031, 'FormatNumber(1234.567, 2, -1, -1, -1)', '1.234,57'],
  [1031, 'FormatNumber(1234.567, 2, -1, -1, 0)', '1234,57'],
  [1031, 'FormatNumber(1234.567, 2, -1, -1, -2)', '1.234,57'],
  [1031, 'FormatNumber(123.456, 2, -1, -1, -1)', '123,46'],
  [1031, 'FormatNumber(-1234.567, 2, -1, 0, -1)', '-1.234,57'],
  [1031, 'FormatNumber(True)', '-1,00'],
  [1031, 'FormatNumber(Empty)', '0,00'],
  [1031, 'FormatNumber(0)', '0,00'],
  [1031, 'FormatNumber(Null)', 'ERR 13'],
  [1031, 'FormatCurrency(-1234.5)', '-1.234,50 €'],
  [1031, 'FormatCurrency(-1234.5, 2, -1, -1, -1)', '(1.234,50 €)'],
  [1031, 'FormatCurrency(-1234.5, 2, -1, 0, -1)', '-1.234,50 €'],
  [1031, 'FormatCurrency(0)', '0,00 €'],
  [1031, 'FormatCurrency(1234.5, 0)', '1.235 €'],
  [1031, 'FormatPercent(0.5)', '50,00%'],
  [1031, 'FormatPercent(-0.5, 2, -1, -1, -1)', '(50,00%)'],
  [1031, 'FormatPercent(0.5, 0)', '50%'],
  [1033, 'FormatNumber(1234.567)', '1,234.57'],
  [1033, 'FormatNumber(1234.567, -2)', 'ERR 5'],
  [1033, 'FormatNumber(0.5, 2, 0)', '.50'],
  [1033, 'FormatNumber(Null)', 'ERR 13'],
  [1033, 'FormatCurrency(-1234.5)', '($1,234.50)'],
  [1033, 'FormatCurrency(-1234.5, 2, -1, -1, -1)', '($1,234.50)'],
  [1033, 'FormatCurrency(-1234.5, 2, -1, 0, -1)', '-$1,234.50'],
  [1033, 'FormatPercent(-0.5, 2, -1, -1, -1)', '(50.00%)'],
];

function answer(expression: string, lcid: number): unknown {
  const engine = new VbsEngine();
  try {
    engine.executeStatement(`SetLocale ${lcid}\nr = (${expression})`);
    const failure = engine.error;
    if (failure) return 'ERR ' + String(failure.number);
    return engine.eval('r');
  } finally {
    engine.executeStatement('SetLocale 1033');
  }
}

describe('the Format functions take the arguments VBScript takes', () => {
  for (const [lcid, expression, expected] of CELLS) {
    it(`under ${lcid}, ${expression} is ${expected}`, () => {
      expect(answer(expression, lcid)).toBe(expected);
    });
  }
});
