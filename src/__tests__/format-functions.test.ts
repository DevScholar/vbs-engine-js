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

// The rules a wider probe pinned down over all 132 locales, for the arguments themselves and for
// FormatDateTime.
//
//   * The three tristate flags are -1 (True), 0 (False) and -2 (the locale's own) and nothing else, and the
//     value is rounded to a whole number with a tie going to the even neighbour first: 1, 2, 3, -3 and -3.5
//     are all error 5, 0.5 is False because it rounds to 0, and -1.5 is True because it rounds to -2 is not
//     True but the locale's own.
//   * The fraction count is rounded the same way: 1.5 keeps two digits, 1.4 keeps one, -0.5 keeps none and
//     -1.5 is error 5.
//   * A Nothing is error 91 wherever a number is read, and a Null for the fraction count is error 94.
//   * FormatDateTime reads its date the way CDate does, so a number is a day count and a text is read with
//     the locale's own order and calendar; a Null is error 13 and a Nothing is 91.
//   * The named format is rounded the same way and must land between 0 and 4: 2.7 is the long time and 4.5
//     the short time, while 5, -1 and True are error 5.
const ARGUMENT_CELLS: ReadonlyArray<readonly [number, string, string]> = [
  [1033, 'FormatNumber(0.5, 2, 1)', 'ERR 5'],
  [1033, 'FormatNumber(0.5, 2, 2)', 'ERR 5'],
  [1033, 'FormatNumber(0.5, 2, -3)', 'ERR 5'],
  [1033, 'FormatNumber(0.5, 2, -3.5)', 'ERR 5'],
  [1033, 'FormatNumber(0.5, 2, "1")', 'ERR 5'],
  [1033, 'FormatNumber(0.5, 2, 0.5)', '.50'],
  [1033, 'FormatNumber(0.5, 2, -0.5)', '.50'],
  [1033, 'FormatNumber(0.5, 2, -1.5)', '0.50'],
  [1033, 'FormatNumber(0.5, 2, -2.5)', '0.50'],
  [1033, 'FormatNumber(0.5, 2, Empty)', '.50'],
  [1033, 'FormatNumber(-1234.5, 2, -1, 1)', 'ERR 5'],
  [1033, 'FormatNumber(-1234.5, 2, -1, -0.5)', '-1,234.50'],
  [1033, 'FormatNumber(1234.5, 2, -1, -1, 1)', 'ERR 5'],
  [1033, 'FormatNumber(1234.5, 2, -1, -1, -0.5)', '1234.50'],
  [1033, 'FormatNumber(1234.5, 1.5)', '1,234.50'],
  [1033, 'FormatNumber(1234.5, 1.4)', '1,234.5'],
  [1033, 'FormatNumber(1234.5, -0.5)', '1,235'],
  [1033, 'FormatNumber(1234.5, -1.5)', 'ERR 5'],
  [1033, 'FormatNumber(1234.5, Null)', 'ERR 94'],
  [1033, 'FormatNumber(1234.5, "x")', 'ERR 13'],
  [1033, 'FormatNumber(Nothing)', 'ERR 91'],
  [1033, 'FormatNumber(1234.5, Nothing)', 'ERR 91'],
  [1033, 'FormatNumber(1234.5, 2, Nothing)', 'ERR 91'],
  [1033, 'FormatCurrency(Nothing)', 'ERR 91'],
  [1033, 'FormatPercent(Nothing)', 'ERR 91'],

  [1033, 'FormatDateTime(DateSerial(2026, 1, 5) + TimeSerial(13, 5, 9))', '1/5/2026 1:05:09 PM'],
  [1033, 'FormatDateTime(DateSerial(2026, 1, 5) + TimeSerial(13, 5, 9), 4.5)', '13:05'],
  [
    1033,
    'FormatDateTime(DateSerial(2026, 1, 5) + TimeSerial(13, 5, 9), 0.4)',
    '1/5/2026 1:05:09 PM',
  ],
  [1033, 'FormatDateTime(DateSerial(2026, 1, 5) + TimeSerial(13, 5, 9), "2.7")', '1:05:09 PM'],
  [1033, 'FormatDateTime(DateSerial(2026, 1, 5) + TimeSerial(13, 5, 9), 5)', 'ERR 5'],
  [1033, 'FormatDateTime(DateSerial(2026, 1, 5) + TimeSerial(13, 5, 9), -1)', 'ERR 5'],
  [1033, 'FormatDateTime(DateSerial(2026, 1, 5) + TimeSerial(13, 5, 9), True)', 'ERR 5'],
  [1033, 'FormatDateTime(DateSerial(2026, 1, 5) + TimeSerial(13, 5, 9), Null)', 'ERR 94'],
  [1033, 'FormatDateTime(DateSerial(2026, 1, 5) + TimeSerial(13, 5, 9), Nothing)', 'ERR 91'],
  [1033, 'FormatDateTime(DateSerial(2026, 1, 5) + TimeSerial(13, 5, 9), "x")', 'ERR 13'],
  [1033, 'FormatDateTime(45000)', '3/15/2023'],
  [1033, 'FormatDateTime(-1)', '12/29/1899'],
  [1033, 'FormatDateTime(Empty)', '12:00:00 AM'],
  [1033, 'FormatDateTime(True)', '12/29/1899'],
  [1033, 'FormatDateTime("1/5/2026")', '1/5/2026'],
  [1033, 'FormatDateTime("nonsense")', 'ERR 13'],
  [1033, 'FormatDateTime(Null)', 'ERR 13'],
  [1033, 'FormatDateTime(Nothing)', 'ERR 91'],
  // Under ar-SA a slash date is read day first, so 1 May and not 5 January, which the engine used to read
  // with the JavaScript date parser and get the other way round.
  [1025, 'FormatDateTime("1/5/2026")', '01/05/2026'],
  // Under th-TH a written year is a Buddhist one, so the text names the Gregorian 1483, and its short date
  // is one of the days Windows refuses.
  [1054, 'FormatDateTime("2026-01-05 13:05:09", 0)', '5/1/2026 13:05:09'],
  [1054, 'FormatDateTime("2026-01-05 13:05:09", 2)', 'ERR 5'],
];

describe('the Format functions read their arguments the way VBScript does', () => {
  for (const [lcid, expression, expected] of ARGUMENT_CELLS) {
    it(`under ${lcid}, ${expression} is ${expected}`, () => {
      expect(answer(expression, lcid)).toBe(expected);
    });
  }
});
