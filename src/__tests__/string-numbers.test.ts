import { describe, it, expect } from "vitest";
import { VbsEngine } from "../index.ts";

// A number written as a string, read with the marks the locale writes. Measured against a fresh cscript in
// all 132 locales Windows answers for.
//
//   * Windows reads four marks and not two: LOCALE_SDECIMAL and LOCALE_SMONDECIMALSEP are both the fraction
//     mark, and LOCALE_STHOUSAND and LOCALE_SMONTHOUSANDSEP both group. The two sets differ in three
//     locales, and the difference shows.
//   * The locale's own fraction mark is the fraction mark and its group marks come out, so under 1031
//     "1234.567" is 1234567 while "1,5" is 1.5, and under 1033 it is the other way round.
//   * A group mark is dropped only where a digit stands before it, which is why ",5" is a half under 1031
//     but a type mismatch under 1033.
//   * The locale's own currency symbol may stand before or after the amount, and only that one: "$1.5" is a
//     number under 1033 and a type mismatch under 1031, and the euro is the other way round.
//   * A pair of parentheses around the amount means a negative, an exponent is read with e, E, d or D, and
//     the VB literal prefixes stay readable.
//   * A plain space groups a number wherever the locale groups with a space of its own (1026, 1036, 2077,
//     3079), and nowhere else, so "1 234" is 1234 there and a type mismatch under 1031 or 1033. No other
//     blank counts anywhere: a tab is a type mismatch even under the locales that group with a space.
//   * The whole string has to be the number: "50%" is a type mismatch, and so are an empty string and a
//     string of spaces. A full stop the locale does not name as one of its two fraction marks is not a
//     number either, so "1.5" is a type mismatch under 1026 where it is 1,5 under its neighbours.
//   * A string beside a number is added rather than joined, and the result is a Double: "1" + 1 is the
//     Double 2, "1" + True is the Double 0 (True counts as -1), and "a" + 1 is a type mismatch. Two strings
//     still join, so "1" + "1" is the text "11".
const CELLS: ReadonlyArray<readonly [number, string, string]> = [
  [1031, 'CDbl("1234.567")', "1234567"],
  [1031, 'CDbl("1,5")', "1,5"],
  [1031, 'CDbl("1.5")', "15"],
  [1031, 'CDbl("1.234,5")', "1234,5"],
  [1031, 'CDbl("1,234.5")', "1,2345"],
  [1031, 'CDbl(" 1,5 ")', "1,5"],
  [1031, 'CDbl("1 234")', "ERR 13"],
  [1031, 'CDbl("0.5")', "5"],
  [1031, 'CDbl("1e3")', "1000"],
  [1031, 'CDbl("1d3")', "1000"],
  [1031, 'CDbl("&H10")', "16"],
  [1031, 'CDbl("50%")', "ERR 13"],
  [1031, 'CDbl("$1.5")', "ERR 13"],
  [1031, 'CDbl("1,5 €")', "1,5"],
  [1031, 'CDbl("")', "ERR 13"],
  [1031, 'CDbl("   ")', "ERR 13"],
  [1031, 'CDbl("abc")', "ERR 13"],
  [1031, 'CDbl("True")', "ERR 13"],
  [1031, 'CDbl("(1,5)")', "-1,5"],
  [1031, 'CDbl("1,5,5")', "ERR 13"],
  [1031, 'CDbl("1,5.")', "1,5"],
  [1031, 'CDbl(",5")', "0,5"],
  [1031, 'CDbl("1.234.567,89")', "1234567,89"],
  [1031, 'CDbl("12.34")', "1234"],
  [1031, 'CInt("1,5")', "2"],
  [1031, 'CInt("1.5")', "15"],
  [1031, 'CCur("1,5")', "1,5"],
  [1031, 'CSng("1.5")', "15"],
  [1031, '"1,5" + 0', "1,5"],
  [1031, '"1" + 1', "2"],
  [1031, '"1" + "1"', "11"],
  [1031, '"a" + 1', "ERR 13"],
  [1031, '"1" + True', "0"],
  [1031, 'TypeName("1" + 1)', "Double"],
  [1031, 'CStr(IsNumeric("1.5"))', "True"],
  [1031, 'CStr(IsNumeric("1,5"))', "True"],
  [1031, 'CStr(IsNumeric("abc"))', "False"],
  [1031, 'CStr(IsNumeric("1 234"))', "False"],
  [1033, 'CDbl("1234.567")', "1234.567"],
  [1033, 'CDbl("1,5")', "15"],
  [1033, 'CDbl("1.5")', "1.5"],
  [1033, 'CDbl("1.234,5")', "1.2345"],
  [1033, 'CDbl("1,234.5")', "1234.5"],
  [1033, 'CDbl("$1.5")', "1.5"],
  [1033, 'CDbl("1,5 €")', "ERR 13"],
  [1033, 'CDbl(",5")', "ERR 13"],
  [1033, 'CDbl("1,5,5")', "155"],
  [1033, 'CDbl("1,5.")', "15"],
  [1033, 'CDbl("12.34")', "12.34"],
  [1033, 'CDbl("1.234.567,89")', "ERR 13"],
  [1033, 'CInt("1,5")', "15"],
  [1033, 'CInt("1.5")', "2"],
  [1033, 'CStr(IsNumeric("1,5"))', "True"],
  // Windows reads four marks and not two: LOCALE_SDECIMAL and LOCALE_SMONDECIMALSEP both stand for the
  // fraction, and LOCALE_STHOUSAND and LOCALE_SMONTHOUSANDSEP both group. 3079 (de-AT) groups a plain
  // number with its money mark, and 1102 (mr-IN) takes its money fraction mark, a backtick.
  [3079, 'CDbl("1.5")', "15"],
  [3079, 'CDbl("1234.567")', "1234567"],
  [3079, 'CDbl("1 234.5")', "12345"],
  [1102, 'CDbl("1`234")', "1.234"],
  [4108, 'CDbl("1,234")', "1,234"],
  [4108, 'CDbl("1.234")', "1,234"],
  // A space groups a number wherever the locale groups with a space of its own, and nowhere else.
  [1026, 'CDbl("1 234")', "1234"],
  [2077, 'CDbl("1 234")', "1234"],
  [3079, 'CDbl("1 234")', "1234"],
  [1031, 'CDbl("1" & Chr(9) & "234")', "ERR 13"],
  [1033, 'CDbl("1 234")', "ERR 13"],
  // A full stop the locale does not name as a fraction mark is not a number mark either.
  [1026, 'CDbl("1.5")', "ERR 13"],
  [1036, 'CDbl("1234.567")', "ERR 13"],
  [1065, 'CDbl("1.5")', "ERR 13"],
  [1065, 'CDbl("1234.567")', "ERR 13"],
];

function answer(expression: string, lcid: number): unknown {
  const engine = new VbsEngine();
  try {
    engine.executeStatement(`SetLocale ${lcid}\nr = CStr(${expression})`);
    const failure = engine.error;
    if (failure) return "ERR " + String(failure.number);
    return engine.eval("r");
  } finally {
    engine.executeStatement("SetLocale 1033");
  }
}

describe("a number written as a string is read with the locale marks", () => {
  for (const [lcid, expression, expected] of CELLS) {
    it(`under ${lcid}, ${expression} is ${expected}`, () => {
      expect(answer(expression, lcid)).toBe(expected);
    });
  }
});
