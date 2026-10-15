import { describe, it, expect } from 'vitest';
import { VbsEngine } from '../index.ts';

// The string functions, measured against a fresh cscript under 1033 and 1031 (the same answers in both).
//
// The corners are where the real engine differs from a plain JavaScript reading of the same idea:
//
//   * InStr takes either two arguments or a start first, and the two forms behave differently: with two, a
//     Null argument answers Null and an empty string to look for answers 1 (or 0 when the text is empty);
//     with a start, a Null is the invalid use of Null, a start below 1 is error 5, a start past the end is 0,
//     and a compare other than 0 or 1 is error 5.
//   * InStrRev searches back from its start, which defaults to -1 meaning the end of the text. Any other
//     negative start, and 0, are error 5, a start past the end answers 0, and an empty needle answers with
//     the start itself.
//   * Split makes no parts at all for a limit of 0 or an empty text, one part for an empty delimiter, and
//     error 5 for a limit below -1 or a compare other than 0 or 1.
//   * Replace with an empty string to look for leaves the text as it stands. Stepping by nothing is also how
//     it used to hang the engine.
//   * Trim, LTrim and RTrim take off spaces and nothing else: a tab, a line break and a non-breaking space
//     all stay.
//   * Mid counts its start from 1 and refuses one below that or a length below zero with error 5, and
//     Left, Right, String and Space refuse a negative count the same way.
//   * String takes a character code for a number and the first character for a string: String(3, 65) is
//     "AAA" while String(3, "ab") is "aaa".
//   * The functions that give text back (Len, Left, Right, Mid, UCase, LCase, LTrim, RTrim, Trim) answer
//     Null for a Null argument, and the ones that build from their arguments (Replace, Split, Join, Filter,
//     String, Space) raise the invalid use of Null instead.
const CELLS: ReadonlyArray<readonly [string, string]> = [
  ['InStr("abcabc", "b")', '2'],
  ['InStr("abc", "z")', '0'],
  ['InStr("abc", "")', '1'],
  ['InStr("", "")', '0'],
  ['InStr("", "a")', '0'],
  ['InStr(3, "abcabc", "b")', '5'],
  ['InStr(10, "abc", "a")', '0'],
  ['InStr(0, "abc", "a")', 'ERR 5'],
  ['InStr(-1, "abc", "a")', 'ERR 5'],
  ['InStr(1, "abc", "")', '1'],
  ['InStr(5, "abc", "")', '0'],
  ['InStr(1, "ABC", "b", 1)', '2'],
  ['InStr(1, "ABC", "b", 0)', '0'],
  ['InStr("ABC", "b")', '0'],
  ['InStr(1, "abc", "a", 2)', 'ERR 5'],
  ['InStr(12345, "3")', '3'],
  ['InStr("a1b", 1)', '2'],
  ['InStr(2, "abcabc", "a")', '4'],
  ['CStr(IsNull(InStr("abc", Null)))', 'True'],
  ['CStr(IsNull(InStr(Null, "a")))', 'True'],
  ['InStr(Null, "abc", "a")', 'ERR 94'],
  ['InStrRev("abcabc", "b")', '5'],
  ['InStrRev("abc", "z")', '0'],
  ['InStrRev("abcabc", "b", 3)', '2'],
  ['InStrRev("abcabc", "b", -1)', '5'],
  ['InStrRev("abcabc", "b", 0)', 'ERR 5'],
  ['InStrRev("abcabc", "b", -2)', 'ERR 5'],
  ['InStrRev("abcabc", "b", 7)', '0'],
  ['InStrRev("abcabc", "a", 1)', '1'],
  ['InStrRev("ABCabc", "b", -1, 1)', '5'],
  ['InStrRev("abc", "")', '3'],
  ['InStrRev("", "")', '0'],
  ['InStrRev("abc", "a", -1, 2)', 'ERR 5'],
  ['InStrRev("abc", Null)', 'ERR 94'],
  ['Join(Split("a,b,c", ","), "|")', 'a|b|c'],
  ['UBound(Split("a,b,c", ","))', '2'],
  ['LBound(Split("a,b,c", ","))', '0'],
  ['Join(Split("a,b,c", ",", 2), "|")', 'a|b,c'],
  ['Join(Split("a,b,c", ",", 1), "|")', 'a,b,c'],
  ['Join(Split("a,b,c", ",", -1), "|")', 'a|b|c'],
  ['Join(Split("a,b,c", ",", 0), "|")', ''],
  ['Join(Split("a b c"), "|")', 'a|b|c'],
  ['UBound(Split("", ","))', '-1'],
  ['Join(Split("", ","), "|")', ''],
  ['Join(Split("a,,b", ","), "|")', 'a||b'],
  ['Join(Split("aXbXc", "x", -1, 1), "|")', 'a|b|c'],
  ['Join(Split("a--b--c", "--"), "|")', 'a|b|c'],
  ['UBound(Split("abc", ","))', '0'],
  ['Join(Split("a,b,", ","), "|")', 'a|b|'],
  ['UBound(Split("abc", ""))', '0'],
  ['TypeName(Split("a,b", ","))', 'Variant()'],
  ['Split(Null)', 'ERR 94'],
  ['UBound(Split(Empty))', '-1'],
  ['Join(Split("a,b,c", ",", -2), "|")', 'ERR 5'],
  ['Join(Split("aXb", "x", -1, 2), "|")', 'ERR 5'],
  ['UBound(Split(",", ","))', '1'],
  ['UBound(Split("a", "a"))', '1'],
  ['Join(Split("abc", Null), "|")', 'ERR 94'],
  ['Replace("abcabc", "b", "X")', 'aXcaXc'],
  ['Replace("abc", "", "X")', 'abc'],
  ['Replace("abcabc", "b", "X", 2)', 'XcaXc'],
  ['Replace("abcabc", "b", "X", 1, 1)', 'aXcabc'],
  ['Replace("abcabc", "b", "X", 1, 0)', 'abcabc'],
  ['Replace("ABCabc", "b", "X", 1, -1, 1)', 'AXCaXc'],
  ['Replace("", "a", "b")', ''],
  ['Replace("abc", "b", "")', 'ac'],
  ['Replace(Null, "a", "b")', 'ERR 94'],
  ['"[" & Trim("  a  ") & "]"', '[a]'],
  ['Len(Trim(Chr(9) & "a" & Chr(9)))', '3'],
  ['Len(Trim(Chr(13) & Chr(10) & "a" & Chr(13) & Chr(10)))', '5'],
  ['Len(Trim(Chr(160) & "a" & Chr(160)))', '3'],
  ['"[" & Trim("   ") & "]"', '[]'],
  ['"[" & LTrim("  a  ") & "]"', '[a  ]'],
  ['"[" & RTrim("  a  ") & "]"', '[  a]'],
  ['Trim(123)', '123'],
  ['"[" & Trim(Empty) & "]"', '[]'],
  ['"[" & Trim("a  b") & "]"', '[a  b]'],
  ['CStr(IsNull(Trim(Null)))', 'True'],
  ['Mid("abc", 0)', 'ERR 5'],
  ['Mid("abc", -1)', 'ERR 5'],
  ['Mid("abc", 2, -1)', 'ERR 5'],
  ['"[" & Mid("abc", 5) & "]"', '[]'],
  ['"[" & Mid("abc", 2, 0) & "]"', '[]'],
  ['Mid("abc", 2, 99)', 'bc'],
  ['Left("abc", -1)', 'ERR 5'],
  ['Right("abc", -1)', 'ERR 5'],
  ['Left("abc", 99)', 'abc'],
  ['"[" & Left("abc", 0) & "]"', '[]'],
  ['"[" & Right("abc", 0) & "]"', '[]'],
  ['String(3, "ab")', 'aaa'],
  ['String(3, 65)', 'AAA'],
  ['"[" & String(0, "a") & "]"', '[]'],
  ['String(-1, "a")', 'ERR 5'],
  ['Space(-1)', 'ERR 5'],
  ['String(2, Null)', 'ERR 94'],
  ['Join(Null, "-")', 'ERR 94'],
  ['Filter(Null, "a")', 'ERR 94'],
  ['CStr(IsNull(Len(Null)))', 'True'],
  ['CStr(IsNull(UCase(Null)))', 'True'],
  ['CStr(IsNull(LCase(Null)))', 'True'],
  ['CStr(IsNull(Left(Null, 1)))', 'True'],
  ['CStr(IsNull(Mid(Null, 1)))', 'True'],
  ['Len(Empty)', '0'],
  ['Len(12345)', '5'],
  ['Len(Space(3))', '3'],
  ['StrReverse("abc")', 'cba'],
  ['Join(Array("a", "b"), "-")', 'a-b'],
  ['UCase("aBc")', 'ABC'],
  ['LCase("aBc")', 'abc'],
];

function answer(expression: string): unknown {
  const engine = new VbsEngine();
  try {
    engine.executeStatement(`SetLocale 1033\nr = CStr(${expression})`);
    const failure = engine.error;
    if (failure) return 'ERR ' + String(failure.number);
    return engine.eval('r');
  } finally {
    engine.executeStatement('SetLocale 1033');
  }
}

describe('the string functions behave like vbscript.dll', () => {
  for (const [expression, expected] of CELLS) {
    it(`${expression} is ${expected}`, () => {
      expect(answer(expression)).toBe(expected);
    });
  }
});
