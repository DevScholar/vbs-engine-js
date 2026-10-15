import { describe, it, expect } from 'vitest';
import { VbsEngine } from '../index.ts';

// Hex and Oct, and the strings CDate reads. Every cell was read off a fresh cscript on Windows build 26200.
//
// Hex and Oct write a negative value as its two's complement, and the width follows the type as well as the
// size:
//
//   * An Integer, a Boolean or a Byte fills sixteen bits while its magnitude fits a signed sixteen bit value,
//     so Hex(-1) is "FFFF" and Oct(-1) is "177777".
//   * Everything else fills thirty two, so Hex(CLng(-1)), Hex(CDbl(-1)) and Hex(CCur(-1)) are all "FFFFFFFF".
//   * A value of -32768 is written in thirty two bits even as an Integer, because its magnitude no longer
//     fits sixteen: Hex(-32768) is "FFFF8000".
//   * The value is rounded to a whole number first, a tie going to the even neighbour: Hex(0.5) is "0",
//     Hex(1.5) is "2", Hex(2.5) is "2". A value past the Long range is Overflow.
const DIGIT_CELLS: ReadonlyArray<readonly [string, string]> = [
  ['Hex(255)', 'FF'],
  ['Hex(0)', '0'],
  ['Hex(Empty)', '0'],
  ['Hex("255")', 'FF'],
  ['Hex(-1)', 'FFFF'],
  ['Hex(True)', 'FFFF'],
  ['Hex(CLng(-1))', 'FFFFFFFF'],
  ['Hex(CDbl(-1))', 'FFFFFFFF'],
  ['Hex(CSng(-1))', 'FFFFFFFF'],
  ['Hex(CCur(-1))', 'FFFFFFFF'],
  ['Hex("-1")', 'FFFFFFFF'],
  ['Hex(-32767)', '8001'],
  ['Hex(-32768)', 'FFFF8000'],
  ['Hex(-32769)', 'FFFF7FFF'],
  ['Hex(-2147483648)', '80000000'],
  ['Hex(2147483647)', '7FFFFFFF'],
  ['Hex(0.5)', '0'],
  ['Hex(1.5)', '2'],
  ['Hex(2.5)', '2'],
  ['Hex(-0.5)', '0'],
  ['Hex(-1.5)', 'FFFFFFFE'],
  ['Hex(CSng(255.6))', '100'],
  ['Hex(CCur(1.5))', '2'],
  ['Hex(4294967295)', 'ERR 6'],
  ['Hex(4294967296)', 'ERR 6'],
  ['Hex(CCur(2147483648))', 'ERR 6'],
  ['Oct(8)', '10'],
  ['Oct(CByte(255))', '377'],
  ['Oct(-1)', '177777'],
  ['Oct(-32767)', '100001'],
  ['Oct(-32768)', '37777700000'],
  ['Oct(CLng(-1))', '37777777777'],
  ['Oct(CDbl(-1))', '37777777777'],
  ['Oct(CCur(-1))', '37777777777'],
  ['Oct(2147483647)', '17777777777'],
  ['Oct(2.5)', '2'],
];

function answer(expression: string, lcid = 1031): unknown {
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

describe('Hex and Oct write the width the value type asks for', () => {
  for (const [expression, expected] of DIGIT_CELLS) {
    it(`${expression} is ${expected}`, () => {
      expect(answer(expression)).toBe(expected);
    });
  }
});

// The strings CDate reads, which are a wider set than IsDate and DateValue read:
//
//   * A plain number is a day count: CDate("2026") is 18 July 1905. Only CDate reads it, since IsDate("2026")
//     is False and DateValue("2026") is a type mismatch. It may carry a sign, spaces, leading zeros or a VB
//     literal prefix, and it has to land inside the range a date holds, which is 1 January 100 to
//     31 December 9999.
//   * A year under a hundred is read as 2000 to 2049 and 1950 to 1999, by VALUE and not by the digits
//     written, so CDate("0001-01-01") is 1 January 2001 while CDate("1/1/100") is year 100.
//   * A year is written with the digits it has: CDate("1/1/100") prints "01.01.100" and not "01.01.0100".
//   * A time may be marked with dots where a dot is not the locale's date separator: under 1033 CDate("2.5")
//     is 2:05 in the morning, while under 1031 a dot is the date separator and the same text is 2 May.
const DATE_STRING_CELLS: ReadonlyArray<readonly [number, string, string]> = [
  [1031, 'CStr(CDate("1/5/49"))', '01.05.2049'],
  [1031, 'CStr(CDate("1/5/50"))', '01.05.1950'],
  [1033, 'CStr(CDate("1/5/99"))', '1/5/1999'],
  [1033, 'CStr(CDate("1/5/50"))', '1/5/1950'],
  [1031, 'CStr(CDate("1/1/100"))', '01.01.100'],
  [1033, 'CStr(CDate("1/1/999"))', '1/1/999'],
  [1031, 'CStr(CDate("0999-01-05"))', '05.01.999'],
  [1031, 'CStr(CDate("0001-01-01"))', '01.01.2001'],
  [1031, 'CStr(CDate("2026"))', '18.07.1905'],
  [1033, 'CStr(CDate("2026"))', '7/18/1905'],
  [1031, 'CStr(CDate("0"))', '00:00:00'],
  [1031, 'CStr(CDate("-1"))', '29.12.1899'],
  [1031, 'CStr(CDate("&H10"))', '15.01.1900'],
  [1031, 'CStr(CDate(" +7"))', '06.01.1900'],
  [1031, 'CStr(CDate("  7  "))', '06.01.1900'],
  [1031, 'CStr(CDate("2958465"))', '31.12.9999'],
  [1031, 'CDate("2958466")', 'ERR 13'],
  [1031, 'CStr(CDate("-657434"))', '01.01.100'],
  [1031, 'CDate("-657435")', 'ERR 13'],
  [1031, 'CDate("99999999999")', 'ERR 13'],
  [1031, 'CStr(IsDate("2026"))', 'False'],
  [1031, 'DateValue("2026")', 'ERR 13'],
  [1031, 'TimeValue("2026")', 'ERR 13'],
  [1033, 'CStr(CDate("2.5"))', '2:05:00 AM'],
  [1031, 'CStr(CDate("2.5"))', '02.05.2026'],
  [1033, 'CStr(CDate("2.5 PM"))', '2:05:00 PM'],
  [1033, 'CStr(IsDate("2.5"))', 'True'],
  [1031, 'CStr(IsDate("2.5"))', 'True'],
  [1033, 'CStr(CDate("13.45"))', '1:45:00 PM'],
  // A two field date, and a comma as a field mark.
  [1031, 'CStr(CDate("12.31"))', '31.12.2026'],
  [1031, 'CStr(CDate("2.30"))', '01.02.2030'],
  [1031, 'CStr(CDate("0.5"))', '01.05.2000'],
  [1031, 'CStr(CDate("100.5"))', '01.05.100'],
  [1031, 'CStr(CDate("31.4"))', '01.04.2031'],
  [1031, 'CStr(CDate("29.2"))', '01.02.2029'],
  [1033, 'CStr(CDate("2,5"))', '2/5/2026'],
  [1033, 'CStr(CDate("5,1"))', '5/1/2026'],
  [1033, 'CStr(CDate("1,234"))', '1/1/234'],
  [1031, 'CStr(CDate("1,234"))', '01.01.234'],
  // A number written with the marks a locale uses between digits, which only CDate reads.
  [1031, 'CStr(CDate("13.45"))', '06.09.1903'],
  [1031, 'CStr(CDate("1.2.3.4"))', '18.05.1903'],
  [1031, 'CStr(CDate("123.456"))', '03.01.2238'],
  [1031, 'CStr(CDate("-1.5"))', '15.12.1899'],
  [1031, 'CStr(CDate("00.00"))', '00:00:00'],
  [1031, 'CStr(CDate("13,45"))', '12.01.1900 10:48:00'],
  [1033, 'CStr(CDate("123.456"))', '5/2/1900 10:56:38 AM'],
  [1033, 'CStr(CDate("-1.5"))', '12/29/1899 12:00:00 PM'],
  [1033, 'CStr(CDate("60.5"))', '2/28/1900 12:00:00 PM'],
  [1033, 'CStr(CDate("32.5"))', '1/31/1900 12:00:00 PM'],
  [1033, 'CStr(CDate("13,45"))', '9/6/1903'],
  [1033, 'CStr(CDate("2958465.5"))', '12/31/9999 12:00:00 PM'],
  [1033, 'CStr(CDate("1.2.3.4"))', 'ERR 13'],
  [1031, 'CStr(IsDate("13.45"))', 'False'],
  [1033, 'CStr(IsDate("2.60"))', 'False'],
  [1033, 'CStr(IsDate("123.456"))', 'False'],
  [1033, 'CStr(IsDate("60.5"))', 'False'],
];

// A day count below zero, which counts its time of day away from zero instead of toward it. Measured the same
// way: the day alone is -1 for 29 December 1899, and adding noon to it gives -1.5 rather than -0.5, while
// 30 December 1899 at noon is the day 0 at noon and reads as 0.5.
const NEGATIVE_SERIAL_CELLS: ReadonlyArray<readonly [string, string]> = [
  ['CStr(CDbl(CDate("12/29/1899")))', '-1'],
  ['CStr(CDbl(CDate("12/29/1899 12:00:00")))', '-1.5'],
  ['CStr(CDbl(CDate("12/28/1899 12:00:00")))', '-2.5'],
  ['CStr(CDbl(CDate("12/29/1899 06:00:00")))', '-1.25'],
  ['CStr(CDbl(CDate("12/30/1899 12:00:00")))', '0.5'],
  ['CStr(CDbl(CDate("12/30/1899 06:00:00")))', '0.25'],
  ['CStr(CDbl(CDate("12/31/1899 12:00:00")))', '1.5'],
  ['CStr(CDate(-1.5))', '12/29/1899 12:00:00 PM'],
  ['CStr(CDate(-0.5))', '12:00:00 PM'],
  ['CStr(CDate(-2.5))', '12/28/1899 12:00:00 PM'],
  ['CStr(CDate(-0.25))', '6:00:00 AM'],
  ['CStr(CDate("12/29/1899") + 0.5)', '12:00:00 PM'],
  ['CStr(DateDiff("d", CDate("12/29/1899"), CDate("12/31/1899")))', '2'],
];

describe('the strings CDate reads are the ones vbscript.dll reads', () => {
  for (const [lcid, expression, expected] of DATE_STRING_CELLS) {
    it(`under ${lcid}, ${expression} is ${expected}`, () => {
      expect(answer(expression, lcid)).toBe(expected);
    });
  }
});

// A date field with more digits than a date can hold, which the original reader keeps in sixteen signed bits.
// Measured under 1031, where a dot is the date separator: a year wraps, so 2958465 becomes 9345 and 65535
// becomes -1, which the year window then reads as 1999. A day or a month does not wrap, so a long one is
// simply not a day or a month and the text falls through to the other readings.
const WRAPPED_FIELD_CELLS: ReadonlyArray<readonly [number, string, string]> = [
  [1031, 'CStr(CDate("2958465.5"))', '01.05.9345'],
  [1031, 'CStr(CDate("2958465.4"))', '01.04.9345'],
  [1031, 'CStr(CDate("2958466.5"))', '01.05.9346'],
  [1031, 'CStr(CDate("5.2958465"))', '01.05.9345'],
  [1031, 'CStr(CDate("65535.5"))', '01.05.1999'],
  [1031, 'CStr(CDate("65536.5"))', '01.05.2000'],
  [1031, 'CStr(CDate("65537.5"))', '01.05.2001'],
  [1031, 'CStr(CDate("131073.5"))', '01.05.2001'],
  [1031, 'CStr(CDate("196609.5"))', '01.05.2001'],
  [1031, 'CStr(CDate("70000.5"))', '01.05.4464'],
  [1031, 'CStr(CDate("32767.5"))', '20.02.2797'],
  [1031, 'CStr(CDate("100000.5"))', '01.12.4637'],
  [1031, 'CStr(CDate("2958465.25"))', 'ERR 13'],
  [1031, 'CStr(CDate("999999.5"))', 'ERR 13'],
  [1031, 'CStr(CDate("2000000.5"))', 'ERR 13'],
  [1033, 'CStr(CDate("1000000.5"))', '11/26/4637 12:00:00 PM'],
  [1033, 'CStr(CDate("2000000.5"))', '10/23/7375 12:00:00 PM'],
  [1033, 'CStr(CDate("2958465.5"))', '12/31/9999 12:00:00 PM'],
  [1033, 'CStr(CDate("65536.5"))', '12:05:00 AM'],
  [1033, 'CStr(CDate("131073.5"))', '1:05:00 AM'],
  [1033, 'CStr(CDate("1.2.65536"))', '1:02:00 AM'],
  [1033, 'CStr(CDate("65535.5"))', '6/4/2079 12:00:00 PM'],
  [1033, 'CStr(CDate("70000.5"))', '8/25/2091 12:00:00 PM'],
  [1033, 'CStr(CDate("1.2.70000"))', 'ERR 13'],
  [1033, 'CStr(CDate("1.70000.5"))', 'ERR 13'],
  // A three field date whose year is too large for a year is read as the whole text with the marks taken
  // out, which is a day count: 1.2.65535 is the day 1265535. A year that fits keeps its own reading
  // (1.2.9999 is 1 February 9999), and a day count past what a date holds is a type mismatch.
  [1031, 'CStr(CDate("1.2.65535"))', '29.11.5364'],
  [1031, 'CStr(CDbl(CDate("1.2.65535")))', '1265535'],
  [1031, 'CStr(CDate("1.2.65536"))', '30.11.5364'],
  [1031, 'CStr(CDate("1.2.70000"))', '19.02.5377'],
  [1031, 'CStr(CDbl(CDate("1.2.70000")))', '1270000'],
  [1031, 'CStr(CDate("1.70000.2"))', '11.06.6554'],
  [1031, 'CStr(CDate("1.2.10000"))', '11.11.5212'],
  [1031, 'CStr(CDate("1.2.9999"))', '01.02.9999'],
  [1031, 'CStr(CDate("5.6.65535"))', 'ERR 13'],
  [1031, 'CStr(CDate("1.2.100000"))', 'ERR 13'],
  [1033, 'CStr(CDate("1.2.70000"))', 'ERR 13'],
];

describe('a date field is held in sixteen signed bits', () => {
  for (const [lcid, expression, expected] of WRAPPED_FIELD_CELLS) {
    it(`under ${lcid}, ${expression} is ${expected}`, () => {
      expect(answer(expression, lcid)).toBe(expected);
    });
  }
});

describe('a day count below zero counts its time away from zero', () => {
  for (const [expression, expected] of NEGATIVE_SERIAL_CELLS) {
    it(`${expression} is ${expected}`, () => {
      expect(answer(expression, 1033)).toBe(expected);
    });
  }
});

// The month names Windows writes and reads, and the calendar Thai counts its years in.
//
//   * Windows writes its own name where CLDR has changed one, and reads it back: under 1031 MonthName(3, True)
//     is "Mrz" while CLDR's short name is "Mär", and both "Mrz" and "Mär" are read as March. The engine keeps
//     the CLDR names and adds Windows' own from its name table.
//   * Thai reads a year as a Buddhist one, so CDate("5 มกราคม 2026") is the Gregorian year 1483 and CStr
//     writes it back as "5/1/2026", while CDate("5 มกราคม 2569") is the Gregorian 2026 and writes "5/1/2569".
//     The Buddhist year is the Gregorian one plus 543, months and days unchanged, and ICU is not asked for
//     it: ICU counts the Buddhist calendar in the Julian calendar before 1582 and would write that same
//     Gregorian 5 January 1483 as 27 December 2025.
const LOCALE_NAME_CELLS: ReadonlyArray<readonly [number, string, string]> = [
  [1031, 'CStr(CDate("5. Mrz 2026"))', '05.03.2026'],
  [1031, 'CStr(CDate("5. Mär 2026"))', '05.03.2026'],
  [1031, 'CStr(CDate("5. März 2026"))', '05.03.2026'],
  [1031, 'CStr(CDate("5. Januar 2026"))', '05.01.2026'],
  [1031, 'CStr(MonthName(3))', 'März'],
  [1031, 'CStr(MonthName(3, True))', 'Mrz'],
  [1032, 'CStr(MonthName(1))', 'Ιανουάριος'],
  [1049, 'CStr(MonthName(1))', 'Январь'],
  [1054, 'CStr(CDate("5 มกราคม 2026"))', '5/1/2026'],
  [1054, 'CStr(CDate("5 มกราคม 2569"))', '5/1/2569'],
  [1054, 'CStr(CDbl(CDate("5/1/2026")))', '-152300'],
  [1054, 'FormatDateTime(DateSerial(2026, 1, 5), 1)', '5 มกราคม 2569'],
  [1054, 'CStr(CDate("5/1/2569"))', '5/1/2569'],
];

describe('the names Windows writes and the years Thai counts', () => {
  for (const [lcid, expression, expected] of LOCALE_NAME_CELLS) {
    it(`under ${lcid}, ${expression} is ${expected}`, () => {
      expect(answer(expression, lcid)).toBe(expected);
    });
  }
});

// A month name read under a locale: how much of the name counts, and which marks may stand around it.
//
//   * Three letters or more of the locale's own month names are a month, matched as a prefix, so "Jan",
//     "Janu", "Januar" and "January" are all January in English while "Ja" is none of them and "Januaryy" is
//     too long.
//   * Only the locale's own names count: German names read under 1031 and not under 1033, and English names
//     are read under 1033 and under the three CJK locales, but not the other way round.
//   * The marks around a month name are a space, a comma, a slash, a hyphen or the locale's own date mark. A
//     dot counts only where the locale writes its dates with dots, so under 1033 "5. Januar 2026" is a type
//     mismatch while "Januar. 5, 2026" reads (the dot there belongs to the name), and under 1031 both read.
const MONTH_WORD_CELLS: ReadonlyArray<readonly [number, string, string]> = [
  [1033, 'CStr(CDate("Januar 5, 2026"))', '1/5/2026'],
  [1033, 'CStr(CDate("Janu 5, 2026"))', '1/5/2026'],
  [1033, 'CStr(CDate("Januar. 5, 2026"))', '1/5/2026'],
  [1033, 'CStr(CDate("January 5, 2026"))', '1/5/2026'],
  [1033, 'CStr(CDate("Decem 5, 2026"))', '12/5/2026'],
  [1033, 'CDate("Ja 5, 2026")', 'ERR 13'],
  [1033, 'CDate("Januaryy 5, 2026")', 'ERR 13'],
  [1033, 'CDate("5. Januar 2026")', 'ERR 13'],
  [1033, 'CDate("5 janvier 2026")', 'ERR 13'],
  [1031, 'CStr(CDate("5. Januar 2026"))', '05.01.2026'],
  [1031, 'CStr(CDate("5 Januar 2026"))', '05.01.2026'],
  [1031, 'CStr(CDate("5/Januar/2026"))', '05.01.2026'],
  [1031, 'CStr(CDate("5-Januar-2026"))', '05.01.2026'],
  [1031, 'CStr(CDate("5, Januar, 2026"))', '05.01.2026'],
  [1031, 'CStr(CDate("Januar 5, 2026"))', '05.01.2026'],
  [1031, 'CStr(CDate("Janu 5, 2026"))', '05.01.2026'],
  [1031, 'CDate("Ja 5, 2026")', 'ERR 13'],
  [1031, 'CDate("January 5, 2026")', 'ERR 13'],
  [1031, 'CDate("Mar 5, 2026")', 'ERR 13'],
  [1031, 'CDate("5 janvier 2026")', 'ERR 13'],
  [1041, 'CStr(CDate("Januar 5, 2026"))', '2026/01/05'],
  [1041, 'CDate("5. Januar 2026")', 'ERR 13'],
  [2052, 'CStr(CDate("Decem 5, 2026"))', '2026/12/5'],
  [2052, 'CDate("5 janvier 2026")', 'ERR 13'],
  [1042, 'CStr(CDate("Januar 5, 2026"))', '2026-01-05'],
];

describe('a month name is read the way Windows reads one', () => {
  for (const [lcid, expression, expected] of MONTH_WORD_CELLS) {
    it(`under ${lcid}, ${expression} is ${expected}`, () => {
      expect(answer(expression, lcid)).toBe(expected);
    });
  }
});
