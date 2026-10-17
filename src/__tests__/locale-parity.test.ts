import { describe, it, expect } from "vitest";
import { VbsEngine } from "../index.ts";

// Cell by cell against what vbscript.dll answers.
//
// Every expected value below was read off a cscript run under one process per locale, on Windows build
// 26200, and written through FileSystemObject.CreateTextFile so the console code page could not lose a
// character. The eight tables the runtime ships were read from the same kind of run.
//
// The first block is the German family and Tsonga, where the work started. The second is what the full
// sweep over all 132 locales turned up afterwards, and each one is a rule rather than a character:
//
//   * th-TH writes its year in the Buddhist calendar in EVERY style, CStr included ("5/1/2569"), and its
//     long date carries no weekday although LOCALE_SLONGDATE asks for one.
//   * ar-SA writes its long date with the short date's separator ("16/رجب/1447", not "16 رجب, 1447") and
//     names the sixth Hijri month "جمادى الثانية" where CLDR says "جمادى الآخرة".
//   * fa-IR pads the day in its short date ("04/09/1405") and not in its long date ("7 اسفند 1405").
//   * hi-IN and mr-IN group by three although LOCALE_SGROUPING asks for the lakh grouping: FormatNumber
//     gives "1,234,567.89", not "12,34,567.89".
//   * en-US writes a negative amount in parentheses as "($1,234.57)", with no sign inside.
//   * hr-HR, hi-IN, fa-IR and ar-DZ each spell one name inside a date differently from CLDR.

function answer(lcid: number, expression: string): unknown {
  const engine = new VbsEngine();
  try {
    engine.executeStatement(`SetLocale ${lcid}\nr = (${expression})`);
    expect(engine.error).toBeNull();
    return engine.eval("r");
  } finally {
    engine.executeStatement("SetLocale 1033");
  }
}

const GERMAN_AND_TSONGA: ReadonlyArray<readonly [number, string, string]> = [
  // de-DE (1031)
  [1031, "CStr(DateSerial(2026,1,5) + TimeSerial(0,0,0))", "05.01.2026"],
  [
    1031,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 0)",
    "05.01.2026 09:05:00",
  ],
  [
    1031,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 1)",
    "Montag, 5. Januar 2026",
  ],
  [
    1031,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 2)",
    "05.01.2026",
  ],
  [
    1031,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 3)",
    "09:05:00",
  ],
  [
    1031,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 4)",
    "09:05",
  ],
  [
    1031,
    "FormatDateTime(DateSerial(2026,11,25) + TimeSerial(22,30,45), 1)",
    "Mittwoch, 25. November 2026",
  ],
  [
    1031,
    "FormatDateTime(DateSerial(2026,11,25) + TimeSerial(22,30,45), 2)",
    "25.11.2026",
  ],
  [
    1031,
    "FormatDateTime(DateSerial(2026,12,31) + TimeSerial(23,59,59), 1)",
    "Donnerstag, 31. Dezember 2026",
  ],
  [
    1031,
    "FormatDateTime(DateSerial(2026,12,31) + TimeSerial(23,59,59), 2)",
    "31.12.2026",
  ],
  [1031, "MonthName(3)", "März"],
  [1031, "MonthName(3, True)", "Mrz"],
  [1031, "MonthName(12, True)", "Dez"],
  [1031, "WeekdayName(2)", "Montag"],
  [1031, "WeekdayName(2, True)", "Mo"],
  [1031, "WeekdayName(7, True)", "Sa"],
  [1031, "CStr(1234567.891)", "1234567,891"],
  [1031, "FormatNumber(1234567.891)", "1.234.567,89"],
  [1031, "FormatCurrency(1234567.891)", "1.234.567,89 €"],
  [1031, "FormatPercent(1234567.891)", "123.456.789,10%"],
  [1031, "FormatNumber(-1234.567)", "-1.234,57"],
  [1031, "FormatCurrency(-1234.567)", "-1.234,57 €"],
  [1031, "FormatPercent(0.125)", "12,50%"],
  [1031, "FormatCurrency(1235)", "1.235,00 €"],
  [1031, "FormatNumber(1234567.891, 2, -1, -1, 0)", "1234567,89"],
  [1031, "FormatNumber(-1234.567, 2, -1, True, -1)", "(1.234,57)"],
  [1031, "FormatNumber(1234567.891, 0)", "1.234.568"],
  [1031, "FormatPercent(0.125, 0)", "13%"],
  // de-CH (2055)
  [2055, "CStr(DateSerial(2026,1,5) + TimeSerial(0,0,0))", "05.01.2026"],
  [
    2055,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 0)",
    "05.01.2026 09:05:00",
  ],
  [
    2055,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 1)",
    "Montag, 5. Januar 2026",
  ],
  [
    2055,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 2)",
    "05.01.2026",
  ],
  [
    2055,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 3)",
    "09:05:00",
  ],
  [
    2055,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 4)",
    "09:05",
  ],
  [
    2055,
    "FormatDateTime(DateSerial(2026,11,25) + TimeSerial(22,30,45), 1)",
    "Mittwoch, 25. November 2026",
  ],
  [
    2055,
    "FormatDateTime(DateSerial(2026,11,25) + TimeSerial(22,30,45), 2)",
    "25.11.2026",
  ],
  [
    2055,
    "FormatDateTime(DateSerial(2026,12,31) + TimeSerial(23,59,59), 1)",
    "Donnerstag, 31. Dezember 2026",
  ],
  [
    2055,
    "FormatDateTime(DateSerial(2026,12,31) + TimeSerial(23,59,59), 2)",
    "31.12.2026",
  ],
  [2055, "MonthName(3)", "März"],
  [2055, "MonthName(3, True)", "Mär"],
  [2055, "MonthName(12, True)", "Dez"],
  [2055, "WeekdayName(2)", "Montag"],
  [2055, "WeekdayName(2, True)", "Mo"],
  [2055, "WeekdayName(7, True)", "Sa"],
  [2055, "CStr(1234567.891)", "1234567.891"],
  [2055, "FormatNumber(1234567.891)", "1’234’567.89"],
  [2055, "FormatCurrency(1234567.891)", "CHF 1’234’567.89"],
  [2055, "FormatPercent(1234567.891)", "123’456’789.10%"],
  [2055, "FormatNumber(-1234.567)", "-1’234.57"],
  [2055, "FormatCurrency(-1234.567)", "CHF-1’234.57"],
  [2055, "FormatPercent(0.125)", "12.50%"],
  [2055, "FormatCurrency(1235)", "CHF 1’235.00"],
  [2055, "FormatNumber(1234567.891, 2, -1, -1, 0)", "1234567.89"],
  [2055, "FormatNumber(-1234.567, 2, -1, True, -1)", "(1’234.57)"],
  [2055, "FormatNumber(1234567.891, 0)", "1’234’568"],
  [2055, "FormatPercent(0.125, 0)", "13%"],
  // de-AT (3079)
  [3079, "CStr(DateSerial(2026,1,5) + TimeSerial(0,0,0))", "05.01.2026"],
  [
    3079,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 0)",
    "05.01.2026 09:05:00",
  ],
  [
    3079,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 1)",
    "Montag, 5. Jänner 2026",
  ],
  [
    3079,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 2)",
    "05.01.2026",
  ],
  [
    3079,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 3)",
    "09:05:00",
  ],
  [
    3079,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 4)",
    "09:05",
  ],
  [
    3079,
    "FormatDateTime(DateSerial(2026,11,25) + TimeSerial(22,30,45), 1)",
    "Mittwoch, 25. November 2026",
  ],
  [
    3079,
    "FormatDateTime(DateSerial(2026,11,25) + TimeSerial(22,30,45), 2)",
    "25.11.2026",
  ],
  [
    3079,
    "FormatDateTime(DateSerial(2026,12,31) + TimeSerial(23,59,59), 1)",
    "Donnerstag, 31. Dezember 2026",
  ],
  [
    3079,
    "FormatDateTime(DateSerial(2026,12,31) + TimeSerial(23,59,59), 2)",
    "31.12.2026",
  ],
  [3079, "MonthName(3)", "März"],
  [3079, "MonthName(3, True)", "Mär"],
  [3079, "MonthName(12, True)", "Dez"],
  [3079, "WeekdayName(2)", "Montag"],
  [3079, "WeekdayName(2, True)", "Mo"],
  [3079, "WeekdayName(7, True)", "Sa"],
  [3079, "CStr(1234567.891)", "1234567,891"],
  [3079, "FormatNumber(1234567.891)", "1\u{00a0}234\u{00a0}567,89"],
  [3079, "FormatCurrency(1234567.891)", "€ 1.234.567,89"],
  [3079, "FormatPercent(1234567.891)", "123\u{00a0}456\u{00a0}789,10%"],
  [3079, "FormatNumber(-1234.567)", "-1\u{00a0}234,57"],
  [3079, "FormatCurrency(-1234.567)", "-€ 1.234,57"],
  [3079, "FormatPercent(0.125)", "12,50%"],
  [3079, "FormatCurrency(1235)", "€ 1.235,00"],
  [3079, "FormatNumber(1234567.891, 2, -1, -1, 0)", "1234567,89"],
  [3079, "FormatNumber(-1234.567, 2, -1, True, -1)", "(1\u{00a0}234,57)"],
  [3079, "FormatNumber(1234567.891, 0)", "1\u{00a0}234\u{00a0}568"],
  [3079, "FormatPercent(0.125, 0)", "13%"],
  // de-LU (4103)
  [4103, "CStr(DateSerial(2026,1,5) + TimeSerial(0,0,0))", "05.01.2026"],
  [
    4103,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 0)",
    "05.01.2026 09:05:00",
  ],
  [
    4103,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 1)",
    "Montag, 5. Januar 2026",
  ],
  [
    4103,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 2)",
    "05.01.2026",
  ],
  [
    4103,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 3)",
    "09:05:00",
  ],
  [
    4103,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 4)",
    "09:05",
  ],
  [
    4103,
    "FormatDateTime(DateSerial(2026,11,25) + TimeSerial(22,30,45), 1)",
    "Mittwoch, 25. November 2026",
  ],
  [
    4103,
    "FormatDateTime(DateSerial(2026,11,25) + TimeSerial(22,30,45), 2)",
    "25.11.2026",
  ],
  [
    4103,
    "FormatDateTime(DateSerial(2026,12,31) + TimeSerial(23,59,59), 1)",
    "Donnerstag, 31. Dezember 2026",
  ],
  [
    4103,
    "FormatDateTime(DateSerial(2026,12,31) + TimeSerial(23,59,59), 2)",
    "31.12.2026",
  ],
  [4103, "MonthName(3)", "März"],
  [4103, "MonthName(3, True)", "Mär"],
  [4103, "MonthName(12, True)", "Dez"],
  [4103, "WeekdayName(2)", "Montag"],
  [4103, "WeekdayName(2, True)", "Mo"],
  [4103, "WeekdayName(7, True)", "Sa"],
  [4103, "CStr(1234567.891)", "1234567,891"],
  [4103, "FormatNumber(1234567.891)", "1.234.567,89"],
  [4103, "FormatCurrency(1234567.891)", "1.234.567,89 €"],
  [4103, "FormatPercent(1234567.891)", "123.456.789,10%"],
  [4103, "FormatNumber(-1234.567)", "-1.234,57"],
  [4103, "FormatCurrency(-1234.567)", "-1.234,57 €"],
  [4103, "FormatPercent(0.125)", "12,50%"],
  [4103, "FormatCurrency(1235)", "1.235,00 €"],
  [4103, "FormatNumber(1234567.891, 2, -1, -1, 0)", "1234567,89"],
  [4103, "FormatNumber(-1234.567, 2, -1, True, -1)", "(1.234,57)"],
  [4103, "FormatNumber(1234567.891, 0)", "1.234.568"],
  [4103, "FormatPercent(0.125, 0)", "13%"],
  // de-LI (5127)
  [5127, "CStr(DateSerial(2026,1,5) + TimeSerial(0,0,0))", "05.01.2026"],
  [
    5127,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 0)",
    "05.01.2026 09:05:00",
  ],
  [
    5127,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 1)",
    "Montag, 5. Januar 2026",
  ],
  [
    5127,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 2)",
    "05.01.2026",
  ],
  [
    5127,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 3)",
    "09:05:00",
  ],
  [
    5127,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 4)",
    "09:05",
  ],
  [
    5127,
    "FormatDateTime(DateSerial(2026,11,25) + TimeSerial(22,30,45), 1)",
    "Mittwoch, 25. November 2026",
  ],
  [
    5127,
    "FormatDateTime(DateSerial(2026,11,25) + TimeSerial(22,30,45), 2)",
    "25.11.2026",
  ],
  [
    5127,
    "FormatDateTime(DateSerial(2026,12,31) + TimeSerial(23,59,59), 1)",
    "Donnerstag, 31. Dezember 2026",
  ],
  [
    5127,
    "FormatDateTime(DateSerial(2026,12,31) + TimeSerial(23,59,59), 2)",
    "31.12.2026",
  ],
  [5127, "MonthName(3)", "März"],
  [5127, "MonthName(3, True)", "Mär"],
  [5127, "MonthName(12, True)", "Dez"],
  [5127, "WeekdayName(2)", "Montag"],
  [5127, "WeekdayName(2, True)", "Mo"],
  [5127, "WeekdayName(7, True)", "Sa"],
  [5127, "CStr(1234567.891)", "1234567.891"],
  [5127, "FormatNumber(1234567.891)", "1’234’567.89"],
  [5127, "FormatCurrency(1234567.891)", "CHF 1’234’567.89"],
  [5127, "FormatPercent(1234567.891)", "123’456’789.10%"],
  [5127, "FormatNumber(-1234.567)", "-1’234.57"],
  [5127, "FormatCurrency(-1234.567)", "-CHF 1’234.57"],
  [5127, "FormatPercent(0.125)", "12.50%"],
  [5127, "FormatCurrency(1235)", "CHF 1’235.00"],
  [5127, "FormatNumber(1234567.891, 2, -1, -1, 0)", "1234567.89"],
  [5127, "FormatNumber(-1234.567, 2, -1, True, -1)", "(1’234.57)"],
  [5127, "FormatNumber(1234567.891, 0)", "1’234’568"],
  [5127, "FormatPercent(0.125, 0)", "13%"],
  // ts-ZA (1073)
  [1073, "CStr(DateSerial(2026,1,5) + TimeSerial(0,0,0))", "2026-01-05"],
  [
    1073,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 0)",
    "2026-01-05 09:05:00",
  ],
  [
    1073,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 1)",
    "2026 Sunguti 5, Musumbhunuku",
  ],
  [
    1073,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 2)",
    "2026-01-05",
  ],
  [
    1073,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 3)",
    "09:05:00",
  ],
  [
    1073,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 4)",
    "09:05",
  ],
  [
    1073,
    "FormatDateTime(DateSerial(2026,11,25) + TimeSerial(22,30,45), 1)",
    "2026 Hukuri 25, Ravunharhu",
  ],
  [
    1073,
    "FormatDateTime(DateSerial(2026,11,25) + TimeSerial(22,30,45), 2)",
    "2026-11-25",
  ],
  [
    1073,
    "FormatDateTime(DateSerial(2026,12,31) + TimeSerial(23,59,59), 1)",
    "2026 N’wendzamhala 31, Ravumune",
  ],
  [
    1073,
    "FormatDateTime(DateSerial(2026,12,31) + TimeSerial(23,59,59), 2)",
    "2026-12-31",
  ],
  [1073, "MonthName(3)", "Nyenyankulu"],
  [1073, "MonthName(3, True)", "Kul"],
  [1073, "MonthName(12, True)", "N’w"],
  [1073, "WeekdayName(2)", "Musumbhunuku"],
  [1073, "WeekdayName(2, True)", "Mus"],
  [1073, "WeekdayName(7, True)", "Mug"],
  [1073, "CStr(1234567.891)", "1234567,891"],
  [1073, "FormatNumber(1234567.891)", "1\u{00a0}234\u{00a0}567,89"],
  [1073, "FormatCurrency(1234567.891)", "R 1\u{00a0}234\u{00a0}567,89"],
  [1073, "FormatPercent(1234567.891)", "123\u{00a0}456\u{00a0}789,10%"],
  [1073, "FormatNumber(-1234.567)", "-1\u{00a0}234,57"],
  [1073, "FormatCurrency(-1234.567)", "-R 1\u{00a0}234,57"],
  [1073, "FormatPercent(0.125)", "12,50%"],
  [1073, "FormatCurrency(1235)", "R 1\u{00a0}235,00"],
  [1073, "FormatNumber(1234567.891, 2, -1, -1, 0)", "1234567,89"],
  [1073, "FormatNumber(-1234.567, 2, -1, True, -1)", "(1\u{00a0}234,57)"],
  [1073, "FormatNumber(1234567.891, 0)", "1\u{00a0}234\u{00a0}568"],
  [1073, "FormatPercent(0.125, 0)", "13%"],
];

const FROM_THE_SWEEP: ReadonlyArray<readonly [number, string, string]> = [
  // th-TH (1054)
  [1054, "CStr(DateSerial(2026,1,5) + TimeSerial(0,0,0))", "5/1/2569"],
  [
    1054,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 0)",
    "5/1/2569 9:05:00",
  ],
  [
    1054,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 1)",
    "5 มกราคม 2569",
  ],
  [
    1054,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 2)",
    "5/1/2569",
  ],
  [
    1054,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 3)",
    "9:05:00",
  ],
  [
    1054,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 4)",
    "09:05",
  ],
  [
    1054,
    "FormatDateTime(DateSerial(2026,11,25) + TimeSerial(22,30,45), 1)",
    "25 พฤศจิกายน 2569",
  ],
  [
    1054,
    "FormatDateTime(DateSerial(2026,11,25) + TimeSerial(22,30,45), 2)",
    "25/11/2569",
  ],
  [
    1054,
    "FormatDateTime(DateSerial(2026,12,31) + TimeSerial(23,59,59), 1)",
    "31 ธันวาคม 2569",
  ],
  [
    1054,
    "FormatDateTime(DateSerial(2026,12,31) + TimeSerial(23,59,59), 2)",
    "31/12/2569",
  ],
  [1054, "MonthName(3)", "มีนาคม"],
  [1054, "MonthName(3, True)", "มี.ค."],
  [1054, "MonthName(12, True)", "ธ.ค."],
  [1054, "WeekdayName(2)", "จันทร์"],
  [1054, "WeekdayName(2, True)", "จ."],
  [1054, "WeekdayName(7, True)", "ส."],
  [1054, "CStr(1234567.891)", "1234567.891"],
  [1054, "FormatNumber(1234567.891)", "1,234,567.89"],
  [1054, "FormatCurrency(1234567.891)", "฿1,234,567.89"],
  [1054, "FormatPercent(1234567.891)", "123,456,789.10%"],
  [1054, "FormatNumber(-1234.567)", "-1,234.57"],
  [1054, "FormatCurrency(-1234.567)", "-฿1,234.57"],
  [1054, "FormatPercent(0.125)", "12.50%"],
  [1054, "FormatCurrency(1235)", "฿1,235.00"],
  [1054, "FormatNumber(1234567.891, 2, -1, -1, 0)", "1234567.89"],
  [1054, "FormatNumber(-1234.567, 2, -1, True, -1)", "(1,234.57)"],
  [1054, "FormatNumber(1234567.891, 0)", "1,234,568"],
  [1054, "FormatPercent(0.125, 0)", "13%"],
  // ar-SA (1025)
  [1025, "CStr(DateSerial(2026,1,5) + TimeSerial(0,0,0))", "05/01/2026"],
  [
    1025,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 0)",
    "05/01/2026 09:05:00 ص",
  ],
  [
    1025,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 1)",
    "\u{200f}16/\u{200f}رجب/\u{200f}1447",
  ],
  [
    1025,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 2)",
    "\u{200f}16/\u{200f}07/\u{200f}47",
  ],
  [
    1025,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 3)",
    "09:05:00 ص",
  ],
  [
    1025,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 4)",
    "09:05",
  ],
  [
    1025,
    "FormatDateTime(DateSerial(2026,11,25) + TimeSerial(22,30,45), 1)",
    "\u{200f}15/\u{200f}جمادى\u{00a0}الثانية/\u{200f}1448",
  ],
  [
    1025,
    "FormatDateTime(DateSerial(2026,11,25) + TimeSerial(22,30,45), 2)",
    "\u{200f}15/\u{200f}06/\u{200f}48",
  ],
  [
    1025,
    "FormatDateTime(DateSerial(2026,12,31) + TimeSerial(23,59,59), 1)",
    "\u{200f}22/\u{200f}رجب/\u{200f}1448",
  ],
  [
    1025,
    "FormatDateTime(DateSerial(2026,12,31) + TimeSerial(23,59,59), 2)",
    "\u{200f}22/\u{200f}07/\u{200f}48",
  ],
  [1025, "MonthName(3)", "مارس"],
  [1025, "MonthName(3, True)", "مارس"],
  [1025, "MonthName(12, True)", "ديسمبر"],
  [1025, "WeekdayName(2)", "الإثنين"],
  [1025, "WeekdayName(2, True)", "الإثنين"],
  [1025, "WeekdayName(7, True)", "السبت"],
  [1025, "CStr(1234567.891)", "1234567.891"],
  [1025, "FormatNumber(1234567.891)", "1,234,567.89"],
  [1025, "FormatCurrency(1234567.891)", "⃁ 1,234,567.89"],
  [1025, "FormatPercent(1234567.891)", "123,456,789.10%"],
  [1025, "FormatNumber(-1234.567)", "-1,234.57"],
  [1025, "FormatCurrency(-1234.567)", "⃁ -1,234.57"],
  [1025, "FormatPercent(0.125)", "12.50%"],
  [1025, "FormatCurrency(1235)", "⃁ 1,235.00"],
  [1025, "FormatNumber(1234567.891, 2, -1, -1, 0)", "1234567.89"],
  [1025, "FormatNumber(-1234.567, 2, -1, True, -1)", "(1,234.57)"],
  [1025, "FormatNumber(1234567.891, 0)", "1,234,568"],
  [1025, "FormatPercent(0.125, 0)", "13%"],
  // fa-IR (1065)
  [1065, "CStr(DateSerial(2026,1,5) + TimeSerial(0,0,0))", "05/01/2026"],
  [
    1065,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 0)",
    "05/01/2026 09:05:00 ق.ظ",
  ],
  [
    1065,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 1)",
    "دوشنبه, 15 دی 1404",
  ],
  [
    1065,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 2)",
    "15/10/1404",
  ],
  [
    1065,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 3)",
    "09:05:00 ق.ظ",
  ],
  [
    1065,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 4)",
    "09:05",
  ],
  [
    1065,
    "FormatDateTime(DateSerial(2026,11,25) + TimeSerial(22,30,45), 1)",
    "چهارشنبه, 4 آذر 1405",
  ],
  [
    1065,
    "FormatDateTime(DateSerial(2026,11,25) + TimeSerial(22,30,45), 2)",
    "04/09/1405",
  ],
  [
    1065,
    "FormatDateTime(DateSerial(2026,12,31) + TimeSerial(23,59,59), 1)",
    "پنجشنبه, 10 دی 1405",
  ],
  [
    1065,
    "FormatDateTime(DateSerial(2026,12,31) + TimeSerial(23,59,59), 2)",
    "10/10/1405",
  ],
  [1065, "MonthName(3)", "مارس"],
  [1065, "MonthName(3, True)", "مارس"],
  [1065, "MonthName(12, True)", "دسامبر"],
  [1065, "WeekdayName(2)", "دوشنبه"],
  [1065, "WeekdayName(2, True)", "دوشنبه"],
  [1065, "WeekdayName(7, True)", "شنبه"],
  [1065, "CStr(1234567.891)", "1234567/891"],
  [1065, "FormatNumber(1234567.891)", "1,234,567/89"],
  [1065, "FormatCurrency(1234567.891)", "1,234,567/89ريال"],
  [1065, "FormatPercent(1234567.891)", "123,456,789/10%"],
  [1065, "FormatNumber(-1234.567)", "1,234/57-"],
  [1065, "FormatCurrency(-1234.567)", "1,234/57-ريال"],
  [1065, "FormatPercent(0.125)", "12/50%"],
  [1065, "FormatCurrency(1235)", "1,235/00ريال"],
  [1065, "FormatNumber(1234567.891, 2, -1, -1, 0)", "1234567/89"],
  [1065, "FormatNumber(-1234.567, 2, -1, True, -1)", "(1,234/57)"],
  [1065, "FormatNumber(1234567.891, 0)", "1,234,568"],
  [1065, "FormatPercent(0.125, 0)", "13%"],
  // hr-HR (1050)
  [1050, "CStr(DateSerial(2026,1,5) + TimeSerial(0,0,0))", "5.1.2026."],
  [
    1050,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 0)",
    "5.1.2026. 9:05:00",
  ],
  [
    1050,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 1)",
    "5. siječnja 2026.",
  ],
  [
    1050,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 2)",
    "5.1.2026.",
  ],
  [
    1050,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 3)",
    "9:05:00",
  ],
  [
    1050,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 4)",
    "09:05",
  ],
  [
    1050,
    "FormatDateTime(DateSerial(2026,11,25) + TimeSerial(22,30,45), 1)",
    "25. studenog 2026.",
  ],
  [
    1050,
    "FormatDateTime(DateSerial(2026,11,25) + TimeSerial(22,30,45), 2)",
    "25.11.2026.",
  ],
  [
    1050,
    "FormatDateTime(DateSerial(2026,12,31) + TimeSerial(23,59,59), 1)",
    "31. prosinca 2026.",
  ],
  [
    1050,
    "FormatDateTime(DateSerial(2026,12,31) + TimeSerial(23,59,59), 2)",
    "31.12.2026.",
  ],
  [1050, "MonthName(3)", "ožujak"],
  [1050, "MonthName(3, True)", "ožu"],
  [1050, "MonthName(12, True)", "pro"],
  [1050, "WeekdayName(2)", "ponedjeljak"],
  [1050, "WeekdayName(2, True)", "pon"],
  [1050, "WeekdayName(7, True)", "sub"],
  [1050, "CStr(1234567.891)", "1234567,891"],
  [1050, "FormatNumber(1234567.891)", "1.234.567,89"],
  [1050, "FormatCurrency(1234567.891)", "1.234.567,89 €"],
  [1050, "FormatPercent(1234567.891)", "123.456.789,10%"],
  [1050, "FormatNumber(-1234.567)", "-1.234,57"],
  [1050, "FormatCurrency(-1234.567)", "-1.234,57 €"],
  [1050, "FormatPercent(0.125)", "12,50%"],
  [1050, "FormatCurrency(1235)", "1.235,00 €"],
  [1050, "FormatNumber(1234567.891, 2, -1, -1, 0)", "1234567,89"],
  [1050, "FormatNumber(-1234.567, 2, -1, True, -1)", "(1.234,57)"],
  [1050, "FormatNumber(1234567.891, 0)", "1.234.568"],
  [1050, "FormatPercent(0.125, 0)", "13%"],
  // hi-IN (1081)
  [1081, "CStr(DateSerial(2026,1,5) + TimeSerial(0,0,0))", "05-01-2026"],
  [
    1081,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 0)",
    "05-01-2026 09:05:00",
  ],
  [
    1081,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 1)",
    "05 जनवरी 2026",
  ],
  [
    1081,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 2)",
    "05-01-2026",
  ],
  [
    1081,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 3)",
    "09:05:00",
  ],
  [
    1081,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 4)",
    "09:05",
  ],
  [
    1081,
    "FormatDateTime(DateSerial(2026,11,25) + TimeSerial(22,30,45), 1)",
    "25 नवम्बर 2026",
  ],
  [
    1081,
    "FormatDateTime(DateSerial(2026,11,25) + TimeSerial(22,30,45), 2)",
    "25-11-2026",
  ],
  [
    1081,
    "FormatDateTime(DateSerial(2026,12,31) + TimeSerial(23,59,59), 1)",
    "31 दिसम्बर 2026",
  ],
  [
    1081,
    "FormatDateTime(DateSerial(2026,12,31) + TimeSerial(23,59,59), 2)",
    "31-12-2026",
  ],
  [1081, "MonthName(3)", "मार्च"],
  [1081, "MonthName(3, True)", "मार्च"],
  [1081, "MonthName(12, True)", "दिसम्बर"],
  [1081, "WeekdayName(2)", "सोमवार"],
  [1081, "WeekdayName(2, True)", "सोम."],
  [1081, "WeekdayName(7, True)", "शनि."],
  [1081, "CStr(1234567.891)", "1234567.891"],
  [1081, "FormatNumber(1234567.891)", "1,234,567.89"],
  [1081, "FormatCurrency(1234567.891)", "₹1,234,567.89"],
  [1081, "FormatPercent(1234567.891)", "123,456,789.10%"],
  [1081, "FormatNumber(-1234.567)", "-1,234.57"],
  [1081, "FormatCurrency(-1234.567)", "₹ -1,234.57"],
  [1081, "FormatPercent(0.125)", "12.50%"],
  [1081, "FormatCurrency(1235)", "₹1,235.00"],
  [1081, "FormatNumber(1234567.891, 2, -1, -1, 0)", "1234567.89"],
  [1081, "FormatNumber(-1234.567, 2, -1, True, -1)", "(1,234.57)"],
  [1081, "FormatNumber(1234567.891, 0)", "1,234,568"],
  [1081, "FormatPercent(0.125, 0)", "13%"],
  // ar-DZ (5121)
  [5121, "CStr(DateSerial(2026,1,5) + TimeSerial(0,0,0))", "05-01-2026"],
  [
    5121,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 0)",
    "05-01-2026 9:05:00",
  ],
  [
    5121,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 1)",
    "\u{200f}05 \u{200f}جانفييه, \u{200f}2026",
  ],
  [
    5121,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 2)",
    "\u{200f}05-\u{200f}01-\u{200f}2026",
  ],
  [
    5121,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 3)",
    "9:05:00",
  ],
  [
    5121,
    "FormatDateTime(DateSerial(2026,1,5) + TimeSerial(9,5,0), 4)",
    "09:05",
  ],
  [
    5121,
    "FormatDateTime(DateSerial(2026,11,25) + TimeSerial(22,30,45), 1)",
    "\u{200f}25 \u{200f}نوفمبر, \u{200f}2026",
  ],
  [
    5121,
    "FormatDateTime(DateSerial(2026,11,25) + TimeSerial(22,30,45), 2)",
    "\u{200f}25-\u{200f}11-\u{200f}2026",
  ],
  [
    5121,
    "FormatDateTime(DateSerial(2026,12,31) + TimeSerial(23,59,59), 1)",
    "\u{200f}31 \u{200f}ديسمبر, \u{200f}2026",
  ],
  [
    5121,
    "FormatDateTime(DateSerial(2026,12,31) + TimeSerial(23,59,59), 2)",
    "\u{200f}31-\u{200f}12-\u{200f}2026",
  ],
  [5121, "MonthName(3)", "مارس"],
  [5121, "MonthName(3, True)", "مارس"],
  [5121, "MonthName(12, True)", "ديسمبر"],
  [5121, "WeekdayName(2)", "الإثنين"],
  [5121, "WeekdayName(2, True)", "الإثنين"],
  [5121, "WeekdayName(7, True)", "السبت"],
  [5121, "CStr(1234567.891)", "1234567.891"],
  [5121, "FormatNumber(1234567.891)", "1,234,567.89"],
  [5121, "FormatCurrency(1234567.891)", "1,234,567.89 د.ج.\u{200f}"],
  [5121, "FormatPercent(1234567.891)", "123,456,789.10%"],
  [5121, "FormatNumber(-1234.567)", "-1,234.57"],
  [5121, "FormatCurrency(-1234.567)", "-1,234.57 د.ج.\u{200f}"],
  [5121, "FormatPercent(0.125)", "12.50%"],
  [5121, "FormatCurrency(1235)", "1,235.00 د.ج.\u{200f}"],
  [5121, "FormatNumber(1234567.891, 2, -1, -1, 0)", "1234567.89"],
  [5121, "FormatNumber(-1234.567, 2, -1, True, -1)", "(1,234.57)"],
  [5121, "FormatNumber(1234567.891, 0)", "1,234,568"],
  [5121, "FormatPercent(0.125, 0)", "13%"],
];

for (const [title, cells] of [
  ["German and Tsonga", GERMAN_AND_TSONGA],
  ["the rest of the sweep", FROM_THE_SWEEP],
] as const) {
  describe(`${title} follow vbscript.dll`, () => {
    for (const [lcid, expression, expected] of cells) {
      it(`${lcid}: ${expression}`, () => {
        expect(answer(lcid, expression)).toBe(expected);
      });
    }
  });
}
