import { describe, it, expect } from "vitest";
import { VbsEngine } from "../index.ts";

// What a Single and a Currency are, and what arithmetic does with them, measured against vbscript.dll.
//
// Every cell was read off a fresh cscript under SetLocale 1031 on Windows build 26200:
//
//   * A Single is an IEEE 754 single, so CSng(1/3) holds 0.3333333432674408 and CSng(1/3) * 3 is 1.
//   * A Currency is a count of ten-thousandths, so CCur(0.1) + CCur(0.2) is exactly 0.3 and
//     CCur(1/3) * 3 is 0.9999. It holds eighteen digits, which a JavaScript number cannot.
//   * A type name that says Single or Currency is the point of the first block: the value alone would not
//     tell them apart from a Double.
//
// The last three cells are the refusals. CCur(1E+15) is past the Currency range, CSng(1E+39) is past the
// Single range, and CCur(922337203685477.5807) is past it too: that is the documented largest amount, and a
// Double cannot hold its last digit, so the literal arrives already rounded up and over the top.

const CELLS: ReadonlyArray<readonly [string, string]> = [
  // The type each operator gives.
  ["TypeName(CSng(1) + CSng(2))", "Single"],
  ["TypeName(CSng(1) + 2)", "Single"],
  ["TypeName(CSng(1) + CLng(2))", "Double"],
  ["TypeName(CSng(1) + CDbl(2))", "Double"],
  ["TypeName(CSng(1) / CSng(2))", "Single"],
  ["TypeName(CSng(1) / 3)", "Single"],
  ["TypeName(CSng(1) / CLng(2))", "Double"],
  ["TypeName(CSng(2) * CLng(3))", "Double"],
  ["TypeName(CCur(1) + CCur(2))", "Currency"],
  ["TypeName(CCur(1) + CDbl(2))", "Currency"],
  ["TypeName(CCur(1) + CSng(2))", "Currency"],
  ["TypeName(CCur(2) * CCur(3))", "Currency"],
  ["TypeName(CCur(2) * 3)", "Currency"],
  ["TypeName(CCur(2) * CDbl(3))", "Double"],
  ["TypeName(CCur(1) / CCur(2))", "Double"],
  ["TypeName(CCur(7) \\ 2)", "Long"],
  ["TypeName(CCur(7) Mod 2)", "Long"],
  ["TypeName(7 Mod 2)", "Integer"],
  ["TypeName(CDbl(7) Mod 2)", "Long"],
  ["TypeName(CByte(2) + CByte(3))", "Byte"],
  ["TypeName(CByte(200) + CByte(200))", "Integer"],
  ["TypeName(-CSng(1))", "Single"],
  ["TypeName(-CCur(1))", "Currency"],
  ["TypeName(Abs(CSng(-1.5)))", "Single"],
  ["TypeName(Abs(CCur(-1.5)))", "Currency"],
  ["TypeName(Int(CCur(1.5)))", "Currency"],
  ["TypeName(Fix(CCur(1.5)))", "Currency"],
  ["TypeName(Sgn(CCur(1.5)))", "Integer"],
  ["TypeName(Round(CCur(1.5)))", "Currency"],
  ["TypeName(Round(CSng(1.5)))", "Single"],
  ["TypeName(Sqr(CCur(4)))", "Double"],
  // What a Single keeps.
  ["CStr(CSng(1 / 3) * 3)", "1"],
  ["CStr(CSng(16777217))", "1,677722E+07"],
  ["CStr(CSng(0.1) + CSng(0.2))", "0,3"],
  ["CStr(CDbl(CSng(1 / 3)))", "0,333333343267441"],
  ["CStr(CSng(1) / 3)", "0,3333333"],
  // What a Currency keeps.
  ["CStr(CCur(0.1) + CCur(0.2))", "0,3"],
  ["CStr(CCur(1 / 3))", "0,3333"],
  ["CStr(CCur(1 / 3) * 3)", "0,9999"],
  ["CStr(CCur(123456789012.3456))", "123456789012,3456"],
  ["CStr(CCur(123456789012.3456) + CCur(0.0001))", "123456789012,3457"],
  ["CStr(CCur(0.1) + CCur(0.2) + CCur(0.3) + CCur(0.4))", "1"],
  ["CStr(CCur(1.005) * CCur(3))", "3,015"],
  ["CStr(CCur(0.00005))", "0,0001"],
  ["CStr(CCur(-0.00005))", "-0,0001"],
  ["CStr(CDbl(CCur(1 / 3)))", "0,3333"],
  ["CStr(CSng(CCur(0.123456789)))", "0,1235"],
  ["CStr(CByte(200) + CByte(200))", "400"],
  ["CStr(Abs(CSng(-1.5)))", "1,5"],
  ["CStr(Int(CCur(1.5)))", "1"],
  // What the types refuse.
  ["CStr(CCur(922337203685477.5807))", "ERR 6"],
  ["CStr(CCur(1E+15))", "ERR 6"],
  ["CStr(CSng(1E+39))", "ERR 6"],
  // A Null is Null for Abs, Int and Fix, which a fresh cscript shows, while Sgn, Sqr, Round, Exp, Log, Sin,
  // Cos, Tan, Atn, CByte, CSng, CCur and CBool raise the invalid use of Null for one.
  ["CStr(IsNull(Abs(Null)))", "True"],
  ["CStr(IsNull(Int(Null)))", "True"],
  ["CStr(IsNull(Fix(Null)))", "True"],
  ["Sgn(Null)", "ERR 94"],
  ["Sqr(Null)", "ERR 94"],
  ["Round(Null)", "ERR 94"],
  ["Exp(Null)", "ERR 94"],
  ["Log(Null)", "ERR 94"],
  ["Sin(Null)", "ERR 94"],
  ["CByte(Null)", "ERR 94"],
  ["CBool(Null)", "ERR 94"],
];

function answer(expression: string): unknown {
  const engine = new VbsEngine();
  try {
    engine.executeStatement(`SetLocale 1031\nr = (${expression})`);
    const failure = engine.error;
    if (failure) return "ERR " + String(failure.number);
    return engine.eval("r");
  } finally {
    engine.executeStatement("SetLocale 1033");
  }
}

describe("a Single and a Currency hold what VBScript holds", () => {
  for (const [expression, expected] of CELLS) {
    it(`${expression} is ${expected}`, () => {
      expect(answer(expression)).toBe(expected);
    });
  }
});

// What CInt, CLng and CByte round to, and where each refuses. Measured the same way, and written with the
// type beside the value because the roundings agree and only the refusals tell them apart.
const CONVERSION_CELLS: ReadonlyArray<readonly [string, string]> = [
  // CInt and CLng round a tie to the even neighbour, and the negative side mirrors it.
  ["CInt(0.5)", "Integer | 0"],
  ["CInt(1.5)", "Integer | 2"],
  ["CInt(2.5)", "Integer | 2"],
  ["CInt(3.5)", "Integer | 4"],
  ["CInt(-0.5)", "Integer | 0"],
  ["CInt(-1.5)", "Integer | -2"],
  ["CInt(-2.5)", "Integer | -2"],
  ["CInt(2.4)", "Integer | 2"],
  ["CInt(2.6)", "Integer | 3"],
  ["CInt(32767)", "Integer | 32767"],
  ["CInt(32767.4)", "Integer | 32767"],
  ["CInt(32768)", "ERR 6"],
  ["CInt(32767.6)", "ERR 6"],
  ["CInt(-32768)", "Integer | -32768"],
  ["CInt(-32769)", "ERR 6"],
  ['CInt("1,5")', "Integer | 2"],
  ["CInt(Empty)", "Integer | 0"],
  ["CInt(CCur(1.5))", "Integer | 2"],
  ["CInt(CSng(1.5))", "Integer | 2"],
  ["CLng(0.5)", "Long | 0"],
  ["CLng(1.5)", "Long | 2"],
  ["CLng(2.5)", "Long | 2"],
  ["CLng(2147483647)", "Long | 2147483647"],
  ["CLng(2147483647.4)", "Long | 2147483647"],
  ["CLng(2147483648)", "ERR 6"],
  ["CLng(-2147483648)", "Long | -2147483648"],
  ["CLng(-2147483649)", "ERR 6"],
  ["CByte(2.5)", "Byte | 2"],
  ["CByte(0.5)", "Byte | 0"],
  ["CByte(1.5)", "Byte | 2"],
  ["CByte(255)", "Byte | 255"],
  ["CByte(255.4)", "Byte | 255"],
  ["CByte(256)", "ERR 6"],
  ["CByte(-1)", "ERR 6"],
  ["CByte(-0.5)", "Byte | 0"],
  ['CByte("255")', "Byte | 255"],
  ["TypeName(CBool(1))", "String | Boolean"],
  ["TypeName(CDbl(1000000000000))", "String | Double"],
];

function typed(expression: string): unknown {
  const engine = new VbsEngine();
  try {
    engine.executeStatement(
      `SetLocale 1031\nr = TypeName(${expression}) & " | " & CStr(${expression})`,
    );
    const failure = engine.error;
    if (failure) return "ERR " + String(failure.number);
    return engine.eval("r");
  } finally {
    engine.executeStatement("SetLocale 1033");
  }
}

describe("CInt, CLng and CByte round and refuse as VBScript does", () => {
  for (const [expression, expected] of CONVERSION_CELLS) {
    it(`${expression} is ${expected}`, () => {
      expect(typed(expression)).toBe(expected);
    });
  }
});

// Rounding a Currency with a fraction count, and comparing a Currency with anything. Measured the same way.
const CURRENCY_EDGE_CELLS: ReadonlyArray<readonly [string, string]> = [
  // A Currency rounds on its own scale, so Round(CCur(1.005), 2) reaches the tie and takes the even
  // neighbour: 1, where two doubles would lose the fifth decimal first and answer 1.01.
  ["Round(CCur(1.2345), 2)", "Currency | 1,23"],
  ["Round(CCur(1.005), 2)", "Currency | 1"],
  ["Round(CCur(2.5), 0)", "Currency | 2"],
  ["Round(CCur(1.5), 0)", "Currency | 2"],
  ["Round(CCur(-1.5), 0)", "Currency | -2"],
  ["Round(CCur(1.5), 1)", "Currency | 1,5"],
  ["Round(CCur(1234.5678), 2)", "Currency | 1234,57"],
  ["Round(CCur(-1234.5678), 2)", "Currency | -1234,57"],
  ["Round(CCur(1.00005), 4)", "Currency | 1,0001"],
  // A comparison brings the other side to the Currency's own scale, then compares the scaled integers.
  ["CCur(1) = 1", "Boolean | True"],
  ["CCur(0.12345) = 0.12345", "Boolean | True"],
  ["CCur(1) = 1.00001", "Boolean | True"],
  ["CCur(123456789012.3456) = CCur(123456789012.3457)", "Boolean | False"],
  ["CCur(123456789012.3456) < CCur(123456789012.3457)", "Boolean | True"],
  ["1 = CCur(1)", "Boolean | True"],
  ["CCur(2) > 1.9999", "Boolean | True"],
  ["CCur(-0.00005) < 0", "Boolean | True"],
  ["TypeName(CCur(1) = 1)", "String | Boolean"],
];

describe("a Currency rounds and compares on its own scale", () => {
  for (const [expression, expected] of CURRENCY_EDGE_CELLS) {
    it(`${expression} is ${expected}`, () => {
      expect(typed(expression)).toBe(expected);
    });
  }
});

// CLngLng and the LongLong type are the engine's own extension. This vbscript.dll has no such function at
// all: a direct call raises error 500 and a call through Eval raises error 13, so there is no reference
// answer to hold these cells against. They pin the rules the engine chose for itself: the rounding CLng
// uses, since CLng is the 32 bit member of the same family, the 64 bit range, and a type that survives
// arithmetic.
const EXTENSION_CELLS: ReadonlyArray<readonly [string, string]> = [
  ["CLngLng(1)", "LongLong | 1"],
  ["CLngLng(1.5)", "LongLong | 2"],
  ["CLngLng(2.5)", "LongLong | 2"],
  ["CLngLng(0.5)", "LongLong | 0"],
  ["CLngLng(-1.5)", "LongLong | -2"],
  ["CLngLng(-2.5)", "LongLong | -2"],
  ["CLngLng(1.4)", "LongLong | 1"],
  ['CLngLng("123")', "LongLong | 123"],
  ['CLngLng("1.5")', "LongLong | 2"],
  ["CLngLng(Empty)", "LongLong | 0"],
  ["CLngLng(CCur(1.5))", "LongLong | 2"],
  ["CLngLng(1000000000000000000)", "LongLong | 1000000000000000000"],
  ["CLngLng(1E+19)", "ERR 6"],
  ["CLngLng(1) + 1", "LongLong | 2"],
  ["CLngLng(1) + CSng(1)", "Double | 2"],
  ["CLngLng(1) + CDbl(1)", "Double | 2"],
  ["CLngLng(1) + CCur(1)", "Double | 2"],
  ["-CLngLng(1)", "LongLong | -1"],
  ["CLngLng(9000000000000000000) + CLngLng(9000000000000000000)", "ERR 6"],
];

describe("the LongLong extension keeps the rules the engine chose", () => {
  for (const [expression, expected] of EXTENSION_CELLS) {
    it(`${expression} is ${expected}`, () => {
      expect(typed(expression)).toBe(expected);
    });
  }
});
