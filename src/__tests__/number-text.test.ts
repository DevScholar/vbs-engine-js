import { describe, it, expect } from "vitest";
import { VbsEngine } from "../index.ts";

// CStr of a number, and the type a whole number gets, measured against vbscript.dll.
//
// The three numeric types carry different precision, and each was measured with cscript //E:vbscript under
// SetLocale 1031 on Windows build 26200, one fresh process per locale:
//
//   * A Double keeps fifteen significant digits (CStr(1/3) is "0,333333333333333", and CStr(0.1+0.2) is
//     "0,3" where the value itself is 0.30000000000000004).
//   * A Single keeps seven (CStr(CSng(1234567.89)) is "1234568").
//   * A Currency keeps four decimal places rather than a count of significant digits, rounds away from
//     zero, and keeps the sign through a rounding to zero (CStr(CCur(-0.00005)) is "-0,0001").
//
// Outside a window of exponents the value is written with one, and the exponent carries a sign and at least
// two digits: a Double writes "1E+15" and "1E-16" either side of its window, a Single writes "1,234568E+07"
// and "4E-08" either side of its own. A whole number is what decides whether that happens at all, because a
// VBScript Long is 32 bits wide and anything past it is a Double.

const CELLS: ReadonlyArray<readonly [string, string]> = [
  // A Double keeps fifteen significant digits, and trailing zeros are not written.
  ["CStr(1/3)", "0,333333333333333"],
  ["CStr(1/7)", "0,142857142857143"],
  ["CStr(1.234567890123456)", "1,23456789012346"],
  ["CStr(0.1 + 0.2)", "0,3"],
  ["CStr(0.12345)", "0,12345"],
  ["CStr(0.000001)", "0,000001"],
  ["CStr(123456789012345)", "123456789012345"],
  ["CStr(1000000.5)", "1000000,5"],
  // An exponent is written only outside the window of that type.
  ["CStr(100000000000000)", "100000000000000"],
  ["CStr(1000000000000000)", "1E+15"],
  ["CStr(1000000000000000000000)", "1E+21"],
  ["CStr(-1000000000000000000000)", "-1E+21"],
  ["CStr(1.234567890123456E+15)", "1,23456789012346E+15"],
  ["CStr(0.000000000000001)", "0,000000000000001"],
  ["CStr(0.0000000000000001)", "1E-16"],
  // A Single keeps seven digits and has its own, narrower window.
  ["CStr(CSng(1/3))", "0,3333333"],
  ["CStr(CSng(1.2345678))", "1,234568"],
  ["CStr(CSng(12345678))", "1,234568E+07"],
  ["CStr(CSng(0.00000004))", "4E-08"],
  ["CStr(CSng(3.4E+38))", "3,4E+38"],
  // A Currency keeps four decimal places and keeps the sign through a rounding to zero.
  ["CStr(CCur(1/3))", "0,3333"],
  ["CStr(CCur(2.675))", "2,675"],
  ["CStr(CCur(1.5))", "1,5"],
  ["CStr(CCur(0.00005))", "0,0001"],
  ["CStr(CCur(-0.00005))", "-0,0001"],
  ["CStr(CCur(0.00004))", "0"],
  ["CStr(CCur(123456789012.3456))", "123456789012,3456"],
  // A whole number past the Long range is a Double, which is what CStr writes with an exponent.
  ["TypeName(32767)", "Integer"],
  ["TypeName(32768)", "Long"],
  ["TypeName(2147483647)", "Long"],
  ["TypeName(2147483648)", "Double"],
  ["TypeName(1000000 * 1000000)", "Double"],
  ["TypeName(1000000000000)", "Double"],
];

function answer(expression: string, lcid = 1031): unknown {
  const engine = new VbsEngine();
  try {
    engine.executeStatement(`SetLocale ${lcid}\nr = (${expression})`);
    expect(engine.error).toBeNull();
    return engine.eval("r");
  } finally {
    engine.executeStatement("SetLocale 1033");
  }
}

describe("a number written as a string follows vbscript.dll", () => {
  for (const [expression, expected] of CELLS) {
    it(`${expression} is ${expected}`, () => {
      expect(answer(expression)).toBe(expected);
    });
  }

  it("writes the fraction mark of the locale it runs under", () => {
    expect(answer("CStr(1/3)", 1031)).toBe("0,333333333333333");
    expect(answer("CStr(1/3)", 2055)).toBe("0.333333333333333");
    expect(answer("CStr(1/3)", 1065)).toBe("0/333333333333333");
  });
});
