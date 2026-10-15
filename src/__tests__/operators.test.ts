import { describe, it, expect } from 'vitest';
import { VbsEngine } from '../index.ts';

// Every operator against every pair of VBScript types, measured against a fresh cscript under 1033: twelve
// operands (Empty, Null, Boolean, Byte, Integer, Long, Single, Double, Currency, Date and two Strings) under
// twenty operators, plus nineteen unary cells, 2899 in all. All of them agree with the engine. The capture is
// arith2-1033.txt in the probe directory.
//
// Why the probe writes every expression out instead of building it as text and calling Eval. Two reads of
// the same cell disagree under Eval, and the written out form is what a compiled script does:
//
//   * A bitwise operator converts an Empty operand in place: after "Empty And Empty" the variable itself is
//     a Long 0. An Eval driven table therefore measured every later cell against a changed operand.
//   * Eval reads a numeric value beside a String as a String comparison, so CByte(200) < "2" answered True
//     where the written out form answers False. The engine, which compiles its input, is on the False side.
//   * The operands of the written out table are literals, so no cell can leave a mark on the next one.
//
// The rules the cells lock in, all of them measured:
//
//   * Is compares two object references and nothing else, so a plain value on either side is error 424.
//   * The five bitwise operators keep the wider of their two operands: Byte And Byte is a Byte, Byte And
//     Integer an Integer, Boolean And Byte an Integer, and everything else a Long. A narrower type wraps:
//     Byte Eqv Byte is the Byte 255 where the whole number is -1.
//   * A floating point operand is rounded to the nearest whole number, a tie to the even neighbour, before a
//     bitwise operator or integer division reads it, so CByte(200) \ CSng(1.5) is 100 and CByte(200) Or
//     CSng(1.5) is 202.
//   * Integer division cuts toward zero and keeps an Integer for narrow operands: CByte(200) \ CInt(-3) is
//     -66, and CByte(200) Mod CByte(200) is the Byte 0.
//   * A text that is no number is a type mismatch under all of these, even beside a Null.
//   * A Null is read operator by operator: And and Or keep the side that already decides the answer, Xor
//     and Eqv are Null, and Imp answers its right side where the left is Null.
//   * Not keeps its operand's type, and unary minus keeps a Single, a Double, a Currency and a Date, turns a
//     numeric String into a Double and a Byte into an Integer.
//   * A power that leaves the Double range is error 5 where no value exists and error 6 where the value
//     exists but is out of range.
//   * Division answers a Single while both operands are narrow, an Empty counting as narrow, and a Double
//     otherwise. Zero divided by zero is error 6 where a dividend that is not zero is error 11.
//   * A Long keeps its width beside anything whole, and a product that leaves the Long range is a Double.
//   * A negative zero keeps its sign in CStr, for a Single and a Double alike, and a Currency and an Integer
//     have no such value.
const CELLS: ReadonlyArray<readonly [string, string]> = [
  // Is compares two object references and nothing else, so a plain value on either side is error 424 rather
  // than False. Measured for all 144 pairs of the twelve types.
  ['Empty Is Empty', 'ERR 424'],
  ['Empty Is Null', 'ERR 424'],
  ['Empty Is True', 'ERR 424'],
  ['Empty Is CByte(1)', 'ERR 424'],
  ['Empty Is CInt(1)', 'ERR 424'],
  ['Empty Is CLng(1)', 'ERR 424'],
  ['Empty Is CSng(1)', 'ERR 424'],
  ['Empty Is CDbl(1)', 'ERR 424'],
  ['Empty Is CCur(1)', 'ERR 424'],
  ['Empty Is CDate("2026-01-05")', 'ERR 424'],
  ['Empty Is "2"', 'ERR 424'],
  ['"2" Is "2"', 'ERR 424'],
  ['Null Is True', 'ERR 424'],
  ['CDate("2026-01-05") Is CCur(1)', 'ERR 424'],
  ['CByte(200) Is CLng(70000)', 'ERR 424'],

  // The type each bitwise operator keeps, and the wrap of the narrower types.
  ['Empty And Empty', 'Long | 0'],
  ['Empty Eqv Empty', 'Long | -1'],
  ['True And True', 'Boolean | True'],
  ['True And CByte(200)', 'Integer | 200'],
  ['CByte(200) And CByte(200)', 'Byte | 200'],
  ['CByte(200) And CInt(-3)', 'Integer | 200'],
  ['CByte(200) Eqv CByte(200)', 'Byte | 255'],
  ['CByte(200) Imp CByte(200)', 'Byte | 255'],
  ['CLng(70000) And CLng(70000)', 'Long | 70000'],

  // A floating point operand is rounded first, a tie going to the even neighbour.
  ['CSng(1.5) Or CSng(1.5)', 'Long | 2'],
  ['CByte(200) Or CSng(1.5)', 'Long | 202'],
  ['CInt(-3) And CSng(1.5)', 'Long | 0'],

  // A text that is no number is a type mismatch, even beside a Null or an Empty.
  ['"abc" And Empty', 'ERR 13'],
  ['Empty And "abc"', 'ERR 13'],
  ['Null And "abc"', 'ERR 13'],
  ['"abc" Or Null', 'ERR 13'],
  ['Not "abc"', 'ERR 13'],
  ['Empty \\ "abc"', 'ERR 13'],

  // A Null, operator by operator.
  ['Empty And Null', 'Long | 0'],
  ['Null And Empty', 'Long | 0'],
  ['Null And True', 'ERR 94'],
  ['Null Or CByte(200)', 'Byte | 200'],
  ['True Or Null', 'Boolean | True'],
  ['Empty Or Null', 'ERR 94'],
  ['Empty Xor Null', 'ERR 94'],
  ['Empty Eqv Null', 'ERR 94'],
  ['Null Imp Empty', 'ERR 94'],
  ['Null Imp CByte(200)', 'Byte | 200'],
  ['CByte(200) Imp Null', 'Byte | 55'],
  ['True Imp Null', 'ERR 94'],

  // Integer division and the remainder: the type, the rounding and the cut toward zero.
  ['Empty \\ True', 'Integer | 0'],
  ['CByte(200) \\ CByte(200)', 'Byte | 1'],
  ['CByte(200) \\ CInt(-3)', 'Integer | -66'],
  ['CByte(200) \\ CSng(1.5)', 'Long | 100'],
  ['CByte(200) Mod CByte(200)', 'Byte | 0'],
  ['CLng(70000) Mod CSng(1.5)', 'Long | 0'],
  ['Empty \\ Empty', 'ERR 11'],
  ['Empty Mod Empty', 'ERR 11'],

  // Not and unary minus.
  ['Not True', 'Boolean | False'],
  ['Not CByte(200)', 'Byte | 55'],
  ['Not CInt(-3)', 'Integer | 2'],
  ['Not CLng(70000)', 'Long | -70001'],
  ['Not Empty', 'Long | -1'],
  ['Not "2"', 'Long | -3'],
  ['-CByte(200)', 'Integer | -200'],
  ['-CInt(-3)', 'Integer | 3'],
  ['-CSng(1.5)', 'Single | -1.5'],
  ['-CCur(1.2345)', 'Currency | -1.2345'],
  ['-CDate("2026-01-05")', 'Date | 12/23/1773'],
  ['-"2"', 'Double | -2'],

  // A power outside the Double range.
  ['Empty ^ True', 'ERR 5'],
  ['CInt(-3) ^ CSng(1.5)', 'ERR 5'],
  ['CByte(200) ^ CByte(200)', 'ERR 6'],

  // Division: the Single it answers, the Double it answers otherwise, and the two zero errors.
  ['Empty / True', 'Single | -0'],
  ['Empty / CInt(-3)', 'Single | -0'],
  ['Empty / CByte(200)', 'Single | 0'],
  ['Empty / CSng(1.5)', 'Single | 0'],
  ['Empty / CLng(70000)', 'Double | 0'],
  ['CByte(200) / CInt(-3)', 'Double | -66.6666666666667'],
  ['True / CSng(1.5)', 'Single | -0.6666667'],
  ['CSng(1.5) / True', 'Single | -1.5'],
  ['Empty / Empty', 'ERR 6'],
  ['CInt(0) / CInt(0)', 'ERR 6'],
  ['CByte(200) / CByte(0)', 'ERR 11'],

  // The width a +, - and * answer, and the two shapes a Date takes part in.
  ['CLng(5) - CLng(5)', 'Long | 0'],
  ['Empty * CLng(70000)', 'Long | 0'],
  ['CLng(70000) * Empty', 'Long | 0'],
  ['CByte(200) * CByte(200)', 'Long | 40000'],
  ['CLng(70000) * CLng(70000)', 'Double | 4900000000'],
  ['CSng(1.5) * CCur(1.2345)', 'Double | 1.85175'],
  ['CDate("2026-01-05") * CDate("2026-01-05")', 'Double | 2118484729'],
  ['CCur(1.2345) + CDate("2026-01-05")', 'Date | 1/6/2026 5:37:41 AM'],
  ['CDate("2026-01-05") + CCur(1.2345)', 'Date | 1/6/2026 5:37:41 AM'],

  // A negative zero keeps its sign where the type has one.
  ['-CDbl(0)', 'Double | -0'],
  ['-CSng(0)', 'Single | -0'],
  ['-CCur(0)', 'Currency | 0'],
  ['-CInt(0)', 'Integer | 0'],
];

function answer(expression: string, lcid = 1033): unknown {
  const engine = new VbsEngine();
  try {
    engine.executeStatement(
      `SetLocale ${lcid}\nr = TypeName(${expression}) & " | " & CStr(${expression})`
    );
    const failure = engine.error;
    if (failure) return 'ERR ' + String(failure.number);
    return engine.eval('r');
  } finally {
    engine.executeStatement('SetLocale 1033');
  }
}

describe('an operator between two types answers what vbscript.dll answers', () => {
  for (const [expression, expected] of CELLS) {
    it(`${expression} is ${expected}`, () => {
      expect(answer(expression)).toBe(expected);
    });
  }
});
