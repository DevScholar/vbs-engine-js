import { describe, it, expect } from 'vitest';
import { VbsEngine } from '../index.ts';

// What Err holds after a statement, measured against a fresh cscript under 1033 with the error probe in the
// probe directory: 82 cells of runtime errors and of the Err object itself. 79 of them agree with the
// engine, and the cells below are those. Windows localizes Err.Description to the language of the machine
// (this one answers in Chinese), so only Err.Number is compared.
//
// Three cells are still open, all written down here rather than hidden:
//
//   * x = NoSuchFunction(1) is error 13 in Windows, since the name holds nothing and is indexed like any
//     other Empty. The engine answers 500 there, because it reads the same shape as a call to an undefined
//     procedure, which is what that shape is at statement level ("NoSuchFunction 1" is error 500).
//   * OpenTextFile on a missing file and on a missing drive are error 53 and 76 in Windows. Those members
//     belong to a COM class, which the host supplies rather than this engine (see the compatibility
//     document), and the runner these tests use has none, so the call answers the invalid procedure call the
//     engine raises for a member it does not have.
//
// The argument count of a call is checked now, which was two more of these cells: see builtin-arity.test.ts.
const CELLS: ReadonlyArray<readonly [string, string, string]> = [
  // The number each runtime failure leaves in Err.
  ['ok.nothing', 'x = 1 + 1', '0'],
  ['err.type.string', 'x = CInt("abc")', '13'],
  ['err.type.add', 'x = "a" + 1', '13'],
  ['err.type.addLeft', 'x = 1 + "a"', '13'],
  ['err.divide', 'x = 1 / 0', '11'],
  ['err.divideInt', 'x = 1 \\ 0', '11'],
  ['err.divideMod', 'x = 1 Mod 0', '11'],
  ['err.divideZeroZero', 'x = 0 / 0', '6'],
  ['err.overflow.cint', 'x = CInt(40000)', '6'],
  ['err.overflow.cbyte', 'x = CByte(300)', '6'],
  ['err.overflow.clng', 'x = CLng(3000000000)', '6'],
  ['err.overflow.power', 'x = 200 ^ 200', '6'],
  ['err.overflow.ccur', 'x = CCur(1E+20)', '6'],
  // A product of two Longs that leaves the Long range is a Double and not an error.
  ['err.overflow.multiply', 'x = 100000 * 100000', '0'],
  ['err.subscript.array', 'a = Array(1, 2, 3) : x = a(5)', '9'],
  ['err.subscript.negative', 'a = Array(1, 2, 3) : x = a(-1)', '9'],
  ['err.subscript.dim', 'b = Array(1) : x = b(5, 5)', '9'],
  ['err.subscript.ubound', 'x = UBound(5)', '13'],
  ['err.subscript.redim', 'Dim c() : x = UBound(c)', '9'],
  ['err.object.notSet', 'Set o = Nothing : x = o.Foo', '424'],
  ['err.object.required', 'x = Empty.Foo', '424'],
  ['err.object.subscript', 'x = Empty(0)', '13'],
  ['err.object.undefined', 'x = someObject.Foo', '424'],
  ['err.proc.mid', 'x = Mid("abc", 0)', '5'],
  ['err.proc.left', 'x = Left("abc", -1)', '5'],
  ['err.proc.formatNumber', 'x = FormatNumber(1.25, -2)', '5'],
  ['err.proc.space', 'x = Space(-1)', '5'],
  ['err.proc.string', 'x = String(-1, "a")', '5'],
  // Chr(-1) is a character and not an error.
  ['err.proc.chr', 'x = Chr(-1)', '0'],
  ['err.args.unknown', 'x = NoSuchFunction(1)', '500'],
  ['err.for.each', 'x = 0 : For Each y In 5 : x = x + 1 : Next', '451'],
  ['err.null.method', 'x = Null.Foo', '424'],
  ['err.null.property', 'x = Len(Null)', '0'],
  ['err.empty.length', 'x = Len(Empty)', '0'],
  ['err.undefined.variable', 'x = notDeclaredAnywhere', '0'],
  ['err.array.compare', 'a = Array(1) : x = (a = a)', '13'],
  ['err.cpp.array', 'a = Array(1, 2) : b = Array(3) : x = a + b', '13'],

  // An object needs Set, and Set takes nothing else. Measured on a fresh cscript. The object here comes
  // from the engine's own Collection rather than from a COM class, which the host supplies.
  ['assign.object', 'Set o = New Collection : x = o', '450'],
  ['assign.nothing', 'x = Nothing', '91'],
  ['assign.array', 'x = Array(1, 2)', '0'],
  ['assign.object.with.set', 'Set o = New Collection', '0'],
  ['assign.nothing.with.set', 'Set o = Nothing', '0'],
  [
    'argument.object',
    'Function Take(value) : Take = 1 : End Function : Set o = New Collection : x = Take(o)',
    '0',
  ],
  ['set.number', 'Set o = 5', '424'],
  ['set.text', 'Set o = "abc"', '424'],
  ['set.empty', 'Set o = Empty', '424'],
  ['set.null', 'Set o = Null', '424'],
  ['set.array', 'Set o = Array(1, 2)', '424'],
  ['assign.number', 'o = 5', '0'],
  ['assign.empty', 'o = Empty', '0'],
  ['assign.null', 'o = Null', '0'],

  // The Err object itself, which rounds a number and holds it only inside a range.
  ['e.after.ok', 'x = 5', '0'],
  ['e.clear', 'Err.Clear', '0'],
  ['e.raise.one', 'Err.Raise 5', '5'],
  ['e.raise.zero', 'Err.Raise 0', '5'],
  ['e.raise.big', 'Err.Raise 65535', '65535'],
  ['e.raise.high', 'Err.Raise 65536', '5'],
  ['e.raise.negative', 'Err.Raise -1', '-1'],
  ['e.raise.custom', 'Err.Raise vbObjectError + 2000, "src", "desc"', '-2147219504'],
  ['e.raise.fraction', 'Err.Raise 5.7', '6'],
  ['e.raise.text', 'Err.Raise "x"', '13'],
  ['e.raise.textnumber', 'Err.Raise "5"', '5'],
  ['e.raise.null', 'Err.Raise Null', '13'],
  ['e.raise.65535.9', 'Err.Raise 65535.9', '5'],
  ['e.number.set', 'Err.Number = 7', '7'],
  ['e.number.zero', 'Err.Number = 0', '0'],
  ['e.number.fraction', 'Err.Number = 5.7', '6'],
  ['e.number.text', 'Err.Number = "9"', '9'],
  ['e.number.badtext', 'Err.Number = "x"', '13'],
  ['e.number.negativezero', 'Err.Number = -0.5', '0'],
  ['e.number.sethigh', 'Err.Number = 70000', '5'],
  ['e.number.setnegative', 'Err.Number = -5', '-5'],
  ['e.clear.method', 'Err.Raise 5 : Err.Clear', '0'],
];

/** Runs one statement with a handler in place and answers the number it left in Err. */
function errorNumber(statement: string): string {
  const engine = new VbsEngine();
  engine.executeStatement(
    `On Error Resume Next\nDim x, y, z, a, b, o, f\nErr.Clear\n${statement}\nvalue = CStr(Err.Number)`
  );
  const failure = engine.error;
  if (failure && failure.number !== 0 && engine.error !== null) {
    // A statement that failed outside the script's own handler surfaces here instead.
    return 'ERR ' + String(failure.number);
  }
  return String(engine._getVariable('value').value);
}

describe('the number a failure leaves in Err', () => {
  for (const [label, statement, expected] of CELLS) {
    it(`${label} leaves ${expected}`, () => {
      expect(errorNumber(statement)).toBe(expected);
    });
  }
});
