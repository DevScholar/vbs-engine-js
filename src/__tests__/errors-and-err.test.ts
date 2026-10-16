import { describe, it, expect } from 'vitest';
import { VbsEngine, type VbsHostError } from '../index.ts';

// What Err holds after a statement, measured against a fresh cscript under 1033 with the error probe in the
// probe directory: 82 cells of runtime errors and of the Err object itself. 80 of them agree with the
// engine, and the cells below are those. Windows localizes Err.Description to the language of the machine
// (this one answers in Chinese), so only Err.Number is compared.
//
// Two cells are still open, and both call a member of a host class the runner does not have (a file that is
// not there and a drive that is not there, error 53 and 76 in Windows). A host class is the host's business
// rather than this engine's, so those two are out of this file's scope.
//
// The argument count of a call is checked now, and so is a call on a name nothing holds, which was two more
// of these cells: see builtin-arity.test.ts and the "undefined name" cells below.
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
  ['err.args.unknown', 'x = NoSuchFunction(1)', '13'],
  // A call on a name nothing holds indexes the Empty it stands for, which is a type mismatch, in every shape
  // a call can be written without Option Explicit: measured, the three statement forms and the four
  // expression forms below are all 13, while a bare reference is the Empty itself.
  ['err.args.unknown.statement', 'NoSuchFunction(1)', '13'],
  ['err.args.unknown.statementBare', 'NoSuchFunction 1', '13'],
  ['err.args.unknown.statementCall', 'Call NoSuchFunction(1)', '13'],
  ['err.args.unknown.twoArgs', 'x = NoSuchFunction(1, 2)', '13'],
  ['err.args.unknown.condition', 'If NoSuchFunction(1) Then x = 1', '13'],
  ['err.args.unknown.argument', 'x = CStr(NoSuchFunction(1))', '13'],
  ['err.args.unknown.arithmetic', 'x = NoSuchFunction(1) + 1', '13'],
  ['err.args.unknown.bareName', 'x = NoSuchFunction', '0'],
  ['err.for.each', 'x = 0 : For Each y In 5 : x = x + 1 : Next', '451'],
  ['err.null.method', 'x = Null.Foo', '424'],
  ['err.null.property', 'x = Len(Null)', '0'],
  ['err.empty.length', 'x = Len(Empty)', '0'],
  ['err.undefined.variable', 'x = notDeclaredAnywhere', '0'],
  ['err.array.compare', 'a = Array(1) : x = (a = a)', '13'],
  ['err.cpp.array', 'a = Array(1, 2) : b = Array(3) : x = a + b', '13'],

  // An object needs Set, and Set takes nothing else. Measured on a fresh cscript, with the object taken
  // from a class the script declares itself, so that nothing here depends on a host class.
  ['assign.object', 'Class Holder : End Class : Set o = New Holder : x = o', '438'],
  ['assign.nothing', 'x = Nothing', '91'],
  ['assign.array', 'x = Array(1, 2)', '0'],
  ['assign.object.with.set', 'Class Holder : End Class : Set o = New Holder', '0'],
  ['assign.nothing.with.set', 'Set o = Nothing', '0'],
  [
    'argument.object',
    'Class Holder : End Class : Function Take(value) : Take = 1 : End Function : Set o = New Holder : x = Take(o)',
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

// With Option Explicit the name itself is the error, so every one of those shapes is 500 instead. Option
// Explicit has to open the script, so these cells are run on their own rather than through the helper above.
const EXPLICIT_CELLS: ReadonlyArray<readonly [string, string]> = [
  ['NoSuchFunction(1)', '500'],
  ['NoSuchFunction 1', '500'],
  ['Call NoSuchFunction(1)', '500'],
  ['x = NoSuchFunction(1)', '500'],
  ['x = NoSuchFunction', '500'],
  ['NoSuchFunction', '500'],
  ['x = NoSuchFunction(1, 2)', '500'],
  ['x = 1 + 1', '0'],
];

describe('an undefined name with Option Explicit', () => {
  for (const [statement, expected] of EXPLICIT_CELLS) {
    it(`${statement} leaves ${expected}`, () => {
      const engine = new VbsEngine();
      engine.executeStatement(
        `Option Explicit\nOn Error Resume Next\nDim x, value\nErr.Clear\n${statement}\nvalue = CStr(Err.Number)`
      );
      const failure = engine.error;
      const number =
        failure && failure.number !== 0
          ? failure.number
          : Number(engine._getVariable('value').value);
      expect(String(number)).toBe(expected);
    });
  }
});

// ---------------------------------------------------------------------------
// The scope of On Error Resume Next
// ---------------------------------------------------------------------------
// Measured on a fresh cscript. The handler is the procedure's own: it is not in force in a procedure that
// procedure calls, and the handler the caller set is in force again once the call returns. So a divide by zero
// inside a Sub answers 11 to a caller that set the handler, the Sub's own statements after it do not run, and
// the caller carries on; and a handler a procedure set for itself is gone once it returns. The engine used to
// keep one flag for the whole run, so a single `On Error Resume Next` written inside a helper swallowed every
// later failure in the script, which is how the example page's clock stopped reporting its own errors.
describe('the scope of On Error Resume Next', () => {
  it('is not in force in a procedure the caller calls, and comes back when the call returns', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`On Error Resume Next
Dim result
Sub Inner()
  result = "before"
  result = 1 / 0
  result = "after"
End Sub
Inner
after = result & "/" & CStr(Err.Number)`);
    expect(engine.error).toBeNull();
    // The Sub stopped where it failed, and the caller's own handler took the error and carried on.
    expect(engine._getVariable('after').value).toBe('before/11');
  });

  it('does not outlive the procedure that set it', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`Dim answer
Function Inside()
  On Error Resume Next
  Inside = 1 / 0
  If Err.Number <> 0 Then Inside = "handled"
End Function
answer = Inside()
x = 1 / 0`);
    // The divide at the end is the script's own and nobody is handling it, which is what the documented rule
    // says: the handler the function set went with the function.
    expect(engine.error).not.toBeNull();
    expect(engine.error!.number).toBe(11);
    expect(engine._getVariable('answer').value).toBe('handled');
  });
});

// ---------------------------------------------------------------------------
// What On Error Resume Next does to Err
// ---------------------------------------------------------------------------
// Measured on a fresh cscript, and written in the Err object's own documentation: "The Err object's properties
// are reset to zero or zero-length strings ("") after an On Error Resume Next statement." The statement resets
// them itself, so a raise that a handler already dealt with does not stay in Err because a later handler
// happens to be written.
describe('what On Error Resume Next does to Err', () => {
  it('resets Err itself, the way Err.Clear and On Error GoTo 0 do', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`On Error Resume Next
Err.Raise 4321, "a source", "a description"
raised = CStr(Err.Number)
raisedSource = Err.Source
On Error Resume Next
after = CStr(Err.Number)
afterSource = Err.Source
afterDescription = Err.Description`);
    expect(engine._getVariable('raised').value).toBe('4321');
    expect(engine._getVariable('raisedSource').value).toBe('a source');
    expect(engine._getVariable('after').value).toBe('0');
    expect(engine._getVariable('afterSource').value).toBe('');
    expect(engine._getVariable('afterDescription').value).toBe('');
  });
});

// ---------------------------------------------------------------------------
// Err.Raise does not only fill Err, it raises
// ---------------------------------------------------------------------------
// Measured on a fresh cscript: a script that writes a line, then `Err.Raise 4321, "a source", "a description"`,
// then another line prints the first line and then `(4, 1) a source: a description`, and the line after it never
// runs. So the statement raises a run-time error like any other, and a handler is what keeps a script going.
describe('Err.Raise', () => {
  it('stops a script that has no handler', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`Dim after
after = 0
before = 1
Err.Raise 4321, "a source", "a description"
after = 1`);
    expect(engine._getVariable('before').value).toBe(1);
    expect(engine._getVariable('after').value).toBe(0);
    expect(engine.error!.number).toBe(4321);
    expect(engine.error!.description).toBe('a description');
    expect(engine.error!.source).toBe('a source');
  });

  it('reaches a JavaScript caller as the HRESULT for the number it raised', () => {
    const engine = new VbsEngine();
    let caught: VbsHostError | null = null;
    try {
      engine.executeStatementThrows('Err.Raise 4321, "a source", "a description"');
    } catch (e) {
      caught = e as VbsHostError;
    }
    expect(caught).not.toBeNull();
    // 4321 is 0x10E1, and the VBScript facility puts it at 0x800A10E1. Measured the same way out of
    // Internet Explorer, whose caught error for Err.Raise 4321 is -2146823967.
    expect(caught!.number).toBe(-2146823967);
    expect(caught!.description).toBe('a description');
  });
});
