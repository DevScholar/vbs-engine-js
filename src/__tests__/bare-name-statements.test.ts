import { describe, it, expect } from 'vitest';
import { VbsEngine } from '../index.ts';

// A statement that is just a name, measured against a fresh cscript under 1033 with the bare name probes in
// the probe directory. What the measurement settled:
//
//   * A whole statement naming a value is a Sub call on it with no arguments, so a name that holds a number,
//     a text, an Empty, a Null or an array is error 13, and so is a name that holds nothing at all.
//   * A name that holds an object is 438 instead, since the object is asked to be a Sub: `heldObject` where
//     the class has no default member, and `WScript` and `WScript.Arguments` with the host in place. A
//     Nothing is the exception and stays 13.
//   * A name that is a procedure is called: `NoArgs` runs the Sub, `NoArgFunction` runs the Function,
//     `heldObject.DoIt`, `heldObject.GetIt` and `heldObject.Prop` run the method or the property, and a
//     procedure that needs an argument answers 450.
//   * A member that holds a field rather than a procedure is error 5 with no arguments and 438 with them,
//     which `heldObject.Value` and `heldObject.Value(1)` show, and a Property Get is the exception again:
//     `heldObject.Prop` is 0 because the getter runs and its value is dropped.
//   * An array is never indexed by a statement: `arr(9)` alone is 13 where `x = arr(9)` is 9. The same
//     holds for a member array, `holder.Items(0)` alone being 13.
//   * A literal as a statement is a compile error in Windows, so no cell here writes one.
//   * A Sub has no value: `result = NoArgs` is 13 for a Sub of any arity, while a Function is called,
//     `result = NoArgFunction` being 0 and `result = OneArgFunction` the 450 it raises with no arguments.
//
// Eval is handed an expression rather than a statement, so it reads the same text the other way:
// Eval("number") is 5 and Eval("NoSuchFunction") is 0, where those names written as statements are both 13.
const TARGET_CELLS: ReadonlyArray<readonly [string, string, string]> = [
  ['bare.unknown', 'NoSuchFunction', 'ERR 13'],
  ['bare.number', 'number', 'ERR 13'],
  ['bare.emptyVariable', 'blank', 'ERR 13'],
  ['bare.null', 'nothingValue', 'ERR 13'],
  ['bare.text', 'text', 'ERR 13'],
  ['bare.array', 'arrayValue', 'ERR 13'],
  ['bare.nothing', 'Set heldObject = Nothing\nheldObject', 'ERR 13'],
  ['bare.object', 'Set heldObject = New Holder\nheldObject', 'ERR 438'],
  ['member.sub', 'Set heldObject = New Holder\nheldObject.DoIt', 'ok'],
  ['member.function', 'Set heldObject = New Holder\nheldObject.GetIt', 'ok'],
  ['member.property', 'Set heldObject = New Holder\nheldObject.Prop', 'ok'],
  ['member.field', 'Set heldObject = New Holder\nheldObject.Value', 'ERR 5'],
  ['member.methodTooFewArgs', 'Set heldObject = New Holder\nheldObject.WithArg', 'ERR 450'],
  ['sub.noArgs', 'NoArgs', 'ok'],
  ['sub.oneArgNoArgs', 'OneArg', 'ERR 450'],
  ['function.noArgs', 'NoArgFunction', 'ok'],
  ['builtin.oneArg', 'Len', 'ERR 450'],
  ['builtin.noArgs', 'Now', 'ok'],
];

const MEMBER_CELLS: ReadonlyArray<readonly [string, string, string]> = [
  ['call.objectWithArg', 'heldObject(1)', 'ERR 438'],
  ['call.fieldWithArg', 'heldObject.Value(1)', 'ERR 438'],
  ['bare.field', 'heldObject.Value', 'ERR 5'],
  ['call.methodTooManyArgs', 'Call heldObject.DoIt(1, 2)', 'ERR 450'],
  ['bare.withField', 'With heldObject\n.Value\nEnd With', 'ERR 5'],
  ['bare.withMethod', 'With heldObject\n.DoIt\nEnd With', 'ok'],
  ['bare.constant', 'constantValue', 'ERR 13'],
];

const PROCEDURE_CELLS: ReadonlyArray<readonly [string, string, string]> = [
  ['value.subNoArgs', 'result = NoArgs', 'ERR 13'],
  ['value.subOneArg', 'result = OneArg', 'ERR 13'],
  ['value.functionNoArgs', 'result = NoArgFunction', 'ok'],
  ['value.functionOneArg', 'result = OneArgFunction', 'ERR 450'],
  ['eval.subNoArgs', 'result = Eval("NoArgs")', 'ERR 13'],
  ['eval.functionNoArgs', 'result = Eval("NoArgFunction")', 'ok'],
  ['typename.functionName', 'result = TypeName(NoArgFunction)', 'ok'],
];

/** Runs one cell on its own and answers "ok" or the error it left in Err. */
function run(setup: string, code: string): string {
  const engine = new VbsEngine();
  engine.executeStatement(`SetLocale 1033
On Error Resume Next
${setup}`);
  engine.executeStatement(`Err.Clear
${code}
number = CStr(Err.Number)`);
  const failure = engine.error;
  const number =
    failure && failure.number !== 0 ? failure.number : Number(engine._getVariable('number').value);
  return number !== 0 ? `ERR ${number}` : 'ok';
}

describe('a statement that is just a name is a call on it', () => {
  for (const [label, code, expected] of TARGET_CELLS) {
    it(`${label} is ${expected}`, () => {
      expect(run(TARGET_SETUP, code)).toBe(expected);
    });
  }
});

describe('a member named by a whole statement', () => {
  for (const [label, code, expected] of MEMBER_CELLS) {
    it(`${label} is ${expected}`, () => {
      expect(run(MEMBER_SETUP, code)).toBe(expected);
    });
  }
});

describe('a procedure name where a value is expected', () => {
  for (const [label, code, expected] of PROCEDURE_CELLS) {
    it(`${label} is ${expected}`, () => {
      expect(run(PROCEDURE_SETUP, code)).toBe(expected);
    });
  }
});

const TARGET_SETUP = `Class Holder
  Public Value
  Public Sub DoIt()
    doItCount = doItCount + 1
  End Sub
  Public Function GetIt()
    getItCount = getItCount + 1
    GetIt = 1
  End Function
  Public Property Get Prop
    propCount = propCount + 1
    Prop = 5
  End Property
  Public Function WithArg(one)
    WithArg = one
  End Function
End Class
Dim result, number, text, blank, nothingValue, arrayValue, heldObject
Dim doItCount, getItCount, propCount, subCount, funcCount
blank = Empty
nothingValue = Null
number = 5
text = "abc"
arrayValue = Array(1, 2)
Set heldObject = New Holder
Sub NoArgs()
  subCount = subCount + 1
End Sub
Sub OneArg(one)
  subCount = subCount + 1
End Sub
Function NoArgFunction()
  funcCount = funcCount + 1
  NoArgFunction = 7
End Function
`;
const MEMBER_SETUP = `Class Holder
  Public Value
  Public Sub DoIt()
  End Sub
  Public Function WithArg(one)
    WithArg = one
  End Function
End Class
Dim result, heldObject, constantValue
Set heldObject = New Holder
constantValue = vbCrLf
`;
const PROCEDURE_SETUP = `Dim result
Sub NoArgs()
End Sub
Sub OneArg(one)
End Sub
Function NoArgFunction()
  NoArgFunction = 7
End Function
Function OneArgFunction(one)
  OneArgFunction = one
End Function
`;
