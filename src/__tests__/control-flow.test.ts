import { describe, it, expect } from 'vitest';
import { VbsEngine } from '../index.ts';

// Control flow, measured against a fresh cscript under 1033 with the control probe in the probe directory:
// 55 blocks, each one written out at script level with its answer left in result. Every one of them agrees
// with the engine. The corners the measurement settled:
//
//   * Select Case compares text case sensitively and raises a type mismatch for a number written out beside
//     a text, which is the = operator's own reading (see comparisons-and-conditions.test.ts). A Null value
//     and a Case Null both fall through to Case Else, and the first matching case wins.
//   * A Do loop reads its condition at the end when the condition sits after Loop, so the body runs once
//     even when the condition is false. Exit Do leaves only the loop it sits in, and a Null or an Empty
//     condition is false rather than an error, while a text that is no number is error 13.
//   * For...Next keeps the value that stopped it in its counter, where For Each empties its own loop
//     variable after the loop. A bound or a step that is no number is the loop's own error 92.
const CELLS: ReadonlyArray<readonly [string, string, string]> = [
  [
    'sc.simple',
    'Select Case 3\nCase 1\nresult = "one"\nCase 3\nresult = "three"\nEnd Select',
    'three',
  ],
  [
    'sc.else',
    'Select Case 9\nCase 1\nresult = "one"\nCase Else\nresult = "else"\nEnd Select',
    'else',
  ],
  ['sc.nomatch', 'result = "start"\nSelect Case 9\nCase 1\nresult = "one"\nEnd Select', 'start'],
  [
    'sc.list',
    'Select Case 3\nCase 1, 3, 5\nresult = "oddSmall"\nCase Else\nresult = "other"\nEnd Select',
    'oddSmall',
  ],
  [
    'sc.listFirstWins',
    'Select Case 2\nCase 1, 2\nresult = "first"\nCase 2, 3\nresult = "second"\nEnd Select',
    'first',
  ],
  ['sc.string', 'Select Case "b"\nCase "a"\nresult = "a"\nCase "b"\nresult = "b"\nEnd Select', 'b'],
  [
    'sc.stringCase',
    'Select Case "b"\nCase "B"\nresult = "upper"\nCase Else\nresult = "else"\nEnd Select',
    'else',
  ],
  [
    'sc.emptyZero',
    'Select Case Empty\nCase 0\nresult = "zero"\nCase Else\nresult = "else"\nEnd Select',
    'zero',
  ],
  [
    'sc.emptyCase',
    'Select Case 0\nCase Empty\nresult = "empty"\nCase Else\nresult = "else"\nEnd Select',
    'empty',
  ],
  [
    'sc.nullValue',
    'Select Case Null\nCase 1\nresult = "one"\nCase Else\nresult = "else"\nEnd Select',
    'else',
  ],
  [
    'sc.nullCase',
    'Select Case 1\nCase Null\nresult = "null"\nCase Else\nresult = "else"\nEnd Select',
    'else',
  ],
  [
    'sc.trueMinusOne',
    'Select Case -1\nCase True\nresult = "true"\nCase Else\nresult = "else"\nEnd Select',
    'true',
  ],
  [
    'sc.boolean',
    'Select Case True\nCase True\nresult = "true"\nCase Else\nresult = "else"\nEnd Select',
    'true',
  ],
  [
    'sc.date',
    'Select Case CDate("2026-01-05")\nCase CDate("2026-01-05")\nresult = "same"\nCase Else\nresult = "else"\nEnd Select',
    'same',
  ],
  [
    'sc.nested',
    'Select Case 1\nCase 1\nSelect Case 2\nCase 2\nresult = "inner"\nEnd Select\nEnd Select',
    'inner',
  ],
  [
    'sc.expressionOnce',
    'n = 0\nresult = ""\nSelect Case NextCount(n)\nCase 1\nresult = CStr(n)\nEnd Select',
    '1',
  ],
  [
    'sc.typeMismatch',
    'Select Case "abc"\nCase 1\nresult = "one"\nCase Else\nresult = "else"\nEnd Select',
    'ERR 13',
  ],
  ['sc.singleLine', 'Select Case 2 : Case 2 : result = "two" : End Select', 'two'],
  ['sc.insideFor', 'For i = 1 To 1\nSelect Case 1\nCase 1\nresult = "in"\nEnd Select\nNext', 'in'],
  ['do.while', 'n = 0\nDo While n < 3\nn = n + 1\nLoop\nresult = CStr(n)', '3'],
  ['do.until', 'n = 0\nDo Until n = 3\nn = n + 1\nLoop\nresult = CStr(n)', '3'],
  ['do.whileFalse', 'n = 5\nDo While n < 3\nn = n + 1\nLoop\nresult = CStr(n)', '5'],
  ['do.loopWhile', 'n = 0\nDo\nn = n + 1\nLoop While n < 3\nresult = CStr(n)', '3'],
  ['do.loopWhileOnce', 'n = 9\nDo\nn = n + 1\nLoop While n < 3\nresult = CStr(n)', '10'],
  ['do.loopUntil', 'n = 0\nDo\nn = n + 1\nLoop Until n = 3\nresult = CStr(n)', '3'],
  ['do.exitDo', 'n = 0\nDo\nn = n + 1\nIf n = 2 Then Exit Do\nLoop\nresult = CStr(n)', '2'],
  [
    'do.exitDoInner',
    'n = 0\nDo While n < 5\nn = n + 1\nDo\nExit Do\nLoop\nLoop\nresult = CStr(n)',
    '5',
  ],
  [
    'do.exitDoFromFor',
    'n = 0\nDo While n < 5\nFor i = 1 To 3\nn = n + 1\nIf n = 2 Then Exit Do\nNext\nLoop\nresult = CStr(n)',
    '2',
  ],
  ['do.emptyCondition', 'n = 0\nDo While Empty\nn = n + 1\nLoop\nresult = CStr(n)', '0'],
  ['do.nullCondition', 'n = 0\nDo While Null\nn = n + 1\nLoop\nresult = CStr(n)', '0'],
  ['do.zeroCondition', 'n = 0\nDo While 0\nn = n + 1\nLoop\nresult = CStr(n)', '0'],
  [
    'do.stringCondition',
    'n = 0\nDo While "x"\nn = n + 1\nIf n = 2 Then Exit Do\nLoop\nresult = CStr(n)',
    'ERR 13',
  ],
  ['do.noCondition', 'n = 0\nDo\nn = n + 1\nIf n = 4 Then Exit Do\nLoop\nresult = CStr(n)', '4'],
  ['for.simple', 'n = 0\nFor i = 1 To 3\nn = n + i\nNext\nresult = CStr(n) & "/" & CStr(i)', '6/4'],
  [
    'for.noIterations',
    'For i = 3 To 1\nresult = "ran"\nNext\nresult = "skipped/" & CStr(i)',
    'skipped/3',
  ],
  ['for.negativeStep', 'n = 0\nFor i = 3 To 1 Step -1\nn = n + i\nNext\nresult = CStr(n)', '6'],
  [
    'for.stepTwo',
    'n = 0\nFor i = 1 To 5 Step 2\nn = n + 1\nNext\nresult = CStr(n) & "/" & CStr(i)',
    '3/7',
  ],
  [
    'for.exitFor',
    'n = 0\nFor i = 1 To 5\nn = n + 1\nIf i = 2 Then Exit For\nNext\nresult = CStr(n)',
    '2',
  ],
  [
    'for.modifyCounter',
    'n = 0\nFor i = 1 To 5\ni = i + 1\nn = n + 1\nNext\nresult = CStr(n) & "/" & CStr(i)',
    '3/7',
  ],
  ['for.decimalStep', 'n = 0\nFor i = 1 To 2 Step 0.5\nn = n + 1\nNext\nresult = CStr(n)', '3'],
  ['for.nullBound', 'For i = 1 To Null\nresult = "ran"\nNext', 'ERR 92'],
  ['for.emptyBound', 'n = 0\nFor i = 1 To Empty\nn = n + 1\nNext\nresult = CStr(n)', '0'],
  ['for.stringBound', 'n = 0\nFor i = 1 To "3"\nn = n + 1\nNext\nresult = CStr(n)', '3'],
  ['for.stringCounter', 'n = 0\nFor i = "a" To "c"\nn = n + 1\nNext\nresult = CStr(n)', 'ERR 92'],
  ['fe.array', 'n = 0\nFor Each i In Array(1, 2, 3)\nn = n + i\nNext\nresult = CStr(n)', '6'],
  ['fe.emptyArray', 'n = 0\nFor Each i In Array()\nn = n + 1\nNext\nresult = CStr(n)', '0'],
  [
    'fe.exitFor',
    'n = 0\nFor Each i In Array(1, 2, 3)\nn = n + 1\nIf n = 2 Then Exit For\nNext\nresult = CStr(n)',
    '2',
  ],
  [
    'fe.types',
    'result = ""\nFor Each i In Array(1, 2.5, "x")\nresult = result & TypeName(i) & ","\nNext',
    'Integer,Double,String,',
  ],
  [
    'fe.modifyArray',
    'a = Array(1, 2, 3)\nresult = ""\nFor Each i In a\na(0) = 9\nresult = result & CStr(i)\nNext',
    '123',
  ],
  ['fe.number', 'For Each i In 5\nresult = "ran"\nNext', 'ERR 451'],
  ['fe.afterLoop', 'For Each i In Array(7)\nNext\nresult = CStr(i)', ''],
  ['while.simple', 'n = 0\nWhile n < 3\nn = n + 1\nWend\nresult = CStr(n)', '3'],
  ['while.false', 'n = 5\nWhile n < 3\nn = n + 1\nWend\nresult = CStr(n)', '5'],
  ['while.null', 'n = 0\nWhile Null\nn = n + 1\nWend\nresult = CStr(n)', '0'],
  ['while.empty', 'n = 0\nWhile Empty\nn = n + 1\nWend\nresult = CStr(n)', '0'],
];

/** Runs one cell in its own engine and answers its result or the error it raised. */
function run(code: string): string {
  const engine = new VbsEngine();
  engine.executeStatement(
    `On Error Resume Next
SetLocale 1033
Function NextCount(counter)
  counter = counter + 1
  NextCount = counter
End Function
Dim result, n, i, j, a, c
Err.Clear
${code}
number = CStr(Err.Number)`
  );
  const failure = engine.error;
  const number =
    failure && failure.number !== 0 ? failure.number : Number(engine._getVariable('number').value);
  if (number !== 0) return `ERR ${number}`;
  const answer = engine._getVariable('result');
  return answer === undefined || answer.value === undefined ? '' : String(answer.value);
}

describe('the control flow statements behave like vbscript.dll', () => {
  for (const [label, code, expected] of CELLS) {
    it(`${label} is ${expected}`, () => {
      expect(run(code)).toBe(expected);
    });
  }
});
