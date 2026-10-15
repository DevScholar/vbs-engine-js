import { describe, it, expect } from 'vitest';
import { VbsEngine } from '../index.ts';

// How many arguments each built-in takes, measured against a fresh cscript under 1033 with the arity probe
// in the probe directory: every function called with 0, 1, 2 ... arguments, 326 cells over 89 functions. A
// count out of range is error 450, "Wrong number of arguments or invalid property assignment", and anything
// else means the count was accepted. The table itself is src/runtime/builtin-arity.ts.
//
// The measurement settled three things the engine used to get wrong:
//
//   * A count that is too high is refused as well as one that is too low. Mid("abc") and
//     Mid("abc", 1, 1, 1) are both error 450, where the engine ignored the extra argument.
//   * A declared procedure is checked the same way, from its own parameter list: Function One(value) called
//     as One() or as One(1, 2) is error 450, and a Sub with no parameters called with one is 450 too.
//     VBScript has no Optional and no ParamArray word, so every declared parameter is required; the engine
//     accepts both as extensions and counts them the way they read.
//   * Array is the one function with no upper bound: Array() and Array(1, 2, 3, 4) are both accepted.
//
// Three cells of the probe are not about the count at all and are left out here. GetRef("Foo"),
// LoadPicture("x.ico") and CreateObject("Scripting.Dictionary", "localhost") are all functions whose work is
// the host's or is not implemented, and each is assigned to a variable without Set in the probe: Windows
// raises its own failure (5, 432 and 429) before the assignment rule can apply, while the engine hands back
// an object and answers 450 for the assignment. The counts themselves are accepted on both sides.
const CELLS: ReadonlyArray<readonly [string, string]> = [
  ['Mid("abc", 1)', 'ok'],
  ['Mid("abc", 1, 1)', 'ok'],
  ['Mid("abc")', '450'],
  ['Mid("abc", 1, 1, 1)', '450'],
  ['Left("abc", 1)', 'ok'],
  ['Left("abc")', '450'],
  ['Left("abc", 1, 1)', '450'],
  ['Abs(1)', 'ok'],
  ['Abs()', '450'],
  ['Abs(1, 1)', '450'],
  ['DateDiff("d", "2026-01-05", "2026-01-06")', 'ok'],
  ['DateDiff("d", "2026-01-05", "2026-01-06", 1, 1)', 'ok'],
  ['DateDiff("d", "2026-01-05")', '450'],
  ['DateDiff("d", "2026-01-05", "2026-01-06", 1, 1, 1)', '450'],
  ['FormatNumber(1)', 'ok'],
  ['FormatNumber(1, 2, -1, -1, -1)', 'ok'],
  ['FormatNumber()', '450'],
  ['FormatNumber(1, 2, -1, -1, -1, 1)', '450'],
  ['Replace("abc", "a", "b")', 'ok'],
  ['Replace("abc", "a", "b", 1, 1, 0)', 'ok'],
  ['Replace("abc", "a")', '450'],
  ['InStr("abc", "b")', 'ok'],
  ['InStr(1, "abc", "b", 0)', 'ok'],
  ['InStr("abc")', '450'],
  ['InStr(1, "abc", "b", 0, 0)', '450'],
  ['Split("a b", " ")', 'ok'],
  ['Split("a b", " ", 1, 0)', 'ok'],
  ['Split()', '450'],
  ['Join(Array("a"), "-")', 'ok'],
  ['Join()', '450'],
  ['Round(1.5)', 'ok'],
  ['Round(1.5, 1)', 'ok'],
  ['Round()', '450'],
  ['Round(1.5, 1, 1)', '450'],
  ['Array()', 'ok'],
  ['Array(1, 2, 3, 4)', 'ok'],
  ['Now()', 'ok'],
  ['Now(1)', 'ok'],
  ['Now(1, 2)', '450'],
  ['Escape("a b")', 'ok'],
  ['Escape()', '450'],
  ['Escape("a b", "a b")', '450'],
  ['Unescape("a%20b")', 'ok'],
  ['Unescape()', '450'],
  ['Unescape("a%20b", "a%20b")', '450'],
];

/** Runs one call with a handler in place and answers "450" or "ok". */
function callResult(expression: string): string {
  const engine = new VbsEngine();
  engine.executeStatement(
    `On Error Resume Next\nDim result\nErr.Clear\nresult = ${expression}\nvalue = CStr(Err.Number)`
  );
  const failure = engine.error;
  const number =
    failure && failure.number !== 0 ? failure.number : Number(engine._getVariable('value').value);
  return number === 450 ? '450' : 'ok';
}

describe('the number of arguments a built-in is handed', () => {
  for (const [expression, expected] of CELLS) {
    it(`${expression} is ${expected}`, () => {
      expect(callResult(expression)).toBe(expected);
    });
  }

  // A declared procedure carries its own counts, which is where the engine used to check nothing at all.
  const declared: ReadonlyArray<readonly [string, string, string]> = [
    ['One()', 'Function One(value)\nOne = 1\nEnd Function', '450'],
    ['One(1)', 'Function One(value)\nOne = 1\nEnd Function', 'ok'],
    ['One(1, 2)', 'Function One(value)\nOne = 1\nEnd Function', '450'],
    ['DoNone()', 'Sub DoNone()\nEnd Sub', 'ok'],
    ['DoNone(1)', 'Sub DoNone()\nEnd Sub', '450'],
    [
      'WithOptional(1)',
      'Function WithOptional(first, Optional second)\nWithOptional = 1\nEnd Function',
      'ok',
    ],
    [
      'WithOptional()',
      'Function WithOptional(first, Optional second)\nWithOptional = 1\nEnd Function',
      '450',
    ],
    [
      'WithOptional(1, 2, 3)',
      'Function WithOptional(first, Optional second)\nWithOptional = 1\nEnd Function',
      '450',
    ],
    ['WithArray()', 'Function WithArray(ParamArray items)\nWithArray = 1\nEnd Function', 'ok'],
    [
      'WithArray(1, 2, 3)',
      'Function WithArray(ParamArray items)\nWithArray = 1\nEnd Function',
      'ok',
    ],
  ];

  for (const [expression, declaration, expected] of declared) {
    it(`${expression} is ${expected} with ${declaration.split('\n')[0]}`, () => {
      const engine = new VbsEngine();
      engine.executeStatement(`On Error Resume Next\n${declaration}`);
      engine.executeStatement(`Err.Clear\nresult = ${expression}\nvalue = CStr(Err.Number)`);
      const failure = engine.error;
      const number =
        failure && failure.number !== 0
          ? failure.number
          : Number(engine._getVariable('value').value);
      expect(number === 450 ? '450' : 'ok').toBe(expected);
    });
  }
});
