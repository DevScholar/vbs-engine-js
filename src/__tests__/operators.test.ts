import { describe, it, expect } from 'vitest';
import { VbsEngine } from '../index.ts';

// Every operator against every pair of VBScript types, measured against a fresh cscript under 1033: twelve
// operands (Empty, Null, Boolean, Byte, Integer, Long, Single, Double, Currency, Date and two Strings) under
// twenty operators, 2899 cells. The capture is arith-1033.txt in the probe directory.
//
// The operands are set again before every cell there, because a bitwise operator converts an Empty operand
// in place: after "Empty And Empty" the variable itself is a Long 0 and no longer Empty. That is what
// VBScript does, and the engine holds its values as plain objects, so the engine does not reproduce the
// change to the operand. A script can see it, and it is the one thing this table does not compare.
//
// The cells below are the ones the engine answers as Windows does so far. The rest of the matrix is still
// being worked through, cluster by cluster: the promotion table for "+" and "*", the result type of "\" and
// Mod, the Single that "/" answers, "Object required" for Is, the type mismatch a non numeric string raises
// under a bitwise operator, and CStr of a Null.
const CELLS: ReadonlyArray<readonly [string, string]> = [
  // Is compares two object references and nothing else, so a plain value on either side is error 424 rather
  // than False. Measured for all 144 pairs of the twelve types.
  ['Empty Is Empty', 'ERR 424'],
  ['Empty Is Null', 'ERR 424'],
  ['Empty Is Boolean', 'ERR 424'],
  ['Empty Is Byte', 'ERR 424'],
  ['Empty Is Integer', 'ERR 424'],
  ['Empty Is Long', 'ERR 424'],
  ['Empty Is Single', 'ERR 424'],
  ['Empty Is Double', 'ERR 424'],
  ['Empty Is Currency', 'ERR 424'],
  ['Empty Is Date', 'ERR 424'],
  ['Empty Is "2"', 'ERR 424'],
  ['"2" Is "2"', 'ERR 424'],
  ['Null Is Boolean', 'ERR 424'],
  ['Date Is Currency', 'ERR 424'],
  ['CByte(200) Is CLng(70000)', 'ERR 424'],
];

function answer(expression: string, lcid = 1033): unknown {
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

describe('an operator between two types answers what vbscript.dll answers', () => {
  for (const [expression, expected] of CELLS) {
    it(`${expression} is ${expected}`, () => {
      expect(answer(expression)).toBe(expected);
    });
  }
});
