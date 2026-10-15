import { describe, it, expect } from 'vitest';
import { VbsEngine } from '../index.ts';

// Arrays, Erase and the bounds, measured against a fresh cscript under 1033.
//
//   * Dim a() makes a dynamic array with no size yet: IsArray is True, VarType is 8204, TypeName is
//     "Variant()" and both bounds raise error 9 until a ReDim gives it a size.
//   * A bound is the last index and not a count, so Dim a(3) runs from 0 to 3 and ReDim a(4) gives five
//     elements. ReDim Preserve keeps what fits and may only change the last dimension, where changing any
//     other one is error 9.
//   * Erase turns every element of a fixed array to Empty and keeps its bounds, while a dynamic array is
//     freed: IsArray stays True and UBound raises error 9. Erasing something that is not an array is a type
//     mismatch.
//   * A dimension outside the array and an index outside its bounds are error 9, and a bound of something
//     that is not an array is error 13, a Null included.
//   * Assigning an array copies it, so after b = a neither changing b(0) nor erasing b touches a.
//   * Filter reads both sides as text, so Filter(Array(1, 2), 1) answers with the element 1.
const CELLS: ReadonlyArray<readonly [string[], string]> = [
  [['r = TypeName(Array())'], 'String | Variant()'], // arr.empty.type
  [['r = UBound(Array())'], 'Long | -1'], // arr.empty.ubound
  [['r = LBound(Array())'], 'Long | 0'], // arr.empty.lbound
  [['r = UBound(Array(1))'], 'Long | 0'], // arr.one
  [['r = Join(Array(1, 2, 3), "|")'], 'String | 1|2|3'], // arr.three
  [['r = IsArray(Array())'], 'Boolean | True'], // arr.isarray
  [['a = Array(Array(1, 2))', 'r = UBound(a(0))'], 'Long | 1'], // arr.nested
  [['a = Array(1, 2)', 'r = Join(a, "|")'], 'String | 1|2'], // arr.assign
  [['r = "[" & Join(Array(), "|") & "]"'], 'String | []'], // arr.emptyjoin
  [['a = Array(Null)', 'r = IsNull(a(0))'], 'Boolean | True'], // arr.null.el
  [['a = Array(Empty)', 'r = IsEmpty(a(0))'], 'Boolean | True'], // arr.empty.el
  [['a = Array(1)', 'r = TypeName(a(0))'], 'String | Integer'], // arr.type.el
  [['a = Array(1, "x")', 'r = TypeName(a(0)) & TypeName(a(1))'], 'String | IntegerString'], // arr.multi.type
  [['Dim a(1)', 'Dim b', 'a(0) = 5', 'b = a', 'b(0) = 9', 'r = a(0)'], 'Integer | 5'], // arr.copy
  [['r = IsArray(5)'], 'Boolean | False'], // isarray.scalar
  [['r = IsArray("x")'], 'Boolean | False'], // isarray.string
  [['Dim a(3)', 'r = LBound(a) & "-" & UBound(a)'], 'String | 0-3'], // dim.fixed.bounds
  [['Dim a(0)', 'r = UBound(a)'], 'Long | 0'], // dim.fixed.zero
  [['Dim a(1, 2)', 'r = UBound(a, 2)'], 'Long | 2'], // dim.two.bounds
  [['Dim a(1, 2)', 'r = LBound(a, 2)'], 'Long | 0'], // dim.two.lbound
  [['Dim a(1, 2)', 'r = a(2, 0)'], 'ERR 9'], // dim.two.oob
  [['Dim a(3)', 'r = LBound(a, 1)'], 'Long | 0'], // lb.dim1
  [['Dim a(3)', 'r = LBound(a, 2)'], 'ERR 9'], // lb.dim2
  [['Dim a(3)', 'r = UBound(a, 0)'], 'ERR 9'], // ub.dim0
  [['r = LBound("x")'], 'ERR 13'], // lb.nonarray
  [['r = UBound("x")'], 'ERR 13'], // ub.nonarray
  [['r = LBound(Null)'], 'ERR 13'], // lb.null
  [['Dim a(2)', 'r = a(3)'], 'ERR 9'], // el.read.oob
  [['Dim a(2)', 'r = a(-1)'], 'ERR 9'], // el.read.neg
  [['Dim a(2)', 'a(3) = 1', 'r = "ok"'], 'ERR 9'], // el.write.oob
  [['Dim a(2)', 'a(-1) = 1', 'r = "ok"'], 'ERR 9'], // el.write.neg
  [['Dim a(2)', 'a(2) = 7', 'r = a(2)'], 'Integer | 7'], // el.write.read
  [['Dim a(2)', 'a(0) = 5', 'Erase a', 'r = a(0)'], 'Empty | '], // erase.numeric
  [['Dim a(2)', 'a(0) = 5', 'Erase a', 'r = IsEmpty(a(0))'], 'Boolean | True'], // erase.isempty
  [['Dim a(2)', 'Erase a', 'r = TypeName(a)'], 'String | Variant()'], // erase.type
  [['Dim a(2)', 'Erase a', 'r = UBound(a)'], 'Long | 2'], // erase.ubound
  [['Dim a(2)', 'Erase a', 'r = IsArray(a)'], 'Boolean | True'], // erase.isarray
  [['Dim a(2)', 'a(0) = "x"', 'Erase a', 'r = "[" & a(0) & "]"'], 'String | []'], // erase.string
  [['Dim a()', 'ReDim a(2)', 'Erase a', 'r = "ok"'], 'String | ok'], // erase.dynamic.ok
  [['Dim a()', 'ReDim a(2)', 'Erase a', 'r = UBound(a)'], 'ERR 9'], // erase.dynamic.ubound
  [['Dim a()', 'ReDim a(2)', 'Erase a', 'r = IsArray(a)'], 'Boolean | True'], // erase.dynamic.isarray
  [['x = 5', 'Erase x', 'r = "ok"'], 'ERR 13'], // erase.nonarray
  [
    ['Dim a()', 'ReDim a(2)', 'a(0) = "x"', 'ReDim Preserve a(4)', 'r = a(0) & UBound(a)'],
    'String | x4',
  ], // redim.preserve
  [['Dim a()', 'ReDim a(4)', 'a(4) = "x"', 'ReDim Preserve a(2)', 'r = UBound(a)'], 'Long | 2'], // redim.shrink
  [['Dim a()', 'ReDim a(2)', 'a(0) = "x"', 'ReDim a(4)', 'r = IsEmpty(a(0))'], 'Boolean | True'], // redim.nopreserve
  [['Dim a()', 'ReDim a(1, 2)', 'ReDim Preserve a(2, 2)', 'r = "ok"'], 'ERR 9'], // redim.two.preserve
  [['Dim a()', 'ReDim a(1, 2)', 'ReDim Preserve a(1, 3)', 'r = UBound(a, 2)'], 'Long | 3'], // redim.two.second
  [['Dim a()', 'ReDim a(3)', 'r = LBound(a)'], 'Long | 0'], // redim.lbound.after
  [['a = Split("1,2", ",")', 'r = TypeName(a(0))'], 'String | String'], // split.el.type
  [['r = LBound(Split("a", ","))'], 'Long | 0'], // split.lbound
  [['a = Filter(Array("aa", "bb", "ab"), "a")', 'r = Join(a, "|")'], 'String | aa|ab'], // filter.basic
  [['a = Filter(Array("aa", "bb"), "a", False)', 'r = Join(a, "|")'], 'String | bb'], // filter.exclude
  [['a = Filter(Array(), "a")', 'r = UBound(a)'], 'Long | -1'], // filter.empty
  [['a = Filter(Array("AA", "bb"), "a", True, 1)', 'r = Join(a, "|")'], 'String | AA'], // filter.text
  [['a = Filter(Array(1, 2), 1)', 'r = Join(a, "|")'], 'String | 1'], // filter.numbers
  [['a = Filter(Array("aa"), "z")', 'r = UBound(a)'], 'Long | -1'], // filter.nomatch
  [['r = Join(Array(1, 2.5), "-")'], 'String | 1-2.5'], // join.numbers
  [['r = Join(Array("x"), "-")'], 'String | x'], // join.one
  [['r = "[" & Join(Array("", ""), "|") & "]"'], 'String | [|]'], // join.empty.el
];

function answer(statements: string[]): unknown {
  const engine = new VbsEngine();
  try {
    engine.executeStatement(`SetLocale 1033\n${statements.join('\n')}`);
    const failure = engine.error;
    if (failure) return 'ERR ' + String(failure.number);
    if (engine._getVariable('r').type === 'Null') return 'Null';
    engine.executeStatement('t = TypeName(r)\ns = CStr(r)');
    return String(engine._getVariable('t').value) + ' | ' + String(engine._getVariable('s').value);
  } finally {
    engine.executeStatement('SetLocale 1033');
  }
}

describe('arrays, Erase and the bounds behave like vbscript.dll', () => {
  for (const [statements, expected] of CELLS) {
    it(`${statements.join('; ')} is ${expected}`, () => {
      expect(answer(statements)).toBe(expected);
    });
  }
});
