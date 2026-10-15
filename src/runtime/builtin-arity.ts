/**
 * How many arguments each VBScript built-in takes, measured against a fresh cscript under 1033.
 *
 * The probe calls every function with 0, 1, 2 ... arguments, one written-out call per cell, and reads the
 * number the call left in Err: 450 is "Wrong number of arguments or invalid property assignment" and
 * anything else means the count was accepted (a call that got past the count can still fail on its values,
 * which is how CreateObject answers 429). The capture is arity-1033.txt in the probe directory, 326 cells
 * over 89 functions.
 *
 * Two things the measurement settled, both of which the engine used to get wrong:
 *
 *   * The count is checked on both sides. Mid needs 2 or 3: Mid("abc") is 450 and Mid("abc", 1, 1, 1) is
 *     450, while the engine ignored the extra argument.
 *   * A declared procedure is checked the same way. Function One(value) called as One() or as One(1, 2) is
 *     450, and a Sub with no parameters called with one is 450 too.
 *
 * Array is the one function with no upper bound: Array() and Array(1, 2, 3, 4) are both accepted.
 *
 * A name that is not in this table is not checked, which covers the engine's own extensions and the few
 * functions nobody measured (MsgBox and InputBox are checked on the low side only, since a higher count
 * would have raised a dialog in the probe).
 */
export const BUILTIN_ARITY: Readonly<Record<string, readonly [number, number]>> = {
  abs: [1, 1],
  array: [0, Number.POSITIVE_INFINITY],
  asc: [1, 1],
  atn: [1, 1],
  cbool: [1, 1],
  cbyte: [1, 1],
  ccur: [1, 1],
  cdate: [1, 1],
  cdbl: [1, 1],
  chr: [1, 1],
  cint: [1, 1],
  clng: [1, 1],
  cos: [1, 1],
  createobject: [1, 2],
  cstr: [1, 1],
  dateadd: [3, 3],
  datediff: [3, 5],
  datepart: [2, 4],
  dateserial: [3, 3],
  datevalue: [1, 1],
  day: [1, 1],
  escape: [1, 1],
  eval: [1, 1],
  execute: [1, 1],
  exp: [1, 1],
  filter: [2, 4],
  fix: [1, 1],
  formatcurrency: [1, 5],
  formatdatetime: [1, 2],
  formatnumber: [1, 5],
  formatpercent: [1, 5],
  getlocale: [0, 1],
  getref: [1, 1],
  hex: [1, 1],
  hour: [1, 1],
  inputbox: [1, Number.POSITIVE_INFINITY],
  instr: [2, 4],
  instrrev: [2, 4],
  int: [1, 1],
  isarray: [1, 1],
  isdate: [1, 1],
  isempty: [1, 1],
  isnull: [1, 1],
  isnumeric: [1, 1],
  isobject: [1, 1],
  join: [1, 2],
  lbound: [1, 2],
  lcase: [1, 1],
  left: [2, 2],
  len: [1, 1],
  loadpicture: [1, 1],
  log: [1, 1],
  ltrim: [1, 1],
  mid: [2, 3],
  minute: [1, 1],
  month: [1, 1],
  monthname: [1, 2],
  msgbox: [1, Number.POSITIVE_INFINITY],
  now: [0, 1],
  oct: [1, 1],
  replace: [3, 6],
  rgb: [3, 3],
  right: [2, 2],
  rnd: [0, 1],
  round: [1, 2],
  rtrim: [1, 1],
  second: [1, 1],
  sgn: [1, 1],
  sin: [1, 1],
  space: [1, 1],
  split: [1, 4],
  sqr: [1, 1],
  strcomp: [2, 3],
  string: [2, 2],
  strreverse: [1, 1],
  tan: [1, 1],
  time: [0, 1],
  timer: [0, 1],
  timeserial: [3, 3],
  timevalue: [1, 1],
  trim: [1, 1],
  typename: [1, 1],
  ubound: [1, 2],
  ucase: [1, 1],
  unescape: [1, 1],
  vartype: [1, 1],
  weekday: [1, 2],
  weekdayname: [1, 3],
  year: [1, 1],
};
