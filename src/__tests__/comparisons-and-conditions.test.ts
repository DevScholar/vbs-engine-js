import { describe, it, expect } from "vitest";
import { VbsEngine } from "../index.ts";

// How a comparison reads its two sides, and how a condition reads a value. Both are measured against a fresh
// cscript under 1033 with the text comparison and condition probes in the probe directory.
//
// The comparison rules, each measured:
//
//   * Text is compared character by character and case sensitively: "b" = "B" is False and "B" < "b" is
//     True, in 1033, 1031, 1036 and 2052 alike.
//   * A text beside a number is read as a number when the text is one, so "1" = 1, "1.5" = 1.5, " 1 " = 1
//     and "2" < 10 are True and "2" > 10 is False.
//   * A text that is no number beside a number is a type mismatch, but only when the number was written out
//     as a literal: "abc" = 1 and "abc" = (1) are error 13, while "abc" = 1 * 1, "abc" = 1 + 0,
//     "abc" = CInt(1), "abc" = Len("a") and "abc" = heldNumber (heldNumber = 1) are all False.
//   * A text beside anything else is compared as text, which is why "True" = True is True, "1" = True is
//     False, "" = Empty is True, "x" = Empty is False and "1/5/2026" = CDate("2026-01-05") is True.
//   * A Null on either side answers Null, for every operator: "x" = Null, 1 = Null, "x" < Null and
//     Null <> Null are all Null.
//   * An object on either side is error 91 when it is Nothing and error 438 otherwise, and an array is
//     error 13.
//
// The condition rules, measured for If and for Do While with the same answers in both:
//
//   * A number is true when it is not zero, a Date is true, an Empty is false and a Null is false.
//   * A text is read as a number, so "1" and " 1 " and "1.5" and "&H10" are true and "0" is false, the words
//     True and False are read as themselves, and a text that is no number is a type mismatch: "x", "1x",
//     "" and " " are all error 13.
//   * An array is a type mismatch, Nothing is error 91 and any other object is error 438.
const COMPARISONS: ReadonlyArray<readonly [string, string]> = [
  ['"a" = "a"', "True"],
  ['"b" = "B"', "False"],
  ['"B" <> "b"', "True"],
  ['"b" < "B"', "False"],
  ['"B" < "b"', "True"],
  ['"apple" < "banana"', "True"],
  ['"1" = 1', "True"],
  ['"1.5" = 1.5', "True"],
  ['" 1 " = 1', "True"],
  ['"2" < 10', "True"],
  ['"2" > 10', "False"],
  ['"abc" = 1', "ERR 13"],
  ['1 = "abc"', "ERR 13"],
  ['"abc" < 1', "ERR 13"],
  ['"abc" <= 1', "ERR 13"],
  ['"abc" = (1)', "ERR 13"],
  ['"abc" = 1 * 1', "False"],
  ['"abc" = 1 + 0', "False"],
  ['"abc" = CInt(1)', "False"],
  ['"abc" = CByte(1)', "False"],
  ['"abc" = CLng(1)', "False"],
  ['"abc" = CSng(1)', "False"],
  ['"abc" = CDbl(1)', "False"],
  ['"abc" = CCur(1)', "False"],
  ['"abc" = CDbl(1.5)', "False"],
  ['"abc" = Len("a")', "False"],
  ['"abc" < CInt(1)', "False"],
  ['"abc" > CInt(1)', "True"],
  ['"abc" <> CInt(1)', "True"],
  ['"True" = True', "True"],
  ['True = "True"', "True"],
  ['"1" = True', "False"],
  ['"true" = True', "False"],
  ['"" = Empty', "True"],
  ['"x" = Empty', "False"],
  ['"x" < Empty', "False"],
  ['"1/5/2026" = CDate("2026-01-05")', "True"],
  ['"January 5, 2026" = CDate("2026-01-05")', "False"],
  ['"2026-01-05" = CDate("2026-01-05")', "False"],
  ['"2026-01-05" < CDate("2026-01-06")', "False"],
  ['CDate("2026-01-05") = "abc"', "False"],
  ["Null = 1", "Null"],
  ["1 = Null", "Null"],
  ['"x" = Null', "Null"],
  ['"x" < Null', "Null"],
  ["Null <> Null", "Null"],
  ["Nothing = Nothing", "ERR 91"],
  ["Array(1) = 1", "ERR 13"],
  ["Array(1) = Null", "ERR 13"],
];

describe("how a comparison reads its two sides", () => {
  for (const [expression, expected] of COMPARISONS) {
    it(`${expression} is ${expected}`, () => {
      expect(answer(expression)).toBe(expected);
    });
  }
});

// Select Case asks the same question, which a fresh cscript shows for the same pairs.
const SELECT_CASES: ReadonlyArray<readonly [string, string]> = [
  [
    'Select Case "b"\nCase "B"\nresult = "upper"\nCase Else\nresult = "else"\nEnd Select',
    "else",
  ],
  [
    'Select Case "b"\nCase "b"\nresult = "lower"\nCase Else\nresult = "else"\nEnd Select',
    "lower",
  ],
  [
    'Select Case "abc"\nCase 1\nresult = "one"\nCase Else\nresult = "else"\nEnd Select',
    "ERR 13",
  ],
  [
    'Select Case 1\nCase "abc"\nresult = "text"\nCase Else\nresult = "else"\nEnd Select',
    "ERR 13",
  ],
  [
    'Select Case "1"\nCase 1\nresult = "one"\nCase Else\nresult = "else"\nEnd Select',
    "one",
  ],
  [
    'Select Case "abc"\nCase "abc", 3\nresult = "hit"\nCase Else\nresult = "else"\nEnd Select',
    "hit",
  ],
  [
    'Select Case Null\nCase 1\nresult = "one"\nCase Else\nresult = "else"\nEnd Select',
    "else",
  ],
  [
    'Select Case 1\nCase Null\nresult = "null"\nCase Else\nresult = "else"\nEnd Select',
    "else",
  ],
  [
    'Select Case Empty\nCase 0\nresult = "zero"\nCase Else\nresult = "else"\nEnd Select',
    "zero",
  ],
  [
    'Select Case 0\nCase Empty\nresult = "empty"\nCase Else\nresult = "else"\nEnd Select',
    "empty",
  ],
  [
    'Select Case -1\nCase True\nresult = "true"\nCase Else\nresult = "else"\nEnd Select',
    "true",
  ],
  [
    'Select Case 2\nCase 1, 2\nresult = "first"\nCase 2, 3\nresult = "second"\nEnd Select',
    "first",
  ],
  [
    'Select Case 9\nCase 1\nresult = "one"\nCase Else\nresult = "else"\nEnd Select',
    "else",
  ],
  [
    'result = "start"\nSelect Case 9\nCase 1\nresult = "one"\nEnd Select',
    "start",
  ],
  [
    'Select Case CDate("2026-01-05")\nCase CDate("2026-01-05")\nresult = "same"\nCase Else\nresult = "else"\nEnd Select',
    "same",
  ],
  [
    "n = 0\nSelect Case NextCount(n)\nCase 1\nresult = CStr(n)\nEnd Select",
    "1",
  ],
];

describe("how Select Case picks a case", () => {
  for (const [source, expected] of SELECT_CASES) {
    it(`${source.split("\n")[0]} with ${source.split("\n")[1]} is ${expected}`, () => {
      expect(run(source)).toBe(expected);
    });
  }
});

// A Select Case range or an Is test is not VBScript at all: both are compile errors in a fresh cscript, so
// there is nothing to compare them against. The engine accepts `Case Is > 5` as a VB6 extension and refuses
// `Case 1 To 5` with a parse error, which is the same outcome Windows gives for the range.
const CONDITIONS: ReadonlyArray<readonly [string, string, string]> = [
  ["if.number.one", "1", "ran"],
  ["if.number.zero", "0", "skipped"],
  ["if.number.minusOne", "-1", "ran"],
  ["if.number.decimal", "1.5", "ran"],
  ["if.boolean.true", "True", "ran"],
  ["if.boolean.false", "False", "skipped"],
  ["if.text.one", '"1"', "ran"],
  ["if.text.zero", '"0"', "skipped"],
  ["if.text.padded", '" 1 "', "ran"],
  ["if.text.decimal", '"1.5"', "ran"],
  ["if.text.hex", '"&H10"', "ran"],
  ["if.text.true", '"True"', "ran"],
  ["if.text.letter", '"x"', "ERR 13"],
  ["if.text.empty", '""', "ERR 13"],
  ["if.text.spaces", '" "', "ERR 13"],
  ["if.text.trailing", '"1x"', "ERR 13"],
  ["if.empty", "blank", "skipped"],
  ["if.null", "nothingValue", "skipped"],
  ["if.date", "dateValue", "ran"],
  ["if.array", "Array(1)", "ERR 13"],
  ["if.nothing", "Nothing", "ERR 91"],
  ["if.object", "heldObject", "ERR 438"],
  ["do.text.letter", '"x"', "ERR 13"],
  ["do.null", "nothingValue", "skipped"],
  ["do.object", "heldObject", "ERR 438"],
];

describe("how a condition reads a value", () => {
  for (const [label, expression, expected] of CONDITIONS) {
    it(`${label} is ${expected}`, () => {
      const kind = label.startsWith("do.") ? "do" : "if";
      expect(run(conditionSource(kind, expression))).toBe(expected);
    });
  }
});

/** Runs one expression comparison with a handler in place and answers it as text. */
function answer(expression: string): string {
  const engine = new VbsEngine();
  engine.executeStatement(
    `On Error Resume Next\nSetLocale 1033\nDim blank, nothingValue, dateValue\nblank = Empty\nnothingValue = Null\ndateValue = CDate("2026-01-05")\nErr.Clear\nvalue = (${expression})\nnumber = CStr(Err.Number)`,
  );
  const failure = engine.error;
  const number =
    failure && failure.number !== 0
      ? failure.number
      : Number(engine._getVariable("number").value);
  if (number !== 0) return `ERR ${number}`;
  const result = engine._getVariable("value");
  if (result.type === "Boolean") return result.value ? "True" : "False";
  if (result.type === "Null") return "Null";
  return String(result.value);
}

/** Runs a block that leaves its answer in result, and answers the answer or the error number. */
function run(source: string): string {
  const engine = new VbsEngine();
  engine.executeStatement(
    `On Error Resume Next\nSetLocale 1033\nFunction NextCount(counter)\ncounter = counter + 1\nNextCount = counter\nEnd Function\nClass Holder\nEnd Class\nDim result, n, blank, nothingValue, dateValue, heldObject\nblank = Empty\nnothingValue = Null\ndateValue = CDate("2026-01-05")\nSet heldObject = New Holder\nErr.Clear\n${source}\nnumber = CStr(Err.Number)`,
  );
  const failure = engine.error;
  const number =
    failure && failure.number !== 0
      ? failure.number
      : Number(engine._getVariable("number").value);
  if (number !== 0) return `ERR ${number}`;
  return String(engine._getVariable("result").value);
}

/** The source of one condition cell, written the way the probe writes it. */
function conditionSource(kind: string, expression: string): string {
  if (kind === "do") {
    return `n = 0\nDo While ${expression}\nn = n + 1\nIf n > 3 Then Exit Do\nLoop\nresult = "skipped"\nIf n > 0 Then result = "ran"`;
  }
  return `n = 0\nIf ${expression} Then n = 1\nresult = "ran"\nIf n = 0 Then result = "skipped"`;
}
