import { describe, it, expect } from "vitest";
import { VbsEngine } from "../index.ts";

// Arrays measured against a fresh cscript under 1033: ReDim and ReDim Preserve, Erase and the shapes it
// accepts, the index and the dimension read as whole numbers, and the edges of Join, Filter and Split.
//
//   * A fixed array cannot be sized again: Dim a(2) then ReDim a(3) and then ReDim Preserve a(3) are both
//     error 10, and Erase of it does not free that hold on it.
//   * ReDim Preserve may only change the last dimension of an array that already has one; changing any other
//     dimension, or how many there are, is error 9. An array with no size yet, from Dim a() or from Erase,
//     takes any shape Preserve asks for.
//   * Erase frees an array that Array, Split or Filter made, so UBound and any element read are error 9
//     afterwards. Erase takes one name: Erase a(0) frees what the element holds, Erase(a) and
//     Erase Array(1, 2) are error 13, and Erase a, b is error 450.
//   * An index and a dimension are whole numbers rounded to the nearest one, a tie going to the even
//     neighbour, so a(1.5) is a(2) and UBound(a, 2.6) is the third dimension. A Null index is error 13 where
//     a Null dimension is error 94, and a number too large for a Long is Overflow.
//   * A name Dim declared as an array takes no whole value afterwards: Dim a(1) with a = 5 is error 13, and
//     Dim a() with a = Array(1, 2) is error 9.
//   * Join reads a one dimensional array and refuses a Null element, Filter keeps the elements it found and
//     reads its compare argument, and Split reads its text as text.
//
// Each cell runs in its own engine, with the procedures and the class the probe declares.

const HELPERS = `Sub FillWithNine(values)
  values(0) = 9
End Sub

Sub ClearWholeArray(values)
  Erase values
End Sub

Sub SizeArrayToThree(values)
  ReDim values(3)
End Sub

Function ArrayUpperBound(values)
  ArrayUpperBound = UBound(values)
End Function

Function MakeTwoElementArray()
  MakeTwoElementArray = Array(1, 2)
End Function

Sub ReplaceWholeArray(values)
  values = Array(5, 6)
End Sub

Sub FillWithNineByValue(ByVal values)
  values(0) = 9
End Sub

Class Holder
End Class

Class ArrayHolder
  Public Items
End Class`;

const CELLS: ReadonlyArray<readonly [string[], string]> = [
  [["Dim a(2)", "ReDim a(3)", "r = UBound(a)"], "ERR 10"], // redim.fixed.one
  [["Dim a(2)", "ReDim Preserve a(3)", "r = UBound(a)"], "ERR 10"], // redim.fixed.preserve
  [["Dim a(2)", 'a(0) = "x"', "ReDim a(4)", "r = a(0)"], "ERR 10"], // redim.fixed.keeps.values
  [["Dim a()", "ReDim Preserve a(2)", "r = UBound(a)"], "Long | 2"], // redim.preserve.unsized
  [
    [
      "Dim a()",
      "ReDim a(3)",
      'a(1) = "x"',
      "ReDim Preserve a(3)",
      "r = a(1) & UBound(a)",
    ],
    "String | x3",
  ], // redim.preserve.same
  [
    [
      "Dim a()",
      "ReDim a(1)",
      'a(0) = "x"',
      "ReDim Preserve a(3)",
      "r = a(0) & IsEmpty(a(2)) & IsEmpty(a(3))",
    ],
    "String | xTrueTrue",
  ], // redim.preserve.gap
  [
    [
      "Dim a()",
      "ReDim a(3)",
      'a(0) = "y"',
      'a(3) = "z"',
      "ReDim Preserve a(1)",
      "ReDim Preserve a(3)",
      "r = a(0) & IsEmpty(a(3))",
    ],
    "String | yTrue",
  ], // redim.preserve.shrink.grow
  [
    [
      "Dim a()",
      "ReDim a(4)",
      'a(4) = "x"',
      "ReDim Preserve a(2)",
      "r = UBound(a)",
    ],
    "Long | 2",
  ], // redim.preserve.shrink.below
  [
    [
      "Dim a()",
      "ReDim a(4)",
      "ReDim Preserve a(2)",
      "ReDim Preserve a(4)",
      "r = UBound(a)",
    ],
    "Long | 4",
  ], // redim.preserve.shrink.again
  [["Dim a()", "ReDim a(-1)", 'r = "ok"'], "String | ok"], // redim.negative
  [
    ["Dim a()", "ReDim a(2)", "ReDim Preserve a(-1)", 'r = "ok"'],
    "String | ok",
  ], // redim.negative.preserve
  [["Dim a()", "ReDim a(2.4)", "r = UBound(a)"], "Long | 2"], // redim.round.down
  [["Dim a()", "ReDim a(2.6)", "r = UBound(a)"], "Long | 3"], // redim.round.up
  [["Dim a()", "ReDim a(2.5)", "r = UBound(a)"], "Long | 2"], // redim.round.half.two
  [["Dim a()", "ReDim a(1.5)", "r = UBound(a)"], "Long | 2"], // redim.round.half.one
  [["Dim a()", "ReDim a(0.5)", "r = UBound(a)"], "Long | 0"], // redim.round.half.zero
  [["Dim a()", "ReDim a(0)", "r = UBound(a)"], "Long | 0"], // redim.zero
  [["Dim a()", 'n = "3"', "ReDim a(n)", "r = UBound(a)"], "Long | 3"], // redim.bound.string
  [["Dim a()", "n = Empty", "ReDim a(n)", "r = UBound(a)"], "Long | 0"], // redim.bound.empty
  [["Dim a()", "n = Null", "ReDim a(n)", 'r = "ok"'], "ERR 94"], // redim.bound.null
  [["Dim a()", 'n = "x"', "ReDim a(n)", 'r = "ok"'], "ERR 13"], // redim.bound.text
  [["Dim a()", "n = 3", "ReDim a(n)", "r = UBound(a)"], "Long | 3"], // redim.bound.variable
  [["Dim a()", "n = 1", "ReDim a(n + 2)", "r = UBound(a)"], "Long | 3"], // redim.bound.expression
  [["x = 5", "ReDim x(2)", "r = UBound(x)"], "Long | 2"], // redim.scalar.variant
  [['x = "s"', "ReDim x(2)", "r = UBound(x)"], "Long | 2"], // redim.scalar.text
  [["a = Array(1, 2)", "ReDim a(4)", "r = UBound(a)"], "Long | 4"], // redim.array.result
  [['a = Split("a,b", ",")', "ReDim a(4)", "r = UBound(a)"], "Long | 4"], // redim.split.result
  [
    ["a = Array(1, 2)", "ReDim Preserve a(4)", "r = a(0) & UBound(a)"],
    "String | 14",
  ], // redim.preserve.array.result
  [
    ["Dim a()", "ReDim a(2)", "Erase a", "ReDim a(3)", "r = UBound(a)"],
    "Long | 3",
  ], // redim.after.erase.dynamic
  [
    [
      "Dim a()",
      "ReDim a(2)",
      "Erase a",
      "ReDim Preserve a(3)",
      "r = UBound(a)",
    ],
    "Long | 3",
  ], // redim.preserve.after.erase
  [
    ["Dim a(2)", 'a(0) = "x"', "Erase a", "ReDim a(4)", "r = UBound(a)"],
    "ERR 10",
  ], // redim.after.erase.fixed
  [["Dim a()", "ReDim a(1, 2)", "ReDim a(3)", "r = UBound(a)"], "Long | 3"], // redim.two.to.one
  [
    ["Dim a()", "ReDim a(1, 2)", "ReDim Preserve a(3)", "r = UBound(a)"],
    "ERR 9",
  ], // redim.two.to.one.preserve
  [["Dim a()", "ReDim a(2)", "ReDim a(1, 1)", "r = UBound(a, 2)"], "Long | 1"], // redim.one.to.two
  [
    ["Dim a()", "ReDim a(2)", "ReDim Preserve a(1, 1)", "r = UBound(a, 2)"],
    "ERR 9",
  ], // redim.one.to.two.preserve
  [
    [
      "Dim a()",
      "ReDim a(1, 2)",
      "ReDim a(3, 4)",
      "r = UBound(a, 1) & UBound(a, 2)",
    ],
    "String | 34",
  ], // redim.two.same.dims
  [
    ["Dim a()", "ReDim a(1, 2)", "ReDim Preserve a(1, 2)", 'r = "ok"'],
    "String | ok",
  ], // redim.two.preserve.same
  [["Dim a()", "ReDim a(1, 2)", "ReDim Preserve a(2, 2)", 'r = "ok"'], "ERR 9"], // redim.two.preserve.first
  [["Dim a()", "ReDim a(1, 2)", "ReDim Preserve a(2, 3)", 'r = "ok"'], "ERR 9"], // redim.two.preserve.both
  [
    [
      "Dim a()",
      "ReDim a(1, 2)",
      "ReDim Preserve a(1, 4)",
      "r = a(0, 0) & UBound(a, 2)",
    ],
    "String | 4",
  ], // redim.two.preserve.second.grow
  [
    ["Dim a()", "ReDim a(1, 3)", "ReDim Preserve a(1, 1)", "r = UBound(a, 2)"],
    "Long | 1",
  ], // redim.two.preserve.second.shrink
  [
    [
      "Dim a()",
      "ReDim a(1, 2)",
      'a(0, 0) = "x"',
      'a(1, 2) = "y"',
      "ReDim Preserve a(1, 4)",
      "r = a(0, 0) & a(1, 2) & UBound(a, 2)",
    ],
    "String | xy4",
  ], // redim.two.preserve.values
  [
    [
      "Dim a()",
      "ReDim a(1, 2)",
      'a(1, 2) = "y"',
      "ReDim Preserve a(1, 1)",
      "r = UBound(a, 1)",
    ],
    "Long | 1",
  ], // redim.two.preserve.drop.last
  [
    ["Dim a()", "ReDim a(1, 1, 1)", "ReDim Preserve a(1, 2, 1)", 'r = "ok"'],
    "ERR 9",
  ], // redim.three.preserve.second
  [
    [
      "Dim a()",
      "ReDim a(1, 1, 1)",
      "ReDim Preserve a(1, 1, 3)",
      "r = UBound(a, 3)",
    ],
    "Long | 3",
  ], // redim.three.preserve.third
  [
    [
      "Dim a()",
      "ReDim a(3)",
      "ReDim Preserve a(5)",
      "r = LBound(a) & UBound(a)",
    ],
    "String | 05",
  ], // redim.lower.after.preserve
  [
    [
      "Dim a(1, 1)",
      'a(0, 0) = "x"',
      "Erase a",
      "r = IsEmpty(a(0, 0)) & UBound(a, 2)",
    ],
    "String | True1",
  ], // erase.two.dim
  [["a = Array(1, 2)", "Erase a", "r = IsEmpty(a(1)) & UBound(a)"], "ERR 9"], // erase.array.result
  [["a = Array(1, 2)", "Erase a", "r = TypeName(a)"], "String | Variant()"], // erase.array.result.type
  [
    ['a = Split("a,b", ",")', "Erase a", 'r = "[" & a(0) & "]" & UBound(a)'],
    "ERR 9",
  ], // erase.split.result
  [['a = Filter(Array("aa"), "a")', "Erase a", "r = UBound(a)"], "ERR 9"], // erase.filter.result
  [["a = Array(1, 2)", "Erase a", "ReDim a(3)", "r = UBound(a)"], "Long | 3"], // erase.array.result.then.redim
  [
    ["Dim a()", "Erase a", "r = TypeName(a) & IsArray(a)"],
    "String | Variant()True",
  ], // erase.dynamic.type
  [["Dim a()", "ReDim a(2)", "Erase a", "r = VarType(a)"], "Integer | 8204"], // erase.dynamic.vartype
  [["x = Empty", "Erase x", 'r = "ok"'], "ERR 13"], // erase.empty.variant
  [["Set o = Nothing", "Erase o", 'r = "ok"'], "ERR 13"], // erase.nothing
  [["Dim a(2)", "a(0) = 5", "Erase a", "Erase a", "r = UBound(a)"], "Long | 2"], // erase.twice.fixed
  [
    ["Dim a()", "ReDim a(2)", "Erase a", "Erase a", "r = IsArray(a)"],
    "Boolean | True",
  ], // erase.twice.dynamic
  [["Dim a(1)", "Erase(a)", 'r = "ok"'], "ERR 13"], // erase.sub.form
  [
    [
      "Dim a(1)",
      "Dim b(1)",
      "a(0) = 1",
      "Erase a, b",
      "r = IsEmpty(a(0)) & IsEmpty(b(0))",
    ],
    "ERR 450",
  ], // erase.two.names
  [["Dim a(1)", "Erase", 'r = "ok"'], "ERR 450"], // erase.no.target
  [["Erase 5", 'r = "ok"'], "ERR 13"], // erase.literal.number
  [['Erase "x"', 'r = "ok"'], "ERR 13"], // erase.literal.string
  [["Erase Nothing", 'r = "ok"'], "ERR 13"], // erase.nothing.literal
  [["Dim a(1)", "Erase(1 + 1)", 'r = "ok"'], "ERR 13"], // erase.parenthesized.expression
  [
    [
      "Dim h",
      "Set h = New ArrayHolder",
      "h.Items = Array(1, 2)",
      "Erase h.Items",
      'r = "ok"',
    ],
    "ERR 13",
  ], // erase.member
  [["Erase z", 'r = "ok"'], "ERR 13"], // erase.undeclared.name
  [
    ["Dim a(1)", "a(0) = Array(1, 2)", "Erase a(0)", "r = IsEmpty(a(0))"],
    "Boolean | False",
  ], // erase.element
  [["Erase Array(1, 2)", 'r = "ok"'], "ERR 13"], // erase.expression
  [["Dim a(1)", "a(0) = 5", "Erase(a)", "r = IsEmpty(a(0))"], "ERR 13"], // erase.parenthesized.fixed
  [["Dim a()", "Erase a", "r = IsArray(a)"], "Boolean | True"], // erase.after.dim.dynamic
  [["Dim a(1, 2)", 'r = UBound(a, "2")'], "Long | 2"], // ub.dim.string
  [["Dim a(1, 2)", "r = UBound(a, Empty)"], "ERR 9"], // ub.dim.empty
  [["Dim a(1, 2)", "r = UBound(a, Null)"], "ERR 94"], // ub.dim.null
  [["Dim a(1, 2)", "r = UBound(a, 1.7)"], "Long | 2"], // ub.dim.half
  [["Dim a(1, 2)", "r = UBound(a, 1.5)"], "Long | 2"], // ub.dim.half.one
  [["Dim a(1, 2)", "r = UBound(a, -1)"], "ERR 9"], // ub.dim.negative
  [["Dim a(1, 2)", "r = UBound(a, 3)"], "ERR 9"], // ub.dim.three
  [["Dim a(1, 2)", "r = UBound(a, True)"], "ERR 9"], // ub.dim.boolean
  [["Dim a(1, 2)", 'r = UBound(a, "x")'], "ERR 13"], // ub.dim.text
  [["Dim a(1, 2)", 'r = LBound(a, "2")'], "Long | 0"], // lb.dim.string
  [["Dim a(2)", "r = UBound(a, 1, 2)"], "ERR 450"], // ub.extra.arg
  [["Dim a()", "r = UBound(a, 1)"], "ERR 9"], // ub.unsized.dynamic
  [["Dim a(2, 3)", "r = UBound(a, 4)"], "ERR 9"], // ub.2d.oob.dimension
  [["a = Array(Array())", "r = UBound(a(0))"], "Long | -1"], // ub.of.nested.empty
  [["Dim a(2)", "r = UBound((a))"], "Long | 2"], // ub.parenthesized
  [["Dim a(2, 3)", "r = LBound(a, 1) & UBound(a, 1)"], "String | 02"], // lb.2d.first
  [["Dim a(3)", 'a(1.4) = "x"', "r = a(1)"], "String | x"], // idx.round.down
  [["Dim a(3)", 'a(1.5) = "x"', "r = a(2)"], "String | x"], // idx.round.half
  [["Dim a(3)", 'a(0.5) = "x"', "r = a(0)"], "String | x"], // idx.round.half.one
  [["Dim a(3)", 'a("2") = "x"', "r = a(2)"], "String | x"], // idx.string
  [["Dim a(3)", 'a(Empty) = "x"', "r = a(0)"], "String | x"], // idx.empty
  [["Dim a(3)", 'a(Null) = "x"', 'r = "ok"'], "ERR 13"], // idx.null
  [["Dim a(3)", 'a(True) = "x"', 'r = "ok"'], "ERR 9"], // idx.boolean
  [["Dim a(3)", "r = IsEmpty(a(-0.4))"], "Boolean | True"], // idx.negative.half
  [["Dim a(3)", 'a("x") = 1', 'r = "ok"'], "ERR 13"], // idx.text
  [["Dim a(1, 1)", "r = a(0, 0, 0)"], "ERR 9"], // idx.too.many
  [["Dim a(1, 1)", "r = a(0)"], "ERR 9"], // idx.too.few
  [["Dim a(1, 2)", "a(1) = 5", 'r = "ok"'], "ERR 9"], // idx.one.for.two
  [["Dim a(3)", 'a((1)) = "x"', "r = a(1)"], "String | x"], // idx.parenthesized
  [["Dim a(3)", "n = 2", 'a(n + 1) = "x"', "r = a(3)"], "String | x"], // idx.expression
  [["Dim a(3)", "r = IsEmpty(a(3))"], "Boolean | True"], // idx.unset.isempty
  [
    ["Dim a(1, 1)", 'a(0, 1) = "x"', "r = a(0, 1) & IsEmpty(a(1, 0))"],
    "String | xTrue",
  ], // idx.two.dim.read
  [["Dim a(3)", "r = a(3000000000)"], "ERR 6"], // idx.huge
  [["Dim a(3)", 'r = a("99999999999999999999")'], "ERR 6"], // idx.overflow.text
  [["Dim a()", "r = TypeName(a)"], "String | Variant()"], // dim.dynamic.type
  [["Dim a()", "r = VarType(a)"], "Integer | 8204"], // dim.dynamic.vartype
  [["Dim a()", "r = IsEmpty(a)"], "Boolean | False"], // dim.dynamic.isempty
  [["Dim a()", "r = IsNull(a)"], "Boolean | False"], // dim.dynamic.isnull
  [["Dim a(2)", "r = VarType(a)"], "Integer | 8204"], // dim.fixed.vartype
  [["Dim a(1, 2)", "r = VarType(a)"], "Integer | 8204"], // dim.two.vartype
  [["Dim a(2)", "r = IsEmpty(a)"], "Boolean | False"], // dim.fixed.isempty
  [["Dim a(2)", "r = IsNull(a)"], "Boolean | False"], // dim.fixed.isnull
  [["Dim a(1), b(2)", "r = UBound(a) & UBound(b)"], "String | 12"], // dim.two.names
  [["Dim a(0, 0)", "r = UBound(a, 2)"], "Long | 0"], // dim.two.zero.zero
  [["Dim x", "r = TypeName(x)"], "String | Empty"], // dim.plain.variant.type
  [["Dim x", "r = UBound(x)"], "ERR 13"], // dim.plain.variant.ubound
  [["x = 5", "r = UBound(x)"], "ERR 13"], // dim.assigned.scalar.ubound
  [["Set o = Nothing", "r = UBound(o)"], "ERR 91"], // dim.nothing.ubound
  [
    [
      "r = UBound(Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20))",
    ],
    "Long | 19",
  ], // array.twenty
  [
    [
      'a = Array(1, 2.5, "s", True, #1/2/2020#)',
      "r = TypeName(a(0)) & TypeName(a(1)) & TypeName(a(2)) & TypeName(a(3)) & TypeName(a(4))",
    ],
    "String | IntegerDoubleStringBooleanDate",
  ], // array.types
  [["a = Array(CCur(1))", "r = TypeName(a(0))"], "String | Currency"], // array.currency.type
  [["a = Array(CLng(1))", "r = TypeName(a(0))"], "String | Long"], // array.long.type
  [["a = Array(Nothing)", "r = TypeName(a(0))"], "String | Nothing"], // array.nothing.type
  [
    ["a = Array(Nothing)", "r = IsNull(a(0)) & IsEmpty(a(0))"],
    "String | FalseFalse",
  ], // array.nothing.flags
  [["Dim a(1)", "a = Array(1, 2, 3)", "r = UBound(a)"], "ERR 13"], // array.assign.to.fixed
  [["Dim a()", "a = Array(1, 2)", "r = UBound(a)"], "ERR 9"], // array.assign.to.dynamic
  [["a = Array(Array(1))", "r = IsArray(a(0))"], "Boolean | True"], // array.element.isarray
  [['r = Join(Array(Empty, Null, "x"), "|")'], "ERR 13"], // array.empty.null.join
  [['a = Array("1")', "r = a(0) + 1"], "Double | 2"], // array.string.coerce
  [["a = Array((1), (2))", 'r = Join(a, ",")'], "String | 1,2"], // array.parenthesized
  [["r = TypeName(Array())"], "String | Variant()"], // array.empty.type
  [["r = VarType(Array())"], "Integer | 8204"], // array.empty.vartype
  [['r = Join(Array("a", "b"))'], "String | a b"], // join.default.delim
  [['r = Join(Array("a", "b"), "")'], "String | ab"], // join.empty.delim
  [["r = Join(Null)"], "ERR 94"], // join.null.array
  [['r = Join("x", ",")'], "ERR 13"], // join.nonarray
  [["Dim a()", 'r = "[" & Join(a, ",") & "] "'], "String | [] "], // join.unsized.dynamic
  [["Dim a(1, 1)", 'r = Join(a, ",")'], "ERR 13"], // join.two.dim
  [['r = Join(Array(Null, "a"), ",")'], "ERR 13"], // join.null.element
  [['r = Join(Array(Empty, "a"), ",")'], "String | ,a"], // join.empty.element
  [['r = Join(Array(#1/2/2020#), ",")'], "String | 1/2/2020"], // join.date.element
  [['r = Join(Array(True, False), ",")'], "String | True,False"], // join.boolean.element
  [['r = Join(Array(Array(1, 2)), ",")'], "ERR 13"], // join.nested.array
  [["r = Join(Array(5))"], "String | 5"], // join.one.element
  [['r = Join(Array("a", "b"), Null)'], "ERR 94"], // join.null.delim
  [['r = "[" & Join(Array(), "") & "] "'], "String | [] "], // join.empty.array.empty.delim
  [['r = Join(Array("a"), ",", 1)'], "ERR 450"], // join.extra.arg
  [['r = Join(Array(""), ",")'], "String | "], // join.element.empty.string
  [['r = Join(Filter(Array("a", "b"), ""), "|")'], "String | a|b"], // filter.empty.search
  [['r = Join(Filter(Array(Null, "ab"), "a"), "|")'], "String | ab"], // filter.null.element
  [['r = Join(Filter(Array(Empty, "ab"), "a"), "|")'], "String | ab"], // filter.empty.element
  [['r = Join(Filter(Array("a", "b"), "a", 2), "|")'], "String | a"], // filter.include.two
  [['r = Join(Filter(Array("a", "b"), "a", 0), "|")'], "String | b"], // filter.include.zero
  [['r = Join(Filter(Array("a", "b"), "a", -1), "|")'], "String | a"], // filter.include.negone
  [['r = Join(Filter(Array("a", "b"), "a", "x"), "|")'], "ERR 13"], // filter.include.text
  [['r = Join(Filter(Array("AB", "ab"), "a", True, 0), "|")'], "String | ab"], // filter.compare.binary
  [
    ['r = Join(Filter(Array("AB", "ab"), "a", True, 1), "|")'],
    "String | AB|ab",
  ], // filter.compare.text
  [['r = Join(Filter(Array("AB", "ab"), "a", True, 2), "|")'], "ERR 5"], // filter.compare.two
  [['r = Join(Filter(Array("a"), "a", True, Null), "|")'], "ERR 94"], // filter.compare.null
  [['r = Join(Filter(Array("a"), "a", Null), "|")'], "ERR 94"], // filter.include.null
  [["Dim a(1, 1)", 'r = Join(Filter(a, ""), "|")'], "ERR 13"], // filter.two.dim
  [["Dim a()", 'r = Join(Filter(a, "a"), "|")'], "String | "], // filter.unsized.dynamic
  [['a = Filter(Array("aa"), "a")', "r = TypeName(a(0))"], "String | String"], // filter.element.type
  [['a = Filter(Array("aa"), "a")', "r = TypeName(a)"], "String | Variant()"], // filter.result.type
  [['r = Join(Filter(Array(11, 21), 1), "|")'], "String | 11|21"], // filter.numbers.value
  [['r = Join(Filter(Array(11, 21), "1"), "|")'], "String | 11|21"], // filter.numbers.text
  [['r = Join(Filter(Split("ab,cd", ","), "a"), "|")'], "String | ab"], // filter.of.split
  [['r = Join(Filter(Array("a1", "b"), 1), "|")'], "String | a1"], // filter.value.number.text
  [['r = UBound(Split(""))'], "Long | -1"], // split.empty.source.ubound
  [['a = Split("")', 'r = "[" & a(0) & "] "'], "ERR 9"], // split.empty.source.value
  [['r = UBound(Split("abc", ";"))'], "Long | 0"], // split.delim.missing
  [['a = Split("abc", "")', "r = UBound(a)"], "Long | 0"], // split.empty.delim
  [
    ['a = Split("a,b,c", ",", 1)', 'r = UBound(a) & "[" & a(0) & "] "'],
    "String | 0[a,b,c] ",
  ], // split.limit.one
  [
    ['a = Split("a,b,c", ",", 2)', 'r = a(0) & "/" & a(1) & UBound(a)'],
    "String | a/b,c1",
  ], // split.limit.two
  [['a = Split("a,b", ",", 0)', "r = UBound(a)"], "Long | -1"], // split.limit.zero
  [['a = Split("a,b", ",", -1)', "r = UBound(a)"], "Long | 1"], // split.limit.negative
  [['a = Split("a,b", ",", 10)', "r = UBound(a)"], "Long | 1"], // split.limit.big
  [['a = Split("", ",", 1)', 'r = "[" & a(0) & "] "'], "ERR 9"], // split.limit.empty.source
  [['a = Split("a,b", ",", "1")', "r = UBound(a)"], "Long | 0"], // split.limit.string
  [['a = Split("a,b", ",", Empty)', "r = UBound(a)"], "Long | -1"], // split.limit.empty
  [['a = Split("a,b", ",", Null)', 'r = "ok"'], "ERR 94"], // split.limit.null
  [['r = UBound(Split("aXbxc", "x", -1, 1))'], "Long | 2"], // split.compare.text
  [['r = UBound(Split("aXbxc", "x", -1, 0))'], "Long | 1"], // split.compare.binary
  [['r = UBound(Split("a", ",", -1, Null))'], "ERR 94"], // split.compare.null
  [['r = UBound(Split(Null, ","))'], "ERR 94"], // split.null.source
  [['r = UBound(Split("a", Null))'], "ERR 94"], // split.null.delim
  [['a = Split("1,2", 1)', "r = UBound(a)"], "Long | 1"], // split.number.delim
  [['r = Join(Split(",a", ","), "|")'], "String | |a"], // split.leading.delim
  [['r = Join(Split("a,", ","), "|")'], "String | a|"], // split.trailing.delim
  [['r = Join(Split("a,,b", ","), "|")'], "String | a||b"], // split.double.delim
  [['r = Join(Split("a::b", "::"), "|")'], "String | a|b"], // split.two.char.delim
  [['r = UBound(Split("a-b", "-:"))'], "Long | 0"], // split.two.char.missing
  [['r = Join(Split("a b c"), "|")'], "String | a|b|c"], // split.space.delim
  [['r = UBound(Split(",", ","))'], "Long | 1"], // split.only.delim
  [['r = TypeName(Split("a", ",")(0))'], "String | String"], // split.element.type
  [['a = Split("1,2", ",")', "r = a(0) + a(1)"], "String | 12"], // split.number.coerce
  [['a = Split("a,b,c", ",", -2)', "r = UBound(a)"], "ERR 5"], // split.negative.two
  [['r = UBound(Split(Array("a"), ","))'], "ERR 13"], // split.source.array
  [
    [
      "Dim a(2)",
      "a(0) = 1",
      "a(1) = 2",
      "Erase a",
      "n = 0",
      "For Each i In a",
      "n = n + 1",
      "Next",
      "r = n",
    ],
    "Integer | 3",
  ], // foreach.after.erase
  [
    ["n = 0", "For Each i In Array()", "n = n + 1", "Next", "r = n"],
    "Integer | 0",
  ], // foreach.empty.array
  [
    ["Dim a(1, 1)", "n = 0", "For Each i In a", "n = n + 1", "Next", "r = n"],
    "Integer | 4",
  ], // foreach.two.dim
  [
    ["Dim a()", "n = 0", "For Each i In a", "n = n + 1", "Next", "r = n"],
    "Integer | 0",
  ], // foreach.unsized.dynamic
  [["Dim a(1)", "FillWithNine a", "r = a(0)"], "Integer | 9"], // byref.set.element
  [
    ["Dim a(1)", "a(0) = 1", "ClearWholeArray a", "r = IsEmpty(a(0))"],
    "Boolean | True",
  ], // byref.erase
  [["Dim a()", "SizeArrayToThree a", "r = UBound(a)"], "Long | 3"], // byref.redim
  [["Dim a(2)", "r = ArrayUpperBound(a)"], "Long | 2"], // byref.ubound
  [["a = MakeTwoElementArray()", "r = UBound(a) & a(0)"], "String | 11"], // byref.make.array
  [
    ["Dim a(1)", "a(0) = 1", "ReplaceWholeArray a", "r = UBound(a) & a(0)"],
    "ERR 13",
  ], // byref.replace.whole
  [["a = Array(1)", "FillWithNine(a)", "r = a(0)"], "Integer | 1"], // byref.parenthesized
  [["Dim a(1)", "a(0) = 1", "FillWithNine(a)", "r = a(0)"], "Integer | 1"], // byref.parenthesized.fixed
  [
    ["Dim a(1)", "a(0) = 1", "FillWithNineByValue a", "r = a(0)"],
    "Integer | 1",
  ], // byref.byval.fixed
  [["a = Array(1)", "FillWithNineByValue a", "r = a(0)"], "Integer | 1"], // byref.byval.variant
  [["a = Array(1, 2)", "a = Array(3, 4, 5)", "r = UBound(a)"], "Long | 2"], // assign.plain.reassign
  [["a = Array(1, 2)", "a = 5", "r = a"], "Integer | 5"], // assign.plain.to.scalar
  [["x = 5", "x = Array(1, 2)", "r = UBound(x)"], "Long | 1"], // assign.scalar.to.array
  [["Dim x", "x = Array(1, 2)", "r = UBound(x)"], "Long | 1"], // assign.dim.plain
  [["Dim x", "ReDim x(2)", "x = Array(1, 2)", "r = UBound(x)"], "Long | 1"], // assign.after.redim
  [
    ["Dim x()", "ReDim x(2)", 'x = Split("a,b", ",")', "r = UBound(x)"],
    "ERR 13",
  ], // assign.dim.dynamic.split
  [["Dim x()", "ReDim x(2)", "ReplaceWholeArray x", "r = UBound(x)"], "ERR 13"], // assign.byref.dynamic
  [["Dim a(1)", "a = a", 'r = "ok"'], "ERR 13"], // assign.dim.fixed.same
  [["a = Array(1, 2)", "a = a", "r = UBound(a)"], "Long | 1"], // assign.plain.same
  [["Dim a()", "ReDim a(2)", "ReDim a(-1)", "r = UBound(a)"], "Long | -1"], // redim.negative.then.ubound
  [["Dim a()", "ReDim a(-1)", "r = UBound(a)"], "Long | -1"], // redim.negative.first.ubound
  [["Dim a()", "ReDim a(-1)", "r = IsArray(a)"], "Boolean | True"], // redim.negative.isarray
  [["Dim a()", "ReDim a(-1)", "r = TypeName(a)"], "String | Variant()"], // redim.negative.type
  [
    [
      "Dim a(1, 1)",
      'a(0, 0) = "a"',
      'a(1, 0) = "b"',
      'a(0, 1) = "c"',
      'a(1, 1) = "d"',
      'r = ""',
      "For Each i In a",
      "r = r & i",
      "Next",
    ],
    "String | abcd",
  ], // foreach.two.dim.order
  [
    [
      "Dim a()",
      "ReDim a(1, 2)",
      'a(0, 0) = "a"',
      'a(1, 0) = "b"',
      'a(0, 1) = "c"',
      'a(1, 1) = "d"',
      "ReDim Preserve a(1, 3)",
      'r = ""',
      "For Each i In a",
      "r = r & i",
      "Next",
    ],
    "String | abcd",
  ], // foreach.preserve.order
  [["Dim a(4)", 'a(3.5) = "x"', "r = a(4)"], "String | x"], // idx.round.half.three
  [["Dim a(3)", 'a(-0.5) = "x"', "r = a(0)"], "String | x"], // idx.round.half.negative
  [["Dim a(3)", "r = a(1E+300)"], "ERR 6"], // idx.overflow.double
  [["Dim a(1, 2)", "r = UBound(a, -0.5)"], "ERR 9"], // ub.dim.round.half.negative
  [["Dim a(1, 2)", "r = UBound(a, 2.6)"], "ERR 9"], // ub.dim.round.half.three
  [["Set o = Nothing", "r = LBound(o)"], "ERR 91"], // lb.nothing
  [["Set o = Nothing", 'r = Join(o, ",")'], "ERR 91"], // join.nothing
  [["Set o = New Holder", "r = UBound(o)"], "ERR 438"], // ub.class.instance
  [['r = Join(Filter(Array("AB", "ab"), "a", True, -1), "|")'], "ERR 5"], // filter.compare.negative
  [["Dim a(1)", "a(0) = 1", "a = 5", "r = a"], "ERR 13"], // assign.fixed.to.scalar
  [["Dim a()", "a = 5", "r = a"], "ERR 13"], // assign.dynamic.to.scalar
  [["Dim a()", "a = Empty", "r = IsEmpty(a)"], "ERR 13"], // assign.dynamic.to.empty
  [
    ["Dim a()", "ReDim a(2)", "Erase a", "a = Array(1, 2)", "r = UBound(a)"],
    "ERR 9",
  ], // assign.dynamic.after.erase
  [["Dim a(1)", "a = Array(1, 2)", 'r = "ok"'], "ERR 13"], // assign.fixed.to.array.same.size
  [["Dim a()", "ReDim Preserve a(1, 2)", "r = UBound(a, 2)"], "Long | 2"], // redim.preserve.unsized.two
  [
    [
      "Dim a()",
      "ReDim a(1, 2)",
      "Erase a",
      "ReDim Preserve a(3)",
      "r = UBound(a)",
    ],
    "Long | 3",
  ], // redim.preserve.two.to.one.after.erase
  [
    ["Dim a(1)", "a(0) = Array(1, 2)", "Erase a(0)", "r = UBound(a(0))"],
    "ERR 9",
  ], // erase.element.ubound
  [["Dim a(1)", "a(0) = 5", "Erase a(0)", "r = a(0)"], "ERR 13"], // erase.element.scalar
  [
    [
      "a = Array(1, 2)",
      "Erase a",
      "n = 0",
      "For Each i In a",
      "n = n + 1",
      "Next",
      "r = n",
    ],
    "Integer | 0",
  ], // foreach.array.result.after.erase
  [
    [
      "Dim a(2)",
      "a(0) = 1",
      "Erase a",
      "n = 0",
      "For Each i In a",
      "n = n + 1",
      "Next",
      "r = n",
    ],
    "Integer | 3",
  ], // foreach.fixed.after.erase
  [['r = UBound(Split(1, ","))'], "Long | 0"], // split.first.arg.number
  [['r = UBound(Filter(1, "a"))'], "ERR 13"], // filter.first.arg.number
  [['r = Join(5, ",")'], "ERR 13"], // join.first.arg.number
  [["a = Array(1, 2)", "ReplaceWholeArray a", "r = UBound(a)"], "Long | 1"], // byref.replace.plain
  [["x = 5", "ReplaceWholeArray x", "r = UBound(x)"], "Long | 1"], // byref.replace.scalar
  [['r = Join(Filter(Array("AB", "ab"), "a"), "|")'], "String | ab"], // filter.default.compare
  [["a = Filter(Array(11, 21), 1)", "r = TypeName(a(0))"], "String | Integer"], // filter.result.element.type
  [['a = Filter(Array(Empty), "")', "r = TypeName(a(0))"], "String | Empty"], // filter.empty.element.type
];

function answer(statements: string[]): unknown {
  const engine = new VbsEngine();
  engine.executeStatement(
    `SetLocale 1033
Dim r
${HELPERS}
Sub RunCell()
${statements.join("\n")}
End Sub
RunCell`,
  );
  const failure = engine.error;
  if (failure) return "ERR " + String(failure.number);
  if (engine._getVariable("r").type === "Null") return "Null";
  engine.executeStatement("t = TypeName(r)\ns = CStr(r)");
  return (
    String(engine._getVariable("t").value) +
    " | " +
    String(engine._getVariable("s").value)
  );
}

describe("arrays, ReDim and Erase behave like vbscript.dll", () => {
  for (const [statements, expected] of CELLS) {
    it(`${statements.join("; ")} is ${expected}`, () => {
      expect(answer(statements)).toBe(expected);
    });
  }
});
