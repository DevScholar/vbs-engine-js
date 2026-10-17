import { describe, it, expect } from "vitest";
import { VbsEngine } from "../index.ts";

// Classes and objects, measured against a fresh cscript under 1033 with the class probes in the probe
// directory: 42 cells over one script's own classes, and 5 more over a property that hands back an object.
// Every one of them agrees with the engine. What the measurement settled:
//
//   * Private inside a class means inside that class. A member declared Private is error 438 from script
//     level, both read and written, and stays reachable from the class's own methods.
//   * A member a class does not have is 438 as well, whether it is read, written or called:
//     result = first.NoSuchMember, first.NoSuchMember = 1 and first.NoSuchMethod() are all 438.
//   * A method's argument count is checked the way a procedure's is, so Describe() and Describe("a", "b")
//     for Describe(prefix) are both 450.
//   * An object needs Set, and without it the assignment is 438 for every shape a class object can be
//     written in: x = New Holder, x = Maker(), x = held and x = items(0). x = Nothing is 91 instead.
//   * New on a name that is no class is 506, "Class not defined".
//   * An object of a class has no text of its own, so CStr(anInstance) is 438.
//   * An object may live in an array: Set items(0) = New Holder works, and items(0).Name reads and writes
//     through it. Before this the engine refused the Set with "Invalid assignment target: CallExpression".
//   * A Property Get that hands back an object works in an expression. Before this the engine recursed until
//     the stack ran out, since reading the property asked for the property again.
//
// The cells share one script's state, in the probe's own order, so the ones that change an instance see the
// same instance the probe saw.
const CELLS: ReadonlyArray<readonly [string, string, string]> = [
  ["typename.instance", "result = TypeName(first)", "Holder"],
  ["typename.plain", "result = TypeName(second)", "Plain"],
  ["isobject", "result = CStr(IsObject(first))", "True"],
  ["vartype", "result = CStr(VarType(first))", "9"],
  ["isnotnothing", "result = CStr(first Is Nothing)", "False"],
  [
    "nothing.isnothing",
    "Set heldNothing = Nothing\nresult = CStr(heldNothing Is Nothing)",
    "True",
  ],
  [
    "is.sameReference",
    "Set sameOne = first\nresult = CStr(first Is sameOne)",
    "True",
  ],
  [
    "is.differentInstance",
    "Set otherOne = New Holder\nresult = CStr(first Is otherOne)",
    "False",
  ],
  [
    "equality.objects",
    "Set otherOne = New Holder\nresult = CStr(first = otherOne)",
    "ERR 438",
  ],
  ["public.field", "result = first.Name", "one"],
  [
    "public.field.write",
    'first.Name = "changed"\nresult = first.Name',
    "changed",
  ],
  ["property.get", "result = CStr(first.Count)", "1"],
  ["property.let", "first.Count = 42\nresult = CStr(first.Count)", "42"],
  ["property.readonly", "result = CStr(first.Total)", "49"],
  [
    "property.set.object",
    'Set otherOne = New Holder\notherOne.Name = "linked"\nSet first.Linked = otherOne\nresult = CStr(first.Linked Is otherOne)',
    "True",
  ],
  [
    "property.getWithoutSet",
    "Set otherOne = New Holder\nSet first.Linked = otherOne\nresult = first.Linked",
    "ERR 438",
  ],
  ["private.read", "result = first.hiddenCount", "ERR 438"],
  [
    "private.write",
    "first.hiddenCount = 5\nresult = CStr(first.Count)",
    "ERR 438",
  ],
  ["method.return", 'result = first.Describe("d")', "d:changed:42"],
  ["method.args", "result = CStr(first.Sum(2, 3))", "5"],
  ["method.subChangesState", "first.Bump 5\nresult = CStr(first.Count)", "47"],
  ["method.subNoParens", 'first.Bump 5\nresult = "ran"', "ran"],
  [
    "method.subWithParensObject",
    "Set otherOne = New Holder\nfirst.Bump (5)\nresult = CStr(first.Count)",
    "57",
  ],
  ["method.me", "result = CStr(first.SelfValue() Is first)", "True"],
  [
    "method.factory",
    'Set made = first.MakeOther()\nresult = TypeName(made) & "/" & made.Name & "/" & CStr(made Is first)',
    "Holder/made/False",
  ],
  ["method.missing", "result = first.NoSuchMethod()", "ERR 438"],
  ["member.missing", "result = first.NoSuchMember", "ERR 438"],
  ["method.tooFewArgs", "result = first.Describe()", "ERR 450"],
  ["method.tooManyArgs", 'result = first.Describe("a", "b")', "ERR 450"],
  ["new.withoutSet", "result = New Holder", "ERR 438"],
  ["new.unknownClass", "result = New NoSuchClass", "ERR 506"],
  [
    "nothing.member",
    "Set heldNothing = Nothing\nresult = heldNothing.Name",
    "ERR 424",
  ],
  [
    "nothing.method",
    'Set heldNothing = Nothing\nresult = heldNothing.Describe("d")',
    "ERR 424",
  ],
  [
    "nothing.property",
    "Set heldNothing = Nothing\nresult = CStr(heldNothing.Count)",
    "ERR 424",
  ],
  [
    "plain.property",
    'Set plainOne = New Plain\nplainOne.Value = "v"\nresult = plainOne.Value',
    "v",
  ],
  [
    "object.argument",
    'Sub Taker(item)\nitem.Name = "taken"\nEnd Sub\nTaker first\nresult = first.Name',
    "taken",
  ],
  [
    "object.argumentByVal",
    "Function Checker(ByVal item)\nChecker = TypeName(item)\nEnd Function\nresult = Checker(first)",
    "Holder",
  ],
  [
    "object.functionReturn",
    "Function Maker()\nSet Maker = New Plain\nEnd Function\nSet made = Maker()\nresult = TypeName(made)",
    "Plain",
  ],
  [
    "object.functionReturnPlain",
    "Function Maker()\nSet Maker = New Plain\nEnd Function\nresult = Maker()",
    "ERR 438",
  ],
  [
    "with.block",
    'With first\n.Name = "withName"\n.Bump 1\nresult = .Name & "/" & CStr(.Count)\nEnd With',
    "withName/58",
  ],
  [
    "array.ofObjects",
    'Dim items(1)\nSet items(0) = New Holder\nSet items(1) = New Holder\nitems(0).Name = "zero"\nresult = items(0).Name & "/" & TypeName(items(1))',
    "zero/Holder",
  ],
  ["object.invariantToString", "result = CStr(first)", "ERR 438"],
];

const CLASSES = `Class Holder
  Public Name
  Private hiddenCount
  Private linkedObject

  Private Sub Class_Initialize()
    hiddenCount = 1
    trace = trace & "init(" & TypeName(Me) & ");"
  End Sub

  Private Sub Class_Terminate()
    trace = trace & "term(" & Name & ");"
  End Sub

  Public Property Get Count
    Count = hiddenCount
  End Property

  Public Property Let Count(value)
    hiddenCount = value
  End Property

  Public Property Get Total
    Total = hiddenCount + 7
  End Property

  Public Property Set Linked(other)
    Set linkedObject = other
  End Property

  Public Property Get Linked
    Set Linked = linkedObject
  End Property

  Public Function Describe(prefix)
    Describe = prefix & ":" & Name & ":" & hiddenCount
  End Function

  Public Sub Bump(amount)
    hiddenCount = hiddenCount + amount
  End Sub

  Public Function MakeOther()
    Dim fresh
    Set fresh = New Holder
    fresh.Name = "made"
    Set MakeOther = fresh
  End Function

  Public Function SelfValue()
    Set SelfValue = Me
  End Function

  Public Function Sum(firstValue, secondValue)
    Sum = firstValue + secondValue
  End Function
End Class

Class Plain
  Public Value
End Class

`;

/** Runs every cell in order in one engine, the way the probe's single script does. */
function runAll(): string[] {
  const engine = new VbsEngine();
  engine.executeStatement(`${CLASSES}
On Error Resume Next
Dim result, trace, first, second, heldNothing
trace = ""
Set first = New Holder
first.Name = "one"
Set second = New Plain
`);
  const answers: string[] = [];
  for (const [, code] of CELLS) {
    engine.executeStatement(`Err.Clear
${code}
number = CStr(Err.Number)
answer = result`);
    const failure = engine.error;
    const number =
      failure && failure.number !== 0
        ? failure.number
        : Number(engine._getVariable("number").value);
    const answer = engine._getVariable("answer");
    answers.push(
      number !== 0
        ? `ERR ${number}`
        : answer === undefined
          ? "(unset)"
          : String(answer.value),
    );
  }
  return answers;
}

describe("classes and objects behave like vbscript.dll", () => {
  const answers = runAll();
  CELLS.forEach(([label, , expected], index) => {
    it(`${label} is ${expected}`, () => {
      expect(answers[index]).toBe(expected);
    });
  });
});

// A Property Get that hands back an object, read back in an expression. Measured on a fresh cscript.
const PROPERTY_CELLS: ReadonlyArray<readonly [string, string, string]> = [
  [
    "readWithSet",
    "Set readBack = first.Linked\nresult = CStr(readBack Is otherOne)",
    "True",
  ],
  ["typenameInExpression", "result = TypeName(first.Linked)", "Holder"],
  ["isInExpression", "result = CStr(first.Linked Is otherOne)", "True"],
  ["memberOfProperty", "result = first.Linked.Name", "linked"],
];

const PROPERTY_CLASS = `Class Holder
  Private inner
  Public Name
  Public Property Set Linked(item)
    Set inner = item
  End Property
  Public Property Get Linked
    Set Linked = inner
  End Property
End Class`;

describe("a property that holds an object", () => {
  const answers = runPropertyCells();
  PROPERTY_CELLS.forEach(([label, , expected], index) => {
    it(`${label} is ${expected}`, () => {
      expect(answers[index]).toBe(expected);
    });
  });
});

function runPropertyCells(): string[] {
  const engine = new VbsEngine();
  engine.executeStatement(`${PROPERTY_CLASS}
On Error Resume Next
Dim first, otherOne, result
Set first = New Holder
Set otherOne = New Holder
otherOne.Name = "linked"
Set first.Linked = otherOne
`);
  const answers: string[] = [];
  for (const [, code] of PROPERTY_CELLS) {
    engine.executeStatement(`Err.Clear
${code}
number = CStr(Err.Number)
answer = result`);
    const failure = engine.error;
    const number =
      failure && failure.number !== 0
        ? failure.number
        : Number(engine._getVariable("number").value);
    const answer = engine._getVariable("answer");
    answers.push(
      number !== 0
        ? `ERR ${number}`
        : answer === undefined
          ? "(unset)"
          : String(answer.value),
    );
  }
  return answers;
}
