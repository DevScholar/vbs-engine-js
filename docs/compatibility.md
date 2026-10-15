# VBScript Engine Compatibility

This document tracks the compatibility between this VBScript engine implementation and Microsoft's Internet Explorer VBScript engine.

**Status Legend:**
- ✔️ **Implemented** - Feature is fully or partially implemented
- ❌ **Not Implemented** - Feature is not yet implemented but may be in the future
- ⛔ **Intentionally Not Implemented** - Feature will not be implemented due to platform limitations

---

## Core Language Features

### Data Types

| Feature | Status | Notes |
|---------|--------|-------|
| Empty | ✔️ | |
| Null | ✔️ | |
| Nothing | ✔️ | |
| Boolean | ✔️ | |
| Byte | ✔️ | |
| Integer | ✔️ | |
| Long | ✔️ | |
| Single | ✔️ | |
| Double | ✔️ | |
| Currency | ✔️ | |
| Date | ✔️ | |
| String | ✔️ | |
| Object | ✔️ | |
| Variant | ✔️ | |
| Error | ✔️ | |

### Variables and Constants

| Feature | Status | Notes |
|---------|--------|-------|
| Dim | ✔️ | |
| ReDim | ✔️ | With Preserve support |
| Const | ✔️ | |
| Public | ✔️ | |
| Private | ✔️ | |
| Option Explicit | ✔️ | |

### Operators

| Feature | Status | Notes |
|---------|--------|-------|
| Arithmetic (+, -, *, /) | ✔️ | |
| Integer Division (\) | ✔️ | |
| Modulo (Mod) | ✔️ | |
| Exponentiation (^) | ✔️ | |
| Concatenation (&) | ✔️ | |
| Comparison (=, <>, <, >, <=, >=) | ✔️ | Text is compared character by character and case sensitively, so `"b" = "B"` is False. A text beside a number is read as a number when it is one (`"1" = 1` is True), and beside anything else as text (`"True" = True` is True and `"1" = True` is False). A Null on either side answers Null, an object is error 91 or 438 and an array is error 13 |
| Comparison with a literal number | ✔️ | A text that is no number beside a number written out in the source is the type mismatch Windows raises: `"abc" = 1` and `"abc" = (1)` are error 13. A number that came from anywhere else is compared as text instead and answers False, which is what `"abc" = 1 * 1`, `"abc" = CInt(1)` and `"abc" = Len("a")` do |
| Is | ✔️ | |
| And | ✔️ | |
| Or | ✔️ | |
| Not | ✔️ | |
| Xor | ✔️ | |
| Eqv | ✔️ | |
| Imp | ✔️ | |
| Operator Precedence | ✔️ | |

### Control Flow Statements

| Feature | Status | Notes |
|---------|--------|-------|
| If...Then...Else | ✔️ | Single-line and multi-line |
| ElseIf | ✔️ | |
| For...Next | ✔️ | With Step support. A bound or a step that is no number is error 92, the loop's own "For loop not initialized" |
| For Each...Next | ✔️ | Leaves its loop variable empty after the loop, which is what Windows does; For...Next keeps the value that stopped it |
| Do...Loop | ✔️ | While/Until variants. A Null or an Empty condition is false rather than an error |
| While...Wend | ✔️ | |
| Select Case | ✔️ | With Case Else. A case is a plain expression or a comma list, as in VBScript: `Case 1 To 5` is a parse error there and here. It compares the way the operators do, so it reads text case sensitively and raises error 13 for a text beside a literal number |
| Exit (Do/For/Sub/Function/Property) | ✔️ | |
| With...End With | ✔️ | |
| GoTo | ✔️ | With label support |
| On Error | ✔️ | Resume Next, GoTo 0 |
| Resume | ✔️ | Partial implementation |

---

## Procedures and Classes

### Procedures

| Feature | Status | Notes |
|---------|--------|-------|
| Sub | ✔️ | |
| Function | ✔️ | |
| Call | ✔️ | |
| ByVal | ✔️ | |
| ByRef | ✔️ | Default behavior |
| Optional Parameters | ✔️ | With default values |
| ParamArray | ✔️ | |

### Classes

| Feature | Status | Notes |
|---------|--------|-------|
| Class...End Class | ✔️ | |
| New | ✔️ | `New` on a name that is no class is error 506, "Class not defined" |
| Public members | ✔️ | A field, a Property or a method reachable from anywhere |
| Private members | ✔️ | Reachable from inside the same class only, another instance of it included. From outside, reading, writing or calling one is error 438, as it is in Windows |
| A member the class does not have | ✔️ | Error 438 for a read, a write and a call alike: `first.NoSuchMember`, `first.NoSuchMember = 1` and `first.NoSuchMethod()` |
| Property Get | ✔️ | Handing back an object works (`Set Linked = inner`), including read inside an expression such as `TypeName(first.Linked)` |
| Property Let | ✔️ | |
| Property Set | ✔️ | An object stored this way comes back with `Is` intact |
| Method argument count | ✔️ | Checked the way a procedure's is: `Describe()` and `Describe("a", "b")` for `Describe(prefix)` are both error 450 |
| Me | ✔️ | |
| Initialize Event | ✔️ | Class_Initialize |
| Terminate Event | ✔️ | Class_Terminate, which runs when the last reference to an instance goes |
| Assigning an object | ✔️ | Needs `Set`. Without it the assignment is error 438 (`x = New Holder`, `x = Maker()`, `x = held`, `x = items(0)`), and `x = Nothing` is error 91. Passing an object as an argument needs no `Set` |
| An object in an array | ✔️ | `Set items(0) = New Holder` works, and `items(0).Name` reads and writes through it |
| Reading an object as text | ✔️ | Error 438: an object of a class has no text of its own, so `CStr(anInstance)` is 438 |

---

## Built-in Functions

### String Functions

| Function | Status | Notes |
|----------|--------|-------|
| Len | ✔️ | |
| Left | ✔️ | |
| Right | ✔️ | |
| Mid | ✔️ | |
| InStr | ✔️ | |
| InStrRev | ✔️ | |
| LCase | ✔️ | |
| UCase | ✔️ | |
| LTrim | ✔️ | |
| RTrim | ✔️ | |
| Trim | ✔️ | |
| Replace | ✔️ | |
| StrReverse | ✔️ | |
| Space | ✔️ | |
| String | ✔️ | Create repeated character string |
| Asc | ✔️ | |
| AscW | ✔️ | |
| Chr | ✔️ | |
| ChrW | ✔️ | |
| StrComp | ✔️ | |
| Split | ✔️ | Reads its text as text, so `Split(1, ",")` makes one part, and reads the delimiter, the limit and the compare the way Windows does: a limit below -1 is error 5, 0 makes no parts, and an empty text makes none either. An array or an object as the text is error 13 and 91 or 438 |
| Join | ✔️ | Reads one dimensional arrays only, so a two dimensional one is error 13, and refuses a Null or an array element with error 13 where an Empty element is no text at all. `Join(Null)` is error 94, `Join(Nothing)` is 91, a list that is no array is 13 |
| Filter | ✔️ | Reads its compare argument (0 binary, 1 text, anything else error 5), returns the elements it found with their own types, and reads a Null element as no text at all. A two dimensional array or a list that is no array is error 13 |
| Format | ✔️ | Custom format strings |
| LSet | ✔️ | |
| RSet | ✔️ | |

### Mathematical Functions

| Function | Status | Notes |
|----------|--------|-------|
| Abs | ✔️ | |
| Sgn | ✔️ | |
| Sqr | ✔️ | |
| Int | ✔️ | |
| Fix | ✔️ | |
| Round | ✔️ | |
| Atn | ✔️ | |
| Cos | ✔️ | |
| Sin | ✔️ | |
| Tan | ✔️ | |
| Exp | ✔️ | |
| Log | ✔️ | Natural logarithm |
| Rnd | ✔️ | |
| Randomize | ✔️ | |
| Hex | ✔️ | |
| Oct | ✔️ | |

### Date and Time Functions

| Function | Status | Notes |
|----------|--------|-------|
| Now | ✔️ | |
| Date | ✔️ | |
| Time | ✔️ | |
| Year | ✔️ | |
| Month | ✔️ | |
| Day | ✔️ | |
| Weekday | ✔️ | |
| Hour | ✔️ | |
| Minute | ✔️ | |
| Second | ✔️ | |
| DateAdd | ✔️ | |
| DateDiff | ✔️ | |
| DatePart | ✔️ | |
| DateSerial | ✔️ | |
| TimeSerial | ✔️ | |
| DateValue | ✔️ | |
| TimeValue | ✔️ | |
| MonthName | ✔️ | |
| WeekdayName | ✔️ | |
| Timer | ✔️ | |

### Conversion Functions

| Function | Status | Notes |
|----------|--------|-------|
| CBool | ✔️ | |
| CByte | ✔️ | |
| CCur | ✔️ | |
| CDate | ✔️ | Reads a written date with the locale's own order and separators, from the Windows short date picture, the spaced dotted form the picture itself writes included: `CDate("2026. 05. 01.")` is 1 May 2026 under 1038 and `CDate("1. 5. 2026")` is a date under 1031, 1038, 1051, 1054 and 1060. Measured across the 132 locales that answer |
| CDbl | ✔️ | |
| CInt | ✔️ | |
| CLng | ✔️ | |
| CSng | ✔️ | |
| CStr | ✔️ | |
| CVar | ✔️ | |
| CVErr | ✔️ | |
| FormatNumber | ✔️ | Reads the Windows NLS number table, so the fraction mark, the group marks and the fraction count are the locale's own |
| FormatCurrency | ✔️ | The symbol, its position, the fraction count and the shape of a negative amount come from the locale's row in the same table |
| FormatPercent | ✔️ | Reads the Windows NLS number table, and writes the percent sign the way Windows does |
| FormatDateTime | ✔️ | Reads the Windows NLS date table, so it follows Windows and not the host |

### Inspection Functions

| Function | Status | Notes |
|----------|--------|-------|
| IsArray | ✔️ | |
| IsDate | ✔️ | |
| IsEmpty | ✔️ | |
| IsNull | ✔️ | |
| IsNumeric | ✔️ | Reads a string the way the conversions do, with the locale's own marks, so `IsNumeric("50%")` is `False` and `IsNumeric("1 234")` is `False` under 1031 but `True` under 1026, where a space is the group mark |
| IsObject | ✔️ | |
| VarType | ✔️ | |
| TypeName | ✔️ | |

### Array Functions

| Function | Status | Notes |
|----------|--------|-------|
| Array | ✔️ | The array it makes is dynamic, so `Erase` frees it and `ReDim` may size it again |
| LBound | ✔️ | The dimension is a whole number rounded to the nearest one, a tie to the even neighbour, so `LBound(a, "2")` is the second dimension and `LBound(a, 1.5)` is the second as well. A dimension outside the array is error 9, a Null is error 94, no array at all is 13, and an object is 91 for a Nothing and 438 for any other |
| UBound | ✔️ | Reads its dimension the way LBound does, and a name Dim declared as an array still answers here after `Erase`, which is error 9 for a dynamic one |
| Erase | ✔️ | Takes one name: `Erase a` frees what a holds, `Erase a(0)` frees what the element holds, and `Erase(a)`, `Erase Array(1, 2)` and `Erase h.Items` are error 13 while `Erase a, b` and `Erase` alone are 450. A fixed array keeps its bounds and every element becomes Empty, and a dynamic one, `Array`, `Split` and `Filter` results included, is freed |
| Filter | ✔️ | See String Functions |

### Other Functions

| Function | Status | Notes |
|----------|--------|-------|
| Eval | ✔️ | Reads its text as an expression, so a bare name is the value it holds or the procedure it names rather than the call a bare name statement is: `Eval("number")` is 5 and `Eval("NoSuchFunction")` is 0, where those names written as statements are both error 13 |
| Execute | ✔️ | Dynamic code execution in current scope |
| ExecuteGlobal | ✔️ | Dynamic code execution in global scope |
| MsgBox | ✔️ | Browser simulation using alert/confirm |
| InputBox | ✔️ | Browser simulation using prompt |
| GetRef | ✔️ | |
| GetLocale | ✔️ | Uses Intl API to detect browser locale |
| SetLocale | ✔️ | Sets locale, returns LCID |
| CreateObject | ✔️ | Delegates to the host's `ActiveXObject`; see COM Objects below |
| GetObject | ✔️ | Delegates to the host's `GetObject`; see COM Objects below |
| LoadPicture | ✔️ | Returns stub IPictureDisp object |
| RGB | ✔️ | Returns RGB color value |
| QBColor | ✔️ | Returns legacy color value |
| ScriptEngine | ✔️ | Returns "VBScript" |
| ScriptEngineMajorVersion | ✔️ | Returns 10 |
| ScriptEngineMinorVersion | ✔️ | Returns 8 |
| ScriptEngineBuildVersion | ✔️ | Returns 16384 |

---

## Objects

### Err Object

| Property/Method | Status | Notes |
|-----------------|--------|-------|
| Number | ✔️ | |
| Description | ✔️ | |
| Source | ✔️ | |
| HelpFile | ✔️ | Read/write `String`, default `""` (no help source); only set when `Err.Raise` passes a helpfile argument. No `.hlp` loading — WinHelp was removed from modern Windows |
| HelpContext | ✔️ | Read/write `Long`, default `0` (no help source); only set when `Err.Raise` passes a helpcontext argument. No `.hlp` loading |
| LastDllError | ✔️ | Returns 0 (no DLL loader wired) |
| Clear | ✔️ | |
| Raise | ✔️ | Supports `number, source, description, helpfile, helpcontext` |

### RegExp Object

| Feature | Status | Notes |
|---------|--------|-------|
| RegExp Object | ✔️ | |
| Pattern | ✔️ | |
| Global | ✔️ | |
| IgnoreCase | ✔️ | |
| Multiline | ✔️ | |
| Execute Method | ✔️ | |
| Test Method | ✔️ | |
| Replace Method | ✔️ | |
| Match Object | ✔️ | FirstIndex, Length, Value |
| Matches Collection | ✔️ | Count, Item |
| SubMatches Collection | ✔️ | |

### Class Object

| Feature | Status | Notes |
|---------|--------|-------|
| Class definition | ✔️ | |
| Public members | ✔️ | |
| Private members | ✔️ | |
| Property procedures | ✔️ | |
| Methods | ✔️ | |

### VBArray Object (JavaScript-side)

VBArray is a JavaScript-side object for accessing VBScript arrays from JavaScript.
It provides backward compatibility with early Internet Explorer, which required
wrapping VBScript arrays before accessing them from JavaScript.

| Feature | Status | Notes |
|---------|--------|-------|
| VBArray constructor | ✔️ | Wraps internal VbArray for JS access |
| dimensions() | ✔️ | Returns number of dimensions |
| lbound() | ✔️ | Returns lower bound of a dimension |
| ubound() | ✔️ | Returns upper bound of a dimension |
| getItem() | ✔️ | Gets value at specified indices |
| toArray() | ✔️ | Converts to JavaScript array |

**Configuration:** The `injectVBArrayToGlobalThis` option (default: `true`) controls
whether VBArray is available as `globalThis.VBArray`. When enabled, legacy code
using `new VBArray(vbArray)` will work. Modern code can access VBScript arrays
directly without wrapping.

---

## Statements

| Statement | Status | Notes |
|-----------|--------|-------|
| Call | ✔️ | |
| Class | ✔️ | |
| Const | ✔️ | |
| Dim | ✔️ | The engine reads the bound of an array at run time, so `Dim a(1.5)` and `Dim a(n)` run here. Windows refuses to compile a bound that is not a whole number written out in the source: `Dim a(1.5)`, `Dim a(-1)`, `Dim a(1 + 2)` and `Dim a(namedConstant)` are all compile errors there, and only a literal such as `Dim a(3)` is accepted |
| Do...Loop | ✔️ | |
| Erase | ✔️ | Takes one name, as Windows does: `Erase a` frees what a holds, `Erase a(0)` frees what the element holds, and `Erase(a)`, `Erase Array(1, 2)` and `Erase h.Items` are error 13 while `Erase a, b` and `Erase` alone are 450 |
| Execute | ✔️ | |
| ExecuteGlobal | ✔️ | |
| Exit | ✔️ | |
| For...Next | ✔️ | |
| For Each...Next | ✔️ | |
| Function | ✔️ | |
| GoTo | ✔️ | |
| If...Then...Else | ✔️ | |
| On Error | ✔️ | |
| Option Explicit | ✔️ | An undeclared name is error 500 in every position, where without it a call on such a name indexes the Empty it stands for and is error 13 |
| A statement that is just a name | ✔️ | A whole statement naming something is a call on it with no arguments, as it is in Windows: a name holding a number, a text, an Empty, a Null or an array is error 13, a name holding an object is 438, and a procedure is called, with 450 when it needs an argument. `arr(9)` alone is error 13 where `x = arr(9)` is error 9, a member holding a field is 5 with no arguments and 438 with them, and a Property Get reads and drops its value. A literal as a statement is a compile error in Windows and is not accepted here either |
| Private | ✔️ | |
| Property Get/Let/Set | ✔️ | |
| Public | ✔️ | |
| Randomize | ✔️ | |
| ReDim | ✔️ | A bound is rounded to the nearest whole number, a tie to the even neighbour, and `ReDim a(-1)` makes an array of no elements whose UBound is -1. `ReDim` on a fixed array is error 10, `ReDim Preserve` may change only the last dimension of an array that already has one, and changing any other dimension or how many there are is error 9. An array with no size yet, from `Dim a()` or from `Erase`, takes any shape Preserve asks for |
| Rem | ✔️ | |
| Select Case | ✔️ | |
| Set | ✔️ | |
| Sub | ✔️ | |
| While...Wend | ✔️ | |
| With | ✔️ | |

---

## Constants

### String Constants

| Constant | Status | Notes |
|----------|--------|-------|
| vbCr | ✔️ | |
| vbCrLf | ✔️ | |
| vbFormFeed | ✔️ | |
| vbLf | ✔️ | |
| vbNewLine | ✔️ | |
| vbNullChar | ✔️ | |
| vbNullString | ✔️ | |
| vbTab | ✔️ | |
| vbVerticalTab | ✔️ | |

### VarType Constants

| Constant | Status | Notes |
|----------|--------|-------|
| vbEmpty | ✔️ | |
| vbNull | ✔️ | |
| vbInteger | ✔️ | |
| vbLong | ✔️ | |
| vbSingle | ✔️ | |
| vbDouble | ✔️ | |
| vbCurrency | ✔️ | |
| vbDate | ✔️ | |
| vbString | ✔️ | |
| vbObject | ✔️ | |
| vbError | ✔️ | |
| vbBoolean | ✔️ | |
| vbVariant | ✔️ | |
| vbByte | ✔️ | |
| vbArray | ✔️ | |

### MsgBox Constants

| Constant | Status | Notes |
|----------|--------|-------|
| vbOKOnly | ✔️ | |
| vbOKCancel | ✔️ | |
| vbAbortRetryIgnore | ✔️ | |
| vbYesNoCancel | ✔️ | |
| vbYesNo | ✔️ | |
| vbRetryCancel | ✔️ | |
| vbCritical | ✔️ | |
| vbQuestion | ✔️ | |
| vbExclamation | ✔️ | |
| vbInformation | ✔️ | |
| vbDefaultButton1/2/3 | ✔️ | |
| vbOK | ✔️ | |
| vbCancel | ✔️ | |
| vbAbort | ✔️ | |
| vbRetry | ✔️ | |
| vbIgnore | ✔️ | |
| vbYes | ✔️ | |
| vbNo | ✔️ | |

### Date Constants

| Constant | Status | Notes |
|----------|--------|-------|
| vbSunday - vbSaturday | ✔️ | |
| vbUseSystemDayOfWeek | ✔️ | |
| vbFirstJan1 | ✔️ | |
| vbFirstFourDays | ✔️ | |
| vbFirstFullWeek | ✔️ | |
| vbGeneralDate | ✔️ | |
| vbLongDate | ✔️ | |
| vbShortDate | ✔️ | |
| vbLongTime | ✔️ | |
| vbShortTime | ✔️ | |

### Comparison Constants

| Constant | Status | Notes |
|----------|--------|-------|
| vbBinaryCompare | ✔️ | |
| vbTextCompare | ✔️ | |
| vbDatabaseCompare | ✔️ | Falls back to locale-aware comparison |

### Other Constants

| Constant | Status | Notes |
|----------|--------|-------|
| vbObjectError | ✔️ | |

---

## COM Objects

`CreateObject` and `GetObject` hand their call to the JavaScript host's `ActiveXObject` and
`GetObject`, so no COM class lives in this engine:

- **Internet Explorer** has a native `ActiveXObject`, so every COM class works there as it
  always did.
- **Node.js on Windows** takes them from
  [`@devscholar/node-ps1-dotnet`](https://www.npmjs.com/package/@devscholar/node-ps1-dotnet),
  wired up before a script runs (see the README's "COM / ActiveX Integration" section).
- **Other hosts** have no `ActiveXObject`, so `CreateObject` answers a placeholder object and
  `GetObject` answers `Nothing`.

This document therefore lists the two functions and no COM class. `Scripting.FileSystemObject`,
`Scripting.Dictionary`, `WScript.Shell`, `Word.Application` and the rest belong to the host,
and which of them answer depends on the host rather than on this engine. A table of them here
would read as if each one had been given its own handling.

| Function | Status | Notes |
|----------|--------|-------|
| CreateObject | ✔️ | Hands the call to the host's `ActiveXObject` |
| GetObject | ✔️ | Hands the call to the host's `GetObject` |

---

## WSH Host (`WScript`) and `run-script`

The engine itself does not inject a `WScript` object. Like `ActiveXObject`/`GetObject`,
it is supplied by the host through `globalThis` — present if the host provides it,
absent otherwise. The `run-script` CLI (this package's `bin`) provides a
cscript/wscript-style host on top of this, so scripts written for Windows Script
Host run with minimal changes.

### `WScript` host object

Injected by the host as `globalThis.WScript`, exposing the standard WSH surface:

| Member | Status | Notes |
|--------|--------|-------|
| `WScript.Echo` | ✔️ | Writes to stdout in cscript mode; joined with spaces |
| `WScript.Quit n` | ✔️ | Throws a `WScriptQuitSignal` to unwind; `run-script` maps it to the process exit code |
| `WScript.Sleep ms` | ✔️ | Synchronous sleep |
| `WScript.Arguments` | ✔️ | `Count`, `(n)`, `Item(n)`, `Named`, `Unnamed` |
| `WScript.Arguments.Named("x")` / `.Exists("x")` | ✔️ | Named args from `/name:value` and `.wsf` `<runtime>` |
| `WScript.StdOut` / `StdErr` | ✔️ | `Write` / `WriteLine` / `WriteBlankLines` |
| `WScript.StdIn` | ✔️ | `ReadLine` / `ReadAll` (synchronous) |
| `WScript.ScriptName` / `ScriptFullName` | ✔️ | |
| `WScript.FullName` / `Path` / `Version` / `BuildVersion` | ✔️ | Host paths and a fixed `5.812` version string |
| `WScript.Interactive` | ✔️ | `false` under `//B` |
| `WScript.CreateObject` / `GetObject` | ✔️ | Delegate to the host's `ActiveXObject`/`GetObject` |

`WScript` is implemented as a plain object exposing the engine's `VbObjectValueData`
protocol (`getProperty`/`hasMethod`/`getMethod`). Default-property invocation
(`WScript.Arguments(0)`, `WScript.Arguments.Named("x")`) is supported by the
engine's member-call dispatch.

### `run-script` CLI

```
run-script script.vbs [//options] [script-args...]
run-script script.wsf //job:Job1 [//options] [script-args...]
```

| Option | Meaning |
|--------|---------|
| `//nologo` | Suppress the banner |
| `//job:<id>` | Select a `<job>` in a `.wsf` |
| `//B` / `//I` | Batch (non-interactive) / interactive mode |
| `//E:<engine>` | Target engine (VBScript only for now) |
| `--gui-mode=cscript\|wscript` | Drives `WScript.Echo` (console vs GUI) |

Named script arguments use WSH's `/name:value` form; bare words are positional.
`WScript.Quit n` sets the process exit code.

### Windows Script File (`.wsf`)

`run-script` parses `.wsf` files (via `fast-xml-parser`), supporting:

| Element | Status | Notes |
|---------|--------|-------|
| `<package>` / `<job>` | ✔️ | Multiple jobs; single job without `<package>` |
| `<script language="VBScript">` | ✔️ | Inline text, CDATA, and `src="..."` |
| `<runtime>` | ✔️ | `<named>`, `<unnamed>`, `<description>`, `<example>`; feeds `WScript.Arguments` and `ShowUsage` |
| `<object>` / `<reference>` / `<resource>` | ⚠️ | `<object id progid>` is hoisted into an implicit `Set <id> = CreateObject(progid)` before scripts run; `<reference>` is parsed but its typelib/constant loading is not yet performed |
| `<script language="JScript">` | ✔️ | Executed via the host JS engine in the global context (`vm.runInThisContext`); top-level `var`/`function` become `globalThis` properties, so VBScript and JScript share one global scope |

---

## Events

| Event | Status | Notes |
|-------|--------|-------|
| Initialize | ✔️ | Class_Initialize |
| Terminate | ✔️ | Class_Terminate |

---

## Non-Standard Extensions

These features are not part of the VBScript specification and do not exist in
Internet Explorer's VBScript engine. They are specific to this implementation
and are intended for Node.js/JavaScript interoperability.

| Feature | Status | Notes |
|---------|--------|-------|
| `New` with JavaScript constructors | ✔️ | Resolves dotted names via `globalThis`; `globalThis[A][B][…]` path; `Reflect.construct` |
| `GetRef` callbacks to JavaScript | ✔️ | Wraps VBScript procedures as real JS callables for use in JS event APIs |
| Property access on JavaScript Proxies | ✔️ | Dot-notation get/set via Proxy traps; engine protocol methods only used as own properties |
| Host object access via `globalThis` | ✔️ | All `globalThis` properties available as VBScript variables; case-insensitive |

---

## VB6 / VBA7 Extensions

| Feature | Status | Notes |
|---------|--------|-------|
| `As Type` annotations | ✔️ | `Dim x As Integer` initializes to typed default; applies to variables, parameters, return types |
| `Enum...End Enum` | ✔️ | Global read-only constants; auto-increment from 0 or last explicit value; `Public`/`Private` |
| `Collection` class | ✔️ | `New Collection`; `Add`, `Remove`, `Item`, `Count`; 1-based index or string key. Not a class of `vbscript.dll`, which answers a runtime error for `New Collection`, so there is nothing to compare it against; it is kept because VB6 and VBA have it |
| `Case Is` in Select Case | ✔️ | `Select Case x : Case Is > 5` runs here and is a compile error in `vbscript.dll`, which allows only a plain expression or a comma list. Kept because VB6 and VBA have it. A range (`Case 1 To 5`) is not accepted either way |
| `LongLong` (64-bit integer) | ✔️ | Backed by `BigInt`; `CLngLng()` conversion; `VarType` = 20; arithmetic stays `LongLong` |
| `Val()` and `Str()` | ✔️ | Not functions of `vbscript.dll`, which answers error 13 for both: a direct call to a name it does not know raises error 500 and a call through `Eval` raises error 13. They are kept because VB6 and VBA have them. `Val` reads a leading number and stops at the first character that is not one, and `Str` writes a leading space for the sign |
| `Type...End Type` (user-defined types) | ✔️ | Struct-like; fields typed to VB6 defaults; `Dim x As MyType` creates initialized instance |

