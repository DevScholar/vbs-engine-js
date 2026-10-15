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
| Comparison (=, <>, <, >, <=, >=) | ✔️ | |
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
| For...Next | ✔️ | With Step support |
| For Each...Next | ✔️ | |
| Do...Loop | ✔️ | While/Until variants |
| While...Wend | ✔️ | |
| Select Case | ✔️ | With Case Else |
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
| New | ✔️ | |
| Property Get | ✔️ | |
| Property Let | ✔️ | |
| Property Set | ✔️ | |
| Me | ✔️ | |
| Initialize Event | ✔️ | Class_Initialize |
| Terminate Event | ✔️ | Class_Terminate |

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
| Split | ✔️ | |
| Join | ✔️ | |
| Filter | ✔️ | |
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
| CDate | ✔️ | |
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
| IsNumeric | ✔️ | |
| IsObject | ✔️ | |
| VarType | ✔️ | |
| TypeName | ✔️ | |

### Array Functions

| Function | Status | Notes |
|----------|--------|-------|
| Array | ✔️ | |
| LBound | ✔️ | |
| UBound | ✔️ | |
| Erase | ✔️ | Full implementation for fixed-size arrays |
| Filter | ✔️ | |

### Other Functions

| Function | Status | Notes |
|----------|--------|-------|
| Eval | ✔️ | |
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
| Dim | ✔️ | |
| Do...Loop | ✔️ | |
| Erase | ✔️ | |
| Execute | ✔️ | |
| ExecuteGlobal | ✔️ | |
| Exit | ✔️ | |
| For...Next | ✔️ | |
| For Each...Next | ✔️ | |
| Function | ✔️ | |
| GoTo | ✔️ | |
| If...Then...Else | ✔️ | |
| On Error | ✔️ | |
| Option Explicit | ✔️ | |
| Private | ✔️ | |
| Property Get/Let/Set | ✔️ | |
| Public | ✔️ | |
| Randomize | ✔️ | |
| ReDim | ✔️ | |
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

`CreateObject` and `GetObject` delegate to the JavaScript host's `ActiveXObject` /
`GetObject`, so their availability depends on the environment:

- **Internet Explorer** — native `ActiveXObject` is present, so COM objects
  (FileSystemObject, Word.Application, Excel.Application, WScript.Shell, etc.)
  work directly.
- **Node.js on Windows** — wire up
  [`@devscholar/node-ps1-dotnet`](https://www.npmjs.com/package/@devscholar/node-ps1-dotnet)
  to provide `ActiveXObject`/`GetObject` before running scripts (see the README's
  "COM / ActiveX Integration" section).
- **Other browsers (Chrome, Firefox, Edge, modern Node.js)** — no `ActiveXObject`,
  so `CreateObject` returns a dummy object and `GetObject` returns `Nothing`.

The concrete COM automation classes (FileSystemObject, WScript.Shell, etc.) are
not polyfilled by this engine; they are supplied by the host through the two
functions above.

| Feature | Status | Notes |
|---------|--------|-------|
| CreateObject | ✔️ | Delegates to host `ActiveXObject` |
| GetObject | ✔️ | Delegates to host `GetObject` |
| FileSystemObject | ⛔ | Requires host COM backend |
| Drive Object | ⛔ | Requires host COM backend |
| Folder Object | ⛔ | Requires host COM backend |
| File Object | ⛔ | Requires host COM backend |
| TextStream | ⛔ | Requires host COM backend |
| Dictionary Object | ⛔ | Requires host COM backend |

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
| `Collection` class | ✔️ | `New Collection`; `Add`, `Remove`, `Item`, `Count`; 1-based index or string key |
| `LongLong` (64-bit integer) | ✔️ | Backed by `BigInt`; `CLngLng()` conversion; `VarType` = 20; arithmetic stays `LongLong` |
| `Val()` and `Str()` | ✔️ | Not functions of `vbscript.dll`, which answers error 13 for both: a direct call to a name it does not know raises error 500 and a call through `Eval` raises error 13. They are kept because VB6 and VBA have them. `Val` reads a leading number and stops at the first character that is not one, and `Str` writes a leading space for the sign |
| `Type...End Type` (user-defined types) | ✔️ | Struct-like; fields typed to VB6 defaults; `Dim x As MyType` creates initialized instance |

