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
| Val | ✔️ | |
| Str | ✔️ | |
| FormatNumber | ✔️ | Uses Intl API with locale support |
| FormatCurrency | ✔️ | Auto-detects currency from locale |
| FormatPercent | ✔️ | Uses Intl API with locale support |
| FormatDateTime | ✔️ | Uses Intl API with locale support |

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
| CreateObject | ⛔ | COM objects unavailable in browser |
| GetObject | ⛔ | COM objects unavailable in browser |
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
| Clear | ✔️ | |
| Raise | ✔️ | |
| HelpContext | ⛔ | Windows Help system |
| HelpFile | ⛔ | Windows Help system |

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

All COM objects features are intentionally not implemented due to browser security restrictions.

These functions rely on the JavaScript environment's native `ActiveXObject` support. In Internet Explorer, this allows creating COM objects including FileSystemObject, Word.Application, Excel.Application, etc. However, modern JavaScript environments (Chrome, Firefox, Edge, Node.js) do not support ActiveXObject. The author of this project may create a separate project in the future to implement polyfills for common COM objects (such as FileSystemObject, WScript.Shell, etc.) in modern JavaScript environments. This is beyond the scope of this VBScript engine project.

| Feature | Status | Notes |
|---------|--------|-------|
| FileSystemObject | ⛔ | Browser security |
| Drive Object | ⛔ | Browser security |
| Folder Object | ⛔ | Browser security |
| File Object | ⛔ | Browser security |
| TextStream | ⛔ | Browser security |
| Dictionary Object | ⛔ | Browser security |

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

### `New` with JavaScript Constructors

Standard VBScript's `New` keyword is restricted to user-defined VBScript classes
(`Class...End Class`). This engine extends `New` to also accept dotted names that
resolve to JavaScript constructors on `globalThis`:

```vbscript
' Requires globalThis.Forms = System.Windows.Forms
Set form   = New Forms.Form
Set label  = New Forms.Label
Set font   = New Drawing.Font("Arial", 12)
Set pt     = New Drawing.Point(10, 20)
```

Resolution order:
1. VBScript class registry (existing behavior — backward-compatible)
2. `globalThis[Name]` for a simple name, or `globalThis[A][B][…]` for a dotted
   path — calls `Reflect.construct(ctor, args)` on the resolved constructor

This mirrors VB6/VBA early-binding behavior where `New` can instantiate any
registered COM class.

### `GetRef` Callbacks to JavaScript

Standard VBScript's `GetRef` returns a procedure reference usable only within
VBScript. This engine additionally wraps the reference into a real JavaScript
callable, so it can be passed to JavaScript APIs that expect a callback:

```vbscript
button.add_Click GetRef("OnButtonClick")

Sub OnButtonClick(sender, e)
  ' ...
End Sub
```

### Property Access and Assignment on JavaScript Proxies

Standard VBScript COM objects implement a fixed interface for property
get/set. This engine allows VBScript to access and assign properties on any
JavaScript `Proxy` object (including node-ps1-dotnet .NET proxies) using
natural dot-notation syntax:

```vbscript
form.Text = "Hello"       ' sets property via Proxy set trap
x = form.Width            ' reads property via Proxy get trap
```

VBScript-engine protocol methods (`getProperty`, `setProperty`, etc.) are
only used when they exist as **own properties** of the object, so arbitrary
Proxy objects are treated as plain objects instead.

### Host Object Access via `globalThis`

Standard VBScript only exposes objects explicitly registered by the host
(e.g. `window`, `document` in IE). This engine automatically makes all
properties on JavaScript's `globalThis` available to VBScript as variables,
with case-insensitive name resolution. This includes non-enumerable globals
such as `console`, `Math`, `JSON`, etc.:

```vbscript
console.log "Hello from VBScript"
Dim x : x = Math.round(3.7)   ' 4
Dim s : s = JSON.stringify(obj)
```

User-assigned `globalThis` properties are also accessible:

```vbscript
' JS side: globalThis.Forms = System.Windows.Forms
' VBS side:
Set form = New Forms.Form
```

| Keyword | Status | Notes |
|---------|--------|-------|
| Empty | ✔️ | |
| False | ✔️ | |
| Nothing | ✔️ | |
| Null | ✔️ | |
| True | ✔️ | |
| Me | ✔️ | |

---

## VB6 / VBA7 Extensions

### VB6 Type Syntax (`As Type` Annotations)

Variables, parameters, and function return types can carry explicit type
annotations using the `As <Type>` syntax. The engine uses the annotation to
initialize typed variables to their VB6 default values.

```vbscript
Dim x As Integer        ' initialized to 0 (Integer)
Dim s As String         ' initialized to "" (String)
Dim b As Boolean        ' initialized to False (Boolean)
Dim d As Double         ' initialized to 0 (Double)
Dim n As Long           ' initialized to 0 (Long)
Dim by As Byte          ' initialized to 0 (Byte)

Function Add(ByVal a As Integer, ByVal b As Integer) As Long
  Add = a + b
End Function

Sub Greet(ByVal name As String)
  MsgBox "Hello " & name
End Sub

Property Get Count() As Integer
  Count = m_count
End Property
```

Supported type names: `Integer`, `Long`, `LongLong`, `Single`, `Double`,
`Currency`, `String`, `Boolean`, `Byte`, `Date`, `Object`, `Variant`.

Type annotations are parsed and stored on the AST but are not enforced at
runtime (no implicit coercion on assignment). They affect only the initial
default value of `Dim` variables.

### VB6 Enums (`Enum…End Enum`)

Enumerations declare a set of named integer constants. Members
auto-increment from 0 (or from the previous explicit value + 1).

```vbscript
Enum Color
  Red        ' 0
  Green      ' 1
  Blue       ' 2
End Enum

Enum Status
  Pending = 10
  Active  = 20
  Inactive = 30
End Enum

Enum Priority
  Low          ' 0
  Medium = 10
  High         ' 11
  Critical     ' 12
End Enum
```

Enum members are declared as global read-only constants and can be used in
any expression, `Select Case`, or arithmetic:

```vbscript
heading = East + 45

Select Case pick
  Case Apple  : result = "apple"
  Case Banana : result = "banana"
End Select
```

`Public Enum` is supported (same as unqualified `Enum`). `Private Enum`
is also accepted syntactically.

### VB6 Collection

The built-in `Collection` class provides a VB6-compatible ordered list with
optional string keys. Create instances with `Set col = New Collection`.

```vbscript
Set col = New Collection

' Add items (optional string key)
col.Add "Paris", "FR"
col.Add "London", "UK"
col.Add 42

' Count
result = col.Count   ' 3

' Item access (1-based index or string key)
r1 = col.Item(1)     ' "Paris"
r2 = col.Item("UK")  ' "London"

' Remove by 1-based index or string key
col.Remove 1
col.Remove "UK"

' For Each iteration
total = 0
For Each item In col
  total = total + item
Next
```

Supported methods: `Add`, `Remove`, `Item`.  
Supported properties: `Count`.  
`TypeName(col)` returns `"Collection"`.

### VBA7 LongLong (64-bit Integer)

On 64-bit VBA7 targets, the `LongLong` type provides a signed 64-bit
integer. This engine implements it using JavaScript `BigInt`.

```vbscript
Dim x As LongLong      ' initialized to 0 (LongLong / BigInt)

' Conversion
x = CLngLng(42)
x = CLngLng("9223372036854775807")  ' max Int64 via string
x = CLngLng(-100)

' Arithmetic — LongLong + LongLong stays LongLong
a = CLngLng(1000000000)
b = CLngLng(2000000000)
c = a + b   ' 3000000000 (LongLong)

' Mixed: LongLong + Integer → LongLong
c = CLngLng(9000000000) + 1

' Mixed: LongLong + Double → Double (promoted)
c = CLngLng(100) + 1.5

' Integer division (returns LongLong)
c = CLngLng(10000000000) \ CLngLng(3)   ' 3333333333 (LongLong)

' Inspection
VarType(CLngLng(0))   ' 20
TypeName(CLngLng(0))  ' "LongLong"
CStr(CLngLng(x))      ' string representation
```

`VarType` code for `LongLong` is **20** (matches VBA7 specification).  
Comparison operators (`=`, `<`, `>`, `<=`, `>=`, `<>`) work between
`LongLong` values.

