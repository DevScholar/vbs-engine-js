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

## Keywords

| Keyword | Status | Notes |
|---------|--------|-------|
| Empty | ✔️ | |
| False | ✔️ | |
| Nothing | ✔️ | |
| Null | ✔️ | |
| True | ✔️ | |
| Me | ✔️ | |

