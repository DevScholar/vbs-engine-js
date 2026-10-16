# VBScript Engine Compatibility

This document tracks the compatibility between this VBScript engine implementation and Microsoft's Internet Explorer VBScript engine.

**Status Legend:**

- ✔️ **Implemented** - Feature is fully or partially implemented
- ❌ **Not Implemented** - Feature is not yet implemented but may be in the future
- ⛔ **Intentionally Not Implemented** - Feature will not be implemented due to platform limitations

---

## Core Language Features

### Data Types

| Feature  | Status | Notes |
| -------- | ------ | ----- |
| Empty    | ✔️     |       |
| Null     | ✔️     |       |
| Nothing  | ✔️     |       |
| Boolean  | ✔️     |       |
| Byte     | ✔️     |       |
| Integer  | ✔️     |       |
| Long     | ✔️     |       |
| Single   | ✔️     |       |
| Double   | ✔️     |       |
| Currency | ✔️     |       |
| Date     | ✔️     |       |
| String   | ✔️     |       |
| Object   | ✔️     |       |
| Variant  | ✔️     |       |
| Error    | ✔️     |       |

### Variables and Constants

| Feature         | Status | Notes                 |
| --------------- | ------ | --------------------- |
| Dim             | ✔️     |                       |
| ReDim           | ✔️     | With Preserve support |
| Const           | ✔️     |                       |
| Public          | ✔️     |                       |
| Private         | ✔️     |                       |
| Option Explicit | ✔️     |                       |

### Operators

| Feature                          | Status | Notes                                                                                                                                                                                                                                                                                                                                                 |
| -------------------------------- | ------ | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Arithmetic (+, -, *, /)          | ✔️     |                                                                                                                                                                                                                                                                                                                                                       |
| Integer Division (\)             | ✔️     |                                                                                                                                                                                                                                                                                                                                                       |
| Modulo (Mod)                     | ✔️     |                                                                                                                                                                                                                                                                                                                                                       |
| Exponentiation (^)               | ✔️     |                                                                                                                                                                                                                                                                                                                                                       |
| Concatenation (&)                | ✔️     |                                                                                                                                                                                                                                                                                                                                                       |
| Comparison (=, <>, <, >, <=, >=) | ✔️     | Text is compared character by character and case sensitively, so `"b" = "B"` is False. A text beside a number is read as a number when it is one (`"1" = 1` is True), and beside anything else as text (`"True" = True` is True and `"1" = True` is False). A Null on either side answers Null, an object is error 91 or 438 and an array is error 13 |
| Comparison with a literal number | ✔️     | A text that is no number beside a number written out in the source is the type mismatch Windows raises: `"abc" = 1` and `"abc" = (1)` are error 13. A number that came from anywhere else is compared as text instead and answers False, which is what `"abc" = 1 * 1`, `"abc" = CInt(1)` and `"abc" = Len("a")` do                                   |
| Is                               | ✔️     |                                                                                                                                                                                                                                                                                                                                                       |
| And                              | ✔️     |                                                                                                                                                                                                                                                                                                                                                       |
| Or                               | ✔️     |                                                                                                                                                                                                                                                                                                                                                       |
| Not                              | ✔️     |                                                                                                                                                                                                                                                                                                                                                       |
| Xor                              | ✔️     |                                                                                                                                                                                                                                                                                                                                                       |
| Eqv                              | ✔️     |                                                                                                                                                                                                                                                                                                                                                       |
| Imp                              | ✔️     |                                                                                                                                                                                                                                                                                                                                                       |
| Operator Precedence              | ✔️     |                                                                                                                                                                                                                                                                                                                                                       |

### Control Flow Statements

| Feature                             | Status | Notes                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 |
| ----------------------------------- | ------ | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| If...Then...Else                    | ✔️     | Single-line and multi-line                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            |
| ElseIf                              | ✔️     |                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       |
| For...Next                          | ✔️     | With Step support. A bound or a step that is no number is error 92, the loop's own "For loop not initialized"                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         |
| For Each...Next                     | ✔️     | Leaves its loop variable empty after the loop, which is what Windows does; For...Next keeps the value that stopped it                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 |
| Do...Loop                           | ✔️     | While/Until variants. A Null or an Empty condition is false rather than an error                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      |
| While...Wend                        | ✔️     |                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       |
| Select Case                         | ✔️     | With Case Else. A case is a plain expression or a comma list, as in VBScript: `Case 1 To 5` is a parse error there and here. It compares the way the operators do, so it reads text case sensitively and raises error 13 for a text beside a literal number                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           |
| Exit (Do/For/Sub/Function/Property) | ✔️     |                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       |
| With...End With                     | ✔️     |                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       |
| GoTo                                | ✔️     | With label support                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    |
| On Error                            | ✔️     | Resume Next, GoTo 0. Its reach is the procedure it stands in, measured on a fresh cscript: a handler set in a procedure is not in force in a procedure that one calls, and the handler the caller set is in force again once the call returns. So with the handler at script level, `y = 1 / 0` inside a Sub answers 11 and the Sub's own statements after it do not run, while the caller carries on at the statement after the call; and a handler a Function set for itself is gone once it returns, so a later failure at script level is not swallowed. Both statements also reset the `Err` object themselves, which the Err object's own documentation says and a fresh cscript shows: `Err.Raise 4321, "a source", "a description"` followed by another `On Error Resume Next` leaves `Err.Number` 0 with an empty `Source` and `Description` |
| Resume                              | ✔️     | Partial implementation                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                |

---

## Procedures and Classes

### Procedures

| Feature             | Status | Notes               |
| ------------------- | ------ | ------------------- |
| Sub                 | ✔️     |                     |
| Function            | ✔️     |                     |
| Call                | ✔️     |                     |
| ByVal               | ✔️     |                     |
| ByRef               | ✔️     | Default behavior    |
| Optional Parameters | ✔️     | With default values |
| ParamArray          | ✔️     |                     |

### Classes

| Feature                          | Status | Notes                                                                                                                                                                                                                                                           |
| -------------------------------- | ------ | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Class...End Class                | ✔️     |                                                                                                                                                                                                                                                                 |
| New                              | ✔️     | `New` on a name that is no class is error 506, "Class not defined"                                                                                                                                                                                              |
| Public members                   | ✔️     | A field, a Property or a method reachable from anywhere                                                                                                                                                                                                         |
| Private members                  | ✔️     | Reachable from inside the same class only, another instance of it included. From outside, reading, writing or calling one is error 438, as it is in Windows                                                                                                     |
| A member the class does not have | ✔️     | Error 438 for a read, a write and a call alike: `first.NoSuchMember`, `first.NoSuchMember = 1` and `first.NoSuchMethod()`                                                                                                                                       |
| Property Get                     | ✔️     | Handing back an object works (`Set Linked = inner`), including read inside an expression such as `TypeName(first.Linked)`                                                                                                                                       |
| Property Let                     | ✔️     |                                                                                                                                                                                                                                                                 |
| Property Set                     | ✔️     | An object stored this way comes back with `Is` intact                                                                                                                                                                                                           |
| Method argument count            | ✔️     | Checked the way a procedure's is: `Describe()` and `Describe("a", "b")` for `Describe(prefix)` are both error 450                                                                                                                                               |
| Me                               | ✔️     |                                                                                                                                                                                                                                                                 |
| Initialize Event                 | ✔️     | Class_Initialize, run by `New`; naming it from outside the class is error 438 without running it                                                                                                                                                                |
| Terminate Event                  | ✔️     | Class_Terminate, run when the last reference to an instance goes: a variable given another value or left behind by a procedure, an array element written over or freed, a field of an instance that is freed, and the end of the script. See the Events section |
| Assigning an object              | ✔️     | Needs `Set`. Without it the assignment is error 438 (`x = New Holder`, `x = Maker()`, `x = held`, `x = items(0)`), and `x = Nothing` is error 91. Passing an object as an argument needs no `Set`                                                               |
| An object in an array            | ✔️     | `Set items(0) = New Holder` works, and `items(0).Name` reads and writes through it                                                                                                                                                                              |
| Reading an object as text        | ✔️     | Error 438: an object of a class has no text of its own, so `CStr(anInstance)` is 438                                                                                                                                                                            |

---

## Built-in Functions

A Null argument is read the way Windows reads one, and the functions differ on it. All of it is measured on a
fresh cscript:

- A function that gives a value back from the argument it was handed answers Null for a Null there. `Len(Null)`,
  `UCase(Null)`, `Trim(Null)`, `Left(Null, 2)`, `Mid(Null, 1, 1)`, `InStr("abc", Null)`, `InStr(1, "abc", Null)`,
  `Hex(Null)`, `Oct(Null)`, `Abs(Null)`, `Int(Null)`, `Fix(Null)`, `Year(Null)`, `Month(Null)`, `Day(Null)`,
  `Hour(Null)`, `Minute(Null)`, `Second(Null)`, `Weekday(Null)`, `DatePart("d", Null)`, `DateAdd("d", 1, Null)`
  and `DateDiff("d", Null, Date)` are all Null.
- A Null anywhere else in those calls is the invalid use of Null, error 94: `Left("abc", Null)`,
  `Right("abc", Null)`, `Mid("abc", Null, 1)`, `InStr(Null, "abc", "b")`, `InStr(1, "abc", "b", Null)`,
  `Weekday(Date, Null)`, `DatePart(Null, Date)`, `DateAdd("d", Null, Date)`, `DateDiff(Null, Date, Date)`,
  `DateSerial(Null, 1, 1)` and `TimeSerial(Null, 1, 1)` alike.
- The functions that build a result from their arguments refuse a Null anywhere with error 94: `Replace`,
  `Split`, `Join`, `Filter`, `String`, `Space`, `StrReverse`, `CStr`, `Sgn`, `Sqr`, `Round`, `Exp`, `Log`,
  `Sin`, `Cos`, `Tan`, `Atn` and the other `C` conversions.
- `FormatNumber`, `FormatCurrency`, `FormatPercent` and `FormatDateTime` are the exception, with a type
  mismatch, error 13, for a Null.

### String Functions

| Function   | Status | Notes                                                                                                                                                                                                                                                                               |
| ---------- | ------ | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Len        | ✔️     |                                                                                                                                                                                                                                                                                     |
| Left       | ✔️     |                                                                                                                                                                                                                                                                                     |
| Right      | ✔️     |                                                                                                                                                                                                                                                                                     |
| Mid        | ✔️     |                                                                                                                                                                                                                                                                                     |
| InStr      | ✔️     |                                                                                                                                                                                                                                                                                     |
| InStrRev   | ✔️     |                                                                                                                                                                                                                                                                                     |
| LCase      | ✔️     |                                                                                                                                                                                                                                                                                     |
| UCase      | ✔️     |                                                                                                                                                                                                                                                                                     |
| LTrim      | ✔️     |                                                                                                                                                                                                                                                                                     |
| RTrim      | ✔️     |                                                                                                                                                                                                                                                                                     |
| Trim       | ✔️     |                                                                                                                                                                                                                                                                                     |
| Replace    | ✔️     |                                                                                                                                                                                                                                                                                     |
| StrReverse | ✔️     |                                                                                                                                                                                                                                                                                     |
| Space      | ✔️     |                                                                                                                                                                                                                                                                                     |
| String     | ✔️     | Create repeated character string                                                                                                                                                                                                                                                    |
| Asc        | ✔️     |                                                                                                                                                                                                                                                                                     |
| AscW       | ✔️     |                                                                                                                                                                                                                                                                                     |
| Chr        | ✔️     |                                                                                                                                                                                                                                                                                     |
| ChrW       | ✔️     |                                                                                                                                                                                                                                                                                     |
| StrComp    | ✔️     |                                                                                                                                                                                                                                                                                     |
| Split      | ✔️     | Reads its text as text, so `Split(1, ",")` makes one part, and reads the delimiter, the limit and the compare the way Windows does: a limit below -1 is error 5, 0 makes no parts, and an empty text makes none either. An array or an object as the text is error 13 and 91 or 438 |
| Join       | ✔️     | Reads one dimensional arrays only, so a two dimensional one is error 13, and refuses a Null or an array element with error 13 where an Empty element is no text at all. `Join(Null)` is error 94, `Join(Nothing)` is 91, a list that is no array is 13                              |
| Filter     | ✔️     | Reads its compare argument (0 binary, 1 text, anything else error 5), returns the elements it found with their own types, and reads a Null element as no text at all. A two dimensional array or a list that is no array is error 13                                                |
| Format     | ✔️     | Custom format strings                                                                                                                                                                                                                                                               |
| LSet       | ✔️     |                                                                                                                                                                                                                                                                                     |
| RSet       | ✔️     |                                                                                                                                                                                                                                                                                     |

### Mathematical Functions

| Function  | Status | Notes             |
| --------- | ------ | ----------------- |
| Abs       | ✔️     |                   |
| Sgn       | ✔️     |                   |
| Sqr       | ✔️     |                   |
| Int       | ✔️     |                   |
| Fix       | ✔️     |                   |
| Round     | ✔️     |                   |
| Atn       | ✔️     |                   |
| Cos       | ✔️     |                   |
| Sin       | ✔️     |                   |
| Tan       | ✔️     |                   |
| Exp       | ✔️     |                   |
| Log       | ✔️     | Natural logarithm |
| Rnd       | ✔️     |                   |
| Randomize | ✔️     |                   |
| Hex       | ✔️     |                   |
| Oct       | ✔️     |                   |

### Date and Time Functions

| Function    | Status | Notes                                                                                                                                                                                                                                                                                                                                                                  |
| ----------- | ------ | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Now         | ✔️     |                                                                                                                                                                                                                                                                                                                                                                        |
| Date        | ✔️     |                                                                                                                                                                                                                                                                                                                                                                        |
| Time        | ✔️     |                                                                                                                                                                                                                                                                                                                                                                        |
| Year        | ✔️     |                                                                                                                                                                                                                                                                                                                                                                        |
| Month       | ✔️     |                                                                                                                                                                                                                                                                                                                                                                        |
| Day         | ✔️     |                                                                                                                                                                                                                                                                                                                                                                        |
| Weekday     | ✔️     |                                                                                                                                                                                                                                                                                                                                                                        |
| Hour        | ✔️     |                                                                                                                                                                                                                                                                                                                                                                        |
| Minute      | ✔️     |                                                                                                                                                                                                                                                                                                                                                                        |
| Second      | ✔️     |                                                                                                                                                                                                                                                                                                                                                                        |
| DateAdd     | ✔️     |                                                                                                                                                                                                                                                                                                                                                                        |
| DateDiff    | ✔️     |                                                                                                                                                                                                                                                                                                                                                                        |
| DatePart    | ✔️     |                                                                                                                                                                                                                                                                                                                                                                        |
| DateSerial  | ✔️     |                                                                                                                                                                                                                                                                                                                                                                        |
| TimeSerial  | ✔️     |                                                                                                                                                                                                                                                                                                                                                                        |
| DateValue   | ✔️     |                                                                                                                                                                                                                                                                                                                                                                        |
| TimeValue   | ✔️     |                                                                                                                                                                                                                                                                                                                                                                        |
| MonthName   | ✔️     | The index is rounded to a whole number, a tie to the even neighbour, and anything outside 1 to 12 is error 5 rather than an empty name: `MonthName(0)`, `MonthName(13)`, `MonthName(-1)` and `MonthName(Empty)` are all error 5, `MonthName(3.6)` is April and `MonthName(2.5)` is February. A Nothing is 91 and a Null is 94                                          |
| WeekdayName | ✔️     | Its index is read the way MonthName reads one. The third argument counts the week from a day of its own, so `WeekdayName(1, False, 2)` is Monday and `WeekdayName(1, False, 7)` is Saturday, while 0 counts from the locale's own first day, which the engine ships per locale; leaving the argument out always counts from Sunday. Anything outside 0 to 7 is error 5 |
| Timer       | ✔️     |                                                                                                                                                                                                                                                                                                                                                                        |

### Conversion Functions

| Function       | Status | Notes                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     |
| -------------- | ------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| CBool          | ✔️     |                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           |
| CByte          | ✔️     |                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           |
| CCur           | ✔️     |                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           |
| CDate          | ✔️     | Reads a written date with the locale's own order and separators, from the Windows short date picture, the spaced dotted form the picture itself writes included: `CDate("2026. 05. 01.")` is 1 May 2026 under 1038 and `CDate("1. 5. 2026")` is a date under 1031, 1038, 1051, 1054 and 1060. Fields separated by spaces are read the way the same fields with the locale's own mark are, and a first field that is no day and no month is a year: `CDate("15 1 26")` is 26 January 2015 and `CDate("31/12/26")` is 26 December 2031 under 1033. A comma may separate the fields of a date but not a date from its time, except under the five locales that read it there. A space where the locale groups its numbers is a number, so `CDate("2026 15 1")` is the day count 2026151 under 1026 and error 13 under 1033. Measured across the 132 locales that answer. A month name is read when it is the locale's own, and a word that _is_ one of those names matches whatever its length, while a shorter word matches as a prefix only from three letters: Catalan "ag.", Arabic "آب" and Persian "مه" are months, Czech "červen" is June and not July, and a name a locale writes with a dot or a space (Thai "ม.ค.", Vietnamese "Thg1") is read whole. A name that is only digits is no name, so "29.9.26" is error 13 under 1051. A dot beside a month name is read the way Windows reads one, all of it measured over the 132 locales: one dot after a name is the dot of an abbreviation and reads almost everywhere (`CDate("15 Jan. 2026")` is 15 January 2026 under 1033), two dots read nowhere, a name itself written with a dot refuses a second one (Catalan "15 ag. 2026" is August and "15 ag.. 2026" is error 13), and where the locale writes its dates with a dot, a name Windows writes with a dot is read only with a space or the end of the text after it (`CDate("15/jan./2026")` and `CDate("15 jan., 2026")` are error 13 under 1038, while the same texts read under 1027, whose own date mark is a slash, and under 1054, whose dotted names are read wherever they stand). Chinese, Japanese and Korean read only English names, and the names each locale reads back are measured per name and per spelling: Yiddish and Tamil read none of their own, Hindi only May, Tsonga everything but December, and Greek reads the names Windows writes and not the inflected long ones ICU gives |
| CDbl           | ✔️     |                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           |
| CInt           | ✔️     |                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           |
| CLng           | ✔️     |                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           |
| CSng           | ✔️     |                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           |
| CStr           | ✔️     |                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           |
| CVar           | ✔️     | Not a function of `vbscript.dll`, which answers error 13 for every call, so `CVar(1)`, `"1" = CVar(1)` and `IsEmpty(CVar(1))` are all error 13 there. Kept because VB6 and VBA have it                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    |
| CVErr          | ✔️     | Not a function of `vbscript.dll` either, which answers error 13 for `CVErr(1)` as well. Kept because VB6 and VBA have it                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  |
| FormatNumber   | ✔️     | Reads the Windows NLS number table, so the fraction mark, the group marks and the fraction count are the locale's own. The fraction count and the three tristate flags are rounded to a whole number first, a tie to the even neighbour, and a flag that is not -1, 0 or -2 is error 5: `FormatNumber(0.5, 2, 1)` is error 5 while `-0.5` is False and `-1.5` is the locale's own. A Null is error 13 for the value and 94 for the count, and a Nothing is 91                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             |
| FormatCurrency | ✔️     | The symbol, its position, the fraction count and the shape of a negative amount come from the locale's row in the same table, and its arguments are read the way FormatNumber's are                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       |
| FormatPercent  | ✔️     | Reads the Windows NLS number table, and writes the percent sign the way Windows does, with the same reading of its arguments                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                              |
| FormatDateTime | ✔️     | Reads the Windows NLS date table, so it follows Windows and not the host. Its date argument is read the way CDate reads one, so `FormatDateTime(45000)` is 15 March 2023 where the JavaScript date parser gives the year 45000, and a text is read with the locale's own order and calendar. A Null date is error 13 and a Nothing is 91. The named format is rounded the same way and must land between 0 and 4: 2.7 is the long time and 4.5 the short time, while 5, -1 and True are error 5. The long and the short date are error 5 for a day Windows' own date writers cannot write, which is anything before 1601-01-01 in every locale and anything outside the Hijri range 1318-01-01 to 1500-12-30 for ar-SA; the Hebrew locale is the one that answers 32812 there instead. The Hijri date comes from the Umm al-Qura table, which is the one .NET ships, because ICU's islamic-umalqura calendar is a day behind Windows for about a fifth of the days. A picture that asks for a two digit year writes four digits outside 1950 to 2049 for GeneralDate and CStr, while the short date keeps two                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             |

### Inspection Functions

| Function  | Status | Notes                                                                                                                                                                                                          |
| --------- | ------ | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| IsArray   | ✔️     |                                                                                                                                                                                                                |
| IsDate    | ✔️     |                                                                                                                                                                                                                |
| IsEmpty   | ✔️     |                                                                                                                                                                                                                |
| IsNull    | ✔️     |                                                                                                                                                                                                                |
| IsNumeric | ✔️     | Reads a string the way the conversions do, with the locale's own marks, so `IsNumeric("50%")` is `False` and `IsNumeric("1 234")` is `False` under 1031 but `True` under 1026, where a space is the group mark |
| IsObject  | ✔️     |                                                                                                                                                                                                                |
| VarType   | ✔️     |                                                                                                                                                                                                                |
| TypeName  | ✔️     |                                                                                                                                                                                                                |

### Array Functions

| Function | Status | Notes                                                                                                                                                                                                                                                                                                                                                   |
| -------- | ------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Array    | ✔️     | The array it makes is dynamic, so `Erase` frees it and `ReDim` may size it again                                                                                                                                                                                                                                                                        |
| LBound   | ✔️     | The dimension is a whole number rounded to the nearest one, a tie to the even neighbour, so `LBound(a, "2")` is the second dimension and `LBound(a, 1.5)` is the second as well. A dimension outside the array is error 9, a Null is error 94, no array at all is 13, and an object is 91 for a Nothing and 438 for any other                           |
| UBound   | ✔️     | Reads its dimension the way LBound does, and a name Dim declared as an array still answers here after `Erase`, which is error 9 for a dynamic one                                                                                                                                                                                                       |
| Erase    | ✔️     | Takes one name: `Erase a` frees what a holds, `Erase a(0)` frees what the element holds, and `Erase(a)`, `Erase Array(1, 2)` and `Erase h.Items` are error 13 while `Erase a, b` and `Erase` alone are 450. A fixed array keeps its bounds and every element becomes Empty, and a dynamic one, `Array`, `Split` and `Filter` results included, is freed |
| Filter   | ✔️     | See String Functions                                                                                                                                                                                                                                                                                                                                    |

### Other Functions

| Function                 | Status | Notes                                                                                                                                                                                                                                                             |
| ------------------------ | ------ | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Eval                     | ✔️     | Reads its text as an expression, so a bare name is the value it holds or the procedure it names rather than the call a bare name statement is: `Eval("number")` is 5 and `Eval("NoSuchFunction")` is 0, where those names written as statements are both error 13 |
| Execute                  | ✔️     | Dynamic code execution in current scope                                                                                                                                                                                                                           |
| ExecuteGlobal            | ✔️     | Dynamic code execution in global scope                                                                                                                                                                                                                            |
| MsgBox                   | ✔️     | Browser simulation using alert/confirm                                                                                                                                                                                                                            |
| InputBox                 | ✔️     | Browser simulation using prompt                                                                                                                                                                                                                                   |
| GetRef                   | ✔️     |                                                                                                                                                                                                                                                                   |
| GetLocale                | ✔️     | Uses Intl API to detect browser locale                                                                                                                                                                                                                            |
| SetLocale                | ✔️     | Sets locale, returns LCID                                                                                                                                                                                                                                         |
| CreateObject             | ✔️     | Delegates to the host's `ActiveXObject`; see COM Objects below                                                                                                                                                                                                    |
| GetObject                | ✔️     | Delegates to the host's `GetObject`; see COM Objects below                                                                                                                                                                                                        |
| LoadPicture              | ✔️     | Returns stub IPictureDisp object                                                                                                                                                                                                                                  |
| RGB                      | ✔️     | Returns RGB color value                                                                                                                                                                                                                                           |
| QBColor                  | ✔️     | Returns legacy color value                                                                                                                                                                                                                                        |
| ScriptEngine             | ✔️     | Returns "VBScript"                                                                                                                                                                                                                                                |
| ScriptEngineMajorVersion | ✔️     | Returns 10                                                                                                                                                                                                                                                        |
| ScriptEngineMinorVersion | ✔️     | Returns 8                                                                                                                                                                                                                                                         |
| ScriptEngineBuildVersion | ✔️     | Returns 16384                                                                                                                                                                                                                                                     |

---

## Objects

### Err Object

| Property/Method | Status | Notes                                                                                                                                                                                                                                                                                                                                                                                                                                                                                   |
| --------------- | ------ | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Number          | ✔️     |                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         |
| Description     | ✔️     |                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         |
| Source          | ✔️     |                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         |
| HelpFile        | ✔️     | Read/write `String`, default `""` (no help source); only set when `Err.Raise` passes a helpfile argument. No `.hlp` loading — WinHelp was removed from modern Windows                                                                                                                                                                                                                                                                                                                   |
| HelpContext     | ✔️     | Read/write `Long`, default `0` (no help source); only set when `Err.Raise` passes a helpcontext argument. No `.hlp` loading                                                                                                                                                                                                                                                                                                                                                             |
| LastDllError    | ⚠️     | Not a member of `vbscript.dll`'s `Err` object at all: measured on a fresh `cscript`, reading it is error 438, the same as any other name that is not there. Kept as a VB6 extension, and it answers 0 because no DLL loader is wired                                                                                                                                                                                                                                                    |
| Clear           | ✔️     |                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         |
| Raise           | ✔️     | Supports `number, source, description, helpfile, helpcontext`. It fills `Err` and then raises, so a script with no handler stops at that statement, measured on a fresh cscript: a script that writes a line, then `Err.Raise 4321, "a source", "a description"`, then another line prints the first line and `(4, 1) a source: a description`, and the line after it never runs. A handler is what keeps the script going, and the fields it filled are the ones the script then reads |

### RegExp Object

| Feature               | Status | Notes                     |
| --------------------- | ------ | ------------------------- |
| RegExp Object         | ✔️     |                           |
| Pattern               | ✔️     |                           |
| Global                | ✔️     |                           |
| IgnoreCase            | ✔️     |                           |
| Multiline             | ✔️     |                           |
| Execute Method        | ✔️     |                           |
| Test Method           | ✔️     |                           |
| Replace Method        | ✔️     |                           |
| Match Object          | ✔️     | FirstIndex, Length, Value |
| Matches Collection    | ✔️     | Count, Item               |
| SubMatches Collection | ✔️     |                           |

### Class Object

| Feature             | Status | Notes |
| ------------------- | ------ | ----- |
| Class definition    | ✔️     |       |
| Public members      | ✔️     |       |
| Private members     | ✔️     |       |
| Property procedures | ✔️     |       |
| Methods             | ✔️     |       |

### VBArray Object (JavaScript-side)

VBArray is a JavaScript-side object for accessing VBScript arrays from JavaScript.
It provides backward compatibility with early Internet Explorer, which required
wrapping VBScript arrays before accessing them from JavaScript.

| Feature             | Status | Notes                                |
| ------------------- | ------ | ------------------------------------ |
| VBArray constructor | ✔️     | Wraps internal VbArray for JS access |
| dimensions()        | ✔️     | Returns number of dimensions         |
| lbound()            | ✔️     | Returns lower bound of a dimension   |
| ubound()            | ✔️     | Returns upper bound of a dimension   |
| getItem()           | ✔️     | Gets value at specified indices      |
| toArray()           | ✔️     | Converts to JavaScript array         |

**Configuration:** The `injectVBArrayToGlobalThis` option (default: `true`) controls
whether VBArray is available as `globalThis.VBArray`. When enabled, legacy code
using `new VBArray(vbArray)` will work. Modern code can access VBScript arrays
directly without wrapping.

---

## Statements

| Statement                       | Status | Notes                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                   |
| ------------------------------- | ------ | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Call                            | ✔️     |                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         |
| Class                           | ✔️     |                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         |
| Const                           | ✔️     |                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         |
| Dim                             | ✔️     | The engine reads the bound of an array at run time, so `Dim a(1.5)` and `Dim a(n)` run here. Windows refuses to compile a bound that is not a whole number written out in the source: `Dim a(1.5)`, `Dim a(-1)`, `Dim a(1 + 2)` and `Dim a(namedConstant)` are all compile errors there, and only a literal such as `Dim a(3)` is accepted                                                                                                                                                                                              |
| Do...Loop                       | ✔️     |                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         |
| Erase                           | ✔️     | Takes one name, as Windows does: `Erase a` frees what a holds, `Erase a(0)` frees what the element holds, and `Erase(a)`, `Erase Array(1, 2)` and `Erase h.Items` are error 13 while `Erase a, b` and `Erase` alone are 450                                                                                                                                                                                                                                                                                                             |
| Execute                         | ✔️     |                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         |
| ExecuteGlobal                   | ✔️     |                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         |
| Exit                            | ✔️     |                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         |
| For...Next                      | ✔️     |                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         |
| For Each...Next                 | ✔️     |                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         |
| Function                        | ✔️     |                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         |
| GoTo                            | ✔️     |                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         |
| If...Then...Else                | ✔️     |                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         |
| On Error                        | ✔️     |                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         |
| Option Explicit                 | ✔️     | An undeclared name is error 500 in every position, where without it a call on such a name indexes the Empty it stands for and is error 13                                                                                                                                                                                                                                                                                                                                                                                               |
| A statement that is just a name | ✔️     | A whole statement naming something is a call on it with no arguments, as it is in Windows: a name holding a number, a text, an Empty, a Null or an array is error 13, a name holding an object is 438, and a procedure is called, with 450 when it needs an argument. `arr(9)` alone is error 13 where `x = arr(9)` is error 9, a member holding a field is 5 with no arguments and 438 with them, and a Property Get reads and drops its value. A literal as a statement is a compile error in Windows and is not accepted here either |
| Private                         | ✔️     |                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         |
| Property Get/Let/Set            | ✔️     |                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         |
| Public                          | ✔️     |                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         |
| Randomize                       | ✔️     |                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         |
| ReDim                           | ✔️     | A bound is rounded to the nearest whole number, a tie to the even neighbour, and `ReDim a(-1)` makes an array of no elements whose UBound is -1. `ReDim` on a fixed array is error 10, `ReDim Preserve` may change only the last dimension of an array that already has one, and changing any other dimension or how many there are is error 9. An array with no size yet, from `Dim a()` or from `Erase`, takes any shape Preserve asks for                                                                                            |
| Rem                             | ✔️     |                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         |
| Select Case                     | ✔️     |                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         |
| Set                             | ✔️     |                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         |
| Sub                             | ✔️     |                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         |
| While...Wend                    | ✔️     |                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         |
| With                            | ✔️     |                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         |

---

## Constants

### String Constants

| Constant      | Status | Notes |
| ------------- | ------ | ----- |
| vbCr          | ✔️     |       |
| vbCrLf        | ✔️     |       |
| vbFormFeed    | ✔️     |       |
| vbLf          | ✔️     |       |
| vbNewLine     | ✔️     |       |
| vbNullChar    | ✔️     |       |
| vbNullString  | ✔️     |       |
| vbTab         | ✔️     |       |
| vbVerticalTab | ✔️     |       |

### VarType Constants

| Constant   | Status | Notes |
| ---------- | ------ | ----- |
| vbEmpty    | ✔️     |       |
| vbNull     | ✔️     |       |
| vbInteger  | ✔️     |       |
| vbLong     | ✔️     |       |
| vbSingle   | ✔️     |       |
| vbDouble   | ✔️     |       |
| vbCurrency | ✔️     |       |
| vbDate     | ✔️     |       |
| vbString   | ✔️     |       |
| vbObject   | ✔️     |       |
| vbError    | ✔️     |       |
| vbBoolean  | ✔️     |       |
| vbVariant  | ✔️     |       |
| vbByte     | ✔️     |       |
| vbArray    | ✔️     |       |

### MsgBox Constants

| Constant             | Status | Notes |
| -------------------- | ------ | ----- |
| vbOKOnly             | ✔️     |       |
| vbOKCancel           | ✔️     |       |
| vbAbortRetryIgnore   | ✔️     |       |
| vbYesNoCancel        | ✔️     |       |
| vbYesNo              | ✔️     |       |
| vbRetryCancel        | ✔️     |       |
| vbCritical           | ✔️     |       |
| vbQuestion           | ✔️     |       |
| vbExclamation        | ✔️     |       |
| vbInformation        | ✔️     |       |
| vbDefaultButton1/2/3 | ✔️     |       |
| vbOK                 | ✔️     |       |
| vbCancel             | ✔️     |       |
| vbAbort              | ✔️     |       |
| vbRetry              | ✔️     |       |
| vbIgnore             | ✔️     |       |
| vbYes                | ✔️     |       |
| vbNo                 | ✔️     |       |

### Date Constants

| Constant              | Status | Notes |
| --------------------- | ------ | ----- |
| vbSunday - vbSaturday | ✔️     |       |
| vbUseSystemDayOfWeek  | ✔️     |       |
| vbFirstJan1           | ✔️     |       |
| vbFirstFourDays       | ✔️     |       |
| vbFirstFullWeek       | ✔️     |       |
| vbGeneralDate         | ✔️     |       |
| vbLongDate            | ✔️     |       |
| vbShortDate           | ✔️     |       |
| vbLongTime            | ✔️     |       |
| vbShortTime           | ✔️     |       |

### Comparison Constants

| Constant          | Status | Notes                                 |
| ----------------- | ------ | ------------------------------------- |
| vbBinaryCompare   | ✔️     |                                       |
| vbTextCompare     | ✔️     |                                       |
| vbDatabaseCompare | ✔️     | Falls back to locale-aware comparison |

### Other Constants

| Constant      | Status | Notes |
| ------------- | ------ | ----- |
| vbObjectError | ✔️     |       |

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

| Function     | Status | Notes                                        |
| ------------ | ------ | -------------------------------------------- |
| CreateObject | ✔️     | Hands the call to the host's `ActiveXObject` |
| GetObject    | ✔️     | Hands the call to the host's `GetObject`     |

---

## WSH Host (`WScript`) and `run-script`

The engine itself does not inject a `WScript` object. Like `ActiveXObject`/`GetObject`,
it is supplied by the host through `globalThis` — present if the host provides it,
absent otherwise. The `run-script` CLI (this package's `bin`) provides a
cscript/wscript-style host on top of this, so scripts written for Windows Script
Host run with minimal changes.

### `WScript` host object

Injected by the host as `globalThis.WScript`, exposing the standard WSH surface:

| Member                                                   | Status | Notes                                                                                 |
| -------------------------------------------------------- | ------ | ------------------------------------------------------------------------------------- |
| `WScript.Echo`                                           | ✔️     | Writes to stdout in cscript mode; joined with spaces                                  |
| `WScript.Quit n`                                         | ✔️     | Throws a `WScriptQuitSignal` to unwind; `run-script` maps it to the process exit code |
| `WScript.Sleep ms`                                       | ✔️     | Synchronous sleep                                                                     |
| `WScript.Arguments`                                      | ✔️     | `Count`, `(n)`, `Item(n)`, `Named`, `Unnamed`                                         |
| `WScript.Arguments.Named("x")` / `.Exists("x")`          | ✔️     | Named args from `/name:value` and `.wsf` `<runtime>`                                  |
| `WScript.StdOut` / `StdErr`                              | ✔️     | `Write` / `WriteLine` / `WriteBlankLines`                                             |
| `WScript.StdIn`                                          | ✔️     | `ReadLine` / `ReadAll` (synchronous)                                                  |
| `WScript.ScriptName` / `ScriptFullName`                  | ✔️     |                                                                                       |
| `WScript.FullName` / `Path` / `Version` / `BuildVersion` | ✔️     | Host paths and a fixed `5.812` version string                                         |
| `WScript.Interactive`                                    | ✔️     | `false` under `//B`                                                                   |
| `WScript.CreateObject` / `GetObject`                     | ✔️     | Delegate to the host's `ActiveXObject`/`GetObject`                                    |

`WScript` is implemented as a plain object exposing the engine's `VbObjectValueData`
protocol (`getProperty`/`hasMethod`/`getMethod`). Default-property invocation
(`WScript.Arguments(0)`, `WScript.Arguments.Named("x")`) is supported by the
engine's member-call dispatch.

### `run-script` CLI

```
run-script script.vbs [//options] [script-args...]
run-script script.wsf //job:Job1 [//options] [script-args...]
```

| Option                        | Meaning                                    |
| ----------------------------- | ------------------------------------------ |
| `//nologo`                    | Suppress the banner                        |
| `//job:<id>`                  | Select a `<job>` in a `.wsf`               |
| `//B` / `//I`                 | Batch (non-interactive) / interactive mode |
| `//E:<engine>`                | Target engine (VBScript only for now)      |
| `--gui-mode=cscript\|wscript` | Drives `WScript.Echo` (console vs GUI)     |

Named script arguments use WSH's `/name:value` form; bare words are positional.
`WScript.Quit n` sets the process exit code.

### Windows Script File (`.wsf`)

`run-script` parses `.wsf` files (via `fast-xml-parser`), supporting:

| Element                                   | Status | Notes                                                                                                                                                                                    |
| ----------------------------------------- | ------ | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `<package>` / `<job>`                     | ✔️     | Multiple jobs; single job without `<package>`                                                                                                                                            |
| `<script language="VBScript">`            | ✔️     | Inline text, CDATA, and `src="..."`                                                                                                                                                      |
| `<runtime>`                               | ✔️     | `<named>`, `<unnamed>`, `<description>`, `<example>`; feeds `WScript.Arguments` and `ShowUsage`                                                                                          |
| `<object>` / `<reference>` / `<resource>` | ⚠️     | `<object id progid>` is hoisted into an implicit `Set <id> = CreateObject(progid)` before scripts run; `<reference>` is parsed but its typelib/constant loading is not yet performed     |
| `<script language="JScript">`             | ✔️     | Executed via the host JS engine in the global context (`vm.runInThisContext`); top-level `var`/`function` become `globalThis` properties, so VBScript and JScript share one global scope |

---

## Events

Every rule below was read off a fresh `cscript`, one cell per process, since an instance that outlives its
cell is still alive when the next cell starts. The probe, its capture and the comparison script live in the
temporary probe directory, and `src/__tests__/class-events.test.ts` is the durable copy of its 59 cells.

| Event                          | Status | Notes                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            |
| ------------------------------ | ------ | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Initialize                     | ✔️     | Class_Initialize, run by `New` once the instance exists and before the value reaches the caller                                                                                                                                                                                                                                                                                                                                                                                                                                  |
| Terminate                      | ✔️     | Class_Terminate, run when the last reference to the instance goes                                                                                                                                                                                                                                                                                                                                                                                                                                                                |
| A variable given another value | ✔️     | The old instance's Terminate runs at that statement, between the new instance's Initialize and the assignment: `Set o = New Thing` answers init, term, init                                                                                                                                                                                                                                                                                                                                                                      |
| A variable a procedure leaves  | ✔️     | A local variable's instance is freed when the procedure returns, and so is the one a ByVal parameter holds, since that parameter has a copy of its own                                                                                                                                                                                                                                                                                                                                                                           |
| A ByRef parameter              | ✔️     | It is another name for the caller's variable and holds no reference of its own, so the instance the caller handed it stays alive after the return, even when the parameter was given another value inside                                                                                                                                                                                                                                                                                                                        |
| An array element               | ✔️     | An element that is written over, freed by `Erase` or dropped by `ReDim` lets go of its instance; an instance only an array element holds is kept alive by it                                                                                                                                                                                                                                                                                                                                                                     |
| An instance field              | ✔️     | Writing over a plain field lets go of what it held, and an instance that is itself freed lets go of the instances its fields hold                                                                                                                                                                                                                                                                                                                                                                                                |
| A function's answer            | ✔️     | The answer carries a reference of its own, so a `Function` that answers a fresh instance hands the caller something it can still name                                                                                                                                                                                                                                                                                                                                                                                            |
| The end of a script            | ✔️     | An instance the script never dropped has its Terminate event run as the script ends. `VbsEngine.releaseGlobals()` does this. The engine that claimed host integration calls it by itself when its host ends, which is the process in Node and the page going away in a browser, so a script run through this repository's own `run-script` host ends the way `cscript` ends. A host that runs several scripts through one engine calls it between them, since the engine cannot tell where one script stopped and the next began |
| Naming an event from outside   | ✔️     | `o.Class_Initialize`, `o.Class_Initialize()`, `o.Class_Terminate` and `Call o.Class_Terminate` are all error 438 without running it, whether the class wrote it Private or Public                                                                                                                                                                                                                                                                                                                                                |
| Naming an event from inside    | ✔️     | A method whose body is the bare name `Class_Initialize` runs the Initialize event again, and one naming `Class_Terminate` runs that                                                                                                                                                                                                                                                                                                                                                                                              |
| An error in Initialize         | ✔️     | The instance's Terminate still runs, and the Initialize error is the one the caller sees                                                                                                                                                                                                                                                                                                                                                                                                                                         |
| An error in Terminate          | ✔️     | It is the error the statement answers with, and the work after it continues                                                                                                                                                                                                                                                                                                                                                                                                                                                      |

---

## Execute and ExecuteGlobal

Every rule below was read off a fresh `cscript`, one family of cells per process, since the cells of a family
are meant to see what that family's earlier cells left behind. The probe, its capture and the comparison
script live in the temporary probe directory, and `src/__tests__/execute-and-executeglobal.test.ts` is the
durable copy of its 43 cells.

The families call helper procedures. The probe holds them above the family marker, and the comparison script and
the durable copy hold them as well. They have to stand in the source on both sides: without them every helper
call is a name that was never defined, `On Error Resume Next` swallows the error, and the cell compares the empty
value the name still had instead of an answer. Two scope cells were read that way for a while and looked like an
engine difference, while the engine agreed with Windows all along.

| Rule                                 | Status | Notes                                                                                                                                                                                                                                                                                                                                                                                                             |
| ------------------------------------ | ------ | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| The scope it runs in                 | ✔️     | `Execute` runs in the caller's scope, so a variable the caller holds is read and written: inside a Sub it reads that Sub's own locals, and at script level a name it writes stays for the statements after it                                                                                                                                                                                                     |
| A name it declares                   | ✔️     | `Dim` in the handed over text declares into the scope the text runs in, so a `Dim x` inside `Execute` in a Sub is that Sub's local and is gone when the Sub returns                                                                                                                                                                                                                                               |
| `ExecuteGlobal`                      | ✔️     | Its text runs in the script's global scope, so a name it declares is the script's own wherever it was called from                                                                                                                                                                                                                                                                                                 |
| `ExecuteGlobal` called from a Sub    | ✔️     | Measured: a Sub that runs `ExecuteGlobal "Dim globalOnly"` leaves `globalOnly` at script level, where the `Execute` form leaves the `Dim` of its own text in the Sub. A name the text writes reaches the script's own either way: a Sub that reads `moduleValue` through `Execute` answers 12 when the script set it to 11, and one that reads it through `ExecuteGlobal` answers 22 when the script set it to 21 |
| Sub and Function definitions         | ✔️     | A procedure the text defines is available to the statements after it, and redefining it replaces what was there before                                                                                                                                                                                                                                                                                            |
| A class definition                   | ✔️     | `Execute` may define a class, which `New` finds afterwards                                                                                                                                                                                                                                                                                                                                                        |
| An empty text                        | ✔️     | `Execute ""` runs and answers nothing                                                                                                                                                                                                                                                                                                                                                                             |
| A Null argument                      | ✔️     | `Execute Null` runs the empty text rather than raising                                                                                                                                                                                                                                                                                                                                                            |
| A syntax error in the text           | ✔️     | Error 1002, the same reading `Eval` gets                                                                                                                                                                                                                                                                                                                                                                          |
| A run time error in the text         | ✔️     | Error 11 for `x = 1 / 0`, which stops the handed over text where it failed                                                                                                                                                                                                                                                                                                                                        |
| `On Error Resume Next` in the caller | ✔️     | The caller's own run continues after the error, and the text that was handed over still stops at the error rather than resuming inside itself: measured, `Execute "result = ""before"" : x = 1 / 0 : result = ""after"""` under that handler answers 11 with result still "before", while the same text with its own handler written inside it answers 11 with result "after"                                     |
| `Err.Raise` in the text              | ✔️     | The number it raises is the one the caller sees                                                                                                                                                                                                                                                                                                                                                                   |
| `Exit` in the text                   | ✔️     | Error 1039, and it also stops the procedure the call stands in: measured, a Sub whose `Execute "Exit Sub"` answered 1039 does not run its own statements after the call, and a Function in that place answers the value it had reached, or Empty when it had not written its name yet. The same holds for `Exit For`, `Exit Do`, `Execute "Exit Function"` and `ExecuteGlobal`                                    |
| Before and after the text            | ✔️     | A statement before the call stands, and the text's own work up to the `Exit` does too when the call is at script level                                                                                                                                                                                                                                                                                            |

---

## The error a browser host hands a page

A VBScript failure inside a browser reaches the page as a real JavaScript error, which is the shape the browser
mode of `VbsEngine` reproduces. The reading of the error object was taken with an HTA rather than with Internet
Explorer itself: the automation object on a current Windows is created but answers every property with
`0x800706B5`, "the interface is unknown", while an HTA is the same Trident engine in the same document modes and
can write its own report. One HTA was built per document mode, and the reports are `mode-<n>-report.txt` in the
temporary probe directory beside the pages that wrote them.

The part a person reads was taken from Internet Explorer itself. Its automation object cannot be read back, but
it still starts the browser and still navigates, and `Start-Process iexplore.exe` is redirected to Microsoft Edge
on this Windows, where there is no VBScript at all. So the page was opened through the automation object and read
off the F12 console by eye.

What the document mode changes:

| `X-UA-Compatible`  | documentMode | VBScript | `ActiveXObject` | `addEventListener`, `JSON`, `Array.forEach`     |
| ------------------ | ------------ | -------- | --------------- | ----------------------------------------------- |
| `IE=7`             | 7            | runs     | there           | no                                              |
| `IE=8`             | 8            | runs     | there           | `JSON` and `querySelector` only                 |
| `IE=9`             | 9            | runs     | there           | all there                                       |
| `IE=10`            | 10           | runs     | there           | all there, and a caught error carries a `stack` |
| `IE=11`, `IE=edge` | 11           | **gone** | **gone**        | all there                                       |

The highest mode that still runs VBScript is 10, and the lowest that has the modern DOM is 9, so a page that
wants both VBScript and a modern debugger uses `IE=9` or `IE=10`. An HTA that names no mode gets 7.

What a caught error looks like, for a divide by zero, in modes 7 to 10:

| What                                                       | How it reads                                                                       |
| ---------------------------------------------------------- | ---------------------------------------------------------------------------------- |
| `typeof`                                                   | `object`                                                                           |
| `instanceof Error`                                         | `true`                                                                             |
| `name`                                                     | `Error`, and not an own enumerable property                                        |
| `number`                                                   | `-2146828277`, which is `0x800a000b`                                               |
| `description`, `message`                                   | both present and equal, and localized by the host: `被零除` under a Chinese locale |
| `line`, `column`, `lineNumber`, `columnNumber`, `fileName` | absent altogether, in every one of those modes                                     |
| `stack`                                                    | absent in 7 to 9; in 10 it is `Error` with the JavaScript frames                   |
| `for...in`                                                 | `message,description,number,name` in 7 to 9, `description,number,stack` in 10      |

`window.onerror` is passed exactly three arguments, `(message, url, line)`: no column and no error object. `url`
is the document's own URL, and `line` is the line in the **document** rather than in the script block, measured
by putting the failing statement on line 172 of the file and being told 172.

### What a person reading the browser sees

Measured in Internet Explorer 11 on Windows 11 with the page in document mode 8, on a page whose failing
statements sat at written down lines. Nothing on that page cancelled an error. The F12 **Console** lists one
failure as two lines:

```
SCRIPT11: 被零除
ie8-vbs-errors.html (35,3)
```

The prefix is `SCRIPT` and the plain VBScript error number, not the HRESULT: a divide by zero is `SCRIPT11`, a
missing member is `SCRIPT438`, a type mismatch is `SCRIPT13`, and `Err.Raise 4321` is `SCRIPT4321`. The
description is localized by the host. The second line names the file and then the position in brackets.

That position is the line **in the document**, counted from 1, and the column of the statement within that line.
It is the same number the third argument of `window.onerror` carries: four failures placed on known lines were
reported as 27, 36, 49 and 55, which is where those statements were.

The classic error dialog shows one thing more. Microsoft documents it as:

```
Line:<LineNumber>
Char:<CharacterNumber>
Error:<ErrorMessage>
Code: 0
URL: http://Webserver/page.htm
```

`Char` is the only place a column was given a page; `window.onerror` was passed none at all.

Nothing is added to the console for a failure that happened before the console was opened, so a failure while a
page loads is read only if F12 was already open.

### What VBScript itself could read

Nothing. Measured on a fresh `cscript`, the `Err` object answers five members and nothing besides: `Number`,
`Description`, `Source`, `HelpFile` and `HelpContext`. Every other name a script might reach for is error 438,
"Object doesn't support this property or method": `Line`, `Column`, `LineNumber`, `ColumnNumber`, `FileName`,
`StackTrace`, `Stack`, `CallStack`, `Procedure`, `Erl`, `Help` and `LastDllError`. `WScript` carries no stack
either, and asking it for `StackTrace`, `Stack`, `CallStack`, `Line` or `Column` is 438 as well.

So a VBScript program cannot look at where it was called from. The position exists only in what the host says at
the moment it reports a failure, which is why this engine keeps it on `engine.error` rather than on the thrown
error.

What this engine's own thrown error looks like beside those, for a divide by zero:

| What                         | This engine                                 | IE 7 to 9                          | IE 10                              |
| ---------------------------- | ------------------------------------------- | ---------------------------------- | ---------------------------------- |
| `number`                     | `0x800a000b`                                | `0x800a000b`                       | `0x800a000b`                       |
| `description`, `message`     | both, equal, in English                     | both, equal, localized by the host | both, equal, localized by the host |
| `name`                       | `Error`, inherited                          | `Error`, inherited                 | `Error`, inherited                 |
| `line`, `column`             | absent                                      | absent                             | absent                             |
| `stack`                      | the host's own, enumerable                  | absent                             | the host's own, enumerable         |
| `Object.keys` and `for...in` | `number,description,stack`                  | `message,description,number,name`  | `description,number,stack`         |
| `hasOwnProperty`             | `number`, `description`, `message`, `stack` | `number`, `description`, `message` | the same four                      |

So the thrown error is the shape document mode 10 handed JScript, and only that shape. It is not the later
modes' shape either: 7 to 9 kept no stack at all, and every mode up to 10 left `name` to the prototype.

Two things the host never put on the object are deliberately still absent:

- There is no `line` and no `column`. Internet Explorer reported the line of a failing script through the
  arguments of `window.onerror`, and those arguments describe the JavaScript throw site, which for a VBScript
  failure is inside the engine. The position is not lost: the MSScriptControl-style `engine.error` member carries
  `line` and `column` beside the same fields.
- The text stays English, as it is throughout this engine, where the host localizes it.

### Where `engine.error` counts its line

Internet Explorer counted the line in the **document**, so a page that reads a failure from this engine is given
the same number. Verified on the example page: the failing statement stands on line 135 of the HTML the server
hands out, and the engine reports 135.

A document does not keep its own line numbers, so the browser mode counts them from what it does keep. The
element is looked for in `documentElement.innerHTML` by its own `outerHTML`, and the newlines before it are
counted. The lines before the `<html>` tag are gone from the DOM, because the parser drops the whitespace there,
so they are counted back from the doctype and the comments the document node still holds. Three things are taken
to be so, and a page that breaks one of them is given a line that is short by the difference:

- the doctype, each comment before `<html>`, and the `<html>` tag occupy one line each, with no blank line
  between them,
- the whitespace between the `<html>` tag and `<head>` is one newline,
- and the block's text can be found in the written back document.

Should any of that fail, the line falls back to being counted within the script block, which is what it always
was. Node mode and `Execute` have no document to count in, so their positions stay counted within the script that
was handed to them.

---

## Non-Standard Extensions

These features are not part of the VBScript specification and do not exist in
Internet Explorer's VBScript engine. They are specific to this implementation
and are intended for Node.js/JavaScript interoperability.

| Feature                               | Status | Notes                                                                                     |
| ------------------------------------- | ------ | ----------------------------------------------------------------------------------------- |
| `New` with JavaScript constructors    | ✔️     | Resolves dotted names via `globalThis`; `globalThis[A][B][…]` path; `Reflect.construct`   |
| `GetRef` callbacks to JavaScript      | ✔️     | Wraps VBScript procedures as real JS callables for use in JS event APIs                   |
| Property access on JavaScript Proxies | ✔️     | Dot-notation get/set via Proxy traps; engine protocol methods only used as own properties |
| Host object access via `globalThis`   | ✔️     | All `globalThis` properties available as VBScript variables; case-insensitive             |

---

## VB6 / VBA7 Extensions

| Feature                                | Status | Notes                                                                                                                                                                                                                                                                                                                                            |
| -------------------------------------- | ------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `As Type` annotations                  | ✔️     | `Dim x As Integer` initializes to typed default; applies to variables, parameters, return types                                                                                                                                                                                                                                                  |
| `Enum...End Enum`                      | ✔️     | Global read-only constants; auto-increment from 0 or last explicit value; `Public`/`Private`                                                                                                                                                                                                                                                     |
| `Collection` class                     | ✔️     | `New Collection`; `Add`, `Remove`, `Item`, `Count`; 1-based index or string key. Not a class of `vbscript.dll`, which answers a runtime error for `New Collection`, so there is nothing to compare it against; it is kept because VB6 and VBA have it                                                                                            |
| `Case Is` in Select Case               | ✔️     | `Select Case x : Case Is > 5` runs here and is a compile error in `vbscript.dll`, which allows only a plain expression or a comma list. Kept because VB6 and VBA have it. A range (`Case 1 To 5`) is not accepted either way                                                                                                                     |
| `LongLong` (64-bit integer)            | ✔️     | Backed by `BigInt`; `CLngLng()` conversion; `VarType` = 20; arithmetic stays `LongLong`                                                                                                                                                                                                                                                          |
| `Val()` and `Str()`                    | ✔️     | Not functions of `vbscript.dll`, which answers error 13 for both: a direct call to a name it does not know raises error 500 and a call through `Eval` raises error 13. They are kept because VB6 and VBA have them. `Val` reads a leading number and stops at the first character that is not one, and `Str` writes a leading space for the sign |
| `CVar()` and `CVErr()`                 | ✔️     | Not functions of `vbscript.dll` either, which answers error 13 for every call to them. Kept because VB6 and VBA have them                                                                                                                                                                                                                        |
| `Type...End Type` (user-defined types) | ✔️     | Struct-like; fields typed to VB6 defaults; `Dim x As MyType` creates initialized instance                                                                                                                                                                                                                                                        |

---

## How the compatibility was measured

Every rule in this document that has a note beside it was read off a fresh `cscript` process, because
`SetLocale` in a loop is not reliable on Windows. The measurements are kept as development probes outside this
repository, in the temporary probe directory, since the brief asks that no probe or throwaway script goes into
the repository. Each area there is a `.vbs` probe that writes the Windows answers, a UTF-16LE capture, and a
comparison script that runs the same cells through this engine. `run-all-compares.mjs` runs them all and prints
one line with the number of cells that differ, which is the regression net a change in a shared part of the
engine is checked against. The durable net is this repository's own test suite, which holds a cell for every
rule that was measured and fixed.

Three things guide the probes: an expression is always written out, since `Eval` is not faithful and once
reported 24 differences that a written out probe does not have; a label that is missing from a capture is not a
difference, since it means the probe could not write the line; and the events probe runs one cell per process,
because an instance that outlives its cell is still alive when the next cell starts.

Every events cell is therefore a whole script of its own, and the end of that script is read as well, since
Windows runs a Terminate event there for an instance the script never dropped.

## What is not measured yet

The measurements cover the language, the built-in functions, the locales, the class events and the host calls
this engine makes. These parts have no comparison yet, and are the honest limit of the claim that this engine
matches `vbscript.dll`:

- **The order of the releases at the end of a script.** The events probe reads which instance is freed when a
  script ends, and both sides free every one of them, but not in the same order. The order follows the order
  the host keeps its globals in, which neither `cscript` nor this engine writes down, so only the set of
  releases is compared.
- **One `Execute` cell row.** The row the `execute.error.exitInsideTextStops` cell writes comes after the
  statement that stops the script, so the engine never reaches it while the probe's capture holds it. The rule
  itself is measured and followed: a script that calls `Execute "Exit Sub"` at its own level stops there.
- **Parts of the `Err` object.** `Err.Source` has no test, `Err.Description` and `Err.HelpContext` have one
  each, and the arguments of `Err.Raise` are not compared.
- **The error a browser host hands a page.** The section above says what Internet Explorer handed a page, read
  off an HTA by hand, and the browser mode of this engine follows it. There is no comparison script for it: the
  shape is a host's, not `vbscript.dll`'s, and the automation interface of Internet Explorer on a current
  Windows cannot be read from outside anyway.
- **The text of an error.** The probes compare the error number only, since the description is localized.
- **COM and the host.** `RegExp`, `Scripting.Dictionary`, `Scripting.FileSystemObject` and `WScript.Shell` are
  out of scope for this engine, which the brief keeps to the `node-ps1-dotnet` project. `WScript` itself is
  covered by this repository's own tests rather than by a `cscript` comparison.
- **One flagged row above**, the `.wsf` `<object>`, `<reference>` and `<resource>` hoisting, which is marked
  with a warning rather than a tick.
