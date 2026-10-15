' Dumps the names a locale writes, both on their own and inside a date, so the engine can carry the ones
' Intl spells differently (or cannot spell at all).
'
' Three sets, because the two forms are not the same question: MonthName answers with the standalone form
' and a date's picture string with the one CLDR calls the format form (German "Mär" against "März", Catalan
' "gener" against "de gener"). The dated ones are read out of a long date, which is the only place Windows
' writes them.
'
'   nm/ns/dw/ds : MonthName and WeekdayName, long and short, on their own.
'   md/wd       : a long date for the 15th of every Gregorian month and for 5..11 January, which is where a
'                 Gregorian month name and a weekday name are read from.
'   id          : a long date every 28 days for 14 samples, which is what covers the months of a locale
'                 whose calendar is not the Gregorian one: the Islamic year is about 354 days, so its
'                 months cannot be reached from a Gregorian month.
'
' One process per locale, with SetLocale: a fresh script is what a script sees, and a loop that changed the
' locale mid-run once reported values a fresh process does not.
'
' The file is written through FileSystemObject.CreateTextFile(path, True, True), which is UTF-16LE, so the
' console's code page cannot lose a character on the way out.
'
' Usage: cscript //nologo //E:vbscript dump-locale-names.vbs <lcid> <out-file>
Option Explicit

Dim localeId, outputPath, fileSystem, outputFile, nameIndex
localeId = CLng(WScript.Arguments(0))
outputPath = WScript.Arguments(1)
SetLocale localeId

Set fileSystem = CreateObject("Scripting.FileSystemObject")
Set outputFile = fileSystem.CreateTextFile(outputPath, True, True)

For nameIndex = 1 To 12
  outputFile.WriteLine "nm" & nameIndex & "=" & MonthName(nameIndex)
  outputFile.WriteLine "ns" & nameIndex & "=" & MonthName(nameIndex, True)
Next

For nameIndex = 1 To 7
  outputFile.WriteLine "dw" & nameIndex & "=" & WeekdayName(nameIndex)
  outputFile.WriteLine "ds" & nameIndex & "=" & WeekdayName(nameIndex, True)
Next

For nameIndex = 1 To 12
  outputFile.WriteLine "md" & nameIndex & "=" & FormatDateTime(DateSerial(2026, nameIndex, 15), 1)
Next

For nameIndex = 0 To 6
  outputFile.WriteLine "wd" & (nameIndex + 1) & "=" & FormatDateTime(DateSerial(2026, 1, 5) + nameIndex, 1)
Next

For nameIndex = 0 To 13
  outputFile.WriteLine "id" & (nameIndex + 1) & "=" & FormatDateTime(DateSerial(2026, 1, 5) + nameIndex * 28, 1)
Next

outputFile.Close
WScript.Echo "ok " & localeId
