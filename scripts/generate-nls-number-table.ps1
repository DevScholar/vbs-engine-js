<#
  Generates the Windows NLS number table the engine needs to write numbers the way vbscript.dll does.

  Why a table instead of asking Intl:
    * Intl describes a number with CLDR's marks and CLDR's patterns. Windows uses its own two sets of
      marks (one for plain numbers, one for money) and its own negative patterns, and the two disagree
      for most locales. de-AT is the plain example: Windows groups 1234567.89 with U+00A0 in
      FormatNumber but with a full stop in FormatCurrency, and Intl has one set only.
    * Some locales have no ICU data at all. 1073 (Tsonga) is the engine's only such LCID: Intl resolves
      'ts' to the host language, so the marks it hands back belong to another language entirely.

  Independence from the generating machine:
    * Every value is read with an explicit locale name, with LOCALE_NOUSEROVERRIDE, exactly as the date
      table generator does, so neither the display language nor Regional Settings can leak in.

  Usage:
    pwsh -File scripts/generate-nls-number-table.ps1
#>
[CmdletBinding()]
param(
  [string]$LocaleSource = 'src/builtins/locale.ts',
  [string]$Out = 'src/runtime/nls-number-table.generated.ts'
)

$ErrorActionPreference = 'Stop'

Add-Type -TypeDefinition @'
using System;
using System.Text;
using System.Runtime.InteropServices;
public class NlsNumberRead {
  [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
  static extern int GetLocaleInfoEx(string name, uint type, StringBuilder data, int cch);
  [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
  static extern int LCIDToLocaleName(uint lcid, StringBuilder name, int cch, uint flags);
  public static string Get(string name, uint type) {
    var buffer = new StringBuilder(512);
    // 0x80000000 = LOCALE_NOUSEROVERRIDE
    int length = GetLocaleInfoEx(name, type | 0x80000000u, buffer, buffer.Capacity);
    return length > 0 ? buffer.ToString() : "";
  }
  public static string NameOf(uint lcid) {
    var buffer = new StringBuilder(128);
    return LCIDToLocaleName(lcid, buffer, buffer.Capacity, 0) > 0 ? buffer.ToString() : "";
  }
}
'@ -Language CSharp

$LOCALE_SDECIMAL        = [uint32]0x0000000E
$LOCALE_STHOUSAND       = [uint32]0x0000000F
$LOCALE_SGROUPING       = [uint32]0x00000010
$LOCALE_IDIGITS         = [uint32]0x00000011
$LOCALE_ILZERO          = [uint32]0x00000012
$LOCALE_SNATIVEDIGITS   = [uint32]0x00000013
$LOCALE_SCURRENCY       = [uint32]0x00000014
$LOCALE_SINTLSYMBOL     = [uint32]0x00000015
$LOCALE_SMONDECIMALSEP  = [uint32]0x00000016
$LOCALE_SMONTHOUSANDSEP = [uint32]0x00000017
$LOCALE_SMONGROUPING    = [uint32]0x00000018
$LOCALE_ICURRDIGITS     = [uint32]0x00000019
$LOCALE_ICURRENCY       = [uint32]0x0000001B
$LOCALE_INEGCURR        = [uint32]0x0000001C
$LOCALE_SPOSITIVESIGN   = [uint32]0x00000050
$LOCALE_SNEGATIVESIGN   = [uint32]0x00000051
$LOCALE_INEGNUMBER      = [uint32]0x00001010

if (-not (Test-Path $LocaleSource)) { throw "locale map not found: $LocaleSource" }
$source = Get-Content $LocaleSource -Raw
$seen = @{}
foreach ($match in [regex]::Matches($source, "\['([^']+)',\s*(\d+)\]")) {
  $lcid = [int]$match.Groups[2].Value
  if (-not $seen.ContainsKey($lcid)) { $seen[$lcid] = $match.Groups[1].Value }
}

# Windows writes every string in this table. Each one is written so the generated file stays ASCII and
# cannot end the template literal it lives in: a real backslash, backtick or dollar sign is escaped, and
# everything outside printable ASCII becomes \uXXXX. One locale's money fraction mark is a backtick, and
# another's currency symbol ends in a dollar sign, so neither is hypothetical.
function EscapeForDataFile([string]$text) {
  $builder = New-Object System.Text.StringBuilder
  foreach ($character in $text.ToCharArray()) {
    $code = [int]$character
    if ($character -eq '\') { [void]$builder.Append('\\') }
    elseif ($character -eq [char]96) { [void]$builder.Append('\`') }
    elseif ($character -eq '$') { [void]$builder.Append('\u0024') }
    elseif ($code -ge 32 -and $code -lt 127) { [void]$builder.Append($character) }
    else { [void]$builder.Append(('\u{0:x4}' -f $code)) }
  }
  $builder.ToString()
}

function ReadNumber([string]$text) {
  $parsed = 0
  if ([int]::TryParse($text, [ref]$parsed)) { return $parsed }
  return -1
}

$rows = @()
$unresolved = @()
foreach ($lcid in ($seen.Keys | Sort-Object)) {
  $name = [NlsNumberRead]::NameOf([uint32]$lcid)
  if ($name -eq '') { $unresolved += "$lcid($($seen[$lcid]))"; continue }
  $rows += [pscustomobject]@{
    lcid       = $lcid
    decimal    = EscapeForDataFile ([NlsNumberRead]::Get($name, $LOCALE_SDECIMAL))
    thousand   = EscapeForDataFile ([NlsNumberRead]::Get($name, $LOCALE_STHOUSAND))
    grouping   = EscapeForDataFile ([NlsNumberRead]::Get($name, $LOCALE_SGROUPING))
    digits     = ReadNumber ([NlsNumberRead]::Get($name, $LOCALE_IDIGITS))
    leadingZero = ReadNumber ([NlsNumberRead]::Get($name, $LOCALE_ILZERO))
    negative   = ReadNumber ([NlsNumberRead]::Get($name, $LOCALE_INEGNUMBER))
    negativeSign = EscapeForDataFile ([NlsNumberRead]::Get($name, $LOCALE_SNEGATIVESIGN))
    positiveSign = EscapeForDataFile ([NlsNumberRead]::Get($name, $LOCALE_SPOSITIVESIGN))
    symbol     = EscapeForDataFile ([NlsNumberRead]::Get($name, $LOCALE_SCURRENCY))
    code       = EscapeForDataFile ([NlsNumberRead]::Get($name, $LOCALE_SINTLSYMBOL))
    currencyDigits = ReadNumber ([NlsNumberRead]::Get($name, $LOCALE_ICURRDIGITS))
    currencyPosition = ReadNumber ([NlsNumberRead]::Get($name, $LOCALE_ICURRENCY))
    negativeCurrency = ReadNumber ([NlsNumberRead]::Get($name, $LOCALE_INEGCURR))
    monDecimal = EscapeForDataFile ([NlsNumberRead]::Get($name, $LOCALE_SMONDECIMALSEP))
    monThousand = EscapeForDataFile ([NlsNumberRead]::Get($name, $LOCALE_SMONTHOUSANDSEP))
    monGrouping = EscapeForDataFile ([NlsNumberRead]::Get($name, $LOCALE_SMONGROUPING))
    nativeDigits = EscapeForDataFile ([NlsNumberRead]::Get($name, $LOCALE_SNATIVEDIGITS))
  }
}

if ($unresolved.Count -gt 0) { Write-Warning ("locales with no NLS data: " + ($unresolved -join ', ')) }

$packed = ($rows | ForEach-Object {
  "$($_.lcid)|$($_.decimal)|$($_.thousand)|$($_.grouping)|$($_.digits)|$($_.leadingZero)|$($_.negative)|$($_.negativeSign)|$($_.positiveSign)|$($_.symbol)|$($_.code)|$($_.currencyDigits)|$($_.currencyPosition)|$($_.negativeCurrency)|$($_.monDecimal)|$($_.monThousand)|$($_.monGrouping)|$($_.nativeDigits)"
}) -join "`n"

$build = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion').CurrentBuildNumber
$display = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction SilentlyContinue).DisplayVersion
$backtick = [char]96

$header = @"
// Generated by scripts/generate-nls-number-table.ps1 - do not edit by hand.
//
// Windows NLS number and money formats, read with GetLocaleInfoEx(locale, LOCALE_NOUSEROVERRIDE):
//   LOCALE_SDECIMAL, LOCALE_STHOUSAND, LOCALE_SGROUPING, LOCALE_IDIGITS, LOCALE_ILZERO,
//   LOCALE_INEGNUMBER, LOCALE_SNEGATIVESIGN, LOCALE_SPOSITIVESIGN, LOCALE_SCURRENCY,
//   LOCALE_SINTLSYMBOL, LOCALE_ICURRDIGITS, LOCALE_ICURRENCY, LOCALE_INEGCURR,
//   LOCALE_SMONDECIMALSEP, LOCALE_SMONTHOUSANDSEP, LOCALE_SMONGROUPING, LOCALE_SNATIVEDIGITS
//
// Generated on Windows build $build $display, $((Get-Date).ToUniversalTime().ToString('yyyy-MM-dd'))
// Rows: lcid|decimal|thousand|grouping|digits|leadingZero|negative|negativeSign|positiveSign|symbol|code|currencyDigits|currencyPosition|negativeCurrency|monDecimal|monThousand|monGrouping|nativeDigits
//
// Money has its own two marks in Windows (LOCALE_SMONDECIMALSEP, LOCALE_SMONTHOUSANDSEP), and they are
// not always the ones LOCALE_SDECIMAL and LOCALE_STHOUSAND hold: de-AT groups a plain number with U+00A0
// and money with a full stop. A locale's row therefore carries both sets.
//
// Every value came from an explicit locale name with the user override disabled, so neither this
// machine's display language nor its Regional Settings can appear in the data.

/** One locale's Windows number format. Every field is data read from Windows, never inferred. */
export interface NlsNumberFormat {
  /** LOCALE_SDECIMAL: the mark between a whole number and its fraction. */
  decimal: string;
  /** LOCALE_STHOUSAND: the mark between groups in a plain number. */
  thousand: string;
  /** LOCALE_SGROUPING: group sizes from the right, e.g. "3;0" (every three, repeating). */
  grouping: string;
  /** LOCALE_IDIGITS: the fraction count FormatNumber uses when the caller does not name one. */
  digits: number;
  /** LOCALE_ILZERO: 1 writes a zero before a fraction ("0,5"), 0 leaves it out (",5"). */
  leadingZero: number;
  /** LOCALE_INEGNUMBER: 0 "(1)", 1 "-1", 2 "- 1", 3 "1-", 4 "1 -". */
  negativeNumber: number;
  /** LOCALE_SNEGATIVESIGN. */
  negativeSign: string;
  /** LOCALE_SPOSITIVESIGN. */
  positiveSign: string;
  /** LOCALE_SCURRENCY: the symbol FormatCurrency writes. */
  currencySymbol: string;
  /** LOCALE_SINTLSYMBOL: the ISO code, e.g. "EUR", as Windows reports it. */
  currencyCode: string;
  /** LOCALE_ICURRDIGITS: the fraction count the currency itself uses. */
  currencyDigits: number;
  /** LOCALE_ICURRENCY: 0 symbol first, 1 symbol last, 2 symbol first with a space, 3 symbol last with a space. */
  currencyPosition: number;
  /** LOCALE_INEGCURR: which of the sixteen ways a negative amount is written. */
  negativeCurrency: number;
  /** LOCALE_SMONDECIMALSEP: the fraction mark money uses. */
  monDecimal: string;
  /** LOCALE_SMONTHOUSANDSEP: the group mark money uses. */
  monThousand: string;
  /** LOCALE_SMONGROUPING: the group sizes money uses. */
  monGrouping: string;
  /** LOCALE_SNATIVEDIGITS: the locale's own digits, ten characters, ASCII for most locales. */
  nativeDigits: string;
}

const PACKED = $backtick$packed$backtick;

const FIELDS = PACKED.split('\n').map((row) => row.split('|'));

/** The Windows number format for an LCID, or undefined when the table has no row for it. */
export const NLS_NUMBER_FORMATS: ReadonlyMap<number, NlsNumberFormat> = new Map(
  FIELDS.map((row) => [
    Number(row[0]),
    {
      decimal: row[1],
      thousand: row[2],
      grouping: row[3],
      digits: Number(row[4]),
      leadingZero: Number(row[5]),
      negativeNumber: Number(row[6]),
      negativeSign: row[7],
      positiveSign: row[8],
      currencySymbol: row[9],
      currencyCode: row[10],
      currencyDigits: Number(row[11]),
      currencyPosition: Number(row[12]),
      negativeCurrency: Number(row[13]),
      monDecimal: row[14],
      monThousand: row[15],
      monGrouping: row[16],
      nativeDigits: row[17],
    },
  ])
);

/** The Windows number format for an LCID, or undefined when the table has no row for it. */
export function nlsNumberFormat(lcid: number): NlsNumberFormat | undefined {
  return NLS_NUMBER_FORMATS.get(lcid);
}
"@

[System.IO.File]::WriteAllText((Join-Path (Get-Location) $Out), $header, (New-Object System.Text.UTF8Encoding($false)))
Write-Output "rows written : $($rows.Count) of $($seen.Count) locales"
Write-Output "output       : $Out"
$rows | Select-Object -First 3 | ForEach-Object { "  $($_.lcid)|$($_.decimal)|$($_.thousand)|$($_.grouping)|$($_.digits)|$($_.currencyPosition)|$($_.negativeCurrency)|$($_.monThousand)" }
