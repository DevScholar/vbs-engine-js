<#
  Generates the Windows NLS date/time table the engine needs to reproduce VBScript's CStr output.

  Why a generated table instead of calling ICU at run time:
    * VBScript's OLE/NLS layer formats through Windows NLS, and the two disagree for most of the
      locales the engine knows (measured: 46 of 133 agree).
    * Neither source promises stability (Microsoft: "does not guarantee data stability"; Unicode/ICU:
      "Locale data may be updated or changed, yielding different results for ... formatting").
      A generated table is a snapshot: the output it produces cannot drift under the engine.

  Independence from the generating machine's language:
    * Every value is read with an EXPLICIT locale name, so the system/user default language cannot
      substitute for it.
    * LOCALE_NOUSEROVERRIDE is passed so a user's Regional Settings cannot substitute either.
      (Microsoft notes overrides otherwise apply to the system default locale only, so this matters
      for at most one row - but the flag makes that row deterministic too.)
    * The Windows build and the flags used are recorded in the output header, so a regeneration that
      would change behaviour is visible as a diff.

  Usage:
    pwsh -File tools/generate-nls-table.ps1 -LocaleMapPath src/builtins/locale.ts -Out nls-table.generated.ts
#>
[CmdletBinding()]
param(
  [string]$LocaleMapPath = 'src/builtins/locale.ts',
  [string]$Out = 'nls-table.generated.ts'
)

$ErrorActionPreference = 'Stop'

Add-Type -TypeDefinition @'
using System;
using System.Text;
using System.Runtime.InteropServices;
public class NlsRead {
  [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
  static extern int GetLocaleInfoEx(string name, uint type, StringBuilder data, int cch);
  [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
  static extern int LCIDToLocaleName(uint lcid, StringBuilder name, int cch, uint flags);
  public static string Get(string name, uint type) {
    var sb = new StringBuilder(512);
    // 0x80000000 = LOCALE_NOUSEROVERRIDE
    int n = GetLocaleInfoEx(name, type | 0x80000000u, sb, sb.Capacity);
    return n > 0 ? sb.ToString() : "";
  }
  /** Windows' own name for an LCID. The engine's map is handwritten and is not always the same name,
   *  so asking by a wrong name silently reads a different locale's data. */
  public static string NameOf(uint lcid) {
    var sb = new StringBuilder(128);
    return LCIDToLocaleName(lcid, sb, sb.Capacity, 0) > 0 ? sb.ToString() : "";
  }
  /** Deliberately read by name only, to show what a wrong name would have returned. */
  public static string GetByPlainName(string name, uint type) {
    var sb = new StringBuilder(512);
    int n = GetLocaleInfoEx(name, type | 0x80000000u, sb, sb.Capacity);
    return n > 0 ? sb.ToString() : "";
  }
}
'@ -Language CSharp

$LOCALE_SSHORTDATE    = [uint32]0x0000001F
$LOCALE_SLONGDATE     = [uint32]0x00000020
$LOCALE_SSHORTTIME    = [uint32]0x00000079
$LOCALE_STIMEFORMAT   = [uint32]0x00001003
$LOCALE_S1159         = [uint32]0x00000028
$LOCALE_S2359         = [uint32]0x00000029
$LOCALE_ICALENDARTYPE = [uint32]0x00001009

# The engine's own locale map is the single source of truth for which locales to cover.
if (-not (Test-Path $LocaleMapPath)) { throw "locale map not found: $LocaleMapPath" }
$source = Get-Content $LocaleMapPath -Raw
$seen = @{}
foreach ($m in [regex]::Matches($source, "\['([^']+)',\s*(\d+)\]")) {
  $lcid = [int]$m.Groups[2].Value
  if (-not $seen.ContainsKey($lcid)) { $seen[$lcid] = $m.Groups[1].Value }
}

$rows = @()
$unresolved = @()
$renamed = @()
foreach ($lcid in ($seen.Keys | Sort-Object)) {
  # Windows' own name for the LCID, not the engine's handwritten one: GetLocaleInfoEx(name) reads a
  # different locale when the name is wrong, and the engine's map has such entries.
  $name = [NlsRead]::NameOf([uint32]$lcid)
  if ($name -eq '') { $unresolved += "$lcid($($seen[$lcid]))"; continue }
  if ($name -ne $seen[$lcid]) { $renamed += "$lcid : engine '$($seen[$lcid])' -> Windows '$name'" }
  $short = [NlsRead]::Get($name, $LOCALE_SSHORTDATE)
  $long  = [NlsRead]::Get($name, $LOCALE_SLONGDATE)
  $time  = [NlsRead]::Get($name, $LOCALE_STIMEFORMAT)
  $stime = [NlsRead]::Get($name, $LOCALE_SSHORTTIME)
  if ($short -eq '' -or $time -eq '') { $unresolved += "$lcid($name)"; continue }
  if ($long -eq '') { $unresolved += "$lcid($name) longDate empty" }
  if ($stime -eq '') { $unresolved += "$lcid($name) shortTime empty" }
  $rows += [pscustomobject]@{
    lcid   = $lcid
    short  = $short
    long   = $long
    time   = $time
    stime  = $stime
    am     = [NlsRead]::Get($name, $LOCALE_S1159)
    pm     = [NlsRead]::Get($name, $LOCALE_S2359)
    cal    = [NlsRead]::Get($name, $LOCALE_ICALENDARTYPE)
  }
}

if ($renamed.Count -gt 0) {
  Write-Warning ("LCIDs whose Windows name differs from the engine's map (" + $renamed.Count + "):")
  $renamed | ForEach-Object { Write-Warning "  $_" }
}
if ($unresolved.Count -gt 0) { Write-Warning ("locales with no NLS data: " + ($unresolved -join ', ')) }

$packed = ($rows | ForEach-Object { "$($_.lcid)|$($_.short)|$($_.long)|$($_.time)|$($_.stime)|$($_.am)|$($_.pm)|$($_.cal)" }) -join "`n"
$packed = $packed.Replace('\', '\\').Replace('`', '\`').Replace('${', '\${')
# A real U+202F, escaped: the generated file stays ASCII and the irregular-whitespace lint rule stays quiet.
$packed = $packed.Replace([string][char]0x202F, '\u202f')

$build = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion').CurrentBuildNumber
$display = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction SilentlyContinue).DisplayVersion
# A real backtick: writing one directly in a double-quoted here-string would be eaten as an escape.
$bq = [char]96
$header = @"
// Generated by tools/generate-nls-table.ps1 - do not edit by hand.
//
// Windows NLS date/time patterns, read with GetLocaleInfoEx(locale, LOCALE_NOUSEROVERRIDE):
//   LOCALE_SSHORTDATE, LOCALE_SLONGDATE, LOCALE_STIMEFORMAT, LOCALE_SSHORTTIME,
//   LOCALE_S1159, LOCALE_S2359, LOCALE_ICALENDARTYPE
//
// Generated on Windows build $build $display, $((Get-Date).ToUniversalTime().ToString('yyyy-MM-dd'))
// Rows: lcid|shortDate|longDate|timeFormat|shortTime|am|pm|calendarType
//
// Every value came from an explicit locale name with the user override disabled, so neither this
// machine's display language nor its Regional Settings can appear in the data. Regenerating on a
// different Windows build changes this file, and that diff is the review surface.

export interface NlsDateFormat {
  /** LOCALE_SSHORTDATE, a Windows date picture string (e.g. "dd.MM.yyyy"). */
  shortDate: string;
  /** LOCALE_SLONGDATE, the long date picture string with weekday and month names (e.g. "dddd, d. MMMM yyyy"). */
  longDate: string;
  /** LOCALE_STIMEFORMAT, a Windows time picture string (e.g. "HH:mm:ss"). */
  timeFormat: string;
  /** LOCALE_SSHORTTIME, the short time picture string (e.g. "HH:mm"). */
  shortTime: string;
  /** LOCALE_S1159 / LOCALE_S2359, the AM and PM designators ('' where the locale has none). */
  am: string;
  pm: string;
  /** LOCALE_ICALENDARTYPE: 1 = Gregorian, 6 = Hijri, 7 = Thai Buddhist, ... */
  calendarType: number;
}

const PACKED = $bq$packed$bq;

export const NLS_DATE_FORMATS: ReadonlyMap<number, NlsDateFormat> = new Map(
  PACKED.split('\n').map(row => {
    const [lcid, shortDate, longDate, timeFormat, shortTime, am, pm, calendarType] = row.split('|');
    return [
      Number(lcid),
      { shortDate, longDate, timeFormat, shortTime, am, pm, calendarType: Number(calendarType) },
    ] as const;
  })
);

/** The NLS formats for an LCID, or undefined when the table has no row for it. */
export function nlsDateFormat(lcid: number): NlsDateFormat | undefined {
  return NLS_DATE_FORMATS.get(lcid);
}
"@

[System.IO.File]::WriteAllText((Join-Path (Get-Location) $Out), $header, (New-Object System.Text.UTF8Encoding($false)))

$bytes = [System.IO.File]::ReadAllBytes((Join-Path (Get-Location) $Out))
Write-Output "rows written : $($rows.Count) of $($seen.Count) locales"
Write-Output "output       : $Out ($($bytes.Length) bytes)"
Write-Output "sample       :"
$rows | Select-Object -First 4 | ForEach-Object { "  $($_.lcid)|$($_.short)|$($_.time)|$($_.am)|$($_.pm)|cal=$($_.cal)" }
