# Reads Windows' own name for every LCID the engine knows (LCIDToLocaleName) and hands it to
# generate-locale-names.mjs, which keeps a name only where ICU resolves it exactly.
#
#   pwsh -File scripts/generate-locale-names.ps1
#
# The result, src/builtins/locale-data.generated.ts, is committed: the engine ships the names as data,
# because a browser has no Win32 to ask. Regenerating on a different Windows build, or with a different
# ICU, shows up as a reviewable diff.
[CmdletBinding()]
param(
  [string]$LocaleSource = 'src/builtins/locale.ts',
  [string]$Out = 'src/builtins/locale-data.generated.ts'
)

$ErrorActionPreference = 'Stop'
if (-not (Test-Path $LocaleSource)) { throw "locale map not found: $LocaleSource" }

Add-Type -TypeDefinition @'
using System;
using System.Text;
using System.Runtime.InteropServices;
public class LocaleName {
  [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
  static extern int LCIDToLocaleName(uint lcid, StringBuilder name, int cch, uint flags);
  public static string Of(uint lcid) {
    var sb = new StringBuilder(128);
    return LCIDToLocaleName(lcid, sb, sb.Capacity, 0) > 0 ? sb.ToString() : "";
  }
}
'@ -Language CSharp

$lcids = [regex]::Matches((Get-Content $LocaleSource -Raw), "\['([^']+)',\s*(\d+)\]") |
  ForEach-Object { [int]$_.Groups[2].Value } |
  Sort-Object -Unique

$lines = foreach ($lcid in $lcids) {
  $name = [LocaleName]::Of([uint32]$lcid)
  if ($name) { "$lcid|$name" }
}

$namesPath = Join-Path ([System.IO.Path]::GetTempPath()) "lcid-names-$([guid]::NewGuid().ToString('N')).txt"
[System.IO.File]::WriteAllLines($namesPath, $lines, (New-Object System.Text.UTF8Encoding($false)))

$env:WINDOWS_BUILD = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion').CurrentBuildNumber
try {
  & node (Join-Path $PSScriptRoot 'generate-locale-names.mjs') $namesPath $LocaleSource $Out
  if ($LASTEXITCODE -ne 0) { throw "generate-locale-names.mjs failed with exit code $LASTEXITCODE" }
} finally {
  Remove-Item $namesPath -ErrorAction SilentlyContinue
}
