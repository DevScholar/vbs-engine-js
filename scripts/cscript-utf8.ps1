# cscript-utf8.ps1
#
# Run cscript and re-emit its output so a UTF-8 terminal renders it correctly.
#
# Windows Script Host (cscript / wscript) always writes its console output in the
# system ANSI code page (the "Language for non-Unicode programs" setting, a.k.a.
# the ACP). Examples: 936 for Simplified Chinese, 932 for Japanese, 1251 for
# Russian, 1252 for Western European, 950 for Traditional Chinese. A UTF-8
# terminal decodes those bytes with the wrong table and the text turns into
# mojibake.
#
# This wrapper tells .NET to decode cscript's bytes with the SAME ANSI code page
# Windows actually used, then re-encodes the text as UTF-8 for the terminal. The
# code page is read at runtime via GetACP(), so every locale works without any
# hard-coded value and no action is required from the tester.
#
# The script file's own encoding is irrelevant: WSH reads ANSI or UTF-16 source,
# but it always EMITS ANSI console output. That is the thing being fixed here.
#
# Usage:
#   powershell -NoProfile -File scripts/cscript-utf8.ps1 //nologo //E:vbscript script.vbs
#   powershell -NoProfile -File scripts/cscript-utf8.ps1 script.vbs arg1 arg2
#
# The exit code of cscript is preserved.

[CmdletBinding()]
param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$CscriptArgs
)

function Get-AnsiCodePage {
    # The system ANSI code page, which is exactly what WSH uses for 8-bit console
    # output. GetACP() is the source of truth: it is independent of both the
    # console code page (which `chcp` can change) and the user's display language.
    #
    # Deliberately NOT [Text.Encoding]::Default (UTF-8 on PowerShell 7+/ .NET Core)
    # and NOT CurrentCulture.TextInfo.ANSICodePage (follows the user culture, which
    # the user can set independently of the system ACP).
    $signature = '[DllImport("kernel32.dll")] public static extern int GetACP();'
    $native = Add-Type -Namespace Win32 -Name Native -MemberDefinition $signature -PassThru
    return $native::GetACP()
}

$ansi = [System.Text.Encoding]::GetEncoding((Get-AnsiCodePage))
$utf8 = New-Object System.Text.UTF8Encoding($false)   # no BOM

# Decode cscript's native stdout/stderr with the ANSI code page.
[Console]::OutputEncoding = $ansi

$output = & cscript @CscriptArgs 2>&1 | ForEach-Object { $_.ToString() }
$exitCode = $LASTEXITCODE

# Re-emit as UTF-8 so the terminal shows the text correctly.
[Console]::OutputEncoding = $utf8
$output | ForEach-Object { Write-Output $_ }

exit $exitCode
