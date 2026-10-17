import type { ExecFileSyncOptions } from "node:child_process";

/**
 * Node-only helpers shared by the MsgBox and InputBox GUI backends.
 *
 * The library build targets the browser/neutral platform and is deliberately free
 * of Node built-ins. `execFileSync` is therefore NOT statically imported — that
 * would pull `node:child_process` into the browser bundle. Instead it is fetched at
 * runtime via `process.getBuiltinModule`, which is a plain method on the `process`
 * object (no `import`/`require`), so the browser build never touches the module.
 * The `node:child_process` reference here is a type-only import, erased at build.
 */

type ExecFileSync = (
  file: string,
  args?: string[],
  options?: ExecFileSyncOptions,
) => Buffer | string;

let cached: ExecFileSync | null | undefined;

export function getExecFileSync(): ExecFileSync | null {
  if (cached !== undefined) return cached;
  if (
    typeof process === "undefined" ||
    typeof process.getBuiltinModule !== "function"
  ) {
    cached = null;
    return cached;
  }
  try {
    cached = process.getBuiltinModule("node:child_process")
      .execFileSync as ExecFileSync;
  } catch {
    cached = null;
  }
  return cached;
}

/** Quote a value as a PowerShell single-quoted string literal ('' escapes a quote). */
export function psQuote(s: string): string {
  return "'" + s.replace(/'/g, "''") + "'";
}

/**
 * PowerShell preamble that brings a dialog to the foreground. A bare
 * `Interaction.InputBox` / `MessageBox.Show` has no owner window, so it can sink
 * behind other windows. We create an off-screen, top-most `Form` as the owner so
 * the native dialog is owned by (and inherits the foreground of) that form.
 *
 * `Activate()` alone is unreliable here: each dialog runs in its OWN powershell
 * process, and Windows' foreground-lock will refuse a non-foreground process that
 * merely calls Activate(). The canonical workaround is the `AttachThreadInput` +
 * `SetForegroundWindow` pair, which temporarily links our thread to the current
 * foreground thread before forcing the window up. This must run on every dialog
 * (MsgBox and InputBox alike), or a MsgBox→InputBox sequence will leave the second
 * dialog behind the console.
 *
 * `[Console]::OutputEncoding=UTF8` forces PowerShell 5.1 to emit UTF-8 instead of
 * the system OEM code page (GBK on Chinese Windows), so returned text round-trips
 * byte-for-byte through Node's utf8 decoding.
 */
const PS_PREAMBLE = `
[Console]::OutputEncoding=[System.Text.Encoding]::UTF8;
$ProgressPreference = 'SilentlyContinue';
Add-Type -AssemblyName System.Windows.Forms;
Add-Type -AssemblyName System.Drawing;
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class Native {
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
    [DllImport("kernel32.dll")] public static extern uint GetCurrentThreadId();
    [DllImport("user32.dll")] public static extern bool AttachThreadInput(uint a, uint b, bool f);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int c);
}
'@;
$owner = New-Object System.Windows.Forms.Form;
$owner.ShowInTaskbar = $false;
$owner.FormBorderStyle = 'None';
$owner.StartPosition = 'Manual';
$owner.Location = New-Object System.Drawing.Point(-32000,-32000);
$owner.Size = New-Object System.Drawing.Size(1,1);
$owner.TopMost = $true;
$owner.Show();
$h = $owner.Handle;
[Native]::ShowWindow($h, 5) | Out-Null;
$fg = [Native]::GetForegroundWindow();
$pidOut = [uint32]0;
$fgThread = [Native]::GetWindowThreadProcessId($fg, [ref]$pidOut);
$ourThread = [Native]::GetCurrentThreadId();
[Native]::AttachThreadInput($ourThread, $fgThread, $true) | Out-Null;
[Native]::SetForegroundWindow($h) | Out-Null;
[Native]::AttachThreadInput($ourThread, $fgThread, $false) | Out-Null;
$owner.Activate();
`.trim();

/**
 * Runs a PowerShell script (prefixed with the owner-form/foreground preamble) in a
 * single-threaded apartment and returns its UTF-8 stdout. Used by the native
 * Windows MsgBox/InputBox backends.
 *
 * The script is passed via `-EncodedCommand` (base64 UTF-16LE) rather than
 * `-Command` so that quotes, newlines and the embedded C# here-string survive
 * byte-for-byte — no command-line escaping at all.
 */
export function runPowerShellDialog(body: string): string {
  const exec = getExecFileSync()!;
  const script = PS_PREAMBLE + "\n" + body;
  const encoded = Buffer.from(script, "utf16le").toString("base64");
  const out = exec(
    "powershell",
    ["-NoProfile", "-STA", "-EncodedCommand", encoded],
    {
      encoding: "utf8",
    },
  );
  return String(out);
}
