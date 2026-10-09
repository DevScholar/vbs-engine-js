import type { VbValue } from '../runtime/index.ts';
import { getExecFileSync, psQuote, runPowerShellDialog } from './gui.ts';

export interface InputBoxOptions {
  prompt?: (message: string, defaultValue: string) => string | null;
  console?: (message: string) => void;
  /** Synchronous line read; returns null on EOF. The host supplies this. */
  readline?: () => string | null;
  /**
   * Native GUI dialog hook. Returns the entered text, or null if the user cancelled
   * (which VBScript reports as an empty string). Takes precedence over the
   * console/readline fallback, matching real WSH where InputBox is always a
   * blocking modal dialog (in both cscript and wscript).
   */
  gui?: (args: { message: string; title: string; defaultText: string }) => string | null;
}

export function createInputBox(options?: InputBoxOptions) {
  return function (
    promptVal: VbValue,
    title?: VbValue,
    defaultText?: VbValue,
    _xPos?: VbValue,
    _yPos?: VbValue,
    _helpFile?: VbValue,
    _contextVal?: VbValue
  ): VbValue {
    void _xPos;
    void _yPos;
    void _helpFile;
    void _contextVal;
    const message = String(promptVal.value ?? promptVal);
    const titleStr = title ? String(title.value ?? title) : 'Input';
    const defaultVal = defaultText ? String(defaultText.value ?? defaultText) : '';
    const hasDefault = defaultVal !== '';

    if (options?.gui) {
      const result = options.gui({ message, title: titleStr, defaultText: defaultVal });
      return { type: 'String', value: result ?? '' };
    }

    if (options?.prompt) {
      const result = options.prompt(message, defaultVal);
      return { type: 'String', value: result ?? '' };
    }

    const promptText =
      `[${titleStr}]\n${message}\n` +
      (hasDefault ? `Default: ${defaultVal}\n` : '') +
      `[Enter] Confirm / [Esc] Cancel (default is "${defaultVal || '(empty)'}")`;

    if (options?.console) {
      options.console(promptText);
    } else {
      console.log(promptText);
    }

    if (options?.readline) {
      const readline = options.readline;
      while (true) {
        const input = readline();
        if (input === null) {
          return { type: 'String', value: '' };
        }
        if (input === '' && hasDefault) {
          return { type: 'String', value: defaultVal };
        }
        if (input === '' && !hasDefault) {
          if (options?.console) {
            options.console(`Input cannot be empty. Please try again.\n` + promptText);
          } else {
            console.log(`Input cannot be empty. Please try again.\n` + promptText);
          }
          continue;
        }
        return { type: 'String', value: input };
      }
    }

    return { type: 'String', value: '' };
  };
}

export function createBrowserInputBox() {
  return function (
    prompt: VbValue,
    title?: VbValue,
    defaultVal?: VbValue,
    _xPos?: VbValue,
    _yPos?: VbValue,
    _helpFile?: VbValue,
    _contextVal?: VbValue
  ): VbValue {
    void _xPos;
    void _yPos;
    void _helpFile;
    void _contextVal;
    const message = String(prompt.value ?? prompt);
    const titleStr = title ? String(title.value ?? title) : 'Input';
    const def = defaultVal ? String(defaultVal.value ?? defaultVal) : '';
    while (true) {
      const result = window.prompt(`[${titleStr}]\n${message}`, def);
      if (result === null) {
        return { type: 'String', value: '' };
      }
      if (result === '' && def !== '') {
        return { type: 'String', value: def };
      }
      if (result === '' && def === '') {
        alert('Input cannot be empty. Please try again.');
        continue;
      }
      return { type: 'String', value: result };
    }
  };
}

export function registerInputBox(context: {
  functionRegistry: { register: (name: string, func: (...args: VbValue[]) => VbValue) => void };
}): void {
  context.functionRegistry.register(
    'InputBox',
    (
      promptVal: VbValue,
      title?: VbValue,
      defaultText?: VbValue,
      _xPos?: VbValue,
      _yPos?: VbValue,
      _helpFile?: VbValue,
      _contextVal?: VbValue
    ): VbValue => {
      void _xPos;
      void _yPos;
      void _helpFile;
      void _contextVal;
      const message = String(promptVal.value ?? promptVal);
      const titleStr = title ? String(title.value ?? title) : 'Input';
      const defaultVal = defaultText ? String(defaultText.value ?? defaultText) : '';

      if (typeof window !== 'undefined' && typeof window.prompt === 'function') {
        while (true) {
          const result = window.prompt(`[${titleStr}]\n${message}`, defaultVal);
          if (result === null) {
            return { type: 'String', value: '' };
          }
          if (result === '' && defaultVal !== '') {
            return { type: 'String', value: defaultVal };
          }
          if (result === '' && defaultVal === '') {
            alert('Input cannot be empty. Please try again.');
            continue;
          }
          return { type: 'String', value: result };
        }
      }

      // Browser-less (Node) host: prefer a native GUI dialog (matching real WSH,
      // where InputBox is always a blocking modal dialog, even under cscript). If
      // no platform backend is available, fall back to a non-interactive console
      // line. A host that wants custom interactive InputBox overrides this via
      // createInputBox({ readline, console, gui, ... }).
      if (typeof process !== 'undefined' && getExecFileSync() !== null) {
        return { type: 'String', value: showNativeInputBox(message, titleStr, defaultVal) };
      }
      console.log(`[${titleStr}]\n${message}`);
      return { type: 'String', value: '' };
    }
  );
}

// --- Native GUI dialog backend (Node only) -----------------------------------
// Real WSH shows InputBox as a blocking modal dialog under BOTH cscript and
// wscript. In a Node host we reach for a platform dialog (PowerShell InputBox on
// Windows, osascript on macOS, zenity on Unix). The backend is fetched lazily via
// `process.getBuiltinModule` so the browser build never pulls in node:child_process.

/** Shows a native InputBox dialog; returns the entered text, or '' on cancel. */
function showNativeInputBox(message: string, title: string, defaultText: string): string {
  const execFileSync = getExecFileSync()!;

  if (process.platform === 'win32') {
    const body =
      'Add-Type -AssemblyName Microsoft.VisualBasic; ' +
      `$r=[Microsoft.VisualBasic.Interaction]::InputBox(${psQuote(message)}, ${psQuote(title)}, ${psQuote(defaultText)}); ` +
      'if ([string]::IsNullOrEmpty($r)) { Write-Output "__CANCEL__" } else { Write-Output $r }';
    const out = runPowerShellDialog(body).trim();
    return out === '__CANCEL__' ? '' : out;
  }

  if (process.platform === 'darwin') {
    const script =
      `display dialog ${JSON.stringify(message)} default answer ${JSON.stringify(defaultText)}` +
      (title ? ` with title ${JSON.stringify(title)}` : '');
    const out = String(execFileSync('osascript', ['-e', script], { encoding: 'utf8' }));
    const m = /text returned:(.*?)(?:,|$)/.exec(out);
    return m ? m[1]!.trim() : '';
  }

  // Unix-like, non-macOS.
  try {
    const out = String(
      execFileSync(
        'zenity',
        ['--entry', '--title', title, '--text', message, '--entry-text', defaultText],
        { encoding: 'utf8' }
      )
    );
    return out.replace(/\r?\n$/, '');
  } catch {
    return '';
  }
}
