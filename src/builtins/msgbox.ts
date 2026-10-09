import type { VbValue } from '../runtime/index.ts';
import { getExecFileSync, psQuote, runPowerShellDialog } from './gui.ts';

export const MsgBoxConstants = {
  vbOKOnly: 0,
  vbOKCancel: 1,
  vbAbortRetryIgnore: 2,
  vbYesNoCancel: 3,
  vbYesNo: 4,
  vbRetryCancel: 5,

  vbCritical: 16,
  vbQuestion: 32,
  vbExclamation: 48,
  vbInformation: 64,

  vbDefaultButton1: 0,
  vbDefaultButton2: 256,
  vbDefaultButton3: 512,

  vbOK: 1,
  vbCancel: 2,
  vbAbort: 3,
  vbRetry: 4,
  vbIgnore: 5,
  vbYes: 6,
  vbNo: 7,
};

export type MsgBoxButtonType =
  'OK' | 'OKCancel' | 'AbortRetryIgnore' | 'YesNoCancel' | 'YesNo' | 'RetryCancel';
export type MsgBoxIconType = 'Critical' | 'Question' | 'Exclamation' | 'Information' | null;

export interface MsgBoxResult {
  value: number;
  button: string;
}

function parseButtons(buttons: number | undefined): {
  buttonType: MsgBoxButtonType;
  iconType: MsgBoxIconType;
  defaultButton: number;
} {
  const b = buttons ?? 0;
  const buttonTypeValue = b & 0x7;
  const iconValue = b & 0xf0;
  const defaultButton = b & 0x300;

  let buttonType: MsgBoxButtonType;
  switch (buttonTypeValue) {
    case 0:
      buttonType = 'OK';
      break;
    case 1:
      buttonType = 'OKCancel';
      break;
    case 2:
      buttonType = 'AbortRetryIgnore';
      break;
    case 3:
      buttonType = 'YesNoCancel';
      break;
    case 4:
      buttonType = 'YesNo';
      break;
    case 5:
      buttonType = 'RetryCancel';
      break;
    default:
      buttonType = 'OK';
  }

  let iconType: MsgBoxIconType = null;
  switch (iconValue) {
    case 16:
      iconType = 'Critical';
      break;
    case 32:
      iconType = 'Question';
      break;
    case 48:
      iconType = 'Exclamation';
      break;
    case 64:
      iconType = 'Information';
      break;
  }

  return { buttonType, iconType, defaultButton };
}

function getIconPrefix(iconType: MsgBoxIconType): string {
  switch (iconType) {
    case 'Critical':
      return '[X] ';
    case 'Question':
      return '[?] ';
    case 'Exclamation':
      return '[!] ';
    case 'Information':
      return '[i] ';
    default:
      return '';
  }
}

function buildPrompt(title: string | null, iconType: MsgBoxIconType, message: string): string {
  let prompt = getIconPrefix(iconType) + message;
  if (title) {
    prompt = `[${title}]\n` + prompt;
  }
  return prompt;
}

function mapButtonToResult(button: string, buttonType: MsgBoxButtonType): number {
  const b = button.toLowerCase();
  switch (buttonType) {
    case 'OK':
    case 'OKCancel':
      return b === 'ok' || b === 'confirm' ? MsgBoxConstants.vbOK : MsgBoxConstants.vbCancel;
    case 'YesNo':
    case 'YesNoCancel':
      if (b === 'yes' || b === 'y') return MsgBoxConstants.vbYes;
      if (b === 'no' || b === 'n') return MsgBoxConstants.vbNo;
      return MsgBoxConstants.vbCancel;
    case 'RetryCancel':
      if (b === 'retry' || b === 'r') return MsgBoxConstants.vbRetry;
      return MsgBoxConstants.vbCancel;
    case 'AbortRetryIgnore':
      if (b === 'abort' || b === 'a') return MsgBoxConstants.vbAbort;
      if (b === 'retry' || b === 'r') return MsgBoxConstants.vbRetry;
      if (b === 'ignore' || b === 'i') return MsgBoxConstants.vbIgnore;
      return MsgBoxConstants.vbCancel;
    default:
      return MsgBoxConstants.vbCancel;
  }
}

/** Ordered button labels for each MsgBox button type (left-to-right in the dialog). */
export function getButtonLabels(buttonType: MsgBoxButtonType): string[] {
  switch (buttonType) {
    case 'OK':
      return ['OK'];
    case 'OKCancel':
      return ['OK', 'Cancel'];
    case 'AbortRetryIgnore':
      return ['Abort', 'Retry', 'Ignore'];
    case 'YesNoCancel':
      return ['Yes', 'No', 'Cancel'];
    case 'YesNo':
      return ['Yes', 'No'];
    case 'RetryCancel':
      return ['Retry', 'Cancel'];
  }
}

/** Maps a button label (e.g. "Yes") to its MsgBox return value (vbYes). */
export const BUTTON_LABEL_TO_RESULT: Record<string, number> = {
  OK: MsgBoxConstants.vbOK,
  Cancel: MsgBoxConstants.vbCancel,
  Abort: MsgBoxConstants.vbAbort,
  Retry: MsgBoxConstants.vbRetry,
  Ignore: MsgBoxConstants.vbIgnore,
  Yes: MsgBoxConstants.vbYes,
  No: MsgBoxConstants.vbNo,
};

/** 0 → first button, 256 → second, 512 → third, clamped to the button count. */
export function defaultButtonIndex(defaultButton: number, count: number): number {
  const idx = Math.floor(defaultButton / 256);
  return Math.min(Math.max(0, idx), count - 1);
}

function getButtonOptions(buttonType: MsgBoxButtonType, defaultButton: number): string {
  const labels = getButtonLabels(buttonType);
  const index = defaultButtonIndex(defaultButton, labels.length);
  const defaultLabel = labels[index]!.charAt(0);
  const options = labels.map(l => `[${l.charAt(0)}] ${l}`).join(' / ');
  return options + ` (default is "${defaultLabel}"): `;
}

export interface MsgBoxOptions {
  alert?: (message: string) => void;
  confirm?: (message: string) => boolean | undefined;
  prompt?: (message: string, defaultValue: string) => string | null;
  console?: (message: string) => void;
  /** Synchronous line read; returns null on EOF. The host supplies this. */
  readline?: () => string | null;
  /**
   * Native GUI dialog hook. Receives the parsed button/icon config and must return
   * the chosen button's VBScript result code (MsgBoxConstants.vb*). When present it
   * takes precedence over the console/readline fallback, matching real WSH where
   * MsgBox is always a blocking modal dialog (in both cscript and wscript).
   */
  gui?: (args: {
    message: string;
    title: string | null;
    buttonType: MsgBoxButtonType;
    iconType: MsgBoxIconType;
    defaultButton: number;
  }) => number;
}

export function createMsgBox(options?: MsgBoxOptions) {
  return function (promptVal: VbValue, buttons?: VbValue, title?: VbValue): VbValue {
    const message = String(promptVal.value ?? promptVal);
    const buttonsVal = buttons ? Number(buttons.value ?? buttons) : 0;
    const titleStr = title ? String(title.value ?? title) : null;

    const { buttonType, iconType, defaultButton } = parseButtons(buttonsVal);
    const fullPrompt = buildPrompt(titleStr, iconType, message);
    const buttonOptions = getButtonOptions(buttonType, defaultButton);

    const finalPrompt = fullPrompt + '\n' + buttonOptions;

    if (options?.gui) {
      const result = options.gui({ message, title: titleStr, buttonType, iconType, defaultButton });
      return { type: 'Integer', value: result };
    }

    if (options?.alert && buttonType === 'OK') {
      options.alert(finalPrompt);
      return { type: 'Integer', value: MsgBoxConstants.vbOK };
    }

    if (options?.confirm) {
      const isYesNoType =
        buttonType === 'YesNo' || buttonType === 'YesNoCancel' || buttonType === 'OKCancel';
      if (isYesNoType) {
        const confirmResult = options.confirm(finalPrompt);
        if (confirmResult === undefined) {
          // User cancelled - continue to readline
        } else if (confirmResult) {
          return {
            type: 'Integer',
            value: buttonType === 'OKCancel' ? MsgBoxConstants.vbOK : MsgBoxConstants.vbYes,
          };
        } else {
          return {
            type: 'Integer',
            value: buttonType === 'OKCancel' ? MsgBoxConstants.vbCancel : MsgBoxConstants.vbNo,
          };
        }
      }
    }

    if (options?.console) {
      options.console(finalPrompt);
    } else {
      console.log(finalPrompt);
    }

    if (options?.readline) {
      const readline = options.readline;
      while (true) {
        const input = readline();
        if (input === null) {
          return { type: 'Integer', value: MsgBoxConstants.vbCancel };
        }
        const trimmed = input.trim().toLowerCase();
        if (trimmed === '') {
          return { type: 'Integer', value: getDefaultResult(buttonType, defaultButton) };
        }
        const result = mapButtonToResult(input, buttonType);
        if (result !== MsgBoxConstants.vbCancel || isValidExplicitButton(trimmed, buttonType)) {
          return { type: 'Integer', value: result };
        }
        if (options?.console) {
          options.console(`Invalid input: ${input}. Please try again.\n` + finalPrompt);
        } else {
          console.log(`Invalid input: ${input}. Please try again.\n` + finalPrompt);
        }
      }
    }

    return { type: 'Integer', value: MsgBoxConstants.vbCancel };
  };
}

export function getDefaultResult(buttonType: MsgBoxButtonType, defaultButton: number): number {
  const labels = getButtonLabels(buttonType);
  const index = defaultButtonIndex(defaultButton, labels.length);
  return BUTTON_LABEL_TO_RESULT[labels[index]!] ?? MsgBoxConstants.vbOK;
}

function isValidExplicitButton(input: string, buttonType: MsgBoxButtonType): boolean {
  switch (buttonType) {
    case 'OK':
      return input === 'o' || input === 'ok';
    case 'OKCancel':
      return input === 'o' || input === 'ok' || input === 'c' || input === 'cancel';
    case 'YesNo':
      return input === 'y' || input === 'yes' || input === 'n' || input === 'no';
    case 'YesNoCancel':
      return (
        input === 'y' ||
        input === 'yes' ||
        input === 'n' ||
        input === 'no' ||
        input === 'c' ||
        input === 'cancel'
      );
    case 'RetryCancel':
      return input === 'r' || input === 'retry' || input === 'c' || input === 'cancel';
    case 'AbortRetryIgnore':
      return (
        input === 'a' ||
        input === 'abort' ||
        input === 'r' ||
        input === 'retry' ||
        input === 'i' ||
        input === 'ignore'
      );
    default:
      return false;
  }
}

export function createBrowserMsgBox() {
  return function (prompt: VbValue, buttons?: VbValue, title?: VbValue): VbValue {
    const message = String(prompt.value ?? prompt);
    const buttonsVal = buttons ? Number(buttons.value ?? buttons) : 0;
    const titleStr = title ? String(title.value ?? title) : 'VBScript';

    const { buttonType, iconType, defaultButton } = parseButtons(buttonsVal);
    const fullPrompt = buildPrompt(titleStr, iconType, message);
    const buttonOptions = getButtonOptions(buttonType, defaultButton);
    void buttonOptions;

    if (buttonType === 'OK') {
      alert(fullPrompt);
      return { type: 'Integer', value: MsgBoxConstants.vbOK };
    }

    if (buttonType === 'OKCancel') {
      const confirmed = confirm(fullPrompt + '\n[O] OK / [C] Cancel (default is "O"): ');
      return {
        type: 'Integer',
        value: confirmed ? MsgBoxConstants.vbOK : MsgBoxConstants.vbCancel,
      };
    }

    if (buttonType === 'YesNo') {
      const defaultLabel = defaultButton === 256 ? 'Y' : 'N';
      const confirmed = confirm(
        fullPrompt + '\n[Y] Yes / [N] No (default is "' + defaultLabel + '"): '
      );
      return { type: 'Integer', value: confirmed ? MsgBoxConstants.vbYes : MsgBoxConstants.vbNo };
    }

    if (buttonType === 'YesNoCancel') {
      while (true) {
        let defaultLabel = 'N';
        if (defaultButton === 512) defaultLabel = 'C';
        else if (defaultButton === 256) defaultLabel = 'Y';

        const input = window.prompt(
          fullPrompt + '\n[Y] Yes / [N] No / [C] Cancel (default is "' + defaultLabel + '"): ',
          ''
        );
        if (input === null) {
          return { type: 'Integer', value: MsgBoxConstants.vbCancel };
        }
        const normalized = input.trim().toLowerCase();
        if (normalized === '') {
          if (defaultButton === 512) return { type: 'Integer', value: MsgBoxConstants.vbCancel };
          if (defaultButton === 256) return { type: 'Integer', value: MsgBoxConstants.vbYes };
          return { type: 'Integer', value: MsgBoxConstants.vbNo };
        }
        if (normalized === 'y' || normalized === 'yes') {
          return { type: 'Integer', value: MsgBoxConstants.vbYes };
        }
        if (normalized === 'n' || normalized === 'no') {
          return { type: 'Integer', value: MsgBoxConstants.vbNo };
        }
        if (normalized === 'c' || normalized === 'cancel') {
          return { type: 'Integer', value: MsgBoxConstants.vbCancel };
        }
        alert('Invalid input: ' + input + '. Please try again.');
      }
    }

    if (buttonType === 'RetryCancel') {
      const defaultLabel = defaultButton === 256 ? 'R' : 'C';
      const confirmed = confirm(
        fullPrompt + '\n[R] Retry / [C] Cancel (default is "' + defaultLabel + '"): '
      );
      return {
        type: 'Integer',
        value: confirmed ? MsgBoxConstants.vbRetry : MsgBoxConstants.vbCancel,
      };
    }

    if (buttonType === 'AbortRetryIgnore') {
      while (true) {
        let defaultLabel = 'A';
        if (defaultButton === 512) defaultLabel = 'I';
        else if (defaultButton === 256) defaultLabel = 'R';

        const input = window.prompt(
          fullPrompt + '\n[A] Abort / [R] Retry / [I] Ignore (default is "' + defaultLabel + '"): ',
          ''
        );
        if (input === null) {
          return { type: 'Integer', value: MsgBoxConstants.vbCancel };
        }
        const normalized = input.trim().toLowerCase();
        if (normalized === '') {
          if (defaultButton === 512) return { type: 'Integer', value: MsgBoxConstants.vbIgnore };
          if (defaultButton === 256) return { type: 'Integer', value: MsgBoxConstants.vbRetry };
          return { type: 'Integer', value: MsgBoxConstants.vbAbort };
        }
        if (normalized === 'a' || normalized === 'abort') {
          return { type: 'Integer', value: MsgBoxConstants.vbAbort };
        }
        if (normalized === 'r' || normalized === 'retry') {
          return { type: 'Integer', value: MsgBoxConstants.vbRetry };
        }
        if (normalized === 'i' || normalized === 'ignore') {
          return { type: 'Integer', value: MsgBoxConstants.vbIgnore };
        }
        alert('Invalid input: ' + input + '. Please try again.');
      }
    }

    alert(fullPrompt);
    return { type: 'Integer', value: MsgBoxConstants.vbOK };
  };
}

export function registerMsgBox(context: {
  functionRegistry: {
    register: (
      name: string,
      func: (...args: VbValue[]) => VbValue,
      options?: { isSub: boolean }
    ) => void;
  };
}): void {
  context.functionRegistry.register(
    'MsgBox',
    (promptVal: VbValue, buttons?: VbValue, title?: VbValue): VbValue => {
      const message = String(promptVal.value ?? promptVal);
      const buttonsVal = buttons ? Number(buttons.value ?? buttons) : 0;
      const titleStr = title ? String(title.value ?? title) : null;

      const { buttonType, iconType, defaultButton } = parseButtons(buttonsVal);
      const fullPrompt = buildPrompt(titleStr, iconType, message);
      const buttonOptions = getButtonOptions(buttonType, defaultButton);
      const finalPrompt = fullPrompt + '\n' + buttonOptions;

      if (buttonType === 'OK') {
        if (typeof alert !== 'undefined') {
          alert(fullPrompt);
        } else if (typeof process !== 'undefined' && getExecFileSync() !== null) {
          return {
            type: 'Integer',
            value: showNativeMsgBox(
              message,
              titleStr ?? 'Windows Script Host',
              buttonType,
              iconType,
              defaultButton
            ),
          };
        } else {
          console.log(fullPrompt);
        }
        return { type: 'Integer', value: MsgBoxConstants.vbOK };
      }

      if (typeof window !== 'undefined' && typeof window.confirm === 'function') {
        if (buttonType === 'OKCancel') {
          const confirmed = confirm(fullPrompt + '\n[O] OK / [C] Cancel (default is "O"): ');
          return {
            type: 'Integer',
            value: confirmed ? MsgBoxConstants.vbOK : MsgBoxConstants.vbCancel,
          };
        }

        if (buttonType === 'YesNo') {
          const defaultLabel = defaultButton === 256 ? 'Y' : 'N';
          const confirmed = confirm(
            fullPrompt + '\n[Y] Yes / [N] No (default is "' + defaultLabel + '"): '
          );
          return {
            type: 'Integer',
            value: confirmed ? MsgBoxConstants.vbYes : MsgBoxConstants.vbNo,
          };
        }

        if (buttonType === 'YesNoCancel' && typeof window.prompt === 'function') {
          while (true) {
            let defaultLabel = 'N';
            if (defaultButton === 512) defaultLabel = 'C';
            else if (defaultButton === 256) defaultLabel = 'Y';

            const input = window.prompt(
              fullPrompt + '\n[Y] Yes / [N] No / [C] Cancel (default is "' + defaultLabel + '"): ',
              ''
            );
            if (input === null) {
              return { type: 'Integer', value: MsgBoxConstants.vbCancel };
            }
            const normalized = input.trim().toLowerCase();
            if (normalized === '') {
              if (defaultButton === 512)
                return { type: 'Integer', value: MsgBoxConstants.vbCancel };
              if (defaultButton === 256) return { type: 'Integer', value: MsgBoxConstants.vbYes };
              return { type: 'Integer', value: MsgBoxConstants.vbNo };
            }
            if (normalized === 'y' || normalized === 'yes') {
              return { type: 'Integer', value: MsgBoxConstants.vbYes };
            }
            if (normalized === 'n' || normalized === 'no') {
              return { type: 'Integer', value: MsgBoxConstants.vbNo };
            }
            if (normalized === 'c' || normalized === 'cancel') {
              return { type: 'Integer', value: MsgBoxConstants.vbCancel };
            }
            alert('Invalid input: ' + input + '. Please try again.');
          }
        }

        if (buttonType === 'RetryCancel') {
          const defaultLabel = defaultButton === 256 ? 'R' : 'C';
          const confirmed = confirm(
            fullPrompt + '\n[R] Retry / [C] Cancel (default is "' + defaultLabel + '"): '
          );
          return {
            type: 'Integer',
            value: confirmed ? MsgBoxConstants.vbRetry : MsgBoxConstants.vbCancel,
          };
        }

        if (buttonType === 'AbortRetryIgnore' && typeof window.prompt === 'function') {
          while (true) {
            let defaultLabel = 'A';
            if (defaultButton === 512) defaultLabel = 'I';
            else if (defaultButton === 256) defaultLabel = 'R';

            const input = window.prompt(
              fullPrompt +
                '\n[A] Abort / [R] Retry / [I] Ignore (default is "' +
                defaultLabel +
                '"): ',
              ''
            );
            if (input === null) {
              return { type: 'Integer', value: MsgBoxConstants.vbCancel };
            }
            const normalized = input.trim().toLowerCase();
            if (normalized === '') {
              if (defaultButton === 512)
                return { type: 'Integer', value: MsgBoxConstants.vbIgnore };
              if (defaultButton === 256) return { type: 'Integer', value: MsgBoxConstants.vbRetry };
              return { type: 'Integer', value: MsgBoxConstants.vbAbort };
            }
            if (normalized === 'a' || normalized === 'abort') {
              return { type: 'Integer', value: MsgBoxConstants.vbAbort };
            }
            if (normalized === 'r' || normalized === 'retry') {
              return { type: 'Integer', value: MsgBoxConstants.vbRetry };
            }
            if (normalized === 'i' || normalized === 'ignore') {
              return { type: 'Integer', value: MsgBoxConstants.vbIgnore };
            }
            alert('Invalid input: ' + input + '. Please try again.');
          }
        }
      }

      // Browser-less (Node) host: prefer a native GUI dialog (matching real WSH,
      // where MsgBox is always a blocking modal dialog, even under cscript). If no
      // platform backend is available, fall back to a non-interactive console line.
      if (typeof process !== 'undefined' && getExecFileSync() !== null) {
        return {
          type: 'Integer',
          value: showNativeMsgBox(
            message,
            titleStr ?? 'Windows Script Host',
            buttonType,
            iconType,
            defaultButton
          ),
        };
      }
      console.log(finalPrompt);
      return { type: 'Integer', value: MsgBoxConstants.vbCancel };
    },
    { isSub: false }
  );
}

// --- Native GUI dialog backend (Node only) -----------------------------------
// Real WSH shows MsgBox as a blocking modal dialog under BOTH cscript and wscript;
// it never degrades to a console prompt. In a Node host (which has no `alert` /
// `window.confirm`), we reach for a platform dialog: PowerShell MessageBox on
// Windows, osascript on macOS, zenity on Unix. The backend is fetched lazily via
// `process.getBuiltinModule` so the browser build never pulls in node:child_process.

/** Windows MessageBoxButtons / MessageBoxIcon / MessageBoxDefaultButton names. */
const MSGBOX_BUTTONS: Record<MsgBoxButtonType, string> = {
  OK: 'OK',
  OKCancel: 'OKCancel',
  AbortRetryIgnore: 'AbortRetryIgnore',
  YesNoCancel: 'YesNoCancel',
  YesNo: 'YesNo',
  RetryCancel: 'RetryCancel',
};

const MSGBOX_ICONS: Record<Exclude<MsgBoxIconType, null>, string> = {
  Critical: 'Error',
  Question: 'Question',
  Exclamation: 'Warning',
  Information: 'Information',
};

/** Maps a MessageBox.Show DialogResult name back to a VBScript result code. */
const DIALOG_RESULT_TO_VB: Record<string, number> = {
  OK: MsgBoxConstants.vbOK,
  Cancel: MsgBoxConstants.vbCancel,
  Abort: MsgBoxConstants.vbAbort,
  Retry: MsgBoxConstants.vbRetry,
  Ignore: MsgBoxConstants.vbIgnore,
  Yes: MsgBoxConstants.vbYes,
  No: MsgBoxConstants.vbNo,
};

/** Shows a native MsgBox dialog; returns the chosen VBScript result code. */
function showNativeMsgBox(
  message: string,
  title: string,
  buttonType: MsgBoxButtonType,
  iconType: MsgBoxIconType,
  defaultButton: number
): number {
  const execFileSync = getExecFileSync()!;
  const labels = getButtonLabels(buttonType);

  if (process.platform === 'win32') {
    const icon = iconType
      ? `[System.Windows.Forms.MessageBoxIcon]::${MSGBOX_ICONS[iconType]}`
      : '[System.Windows.Forms.MessageBoxIcon]::None';
    const defIndex = defaultButtonIndex(defaultButton, labels.length);
    const def = `[System.Windows.Forms.MessageBoxDefaultButton]::Button${defIndex + 1}`;
    const body =
      `$r=[System.Windows.Forms.MessageBox]::Show($owner, ${psQuote(message)}, ${psQuote(title)}, ` +
      `[System.Windows.Forms.MessageBoxButtons]::${MSGBOX_BUTTONS[buttonType]}, ${icon}, ${def}); ` +
      'Write-Output $r.ToString()';
    const out = runPowerShellDialog(body);
    return DIALOG_RESULT_TO_VB[out.trim()] ?? MsgBoxConstants.vbCancel;
  }

  if (process.platform === 'darwin') {
    const defLabel = labels[defaultButtonIndex(defaultButton, labels.length)]!;
    const buttons = labels.map(l => `"${l}"`).join(', ');
    const script =
      `display dialog ${JSON.stringify(message)} buttons {${buttons}} default button "${defLabel}"` +
      (title ? ` with title ${JSON.stringify(title)}` : '');
    const out = execFileSync('osascript', ['-e', script], { encoding: 'utf8' });
    const m = /button returned:(.+)/.exec(String(out));
    return BUTTON_LABEL_TO_RESULT[(m?.[1] ?? defLabel).trim()] ?? MsgBoxConstants.vbCancel;
  }

  // Unix-like, non-macOS.
  if (labels.length === 1) {
    execFileSync('zenity', ['--info', '--title', title, '--text', message], { stdio: 'inherit' });
    return BUTTON_LABEL_TO_RESULT[labels[0]!] ?? MsgBoxConstants.vbCancel;
  }
  try {
    const out = execFileSync(
      'zenity',
      ['--list', '--title', title, '--text', message, '--column=', ...labels],
      { encoding: 'utf8' }
    );
    return BUTTON_LABEL_TO_RESULT[String(out).trim()] ?? MsgBoxConstants.vbCancel;
  } catch {
    return getDefaultResult(buttonType, defaultButton);
  }
}
