import type { VbValue } from '../runtime/index.ts';

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

export type MsgBoxButtonType = 'OK' | 'OKCancel' | 'AbortRetryIgnore' | 'YesNoCancel' | 'YesNo' | 'RetryCancel';
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
  const iconValue = b & 0xF0;
  const defaultButton = b & 0x300;

  let buttonType: MsgBoxButtonType;
  switch (buttonTypeValue) {
    case 0: buttonType = 'OK'; break;
    case 1: buttonType = 'OKCancel'; break;
    case 2: buttonType = 'AbortRetryIgnore'; break;
    case 3: buttonType = 'YesNoCancel'; break;
    case 4: buttonType = 'YesNo'; break;
    case 5: buttonType = 'RetryCancel'; break;
    default: buttonType = 'OK';
  }

  let iconType: MsgBoxIconType = null;
  switch (iconValue) {
    case 16: iconType = 'Critical'; break;
    case 32: iconType = 'Question'; break;
    case 48: iconType = 'Exclamation'; break;
    case 64: iconType = 'Information'; break;
  }

  return { buttonType, iconType, defaultButton };
}

function getIconPrefix(iconType: MsgBoxIconType): string {
  switch (iconType) {
    case 'Critical': return '[X] ';
    case 'Question': return '[?] ';
    case 'Exclamation': return '[!] ';
    case 'Information': return '[i] ';
    default: return '';
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

function getButtonOptions(buttonType: MsgBoxButtonType, defaultButton: number): string[] {
  const options: string[] = [];
  switch (buttonType) {
    case 'OK':
      options.push('OK');
      break;
    case 'OKCancel':
      options.push(defaultButton === 256 ? 'OK (default), Cancel' : 'OK, Cancel (default)');
      break;
    case 'YesNo':
      options.push(defaultButton === 256 ? 'Yes (default), No' : 'Yes, No (default)');
      break;
    case 'YesNoCancel':
      if (defaultButton === 512) {
        options.push('Yes, No, Cancel (default)');
      } else if (defaultButton === 256) {
        options.push('Yes (default), No, Cancel');
      } else {
        options.push('Yes, No (default), Cancel');
      }
      break;
    case 'RetryCancel':
      options.push(defaultButton === 256 ? 'Retry (default), Cancel' : 'Retry, Cancel (default)');
      break;
    case 'AbortRetryIgnore':
      if (defaultButton === 512) {
        options.push('Abort, Retry, Ignore (default)');
      } else if (defaultButton === 256) {
        options.push('Abort, Retry (default), Ignore');
      } else {
        options.push('Abort (default), Retry, Ignore');
      }
      break;
  }
  return options;
}

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function _getInputPrompt(buttonType: MsgBoxButtonType): string {
  let prompt = '';
  switch (buttonType) {
    case 'OK':
      prompt = '[OK]';
      break;
    case 'OKCancel':
      prompt = '[O]K / [C]ancel';
      break;
    case 'YesNo':
      prompt = '[Y]es / [N]o';
      break;
    case 'YesNoCancel':
      prompt = '[Y]es / [N]o / [C]ancel';
      break;
    case 'RetryCancel':
      prompt = '[R]etry / [C]ancel';
      break;
    case 'AbortRetryIgnore':
      prompt = '[A]bort / [R]etry / [I]gnore';
      break;
  }
  return prompt;
}

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function _readFromConsole(): Promise<string | null> {
  return new Promise((resolve) => {
    // eslint-disable-next-line @typescript-eslint/ban-ts-comment
    // @ts-ignore - Node.js specific
    // eslint-disable-next-line @typescript-eslint/no-require-imports
    const readline = require('readline');
    const rl = readline.createInterface({
      // eslint-disable-next-line @typescript-eslint/ban-ts-comment
      // @ts-ignore - Node.js specific
      input: process.stdin,
      // eslint-disable-next-line @typescript-eslint/ban-ts-comment
      // @ts-ignore - Node.js specific
      output: process.stdout
    });
    rl.question('', (answer: string) => {
      rl.close();
      resolve(answer || null);
    });
  });
}

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function _syncReadFromConsole(): string | null {
  try {
    // eslint-disable-next-line @typescript-eslint/ban-ts-comment
    // @ts-ignore - Node.js specific
    // eslint-disable-next-line @typescript-eslint/no-require-imports
    const readline = require('readline');
    const rl = readline.createInterface({
      // eslint-disable-next-line @typescript-eslint/ban-ts-comment
      // @ts-ignore - Node.js specific
      input: process.stdin,
      // eslint-disable-next-line @typescript-eslint/ban-ts-comment
      // @ts-ignore - Node.js specific
      output: process.stdout
    });
    rl.question('', () => {});
    // eslint-disable-next-line @typescript-eslint/no-unused-vars
  } catch (e) {
    return null;
  }
  return null;
}

export interface MsgBoxOptions {
  alert?: (message: string) => void;
  confirm?: (message: string) => boolean | undefined;
  prompt?: (message: string, defaultValue: string) => string | null;
  console?: (message: string) => void;
  readline?: () => Promise<string | null>;
}

export function createMsgBox(options?: MsgBoxOptions) {
  return function(
    _context: { functionRegistry: { register: (name: string, func: (...args: VbValue[]) => VbValue, options?: { isSub: boolean }) => void } },
    promptVal: VbValue,
    buttons?: VbValue,
    title?: VbValue
  ): VbValue {
    const message = String(promptVal.value ?? promptVal);
    const buttonsVal = buttons ? Number(buttons.value ?? buttons) : 0;
    const titleStr = title ? String(title.value ?? title) : null;

    const { buttonType, iconType, defaultButton } = parseButtons(buttonsVal);
    const fullPrompt = buildPrompt(titleStr, iconType, message);
    const buttonOptions = getButtonOptions(buttonType, defaultButton);

    const finalPrompt = fullPrompt + '\n' + buttonOptions.join(' ');

    if (options?.alert && buttonType === 'OK') {
      options.alert(finalPrompt);
      return { type: 'Integer', value: MsgBoxConstants.vbOK };
    }

    if (options?.confirm) {
      const isYesNoType = buttonType === 'YesNo' || buttonType === 'YesNoCancel' || buttonType === 'OKCancel';
      if (isYesNoType) {
        const confirmResult = options.confirm(finalPrompt);
        if (confirmResult === undefined) {
          // User cancelled - continue to readline
        } else if (confirmResult) {
          return { type: 'Integer', value: buttonType === 'OKCancel' ? MsgBoxConstants.vbOK : MsgBoxConstants.vbYes };
        } else {
          return { type: 'Integer', value: buttonType === 'OKCancel' ? MsgBoxConstants.vbCancel : MsgBoxConstants.vbNo };
        }
      }
    }

    if (options?.console) {
      options.console(finalPrompt);
    } else {
      console.log(finalPrompt);
    }

    if (options?.readline) {
      return options.readline().then((input: string | null) => {
        if (input === null) {
          return { type: 'Integer', value: MsgBoxConstants.vbCancel };
        }
        const result = mapButtonToResult(input, buttonType);
        return { type: 'Integer', value: result };
      }) as unknown as VbValue;
    }

    return { type: 'Integer', value: MsgBoxConstants.vbCancel };
  };
}

export function createBrowserMsgBox() {
  return function(prompt: VbValue, buttons?: VbValue, title?: VbValue): VbValue {
    const message = String(prompt.value ?? prompt);
    const buttonsVal = buttons ? Number(buttons.value ?? buttons) : 0;
    const titleStr = title ? String(title.value ?? title) : 'VBScript';

    const { buttonType, iconType } = parseButtons(buttonsVal);
    const fullPrompt = buildPrompt(titleStr, iconType, message);

    if (buttonType === 'OK') {
      alert(fullPrompt);
      return { type: 'Integer', value: MsgBoxConstants.vbOK };
    }

    if (buttonType === 'OKCancel') {
      const confirmed = confirm(fullPrompt);
      return { type: 'Integer', value: confirmed ? MsgBoxConstants.vbOK : MsgBoxConstants.vbCancel };
    }

    if (buttonType === 'YesNo') {
      const confirmed = confirm(fullPrompt + '\n[OK]=Yes / [Cancel]=No');
      return { type: 'Integer', value: confirmed ? MsgBoxConstants.vbYes : MsgBoxConstants.vbNo };
    }

    if (buttonType === 'YesNoCancel') {
      const input = window.prompt(fullPrompt + '\n\nEnter: Y (Yes) / N (No) / C (Cancel) or invalid value for cancel', '');
      if (input === null) {
        return { type: 'Integer', value: MsgBoxConstants.vbCancel };
      }
      const normalized = input.trim().toUpperCase();
      if (normalized === 'Y' || normalized === 'YES') {
        return { type: 'Integer', value: MsgBoxConstants.vbYes };
      }
      if (normalized === 'N' || normalized === 'NO') {
        return { type: 'Integer', value: MsgBoxConstants.vbNo };
      }
      if (normalized === 'C' || normalized === 'CANCEL' || normalized === '') {
        return { type: 'Integer', value: MsgBoxConstants.vbCancel };
      }
      return { type: 'Integer', value: MsgBoxConstants.vbCancel };
    }

    if (buttonType === 'RetryCancel') {
      const confirmed = confirm(fullPrompt + '\n[OK]=Retry / [Cancel]=Cancel');
      return { type: 'Integer', value: confirmed ? MsgBoxConstants.vbRetry : MsgBoxConstants.vbCancel };
    }

    if (buttonType === 'AbortRetryIgnore') {
      const input = window.prompt(fullPrompt + '\n\nEnter: A (Abort) / R (Retry) / I (Ignore)', '');
      if (input === null) {
        return { type: 'Integer', value: MsgBoxConstants.vbCancel };
      }
      const normalized = input.trim().toUpperCase();
      if (normalized === 'A' || normalized === 'ABORT') {
        return { type: 'Integer', value: MsgBoxConstants.vbAbort };
      }
      if (normalized === 'R' || normalized === 'RETRY') {
        return { type: 'Integer', value: MsgBoxConstants.vbRetry };
      }
      if (normalized === 'I' || normalized === 'IGNORE' || normalized === '') {
        return { type: 'Integer', value: MsgBoxConstants.vbIgnore };
      }
      return { type: 'Integer', value: MsgBoxConstants.vbIgnore };
    }

    alert(fullPrompt);
    return { type: 'Integer', value: MsgBoxConstants.vbOK };
  };
}

export function registerMsgBox(context: { functionRegistry: { register: (name: string, func: (...args: VbValue[]) => VbValue, options?: { isSub: boolean }) => void } }): void {
  context.functionRegistry.register('MsgBox', (promptVal: VbValue, buttons?: VbValue, title?: VbValue): VbValue => {
    const message = String(promptVal.value ?? promptVal);
    const buttonsVal = buttons ? Number(buttons.value ?? buttons) : 0;
    const titleStr = title ? String(title.value ?? title) : null;

    const { buttonType, iconType, defaultButton } = parseButtons(buttonsVal);
    const fullPrompt = buildPrompt(titleStr, iconType, message);

    if (buttonType === 'OK') {
      if (typeof alert !== 'undefined') {
        alert(fullPrompt);
      } else {
        console.log(fullPrompt);
      }
      return { type: 'Integer', value: MsgBoxConstants.vbOK };
    }

    if (typeof window !== 'undefined' && typeof window.confirm === 'function') {
      if (buttonType === 'OKCancel') {
        const confirmed = confirm(fullPrompt);
        return { type: 'Integer', value: confirmed ? MsgBoxConstants.vbOK : MsgBoxConstants.vbCancel };
      }

      if (buttonType === 'YesNo') {
        const confirmed = confirm(fullPrompt + '\n[OK]=Yes / [Cancel]=No');
        return { type: 'Integer', value: confirmed ? MsgBoxConstants.vbYes : MsgBoxConstants.vbNo };
      }

      if (buttonType === 'YesNoCancel' && typeof window.prompt === 'function') {
        const input = window.prompt(fullPrompt + '\n\nEnter: Y (Yes) / N (No) / C (Cancel) or invalid value for cancel', '');
        if (input === null) {
          return { type: 'Integer', value: MsgBoxConstants.vbCancel };
        }
        const normalized = input.trim().toUpperCase();
        if (normalized === 'Y' || normalized === 'YES') {
          return { type: 'Integer', value: MsgBoxConstants.vbYes };
        }
        if (normalized === 'N' || normalized === 'NO') {
          return { type: 'Integer', value: MsgBoxConstants.vbNo };
        }
        return { type: 'Integer', value: MsgBoxConstants.vbCancel };
      }

      if (buttonType === 'RetryCancel') {
        const confirmed = confirm(fullPrompt + '\n[OK]=Retry / [Cancel]=Cancel');
        return { type: 'Integer', value: confirmed ? MsgBoxConstants.vbRetry : MsgBoxConstants.vbCancel };
      }

      if (buttonType === 'AbortRetryIgnore' && typeof window.prompt === 'function') {
        const input = window.prompt(fullPrompt + '\n\nEnter: A (Abort) / R (Retry) / I (Ignore)', '');
        if (input === null) {
          return { type: 'Integer', value: MsgBoxConstants.vbCancel };
        }
        const normalized = input.trim().toUpperCase();
        if (normalized === 'A' || normalized === 'ABORT') {
          return { type: 'Integer', value: MsgBoxConstants.vbAbort };
        }
        if (normalized === 'R' || normalized === 'RETRY') {
          return { type: 'Integer', value: MsgBoxConstants.vbRetry };
        }
        if (normalized === 'I' || normalized === 'IGNORE' || normalized === '') {
          return { type: 'Integer', value: MsgBoxConstants.vbIgnore };
        }
        return { type: 'Integer', value: MsgBoxConstants.vbIgnore };
      }
    }

    console.log(fullPrompt + '\n' + getButtonOptions(buttonType, defaultButton).join(' '));
    return { type: 'Integer', value: MsgBoxConstants.vbCancel };
  }, { isSub: false });
}
