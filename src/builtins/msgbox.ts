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

export type MsgBoxButtonType =
  | 'OK'
  | 'OKCancel'
  | 'AbortRetryIgnore'
  | 'YesNoCancel'
  | 'YesNo'
  | 'RetryCancel';
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

function getButtonOptions(buttonType: MsgBoxButtonType, defaultButton: number): string {
  let options = '';
  let defaultLabel = '';
  switch (buttonType) {
    case 'OK':
      options = '[O] OK';
      defaultLabel = 'O';
      break;
    case 'OKCancel':
      if (defaultButton === 256) {
        options = '[O] OK / [C] Cancel';
        defaultLabel = 'O';
      } else {
        options = '[O] OK / [C] Cancel';
        defaultLabel = 'C';
      }
      break;
    case 'YesNo':
      if (defaultButton === 256) {
        options = '[Y] Yes / [N] No';
        defaultLabel = 'Y';
      } else {
        options = '[Y] Yes / [N] No';
        defaultLabel = 'N';
      }
      break;
    case 'YesNoCancel':
      if (defaultButton === 512) {
        options = '[Y] Yes / [N] No / [C] Cancel';
        defaultLabel = 'C';
      } else if (defaultButton === 256) {
        options = '[Y] Yes / [N] No / [C] Cancel';
        defaultLabel = 'Y';
      } else {
        options = '[Y] Yes / [N] No / [C] Cancel';
        defaultLabel = 'N';
      }
      break;
    case 'RetryCancel':
      if (defaultButton === 256) {
        options = '[R] Retry / [C] Cancel';
        defaultLabel = 'R';
      } else {
        options = '[R] Retry / [C] Cancel';
        defaultLabel = 'C';
      }
      break;
    case 'AbortRetryIgnore':
      if (defaultButton === 512) {
        options = '[A] Abort / [R] Retry / [I] Ignore';
        defaultLabel = 'I';
      } else if (defaultButton === 256) {
        options = '[A] Abort / [R] Retry / [I] Ignore';
        defaultLabel = 'R';
      } else {
        options = '[A] Abort / [R] Retry / [I] Ignore';
        defaultLabel = 'A';
      }
      break;
  }
  return options + ` (default is "${defaultLabel}"): `;
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
  return new Promise(resolve => {
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
      output: process.stdout,
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
      output: process.stdout,
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
  return function (
    _context: {
      functionRegistry: {
        register: (
          name: string,
          func: (...args: VbValue[]) => VbValue,
          options?: { isSub: boolean }
        ) => void;
      };
    },
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

    const finalPrompt = fullPrompt + '\n' + buttonOptions;

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
      async function getInput(): Promise<VbValue> {
        while (true) {
          const input = await readline();
          if (input === null) {
            return { type: 'Integer', value: MsgBoxConstants.vbCancel };
          }
          const trimmed = input.trim().toLowerCase();
          if (trimmed === '') {
            const result = getDefaultResult(buttonType, defaultButton);
            return { type: 'Integer', value: result };
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
      return getInput() as unknown as VbValue;
    }

    return { type: 'Integer', value: MsgBoxConstants.vbCancel };
  };
}

function getDefaultResult(buttonType: MsgBoxButtonType, defaultButton: number): number {
  switch (buttonType) {
    case 'OK':
      return MsgBoxConstants.vbOK;
    case 'OKCancel':
      return defaultButton === 256 ? MsgBoxConstants.vbOK : MsgBoxConstants.vbCancel;
    case 'YesNo':
      return defaultButton === 256 ? MsgBoxConstants.vbYes : MsgBoxConstants.vbNo;
    case 'YesNoCancel':
      if (defaultButton === 512) return MsgBoxConstants.vbCancel;
      if (defaultButton === 256) return MsgBoxConstants.vbYes;
      return MsgBoxConstants.vbNo;
    case 'RetryCancel':
      return defaultButton === 256 ? MsgBoxConstants.vbRetry : MsgBoxConstants.vbCancel;
    case 'AbortRetryIgnore':
      if (defaultButton === 512) return MsgBoxConstants.vbIgnore;
      if (defaultButton === 256) return MsgBoxConstants.vbRetry;
      return MsgBoxConstants.vbAbort;
    default:
      return MsgBoxConstants.vbCancel;
  }
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

      console.log(finalPrompt);
      return { type: 'Integer', value: MsgBoxConstants.vbCancel };
    },
    { isSub: false }
  );
}
