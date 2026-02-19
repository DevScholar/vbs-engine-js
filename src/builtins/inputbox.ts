import type { VbValue } from '../runtime/index.ts';

export interface InputBoxOptions {
  prompt?: (message: string, defaultValue: string) => string | null;
  console?: (message: string) => void;
  readline?: () => Promise<string | null>;
}

export function createInputBox(options?: InputBoxOptions) {
  return function(
    context: { functionRegistry: { register: (name: string, func: (...args: VbValue[]) => VbValue) => void } },
    promptVal: VbValue,
    title?: VbValue,
    defaultText?: VbValue,
    xPos?: VbValue,
    yPos?: VbValue,
    helpFile?: VbValue,
    contextVal?: VbValue
  ): VbValue {
    const message = String(promptVal.value ?? promptVal);
    const titleStr = title ? String(title.value ?? title) : 'Input';
    const defaultVal = defaultText ? String(defaultText.value ?? defaultText) : '';

    if (options?.prompt) {
      const result = options.prompt(message, defaultVal);
      return { type: 'String', value: result ?? '' };
    }

    if (options?.console) {
      options.console(`[${titleStr}]\n${message}\nDefault: ${defaultVal}\nPress Enter to confirm or Esc to cancel`);
    } else {
      console.log(`[${titleStr}]\n${message}\nDefault: ${defaultVal}\nPress Enter to confirm or Esc to cancel`);
    }

    if (options?.readline) {
      return options.readline().then((input: string | null) => {
        if (input === null) {
          return { type: 'String', value: '' };
        }
        return { type: 'String', value: input || defaultVal };
      }) as unknown as VbValue;
    }

    return { type: 'String', value: '' };
  };
}

export function createBrowserInputBox() {
  return function(
    prompt: VbValue,
    title?: VbValue,
    defaultVal?: VbValue,
    _xPos?: VbValue,
    _yPos?: VbValue,
    _helpFile?: VbValue,
    _contextVal?: VbValue
  ): VbValue {
    const message = String(prompt.value ?? prompt);
    const titleStr = title ? String(title.value ?? title) : 'Input';
    const def = defaultVal ? String(defaultVal.value ?? defaultVal) : '';
    const result = window.prompt(`[${titleStr}]\n${message}`, def);
    return { type: 'String', value: result ?? '' };
  };
}

export function registerInputBox(context: { functionRegistry: { register: (name: string, func: (...args: VbValue[]) => VbValue) => void } }): void {
  context.functionRegistry.register('InputBox', (promptVal: VbValue, title?: VbValue, defaultText?: VbValue, xPos?: VbValue, yPos?: VbValue, helpFile?: VbValue, contextVal?: VbValue): VbValue => {
    const message = String(promptVal.value ?? promptVal);
    const titleStr = title ? String(title.value ?? title) : 'Input';
    const defaultVal = defaultText ? String(defaultText.value ?? defaultText) : '';

    if (typeof window !== 'undefined' && typeof window.prompt === 'function') {
      const result = window.prompt(`[${titleStr}]\n${message}`, defaultVal);
      return { type: 'String', value: result ?? '' };
    }

    console.log(`[${titleStr}]\n${message}\nDefault: ${defaultVal}`);
    return { type: 'String', value: '' };
  });
}
