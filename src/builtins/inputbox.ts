import type { VbValue } from '../runtime/index.ts';

export interface InputBoxOptions {
  prompt?: (message: string, defaultValue: string) => string | null;
  console?: (message: string) => void;
  readline?: () => Promise<string | null>;
}

export function createInputBox(options?: InputBoxOptions) {
  return function(
    _context: { functionRegistry: { register: (name: string, func: (...args: VbValue[]) => VbValue) => void } },
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

    if (options?.prompt) {
      const result = options.prompt(message, defaultVal);
      return { type: 'String', value: result ?? '' };
    }

    const promptText = `[${titleStr}]\n${message}\n` + 
      (hasDefault ? `Default: ${defaultVal}\n` : '') + 
      `[Enter] Confirm / [Esc] Cancel (default is "${defaultVal || '(empty)'}")`;

    if (options?.console) {
      options.console(promptText);
    } else {
      console.log(promptText);
    }

    if (options?.readline) {
      async function getInput(): Promise<VbValue> {
        while (true) {
          const input = await options.readline!();
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
      return getInput() as unknown as VbValue;
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

export function registerInputBox(context: { functionRegistry: { register: (name: string, func: (...args: VbValue[]) => VbValue) => void } }): void {
  context.functionRegistry.register('InputBox', (promptVal: VbValue, title?: VbValue, defaultText?: VbValue, _xPos?: VbValue, _yPos?: VbValue, _helpFile?: VbValue, _contextVal?: VbValue): VbValue => {
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

    console.log(`[${titleStr}]\n${message}\nDefault: ${defaultVal || '(empty)'}`);
    return { type: 'String', value: '' };
  });
}
