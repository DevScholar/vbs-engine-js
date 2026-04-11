import type { VbValue } from '../runtime/index.ts';

function _syncReadFromConsole(): string | null {
  try {
    // eslint-disable-next-line @typescript-eslint/ban-ts-comment
    // @ts-ignore - Node.js specific
    // eslint-disable-next-line @typescript-eslint/no-require-imports
    const fs = require('fs');
    // eslint-disable-next-line @typescript-eslint/ban-ts-comment
    // @ts-ignore - Node.js specific
    const chunks: Buffer[] = [];

    while (true) {
      // eslint-disable-next-line @typescript-eslint/ban-ts-comment
      // @ts-ignore - Node.js specific
      const buffer = Buffer.alloc(1024);
      let bytesRead = 0;
      try {
        bytesRead = fs.readSync(0, buffer, 0, buffer.length, null);
      } catch {
        return null;
      }
      if (bytesRead === 0) {
        break;
      }
      const chunk = buffer.subarray(0, bytesRead);
      chunks.push(chunk);
      if (chunk.includes(10)) {
        break;
      }
    }

    if (chunks.length === 0) {
      return null;
    }

    // eslint-disable-next-line @typescript-eslint/ban-ts-comment
    // @ts-ignore - Node.js specific
    return Buffer.concat(chunks).toString('utf8').replace(/\r?\n$/, '');
  } catch {
    return null;
  }
}

function _writeToConsole(text: string): void {
  try {
    // eslint-disable-next-line @typescript-eslint/ban-ts-comment
    // @ts-ignore - Node.js specific
    // eslint-disable-next-line @typescript-eslint/no-require-imports
    const fs = require('fs');
    fs.writeSync(1, text);
  } catch {
    console.log(text);
  }
}

export interface InputBoxOptions {
  prompt?: (message: string, defaultValue: string) => string | null;
  console?: (message: string) => void;
  readline?: () => Promise<string | null>;
}

export function createInputBox(options?: InputBoxOptions) {
  return function (
    _context: {
      functionRegistry: { register: (name: string, func: (...args: VbValue[]) => VbValue) => void };
    },
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
      async function getInput(): Promise<VbValue> {
        while (true) {
          const input = await readline();
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

      // Node.js interactive fallback using synchronous stdin
      const promptText =
        `[${titleStr}]\n${message}` +
        (defaultVal ? ` [${defaultVal}]` : '') +
        ': ';
      _writeToConsole(promptText);
      while (true) {
        const input = _syncReadFromConsole();
        if (input === null) {
          _writeToConsole('\n');
          return { type: 'String', value: '' };
        }
        if (input === '' && defaultVal !== '') {
          return { type: 'String', value: defaultVal };
        }
        if (input !== '') {
          return { type: 'String', value: input };
        }
        _writeToConsole('Input cannot be empty. Please try again.\n' + promptText);
      }
    }
  );
}
