import type { VbContext, VbValue } from '../runtime/index.ts';
import {
  createVbError,
  VbError,
  VbErrorCodes,
  VbErrorDescriptions,
  parseVbNumericString,
  roundHalfToEven,
  toNumber,
} from '../runtime/index.ts';
import { stringFunctions } from './string.ts';
import { mathFunctions, constants } from './math.ts';
import { dateFunctions } from './date.ts';
import { conversionFunctions, inspectionFunctions } from './conversion.ts';
import { arrayFunctions } from './array.ts';
import { registerMsgBox } from './msgbox.ts';
import { registerInputBox } from './inputbox.ts';
import { registerRegExp } from './regexp.ts';
import { localeFunctions } from './locale.ts';
import { createCollection } from './collection.ts';

// VbObjectValueData protocol names the VBS engine checks via typeof === 'function'.
// Hide them so the engine falls through to obj[propertyName] → jsfunction path,
// which auto-converts VbValue ↔ JS primitives via vbToJs / jsToVb.
const VBS_PROTOCOL_NAMES = new Set([
  'getProperty',
  'setProperty',
  'hasMethod',
  'getMethod',
  'hasProperty',
  'call',
]);

// Wraps a node-ps1-dotnet COM proxy so the VBS engine treats it as a plain JS
// object. Method return values that are objects/functions are recursively wrapped.
function wrapCOMProxy(ax: unknown): Record<string, unknown> {
  return new Proxy(Object.create(null) as Record<string, unknown>, {
    get(_target: Record<string, unknown>, prop: string | symbol): unknown {
      if (typeof prop === 'symbol') {
        if (prop === Symbol.iterator) {
          const axIterable = ax as Record<symbol, unknown>;
          if (typeof axIterable[Symbol.iterator] === 'function') {
            return function* () {
              for (const item of axIterable as unknown as Iterable<unknown>) {
                if (item !== null && (typeof item === 'object' || typeof item === 'function')) {
                  yield wrapCOMProxy(item);
                } else {
                  yield item;
                }
              }
            };
          }
        }
        return undefined;
      }
      if (VBS_PROTOCOL_NAMES.has(prop)) return undefined;
      const val = (ax as Record<string, unknown>)[prop];
      if (typeof val !== 'function') return val;
      // Wrap the method function so returned COM objects are also wrapped.
      return (...args: unknown[]) => {
        const result = (val as (...args: unknown[]) => unknown)(...args);
        if (result !== null && (typeof result === 'object' || typeof result === 'function'))
          return wrapCOMProxy(result);
        return result;
      };
    },
    has(_target: Record<string, unknown>, prop: string | symbol): boolean {
      if (typeof prop === 'symbol') {
        if (prop === Symbol.iterator) {
          const axIterable = ax as Record<symbol, unknown>;
          return typeof axIterable[Symbol.iterator] === 'function';
        }
        return false;
      }
      if (VBS_PROTOCOL_NAMES.has(prop)) return false;
      return prop in (ax as Record<string, unknown>);
    },
    set(_target: Record<string, unknown>, prop: string | symbol, value: unknown): boolean {
      if (typeof prop === 'symbol') return false;
      (ax as Record<string, unknown>)[prop] = value;
      return true;
    },
  });
}

/**
 * The number an error keeps, which Windows rounds to a whole number and then holds only inside a range.
 *
 * Measured with the error probe: 7 stays 7, 5.7 becomes 6, "9" becomes 9, -5 stays -5, and 70000 becomes 5,
 * the invalid procedure call. A negative number is kept, which is how Err.Raise vbObjectError + 2000 arrives
 * as -2147219504. A text that is no number is a type mismatch, and a Null is one too.
 *
 * `zeroIsKept` is the one place Err.Number and Err.Raise differ: writing the property a zero leaves a zero,
 * while Err.Raise 0 answers 5, measured on a fresh cscript.
 */
function readErrorNumber(value: VbValue, zeroIsKept: boolean): number {
  if (value.type === 'Null') {
    throw createVbError(VbErrorCodes.TypeMismatch, 'Type mismatch: Null');
  }
  let written: number;
  if (value.type === 'String') {
    const parsed = parseVbNumericString(value.value.trim());
    if (parsed === null) throw createVbError(VbErrorCodes.TypeMismatch, 'Type mismatch');
    written = parsed;
  } else {
    written = toNumber(value);
  }
  const whole = roundHalfToEven(written);
  if (whole === 0) return zeroIsKept ? 0 : 5;
  if (whole < 0) return whole;
  return whole <= 65535 ? whole : 5;
}

/** The source Windows puts on an error nobody named one for. */
const RAISED_SOURCE = 'Microsoft VBScript runtime error';

export function registerBuiltins(context: VbContext): void {
  Object.entries(stringFunctions).forEach(([name, func]) => {
    context.functionRegistry.register(name, func);
  });

  Object.entries(mathFunctions).forEach(([name, func]) => {
    context.functionRegistry.register(name, func);
  });

  Object.entries(dateFunctions).forEach(([name, func]) => {
    context.functionRegistry.register(name, func);
  });

  Object.entries(conversionFunctions).forEach(([name, func]) => {
    context.functionRegistry.register(name, func);
  });

  Object.entries(inspectionFunctions).forEach(([name, func]) => {
    context.functionRegistry.register(name, func);
  });

  Object.entries(arrayFunctions).forEach(([name, func]) => {
    context.functionRegistry.register(name, func);
  });

  Object.entries(constants).forEach(([name, value]) => {
    context.globalScope.declare(name, value);
  });

  registerMsgBox(context);
  registerInputBox(context);
  registerRegExp(context);

  Object.entries(localeFunctions).forEach(([name, func]) => {
    context.functionRegistry.register(name, func);
  });

  // Register Collection as a native built-in class
  context.classRegistry.registerNative('Collection', createCollection);

  context.functionRegistry.register('ScriptEngine', (): VbValue => {
    return { type: 'String', value: 'VBScript' };
  });

  context.functionRegistry.register('ScriptEngineMajorVersion', (): VbValue => {
    return { type: 'Integer', value: 10 };
  });

  context.functionRegistry.register('ScriptEngineMinorVersion', (): VbValue => {
    return { type: 'Integer', value: 8 };
  });

  context.functionRegistry.register('ScriptEngineBuildVersion', (): VbValue => {
    return { type: 'Long', value: 16384 };
  });

  context.functionRegistry.register('Eval', (expression: VbValue): VbValue => {
    const code = String(expression.value ?? expression);
    const result = context.evaluate?.(code);
    return result ?? { type: 'Empty', value: undefined };
  });

  context.functionRegistry.register(
    'Execute',
    (expression: VbValue): VbValue => {
      const code = String(expression.value ?? expression);
      const result = context.execute?.(code);
      return result ?? { type: 'Empty', value: undefined };
    },
    { isSub: true }
  );

  context.functionRegistry.register(
    'ExecuteGlobal',
    (expression: VbValue): VbValue => {
      const code = String(expression.value ?? expression);
      const result = context.executeGlobal?.(code);
      return result ?? { type: 'Empty', value: undefined };
    },
    { isSub: true }
  );

  context.functionRegistry.register(
    'Print',
    (...args: VbValue[]): VbValue => {
      console.log(...args.map(a => a.value ?? a));
      return { type: 'Empty', value: undefined };
    },
    { isSub: true }
  );

  context.functionRegistry.register('GetRef', (procname: VbValue): VbValue => {
    const name = String(procname.value ?? procname);
    const funcRegistry = context.functionRegistry;
    return {
      type: 'Object',
      value: {
        type: 'vbref',
        name,
        getProperty: (propName: string): VbValue => {
          if (propName.toLowerCase() === 'name') {
            return { type: 'String', value: name };
          }
          return { type: 'Empty', value: undefined };
        },
        hasProperty: (propName: string): boolean => {
          return propName.toLowerCase() === 'name';
        },
        call: (...args: VbValue[]): VbValue => {
          return funcRegistry.call(name, args);
        },
      },
    };
  });

  context.functionRegistry.register('TypeName', (expression: VbValue): VbValue => {
    if (expression.type === 'Object' && expression.value === null)
      return { type: 'String', value: 'Nothing' };
    if (expression.type === 'Object' && expression.value && typeof expression.value === 'object') {
      const obj = expression.value as { classInfo?: { name: string } };
      if (obj.classInfo && obj.classInfo.name) {
        return { type: 'String', value: obj.classInfo.name };
      }
    }
    const typeNames: Record<string, string> = {
      Empty: 'Empty',
      Null: 'Null',
      Integer: 'Integer',
      Long: 'Long',
      LongLong: 'LongLong',
      Single: 'Single',
      Double: 'Double',
      Currency: 'Currency',
      Date: 'Date',
      String: 'String',
      Object: 'Object',
      Error: 'Error',
      Boolean: 'Boolean',
      Variant: 'Variant',
      Byte: 'Byte',
      Array: 'Variant()',
    };
    return { type: 'String', value: typeNames[expression.type] ?? 'Variant' };
  });

  context.functionRegistry.register(
    'Erase',
    (_arrayname: VbValue): VbValue => {
      void _arrayname; // Intentionally unused - matches VBScript signature
      return { type: 'Empty', value: undefined };
    },
    { isSub: true }
  );

  context.functionRegistry.register('GetObject', (pathname?: VbValue, cls?: VbValue): VbValue => {
    const path = pathname ? String(pathname.value ?? pathname) : '';
    const className = cls ? String(cls.value ?? cls) : '';

    const getObjectFn = (
      globalThis as unknown as {
        GetObject?: (pathname?: string, cls?: string) => unknown;
      }
    ).GetObject;

    if (getObjectFn) {
      try {
        const ax = getObjectFn(path || undefined, className || undefined);
        return { type: 'Object', value: wrapCOMProxy(ax) };
      } catch {
        throw createVbError(
          VbErrorCodes.ActiveXComponentCantCreateObject,
          "ActiveX component can't create object"
        );
      }
    }

    return { type: 'Object', value: null };
  });

  context.functionRegistry.register(
    'CreateObject',
    (cls: VbValue, _servername?: VbValue): VbValue => {
      void _servername; // Intentionally unused - matches VBScript signature
      const className = String(cls.value ?? cls);
      const axConstructor = (
        globalThis as unknown as { ActiveXObject?: new (cls: string) => unknown }
      ).ActiveXObject;
      if (axConstructor) {
        try {
          const ax = new axConstructor(className);
          return { type: 'Object', value: wrapCOMProxy(ax) };
        } catch {
          throw createVbError(
            VbErrorCodes.ActiveXComponentCantCreateObject,
            "ActiveX component can't create object"
          );
        }
      }
      return { type: 'Object', value: { className, properties: new Map() } };
    }
  );

  context.functionRegistry.register('LoadPicture', (picturename: VbValue): VbValue => {
    const path = String(picturename.value ?? picturename);
    const imageObj = {
      classInfo: { name: 'IPictureDisp' },
      getProperty: (name: string): VbValue => {
        const lowerName = name.toLowerCase();
        if (lowerName === 'handle') {
          return { type: 'Long', value: 0 };
        }
        if (lowerName === 'width') {
          return { type: 'Long', value: 0 };
        }
        if (lowerName === 'height') {
          return { type: 'Long', value: 0 };
        }
        if (lowerName === 'type') {
          return { type: 'Long', value: 1 };
        }
        return { type: 'Empty', value: undefined };
      },
      hasProperty: (name: string): boolean => {
        const lowerName = name.toLowerCase();
        return ['handle', 'width', 'height', 'type'].includes(lowerName);
      },
      _path: path,
    };
    return { type: 'Object', value: imageObj };
  });

  context.functionRegistry.register(
    'RGB',
    (red: VbValue, green: VbValue, blue: VbValue): VbValue => {
      const r = Math.max(0, Math.min(255, Math.floor(Number(red.value ?? 0))));
      const g = Math.max(0, Math.min(255, Math.floor(Number(green.value ?? 0))));
      const b = Math.max(0, Math.min(255, Math.floor(Number(blue.value ?? 0))));
      const rgbValue = r + g * 256 + b * 65536;
      return { type: 'Long', value: rgbValue };
    }
  );

  context.functionRegistry.register('QBColor', (color: VbValue): VbValue => {
    const colorIndex = Math.floor(Number(color.value ?? 0));
    const colors = [
      0x000000, 0x800000, 0x008000, 0x808000, 0x000080, 0x800080, 0x008080, 0xc0c0c0, 0x808080,
      0xff0000, 0x00ff00, 0xffff00, 0x0000ff, 0xff00ff, 0x00ffff, 0xffffff,
    ];
    const rgbValue = colors[colorIndex] ?? 0;
    return { type: 'Long', value: rgbValue };
  });

  const errObject = {
    getProperty: (name: string): VbValue => {
      const lowerName = name.toLowerCase();
      if (lowerName === 'number') {
        return { type: 'Long', value: context.err.number };
      }
      if (lowerName === 'description') {
        return { type: 'String', value: context.err.description };
      }
      if (lowerName === 'source') {
        return { type: 'String', value: context.err.source };
      }
      if (lowerName === 'helpfile') {
        return { type: 'String', value: context.err.helpFile };
      }
      if (lowerName === 'helpcontext') {
        return { type: 'Long', value: context.err.helpContext };
      }
      if (lowerName === 'lastdllerror') {
        // Not wired to any real DLL loader; VBScript reports 0 when unused.
        return { type: 'Long', value: 0 };
      }
      return { type: 'Empty', value: undefined };
    },
    setProperty: (name: string, value: VbValue): void => {
      const lowerName = name.toLowerCase();
      if (lowerName === 'number') {
        context.err.number = readErrorNumber(value, true);
      } else if (lowerName === 'description') {
        context.err.description = String(value.value ?? '');
      } else if (lowerName === 'source') {
        context.err.source = String(value.value ?? '');
      } else if (lowerName === 'helpfile') {
        context.err.helpFile = String(value.value ?? '');
      } else if (lowerName === 'helpcontext') {
        context.err.helpContext = Number(value.value) || 0;
      }
    },
    hasMethod: (name: string): boolean => {
      return name.toLowerCase() === 'clear' || name.toLowerCase() === 'raise';
    },
    getMethod: (name: string): { func: (...args: VbValue[]) => VbValue } => {
      const lowerName = name.toLowerCase();
      if (lowerName === 'clear') {
        return {
          func: (): VbValue => {
            context.clearError();
            return { type: 'Empty', value: undefined };
          },
        };
      }
      if (lowerName === 'raise') {
        return {
          func: (
            number: VbValue,
            source?: VbValue,
            description?: VbValue,
            helpFile?: VbValue,
            helpContext?: VbValue
          ): VbValue => {
            const raised = readErrorNumber(number, false);
            // Windows fills in what the caller left out: the runtime's own name for the source, and the
            // standard text for the number as the description ("Unknown runtime error" where the number has
            // no standard text of its own, which Err.Raise 60000 answers).
            const raisedSource = source ? String(source.value ?? '') : RAISED_SOURCE;
            const raisedDescription = description
              ? String(description.value ?? '')
              : (VbErrorDescriptions[raised] ?? 'Unknown runtime error');
            const raisedHelpFile = helpFile ? String(helpFile.value ?? '') : '';
            const raisedHelpContext = helpContext ? Number(helpContext.value) || 0 : 0;

            // The statement fills the Err object and then raises, so a script with no handler stops there.
            // Measured on a fresh cscript: a script that writes a line, then `Err.Raise 4321, "a source", "a
            // description"`, then another line prints the first line and then
            // `(4, 1) a source: a description`, and the line after it never runs. The fields are set here as
            // well as carried on the error, because a handler that resumes leaves them in Err to be read.
            context.err.number = raised;
            context.err.source = raisedSource;
            context.err.description = raisedDescription;
            context.err.helpFile = raisedHelpFile;
            context.err.helpContext = raisedHelpContext;

            const raisedError = new VbError(raised, raisedDescription, raisedSource);
            raisedError.helpFile = raisedHelpFile;
            raisedError.helpContext = raisedHelpContext;
            throw raisedError;
          },
        };
      }
      throw new Error(`Unknown Err method: ${name}`);
    },
  };

  context.globalScope.declare('Err', { type: 'Object', value: errObject });
}

export { stringFunctions } from './string.ts';
export { mathFunctions, constants } from './math.ts';
export { dateFunctions } from './date.ts';
export { conversionFunctions, inspectionFunctions } from './conversion.ts';
export { arrayFunctions } from './array.ts';
