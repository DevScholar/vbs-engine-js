import type { VbValue } from './values.ts';
import { VbEmpty } from './values.ts';

export type VbFunction = (...args: VbValue[]) => VbValue;
export type VbSub = (...args: VbValue[]) => void;
export type VbUserFunction = (args: VbValue[]) => VbValue;

export interface VbParameterInfo {
  name: string;
  byRef: boolean;
  isArray: boolean;
}

export interface VbFunctionInfo {
  name: string;
  func: VbFunction | VbSub | VbUserFunction;
  isSub: boolean;
  minArgs: number;
  maxArgs: number;
  params: VbParameterInfo[];
  isUserDefined: boolean;
}

export interface VbArgumentRef {
  value: VbValue;
  variableName?: string;
  setValue?: (value: VbValue) => void;
}

export class VbFunctionRegistry {
  private functions: Map<string, VbFunctionInfo> = new Map();

  register(name: string, func: VbFunction | VbSub | VbUserFunction, options: Partial<VbFunctionInfo> = {}): void {
    const lowerName = name.toLowerCase();
    this.functions.set(lowerName, {
      name,
      func,
      isSub: options.isSub ?? false,
      minArgs: options.minArgs ?? 0,
      maxArgs: options.maxArgs ?? Infinity,
      params: options.params ?? [],
      isUserDefined: options.isUserDefined ?? false,
    });
  }

  get(name: string): VbFunctionInfo | undefined {
    return this.functions.get(name.toLowerCase());
  }

  has(name: string): boolean {
    return this.functions.has(name.toLowerCase());
  }

  call(name: string, args: VbValue[]): VbValue {
    const info = this.functions.get(name.toLowerCase());
    if (!info) {
      throw new Error(`Undefined function or sub: ${name}`);
    }

    if (args.length < info.minArgs) {
      throw new Error(`${name} requires at least ${info.minArgs} argument(s)`);
    }

    if (args.length > info.maxArgs) {
      throw new Error(`${name} accepts at most ${info.maxArgs} argument(s)`);
    }

    let result: VbValue;
    if (info.isUserDefined) {
      result = (info.func as VbUserFunction)(args);
    } else {
      result = (info.func as VbFunction)(...args);
    }
    
    if (info.isSub) {
      return VbEmpty;
    }

    return result ?? VbEmpty;
  }

  callWithRefs(name: string, args: VbArgumentRef[]): VbValue {
    const info = this.functions.get(name.toLowerCase());
    if (!info) {
      throw new Error(`Undefined function or sub: ${name}`);
    }

    if (args.length < info.minArgs) {
      throw new Error(`${name} requires at least ${info.minArgs} argument(s)`);
    }

    if (args.length > info.maxArgs) {
      throw new Error(`${name} accepts at most ${info.maxArgs} argument(s)`);
    }

    const values = args.map(arg => arg.value);
    
    let result: VbValue;
    if (info.isUserDefined) {
      result = (info.func as VbUserFunction)(values);
    } else {
      result = (info.func as VbFunction)(...values);
    }
    
    for (let i = 0; i < args.length && i < info.params.length; i++) {
      const param = info.params[i];
      const arg = args[i];
      if (param && arg && param.byRef && arg.setValue && values[i] !== arg.value) {
        arg.setValue(values[i]!);
      }
    }
    
    if (info.isSub) {
      return VbEmpty;
    }

    return result ?? VbEmpty;
  }

  getAll(): IterableIterator<[string, VbFunctionInfo]> {
    return this.functions.entries();
  }

  getUserDefinedFunctions(): Map<string, VbFunctionInfo> {
    const result = new Map<string, VbFunctionInfo>();
    for (const [name, info] of this.functions.entries()) {
      if (info.isUserDefined) {
        result.set(name, info);
      }
    }
    return result;
  }
}
