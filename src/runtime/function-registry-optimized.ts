/**
 * Optimized Function Registry with String Interning and Fast Path
 * 
 * Performance improvements:
 * 1. String interning for function names
 * 2. Direct function reference cache for hot functions
 * 3. Reduced overhead for common built-in functions
 */

import type { VbValue } from './values.ts';
import { VbEmpty } from './values.ts';
import { globalStringInterner } from './string-interner.ts';

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

/**
 * Hot function cache entry
 */
interface HotFunctionEntry {
  info: VbFunctionInfo;
  callCount: number;
}

export class VbFunctionRegistry {
  private functions: Map<string, VbFunctionInfo> = new Map();
  
  // Hot function cache for frequently called functions
  private hotFunctions: Map<string, HotFunctionEntry> = new Map();
  private readonly HOT_FUNCTION_THRESHOLD = 5;
  
  // Fast path for common built-in functions
  private fastPathFunctions: Map<string, VbFunction> = new Map();

  /**
   * Register a function with string interning for the name.
   */
  register(name: string, func: VbFunction | VbSub | VbUserFunction, options: Partial<VbFunctionInfo> = {}): void {
    const internedName = globalStringInterner.intern(name);
    const funcInfo: VbFunctionInfo = {
      name,
      func,
      isSub: options.isSub ?? false,
      minArgs: options.minArgs ?? 0,
      maxArgs: options.maxArgs ?? Infinity,
      params: options.params ?? [],
      isUserDefined: options.isUserDefined ?? false,
    };
    
    this.functions.set(internedName, funcInfo);
    
    // Add to fast path if it's a simple built-in function
    if (!funcInfo.isUserDefined && !funcInfo.isSub && funcInfo.minArgs === funcInfo.maxArgs) {
      this.fastPathFunctions.set(internedName, func as VbFunction);
    }
  }

  /**
   * Get function info with string interning.
   */
  get(name: string): VbFunctionInfo | undefined {
    const internedName = globalStringInterner.intern(name);
    return this.functions.get(internedName);
  }

  /**
   * Check if function exists with string interning.
   */
  has(name: string): boolean {
    const internedName = globalStringInterner.intern(name);
    return this.functions.has(internedName);
  }

  /**
   * Call a function with optimized fast path for hot functions.
   */
  call(name: string, args: VbValue[]): VbValue {
    const internedName = globalStringInterner.intern(name);
    
    // Try fast path for simple built-in functions
    const fastFunc = this.fastPathFunctions.get(internedName);
    if (fastFunc !== undefined) {
      return fastFunc(...args) ?? VbEmpty;
    }
    
    // Check hot function cache
    const hotEntry = this.hotFunctions.get(internedName);
    if (hotEntry !== undefined) {
      hotEntry.callCount++;
      return this.executeFunction(hotEntry.info, args);
    }
    
    // Standard lookup
    const info = this.functions.get(internedName);
    if (!info) {
      throw new Error(`Undefined function or sub: ${name}`);
    }

    // Promote to hot cache if called frequently
    if (info.isUserDefined) {
      this.hotFunctions.set(internedName, { info, callCount: 1 });
    }

    return this.executeFunction(info, args);
  }

  /**
   * Execute a function with argument validation.
   */
  private executeFunction(info: VbFunctionInfo, args: VbValue[]): VbValue {
    if (args.length < info.minArgs) {
      throw new Error(`${info.name} requires at least ${info.minArgs} argument(s)`);
    }

    if (args.length > info.maxArgs) {
      throw new Error(`${info.name} accepts at most ${info.maxArgs} argument(s)`);
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

  /**
   * Call function with argument references (for ByRef support).
   */
  callWithRefs(name: string, args: VbArgumentRef[]): VbValue {
    const internedName = globalStringInterner.intern(name);
    
    const info = this.functions.get(internedName);
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

    // Handle ByRef parameter updates
    for (let i = 0; i < args.length; i++) {
      const arg = args[i];
      if (arg?.setValue && i < info.params.length && info.params[i]!.byRef) {
        arg.setValue(values[i]!);
      }
    }
    
    if (info.isSub) {
      return VbEmpty;
    }

    return result ?? VbEmpty;
  }

  /**
   * Get all registered function names.
   */
  getAllNames(): string[] {
    return Array.from(this.functions.values()).map(f => f.name);
  }

  /**
   * Clear hot function cache (useful for testing).
   */
  clearHotCache(): void {
    this.hotFunctions.clear();
  }

  /**
   * Get statistics about hot functions.
   */
  getHotFunctionStats(): Array<{ name: string; callCount: number }> {
    return Array.from(this.hotFunctions.entries())
      .map(([name, entry]) => ({ name, callCount: entry.callCount }))
      .sort((a, b) => b.callCount - a.callCount);
  }
}
