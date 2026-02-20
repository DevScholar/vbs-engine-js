/**
 * Optimized VbContext with Performance Improvements
 * 
 * Uses optimized scope and function registry with:
 * 1. String interning for variable names
 * 2. Scope pooling to reduce GC pressure
 * 3. Cached value creation
 * 4. Optimized global variable lookup
 */

import type { VbValue } from './values.ts';
import { VbEmpty } from './values.ts';
import { Vbscope, globalScopePool } from './scope-optimized.ts';
import { VbFunctionRegistry } from './function-registry-optimized.ts';
import { VbClassRegistry, VbObjectInstance } from './class-registry.ts';
import { VbError } from './errors.ts';
import { globalStringInterner } from './string-interner.ts';
import { createCachedVbValue } from './value-pool.ts';

// Pre-intern common global keys for fast lookup
const GLOBAL_KEYS = [
  'eval', 'parseInt', 'parseFloat', 'isNaN', 'isFinite',
  'decodeURI', 'decodeURIComponent', 'encodeURI', 'encodeURIComponent',
  'escape', 'unescape', 'console', 'Math', 'Date', 'Array', 'Object',
  'String', 'Number', 'Boolean', 'JSON', 'RegExp', 'Error',
  'window', 'document', 'navigator', 'location', 'history',
];

// Build a fast lookup set for global keys
const globalKeySet = new Set(GLOBAL_KEYS.map(k => globalStringInterner.intern(k)));

// Cache for globalThis properties
const globalThisCache: Map<string, VbValue> = new Map();

function jsToVbOptimized(value: unknown): VbValue {
  return createCachedVbValue(value);
}

export class VbContextOptimized {
  public globalScope: Vbscope;
  public currentScope: Vbscope;
  public functionRegistry: VbFunctionRegistry;
  public classRegistry: VbClassRegistry;
  public optionExplicit: boolean = false;
  public onErrorResumeNext: boolean = false;
  public lastError: VbError | null = null;
  public err: { number: number; description: string; source: string } = {
    number: 0,
    description: '',
    source: '',
  };
  public currentInstance: VbObjectInstance | null = null;
  public inPropertyGet: boolean = false;
  public propertyGetName: string = '';
  public evaluate: ((code: string) => VbValue) | null = null;
  public checkTimeout: (() => void) | null = null;
  public execute: ((code: string) => VbValue) | null = null;
  public executeGlobal: ((code: string) => VbValue) | null = null;

  private withStack: VbValue[] = [];
  private callStack: string[] = [];
  private exitFlag: 'none' | 'sub' | 'function' | 'property' | 'do' | 'for' | 'select' = 'none';
  
  // Track pooled scopes for cleanup
  private pooledScopes: Vbscope[] = [];

  constructor() {
    this.globalScope = new Vbscope();
    this.currentScope = this.globalScope;
    this.functionRegistry = new VbFunctionRegistry();
    this.classRegistry = new VbClassRegistry();
  }

  /**
   * Push a new scope using the scope pool for better performance.
   */
  pushScope(): Vbscope {
    const scope = globalScopePool.acquire(this.currentScope);
    this.pooledScopes.push(scope);
    this.currentScope = scope;
    return scope;
  }

  /**
   * Pop scope and return it to the pool.
   */
  popScope(): void {
    if (this.currentScope.parent) {
      const scopeToRelease = this.currentScope;
      this.currentScope = this.currentScope.parent;
      
      // Remove from tracked scopes
      const index = this.pooledScopes.indexOf(scopeToRelease);
      if (index > -1) {
        this.pooledScopes.splice(index, 1);
      }
      
      // Return to pool
      globalScopePool.release(scopeToRelease);
    }
  }

  pushWith(obj: VbValue): void {
    this.withStack.push(obj);
  }

  popWith(): void {
    this.withStack.pop();
  }

  getCurrentWith(): VbValue | undefined {
    return this.withStack[this.withStack.length - 1];
  }

  pushCall(name: string): void {
    this.callStack.push(globalStringInterner.intern(name));
  }

  popCall(): void {
    this.callStack.pop();
  }

  getCurrentCall(): string | undefined {
    return this.callStack[this.callStack.length - 1];
  }

  setExitFlag(flag: typeof this.exitFlag): void {
    this.exitFlag = flag;
  }

  getExitFlag(): typeof this.exitFlag {
    return this.exitFlag;
  }

  clearExitFlag(): void {
    this.exitFlag = 'none';
  }

  setError(error: VbError): void {
    this.lastError = error;
    this.err.number = error.number;
    this.err.description = error.description;
    this.err.source = error.source;
  }

  clearError(): void {
    this.lastError = null;
    this.err.number = 0;
    this.err.description = '';
    this.err.source = '';
  }

  declareVariable(name: string, value: VbValue = VbEmpty): void {
    this.currentScope.declare(name, value);
  }

  /**
   * Optimized variable lookup with string interning and caching.
   */
  getVariable(name: string): VbValue {
    const internedName = globalStringInterner.intern(name);
    
    // Check current instance first
    if (this.currentInstance) {
      if (this.currentInstance.hasProperty(name)) {
        return this.currentInstance.getProperty(name);
      }
    }
    
    // Check current scope
    const variable = this.currentScope.get(name);
    if (variable) {
      return variable.value;
    }

    // Check globalThis cache
    const cached = globalThisCache.get(internedName);
    if (cached !== undefined) {
      return cached;
    }

    // Check if it's a known global key
    if (globalKeySet.has(internedName)) {
      for (const key of Object.keys(globalThis)) {
        if (globalStringInterner.intern(key) === internedName) {
          const value = (globalThis as Record<string, unknown>)[key];
          const vbValue = jsToVbOptimized(value, globalThis);
          globalThisCache.set(internedName, vbValue);
          return vbValue;
        }
      }
    }

    // Dynamic global lookup
    for (const key of Object.keys(globalThis)) {
      if (globalStringInterner.intern(key) === internedName) {
        const value = (globalThis as Record<string, unknown>)[key];
        const vbValue = jsToVbOptimized(value, globalThis);
        globalThisCache.set(internedName, vbValue);
        return vbValue;
      }
    }
    
    if (this.optionExplicit) {
      throw new VbError(500, `Variable is undefined: '${name}'`, 'Vbscript');
    }
    return VbEmpty;
  }

  /**
   * Optimized variable setting with string interning.
   */
  setVariable(name: string, value: VbValue): void {
    const internedName = globalStringInterner.intern(name);
    
    if (this.inPropertyGet && internedName === globalStringInterner.intern(this.propertyGetName)) {
      this.currentScope.set(name, value);
      return;
    }
    
    if (this.currentInstance && this.currentInstance.hasProperty(name)) {
      this.currentInstance.setProperty(name, value);
      return;
    }
    
    if (this.optionExplicit && !this.currentScope.has(name)) {
      if (!globalKeySet.has(internedName)) {
        throw new VbError(500, `Variable is undefined: '${name}'`, 'Vbscript');
      }
    }
    
    this.currentScope.set(name, value);
  }

  hasVariable(name: string): boolean {
    if (this.currentScope.has(name)) return true;
    const internedName = globalStringInterner.intern(name);
    return globalKeySet.has(internedName);
  }

  /**
   * Cleanup method to return all pooled scopes.
   * Call this when the context is no longer needed.
   */
  cleanup(): void {
    for (const scope of this.pooledScopes) {
      globalScopePool.release(scope);
    }
    this.pooledScopes.length = 0;
  }
}
