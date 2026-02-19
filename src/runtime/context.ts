import type { VbValue, VbObjectValueData } from './values.ts';
import { VbEmpty } from './values.ts';
import { Vbscope } from './scope.ts';
import { VbFunctionRegistry } from './function-registry.ts';
import { VbClassRegistry, VbObjectInstance } from './class-registry.ts';
import { VbError } from './errors.ts';

function jsToVb(value: unknown, thisArg?: unknown): VbValue {
  if (value === undefined) return { type: 'Empty', value: undefined };
  if (value === null) return { type: 'Null', value: null };
  if (typeof value === 'boolean') return { type: 'Boolean', value };
  if (typeof value === 'number') {
    if (Number.isInteger(value) && value >= -2147483648 && value <= 2147483647) {
      return { type: 'Long', value };
    }
    return { type: 'Double', value };
  }
  if (typeof value === 'string') return { type: 'String', value };
  if (value instanceof Date) return { type: 'Date', value };
  if (Array.isArray(value)) {
    return { type: 'Array', value };
  }
  if (typeof value === 'function') {
    return { type: 'Object', value: { type: 'jsfunction', func: value as (...args: unknown[]) => unknown, thisArg: thisArg ?? null } };
  }
  if (typeof value === 'object') {
    return { type: 'Object', value: value as VbObjectValueData };
  }
  return { type: 'String', value: String(value) };
}

export class VbContext {
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

  constructor() {
    this.globalScope = new Vbscope();
    this.currentScope = this.globalScope;
    this.functionRegistry = new VbFunctionRegistry();
    this.classRegistry = new VbClassRegistry();
  }

  pushScope(): Vbscope {
    const scope = new Vbscope(this.currentScope);
    this.currentScope = scope;
    return scope;
  }

  popScope(): void {
    if (this.currentScope.parent) {
      this.currentScope = this.currentScope.parent;
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
    this.callStack.push(name);
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

  getVariable(name: string): VbValue {
    if (this.currentInstance) {
      if (this.currentInstance.hasProperty(name)) {
        return this.currentInstance.getProperty(name);
      }
    }
    
    const variable = this.currentScope.get(name);
    if (variable) {
      return variable.value;
    }

    const lowerName = name.toLowerCase();
    for (const key of Object.keys(globalThis)) {
      if (key.toLowerCase() === lowerName) {
        const value = (globalThis as Record<string, unknown>)[key];
        return jsToVb(value, globalThis);
      }
    }

    const globalKeys = ['eval', 'parseInt', 'parseFloat', 'isNaN', 'isFinite', 'decodeURI', 'decodeURIComponent', 'encodeURI', 'encodeURIComponent', 'escape', 'unescape'];
    for (const key of globalKeys) {
      if (key.toLowerCase() === lowerName && key in globalThis) {
        const value = (globalThis as Record<string, unknown>)[key];
        return jsToVb(value, globalThis);
      }
    }
    
    if (this.optionExplicit) {
      throw new VbError(500, `Variable is undefined: '${name}'`, 'Vbscript');
    }
    return VbEmpty;
  }

  setVariable(name: string, value: VbValue): void {
    if (this.inPropertyGet && name.toLowerCase() === this.propertyGetName) {
      this.currentScope.set(name, value);
      return;
    }
    
    if (this.currentInstance && this.currentInstance.hasProperty(name)) {
      this.currentInstance.setProperty(name, value);
      return;
    }
    
    if (this.optionExplicit && !this.currentScope.has(name)) {
      const lowerName = name.toLowerCase();
      if (!(lowerName in globalThis)) {
        throw new VbError(500, `Variable is undefined: '${name}'`, 'Vbscript');
      }
    }
    
    this.currentScope.set(name, value);
  }

  hasVariable(name: string): boolean {
    if (this.currentScope.has(name)) return true;
    const lowerName = name.toLowerCase();
    return lowerName in globalThis;
  }
}
