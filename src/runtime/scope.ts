/**
 * Scope implementation with string interning for case-insensitive variable lookup.
 */

import type { VbValue } from './values.ts';
import { globalStringInterner } from './string-interner.ts';

export class VbVariable {
  constructor(
    public name: string, // Already interned (lowercased)
    public originalName: string, // Original casing for error messages
    public value: VbValue,
    public isByRef: boolean = false,
    public isArray: boolean = false,
    public isConst: boolean = false
  ) {}
}

export class Vbscope {
  private variables: Map<string, VbVariable> = new Map();
  public parent: Vbscope | null;

  constructor(parent: Vbscope | null = null) {
    this.parent = parent;
  }

  declare(
    name: string,
    value: VbValue,
    options: Partial<Omit<VbVariable, 'name' | 'originalName' | 'value'>> = {}
  ): VbVariable {
    const internedName = globalStringInterner.intern(name);
    const variable = new VbVariable(
      internedName,
      name,
      value,
      options.isByRef ?? false,
      options.isArray ?? false,
      options.isConst ?? false
    );
    this.variables.set(internedName, variable);
    return variable;
  }

  get(name: string): VbVariable | undefined {
    const internedName = globalStringInterner.intern(name);
    const variable = this.variables.get(internedName);
    if (variable) return variable;
    if (this.parent) return this.parent.get(name);
    return undefined;
  }

  set(name: string, value: VbValue): void {
    const internedName = globalStringInterner.intern(name);
    const variable = this.variables.get(internedName);
    if (variable) {
      if (variable.isConst) {
        throw new Error(`Cannot assign to constant '${name}'`);
      }
      variable.value = value;
      return;
    }
    if (this.parent) {
      this.parent.set(name, value);
      return;
    }
    this.declare(name, value);
  }

  has(name: string): boolean {
    const internedName = globalStringInterner.intern(name);
    if (this.variables.has(internedName)) return true;
    if (this.parent) return this.parent.has(name);
    return false;
  }

  hasLocal(name: string): boolean {
    const internedName = globalStringInterner.intern(name);
    return this.variables.has(internedName);
  }

  getParent(): Vbscope | null {
    return this.parent;
  }

  getAllVariables(): Map<string, VbVariable> {
    const all = new Map<string, VbVariable>();
    if (this.parent) {
      const parentVars = this.parent.getAllVariables();
      parentVars.forEach((v, k) => all.set(k, v));
    }
    this.variables.forEach((v, k) => all.set(k, v));
    return all;
  }

  getLocalVariables(): Map<string, VbVariable> {
    return new Map(this.variables);
  }

  clear(): void {
    this.variables.clear();
  }

  get size(): number {
    return this.variables.size;
  }
}

/**
 * Scope pool for reusing scope objects.
 * Reduces GC pressure during function calls.
 */
export class ScopePool {
  private pool: Vbscope[] = [];
  private maxSize: number;

  constructor(maxSize: number = 100) {
    this.maxSize = maxSize;
  }

  acquire(parent: Vbscope | null = null): Vbscope {
    const scope = this.pool.pop();
    if (scope) {
      scope.parent = parent;
      return scope;
    }
    return new Vbscope(parent);
  }

  release(scope: Vbscope): void {
    if (this.pool.length < this.maxSize) {
      scope.clear();
      scope.parent = null;
      this.pool.push(scope);
    }
  }

  clear(): void {
    this.pool.length = 0;
  }

  get size(): number {
    return this.pool.length;
  }
}

export const globalScopePool = new ScopePool();
