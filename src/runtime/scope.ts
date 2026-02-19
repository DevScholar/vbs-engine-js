import type { VbValue } from './values.ts';

export class VbVariable {
  constructor(
    public name: string,
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

  declare(name: string, value: VbValue, options: Partial<VbVariable> = {}): VbVariable {
    const variable = new VbVariable(
      name,
      value,
      options.isByRef ?? false,
      options.isArray ?? false,
      options.isConst ?? false
    );
    this.variables.set(name.toLowerCase(), variable);
    return variable;
  }

  get(name: string): VbVariable | undefined {
    const lowerName = name.toLowerCase();
    const variable = this.variables.get(lowerName);
    if (variable) return variable;
    if (this.parent) return this.parent.get(name);
    return undefined;
  }

  set(name: string, value: VbValue): void {
    const lowerName = name.toLowerCase();
    const variable = this.variables.get(lowerName);
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
    const lowerName = name.toLowerCase();
    if (this.variables.has(lowerName)) return true;
    if (this.parent) return this.parent.has(name);
    return false;
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
}
