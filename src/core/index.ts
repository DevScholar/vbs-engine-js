import { Lexer } from '../lexer/index.ts';
import { Parser } from '../parser/index.ts';
import { Interpreter } from '../interpreter/index.ts';
import type { VbValue } from '../runtime/index.ts';
import { jsToVb, vbToJs } from './conversion.ts';

function vbToJsAuto(value: VbValue): unknown {
  switch (value.type) {
    case 'Empty':
      return undefined;
    case 'Null':
      return null;
    case 'Boolean':
    case 'Long':
    case 'Double':
    case 'Integer':
    case 'String':
    case 'Date':
      return value.value;
    case 'Array': {
      const arr = value.value as { toArray(): VbValue[] };
      return arr.toArray().map(vbToJsAuto);
    }
    case 'Object':
      return value.value;
    default:
      return value.value;
  }
}

export interface VbsEngineOptions {
  maxExecutionTime?: number;
  injectGlobalThis?: boolean;
}

export class VbsEngine {
  private interpreter: Interpreter;
  private maxExecutionTime: number = -1;
  private options: Required<VbsEngineOptions>;

  constructor(options: VbsEngineOptions = {}) {
    this.options = {
      maxExecutionTime: options.maxExecutionTime ?? -1,
      injectGlobalThis: options.injectGlobalThis ?? false,
    };

    this.interpreter = new Interpreter();
    this.interpreter.getContext().evaluate = (code: string) => this.interpreter.evaluate(code);
    this.interpreter.getContext().execute = (code: string) => this.interpreter.executeInCurrentScope(code);
    this.interpreter.getContext().executeGlobal = (code: string) => this.interpreter.executeInGlobalScope(code);

    if (this.options.maxExecutionTime > 0) {
      this.setMaxExecutionTime(this.options.maxExecutionTime);
    }
  }

  setMaxExecutionTime(ms: number): void {
    this.maxExecutionTime = ms;
    this.interpreter.setMaxExecutionTime(ms);
    this.interpreter.getContext().checkTimeout = () => this.interpreter.checkTimeout();
  }

  run(source: string): VbValue {
    const lexer = new Lexer(source);
    const tokens = lexer.tokenize();
    const parser = new Parser(tokens);
    const program = parser.parse();
    const result = this.interpreter.run(program);
    
    if (this.options.injectGlobalThis) {
      this.syncFunctionsToGlobalThis();
    }
    
    return result;
  }

  getVariable(name: string): VbValue {
    return this.interpreter.getVariable(name);
  }

  getVariableAsJs(name: string): unknown {
    return vbToJsAuto(this.interpreter.getVariable(name));
  }

  setVariable(name: string, value: VbValue): void {
    this.interpreter.setVariable(name, value);
  }

  registerFunction(name: string, func: (...args: VbValue[]) => VbValue): void {
    this.interpreter.registerFunction(name, func);
  }

  getContext() {
    return this.interpreter.getContext();
  }

  private syncFunctionsToGlobalThis(): void {
    if (typeof globalThis === 'undefined') return;

    const context = this.getContext();
    if (!context) return;

    const funcRegistry = context.functionRegistry;
    if (!funcRegistry) return;

    const userFuncs = funcRegistry.getUserDefinedFunctions?.();
    if (!userFuncs) return;

    for (const [, info] of userFuncs) {
      const funcName = info.name;
      if (!(funcName in (globalThis as Record<string, unknown>))) {
        (globalThis as Record<string, unknown>)[funcName] = (...args: unknown[]) => {
          const vbArgs = args.map(a => jsToVb(a));
          return vbToJs(funcRegistry.call(funcName, vbArgs));
        };
      }
    }

    if (context.globalScope) {
      const allVars = context.globalScope.getAllVariables();
      for (const [varName, vbVar] of allVars) {
        if (vbVar.value && vbVar.value.type !== 'Empty') {
          if (!(varName in (globalThis as Record<string, unknown>))) {
            (globalThis as Record<string, unknown>)[varName] = vbToJs(vbVar.value);
          }
        }
      }
    }
  }
}

export function runVbscript(source: string): VbValue {
  const engine = new VbsEngine();
  return engine.run(source);
}
