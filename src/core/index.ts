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

/**
 * Configuration options for the VbsEngine.
 */
export interface VbsEngineOptions {
  /**
   * Maximum execution time in milliseconds.
   * Set to -1 for unlimited execution time (default).
   * Useful to prevent infinite loops from hanging the application.
   */
  maxExecutionTime?: number;
  /**
   * When true, VBScript functions and variables are automatically
   * injected into the global scope (globalThis), enabling IE-style
   * interoperability between VBScript and JavaScript.
   * @default false
   */
  injectGlobalThis?: boolean;
}

/**
 * A VBScript engine that can parse and execute VBScript code.
 *
 * This engine supports the full VBScript language including:
 * - Variables, constants, and arrays
 * - Control flow statements (If, For, Do, While, Select Case)
 * - Procedures (Sub and Function)
 * - Classes with properties and methods
 * - Error handling (On Error)
 * - Built-in functions (string, math, date, conversion, etc.)
 *
 * @example
 * ```typescript
 * // Basic usage
 * const engine = new VbsEngine();
 * engine.run('x = 1 + 2');
 * console.log(engine.getVariableAsJs('x')); // 3
 *
 * // With options
 * const engine = new VbsEngine({ maxExecutionTime: 5000 });
 * ```
 */
export class VbsEngine {
  private interpreter: Interpreter;
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

  /**
   * Sets the maximum execution time for script execution.
   * @param ms - Maximum time in milliseconds, or -1 for unlimited
   */
  setMaxExecutionTime(ms: number): void {
    this.interpreter.setMaxExecutionTime(ms);
    this.interpreter.getContext().checkTimeout = () => this.interpreter.checkTimeout();
  }

  /**
   * Executes VBScript source code.
   *
   * @param source - The VBScript code to execute
   * @returns The result of the last evaluated expression
   * @throws Error if the code contains syntax errors or runtime errors
   *
   * @example
   * ```typescript
   * engine.run('x = 5: y = 10: result = x + y');
   * ```
   */
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

  /**
   * Gets a variable value from the VBScript context.
   *
   * @param name - The variable name (case-insensitive)
   * @returns The VbValue representation of the variable
   */
  getVariable(name: string): VbValue {
    return this.interpreter.getVariable(name);
  }

  /**
   * Gets a variable value converted to a JavaScript type.
   *
   * @param name - The variable name (case-insensitive)
   * @returns The JavaScript value (string, number, boolean, object, etc.)
   *
   * @example
   * ```typescript
   * engine.run('name = "John"');
   * const name = engine.getVariableAsJs('name'); // "John" (string)
   * ```
   */
  getVariableAsJs(name: string): unknown {
    return vbToJsAuto(this.interpreter.getVariable(name));
  }

  /**
   * Sets a variable in the VBScript context.
   *
   * @param name - The variable name
   * @param value - The VbValue to set
   */
  setVariable(name: string, value: VbValue): void {
    this.interpreter.setVariable(name, value);
  }

  /**
   * Registers a custom function that can be called from VBScript code.
   *
   * @param name - The function name as it will appear in VBScript
   * @param func - The function implementation
   *
   * @example
   * ```typescript
   * engine.registerFunction('DoubleIt', (val) => ({
   *   type: 'Long',
   *   value: val.value * 2
   * }));
   * engine.run('result = DoubleIt(5)'); // result = 10
   * ```
   */
  registerFunction(name: string, func: (...args: VbValue[]) => VbValue): void {
    this.interpreter.registerFunction(name, func);
  }

  /**
   * Gets the internal execution context.
   * Use this for advanced scenarios requiring direct context manipulation.
   *
   * @returns The VbContext instance
   */
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

/**
 * A convenience function to quickly execute VBScript code.
 * Creates a new VbsEngine instance, runs the code, and returns the result.
 *
 * @param source - The VBScript code to execute
 * @returns The result of the last evaluated expression
 *
 * @example
 * ```typescript
 * const result = runVbscript('x = 5 + 3: x * 2');
 * // result.value === 16
 * ```
 */
export function runVbscript(source: string): VbValue {
  const engine = new VbsEngine();
  return engine.run(source);
}
