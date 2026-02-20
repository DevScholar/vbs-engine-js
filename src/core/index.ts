import { Lexer } from '../lexer/index.ts';
import { Parser, globalParserCache } from '../parser/index.ts';
import { Interpreter } from '../interpreter/index.ts';
import type { VbValue } from '../runtime/index.ts';
import { jsToVb, vbToJs } from './conversion.ts';
import { initializeBrowserEngine } from '../browser/index.ts';

export { jsToVb, vbToJs } from './conversion.ts';

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
 * The mode of operation for the VbsEngine.
 * - 'general': Standard mode for Node.js or custom environments
 * - 'browser': Browser mode with automatic DOM integration, event handling, etc.
 */
export type VbsEngineMode = 'general' | 'browser';

/**
 * Browser-specific options for VbsEngine.
 */
export interface BrowserEngineOptions {
  /**
   * Automatically parse and execute `<script type="text/vbscript">` elements.
   * @default true
   */
  parseScriptElement?: boolean;
  /**
   * Automatically parse inline event attributes like `onclick="vbscript:..."`.
   * @default true
   */
  parseInlineEventAttributes?: boolean;
  /**
   * Automatically bind event handlers from Sub names like `Button1_OnClick`.
   * Requires `injectGlobalThis` to be true.
   * @default true
   */
  parseEventSubNames?: boolean;
  /**
   * Override JavaScript's eval, setTimeout, and setInterval to support VBScript code.
   * @default true
   */
  overrideJSEvalFunctions?: boolean;
  /**
   * Enable the `vbscript:` protocol handler for links and forms.
   * @default true
   */
  parseVbsProtocol?: boolean;
}

/**
 * Error information from script execution.
 */
export interface VbsError {
  /** Error number/code */
  number: number;
  /** Error description */
  description: string;
  /** Source line number where error occurred */
  line?: number;
  /** Source column where error occurred */
  column?: number;
  /** Source text where error occurred */
  text?: string;
}

/**
 * Configuration options for the VbsEngine.
 */
export interface VbsEngineOptions extends BrowserEngineOptions {
  /**
   * Maximum execution time in milliseconds.
   * Set to -1 for unlimited execution time (default).
   * Useful to prevent infinite loops from hanging the application.
   */
  maxExecutionTime?: number;
  /**
   * When true (default), JavaScript globals from globalThis are automatically
   * shared with VBScript, and VBScript functions/variables are injected
   * into globalThis for bidirectional interoperability.
   * @default true
   */
  injectGlobalThis?: boolean;
  /**
   * The mode of operation for the engine.
   * - 'general': Standard mode for Node.js or custom environments (default)
   * - 'browser': Browser mode with automatic DOM integration
   * @default 'general'
   */
  mode?: VbsEngineMode;
}

/**
 * A VBScript engine that can parse and execute VBScript code.
 *
 * This API is designed to be compatible with MSScriptControl.ScriptControl:
 * - `addCode(code)` - Add script code (functions, classes)
 * - `executeStatement(statement)` - Execute a single statement
 * - `run(procedureName, ...args)` - Call a script function
 * - `addObject(name, object, addMembers?)` - Expose a JS object to script
 * - `eval(expression)` - Evaluate an expression
 *
 * @example
 * ```typescript
 * // Basic usage
 * const engine = new VbsEngine();
 *
 * // Add code (function definitions)
 * engine.addCode(`
 *   Function Add(a, b)
 *     Add = a + b
 *   End Function
 * `);
 *
 * // Call the function
 * const result = engine.run('Add', 5, 3);  // 8
 *
 * // Execute a statement
 * engine.executeStatement('MsgBox "Hello"');
 *
 * // Evaluate an expression
 * const value = engine.eval('2 + 3 * 4');  // 14
 *
 * // Expose an object to script
 * engine.addObject('console', console, true);
 * engine.executeStatement('console.log "Hello from VBScript"');
 * ```
 */
export class VbsEngine {
  private interpreter: Interpreter;
  private options: Required<VbsEngineOptions>;
  private browserCleanup: (() => void) | null = null;
  private lastError: VbsError | null = null;

  constructor(options: VbsEngineOptions = {}) {
    this.options = {
      maxExecutionTime: options.maxExecutionTime ?? -1,
      injectGlobalThis: options.injectGlobalThis ?? true,
      mode: options.mode ?? 'general',
      parseScriptElement: options.parseScriptElement ?? true,
      parseInlineEventAttributes: options.parseInlineEventAttributes ?? true,
      parseEventSubNames: options.parseEventSubNames ?? true,
      overrideJSEvalFunctions: options.overrideJSEvalFunctions ?? true,
      parseVbsProtocol: options.parseVbsProtocol ?? true,
    };

    this.interpreter = new Interpreter();
    this.interpreter.getContext().evaluate = (code: string) => this.interpreter.evaluate(code);
    this.interpreter.getContext().execute = (code: string) => this.interpreter.executeInCurrentScope(code);
    this.interpreter.getContext().executeGlobal = (code: string) => this.interpreter.executeInGlobalScope(code);

    if (this.options.maxExecutionTime > 0) {
      this.setMaxExecutionTime(this.options.maxExecutionTime);
    }

    if (this.options.mode === 'browser' && typeof window !== 'undefined') {
      this.initializeBrowserMode();
    }
  }

  private initializeBrowserMode(): void {
    this.browserCleanup = initializeBrowserEngine(this, this.options);
  }

  private setMaxExecutionTime(ms: number): void {
    this.interpreter.setMaxExecutionTime(ms);
    this.interpreter.getContext().checkTimeout = () => this.interpreter.checkTimeout();
  }

  /**
   * Gets the last error that occurred during script execution.
   * Returns null if no error occurred.
   */
  get error(): VbsError | null {
    return this.lastError;
  }

  /**
   * Clears the last error.
   */
  clearError(): void {
    this.lastError = null;
  }

  /**
   * Adds script code to the engine.
   * Use this to define functions, classes, or variables that can be used later.
   *
   * @param code - The VBScript code to add (function/class definitions)
   *
   * @example
   * ```typescript
   * engine.addCode(`
   *   Function Multiply(a, b)
   *     Multiply = a * b
   *   End Function
   *
   *   Class Calculator
   *     Public Function Add(x, y)
   *       Add = x + y
   *     End Function
   *   End Class
   * `);
   * ```
   */
  addCode(code: string): void {
    this.clearError();
    try {
      // Try to get from cache first
      let program = globalParserCache.get(code);
      
      if (!program) {
        // Parse and cache
        const lexer = new Lexer(code);
        const tokens = lexer.tokenize();
        const parser = new Parser(tokens);
        program = parser.parse();
        globalParserCache.set(code, program);
      }
      
      this.interpreter.run(program);
      this.syncFunctionsToGlobalThis();
    } catch (err) {
      this.handleError(err);
    }
  }

  /**
   * Executes a single VBScript statement.
   *
   * @param statement - The statement to execute
   *
   * @example
   * ```typescript
   * engine.executeStatement('x = 10');
   * engine.executeStatement('MsgBox "Hello World"');
   * ```
   */
  executeStatement(statement: string): void {
    this.clearError();
    try {
      // Try to get from cache first
      let program = globalParserCache.get(statement);
      
      if (!program) {
        // Parse and cache
        const lexer = new Lexer(statement);
        const tokens = lexer.tokenize();
        const parser = new Parser(tokens);
        program = parser.parse();
        globalParserCache.set(statement, program);
      }
      
      this.interpreter.run(program);
    } catch (err) {
      this.handleError(err);
    }
  }

  /**
   * Calls a function defined in the script and returns the result.
   *
   * @param procedureName - The name of the function or sub to call
   * @param args - Arguments to pass to the function
   * @returns The return value of the function (undefined for Subs)
   *
   * @example
   * ```typescript
   * engine.addCode('Function Add(a, b): Add = a + b: End Function');
   * const result = engine.run('Add', 5, 3);  // 8
   * ```
   */
  run(procedureName: string, ...args: unknown[]): unknown {
    this.clearError();
    try {
      const context = this.interpreter.getContext();
      const funcRegistry = context.functionRegistry;
      
      const vbArgs = args.map(arg => jsToVb(arg));
      const result = funcRegistry.call(procedureName, vbArgs);
      
      return vbToJs(result);
    } catch (err) {
      this.handleError(err);
      return undefined;
    }
  }

  /**
   * Exposes a JavaScript object to the VBScript context.
   * The object can then be accessed by name in VBScript code.
   *
   * @param name - The name by which the object will be known in VBScript
   * @param object - The JavaScript object to expose
   * @param addMembers - If true, the object's methods/properties can be called directly
   *
   * @example
   * ```typescript
   * // Expose console to VBScript
   * engine.addObject('console', console, true);
   * engine.executeStatement('console.log "Hello from VBScript"');
   *
   * // Expose a custom object
   * const myApp = {
   *   name: 'MyApp',
   *   getVersion: () => '1.0.0',
   *   doSomething: (x) => x * 2
   * };
   * engine.addObject('MyApp', myApp, true);
   * engine.executeStatement('result = MyApp.doSomething(5)');  // result = 10
   * ```
   */
  addObject(name: string, object: unknown, addMembers: boolean = true): void {
    this.clearError();
    try {
      const vbValue = jsToVb(object);
      this.interpreter.setVariable(name, vbValue);
      
      if (addMembers && typeof object === 'object' && object !== null) {
        const obj = object as Record<string, unknown>;
        for (const key of Object.keys(obj)) {
          const memberName = `${name}.${key}`;
          const member = obj[key];
          if (typeof member === 'function') {
            this.interpreter.registerFunction(memberName, (...args: VbValue[]) => {
              const jsArgs = args.map(vbToJs);
              const result = (member as (...args: unknown[]) => unknown)(...jsArgs);
              return jsToVb(result);
            });
          }
        }
      }
    } catch (err) {
      this.handleError(err);
    }
  }

  /**
   * Evaluates a VBScript expression and returns the result.
   *
   * @param expression - The expression to evaluate
   * @returns The result of the expression
   *
   * @example
   * ```typescript
   * const result = engine.eval('2 + 3 * 4');  // 14
   * const greeting = engine.eval('"Hello" & " " & "World"');  // "Hello World"
   * ```
   */
  eval(expression: string): unknown {
    this.clearError();
    try {
      const result = this.interpreter.evaluate(expression);
      return vbToJs(result);
    } catch (err) {
      this.handleError(err);
      return undefined;
    }
  }

  /**
   * Cleans up resources and restores any overridden browser functions.
   * Call this method when you no longer need the engine.
   */
  destroy(): void {
    if (this.browserCleanup) {
      this.browserCleanup();
      this.browserCleanup = null;
    }
  }

  /**
   * @internal Registers a built-in function. Used by browser module.
   */
  _registerFunction(name: string, func: (...args: VbValue[]) => VbValue): void {
    this.interpreter.registerFunction(name, func);
  }

  /**
   * @internal Gets the interpreter context. Used by browser module.
   */
  _getContext() {
    return this.interpreter.getContext();
  }

  /**
   * @internal Gets a variable value. Used for testing.
   */
  _getVariable(name: string): VbValue {
    return this.interpreter.getVariable(name);
  }

  private handleError(err: unknown): void {
    if (err instanceof Error) {
      this.lastError = {
        number: -1,
        description: err.message,
      };
    } else {
      this.lastError = {
        number: -1,
        description: String(err),
      };
    }
  }

  private syncFunctionsToGlobalThis(): void {
    if (typeof globalThis === 'undefined') return;
    if (!this.options.injectGlobalThis) return;

    const context = this.interpreter.getContext();
    if (!context) return;

    const funcRegistry = context.functionRegistry;
    if (!funcRegistry) return;

    const userFuncs = funcRegistry.getUserDefinedFunctions?.();
    if (!userFuncs) return;

    for (const [, info] of userFuncs) {
      const funcName = info.name;
      if (!(funcName in (globalThis as Record<string, unknown>))) {
        (globalThis as Record<string, unknown>)[funcName] = (...args: unknown[]) => {
          return this.run(funcName, ...args);
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
 * A convenience function to quickly evaluate a VBScript expression.
 * Creates a new VbsEngine instance, evaluates the expression, and returns the result.
 *
 * @param expression - The VBScript expression to evaluate
 * @returns The result of the expression
 *
 * @example
 * ```typescript
 * const result = evalVbscript('2 + 3 * 4');  // 14
 * ```
 */
export function evalVbscript(expression: string): unknown {
  const engine = new VbsEngine();
  return engine.eval(expression);
}
