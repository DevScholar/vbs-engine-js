import { Lexer } from '../lexer/index.ts';
import { Parser } from '../parser/index.ts';
import { Interpreter } from '../interpreter/index.ts';
import type { VbValue } from '../runtime/index.ts';
import { jsToVb, vbToJs } from './conversion.ts';

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
 * // Basic usage (Node.js or general)
 * const engine = new VbsEngine();
 * engine.run('x = 1 + 2');
 * console.log(engine.getVariableAsJs('x')); // 3
 *
 * // Browser mode
 * const engine = new VbsEngine({ mode: 'browser' });
 *
 * // With options
 * const engine = new VbsEngine({ maxExecutionTime: 5000 });
 * ```
 */
export class VbsEngine {
  private interpreter: Interpreter;
  private options: Required<VbsEngineOptions>;
  private browserCleanup: (() => void) | null = null;

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

    // Initialize browser mode if requested
    if (this.options.mode === 'browser' && typeof window !== 'undefined') {
      this.initializeBrowserMode();
    }
  }

  private initializeBrowserMode(): void {
    // Dynamic import to avoid loading browser code in Node.js
    import('../browser/index.ts').then(({ initializeBrowserEngine }) => {
      this.browserCleanup = initializeBrowserEngine(this, this.options);
    }).catch(err => {
      console.error('Failed to initialize browser mode:', err);
    });
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
