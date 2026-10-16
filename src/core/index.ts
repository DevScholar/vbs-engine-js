import { Lexer } from '../lexer/index.ts';
import type { Token } from '../lexer/token.ts';
import { Parser, globalParserCache } from '../parser/index.ts';
import { Interpreter } from '../interpreter/index.ts';
import type { VbValue } from '../runtime/index.ts';
import { VbError, toHresult } from '../runtime/index.ts';
import { jsToVb, vbToJs } from './conversion.ts';
import { withLocaleState } from '../builtins/locale.ts';
import type { LocaleState } from '../builtins/locale.ts';
import { initializeBrowserEngine } from '../browser/index.ts';
import { VBArray } from '../runtime/vbarray-wrapper.ts';

export { jsToVb, vbToJs } from './conversion.ts';
export { VBArray } from '../runtime/vbarray-wrapper.ts';

// eslint-disable-next-line @typescript-eslint/no-unused-vars
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
 * - 'auto': Automatically detect environment (browser or general)
 */
export type VbsEngineMode = 'general' | 'browser' | 'auto';

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
  overrideJsEvalFunctions?: boolean;
  /**
   * Enable the `vbscript:` protocol handler for links and forms.
   * @default true
   */
  parseVbsProtocol?: boolean;
  /**
   * Inject the VBArray class into globalThis for compatibility with legacy code.
   * When true, `new VBArray(vbScriptArray)` can be used from JavaScript to wrap
   * VBScript arrays. This is primarily for backward compatibility with early IE.
   * Modern code can access VBScript arrays directly without wrapping.
   * @default true
   */
  injectVBArrayToGlobalThis?: boolean;
}

/**
 * Error information from script execution.
 *
 * This mirrors the MSScriptControl `ScriptControl.Error` object (KB Q184742):
 * a pure error code (not an HRESULT) plus the source string and position fields.
 * The IE/WSH-style HRESULT shape is a separate surface — see
 * {@link VbsEngine.executeStatementThrows}.
 */
export interface VbsError {
  /** Error number/code (pure VBScript code, e.g. 11 for division by zero) */
  number: number;
  /** Error description */
  description: string;
  /** Component that raised the error ("Microsoft VBScript runtime error" etc.) */
  source: string;
  /** Source line number where error occurred (1-based) */
  line?: number;
  /** Source column where error occurred */
  column?: number;
  /** Source text of the offending line (compile errors only) */
  text?: string;
  /** Help file name, populated only for user-raised errors */
  helpFile?: string;
  /** Help context id, populated only for user-raised errors */
  helpContext?: number;
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
   * Whether this engine integrates with the host it is running in, the way a script host does in IE or
   * WSH: it may own the document, install the `vbscript:` handler, replace the host's eval and timers,
   * and publish its functions and variables into the host's global object.
   *
   * A host has one of these, so the first engine constructed with it claims it for the process. Every
   * later engine is isolated instead, and cannot ask to integrate: with two engines both writing host
   * globals, binding the same event names and installing the same protocol handler, they fight, and one
   * engine's `SetLocale` was visible to the other. An isolated engine keeps reading host globals - that
   * is harmless - and keeps calling any object handed to it with `addObject`.
   *
   * Like the host it stands for, the choice is final: it is fixed when the engine is constructed.
   * @default true, unless an engine already claimed it in this process
   */
  hostIntegration?: boolean;
  /**
   * How an error surfaces when the *host* runs script that no JavaScript call asked for: an inline
   * event attribute, a timer handler, a script element. In IE the error reaches the page as a real
   * JavaScript error, which `window.onerror` and `catch` can see; in WSH the host prints it and sets an
   * exit code. This picks the first; `false` papers over it in `engine.error` instead.
   *
   * It does not touch `executeStatement`, which never throws, or `executeStatementThrows`, which always
   * does: when JavaScript asks for the script to run, the method that was called says how an error
   * surfaces, and both outlets stay available. `engine.error` is set either way, as a record.
   * @default true
   */
  useJsError?: boolean;
  /**
   * When true (default), JavaScript globals from globalThis are automatically
   * shared with VBScript, and VBScript functions/variables are injected
   * into globalThis for bidirectional interoperability.
   * @default true
   */
  injectGlobalThis?: boolean;
  /**
   * The mode of operation for the engine.
   * - 'general': Standard mode for Node.js or custom environments
   * - 'browser': Browser mode with automatic DOM integration
   * When not specified, automatically detects environment:
   * - 'browser' if running in a browser (window is defined)
   * - 'general' otherwise (Node.js or other environments)
   * @default 'auto'
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
/**
 * Determines the appropriate mode based on the environment.
 * If running in a browser (window exists), returns 'browser'.
 * Otherwise, returns 'general' for Node.js or other environments.
 */
function determineMode(): VbsEngineMode {
  if (typeof window !== 'undefined') {
    return 'browser';
  }
  return 'general';
}

/** A process has one host, the way IE has one window and WSH one process. */
let hostIntegrationClaimed = false;

/**
 * The first engine to ask for host integration gets it; every later engine is isolated and cannot ask
 * for it. An engine that opts out does not consume the claim, so it can still be taken by a later one.
 */
function resolveHostIntegration(requested: boolean | undefined): boolean {
  if (hostIntegrationClaimed) return false;
  if (requested === false) return false;
  hostIntegrationClaimed = true;
  return true;
}

function resolveMode(mode: VbsEngineMode | undefined): VbsEngineMode {
  if (mode === undefined || mode === 'auto') {
    return determineMode();
  }
  return mode;
}

/**
 * The native `Error` a VBScript failure reaches a JavaScript caller as, which mirrors what IE handed JScript:
 * `instanceof Error`, `name === 'Error'`, `message` equal to `description`, and `number` as an HRESULT.
 *
 * `line` and `column` say where the failing VBScript statement began. IE reported that position through the
 * arguments of `window.onerror` rather than on the error object; carrying it here lets a page report a script
 * error the way the Windows Script Host does, and lets a reporter take over the stack.
 */
export interface VbsHostError extends Error {
  /** The VBScript error code as an HRESULT, for example 0x800a000b for a division by zero. */
  number: number;
  description: string;
  line?: number;
  column?: number;
}

/**
 * Puts the Windows Script Host's own reading of a failure at the head of the stack: the position, the source and
 * the description, the way `cscript` prints them, as in `(4, 3) Microsoft VBScript runtime error: Division by
 * zero`. A host that knows the file it ran can put the name in front; this engine is handed statements and not
 * a file, so the position stands alone.
 *
 * Writing `error.stack` is how the format is taken over. It is an ordinary writable own property, so this works
 * in every browser and touches this one error only. V8 also offers `Error.prepareStackTrace`, but that is a hook
 * on the whole page: it would rewrite the stack of every other error the host reports, which is not a library's
 * to do. Firefox has no such hook and keeps the throw position in the non-standard `lineNumber` and
 * `columnNumber` properties, which cannot be written.
 */
function withWindowsScriptHostPosition(stack: string, err: VbError): string {
  if (err.line === undefined) return stack;
  const position = `(${err.line}, ${err.column ?? 0})`;
  const source = err.source || 'Microsoft VBScript runtime error';
  return `${position} ${source}: ${err.description}\n${stack}`;
}

export class VbsEngine {
  private interpreter: Interpreter;
  private options: Required<VbsEngineOptions>;
  private browserCleanup: (() => void) | null = null;
  /** Removes the hook this engine installed for the end of its host, or null when it installed none. */
  private hostEndCleanup: (() => void) | null = null;
  private lastError: VbsError | null = null;
  /** This engine's own locale, installed while its code runs (see withLocaleState). */
  private readonly localeState: LocaleState = { locale: null };

  constructor(options: VbsEngineOptions = {}) {
    // Host integration decides what the engine may own before anything else is resolved: the options
    // below only choose what to do with it.
    const hostIntegration = resolveHostIntegration(options.hostIntegration);
    const resolvedMode = resolveMode(options.mode);

    this.options = {
      maxExecutionTime: options.maxExecutionTime ?? -1,
      hostIntegration,
      useJsError: options.useJsError ?? true,
      // Kept on: an isolated engine still reads host globals. This option only decides whether the
      // engine publishes its own functions and variables, and that is the part that collides.
      injectGlobalThis: hostIntegration && (options.injectGlobalThis ?? true),
      mode: hostIntegration ? resolvedMode : 'general',
      parseScriptElement: hostIntegration && (options.parseScriptElement ?? true),
      parseInlineEventAttributes: hostIntegration && (options.parseInlineEventAttributes ?? true),
      parseEventSubNames: hostIntegration && (options.parseEventSubNames ?? true),
      overrideJsEvalFunctions: hostIntegration && (options.overrideJsEvalFunctions ?? true),
      parseVbsProtocol: hostIntegration && (options.parseVbsProtocol ?? true),
      injectVBArrayToGlobalThis: hostIntegration && (options.injectVBArrayToGlobalThis ?? true),
    };

    this.interpreter = new Interpreter();
    this.interpreter.getContext().evaluate = (code: string) => this.interpreter.evaluate(code);
    this.interpreter.getContext().execute = (code: string) =>
      this.interpreter.executeInCurrentScope(code);
    this.interpreter.getContext().executeGlobal = (code: string) =>
      this.interpreter.executeInGlobalScope(code);

    if (this.options.maxExecutionTime > 0) {
      this.setMaxExecutionTime(this.options.maxExecutionTime);
    }

    // Inject VBArray into globalThis for legacy compatibility (in all modes)
    if (this.options.injectVBArrayToGlobalThis && typeof globalThis !== 'undefined') {
      (globalThis as Record<string, unknown>).VBArray = VBArray;
    }

    if (this.options.mode === 'browser' && typeof window !== 'undefined') {
      this.initializeBrowserMode();
    }

    this.installHostEndHook();
  }

  private initializeBrowserMode(): void {
    this.browserCleanup = initializeBrowserEngine(this, this.options);
  }

  /**
   * Lets the script go when the host it stands for ends, which is what a script host does.
   *
   * Only the engine that claimed host integration installs this. That engine stands for the host, and a host
   * has an end: `cscript` runs out of script and exits, and an instance the script never dropped has its
   * `Class_Terminate` run at that moment. An isolated engine is a library the host drives, so it has no end of
   * its own and the host calls `releaseGlobals()` itself.
   *
   * The host's end is the whole process in Node and the page going away in a browser. A page that goes into the
   * back and forward cache is not ending, so it keeps its instances: `pagehide` says which of the two it is.
   */
  private installHostEndHook(): void {
    if (!this.options.hostIntegration) return;

    const release = () => {
      try {
        this.releaseGlobals();
      } catch {
        // The host is ending. A Terminate event that throws must not stop the host from ending, and there is
        // no one left to hand the error to.
      }
    };

    if (this.options.mode === 'browser' && typeof window !== 'undefined') {
      const onPageHide = (event: { persisted?: boolean }) => {
        if (event.persisted) return;
        release();
      };
      window.addEventListener('pagehide', onPageHide);
      this.hostEndCleanup = () => window.removeEventListener('pagehide', onPageHide);
      return;
    }

    // A listener on the host's own end does not keep the host alive, and every statement a Terminate event
    // runs is synchronous, which is all an exit handler may do.
    if (typeof process !== 'undefined' && typeof process.on === 'function') {
      process.on('exit', release);
      this.hostEndCleanup = () => process.off('exit', release);
    }
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
      withLocaleState(this.localeState, () => this.executeStatementInternal(statement));
    } catch (err) {
      this.handleError(err);
    }
  }

  /**
   * Executes a statement and rethrows any VBScript error as a native JavaScript
   * `Error`, mirroring the IE/WSH cross-language `catch(e)` surface rather than
   * the MSScriptControl-style `engine.error` member.
   *
   * The thrown object is a genuine `Error` instance where `name === 'Error'`,
   * `message === description`, and `number` is the HRESULT encoding
   * (`0x800A0000 | code`) of the VBScript error — the shape JScript receives
   * when it catches a VBScript runtime error.
   *
   * @example
   * ```typescript
   * try {
   *   engine.executeStatementThrows('x = 1/0');
   * } catch (e) {
   *   e.number; // -2146828277 (0x800A000B)
   *   e.message; // "Division by zero"
   * }
   * ```
   */
  executeStatementThrows(statement: string): void {
    this.clearError();
    try {
      withLocaleState(this.localeState, () => this.executeStatementInternal(statement));
    } catch (err) {
      this.handleError(err);
      throw this.toHostError(err);
    }
  }

  /**
   * Runs script the host asked for rather than a JavaScript caller: an inline event attribute, a timer
   * handler, a script element. Nobody chose one of the two error outlets, so `useJsError` does.
   */
  executeHostScript(code: string): void {
    if (this.options.useJsError) {
      this.executeStatementThrows(code);
    } else {
      this.executeStatement(code);
    }
  }

  private executeStatementInternal(statement: string): void {
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
      const result = withLocaleState(this.localeState, () =>
        funcRegistry.call(procedureName, vbArgs)
      );

      return vbToJs(result);
    } catch (err) {
      this.handleError(err);
      return undefined;
    }
  }

  /**
   * Ends a script, which runs the Terminate event of every class instance the script still holds.
   *
   * The engine that claimed host integration does this by itself when its host ends: the process in Node, the
   * page going away in a browser. A host that runs several scripts through one engine calls this between them,
   * since the engine cannot tell where one script stopped and the next began.
   *
   * @example
   * ```typescript
   * engine.addCode('Class Counter\nPrivate Sub Class_Terminate()\nWScript.Echo "gone"\nEnd Sub\nEnd Class');
   * engine.executeStatement('Set held = New Counter');
   * engine.releaseGlobals();  // runs Class_Terminate, so "gone" is written
   * ```
   */
  releaseGlobals(): void {
    this.interpreter.releaseGlobals();
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
      const result = withLocaleState(this.localeState, () => this.interpreter.evaluate(expression));
      return vbToJs(result);
    } catch (err) {
      this.handleError(err);
      return undefined;
    }
  }

  /**
   * Cleans up resources and restores any overridden browser functions.
   * Call this method when you no longer need the engine.
   *
   * It also takes back the hook this engine installed for the end of its host, so a destroyed engine has no
   * Terminate events left to run. It does not release the script's own globals: that is `releaseGlobals()`, and
   * a host that wants it before dropping the engine calls it first.
   */
  destroy(): void {
    if (this.browserCleanup) {
      this.browserCleanup();
      this.browserCleanup = null;
    }
    if (this.hostEndCleanup) {
      this.hostEndCleanup();
      this.hostEndCleanup = null;
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
    if (!(err instanceof Error)) {
      this.lastError = { number: -1, description: String(err), source: '' };
      return;
    }

    // VbError carries the real VBScript code, source, and position. A plain parse
    // Error carries only the token it stopped on. Either way the fields below map
    // onto the MSScriptControl-style VbsError surface.
    const raised = err as Error & {
      number?: unknown;
      source?: unknown;
      helpFile?: unknown;
      helpContext?: unknown;
      token?: Token;
      line?: unknown;
      column?: unknown;
    };
    const number = typeof raised.number === 'number' ? raised.number : -1;
    const source = typeof raised.source === 'string' ? raised.source : '';

    // A parse error carries the token it stopped on; a runtime error was stamped with its
    // statement's position on the way out. Either way VbsError's declared line and column can
    // finally be filled instead of staying undefined.
    const at =
      raised.token?.loc?.start ??
      (typeof raised.line === 'number'
        ? { line: raised.line, column: typeof raised.column === 'number' ? raised.column : 0 }
        : undefined);

    this.lastError = {
      number,
      description: err.message,
      source,
      ...(at ? { line: at.line, column: at.column } : {}),
      ...(typeof raised.helpFile === 'string' ? { helpFile: raised.helpFile } : {}),
      ...(typeof raised.helpContext === 'number' ? { helpContext: raised.helpContext } : {}),
    };
  }

  /**
   * Convert a caught internal error into the IE/WSH-style native `Error` that a
   * JavaScript host would see in its own `catch`. Mirrors JScript catching a
   * VBScript runtime error: `instanceof Error`, `name === 'Error'`,
   * `message === description`, and `number` as an HRESULT.
   *
   * The position of the failing VBScript statement rides along on `line` and `column`, which IE reported
   * through the arguments of `window.onerror` rather than on the error object, and which `engine.error`
   * carries under the same names.
   */
  private toHostError(err: unknown): VbsHostError {
    if (err instanceof VbError) {
      const hostError = new Error(err.description) as VbsHostError;
      hostError.name = 'Error';
      hostError.number = toHresult(err.number);
      hostError.description = err.description;
      hostError.line = err.line;
      hostError.column = err.column;
      hostError.stack = withWindowsScriptHostPosition(String(hostError.stack), err);
      return hostError;
    }
    const hostError = new Error(err instanceof Error ? err.message : String(err)) as VbsHostError;
    hostError.name = 'Error';
    hostError.number = toHresult(-1);
    hostError.description = hostError.message;
    return hostError;
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
      for (const [, vbVar] of allVars) {
        // `vbVar.name` is interned to lower case for the case insensitive lookup inside the engine, and
        // `originalName` is the name the script wrote. The host gets the name the script wrote, since
        // JavaScript is case sensitive: publishing the interned name put `vbsmessage` on the page where its
        // own JavaScript read `vbsMessage`, so every read of a VBScript variable from the page came back
        // undefined. The lookup below interns either way, so it takes the same name.
        const publishedName = vbVar.originalName;
        if (publishedName in (globalThis as Record<string, unknown>)) continue;

        try {
          Object.defineProperty(globalThis, publishedName, {
            get() {
              const variable = context.globalScope.get(publishedName);
              if (variable) {
                return vbToJs(variable.value);
              }
              return undefined;
            },
            set(value: unknown) {
              context.globalScope.set(publishedName, jsToVb(value));
            },
            configurable: true,
            enumerable: true,
          });
        } catch {
          // Fallback: static copy for non-configurable properties
          if (vbVar.value && vbVar.value.type !== 'Empty') {
            (globalThis as Record<string, unknown>)[publishedName] = vbToJs(vbVar.value);
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
