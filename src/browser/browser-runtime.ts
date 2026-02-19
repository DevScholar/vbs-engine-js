import { VbsEngine, type VbsEngineOptions } from '../core/index.ts';
import { createBrowserMsgBox } from '../builtins/msgbox.ts';
import { createBrowserInputBox } from '../builtins/inputbox.ts';
import type { BrowserRuntimeOptions, TimerOverrideState, EvalOverrideState } from './types.ts';
import { overrideTimers, restoreTimers } from './timer-override.ts';
import { overrideEval, restoreEval } from './eval-override.ts';
import { setupVbscriptProtocol, cleanupVbscriptProtocol, type ProtocolHandlerResult } from './protocol-handler.ts';
import { createObject, getObject } from './activex.ts';
import { autoRunScripts, setupForEventScript } from './script-parser.ts';
import { setupInlineEventHandlers, setupNamedEventHandlers, cleanupNamedEventHandlers, type BoundNamedHandler } from './event-handlers.ts';
import { startObserver, stopObserver, type ObserverContext } from './dom-observer.ts';

const DEFAULT_BROWSER_OPTIONS = {
  parseScriptElement: true,
  parseInlineEventAttributes: true,
  parseEventSubNames: true,
  overrideJSEvalFunctions: true,
  parseVbsProtocol: true,
};

const DEFAULT_ENGINE_OPTIONS: VbsEngineOptions = {
  maxExecutionTime: -1,
  injectGlobalThis: true,
};

/**
 * A VBScript engine designed for browser environments.
 *
 * This engine extends VbsEngine with browser-specific features:
 * - Automatic parsing of `<script type="text/vbscript">` elements
 * - Inline event attribute handling (`onclick="vbscript:..."`)
 * - IE-style event binding via Sub names (`Button1_OnClick`)
 * - `vbscript:` protocol handler for links
 * - Integration with browser's `alert`, `confirm`, and `prompt` for MsgBox/InputBox
 *
 * @example
 * ```html
 * <!DOCTYPE html>
 * <html>
 * <head>
 *   <script type="module">
 *     import { VbsBrowserEngine } from './src/index.ts';
 *     new VbsBrowserEngine();
 *   </script>
 * </head>
 * <body>
 *   <script type="text/vbscript">
 *     MsgBox "Hello from VBScript!"
 *   </script>
 * </body>
 * </html>
 * ```
 *
 * @example
 * ```typescript
 * // With custom options
 * const engine = new VbsBrowserEngine({
 *   maxExecutionTime: 5000,
 *   parseScriptElement: true,
 *   parseInlineEventAttributes: true
 * });
 * ```
 */
export class VbsBrowserEngine {
  private engine: VbsEngine;
  private initialized: boolean = false;
  private timerState: TimerOverrideState = { originalSetTimeout: null, originalSetInterval: null };
  private evalState: EvalOverrideState = { originalEval: null };
  private protocolState: ProtocolHandlerResult = { navigateHandler: null, clickHandler: null };
  private observerContext: ObserverContext | null = null;
  private boundNamedHandlers: Map<string, BoundNamedHandler> = new Map();
  private options: Required<BrowserRuntimeOptions>;

  constructor(options: BrowserRuntimeOptions = {}) {
    const maxExecTime = options.maxExecutionTime ?? -1;
    const injectGlobal = options.injectGlobalThis ?? true;

    this.engine = new VbsEngine({ maxExecutionTime: maxExecTime, injectGlobalThis: injectGlobal });

    this.options = {
      maxExecutionTime: maxExecTime,
      injectGlobalThis: injectGlobal,
      parseScriptElement: options.parseScriptElement ?? DEFAULT_BROWSER_OPTIONS.parseScriptElement,
      parseInlineEventAttributes: options.parseInlineEventAttributes ?? DEFAULT_BROWSER_OPTIONS.parseInlineEventAttributes,
      parseEventSubNames: options.parseEventSubNames ?? DEFAULT_BROWSER_OPTIONS.parseEventSubNames,
      overrideJSEvalFunctions: options.overrideJSEvalFunctions ?? DEFAULT_BROWSER_OPTIONS.overrideJSEvalFunctions,
      parseVbsProtocol: options.parseVbsProtocol ?? DEFAULT_BROWSER_OPTIONS.parseVbsProtocol,
    };

    if (typeof window !== 'undefined') {
      if (this.options.overrideJSEvalFunctions) {
        this.timerState = overrideTimers(this.engine);
        this.evalState = overrideEval(this.engine);
      }
      this.registerBrowserFunctions();
      if (this.options.parseVbsProtocol) {
        this.protocolState = setupVbscriptProtocol(this.engine);
      }

      if (this.options.parseScriptElement) {
        this.autoRunScripts();
      }

      if (this.options.parseInlineEventAttributes) {
        setupInlineEventHandlers(this.engine);
      }

      if (this.options.injectGlobalThis && this.options.parseEventSubNames) {
        setupNamedEventHandlers(this.engine, this.boundNamedHandlers);
      }

      if (this.options.parseScriptElement || this.options.parseInlineEventAttributes) {
        this.observerContext = startObserver(this.engine, this.options, this.boundNamedHandlers);
      }

      if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', () => {
          if (this.options.injectGlobalThis && this.options.parseEventSubNames) {
            setupNamedEventHandlers(this.engine, this.boundNamedHandlers);
          }
        });
      }

      this.initialized = true;
    }
  }

  private registerBrowserFunctions(): void {
    this.engine.registerFunction('MsgBox', createBrowserMsgBox());
    this.engine.registerFunction('InputBox', createBrowserInputBox());
    this.engine.registerFunction('CreateObject', createObject);
    this.engine.registerFunction('GetObject', getObject);
  }

  private autoRunScripts(): void {
    autoRunScripts(this.engine);
  }

  /**
   * Executes VBScript code in the browser context.
   *
   * @param code - The VBScript code to execute
   * @returns The result of the execution
   */
  run(code: string): unknown {
    return this.engine.run(code);
  }

  /**
   * Gets the underlying VbsEngine instance.
   * Use this to access lower-level engine functionality.
   *
   * @returns The VbsEngine instance
   */
  getEngine(): VbsEngine {
    return this.engine;
  }

  /**
   * Cleans up all browser-specific overrides and event handlers.
   * Call this method when you no longer need the engine to restore
   * the original browser behavior and free resources.
   */
  destroy(): void {
    if (this.observerContext) {
      stopObserver(this.observerContext);
      this.observerContext = null;
    }

    cleanupNamedEventHandlers(this.boundNamedHandlers);

    if (this.protocolState.navigateHandler || this.protocolState.clickHandler) {
      cleanupVbscriptProtocol(this.protocolState);
    }

    if (this.options.overrideJSEvalFunctions) {
      restoreTimers(this.timerState);
      restoreEval(this.evalState);
    }

    this.initialized = false;
  }
}

/**
 * Creates and initializes a browser VBScript runtime.
 * This is a convenience function that creates a new VbsBrowserEngine instance.
 *
 * @param options - Optional configuration for the browser runtime
 * @returns A new VbsBrowserEngine instance
 *
 * @example
 * ```typescript
 * const runtime = createBrowserRuntime({
 *   maxExecutionTime: 5000
 * });
 * runtime.run('MsgBox "Hello!"');
 * ```
 */
export function createBrowserRuntime(options?: BrowserRuntimeOptions): VbsBrowserEngine {
  return new VbsBrowserEngine(options);
}

export { vbToJs, jsToVb } from './conversion.ts';
export type { BrowserRuntimeOptions } from './types.ts';
