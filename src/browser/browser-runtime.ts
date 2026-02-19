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
    const engineOptions: VbsEngineOptions = {
      maxExecutionTime: options.maxExecutionTime ?? DEFAULT_ENGINE_OPTIONS.maxExecutionTime,
      injectGlobalThis: options.injectGlobalThis ?? DEFAULT_ENGINE_OPTIONS.injectGlobalThis,
    };

    this.engine = new VbsEngine(engineOptions);

    this.options = {
      ...engineOptions,
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

  run(code: string): unknown {
    return this.engine.run(code);
  }

  getEngine(): VbsEngine {
    return this.engine;
  }

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

export function createBrowserRuntime(options?: BrowserRuntimeOptions): VbsBrowserEngine {
  return new VbsBrowserEngine(options);
}

export { vbToJs, jsToVb } from './conversion.ts';
export type { BrowserRuntimeOptions } from './types.ts';
