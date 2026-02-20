import type { VbsEngine, VbsEngineOptions } from '../core/index.ts';
import { createBrowserMsgBox } from '../builtins/msgbox.ts';
import { createBrowserInputBox } from '../builtins/inputbox.ts';
import type { TimerOverrideState, EvalOverrideState } from './types.ts';
import type { ProtocolHandlerResult } from './protocol-handler.ts';
import { overrideTimers, restoreTimers } from './timer-override.ts';
import { overrideEval, restoreEval } from './eval-override.ts';
import { setupVbscriptProtocol, cleanupVbscriptProtocol } from './protocol-handler.ts';
import { createObject, getObject } from './activex.ts';
import { autoRunScripts } from './script-parser.ts';
import { setupInlineEventHandlers, setupNamedEventHandlers, cleanupNamedEventHandlers, type BoundNamedHandler } from './event-handlers.ts';
import { startObserver, stopObserver, type ObserverContext } from './dom-observer.ts';

/**
 * Initializes the browser mode for a VbsEngine instance.
 * This function sets up all browser-specific features.
 *
 * @param engine - The VbsEngine instance to initialize
 * @param options - Browser-specific options
 * @returns A cleanup function to restore original browser state
 */
export function initializeBrowserEngine(
  engine: VbsEngine,
  options: Required<VbsEngineOptions>
): () => void {
  const timerState: TimerOverrideState = { originalSetTimeout: null, originalSetInterval: null };
  const evalState: EvalOverrideState = { originalEval: null };
  let protocolState: ProtocolHandlerResult = { navigateHandler: null, clickHandler: null };
  let observerContext: ObserverContext | null = null;
  const boundNamedHandlers: Map<string, BoundNamedHandler> = new Map();

  // Register browser-specific functions
  engine._registerFunction('MsgBox', createBrowserMsgBox());
  engine._registerFunction('InputBox', createBrowserInputBox());
  engine._registerFunction('CreateObject', createObject);
  engine._registerFunction('GetObject', getObject);

  // Override JS functions if requested
  if (options.overrideJSEvalFunctions) {
    timerState.originalSetTimeout = overrideTimers(engine).originalSetTimeout;
    timerState.originalSetInterval = overrideTimers(engine).originalSetInterval;
    evalState.originalEval = overrideEval(engine).originalEval;
  }

  // Setup vbscript: protocol handler
  if (options.parseVbsProtocol) {
    protocolState = setupVbscriptProtocol(engine);
  }

  // Parse and execute script elements
  if (options.parseScriptElement) {
    autoRunScripts(engine);
  }

  // Setup inline event handlers
  if (options.parseInlineEventAttributes) {
    setupInlineEventHandlers(engine);
  }

  // Setup named event handlers (Button1_OnClick style)
  if (options.injectGlobalThis && options.parseEventSubNames) {
    setupNamedEventHandlers(engine, boundNamedHandlers);
  }

  // Start DOM observer for dynamic content
  if (options.parseScriptElement || options.parseInlineEventAttributes) {
    observerContext = startObserver(engine, options, boundNamedHandlers);
  }

  // Handle DOMContentLoaded for late initialization
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => {
      if (options.injectGlobalThis && options.parseEventSubNames) {
        setupNamedEventHandlers(engine, boundNamedHandlers);
      }
    });
  }

  // Return cleanup function
  return () => {
    if (observerContext) {
      stopObserver(observerContext);
    }
    cleanupNamedEventHandlers(boundNamedHandlers);
    if (protocolState.navigateHandler || protocolState.clickHandler) {
      cleanupVbscriptProtocol(protocolState);
    }
    if (options.overrideJSEvalFunctions) {
      restoreTimers(timerState);
      restoreEval(evalState);
    }
  };
}
