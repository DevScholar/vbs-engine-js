import type { VbsEngineOptions } from '../core/index.ts';

/**
 * Configuration options for the browser runtime.
 * Extends VbsEngineOptions with browser-specific features.
 */
export interface BrowserRuntimeOptions extends VbsEngineOptions {
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
}

export type OriginalSetTimeout = typeof setTimeout;
export type OriginalSetInterval = typeof setInterval;
export type OriginalEval = typeof eval;

export interface TimerOverrideState {
  originalSetTimeout: OriginalSetTimeout | null;
  originalSetInterval: OriginalSetInterval | null;
}

export interface EvalOverrideState {
  originalEval: OriginalEval | null;
}
