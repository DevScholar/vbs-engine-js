import type { VbsEngineOptions } from '../core/index.ts';

export interface BrowserRuntimeOptions extends VbsEngineOptions {
  parseScriptElement?: boolean;
  parseInlineEventAttributes?: boolean;
  parseEventSubNames?: boolean;
  overrideJSEvalFunctions?: boolean;
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
