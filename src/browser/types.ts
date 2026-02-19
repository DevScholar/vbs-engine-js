export interface BrowserRuntimeOptions {
  parseScriptElement?: boolean;
  parseInlineEventAttributes?: boolean;
  injectGlobalThis?: boolean;
  parseEventSubNames?: boolean;
  maxExecutingTime?: number;
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
