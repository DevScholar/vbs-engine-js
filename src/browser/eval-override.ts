import type { VbsEngine } from '../core/index.ts';
// eslint-disable-next-line @typescript-eslint/no-unused-vars
import type { OriginalEval, EvalOverrideState } from './types.ts';

export function overrideEval(engine: VbsEngine): EvalOverrideState {
  if (typeof window === 'undefined') {
    return { originalEval: null };
  }

  const originalEval = window.eval;

  (window as unknown as Record<string, unknown>).vbsEval = function(code: unknown): unknown {
    return engine.eval(String(code));
  };

  (window as unknown as Record<string, unknown>).eval = function(code: unknown, language?: unknown): unknown {
    if (typeof language === 'string') {
      const lang = language.toLowerCase();
      if (lang === 'vbscript' || lang === 'vbs') {
        return engine.eval(String(code));
      }
    }
    return originalEval.call(window, String(code));
  };

  return { originalEval };
}

export function restoreEval(state: EvalOverrideState): void {
  if (typeof window === 'undefined') return;

  if (state.originalEval) {
    (window as unknown as Record<string, unknown>).eval = state.originalEval;
  }
  delete (window as unknown as Record<string, unknown>).vbsEval;
}
