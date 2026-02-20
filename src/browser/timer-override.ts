import type { VbsEngine } from '../core/index.ts';
import type { OriginalSetTimeout, OriginalSetInterval, TimerOverrideState } from './types.ts';

export function overrideTimers(engine: VbsEngine): TimerOverrideState {
  if (typeof window === 'undefined') {
    return { originalSetTimeout: null, originalSetInterval: null };
  }

  const originalSetTimeout = window.setTimeout;
  const originalSetInterval = window.setInterval;

  (window as unknown as Record<string, unknown>).setTimeout = function(
    handler: unknown,
    delay?: unknown,
    languageOrArg?: unknown,
    ...args: unknown[]
  ): number {
    if (typeof handler === 'string') {
      const language = typeof languageOrArg === 'string' ? languageOrArg.toLowerCase() : null;
      const actualDelay = typeof delay === 'number' ? delay : 0;

      if (language === 'vbscript' || language === 'vbs') {
        return originalSetTimeout.call(window, () => {
          const funcRegistry = engine._getContext()?.functionRegistry;
          if (funcRegistry?.has(handler)) {
            engine.run(handler);
          } else {
            engine.executeStatement(handler);
          }
        }, actualDelay);
      }

      return originalSetTimeout.call(window, () => {
        const funcRegistry = engine._getContext()?.functionRegistry;
        if (funcRegistry?.has(handler)) {
          engine.run(handler);
        } else {
          engine.executeStatement(handler);
        }
      }, actualDelay);
    }

    return originalSetTimeout.call(window, handler as TimerHandler, delay as number | undefined, ...[languageOrArg, ...args].filter(a => a !== undefined));
  };

  (window as unknown as Record<string, unknown>).setInterval = function(
    handler: unknown,
    delay?: unknown,
    languageOrArg?: unknown,
    ...args: unknown[]
  ): number {
    if (typeof handler === 'string') {
      const language = typeof languageOrArg === 'string' ? languageOrArg.toLowerCase() : null;
      const actualDelay = typeof delay === 'number' ? delay : 0;

      if (language === 'vbscript' || language === 'vbs') {
        return originalSetInterval.call(window, () => {
          const funcRegistry = engine._getContext()?.functionRegistry;
          if (funcRegistry?.has(handler)) {
            engine.run(handler);
          } else {
            engine.executeStatement(handler);
          }
        }, actualDelay);
      }

      return originalSetInterval.call(window, () => {
        const funcRegistry = engine._getContext()?.functionRegistry;
        if (funcRegistry?.has(handler)) {
          engine.run(handler);
        } else {
          engine.executeStatement(handler);
        }
      }, actualDelay);
    }

    return originalSetInterval.call(window, handler as TimerHandler, delay as number | undefined, ...[languageOrArg, ...args].filter(a => a !== undefined));
  };

  return { originalSetTimeout, originalSetInterval };
}

export function restoreTimers(state: TimerOverrideState): void {
  if (typeof window === 'undefined') return;

  if (state.originalSetTimeout) {
    (window as unknown as Record<string, unknown>).setTimeout = state.originalSetTimeout;
  }
  if (state.originalSetInterval) {
    (window as unknown as Record<string, unknown>).setInterval = state.originalSetInterval;
  }
}
