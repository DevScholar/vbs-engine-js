import type { VbsEngine } from '../core/index.ts';

export function isVbscriptElement(element: Element): boolean {
  const language = element.getAttribute('language');
  const type = element.getAttribute('type');

  if (language?.toLowerCase() === 'vbscript') return true;
  if (type?.toLowerCase() === 'text/vbscript') return true;
  if (type?.toLowerCase() === 'application/x-vbscript') return true;

  return false;
}

export function parseScriptElement(
  engine: VbsEngine,
  script: Element,
  onScriptRun?: () => void
): void {
  if (!isVbscriptElement(script)) return;

  const code = script.textContent ?? '';
  const forAttr = script.getAttribute('for');
  const eventAttr = script.getAttribute('event');

  if (forAttr && eventAttr) {
    setupForEventScript(engine, forAttr, eventAttr, code);
  } else {
    try {
      engine.addCode(code);
      onScriptRun?.();
    } catch (error) {
      console.error('Vbscript error:', error);
    }
  }
}

export function setupForEventScript(
  engine: VbsEngine,
  targetId: string,
  eventName: string,
  code: string
): void {
  const eventType = eventName.toLowerCase().replace(/^on/, '');

  const bindHandler = (): boolean => {
    const target = document.getElementById(targetId);
    if (target) {
      const handler = (): void => {
        try {
          engine.addCode(code);
        } catch (error) {
          console.error('VBScript event handler error:', error);
        }
      };
      target.addEventListener(eventType, handler);
      return true;
    }
    return false;
  };

  if (!bindHandler()) {
    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', () => {
        bindHandler();
      });
    } else {
      setTimeout(bindHandler, 0);
    }
  }
}

export function autoRunScripts(engine: VbsEngine, onScriptRun?: () => void): void {
  if (typeof document === 'undefined') return;

  const scripts = document.querySelectorAll('script');
  scripts.forEach(script => {
    parseScriptElement(engine, script, onScriptRun);
  });
}
