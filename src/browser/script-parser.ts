import type { VbsEngine } from '../core/index.ts';

export function isVbscriptElement(element: Element): boolean {
  const language = element.getAttribute('language');
  const type = element.getAttribute('type');

  if (language?.toLowerCase() === 'vbscript') return true;
  if (type?.toLowerCase() === 'text/vbscript') return true;
  if (type?.toLowerCase() === 'text/vbs') return true;
  if (type?.toLowerCase() === 'application/x-vbscript') return true;

  return false;
}

/**
 * Extract VBScript code from script element, handling HTML comment wrapping
 * Supports the legacy pattern: <!-- ... //-->
 */
function extractVbscriptCode(script: Element): string {
  let code = script.textContent ?? '';

  // Trim whitespace
  code = code.trim();

  // Handle HTML comment wrapping: <!-- ... //-->
  // This was used in old browsers (1998 era) to hide scripts from browsers that didn't support <script>
  if (code.startsWith('<!--')) {
    // Find the end of the comment
    // Support both //--> and -->
    const endIndex = code.indexOf('//-->');
    if (endIndex !== -1) {
      code = code.substring(4, endIndex).trim();
    } else if (code.endsWith('-->')) {
      code = code.substring(4, code.length - 3).trim();
    }
  }

  return code;
}

export function parseScriptElement(
  engine: VbsEngine,
  script: Element,
  onScriptRun?: () => void
): void {
  if (!isVbscriptElement(script)) return;

  const code = extractVbscriptCode(script);
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
