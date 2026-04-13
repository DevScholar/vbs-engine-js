import type { VbsEngine } from '../core/index.ts';

export interface ProtocolHandlerResult {
  navigateHandler: ((event: Event) => void) | null;
  clickHandler: ((event: MouseEvent) => void) | null;
}

function decodeVbscriptCode(rawCode: string): string {
  try {
    return decodeURIComponent(rawCode);
  } catch {
    return rawCode;
  }
}

function writeVbscriptResult(result: unknown): void {
  if (typeof document === 'undefined') return;
  if (result === undefined) return;

  document.open();
  document.write(String(result));
  document.close();
}

function executeVbscriptUrl(engine: VbsEngine, url: string): void {
  const code = decodeVbscriptCode(url.substring(9));

  // IE evaluates vbscript: URLs as expressions.
  // If the expression yields a non-Empty result, IE calls document.write(result),
  // replacing the page.  Void results (Sub calls, Empty) leave the page unchanged.
  const result = engine.eval(code);
  writeVbscriptResult(result);
}

export function setupVbscriptProtocol(engine: VbsEngine): ProtocolHandlerResult {
  if (typeof window === 'undefined') {
    return { navigateHandler: null, clickHandler: null };
  }

  let navigateHandler: ((event: Event) => void) | null = null;
  let clickHandler: ((event: MouseEvent) => void) | null = null;

  if ('navigation' in window && window.navigation) {
    navigateHandler = (event: Event): void => {
      const navEvent = event as { destination?: { url?: string } };
      const url = navEvent.destination?.url;
      if (url && url.toLowerCase().startsWith('vbscript:')) {
        event.preventDefault();
        executeVbscriptUrl(engine, url);
      }
    };

    (window.navigation as EventTarget).addEventListener('navigate', navigateHandler);
  } else {
    clickHandler = (event: MouseEvent): void => {
      const target = event.target as Element | null;
      const link = target?.closest('a,area');
      if (!link) return;

      const href = link.getAttribute('href');
      if (href && href.toLowerCase().startsWith('vbscript:')) {
        event.preventDefault();
        executeVbscriptUrl(engine, href);
      }
    };

    document.addEventListener('click', clickHandler, true);
  }

  return { navigateHandler, clickHandler };
}

export function cleanupVbscriptProtocol(result: ProtocolHandlerResult): void {
  if (typeof window === 'undefined') return;

  if (result.navigateHandler && 'navigation' in window && window.navigation) {
    (window.navigation as EventTarget).removeEventListener('navigate', result.navigateHandler);
  }

  if (result.clickHandler) {
    document.removeEventListener('click', result.clickHandler, true);
  }
}
