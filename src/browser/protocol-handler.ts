import type { VbsEngine } from '../core/index.ts';

export interface ProtocolHandlerResult {
  navigateHandler: ((event: Event) => void) | null;
  clickHandler: ((event: MouseEvent) => void) | null;
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
        const code = url.substring(9);
        try {
          engine.run(code);
        } catch (error) {
          console.error('VBScript protocol error:', error);
        }
      }
    };

    (window.navigation as EventTarget).addEventListener('navigate', navigateHandler);
  } else {
    clickHandler = (event: MouseEvent): void => {
      const target = event.target as HTMLElement;

      if (target.tagName === 'A' || target.tagName === 'AREA') {
        const href = target.getAttribute('href');
        if (href && href.toLowerCase().startsWith('vbscript:')) {
          event.preventDefault();
          const code = href.substring(9);
          try {
            engine.run(code);
          } catch (error) {
            console.error('VBScript protocol error:', error);
          }
        }
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
