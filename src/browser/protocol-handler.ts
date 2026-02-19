import type { VbsEngine } from '../core/index.ts';

export interface ProtocolHandlerResult {
  navigateHandler: ((event: NavigateEvent) => void) | null;
  clickHandler: ((event: MouseEvent) => void) | null;
}

export function setupVbscriptProtocol(engine: VbsEngine): ProtocolHandlerResult {
  if (typeof window === 'undefined') {
    return { navigateHandler: null, clickHandler: null };
  }

  let navigateHandler: ((event: NavigateEvent) => void) | null = null;
  let clickHandler: ((event: MouseEvent) => void) | null = null;

  if ('navigation' in window && window.navigation) {
    navigateHandler = (event: NavigateEvent): void => {
      const url = event.destination.url;
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

    window.navigation.addEventListener('navigate', navigateHandler);
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
    window.navigation.removeEventListener('navigate', result.navigateHandler);
  }

  if (result.clickHandler) {
    document.removeEventListener('click', result.clickHandler, true);
  }
}
