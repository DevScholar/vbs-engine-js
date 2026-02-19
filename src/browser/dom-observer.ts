import type { VbsEngine } from '../core/index.ts';
import { isVbscriptElement, parseScriptElement, setupForEventScript } from './script-parser.ts';
import { setupElementInlineEvents, setupNamedEventHandlers, syncFunctionsToGlobalThis } from './event-handlers.ts';
import type { BrowserRuntimeOptions } from './types.ts';
import type { BoundNamedHandler } from './event-handlers.ts';

export interface ObserverContext {
  observer: MutationObserver;
}

export function startObserver(
  engine: VbsEngine,
  options: Required<BrowserRuntimeOptions>,
  boundNamedHandlers: Map<string, BoundNamedHandler>
): ObserverContext {
  const observer = new MutationObserver((mutations) => {
    mutations.forEach((mutation) => {
      mutation.addedNodes.forEach((node) => {
        if (node.nodeType === Node.ELEMENT_NODE) {
          const element = node as Element;

          if (options.parseScriptElement && element.tagName.toLowerCase() === 'script' && isVbscriptElement(element)) {
            const code = element.textContent ?? '';
            const forAttr = element.getAttribute('for');
            const eventAttr = element.getAttribute('event');

            if (forAttr && eventAttr) {
              setupForEventScript(engine, forAttr, eventAttr, code);
            } else {
              try {
                engine.run(code);
                if (options.injectGlobalThis) {
                  syncFunctionsToGlobalThis(engine);
                }
                if (options.injectGlobalThis && options.parseEventSubNames) {
                  setupNamedEventHandlers(engine, boundNamedHandlers);
                }
              } catch (error) {
                console.error('Vbscript error:', error);
              }
            }
          }

          if (options.parseInlineEventAttributes) {
            setupElementInlineEvents(engine, element);
          }

          if (options.parseScriptElement) {
            const scripts = element.querySelectorAll('script');
            scripts.forEach(script => {
              if (isVbscriptElement(script)) {
                const code = script.textContent ?? '';
                const forAttr = script.getAttribute('for');
                const eventAttr = script.getAttribute('event');

                if (forAttr && eventAttr) {
                  setupForEventScript(engine, forAttr, eventAttr, code);
                } else {
                  try {
                    engine.run(code);
                    if (options.injectGlobalThis) {
                      syncFunctionsToGlobalThis(engine);
                    }
                    if (options.injectGlobalThis && options.parseEventSubNames) {
                      setupNamedEventHandlers(engine, boundNamedHandlers);
                    }
                  } catch (error) {
                    console.error('Vbscript error:', error);
                  }
                }
              }
            });
          }

          if (options.parseInlineEventAttributes) {
            const childElements = element.querySelectorAll('*');
            childElements.forEach(child => {
              setupElementInlineEvents(engine, child);
            });
          }
        }
      });
    });
  });

  observer.observe(document.documentElement, {
    childList: true,
    subtree: true,
  });

  return { observer };
}

export function stopObserver(context: ObserverContext): void {
  context.observer.disconnect();
}
