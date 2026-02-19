import { VbsEngine } from '../core/index.ts';
import type { VbValue } from '../runtime/index.ts';
import { createBrowserMsgBox } from '../builtins/msgbox.ts';
import { createBrowserInputBox } from '../builtins/inputbox.ts';

type OriginalSetTimeout = typeof setTimeout;
type OriginalSetInterval = typeof setInterval;
type OriginalEval = typeof eval;

function vbToJs(value: VbValue): unknown {
  switch (value.type) {
    case 'Empty':
      return undefined;
    case 'Null':
      return null;
    case 'Boolean':
    case 'Long':
    case 'Double':
    case 'Integer':
    case 'String':
      return value.value;
    case 'Date':
      if (value.value instanceof Date) {
        return value.value;
      }
      return new Date(value.value as string);
    case 'Array':
      return value.value;
    case 'Object':
      return value.value;
    default:
      return value.value;
  }
}

function jsToVb(value: unknown, thisArg?: unknown): VbValue {
  if (value === undefined) return { type: 'Empty', value: undefined };
  if (value === null) return { type: 'Null', value: null };
  if (typeof value === 'boolean') return { type: 'Boolean', value };
  if (typeof value === 'number') {
    if (Number.isInteger(value) && value >= -2147483648 && value <= 2147483647) {
      return { type: 'Long', value };
    }
    return { type: 'Double', value };
  }
  if (typeof value === 'string') return { type: 'String', value };
  if (value instanceof Date) return { type: 'Date', value };
  if (Array.isArray(value)) {
    return { type: 'Array', value };
  }
  if (typeof value === 'function') {
    return { type: 'Object', value: { type: 'jsfunction', func: value as (...args: unknown[]) => unknown, thisArg: thisArg ?? null } };
  }
  if (typeof value === 'object') {
    return { type: 'Object', value };
  }
  return { type: 'String', value: String(value) };
}

export interface BrowserRuntimeOptions {
  parseScriptElement?: boolean;
  parseInlineEventAttributes?: boolean;
  injectGlobalThis?: boolean;
  parseEventSubNames?: boolean;
  maxExecutingTime?: number;
  overrideJSEvalFunctions?: boolean;
  parseVbsProtocol?: boolean;
}

export class VbsBrowserEngine {
  private engine: VbsEngine;
  private initialized: boolean = false;
  private originalSetTimeout: OriginalSetTimeout | null = null;
  private originalSetInterval: OriginalSetInterval | null = null;
  private originalEval: OriginalEval | null = null;
  private observer: MutationObserver | null = null;
  private boundNamedHandlers: Map<string, { target: EventTarget; handler: EventListener }> = new Map();
  private navigateHandler: ((event: NavigateEvent) => void) | null = null;
  private options: Required<BrowserRuntimeOptions>;

  constructor(options: BrowserRuntimeOptions = {}) {
    this.engine = new VbsEngine();
    
    this.options = {
      parseScriptElement: options.parseScriptElement ?? true,
      parseInlineEventAttributes: options.parseInlineEventAttributes ?? true,
      injectGlobalThis: options.injectGlobalThis ?? true,
      parseEventSubNames: options.parseEventSubNames ?? true,
      maxExecutingTime: options.maxExecutingTime ?? -1,
      overrideJSEvalFunctions: options.overrideJSEvalFunctions ?? true,
      parseVbsProtocol: options.parseVbsProtocol ?? true,
    };

    if (this.options.maxExecutingTime > 0) {
      this.engine.setMaxExecutionTime(this.options.maxExecutingTime);
    }

    if (typeof window !== 'undefined') {
      if (this.options.overrideJSEvalFunctions) {
        this.overrideTimers();
        this.overrideEval();
      }
      this.registerBrowserFunctions();
      if (this.options.parseVbsProtocol) {
        this.setupVbscriptProtocol();
      }

      if (this.options.parseScriptElement) {
        this.autoRunScripts();
      }
      
      if (this.options.parseInlineEventAttributes) {
        this.setupInlineEventHandlers();
      }
      
      if (this.options.injectGlobalThis && this.options.parseEventSubNames) {
        this.setupNamedEventHandlers();
      }
      
      if (this.options.parseScriptElement || this.options.parseInlineEventAttributes) {
        this.startObserver();
      }
      
      if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', () => {
          if (this.options.injectGlobalThis && this.options.parseEventSubNames) {
            this.setupNamedEventHandlers();
          }
        });
      }
      
      this.initialized = true;
    }
  }

  private registerBrowserFunctions(): void {
    this.engine.registerFunction('MsgBox', createBrowserMsgBox());
    this.engine.registerFunction('InputBox', createBrowserInputBox());

    this.engine.registerFunction('GetRef', (procname: VbValue): VbValue => {
      const name = String(procname.value);
      const self = this;
      const func = function(this: unknown, ...args: unknown[]): unknown {
        const vbArgs = args.map(a => jsToVb(a, this));
        const context = self.engine.getContext();
        if (!context) return undefined;
        const result = context.functionRegistry.call(name, vbArgs);
        return vbToJs(result);
      };
      return { type: 'Object', value: { type: 'vbref', name, func } };
    });

    this.engine.registerFunction('CreateObject', (cls: VbValue, _servername?: VbValue): VbValue => {
      const className = String(cls.value ?? cls);
      
      if (typeof window.ActiveXObject !== 'undefined') {
        try {
          const ax = new window.ActiveXObject(className);
          return { type: 'Object', value: { 
            type: 'activex', 
            object: ax, 
            className 
          }};
        } catch (e) {
          throw new Error(`ActiveX component can't create object: '${className}'`);
        }
      }
      
      throw new Error(`ActiveXObject is not supported in this browser environment. Cannot create: '${className}'`);
    });

    this.engine.registerFunction('GetObject', (pathname?: VbValue, cls?: VbValue): VbValue => {
      if (typeof window.ActiveXObject !== 'undefined') {
        const path = pathname ? String(pathname.value ?? pathname) : '';
        const className = cls ? String(cls.value ?? cls) : '';
        
        try {
          if (path) {
            const ax = new window.ActiveXObject(className || 'Scripting.FileSystemObject');
            return { type: 'Object', value: { 
              type: 'activex', 
              object: ax, 
              className 
            }};
          }
        } catch (e) {
          throw new Error(`ActiveX component can't create object: '${className || path}'`);
        }
      }
      
      const className = cls ? String(cls.value ?? cls) : '';
      throw new Error(`ActiveXObject is not supported in this browser environment. Cannot get: '${className}'`);
    });
  }

  private setupVbscriptProtocol(): void {
    if (typeof window === 'undefined') return;

    if ('navigation' in window && window.navigation) {
      this.navigateHandler = (event: NavigateEvent): void => {
        const url = event.destination.url;
        if (url && url.toLowerCase().startsWith('vbscript:')) {
          event.preventDefault();
          const code = url.substring(9);
          try {
            this.engine.run(code);
          } catch (error) {
            console.error('VBScript protocol error:', error);
          }
        }
      };
      
      window.navigation.addEventListener('navigate', this.navigateHandler);
    } else {
      document.addEventListener('click', (event: MouseEvent) => {
        const target = event.target as HTMLElement;
        
        if (target.tagName === 'A' || target.tagName === 'AREA') {
          const href = target.getAttribute('href');
          if (href && href.toLowerCase().startsWith('vbscript:')) {
            event.preventDefault();
            const code = href.substring(9);
            try {
              this.engine.run(code);
            } catch (error) {
              console.error('VBScript protocol error:', error);
            }
          }
        }
      }, true);
    }
  }

  private overrideTimers(): void {
    if (typeof window === 'undefined') return;

    const self = this;
    this.originalSetTimeout = window.setTimeout;
    this.originalSetInterval = window.setInterval;

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
          return self.originalSetTimeout!.call(window, () => {
            const funcRegistry = self.engine.getContext()?.functionRegistry;
            if (funcRegistry?.has(handler)) {
              funcRegistry.call(handler, []);
            } else {
              self.engine.run(handler);
            }
          }, actualDelay);
        }
        
        return self.originalSetTimeout!.call(window, () => {
          const funcRegistry = self.engine.getContext()?.functionRegistry;
          if (funcRegistry?.has(handler)) {
            funcRegistry.call(handler, []);
          } else {
            self.engine.run(handler);
          }
        }, actualDelay);
      }

      return self.originalSetTimeout!.call(window, handler as TimerHandler, delay as number | undefined, ...[languageOrArg, ...args].filter(a => a !== undefined));
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
          return self.originalSetInterval!.call(window, () => {
            const funcRegistry = self.engine.getContext()?.functionRegistry;
            if (funcRegistry?.has(handler)) {
              funcRegistry.call(handler, []);
            } else {
              self.engine.run(handler);
            }
          }, actualDelay);
        }
        
        return self.originalSetInterval!.call(window, () => {
          const funcRegistry = self.engine.getContext()?.functionRegistry;
          if (funcRegistry?.has(handler)) {
            funcRegistry.call(handler, []);
          } else {
            self.engine.run(handler);
          }
        }, actualDelay);
      }

      return self.originalSetInterval!.call(window, handler as TimerHandler, delay as number | undefined, ...[languageOrArg, ...args].filter(a => a !== undefined));
    };
  }

  private overrideEval(): void {
    if (typeof window === 'undefined') return;

    const self = this;
    this.originalEval = window.eval;

    (window as unknown as Record<string, unknown>).vbsEval = function(code: unknown): unknown {
      const result = self.engine.run(String(code));
      return vbToJs(result);
    };

    (window as unknown as Record<string, unknown>).eval = function(code: unknown, language?: unknown): unknown {
      if (typeof language === 'string') {
        const lang = language.toLowerCase();
        if (lang === 'vbscript' || lang === 'vbs') {
          const result = self.engine.run(String(code));
          return vbToJs(result);
        }
      }
      return self.originalEval!.call(window, code);
    };
  }

  private isVbscriptElement(element: Element): boolean {
    const language = element.getAttribute('language');
    const type = element.getAttribute('type');

    if (language?.toLowerCase() === 'vbscript') return true;
    if (type?.toLowerCase() === 'text/vbscript') return true;
    if (type?.toLowerCase() === 'application/x-vbscript') return true;

    return false;
  }

  private autoRunScripts(): void {
    if (typeof document === 'undefined') return;

    const scripts = document.querySelectorAll('script');
    scripts.forEach(script => {
      if (this.isVbscriptElement(script)) {
        const code = script.textContent ?? '';
        const forAttr = script.getAttribute('for');
        const eventAttr = script.getAttribute('event');
        
        if (forAttr && eventAttr) {
          this.setupForEventScript(forAttr, eventAttr, code);
        } else {
          try {
            this.engine.run(code);
          } catch (error) {
            console.error('Vbscript error:', error);
          }
        }
      }
    });

    if (this.options.injectGlobalThis) {
      this.syncFunctionsToGlobalThis();
    }
  }

  private setupForEventScript(targetId: string, eventName: string, code: string): void {
    const eventType = eventName.toLowerCase().replace(/^on/, '');
    
    const bindHandler = (): boolean => {
      const target = document.getElementById(targetId);
      if (target) {
        const handler = (): void => {
          try {
            this.engine.run(code);
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

  private setupInlineEventHandlers(): void {
    if (typeof document === 'undefined') return;

    const allElements = document.querySelectorAll('*');
    allElements.forEach(element => {
      this.setupElementInlineEvents(element);
    });
  }

  private setupElementInlineEvents(element: Element): void {
    const attrs = element.attributes;
    for (let i = 0; i < attrs.length; i++) {
      const attr = attrs[i];
      const attrName = attr.name.toLowerCase();
      
      if (attrName.startsWith('on') && attrName.length > 2) {
        const value = attr.value.trim();
        if (value.toLowerCase().startsWith('vbscript:')) {
          const code = value.substring(9).trim();
          const eventType = attrName.substring(2);
          
          (element as Record<string, unknown>)[attrName] = (event: Event): void => {
            try {
              this.engine.run(code);
            } catch (error) {
              console.error('VBScript event handler error:', error);
            }
          };
        }
      }
    }
  }

  private setupNamedEventHandlers(): void {
    if (typeof window === 'undefined') return;

    const funcRegistry = this.engine.getContext()?.functionRegistry;
    if (!funcRegistry) return;

    const userFuncs = funcRegistry.getUserDefinedFunctions?.();
    if (!userFuncs) return;

    for (const [funcName] of userFuncs) {
      const onIndex = funcName.toLowerCase().indexOf('_on');
      if (onIndex > 0) {
        const objectName = funcName.substring(0, onIndex);
        const eventName = funcName.substring(onIndex + 3);
        
        const target = this.resolveEventTarget(objectName);
        
        if (target) {
          const eventType = eventName.toLowerCase();
          
          if (this.boundNamedHandlers.has(funcName)) {
            const existing = this.boundNamedHandlers.get(funcName)!;
            existing.target.removeEventListener(eventType, existing.handler);
          }
          
          const handler = (): void => {
            try {
              this.engine.getContext()?.functionRegistry.call(funcName, []);
            } catch (error) {
              console.error(`VBScript ${funcName} error:`, error);
            }
          };
          
          target.addEventListener(eventType, handler);
          this.boundNamedHandlers.set(funcName, { target, handler });
        }
      }
    }
  }

  private resolveEventTarget(objectName: string): EventTarget | null {
    if (objectName in globalThis) {
      const obj = (globalThis as Record<string, unknown>)[objectName];
      if (obj && typeof (obj as EventTarget).addEventListener === 'function') {
        return obj as EventTarget;
      }
    }
    
    const lowerName = objectName.toLowerCase();
    for (const key of Object.keys(globalThis)) {
      if (key.toLowerCase() === lowerName) {
        const obj = (globalThis as Record<string, unknown>)[key];
        if (obj && typeof (obj as EventTarget).addEventListener === 'function') {
          return obj as EventTarget;
        }
      }
    }
    
    if (typeof document !== 'undefined') {
      const element = document.getElementById(objectName);
      if (element) return element;
      
      const allElements = document.querySelectorAll('[id]');
      for (let i = 0; i < allElements.length; i++) {
        if (allElements[i].id.toLowerCase() === lowerName) {
          return allElements[i];
        }
      }
    }
    
    return null;
  }

  private startObserver(): void {
    if (typeof document === 'undefined') return;

    this.observer = new MutationObserver((mutations) => {
      mutations.forEach((mutation) => {
        mutation.addedNodes.forEach((node) => {
          if (node.nodeType === Node.ELEMENT_NODE) {
            const element = node as Element;

            if (this.options.parseScriptElement && element.tagName.toLowerCase() === 'script' && this.isVbscriptElement(element)) {
              const code = element.textContent ?? '';
              const forAttr = element.getAttribute('for');
              const eventAttr = element.getAttribute('event');
              
              if (forAttr && eventAttr) {
                this.setupForEventScript(forAttr, eventAttr, code);
              } else {
                try {
                  this.engine.run(code);
                  if (this.options.injectGlobalThis) {
                    this.syncFunctionsToGlobalThis();
                  }
                  if (this.options.injectGlobalThis && this.options.parseEventSubNames) {
                    this.setupNamedEventHandlers();
                  }
                } catch (error) {
                  console.error('Vbscript error:', error);
                }
              }
            }

            if (this.options.parseInlineEventAttributes) {
              this.setupElementInlineEvents(element);
            }

            if (this.options.parseScriptElement) {
              const scripts = element.querySelectorAll('script');
              scripts.forEach(script => {
                if (this.isVbscriptElement(script)) {
                  const code = script.textContent ?? '';
                  const forAttr = script.getAttribute('for');
                  const eventAttr = script.getAttribute('event');
                  
                  if (forAttr && eventAttr) {
                    this.setupForEventScript(forAttr, eventAttr, code);
                  } else {
                    try {
                      this.engine.run(code);
                      if (this.options.injectGlobalThis) {
                        this.syncFunctionsToGlobalThis();
                      }
                      if (this.options.injectGlobalThis && this.options.parseEventSubNames) {
                        this.setupNamedEventHandlers();
                      }
                    } catch (error) {
                      console.error('Vbscript error:', error);
                    }
                  }
                }
              });
            }

            if (this.options.parseInlineEventAttributes) {
              const childElements = element.querySelectorAll('*');
              childElements.forEach(child => {
                this.setupElementInlineEvents(child);
              });
            }
          }
        });
      });
    });

    this.observer.observe(document.documentElement, {
      childList: true,
      subtree: true,
    });
  }

  private syncFunctionsToGlobalThis(): void {
    if (typeof globalThis === 'undefined') return;

    const context = this.engine.getContext?.();
    if (!context) return;

    const funcRegistry = context.functionRegistry;
    if (!funcRegistry) return;

    const userFuncs = funcRegistry.getUserDefinedFunctions?.();
    if (!userFuncs) return;

    for (const [key, info] of userFuncs) {
      const funcName = info.name;
      if (!(funcName in (globalThis as Record<string, unknown>))) {
        (globalThis as Record<string, unknown>)[funcName] = (...args: unknown[]) => {
          const vbArgs = args.map(a => jsToVb(a));
          return vbToJs(funcRegistry.call(funcName, vbArgs));
        };
      }
    }

    if (context.globalScope) {
      const allVars = context.globalScope.getAllVariables();
      for (const [varName, vbVar] of allVars) {
        if (vbVar.value && vbVar.value.type !== 'Empty') {
          (globalThis as Record<string, unknown>)[vbVar.name] = vbToJs(vbVar.value);
        }
      }
    }
  }

  run(code: string): VbValue {
    const result = this.engine.run(code);
    if (this.options.injectGlobalThis) {
      this.syncFunctionsToGlobalThis();
    }
    if (this.options.injectGlobalThis && this.options.parseEventSubNames) {
      this.setupNamedEventHandlers();
    }
    return result;
  }

  getVariable(name: string): VbValue {
    return this.engine.getVariable(name);
  }

  setVariable(name: string, value: VbValue): void {
    this.engine.setVariable(name, value);
  }

  registerFunction(name: string, func: (...args: VbValue[]) => VbValue): void {
    this.engine.registerFunction(name, func);
  }

  getEngine(): VbsEngine {
    return this.engine;
  }

  getContext() {
    return this.engine.getContext();
  }

  isInitialized(): boolean {
    return this.initialized;
  }

  destroy(): void {
    if (this.observer) {
      this.observer.disconnect();
      this.observer = null;
    }

    this.boundNamedHandlers.forEach(({ target, handler }, funcName) => {
      const onIndex = funcName.toLowerCase().indexOf('_on');
      if (onIndex > 0) {
        const eventName = funcName.substring(onIndex + 3).toLowerCase();
        target.removeEventListener(eventName, handler);
      }
    });
    this.boundNamedHandlers.clear();

    if (this.navigateHandler && typeof window !== 'undefined' && window.navigation) {
      window.navigation.removeEventListener('navigate', this.navigateHandler);
      this.navigateHandler = null;
    }

    if (this.options.overrideJSEvalFunctions) {
      if (this.originalSetTimeout && typeof window !== 'undefined') {
        (window as unknown as Record<string, unknown>).setTimeout = this.originalSetTimeout;
      }
      if (this.originalSetInterval && typeof window !== 'undefined') {
        (window as unknown as Record<string, unknown>).setInterval = this.originalSetInterval;
      }
      if (this.originalEval && typeof window !== 'undefined') {
        (window as unknown as Record<string, unknown>).eval = this.originalEval;
      }
    }
  }
}

export function createBrowserRuntime(options?: BrowserRuntimeOptions): VbsBrowserEngine {
  return new VbsBrowserEngine(options);
}
