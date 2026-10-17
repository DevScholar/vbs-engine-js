import type { VbsEngine } from "../core/index.ts";

export interface BoundNamedHandler {
  target: EventTarget;
  handler: EventListener;
}

export function setupElementInlineEvents(
  engine: VbsEngine,
  element: Element,
): void {
  const attrs = element.attributes;
  for (let i = 0; i < attrs.length; i++) {
    const attr = attrs[i];
    if (!attr) continue;
    const attrName = attr.name.toLowerCase();

    if (attrName.startsWith("on") && attrName.length > 2) {
      const value = attr.value.trim();
      if (value.toLowerCase().startsWith("vbscript:")) {
        const code = value.substring(9).trim();

        (element as unknown as Record<string, unknown>)[attrName] =
          (): void => {
            // `executeHostScript` is what decides how an error from script the host asked for surfaces: with
            // `useJsError` it is a real JavaScript error, which is the shape IE's script error reaches the page
            // in, and `window.onerror` sees it; without it the error waits in `engine.error`. A prefix of this
            // module's own making told the page nothing and matched neither host.
            engine.executeHostScript(code);
          };
      }
    }
  }
}

export function setupInlineEventHandlers(engine: VbsEngine): void {
  if (typeof document === "undefined") return;

  const allElements = document.querySelectorAll("*");
  allElements.forEach((element) => {
    setupElementInlineEvents(engine, element);
  });
}

export function resolveEventTarget(objectName: string): EventTarget | null {
  if (objectName in globalThis) {
    const obj = (globalThis as Record<string, unknown>)[objectName];
    if (obj && typeof (obj as EventTarget).addEventListener === "function") {
      return obj as EventTarget;
    }
  }

  const lowerName = objectName.toLowerCase();
  for (const key of Object.keys(globalThis)) {
    if (key.toLowerCase() === lowerName) {
      const obj = (globalThis as Record<string, unknown>)[key];
      if (obj && typeof (obj as EventTarget).addEventListener === "function") {
        return obj as EventTarget;
      }
    }
  }

  if (typeof document !== "undefined") {
    const element = document.getElementById(objectName);
    if (element) return element;

    const allElements = document.querySelectorAll("[id]");
    for (let i = 0; i < allElements.length; i++) {
      const el = allElements[i];
      if (el && el.id.toLowerCase() === lowerName) {
        return el;
      }
    }
  }

  return null;
}

export function setupNamedEventHandlers(
  engine: VbsEngine,
  boundNamedHandlers: Map<string, BoundNamedHandler>,
): void {
  if (typeof window === "undefined") return;

  const funcRegistry = engine._getContext()?.functionRegistry;
  if (!funcRegistry) return;

  const userFuncs = funcRegistry.getUserDefinedFunctions?.();
  if (!userFuncs) return;

  for (const [funcName] of userFuncs) {
    const onIndex = funcName.toLowerCase().indexOf("_on");
    if (onIndex > 0) {
      const objectName = funcName.substring(0, onIndex);
      const eventName = funcName.substring(onIndex + 3);

      const target = resolveEventTarget(objectName);

      if (target) {
        const eventType = eventName.toLowerCase();

        if (boundNamedHandlers.has(funcName)) {
          const existing = boundNamedHandlers.get(funcName)!;
          existing.target.removeEventListener(eventType, existing.handler);
        }

        // The Sub is called the way the host calls it, so an error it raises surfaces the same way an inline
        // event attribute's error does rather than being swallowed into a console line of this module's own.
        const handler = (): void => {
          engine.executeHostScript(funcName);
        };

        target.addEventListener(eventType, handler);
        boundNamedHandlers.set(funcName, { target, handler });
      }
    }
  }
}

export function cleanupNamedEventHandlers(
  boundNamedHandlers: Map<string, BoundNamedHandler>,
): void {
  boundNamedHandlers.forEach(({ target, handler }, funcName) => {
    const onIndex = funcName.toLowerCase().indexOf("_on");
    if (onIndex > 0) {
      const eventName = funcName.substring(onIndex + 3).toLowerCase();
      target.removeEventListener(eventName, handler);
    }
  });
  boundNamedHandlers.clear();
}
