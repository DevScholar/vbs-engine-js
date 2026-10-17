import type { VbsEngine } from "../core/index.ts";
import { withDocumentLineNumbers } from "./document-lines.ts";

export function isVbscriptElement(element: Element): boolean {
  const language = element.getAttribute("language");
  const type = element.getAttribute("type");

  if (language?.toLowerCase() === "vbscript") return true;
  if (type?.toLowerCase() === "text/vbscript") return true;
  if (type?.toLowerCase() === "text/vbs") return true;
  if (type?.toLowerCase() === "application/x-vbscript") return true;

  return false;
}

/**
 * What a script element holds, written the way the engine reads it, with the newlines that put its first line
 * where the block's first line stands in the document. Internet Explorer reported a failure in an inline block at
 * the line in the document, so a page that reads a failure from this engine sees the same number it would have
 * seen from the browser. See `document-lines.ts` for how that line is worked out and what it cannot see.
 *
 * The legacy pattern `<!-- ... //-->` is undone first, since old pages wrapped their script that way to hide it
 * from browsers that did not know the element.
 */
export function extractVbscriptCode(script: Element): string {
  let code = script.textContent ?? "";

  // Trim whitespace
  code = code.trim();

  // Handle HTML comment wrapping: <!-- ... //-->
  // This was used in old browsers (1998 era) to hide scripts from browsers that didn't support <script>
  if (code.startsWith("<!--")) {
    // Find the end of the comment
    // Support both //--> and -->
    const endIndex = code.indexOf("//-->");
    if (endIndex !== -1) {
      code = code.substring(4, endIndex).trim();
    } else if (code.endsWith("-->")) {
      code = code.substring(4, code.length - 3).trim();
    }
  }

  return withDocumentLineNumbers(code, script);
}

export function parseScriptElement(
  engine: VbsEngine,
  script: Element,
  onScriptRun?: () => void,
): void {
  if (!isVbscriptElement(script)) return;

  const code = extractVbscriptCode(script);
  const forAttr = script.getAttribute("for");
  const eventAttr = script.getAttribute("event");

  if (forAttr && eventAttr) {
    setupForEventScript(engine, forAttr, eventAttr, code);
  } else {
    // `addCode` keeps a failure in `engine.error` instead of throwing, so there is nothing to catch here and
    // no prefix for this module to invent.
    engine.addCode(code);
    onScriptRun?.();
  }
}

export function setupForEventScript(
  engine: VbsEngine,
  targetId: string,
  eventName: string,
  code: string,
): void {
  const eventType = eventName.toLowerCase().replace(/^on/, "");

  const bindHandler = (): boolean => {
    const target = document.getElementById(targetId);
    if (target) {
      // The handler's own text is script the host asked for, so it goes through `executeHostScript` and its
      // error surfaces the way every other host run error does: a real JavaScript error when `useJsError` is
      // on, otherwise a record in `engine.error`. `addCode` would have kept it in `engine.error` either way.
      const handler = (): void => {
        engine.executeHostScript(code);
      };
      target.addEventListener(eventType, handler);
      return true;
    }
    return false;
  };

  if (!bindHandler()) {
    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", () => {
        bindHandler();
      });
    } else {
      setTimeout(bindHandler, 0);
    }
  }
}

export function autoRunScripts(
  engine: VbsEngine,
  onScriptRun?: () => void,
): void {
  if (typeof document === "undefined") return;

  const scripts = document.querySelectorAll("script");
  scripts.forEach((script) => {
    parseScriptElement(engine, script, onScriptRun);
  });
}
