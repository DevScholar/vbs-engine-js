/**
 * The line a failure inside an inline `<script>` block is reported at.
 *
 * Internet Explorer reported the line in the **document**, not the line in the block: measured in Internet
 * Explorer 11 with the page in document mode 8, a statement on line 35 of the file is listed by the F12 console,
 * and passed to `window.onerror`, as line 35. This module works out the same number for a block the engine reads
 * by itself, so a page that reads a failure from this engine sees the number it would have seen from the browser.
 *
 * The document does not keep its own line numbers, so they are counted from what it does keep:
 *
 * * Everything from the `<html>` tag onward is written back out by `documentElement.innerHTML`, and a block is
 *   looked for in there by its own `outerHTML`, which comes from the same serialiser and so matches. The two are
 *   used together rather than a document wide serialisation looked up by the element, because serialising one
 *   element on its own writes the `xmlns` attribute onto it, which the document wide form leaves on `<html>`.
 * * The lines before the `<html>` tag are gone from the DOM: the parser drops the whitespace there, measured by a
 *   page of 62 source lines serialising to 59. They are counted back from the doctype and the comments the
 *   document node still holds, taking each of those, and the `<html>` tag itself, to occupy one line. A blank
 *   line, or a tag split across lines, in that region makes the count short by the lines it does not know about.
 *   The same goes for the whitespace between the `<html>` tag and `<head>`, which the parser also drops and which
 *   is taken to be one newline.
 *
 * @internal
 */

/** How many newlines a piece of text holds. */
export function countNewlines(text: string): number {
  let total = 0;
  for (let at = 0; at < text.length; at++) {
    if (text.charCodeAt(at) === 10) total += 1;
  }
  return total;
}

/** Where the first character that is not whitespace is, or -1 when the text is all whitespace. */
export function firstNonWhitespaceIndex(text: string): number {
  for (let at = 0; at < text.length; at++) {
    if (!/\s/.test(text[at] as string)) return at;
  }
  return -1;
}

/** How many spaces stand between the start of `index`'s line and `index` itself. */
export function indentationAtIndex(text: string, index: number): number {
  let spaces = 0;
  for (let at = index - 1; at >= 0; at--) {
    const character = text[at] as string;
    if (character === ' ') spaces += 1;
    else if (character === '\n') break;
    else return 0;
  }
  return spaces;
}

/**
 * Where the element's own text begins within the element's own serialisation, which is the character after the
 * opening tag. Quotes are tracked, so a `>` inside an attribute value is not mistaken for the end of the tag.
 */
export function textStartWithinElement(serialisedElement: string): number {
  let quote: string | null = null;
  for (let at = 0; at < serialisedElement.length; at++) {
    const character = serialisedElement[at] as string;
    if (quote !== null) {
      if (character === quote) quote = null;
      continue;
    }
    if (character === '"' || character === "'") {
      quote = character;
      continue;
    }
    if (character === '>') return at + 1;
  }
  return -1;
}

/**
 * How many lines stand before the `<html>` tag: the doctype, and the comments the document node still holds
 * before that element, each taken to occupy the lines its own text holds. Measured on a page with a doctype and
 * one single-line comment before `<html>`, the answer is 2, and the `<html>` tag is then on line 3, which is
 * where the browser's own console counted from.
 */
export function linesBeforeHtmlElement(
  hasDoctype: boolean,
  commentLineCounts: readonly number[]
): number {
  return (hasDoctype ? 1 : 0) + commentLineCounts.reduce((total, lines) => total + lines, 0);
}

/** The line counts of the comments the document still holds from before the `<html>` element. */
function commentsBeforeHtml(): number[] {
  // 8 is a comment node, written out rather than read from a global so that this file needs no DOM to load.
  const lineCounts: number[] = [];
  for (const child of Array.from(document.childNodes)) {
    if (child.nodeType === 8) lineCounts.push(1 + countNewlines((child as Comment).data));
  }
  return lineCounts;
}

/** Where the next block is looked for, so blocks are found in document order and repeated markup is not confused. */
let searchFrom = 0;

/**
 * The line the element's own text starts on in the document, or null when it cannot be worked out, which leaves
 * the caller with the line within the block.
 */
export function documentLineOfText(element: Element): number | null {
  const raw = element.textContent ?? '';
  if (raw === '') return null;

  const firstContent = firstNonWhitespaceIndex(raw);
  if (firstContent < 0) return null;

  const serialisedElement = element.outerHTML;
  const textStart = textStartWithinElement(serialisedElement);
  if (textStart < 0) return null;

  const innerHtml = document.documentElement.innerHTML;
  let elementAt = innerHtml.indexOf(serialisedElement, searchFrom);
  if (elementAt < 0) {
    // The document may have been written to since the last block was looked for, so it is read again from the top.
    searchFrom = 0;
    elementAt = innerHtml.indexOf(serialisedElement);
    if (elementAt < 0) return null;
  }
  searchFrom = elementAt + serialisedElement.length;

  const lineOfHtmlElement =
    linesBeforeHtmlElement(document.doctype !== null, commentsBeforeHtml()) + 1;
  const textAt = elementAt + textStart;
  // The whitespace between the `<html>` tag and `<head>` is dropped as well, measured by a served page of 285
  // lines whose inner HTML holds 283, so one newline stands there that cannot be counted. `<head>` is taken to
  // follow the html tag on the next line, which is what a page does unless it leaves a blank line between them.
  return (
    lineOfHtmlElement +
    1 +
    countNewlines(innerHtml.slice(0, textAt)) +
    countNewlines(raw.slice(0, firstContent))
  );
}

/**
 * The code to hand the engine, with the newlines that put its first line where the block's first line stands in
 * the document, and with the indentation the trim took off that first line put back so its column stands too. The
 * code itself is unchanged: leading blank lines and spaces mean nothing to a script.
 */
export function withDocumentLineNumbers(code: string, element: Element): string {
  const line = documentLineOfText(element);
  if (line === null || line <= 1) return code;

  const raw = element.textContent ?? '';
  const firstContent = firstNonWhitespaceIndex(raw);
  const indentation = firstContent < 0 ? 0 : indentationAtIndex(raw, firstContent);
  return `${'\n'.repeat(line - 1)}${' '.repeat(indentation)}${code}`;
}
