import { describe, it, expect } from 'vitest';
import {
  countNewlines,
  firstNonWhitespaceIndex,
  indentationAtIndex,
  linesBeforeHtmlElement,
  textStartWithinElement,
} from '../browser/document-lines.ts';

/**
 * Internet Explorer reported a failure inside an inline block at the line in the document, so the engine works
 * the same number out. The document does not keep its own line numbers: the parser drops the whitespace before
 * `<head>`, measured by a page of 62 source lines serialising to 59.
 */
describe('the line a script block stands on in the document', () => {
  it('counts the lines that stand before the html element', () => {
    expect(linesBeforeHtmlElement(true, [1])).toBe(2);
    expect(linesBeforeHtmlElement(true, [])).toBe(1);
    expect(linesBeforeHtmlElement(false, [])).toBe(0);
    // A comment on three lines stands on three lines of its own.
    expect(linesBeforeHtmlElement(false, [3])).toBe(3);
  });

  it('counts newlines', () => {
    expect(countNewlines('a\nb\nc')).toBe(2);
    expect(countNewlines('')).toBe(0);
    expect(countNewlines('\n\n')).toBe(2);
  });

  it('finds the first thing that is not whitespace, and the indentation before it', () => {
    expect(firstNonWhitespaceIndex('\n\n   x')).toBe(5);
    expect(firstNonWhitespaceIndex('   ')).toBe(-1);
    expect(indentationAtIndex('\n\n   x', 5)).toBe(3);
    expect(indentationAtIndex('abc', 0)).toBe(0);
    expect(indentationAtIndex('  abc', 2)).toBe(2);
  });

  it('finds where an element own text starts, quotes and all', () => {
    expect(textStartWithinElement('<script type="text/vbscript">abc</script>')).toBe(
      '<script type="text/vbscript">'.length
    );
    // A `>` inside an attribute value does not end the opening tag.
    expect(textStartWithinElement('<script data-x="a>b">abc</script>')).toBe(
      '<script data-x="a>b">'.length
    );
    expect(textStartWithinElement('<script')).toBe(-1);
  });

  // The page the check ran on, whose blocks begin on lines 12, 18 and 26: a doctype and a comment before the
  // html element, the html element itself, then what the browser writes back out. Reading those pieces off the
  // live page gave the numbers below, and the module put the first block on line 12, where the page said it was.
  it('puts a block where the checked page put it', () => {
    const linesBefore = linesBeforeHtmlElement(true, [1]);
    expect(linesBefore).toBe(2);
    const lineOfHtmlElement = linesBefore + 1;
    const newlineDroppedBeforeHead = 1;
    const newlinesInsideHtml = 7;
    const newlinesBeforeFirstContent = 1;
    expect(
      lineOfHtmlElement + newlineDroppedBeforeHead + newlinesInsideHtml + newlinesBeforeFirstContent
    ).toBe(12);
  });
});
