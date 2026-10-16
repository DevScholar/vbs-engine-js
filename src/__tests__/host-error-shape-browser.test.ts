import { afterAll, beforeAll, describe, expect, it, vi } from 'vitest';
import { VbsEngine, type VbsHostError } from '../core/index.ts';

/**
 * The shape a browser hands a page is the one Internet Explorer handed JScript in document mode 10, read off an
 * HTA: `name` comes from the prototype rather than being written on the error, `message` is JavaScript's own
 * non-enumerable property, `description` and `number` are the engine's two additions, and there is nothing else at
 * all. In particular there is no `line` and no `column`, which the host never put on the object: a page that wants
 * the position reads `engine.error`.
 *
 * The browser shape belongs to an engine standing for a browser host, and the first engine in a process is the one
 * that claims host integration, so this file holds that engine and nothing else. The general environment gets a
 * different message instead, the one line `cscript` printed, which `features.test.ts` holds.
 */
const BROWSER_OPTIONS = {
  mode: 'browser' as const,
  parseScriptElement: false,
  parseInlineEventAttributes: false,
  parseEventSubNames: false,
  parseVbsProtocol: false,
  overrideJsEvalFunctions: false,
  injectVBArrayToGlobalThis: false,
  injectGlobalThis: false,
};

beforeAll(() => {
  vi.stubGlobal('window', { addEventListener() {}, removeEventListener() {} });
  vi.stubGlobal('document', { readyState: 'complete' });
});

afterAll(() => {
  vi.unstubAllGlobals();
});

describe('the error shape a browser hands a page', () => {
  it('hands a page the shape Internet Explorer handed JScript', () => {
    const engine = new VbsEngine(BROWSER_OPTIONS);
    const owns = (object: object, name: string): boolean =>
      Object.prototype.hasOwnProperty.call(object, name);
    let caught: VbsHostError | null = null;
    try {
      engine.executeStatementThrows('Dim divisor\n  divisor = 0\nx = 1 / divisor');
    } catch (e) {
      caught = e as VbsHostError;
    }

    expect(caught).toBeInstanceOf(Error);
    expect(caught!.name).toBe('Error');
    expect(caught!.number).toBe(-2146828277); // 0x800A000B
    expect(caught!.description).toBe('Division by zero');
    expect(caught!.message).toBe('Division by zero');

    // Written by the engine, and so enumerable the way the host's own `description` and `number` were. The host
    // listed its stack here too, so this one is enumerated as well. The order is the host's business, so only the
    // set is compared: the host listed `description,number,stack` and this one sorts out as the same.
    expect(Object.keys(caught!).sort()).toEqual(['description', 'number', 'stack']);
    // JavaScript's own, so not enumerated, and inherited from `Error.prototype` rather than written.
    expect(owns(caught!, 'message')).toBe(true);
    expect(owns(caught!, 'name')).toBe(false);
    // The host kept its stack as an own property, and this one is the host's own rather than a copy.
    expect(owns(caught!, 'stack')).toBe(true);
    expect(caught!.stack).toContain('Division by zero');
    // The host gave a page no position on the error object, so there is none here.
    expect(owns(caught!, 'line')).toBe(false);
    expect(owns(caught!, 'column')).toBe(false);
    expect((caught as unknown as Record<string, unknown>).lineNumber).toBeUndefined();
    expect((caught as unknown as Record<string, unknown>).columnNumber).toBeUndefined();

    // A browser host that names its code, and says where that code stands, still gets the description as the
    // message: Internet Explorer printed no line like the one Windows Script Host printed. The position reaches
    // `engine.error` either way, which is where a page reads it.
    //
    // This engine also stands for a host, because it is the first of this file and claimed the integration, so a
    // failure in added code is recorded and not raised: a host owns its failures, and Internet Explorer reports a
    // failing `<script>` block and carries on. The other half of that rule, an engine that stands for no host and
    // raises, is held in `host-integration.test.ts`.
    expect(() =>
      engine.addCode('Dim divisor\n  divisor = 0\nx = 1 / divisor', {
        name: 'C:\\work\\demo.vbs',
        // Counted from zero, so three puts the code's first line on the script's fourth line and its failing third
        // line on the sixth.
        startingLineNumber: 3,
      })
    ).not.toThrow();
    expect(engine.error!.line).toBe(6);
    let named = '';
    try {
      engine.executeStatementThrows('x = 1 / divisor');
    } catch (e) {
      named = (e as Error).message;
    }
    expect(named).toBe('Division by zero');
  });
});
