import { afterAll, beforeAll, describe, expect, it, vi } from 'vitest';
import { VbsEngine } from '../core/index.ts';

/**
 * The host's end in a browser is the page going away, and `pagehide` is the event that says so. The engine wants
 * a `window` to stand for, so this file gives it a small one and switches every browser feature off: what is
 * under test is only the moment the engine listens for, not the DOM work around it.
 *
 * A page that goes into the back and forward cache is not ending. It can come back, and the instances it holds
 * would then be terminated while the script still names them, so only a real departure lets the script go.
 */
type PageHideListener = (event: { persisted?: boolean }) => void;

const pageHideListeners = new Set<PageHideListener>();

const fakeWindow = {
  addEventListener(type: string, listener: PageHideListener) {
    if (type === 'pagehide') pageHideListeners.add(listener);
  },
  removeEventListener(type: string, listener: PageHideListener) {
    if (type === 'pagehide') pageHideListeners.delete(listener);
  },
};

/** The page goes away, either for good or into the back and forward cache. */
function leavePage(persisted: boolean): void {
  for (const listener of [...pageHideListeners]) listener({ persisted });
}

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

const SCRIPT = `Dim log
log = ""
Class Counter
  Private Sub Class_Terminate()
    log = log & "terminate;"
  End Sub
End Class
Set held = New Counter
`;

beforeAll(() => {
  vi.stubGlobal('window', fakeWindow);
  vi.stubGlobal('document', { readyState: 'complete' });
});

afterAll(() => {
  vi.unstubAllGlobals();
});

describe('the end of the host in a browser', () => {
  it('lets the script go when the page leaves, but not when it only goes into the cache', () => {
    // The first engine of this file claims host integration, which is the one that stands for the host.
    const engine = new VbsEngine(BROWSER_OPTIONS);
    engine.executeStatement(SCRIPT);
    expect(engine.eval('log')).toBe('');
    expect(pageHideListeners.size).toBe(1);

    leavePage(true);
    // The page can come back, so the instance the script still names has to stay alive.
    expect(engine.eval('log')).toBe('');

    leavePage(false);
    expect(engine.eval('log')).toBe('terminate;');

    engine.destroy();
    expect(pageHideListeners.size).toBe(0);
  });

  it('leaves an isolated engine with no end of its own', () => {
    const isolated = new VbsEngine({ ...BROWSER_OPTIONS, hostIntegration: false });
    isolated.executeStatement(SCRIPT);
    expect(pageHideListeners.size).toBe(0);

    leavePage(false);

    expect(isolated.eval('log')).toBe('');
  });
});
