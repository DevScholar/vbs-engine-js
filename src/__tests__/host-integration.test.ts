import { describe, it, expect } from 'vitest';
import { VbsEngine } from '../index.ts';

const globals = globalThis as Record<string, unknown>;

// A host has one host-integrated engine: the first one to ask for it. Every later engine is isolated,
// so two of them cannot fight over the host globals, the event names or the locale.
//
// The claim lasts for the process, so these tests run in the order written: the first one opts out,
// which leaves the claim to the engine that follows it.
describe('Host integration is claimed once per process', () => {
  it('leaves the claim to a later engine when an engine opts out', () => {
    const optingOut = new VbsEngine({ hostIntegration: false });
    optingOut.addCode('Function OptedOut()\nOptedOut = 1\nEnd Function');
    expect(globals.OptedOut).toBeUndefined();

    // The isolated engine did not consume the claim, so this one still integrates.
    const host = new VbsEngine();
    host.addCode('Function HostFn()\nHostFn = 42\nEnd Function');
    expect(typeof globals.HostFn).toBe('function');

    // A variable is published under the name the script wrote, not the lower case name the engine looks names
    // up with: JavaScript is case sensitive, so the two are different names on the host. Publishing the lower
    // case one left `vbsMessage` undefined on the example page while `vbsmessage` held the value.
    host.addCode('publishedMessage = "Hello from VBS"');
    expect(globals.publishedMessage).toBe('Hello from VBS');
    expect(globals.publishedmessage).toBeUndefined();
  });

  it('isolates every engine constructed after the claim is taken', () => {
    const later = new VbsEngine();
    later.addCode('Function LaterFn()\nLaterFn = 1\nEnd Function');
    expect(globals.LaterFn).toBeUndefined();
    // It still runs its own code, and it is still callable through the engine.
    expect(later.run('LaterFn')).toBe(1);
  });

  it('still lets an isolated engine read host globals and host objects', () => {
    const isolated = new VbsEngine();
    isolated.addObject('probe', { value: 7 });
    isolated.executeStatement('r = probe.value + Math.round(1.4)');
    expect(isolated.eval('r')).toBe(8);
  });
});
