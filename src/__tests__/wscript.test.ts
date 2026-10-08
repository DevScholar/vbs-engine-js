import { describe, it, expect, beforeEach } from 'vitest';
import { VbsEngine } from '../index.ts';
import { createWScript, WScriptQuitSignal, type WshEnvironment } from '../host/wscript.ts';

function makeEnv(overrides: Partial<WshEnvironment> = {}): {
  env: WshEnvironment;
  out: string[];
  lines: string[];
} {
  const out: string[] = [];
  const lines: string[] = [];
  const env: WshEnvironment = {
    mode: 'cscript',
    scriptFullName: 'C:\\scripts\\test.vbs',
    scriptName: 'test.vbs',
    interactive: true,
    unnamedArgs: ['apple', 'banana'],
    namedArgs: { server: 'prod', port: '8080' },
    usage: 'Usage: test.vbs /server:<name>\n',
    writeStdout: (t: string) => out.push(t),
    writeStderr: (t: string) => out.push('ERR:' + t),
    readLine: () => lines.shift() ?? null,
    sleep: () => {},
    fullName: 'C:\\bin\\run-script.exe',
    path: 'C:\\bin',
    version: '5.812',
    buildVersion: 5812,
    currentDirectory: 'C:\\scripts',
    ...overrides,
  };
  return { env, out, lines };
}

describe('WScript host object', () => {
  let engine: VbsEngine;

  beforeEach(() => {
    engine = new VbsEngine({ mode: 'general' });
  });

  function inject(wscript: unknown): void {
    (globalThis as Record<string, unknown>).WScript = wscript;
  }

  it('Echo writes joined args to stdout', () => {
    const { env, out } = makeEnv();
    inject(createWScript(env).object);
    engine.addCode('WScript.Echo "a", 1, True');
    expect(out.join('')).toBe('a 1 true\n');
  });

  it('exposes ScriptName / ScriptFullName / Version', () => {
    const { env } = makeEnv();
    inject(createWScript(env).object);
    expect(engine.eval('WScript.ScriptName')).toBe('test.vbs');
    expect(engine.eval('WScript.ScriptFullName')).toBe('C:\\scripts\\test.vbs');
    expect(engine.eval('WScript.Version')).toBe('5.812');
    expect(engine.eval('WScript.BuildVersion')).toBe(5812);
  });

  it('Arguments supports Count, positional (n), and Item(n)', () => {
    const { env } = makeEnv();
    inject(createWScript(env).object);
    expect(engine.eval('WScript.Arguments.Count')).toBe(2);
    expect(engine.eval('WScript.Arguments(0)')).toBe('apple');
    expect(engine.eval('WScript.Arguments.Item(1)')).toBe('banana');
  });

  it('Arguments.Named supports Item, default invocation, and Exists', () => {
    const { env } = makeEnv();
    inject(createWScript(env).object);
    expect(engine.eval('WScript.Arguments.Named("server")')).toBe('prod');
    expect(engine.eval('WScript.Arguments.Named.Exists("server")')).toBe(true);
    expect(engine.eval('WScript.Arguments.Named.Exists("missing")')).toBe(false);
  });

  it('Arguments.Unnamed is indexed positionally', () => {
    const { env } = makeEnv();
    inject(createWScript(env).object);
    expect(engine.eval('WScript.Arguments.Unnamed(0)')).toBe('apple');
  });

  it('StdOut.WriteLine writes to stdout; StdIn.ReadLine reads a line', () => {
    const { env, out, lines } = makeEnv();
    lines.push('first-line');
    inject(createWScript(env).object);
    engine.addCode('WScript.StdOut.WriteLine "hello"');
    expect(out.join('')).toBe('hello\n');
    expect(engine.eval('WScript.StdIn.ReadLine')).toBe('first-line');
  });

  it('Quit records the exit code and unwinds the script', () => {
    const { env, out } = makeEnv();
    const { object, getQuitCode } = createWScript(env);
    inject(object);
    engine.addCode('WScript.Echo "before"\nWScript.Quit 7\nWScript.Echo "after"');
    expect(out.join('')).toBe('before\n');
    expect(getQuitCode()).toBe(7);
  });

  it('Quit throws WScriptQuitSignal so the host can map it to an exit code', () => {
    const { env } = makeEnv();
    const { object } = createWScript(env);
    const quit = object.getMethod!('Quit');
    expect(() => quit.func({ type: 'Integer', value: 3 })).toThrowError(WScriptQuitSignal);
    try {
      quit.func({ type: 'Integer', value: 3 });
    } catch (e) {
      expect((e as WScriptQuitSignal).exitCode).toBe(3);
    }
  });

  it('Sleep calls the environment sleep with the requested duration', () => {
    let slept = -1;
    const { env } = makeEnv({ sleep: (ms: number) => (slept = ms) });
    inject(createWScript(env).object);
    engine.addCode('WScript.Sleep 1234');
    expect(slept).toBe(1234);
  });
});
