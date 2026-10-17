import type { VbValue, VbObjectValueData } from "../runtime/values.ts";
import { VbEmpty } from "../runtime/values.ts";
import { createVbError, VbErrorCodes } from "../runtime/errors.ts";

/**
 * The WSH-compatible `WScript` host object, provided by the host (the run-script
 * CLI) rather than by the engine itself — mirroring how `ActiveXObject`/`GetObject`
 * are supplied by the host. When the host injects a `WScript` global, VBScript
 * scripts written for cscript/wscript run unchanged; when it is absent, the
 * engine behaves exactly as before.
 *
 * Real WSH's `WScript` is a COM/IDispatch object, so a script calls it identically
 * from VBScript and JScript. This module reproduces that dual surface: every host
 * object implements the engine's `VbObjectValueData` protocol (`getProperty` /
 * `hasMethod` / `getMethod`) for VBScript dispatch, *and* carries the same members
 * as plain JS functions/properties for JScript (which reads them directly off
 * `globalThis`). Collections that are called by default member (`WScript.Arguments(0)`,
 * `WScript.Arguments.Named("x")`) are wrapped in a callable proxy that also hides
 * `call`/`apply` from the engine so its default-property dispatch still fires.
 *
 * `createWScript` is pure and testable: every environment-specific behavior (stdin
 * read, stdout/stderr write, sleep, host paths, exit) is supplied through
 * {@link WshEnvironment}. Only `WScript.Quit` has a side effect by contract: it
 * throws {@link WScriptQuitSignal} to unwind the interpreter, and records the exit
 * code for the host to pick up via {@link WScriptHandle.getQuitCode}.
 */

/** Thrown by `WScript.Quit` to terminate script execution. */
export class WScriptQuitSignal extends Error {
  public constructor(public readonly exitCode: number) {
    super(`WScript.Quit(${exitCode})`);
    this.name = "WScriptQuitSignal";
  }
}

/** Host-supplied behavior for the WScript object (I/O, paths, sleep, args). */
export interface WshEnvironment {
  /** 'cscript' (console) or 'wscript' (GUI). Drives `WScript.Echo` and `FullName`. */
  mode: "cscript" | "wscript";
  scriptFullName: string;
  scriptName: string;
  interactive: boolean;
  /** Positional command-line arguments (index 0 = first arg). */
  unnamedArgs: string[];
  /** Named arguments keyed by name (lowercase), from `/name:value` or a `.wsf` <runtime>. */
  namedArgs: Record<string, string>;
  /** Usage text emitted by `WScript.Arguments.ShowUsage()` / generated from `.wsf`. */
  usage: string;
  writeStdout(text: string): void;
  writeStderr(text: string): void;
  /** Synchronous line read; returns null on EOF. */
  readLine(): string | null;
  /** Synchronous sleep (milliseconds). */
  sleep(ms: number): void;
  /**
   * Optional GUI dialog hook for `wscript`-mode Echo. When present, Echo routes
   * through it (a blocking dialog); when absent, Echo falls back to the console.
   */
  showDialog?: (message: string, title?: string) => void;
  /** Full path to the host executable (e.g. the run-script CLI). */
  fullName: string;
  /** Directory of the host executable. */
  path: string;
  /** WSH version string, e.g. "5.812". */
  version: string;
  buildVersion: number;
  currentDirectory: string;
}

/** Handle returned by {@link createWScript} — the object plus its Quit state. */
export interface WScriptHandle {
  /** The object to inject as the `WScript` global. */
  object: VbObjectValueData;
  /** The exit code set by `WScript.Quit`, or null if Quit was never called. */
  getQuitCode(): number | null;
}

/** VBScript-facing protocol shape shared by every host object below. */
interface HostObject extends VbObjectValueData {
  getProperty(name: string): VbValue;
  hasMethod(name: string): boolean;
  getMethod(name: string): { func: (...a: VbValue[]) => VbValue };
}

/**
 * Makes an object callable from JS (for JScript default-member calls like
 * `Arguments(0)`) while hiding `call`/`apply`/`bind` so the engine's own
 * default-property dispatch (which checks `typeof obj.call === 'function'`) does
 * not mistake it for a raw JS function. A Proxy is only callable when its target
 * is a function, so the target is a function; the `get` trap forwards property
 * reads to `target` (the HostObject), so all members (both the VBScript protocol
 * and the JS-facing ones) are still reachable. `invoke` receives the JS arg list.
 */
function callableObject<T extends object>(
  target: T,
  invoke: (args: unknown[]) => unknown,
): T {
  const fn = function (...args: unknown[]): unknown {
    return invoke(args);
  };
  const targetRecord = target as unknown as Record<string | symbol, unknown>;
  return new Proxy(fn, {
    get(_t: typeof fn, prop: string | symbol): unknown {
      if (prop === "call" || prop === "apply" || prop === "bind")
        return undefined;
      return targetRecord[prop];
    },
    set(_t: typeof fn, prop: string | symbol, value: unknown): boolean {
      targetRecord[prop] = value;
      return true;
    },
    // Forward descriptor/introspection traps so `hasOwnProperty(proxy, 'getProperty')`
    // still sees the target's protocol members as own properties — the engine relies
    // on that check to route VBScript member reads through getProperty/getMethod.
    has(_t: typeof fn, prop: string | symbol): boolean {
      void _t;
      return prop in targetRecord;
    },
    getOwnPropertyDescriptor(
      _t: typeof fn,
      prop: string | symbol,
    ): PropertyDescriptor | undefined {
      if (prop === "call" || prop === "apply" || prop === "bind")
        return undefined;
      const d = Object.getOwnPropertyDescriptor(targetRecord, prop);
      return d;
    },
    ownKeys(): ArrayLike<string | symbol> {
      return Reflect.ownKeys(targetRecord);
    },
  }) as unknown as T;
}

function makeItemMethod(items: string[]): (index: VbValue) => VbValue {
  return (index: VbValue): VbValue => {
    // A collection called without its argument is a wrong number of arguments rather than a JavaScript
    // failure, which `WScript.Arguments.Item` reaches: a fresh cscript answers 450 for that both as a
    // statement and inside an expression.
    if (index === undefined) {
      throw createVbError(
        VbErrorCodes.InvalidQualifier,
        "Wrong number of arguments or invalid property assignment: 'Item'",
      );
    }
    const n = Math.floor(Number(index.value ?? index));
    const value = items[n];
    return value === undefined ? VbEmpty : { type: "String", value };
  };
}

/** Builds a callable positional (`WshUnnamed`) collection, dual-facing. */
function createUnnamed(args: string[]): HostObject {
  const item = (index: number): string | undefined => args[index];

  const protocol: HostObject = {
    classInfo: { name: "WshUnnamed" },
    getProperty(name: string): VbValue {
      const lower = name.toLowerCase();
      if (lower === "count" || lower === "length") {
        return { type: "Long", value: args.length };
      }
      const num = parseInt(name, 10);
      if (
        !isNaN(num) &&
        String(num) === name &&
        num >= 0 &&
        num < args.length
      ) {
        return { type: "String", value: args[num] };
      }
      return VbEmpty;
    },
    hasMethod(name: string): boolean {
      const lower = name.toLowerCase();
      return lower === "item" || lower === "default";
    },
    getMethod(name: string): { func: (...a: VbValue[]) => VbValue } {
      const lower = name.toLowerCase();
      if (lower === "item" || lower === "default") {
        return { func: makeItemMethod(args) };
      }
      throw new Error(`WshUnnamed: unknown method '${name}'`);
    },
  };

  // JS-facing members for JScript.
  const js = {
    Count: args.length,
    length: args.length,
    Item: (index: number): string | undefined => item(Math.floor(index)),
  };

  const merged = Object.assign(protocol, js) as HostObject;
  return callableObject(merged, (a) => item(Math.floor(Number(a[0] ?? 0))));
}

/** Builds a callable named (`WshNamed`) collection, dual-facing. */
function createNamed(named: Record<string, string>): HostObject {
  const item = (key: string): string | undefined => named[key.toLowerCase()];
  const exists = (key: string): boolean =>
    Object.prototype.hasOwnProperty.call(named, key.toLowerCase());

  const protocol: HostObject = {
    classInfo: { name: "WshNamed" },
    getProperty(name: string): VbValue {
      const lower = name.toLowerCase();
      if (lower === "count" || lower === "length") {
        return { type: "Long", value: Object.keys(named).length };
      }
      return VbEmpty;
    },
    hasMethod(name: string): boolean {
      const lower = name.toLowerCase();
      return lower === "item" || lower === "default" || lower === "exists";
    },
    getMethod(name: string): { func: (...a: VbValue[]) => VbValue } {
      const lower = name.toLowerCase();
      if (lower === "exists") {
        return {
          func: (key: VbValue): VbValue => ({
            type: "Boolean",
            value: exists(String(key.value ?? key)),
          }),
        };
      }
      if (lower === "item" || lower === "default") {
        return {
          func: (key: VbValue): VbValue => {
            const value = item(String(key.value ?? key));
            return value === undefined ? VbEmpty : { type: "String", value };
          },
        };
      }
      throw new Error(`WshNamed: unknown method '${name}'`);
    },
  };

  const js = {
    Count: Object.keys(named).length,
    length: Object.keys(named).length,
    Item: item,
    Exists: exists,
  };

  const merged = Object.assign(protocol, js) as HostObject;
  return callableObject(merged, (a) => item(String(a[0] ?? "")));
}

/** Builds the `WScript.Arguments` collection, dual-facing. */
function createArguments(env: WshEnvironment): HostObject {
  const unnamed = createUnnamed(env.unnamedArgs);
  const named = createNamed(env.namedArgs);
  const item = (index: number): string | undefined => env.unnamedArgs[index];

  const protocol: HostObject = {
    classInfo: { name: "WshArguments" },
    getProperty(name: string): VbValue {
      const lower = name.toLowerCase();
      if (lower === "count" || lower === "length") {
        return { type: "Long", value: env.unnamedArgs.length };
      }
      if (lower === "named") {
        return { type: "Object", value: named };
      }
      if (lower === "unnamed") {
        return { type: "Object", value: unnamed };
      }
      const num = parseInt(name, 10);
      if (
        !isNaN(num) &&
        String(num) === name &&
        num >= 0 &&
        num < env.unnamedArgs.length
      ) {
        return { type: "String", value: env.unnamedArgs[num] };
      }
      return VbEmpty;
    },
    hasMethod(name: string): boolean {
      const lower = name.toLowerCase();
      return lower === "item" || lower === "default" || lower === "showusage";
    },
    getMethod(name: string): { func: (...a: VbValue[]) => VbValue } {
      const lower = name.toLowerCase();
      if (lower === "item" || lower === "default") {
        return { func: makeItemMethod(env.unnamedArgs) };
      }
      if (lower === "showusage") {
        return {
          func: (): VbValue => {
            env.writeStdout(env.usage);
            return VbEmpty;
          },
        };
      }
      throw new Error(`WshArguments: unknown method '${name}'`);
    },
  };

  const js = {
    Count: env.unnamedArgs.length,
    length: env.unnamedArgs.length,
    Named: named,
    Unnamed: unnamed,
    Item: (index: number): string | undefined => item(Math.floor(index)),
    ShowUsage: (): void => env.writeStdout(env.usage),
  };

  const merged = Object.assign(protocol, js) as HostObject;
  return callableObject(merged, (a) => item(Math.floor(Number(a[0] ?? 0))));
}

/** Builds a `TextStream`-shaped object for StdIn/StdOut/StdErr, dual-facing. */
function createTextStream(
  kind: "stdin" | "stdout" | "stderr",
  env: WshEnvironment,
): HostObject {
  const write = kind === "stderr" ? env.writeStderr : env.writeStdout;
  const writable = kind === "stdout" || kind === "stderr";
  const readable = kind === "stdin";

  const readLine = (): VbValue => {
    const line = env.readLine();
    return line === null ? VbEmpty : { type: "String", value: line };
  };

  const protocol: HostObject = {
    classInfo: { name: "TextStream" },
    getProperty(name: string): VbValue {
      const lower = name.toLowerCase();
      if (lower === "atendofline" || lower === "atendofstream") {
        return { type: "Boolean", value: false };
      }
      if (lower === "line" || lower === "column") {
        return { type: "Long", value: 1 };
      }
      return VbEmpty;
    },
    hasMethod(name: string): boolean {
      const lower = name.toLowerCase();
      if (readable) {
        return [
          "read",
          "readline",
          "readall",
          "skip",
          "skipline",
          "close",
        ].includes(lower);
      }
      if (writable) {
        return ["write", "writeline", "writeblanklines", "close"].includes(
          lower,
        );
      }
      return false;
    },
    getMethod(name: string): { func: (...a: VbValue[]) => VbValue } {
      const lower = name.toLowerCase();
      if (readable) {
        if (lower === "readline") return { func: (): VbValue => readLine() };
        if (lower === "readall") {
          return {
            func: (): VbValue => {
              let result = "";
              for (;;) {
                const line = env.readLine();
                if (line === null) break;
                result += line + "\n";
              }
              return { type: "String", value: result };
            },
          };
        }
        if (
          lower === "skip" ||
          lower === "skipline" ||
          lower === "read" ||
          lower === "close"
        ) {
          return { func: (): VbValue => VbEmpty };
        }
      }
      if (writable) {
        if (lower === "write") {
          return {
            func: (text: VbValue): VbValue => {
              write(String(text.value ?? text));
              return VbEmpty;
            },
          };
        }
        if (lower === "writeline") {
          return {
            func: (text?: VbValue): VbValue => {
              write(String(text?.value ?? text ?? "") + "\n");
              return VbEmpty;
            },
          };
        }
        if (lower === "writeblanklines") {
          return {
            func: (n: VbValue): VbValue => {
              write("\n".repeat(Math.max(0, Math.floor(Number(n.value ?? n)))));
              return VbEmpty;
            },
          };
        }
        if (lower === "close") {
          return { func: (): VbValue => VbEmpty };
        }
      }
      throw new Error(`TextStream: unknown method '${name}'`);
    },
  };

  const js: Record<string, unknown> = {
    AtEndOfLine: false,
    AtEndOfStream: false,
    Line: 1,
    Column: 1,
    Close: (): void => {},
  };
  if (readable) {
    js.ReadLine = (): string | null => env.readLine();
    js.ReadAll = (): string => {
      let result = "";
      for (;;) {
        const line = env.readLine();
        if (line === null) break;
        result += line + "\n";
      }
      return result;
    };
    js.Read = (): string => "";
    js.Skip = (): void => {};
    js.SkipLine = (): void => {};
  }
  if (writable) {
    js.Write = (text: string): void => write(String(text));
    js.WriteLine = (text?: string): void => write(String(text ?? "") + "\n");
    js.WriteBlankLines = (n: number): void =>
      write("\n".repeat(Math.max(0, Math.floor(Number(n ?? 0)))));
  }

  return Object.assign(protocol, js) as HostObject;
}

/**
 * Creates a WSH-compatible `WScript` host object, dual-facing for VBScript and
 * JScript. Inject the returned `object` as `globalThis.WScript`.
 */
export function createWScript(env: WshEnvironment): WScriptHandle {
  let quitCode: number | null = null;

  const argumentsObj = createArguments(env);
  const stdIn = createTextStream("stdin", env);
  const stdOut = createTextStream("stdout", env);
  const stdErr = createTextStream("stderr", env);

  const echo = (mode: "cscript" | "wscript", args: unknown[]): void => {
    const text = args
      .map((a) => String((a as VbValue | undefined)?.value ?? a))
      .join(mode === "cscript" ? " " : "\n");
    // wscript mode surfaces Echo as a dialog when a GUI backend is available;
    // otherwise it degrades to the console, same as the real host.
    if (mode === "wscript" && env.showDialog) {
      env.showDialog(text, env.scriptName);
      return;
    }
    env.writeStdout(text + "\n");
  };

  const quit = (code?: VbValue): never => {
    const n = code === undefined ? 0 : Math.floor(Number(code.value ?? code));
    quitCode = n;
    throw new WScriptQuitSignal(n);
  };

  const createOrGetObject = (
    which: "createobject" | "getobject",
    args: VbValue[],
  ): VbValue => {
    const g = (globalThis as Record<string, unknown>)[
      which === "createobject" ? "ActiveXObject" : "GetObject"
    ];
    if (typeof g === "function") {
      const jsArgs = args.map((a) => a.value ?? a);
      return {
        type: "Object",
        value: (g as (...a: unknown[]) => unknown)(...jsArgs),
      };
    }
    return { type: "Object", value: null };
  };

  const protocol: HostObject = {
    classInfo: { name: "WScript" },

    getProperty(name: string): VbValue {
      const lower = name.toLowerCase();
      switch (lower) {
        case "arguments":
          return { type: "Object", value: argumentsObj };
        case "stdin":
          return { type: "Object", value: stdIn };
        case "stdout":
          return { type: "Object", value: stdOut };
        case "stderr":
          return { type: "Object", value: stdErr };
        case "scriptname":
          return { type: "String", value: env.scriptName };
        case "scriptfullname":
          return { type: "String", value: env.scriptFullName };
        case "fullname":
          return { type: "String", value: env.fullName };
        case "path":
          return { type: "String", value: env.path };
        case "version":
          return { type: "String", value: env.version };
        case "buildversion":
          return { type: "Long", value: env.buildVersion };
        case "interactive":
          return { type: "Boolean", value: env.interactive };
        case "currentdirectory":
          return { type: "String", value: env.currentDirectory };
        default:
          return VbEmpty;
      }
    },

    setProperty(name: string, value: VbValue): void {
      if (name.toLowerCase() === "currentdirectory") {
        void value; // read-only in practice; assignment is accepted but a no-op
      }
    },

    hasMethod(name: string): boolean {
      const lower = name.toLowerCase();
      return (
        lower === "echo" ||
        lower === "quit" ||
        lower === "sleep" ||
        lower === "createobject" ||
        lower === "getobject" ||
        lower === "connectobject" ||
        lower === "disconnectobject"
      );
    },

    getMethod(name: string): { func: (...a: VbValue[]) => VbValue } {
      const lower = name.toLowerCase();
      if (lower === "echo") {
        return {
          func: (...args: VbValue[]): VbValue => {
            echo(env.mode, args);
            return VbEmpty;
          },
        };
      }
      if (lower === "quit") {
        return { func: (code?: VbValue): VbValue => quit(code) };
      }
      if (lower === "sleep") {
        return {
          func: (ms: VbValue): VbValue => {
            env.sleep(Math.max(0, Math.floor(Number(ms.value ?? ms))));
            return VbEmpty;
          },
        };
      }
      if (lower === "createobject" || lower === "getobject") {
        return {
          func: (...args: VbValue[]): VbValue =>
            createOrGetObject(lower as "createobject" | "getobject", args),
        };
      }
      if (lower === "connectobject" || lower === "disconnectobject") {
        return { func: (): VbValue => VbEmpty };
      }
      throw new Error(`WScript: unknown method '${name}'`);
    },
  };

  // JS-facing members for JScript.
  const js: Record<string, unknown> = {
    Arguments: argumentsObj,
    StdIn: stdIn,
    StdOut: stdOut,
    StdErr: stdErr,
    ScriptName: env.scriptName,
    ScriptFullName: env.scriptFullName,
    FullName: env.fullName,
    Path: env.path,
    Version: env.version,
    BuildVersion: env.buildVersion,
    Interactive: env.interactive,
    CurrentDirectory: env.currentDirectory,
    Echo: (...a: unknown[]): void => echo(env.mode, a),
    Quit: (code?: number): void =>
      quit(code === undefined ? undefined : { type: "Integer", value: code }),
    Sleep: (ms: number): void => env.sleep(Math.max(0, Math.floor(Number(ms)))),
    CreateObject: (progid: string): unknown => {
      const g = (globalThis as Record<string, unknown>).ActiveXObject;
      return typeof g === "function"
        ? (g as (...a: unknown[]) => unknown)(progid)
        : null;
    },
    GetObject: (pathname?: string, cls?: string): unknown => {
      const g = (globalThis as Record<string, unknown>).GetObject;
      return typeof g === "function"
        ? (g as (...a: unknown[]) => unknown)(pathname, cls)
        : null;
    },
    ConnectObject: (): void => {},
    DisconnectObject: (): void => {},
  };

  const object = Object.assign(protocol, js) as HostObject;

  return {
    object,
    getQuitCode: () => quitCode,
  };
}
