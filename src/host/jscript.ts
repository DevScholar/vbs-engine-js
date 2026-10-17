/**
 * A minimal JScript host for WSH-style mixed-language scripts.
 *
 * JScript is (historically) ECMAScript 3, and every environment this engine runs
 * in already has a real JavaScript engine (Node/V8). So rather than reimplement a
 * second interpreter, this runner executes a JScript block with the host's own
 * engine, in the *global* context, and bridges the two languages through the
 * shared `globalThis`:
 *   - JScript can call VBScript functions: the engine exposes user-defined
 *     functions on `globalThis` via `syncFunctionsToGlobalThis`.
 *   - VBScript can read JScript globals: `VbContext.getVariable` falls back to
 *     `globalThis` lookup.
 *
 * Running in the global context (via `vm.runInThisContext`) means top-level `var`
 * and `function` declarations in a JScript block become real `globalThis`
 * properties — the same shared-global behavior real WSH gives mixed-language
 * scripts. `WScript`, `ActiveXObject`, `GetObject`, and `Enumerator` are all
 * resolved from `globalThis`, so no special scope wiring is needed here.
 */

import vm from "node:vm";

/**
 * Runs a JScript source block in the host's JS engine, in the global context.
 *
 * @param code - JScript source.
 * @returns The completion value of the block, if any.
 */
export function runJScript(code: string): unknown {
  return vm.runInThisContext(code, { filename: "jscript.wsf" });
}
