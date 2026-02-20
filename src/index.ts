/**
 * VBScript Engine for JavaScript - A VBScript interpreter written in TypeScript.
 *
 * This module provides a complete VBScript engine that can execute VBScript code
 * in both Node.js and browser environments.
 *
 * The API is designed to be compatible with MSScriptControl.ScriptControl:
 * - `addCode(code)` - Add script code (functions, classes)
 * - `executeStatement(statement)` - Execute a single statement
 * - `run(procedureName, ...args)` - Call a script function
 * - `addObject(name, object, addMembers?)` - Expose a JS object to script
 * - `eval(expression)` - Evaluate an expression
 *
 * @packageDocumentation
 *
 * @example
 * ```typescript
 * // Node.js usage (general mode)
 * import { VbsEngine, evalVbscript } from 'vbs-engine-js';
 *
 * const engine = new VbsEngine();
 *
 * // Add code (function definitions)
 * engine.addCode(`
 *   Function Add(a, b)
 *     Add = a + b
 *   End Function
 * `);
 *
 * // Call the function
 * const result = engine.run('Add', 5, 3);  // 8
 *
 * // Execute a statement
 * engine.executeStatement('MsgBox "Hello"');
 *
 * // Evaluate an expression
 * const value = engine.eval('2 + 3 * 4');  // 14
 *
 * // Expose an object to script
 * engine.addObject('console', console, true);
 * engine.executeStatement('console.log "Hello from VBScript"');
 *
 * // Or use the convenience function
 * const result = evalVbscript('2 + 3 * 4');  // 14
 * ```
 *
 * @example
 * ```html
 * <!-- Browser usage -->
 * <script type="module">
 *   import { VbsEngine } from 'vbs-engine-js';
 *   // Use browser mode for automatic DOM integration
 *   new VbsEngine({ mode: 'browser' });
 * </script>
 * <script type="text/vbscript">
 *   MsgBox "Hello from VBScript!"
 * </script>
 * ```
 */
export {
  VbsEngine,
  evalVbscript,
  type VbsEngineOptions,
  type VbsEngineMode,
  type BrowserEngineOptions,
  type VbsError,
  jsToVb,
  vbToJs
} from './core/index.ts';

// Browser-specific exports
export type { BrowserRuntimeOptions } from './browser/index.ts';
