/**
 * VBScript Engine for JavaScript - A VBScript interpreter written in TypeScript.
 *
 * This module provides a complete VBScript engine that can execute VBScript code
 * in both Node.js and browser environments.
 *
 * @packageDocumentation
 *
 * @example
 * ```typescript
 * // Node.js usage
 * import { VbsEngine, runVbscript } from 'vbs-engine-js';
 *
 * const engine = new VbsEngine();
 * engine.run('x = 5 + 3');
 * console.log(engine.getVariableAsJs('x')); // 8
 *
 * // Or use the convenience function
 * const result = runVbscript('MsgBox "Hello"');
 * ```
 *
 * @example
 * ```html
 * <!-- Browser usage -->
 * <script type="module">
 *   import { VbsBrowserEngine } from 'vbs-engine-js';
 *   new VbsBrowserEngine();
 * </script>
 * <script type="text/vbscript">
 *   MsgBox "Hello from VBScript!"
 * </script>
 * ```
 */
export { VbsEngine, runVbscript, type VbsEngineOptions, jsToVb, vbToJs } from './core/index.ts';
export { VbsBrowserEngine, createBrowserRuntime, type BrowserRuntimeOptions } from './browser/index.ts';
