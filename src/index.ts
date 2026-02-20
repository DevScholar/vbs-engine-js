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
 * // Node.js usage (general mode)
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
 *   import { VbsEngine } from 'vbs-engine-js';
 *   // Use browser mode for automatic DOM integration
 *   new VbsEngine({ mode: 'browser' });
 * </script>
 * <script type="text/vbscript">
 *   MsgBox "Hello from VBScript!"
 * </script>
 * ```
 *
 * @example
 * ```typescript
 * // Browser mode with options
 * const engine = new VbsEngine({
 *   mode: 'browser',
 *   parseScriptElement: true,
 *   parseInlineEventAttributes: true,
 *   parseEventSubNames: true
 * });
 * ```
 */
export {
  VbsEngine,
  runVbscript,
  type VbsEngineOptions,
  type VbsEngineMode,
  type BrowserEngineOptions,
  jsToVb,
  vbToJs
} from './core/index.ts';

// Browser-specific exports
export type { BrowserRuntimeOptions } from './browser/index.ts';
