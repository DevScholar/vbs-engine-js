/**
 * How a failure reads in the general environment, which stands for Windows Script Host rather than for a browser.
 *
 * `cscript` and `wscript` print one line and then stop. Measured on Windows 11, with the failing statement on
 * line 4 of the file, in a console that speaks Chinese:
 *
 *   C:\work\demo.vbs(4, 1) Microsoft VBScript 运行时错误: 被零除
 *
 * Read off several scripts, the parts are:
 *
 * * The name is the script file, made absolute even when the script was named relative to the current directory.
 *   A block taken out of a `.wsf` is reported under the `.wsf`'s own name.
 * * The line is the line in that file, counted from 1, and the column is the column in that line.
 * * The source is what raised it: `Microsoft VBScript runtime error` for a failure of the engine,
 *   `Microsoft VBScript compilation error` for a parse failure, and the source text a script gave to `Err.Raise`
 *   when it raised one itself.
 *
 * A host that never told the engine what its script is named gets the same line without that part, which is the
 * one piece the engine cannot work out for itself.
 *
 * @internal
 */

/** A failure, as far as a report about it goes. */
export interface WshFailure {
  description: string;
  /** What raised the failure, when anything did. */
  source?: string;
  /** The line in the script, when the host said where the script's code began. */
  line?: number;
  column?: number;
}

/**
 * The one line Windows Script Host would print for a failure, or the same line without the script's name when the
 * host never named one.
 */
export function formatWshErrorLine(failure: WshFailure, scriptName?: string | null): string {
  const described =
    failure.source === undefined || failure.source === ''
      ? failure.description
      : `${failure.source}: ${failure.description}`;

  // A failure with no line of its own, which the engine can raise before any script ran, has no position to show.
  if (typeof failure.line !== 'number') return described;

  const column = typeof failure.column === 'number' ? failure.column : 1;
  return `${scriptName ?? ''}(${failure.line}, ${column}) ${described}`;
}
