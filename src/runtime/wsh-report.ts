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
 * The column this reads is the engine's own, which is counted from zero to match Microsoft's `ScriptControl`,
 * whose `Error` object the engine's `error` member mirrors. Both `cscript` and Internet Explorer print a column
 * counted from one, so the line this function writes adds one to it. Measured: with the failing statement at
 * column 1, 3, 5 and 9 of its line, a `ScriptControl` reports 0, 2, 4 and 8 where `cscript` reports 1, 3, 5 and 9,
 * and both count the line from one.
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
  /** The column in that line, counted from zero, as the engine's own `error` member carries it. */
  column?: number;
}

/**
 * The one line Windows Script Host would print for a failure, or the same line without the script's name when the
 * host never named one.
 */
export function formatWshErrorLine(
  failure: WshFailure,
  scriptName?: string | null,
): string {
  const described =
    failure.source === undefined || failure.source === ""
      ? failure.description
      : `${failure.source}: ${failure.description}`;

  // A failure with no line of its own, which the engine can raise before any script ran, has no position to show.
  if (typeof failure.line !== "number") return described;

  // The engine counts the column from zero the way `ScriptControl` does, and every host that prints counts from
  // one, so the number this writes is the engine's own plus one.
  const column = (typeof failure.column === "number" ? failure.column : 0) + 1;
  return `${scriptName ?? ""}(${failure.line}, ${column}) ${described}`;
}
