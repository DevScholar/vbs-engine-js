import { describe, it, expect } from "vitest";
import { formatWshErrorLine } from "../runtime/wsh-report.ts";

/**
 * Windows Script Host prints one line for a failure and then stops. Measured on Windows 11 with `cscript`, with
 * the failing statement on line 4 of the file, in a console that speaks Chinese:
 *
 *   C:\Users\UserName\AppData\Local\Temp\dsh-de-probe\wsh\divide.vbs(4, 1) Microsoft VBScript 运行时错误: 被零除
 *
 * The engine writes that line in English, as it writes every other message, and without the name when the host
 * never gave one. The column it is given is the engine's own, counted from zero the way Microsoft's ScriptControl
 * counts it, and the line it writes adds one, the way both hosts that print count.
 */
describe("the line Windows Script Host prints for a failure", () => {
  it("names the script, the position, and what raised the failure", () => {
    expect(
      formatWshErrorLine(
        {
          description: "Division by zero",
          source: "Microsoft VBScript runtime error",
          line: 4,
          column: 0,
        },
        "C:\\work\\divide.vbs",
      ),
    ).toBe(
      "C:\\work\\divide.vbs(4, 1) Microsoft VBScript runtime error: Division by zero",
    );
  });

  // The source is the script's own when it raised the failure itself, which is what `Err.Raise 4321, "a source",
  // "a description"` printed: the script said what raised it, so the engine does not say it for them.
  it("keeps the source a script gave to Err.Raise", () => {
    expect(
      formatWshErrorLine(
        {
          description: "a description",
          source: "a source",
          line: 2,
          column: 0,
        },
        "C:\\work\\raised.vbs",
      ),
    ).toBe("C:\\work\\raised.vbs(2, 1) a source: a description");
  });

  it("leaves the name out when the host never gave one", () => {
    expect(
      formatWshErrorLine({
        description: "Division by zero",
        source: "Microsoft VBScript runtime error",
        line: 4,
        column: 0,
      }),
    ).toBe("(4, 1) Microsoft VBScript runtime error: Division by zero");
  });

  it("shows no position for a failure that has none", () => {
    expect(
      formatWshErrorLine(
        { description: "Something went wrong" },
        "C:\\work\\demo.vbs",
      ),
    ).toBe("Something went wrong");
  });

  // Measured with a script whose statement stood at the fifth column of its line: `cscript` printed column 5,
  // and the engine holds that as four.
  it("counts the column one greater than the engine does", () => {
    expect(
      formatWshErrorLine(
        { description: "Division by zero", line: 4, column: 4 },
        "demo.vbs",
      ),
    ).toBe("demo.vbs(4, 5) Division by zero");
    expect(
      formatWshErrorLine(
        { description: "Division by zero", line: 4 },
        "demo.vbs",
      ),
    ).toBe("demo.vbs(4, 1) Division by zero");
  });
});
