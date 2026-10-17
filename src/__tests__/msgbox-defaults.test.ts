import { describe, it, expect } from "vitest";
import { getDefaultResult, createMsgBox } from "../builtins/msgbox.ts";
import type { VbValue } from "../runtime/index.ts";

describe("MsgBox default-button semantics", () => {
  const cases: [string, number, number][] = [
    ["OK", 0, 1],
    ["OKCancel", 0, 1],
    ["OKCancel", 256, 2],
    ["YesNo", 0, 6],
    ["YesNo", 256, 7],
    ["YesNoCancel", 0, 6],
    ["YesNoCancel", 256, 7],
    ["YesNoCancel", 512, 2],
    ["RetryCancel", 0, 4],
    ["RetryCancel", 256, 2],
    ["AbortRetryIgnore", 0, 3],
    ["AbortRetryIgnore", 256, 4],
    ["AbortRetryIgnore", 512, 5],
  ];
  it.each(cases)("%s defaultButton=%d -> %d", (bt, db, exp) => {
    expect(getDefaultResult(bt as any, db)).toBe(exp);
  });

  it("OK is not affected by defaultButton bits (only one button)", () => {
    expect(getDefaultResult("OK", 512)).toBe(1);
  });
});

describe("MsgBox gui hook precedence", () => {
  it("calls gui and returns its result", () => {
    const mb = createMsgBox({ gui: () => 6 });
    const r = mb(
      { type: "String", value: "q" } as VbValue,
      { type: "Integer", value: 4 } as VbValue,
      { type: "String", value: "T" } as VbValue,
    );
    expect(r.value).toBe(6);
  });
});
