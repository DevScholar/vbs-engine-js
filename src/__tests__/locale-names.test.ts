import { describe, it, expect } from "vitest";
import { LCID_TO_LOCALE } from "../builtins/locale-data.ts";

// The LCIDs whose name the engine used to get wrong. Their old spellings resolved to nothing, so ICU
// fell back to the host locale for both of them - on a Chinese machine they were read as Chinese - and
// 2108 was read as Scottish Gaelic instead of Irish.
describe("The LCID names read from Windows and ICU", () => {
  it("names the LCIDs whose old spellings resolved to nothing", () => {
    expect(LCID_TO_LOCALE.get(1070)).toBe("hsb-de"); // Upper Sorbian; "sb" resolved to nothing
    expect(LCID_TO_LOCALE.get(1072)).toBe("st-za"); // Southern Sotho; "sx" resolved to nothing
    expect(LCID_TO_LOCALE.get(2108)).toBe("ga-ie"); // Irish; "gd-ie" resolved to Scottish Gaelic
  });

  it("uses names this ICU resolves exactly, apart from the documented exceptions", () => {
    // Kept because ICU cannot resolve them exactly: es-ES_tradnl throws, en-029 and the sr-*-CS pair
    // fall back to a relative, and nb-NO, ts-ZA, yi-001 and 1084 have no Windows name to use.
    const documented = new Set([
      1034, 1044, 1073, 1084, 1085, 2074, 3098, 9225,
    ]);

    for (const [lcid, name] of LCID_TO_LOCALE) {
      if (documented.has(lcid)) continue;
      const resolved = new Intl.DateTimeFormat(name).resolvedOptions().locale;
      expect(
        resolved.toLowerCase().replace(/_/g, "-"),
        `LCID ${lcid} (${name})`,
      ).toBe(name);
    }
  });

  it("covers every LCID the engine knows", () => {
    expect(LCID_TO_LOCALE.size).toBe(133);
  });
});
