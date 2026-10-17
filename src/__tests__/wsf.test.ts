import { describe, it, expect } from "vitest";
import { findBlockStart, parseWsf } from "../host/wsf.ts";

describe("WSF parser", () => {
  it("parses a single <job> without a <package> wrapper", () => {
    const xml = `<job id="J1"><script language="VBScript">WScript.Echo 1</script></job>`;
    const doc = parseWsf(xml);
    expect(doc.jobs).toHaveLength(1);
    expect(doc.jobs[0]!.id).toBe("J1");
    expect(doc.jobs[0]!.scripts[0]!.code).toBe("WScript.Echo 1");
  });

  it("parses multiple jobs under a <package>", () => {
    const xml = `<package>
      <job id="A"><script language="VBScript">a</script></job>
      <job id="B"><script language="JScript">b</script></job>
    </package>`;
    const doc = parseWsf(xml);
    expect(doc.jobs).toHaveLength(2);
    expect(doc.jobs.map((j) => j.id)).toEqual(["A", "B"]);
    expect(doc.jobs[1]!.scripts[0]!.language).toBe("JScript");
  });

  it("extracts CDATA-wrapped script source verbatim", () => {
    const xml = `<?XML version="1.0"?><job id="J"><script language="VBScript"><![CDATA[
WScript.Echo "x"
]]></script></job>`;
    const doc = parseWsf(xml);
    expect(doc.jobs[0]!.scripts[0]!.code).toContain('WScript.Echo "x"');
  });

  it("parses <runtime> named/unnamed params and description/example", () => {
    const xml = `<job id="J">
      <runtime>
        <description>Demo.</description>
        <named name="Server" helpstring="Server name" type="string" required="false"/>
        <unnamed name="Path" helpstring="File path"/>
        <example>Example: x.wsf /Server:prod</example>
      </runtime>
      <script language="VBScript">x</script>
    </job>`;
    const job = parseWsf(xml).jobs[0]!;
    expect(job.description).toBe("Demo.");
    expect(job.example).toBe("Example: x.wsf /Server:prod");
    expect(job.named).toHaveLength(1);
    expect(job.named[0]!.name).toBe("Server");
    expect(job.named[0]!.required).toBe(false);
    expect(job.unnamed).toHaveLength(1);
    expect(job.unnamed[0]!.name).toBe("Path");
  });

  it("parses <object>/<reference>/<resource> blocks", () => {
    const xml = `<job id="J">
      <object id="fso" progid="Scripting.FileSystemObject"/>
      <reference object="Scripting.FileSystemObject"/>
      <resource id="text">hello</resource>
      <script language="VBScript">x</script>
    </job>`;
    const job = parseWsf(xml).jobs[0]!;
    expect(job.objects[0]!.progid).toBe("Scripting.FileSystemObject");
    expect(job.references[0]!.object).toBe("Scripting.FileSystemObject");
    expect(job.resources["text"]).toBe("hello");
  });

  it("rejects a document with no job", () => {
    expect(() => parseWsf("<package></package>")).toThrow();
  });

  // What stands between the tags is the script, and a CDATA section is one of the two ways to write it. A block
  // that puts its CDATA section on a line of its own carries whitespace in front of it, and taking only that
  // whitespace as the code leaves the block empty, which is what a script that then never runs looks like.
  it("takes the whitespace in front of a CDATA section as part of the block", () => {
    const xml = `<package>
  <job id="Main">
    <script language="VBScript">
    <![CDATA[
        Dim answer
        answer = 1 / 0
    ]]>
    </script>
  </job>
</package>`;
    const code = parseWsf(xml).jobs[0]!.scripts[0]!.code!;
    expect(code).toContain("answer = 1 / 0");
    expect(code.split("\n")[0]).toBe("");
  });
});

/**
 * Where a block stands in the `.wsf` file, which is what a host has to tell the engine for a failure to be
 * reported at the file's own line. Measured with `cscript`: a failure on wsf line 6 of a block beginning on wsf
 * line 3 came back as line 6, and one on wsf line 9 of a block written as CDATA beginning on wsf line 4 came back
 * as line 9.
 */
describe("finding where a block stands in a .wsf", () => {
  it("counts the block's line in the file, plain text", () => {
    const source = [
      '<?xml version="1.0" encoding="utf-8"?>', // 1
      "<job>", // 2
      '<script language="VBScript">', // 3
      "Dim divisor, answer", // 4
      "divisor = 0", // 5
      "answer = 1 / divisor", // 6
      "</script>", // 7
      "</job>", // 8
    ].join("\n");
    const code = parseWsf(source).jobs[0]!.scripts[0]!.code!;
    expect(findBlockStart(source, code, 0)).toEqual({
      line: 3,
      next: expect.any(Number),
    });
  });

  it("counts the block's line in the file when it is indented", () => {
    const source = [
      '<?xml version="1.0" encoding="utf-8"?>',
      "<job>",
      '  <script language="VBScript">',
      "    Dim divisor, answer",
      "    answer = 1 / divisor",
      "  </script>",
      "</job>",
    ].join("\n");
    const code = parseWsf(source).jobs[0]!.scripts[0]!.code!;
    expect(findBlockStart(source, code, 0)!.line).toBe(3);
  });

  // The block's own text does not stand in the file as one run, because the `CDATA` marker stands between the
  // whitespace and the script, and the parser reports the whitespace around the section separately from it. The
  // first line that carries anything does stand in the file, and the block begins that many lines above it.
  it("counts the block's line in the file when it is written as CDATA", () => {
    const source = [
      '<?xml version="1.0" encoding="utf-8"?>', // 1
      "<package>", // 2
      '  <job id="Main">', // 3
      '    <script language="VBScript">', // 4
      "    <![CDATA[", // 5
      "        Dim divisor, answer", // 6
      "        divisor = 0", // 7
      "        answer = 1 / divisor", // 8
      "    ]]>", // 9
      "    </script>", // 10
      "  </job>", // 11
      "</package>", // 12
    ].join("\n");
    const code = parseWsf(source).jobs[0]!.scripts[0]!.code!;
    // The section's own text begins on the line the marker is on, and the parser drops the whitespace around it.
    expect(code.split("\n")[0]).toBe("");
    expect(code).toContain("answer = 1 / divisor");
    expect(findBlockStart(source, code, 0)!.line).toBe(5);
  });

  // An author who escapes what XML reserves writes `&amp;` where the script means `&`, so the code the parser
  // hands back is not the text the file holds. The file's own text is looked for as well.
  it("finds a block written with an escape rather than a CDATA section", () => {
    const source = [
      '<?xml version="1.0" encoding="utf-8"?>',
      "<job>",
      '<script language="VBScript">',
      'text = "a" &amp; "b"',
      "answer = 1 / 0",
      "</script>",
      "</job>",
    ].join("\n");
    const code = parseWsf(source).jobs[0]!.scripts[0]!.code!;
    expect(code).toContain('text = "a" & "b"');
    expect(findBlockStart(source, code, 0)!.line).toBe(3);
  });

  it("keeps to document order, so one block is not mistaken for the next", () => {
    const source = [
      '<?xml version="1.0" encoding="utf-8"?>',
      "<package>",
      '  <job id="Main">',
      '    <script language="VBScript">', // 4
      '      WScript.Echo "one"',
      "    </script>",
      '    <script language="VBScript">', // 7
      '      WScript.Echo "one"',
      "    </script>",
      "  </job>",
      "</package>",
    ].join("\n");
    const blocks = parseWsf(source).jobs[0]!.scripts;
    const first = findBlockStart(source, blocks[0]!.code!, 0)!;
    const second = findBlockStart(source, blocks[1]!.code!, first.next)!;
    expect(first.line).toBe(4);
    expect(second.line).toBe(7);
  });

  it("answers null for code that is not in the file", () => {
    expect(findBlockStart("<job/>", "WScript.Echo 1", 0)).toBeNull();
  });
});
