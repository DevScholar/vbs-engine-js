import { XMLParser } from "fast-xml-parser";

/**
 * Parser for Windows Script File (`.wsf`) documents, the XML wrapper WSH uses to
 * bundle one or more jobs (each possibly mixing VBScript and JScript), declare
 * command-line parameters via `<runtime>`, and wire COM `<object>`/`<reference>`
 * and `<resource>` blocks.
 *
 * The result is a plain, typed {@link WsfDocument} the run-script CLI turns into a
 * `WScript` environment (arguments, usage text, the selected job's scripts).
 *
 * fast-xml-parser already handles the CDATA, entity, and attribute quoting that a
 * hand-rolled scanner would have to; this module only normalizes its generic tree
 * (which uses arrays for repeated elements and objects for singletons) into the
 * fixed WSF shape below.
 */

export interface WsfScript {
  /** e.g. "VBScript" or "JScript" (case as authored). */
  language: string;
  /** Inline source text, when present. */
  code?: string;
  /** External script path via `src="..."`, when present. */
  src?: string;
}

export interface WsfRuntimeParam {
  name: string;
  helpString?: string;
  /** "string" | "boolean" | "simple" (unvalidated passthrough). */
  type?: string;
  required?: boolean;
}

export interface WsfJob {
  /** The `<job id="...">` value; undefined for a single implicit job. */
  id?: string;
  description?: string;
  example?: string;
  named: WsfRuntimeParam[];
  unnamed: WsfRuntimeParam[];
  scripts: WsfScript[];
  objects: Array<{ id?: string; progid?: string; classid?: string }>;
  references: Array<{ object?: string; guid?: string; version?: string }>;
  resources: Record<string, string>;
}

export interface WsfDocument {
  jobs: WsfJob[];
}

/** Coerces a fast-xml-parser value (object | array | undefined) to an array. */
function toArray<T>(v: T | T[] | undefined): T[] {
  if (v === undefined || v === null) return [];
  return Array.isArray(v) ? v : [v];
}

function toBool(v: unknown): boolean | undefined {
  if (typeof v === "boolean") return v;
  if (typeof v === "string") return v.toLowerCase() === "true";
  return undefined;
}

function parseRuntime(
  runtime: unknown,
): Pick<WsfJob, "description" | "example" | "named" | "unnamed"> {
  if (runtime === null || typeof runtime !== "object") {
    return {
      description: undefined,
      example: undefined,
      named: [],
      unnamed: [],
    };
  }
  const r = runtime as Record<string, unknown>;

  const description =
    typeof r.description === "string"
      ? r.description
      : typeof r.description === "object"
        ? String((r.description as Record<string, unknown>)["#text"] ?? "")
        : undefined;

  const example =
    typeof r.example === "string"
      ? r.example
      : typeof r.example === "object"
        ? String((r.example as Record<string, unknown>)["#text"] ?? "")
        : undefined;

  const named = toArray<Record<string, unknown>>(
    r.named as Record<string, unknown>[] | Record<string, unknown> | undefined,
  ).map((n) => ({
    name: String(n["name"] ?? ""),
    helpString:
      n["helpstring"] === undefined ? undefined : String(n["helpstring"]),
    type: n["type"] === undefined ? undefined : String(n["type"]),
    required: toBool(n["required"]),
  }));

  const unnamed = toArray<Record<string, unknown>>(
    r.unnamed as
      Record<string, unknown>[] | Record<string, unknown> | undefined,
  ).map((n) => ({
    name: String(n["name"] ?? ""),
    helpString:
      n["helpstring"] === undefined ? undefined : String(n["helpstring"]),
    type: n["type"] === undefined ? undefined : String(n["type"]),
    required: toBool(n["required"]),
  }));

  return { description, example, named, unnamed };
}

function parseJob(job: unknown): WsfJob {
  const j = (job ?? {}) as Record<string, unknown>;

  const scripts = toArray<Record<string, unknown>>(
    j.script as Record<string, unknown>[] | Record<string, unknown> | undefined,
  ).map((s) => {
    const language = String(s["language"] ?? "");
    const src = s["src"] === undefined ? undefined : String(s["src"]);
    // What stands between the tags is the script, and a CDATA section is one of the two ways an author may write
    // it. The parser reports the section separately from the whitespace around it, and joins the whitespace
    // before that section to the whitespace after it, so the two cannot simply be put back together in order.
    // Where that whitespace is all there is, the section is the script, and taking the whitespace with it would
    // leave a blank line the file does not have.
    const text = s["#text"] === undefined ? "" : String(s["#text"]);
    const cdata = s["__cdata"] === undefined ? "" : String(s["__cdata"]);
    let code: string | undefined;
    if (cdata !== "" && text.trim() === "") code = cdata;
    else if (text === "" && cdata === "") code = undefined;
    else code = `${text}${cdata}`;
    return { language, src, code };
  });

  const objects = toArray<Record<string, unknown>>(
    j.object as Record<string, unknown>[] | Record<string, unknown> | undefined,
  ).map((o) => ({
    id: o["id"] === undefined ? undefined : String(o["id"]),
    progid: o["progid"] === undefined ? undefined : String(o["progid"]),
    classid: o["classid"] === undefined ? undefined : String(o["classid"]),
  }));

  const references = toArray<Record<string, unknown>>(
    j.reference as
      Record<string, unknown>[] | Record<string, unknown> | undefined,
  ).map((o) => ({
    object: o["object"] === undefined ? undefined : String(o["object"]),
    guid: o["guid"] === undefined ? undefined : String(o["guid"]),
    version: o["version"] === undefined ? undefined : String(o["version"]),
  }));

  const resources: Record<string, string> = {};
  for (const res of toArray<Record<string, unknown>>(
    j.resource as
      Record<string, unknown>[] | Record<string, unknown> | undefined,
  )) {
    const id = res["id"] === undefined ? undefined : String(res["id"]);
    if (id !== undefined) {
      resources[id] =
        res["#text"] !== undefined
          ? String(res["#text"])
          : res["__cdata"] !== undefined
            ? String(res["__cdata"])
            : "";
    }
  }

  const runtime = parseRuntime(j.runtime);

  return {
    id: j["id"] === undefined ? undefined : String(j["id"]),
    description: runtime.description,
    example: runtime.example,
    named: runtime.named,
    unnamed: runtime.unnamed,
    scripts,
    objects,
    references,
    resources,
  };
}

/**
 * Parses a `.wsf` XML document into its constituent jobs.
 *
 * @param xml - Raw WSF source text.
 * @returns A {@link WsfDocument} (always at least one job).
 * @throws Error if the XML is malformed or the document has no `<job>`.
 */
export function parseWsf(xml: string): WsfDocument {
  const parser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: "",
    textNodeName: "#text",
    cdataPropName: "__cdata",
    processEntities: true,
    trimValues: false,
  });

  const root = parser.parse(xml) as Record<string, unknown>;

  // The top level is either a <package> wrapping one or more <job>, or a bare
  // <job> when the file holds a single job (WSH allows omitting <package> then).
  const packageNode = root["package"] as Record<string, unknown> | undefined;
  let jobNodes: unknown[];
  if (packageNode) {
    jobNodes = toArray(packageNode["job"]);
  } else if (root["job"] !== undefined) {
    jobNodes = toArray(root["job"]);
  } else {
    throw new Error("WSF: no <job> element found");
  }

  const jobs = jobNodes.map(parseJob);
  if (jobs.length === 0) {
    throw new Error("WSF: no <job> element found");
  }
  return { jobs };
}

/** How many newlines a piece of text holds. */
function countNewlines(text: string): number {
  let total = 0;
  for (let at = 0; at < text.length; at++) {
    if (text.charCodeAt(at) === 10) total += 1;
  }
  return total;
}

/** The text as XML would carry it when an author escapes what XML reserves instead of using CDATA. */
function escapeXmlText(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

/**
 * Where a script block's code stands in the `.wsf` file, so the engine can be told which line it begins on and
 * report a failure at the file's own line, the way Windows Script Host does. Measured with `cscript`: a failure
 * on wsf line 6 of a block that begins on wsf line 3 came back as line 6, and one on wsf line 9 of a block written
 * as CDATA that begins on wsf line 4 came back as line 9.
 *
 * `fast-xml-parser` carries no source positions, so the block is found by the first of its lines that carries
 * anything, searched from where the last block was found to keep blocks in document order. That line stands in the
 * file as the code has it whether it was written as plain text or inside a CDATA section, where the whole block
 * would not be found, since the file holds the `CDATA` marker between the two. A block written with escapes
 * rather than CDATA is found by its escaped text too, since that is what the file holds.
 *
 * @param source - The whole `.wsf` file.
 * @param code - The block's code, as the parser handed it back.
 * @param from - Where to start looking, which the caller moves past each block it finds.
 * @returns The line the code begins on, counted from 1, and where to look for the next block, or null when the
 * code is not in the file.
 */
export function findBlockStart(
  source: string,
  code: string,
  from: number,
): { line: number; next: number } | null {
  const codeLines = code.split("\n");
  const firstContent = codeLines.findIndex((line) => line.trim() !== "");
  if (firstContent < 0) return null;

  const anchor = codeLines[firstContent] as string;
  let bestAt = -1;
  for (const candidate of [anchor, escapeXmlText(anchor)]) {
    const at = source.indexOf(candidate, from);
    if (at >= 0 && (bestAt < 0 || at < bestAt)) bestAt = at;
  }
  if (bestAt < 0) return null;

  // The anchor is the block's own `firstContent`th line, so the block begins that many lines above it.
  const line = Math.max(
    1,
    countNewlines(source.slice(0, bestAt)) + 1 - firstContent,
  );
  return { line, next: bestAt + anchor.length };
}
