import { XMLParser } from 'fast-xml-parser';

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
  if (typeof v === 'boolean') return v;
  if (typeof v === 'string') return v.toLowerCase() === 'true';
  return undefined;
}

function parseRuntime(runtime: unknown): Pick<WsfJob, 'description' | 'example' | 'named' | 'unnamed'> {
  if (runtime === null || typeof runtime !== 'object') {
    return { description: undefined, example: undefined, named: [], unnamed: [] };
  }
  const r = runtime as Record<string, unknown>;

  const description =
    typeof r.description === 'string'
      ? r.description
      : typeof r.description === 'object'
        ? String((r.description as Record<string, unknown>)['#text'] ?? '')
        : undefined;

  const example =
    typeof r.example === 'string'
      ? r.example
      : typeof r.example === 'object'
        ? String((r.example as Record<string, unknown>)['#text'] ?? '')
        : undefined;

  const named = toArray<Record<string, unknown>>(
    r.named as Record<string, unknown>[] | Record<string, unknown> | undefined
  ).map(n => ({
    name: String(n['name'] ?? ''),
    helpString: n['helpstring'] === undefined ? undefined : String(n['helpstring']),
    type: n['type'] === undefined ? undefined : String(n['type']),
    required: toBool(n['required']),
  }));

  const unnamed = toArray<Record<string, unknown>>(
    r.unnamed as Record<string, unknown>[] | Record<string, unknown> | undefined
  ).map(n => ({
    name: String(n['name'] ?? ''),
    helpString: n['helpstring'] === undefined ? undefined : String(n['helpstring']),
    type: n['type'] === undefined ? undefined : String(n['type']),
    required: toBool(n['required']),
  }));

  return { description, example, named, unnamed };
}

function parseJob(job: unknown): WsfJob {
  const j = (job ?? {}) as Record<string, unknown>;

  const scripts = toArray<Record<string, unknown>>(
    j.script as Record<string, unknown>[] | Record<string, unknown> | undefined
  ).map(s => {
    const language = String(s['language'] ?? '');
    const src = s['src'] === undefined ? undefined : String(s['src']);
    const code =
      s['#text'] !== undefined
        ? String(s['#text'])
        : s['__cdata'] !== undefined
          ? String(s['__cdata'])
          : undefined;
    return { language, src, code };
  });

  const objects = toArray<Record<string, unknown>>(
    j.object as Record<string, unknown>[] | Record<string, unknown> | undefined
  ).map(o => ({
    id: o['id'] === undefined ? undefined : String(o['id']),
    progid: o['progid'] === undefined ? undefined : String(o['progid']),
    classid: o['classid'] === undefined ? undefined : String(o['classid']),
  }));

  const references = toArray<Record<string, unknown>>(
    j.reference as Record<string, unknown>[] | Record<string, unknown> | undefined
  ).map(o => ({
    object: o['object'] === undefined ? undefined : String(o['object']),
    guid: o['guid'] === undefined ? undefined : String(o['guid']),
    version: o['version'] === undefined ? undefined : String(o['version']),
  }));

  const resources: Record<string, string> = {};
  for (const res of toArray<Record<string, unknown>>(
    j.resource as Record<string, unknown>[] | Record<string, unknown> | undefined
  )) {
    const id = res['id'] === undefined ? undefined : String(res['id']);
    if (id !== undefined) {
      resources[id] =
        res['#text'] !== undefined
          ? String(res['#text'])
          : res['__cdata'] !== undefined
            ? String(res['__cdata'])
            : '';
    }
  }

  const runtime = parseRuntime(j.runtime);

  return {
    id: j['id'] === undefined ? undefined : String(j['id']),
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
    attributeNamePrefix: '',
    textNodeName: '#text',
    cdataPropName: '__cdata',
    processEntities: true,
    trimValues: false,
  });

  const root = parser.parse(xml) as Record<string, unknown>;

  // The top level is either a <package> wrapping one or more <job>, or a bare
  // <job> when the file holds a single job (WSH allows omitting <package> then).
  const packageNode = root['package'] as Record<string, unknown> | undefined;
  let jobNodes: unknown[];
  if (packageNode) {
    jobNodes = toArray(packageNode['job']);
  } else if (root['job'] !== undefined) {
    jobNodes = toArray(root['job']);
  } else {
    throw new Error('WSF: no <job> element found');
  }

  const jobs = jobNodes.map(parseJob);
  if (jobs.length === 0) {
    throw new Error('WSF: no <job> element found');
  }
  return { jobs };
}
