/**
 * The engine's internal source of truth for a VBScript error.
 *
 * This class is intentionally NOT part of the public API. It is the single
 * internal representation that the two public error surfaces both derive from:
 *   - the MSScriptControl-style `engine.error` member (pure-code number), and
 *   - the IE/WSH-style thrown native `Error` (HRESULT number).
 * See core/index.ts for those two surfaces.
 *
 * @internal
 */
export class VbError extends Error {
  public number: number;
  public source: string;
  public description: string;
  public helpFile?: string;
  public helpContext?: number;
  /**
   * Where the failing statement began. Filled by the first frame that catches the error, which
   * is the innermost statement -- the one the real engine points at.
   */
  public line?: number;
  public column?: number;

  constructor(number: number, description: string, source: string = '') {
    super(description);
    this.name = 'VbError';
    this.number = number;
    this.source = source;
    this.description = description;
  }
}

/**
 * The two localized-by-the-host source strings Microsoft's engines stamp on an
 * error depending on whether it is a compile-time or run-time failure.
 */
export const VBS_RUNTIME_SOURCE = 'Microsoft VBScript runtime error';
export const VBS_COMPILE_SOURCE = 'Microsoft VBScript compilation error';

/**
 * Encode a VBScript error code into the HRESULT shape that IE/WSH expose on the
 * JavaScript side of a cross-language `catch`.
 *
 * Measured against the real engine: a plain error code is OR'd with the
 * VBScript facility `0x800A0000` (FACILITY_VBS), while a value that already
 * carries HRESULT high bits (e.g. `Err.Raise vbObjectError + N`) is passed
 * through unchanged.
 */
export function toHresult(code: number): number {
  return (code & 0xffff0000) !== 0 ? code : (0x800a0000 | code) | 0;
}

export const VbErrorCodes = {
  TypeMismatch: 13,
  InvalidUseOfNull: 94,
  SubscriptOutOfRange: 9,
  OutOfMemory: 7,
  DivisionByZero: 11,
  Overflow: 6,
  InvalidProcedureCall: 5,
  ObjectRequired: 424,
  ActiveXComponentCantCreateObject: 429,
  ObjectDoesntSupportPropertyOrMethod: 438,
  VariableNotDefined: 500,
  InvalidQualifier: 450,
  PermissionDenied: 70,
  FileNotFound: 53,
  PathNotFound: 76,
  DeviceIOError: 57,
  FileAlreadyExists: 58,
  DiskFull: 61,
  BadFileNameOrNumber: 52,
  TooManyFiles: 67,
  DeviceUnavailable: 68,
};

export function createVbError(code: number, description: string, source?: string): VbError {
  return new VbError(code, description, source ?? VBS_RUNTIME_SOURCE);
}

/**
 * English descriptions for the standard VBScript runtime error codes, keyed by code.
 *
 * Used to map a COM HRESULT back to the human-readable message the real engine reports.
 * Real VBScript localizes these; this engine's own error strings are English throughout
 * (see "Type mismatch", "Division by zero" elsewhere), so the table stays English too.
 */
export const VbErrorDescriptions: Record<number, string> = {
  5: 'Invalid procedure call or argument',
  6: 'Overflow',
  7: 'Out of memory',
  9: 'Subscript out of range',
  10: 'This array is fixed or temporarily locked',
  11: 'Division by zero',
  13: 'Type mismatch',
  14: 'Out of string space',
  17: "Can't perform requested operation",
  28: 'Out of stack space',
  35: 'Sub or Function not defined',
  48: 'Error in loading DLL',
  51: 'Internal error',
  52: 'Bad file name or number',
  53: 'File not found',
  54: 'Bad file mode',
  55: 'File already open',
  57: 'Device I/O error',
  58: 'File already exists',
  61: 'Disk full',
  62: 'Input past end of file',
  67: 'Too many files',
  68: 'Device unavailable',
  70: 'Permission denied',
  71: 'Disk not ready',
  74: "Can't rename with different drive",
  75: 'Path/File access error',
  76: 'Path not found',
  91: 'Object variable or With block variable not set',
  92: 'For loop not initialized',
  94: 'Invalid use of Null',
  322: "Can't create necessary temporary file",
  424: 'Object required',
  429: "ActiveX component can't create object",
  430: "Class doesn't support Automation",
  432: 'File name or class name not found during Automation operation',
  438: "Object doesn't support this property or method",
  440: 'Automation error',
  442:
    'Connection to type library or object library for remote process has been lost. ' +
    'Press OK for dialog box to remove reference.',
  443: "Automation object doesn't have a default value",
  445: "Object doesn't support this action",
  446: "Object doesn't support named arguments",
  447: "Object doesn't support current locale setting",
  448: 'Named argument not found',
  449: 'Argument not optional',
  450: 'Wrong number of arguments or invalid property assignment',
  451: 'Object not a collection',
  453: 'Specified DLL function not found',
  455: 'Code resource lock error',
  457: 'This key is already associated with an element of this collection',
  458: 'Variable uses an Automation type not supported in VBScript',
  462: 'The remote server machine does not exist or is unavailable',
  481: 'Invalid picture',
  500: 'Variable is undefined',
  501: 'Illegal assignment',
  502: 'Object not safe for scripting',
  503: 'Object not safe for initializing',
  504: 'Object not safe for creating',
  505: 'Invalid or unqualified reference',
  1001: 'Out of Memory',
  1002: 'Syntax error',
};

/**
 * Recover the plain VBScript error code from a 32-bit HRESULT.
 *
 * The inverse of {@link toHresult}: the low 16 bits are the code (0x800A0035 -> 53),
 * and the high 16 carry the severity + facility (FACILITY_VBS = 0x800A). Only meaningful
 * when the value actually carries HRESULT high bits; a plain code passes through unchanged.
 */
export function hresultToCode(hresult: number): number {
  return hresult & 0xffff;
}

/**
 * True when a value carries HRESULT high bits (severity + facility), as opposed to a plain
 * VBScript error code like 53. Used to tell a COM HRESULT apart from an already-mapped code.
 */
export function isHresult(hresult: number): boolean {
  return (hresult & 0xffff0000) !== 0;
}

/**
 * Build a {@link VbError} from an error thrown by a host object (e.g. a COM call via
 * node-ps1-dotnet). The bridge stamps the real HRESULT onto the Error as `.hresult` /
 * `.number`; recover it and map it to the VBScript code + description instead of the
 * generic "Automation error" (440) that a flattened message would otherwise produce.
 */
export function createVbErrorFromHost(err: unknown): VbError {
  const e = (err ?? {}) as Error & { hresult?: unknown; number?: unknown };
  const raw = typeof e.hresult === 'number' ? e.hresult : typeof e.number === 'number' ? e.number : undefined;
  if (raw !== undefined && isHresult(raw)) {
    const code = hresultToCode(raw);
    const description = VbErrorDescriptions[code] ?? e.message;
    return createVbError(code, description);
  }
  const message = err instanceof Error ? err.message : String(err);
  return createVbError(440, message);
}
