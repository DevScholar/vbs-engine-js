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
