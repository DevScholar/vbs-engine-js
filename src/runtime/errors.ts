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

  static fromError(error: Error): VbError {
    if (error instanceof VbError) {
      return error;
    }
    return new VbError(440, error.message, 'Vbscript');
  }
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
  return new VbError(code, description, source);
}
