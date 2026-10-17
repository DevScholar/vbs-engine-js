import type { VbValue } from "./values.ts";
import { VbEmpty } from "./values.ts";
import { createVbError, VbErrorCodes } from "./errors.ts";
import { BUILTIN_ARITY } from "./builtin-arity.ts";

export type VbFunction = (...args: VbValue[]) => VbValue;
export type VbSub = (...args: VbValue[]) => void;
export type VbUserFunction = (
  args: VbValue[],
  refs?: VbArgumentRef[],
) => VbValue;

export interface VbParameterInfo {
  name: string;
  byRef: boolean;
  isArray: boolean;
  /** A parameter the caller may leave out, which is what makes it the first the count may stop at. */
  isOptional?: boolean;
  /** A ParamArray, which takes any number of trailing arguments. */
  isParamArray?: boolean;
}

/**
 * The counts a declared procedure accepts, from its own parameter list. Everyone after the first optional
 * parameter may be left out, and a ParamArray takes as many as the caller writes.
 */
export function declaredArity(
  params: VbParameterInfo[],
): readonly [number, number] {
  const variadic = params.some((param) => param.isParamArray);
  const firstOptional = params.findIndex(
    (param) => param.isOptional || param.isParamArray,
  );
  const required = firstOptional < 0 ? params.length : firstOptional;
  return [required, variadic ? Number.POSITIVE_INFINITY : params.length];
}

export interface VbFunctionInfo {
  name: string;
  func: VbFunction | VbSub | VbUserFunction;
  isSub: boolean;
  minArgs: number;
  maxArgs: number;
  params: VbParameterInfo[];
  isUserDefined: boolean;
}

export interface VbArgumentRef {
  value: VbValue;
  variableName?: string;
  setValue?: (value: VbValue) => void;
  /**
   * Whether the argument was a variable a Dim declared as an array, which the parameter takes on: a fresh
   * cscript answers `Dim x() : ReDim x(2)` then a Sub that writes `values = Array(5, 6)` with error 13, while
   * the same Sub handed a plain variable that holds an array takes it.
   */
  isArrayVariable?: boolean;
}

/**
 * Refuses a call whose argument count is out of range, with the error Windows uses.
 *
 * Measured: a count that is too low and a count that is too high are the same answer, error 450, for a
 * built-in and for a declared procedure alike. Mid("abc") and Mid("abc", 1, 1, 1) are both 450, and so are
 * One() and One(1, 2) for Function One(value), and a class method is checked the same way. The message names
 * the procedure, since a script sees it in Err.Description and a long procedure can hold many calls.
 */
export function checkArgumentCount(info: VbFunctionInfo, count: number): void {
  checkArgumentRange(info.name, info.minArgs, info.maxArgs, count);
}

/** The same check for a class method, which keeps its counts beside its own name. */
export function checkArgumentRange(
  name: string,
  minArgs: number,
  maxArgs: number,
  count: number,
): void {
  if (count >= minArgs && count <= maxArgs) return;
  throw createVbError(
    VbErrorCodes.InvalidQualifier,
    `Wrong number of arguments or invalid property assignment: '${name}'`,
  );
}

export class VbFunctionRegistry {
  private functions: Map<string, VbFunctionInfo> = new Map();

  register(
    name: string,
    func: VbFunction | VbSub | VbUserFunction,
    options: Partial<VbFunctionInfo> = {},
  ): void {
    const lowerName = name.toLowerCase();
    // A declared procedure carries its own counts; a built-in takes the measured ones. A built-in is
    // registered with no parameter list at all, so an empty one is never read as "takes nothing".
    const declared =
      options.params === undefined ? null : declaredArity(options.params);
    const standard = BUILTIN_ARITY[lowerName];
    this.functions.set(lowerName, {
      name,
      func,
      isSub: options.isSub ?? false,
      minArgs: options.minArgs ?? declared?.[0] ?? standard?.[0] ?? 0,
      maxArgs:
        options.maxArgs ??
        declared?.[1] ??
        standard?.[1] ??
        Number.POSITIVE_INFINITY,
      params: options.params ?? [],
      isUserDefined: options.isUserDefined ?? false,
    });
  }

  get(name: string): VbFunctionInfo | undefined {
    return this.functions.get(name.toLowerCase());
  }

  has(name: string): boolean {
    return this.functions.has(name.toLowerCase());
  }

  call(name: string, args: VbValue[]): VbValue {
    const info = this.functions.get(name.toLowerCase());
    if (!info) {
      throw new Error(`Undefined function or sub: ${name}`);
    }

    checkArgumentCount(info, args.length);

    let result: VbValue;
    if (info.isUserDefined) {
      result = (info.func as VbUserFunction)(args);
    } else {
      result = (info.func as VbFunction)(...args);
    }

    if (info.isSub) {
      return VbEmpty;
    }

    return result ?? VbEmpty;
  }

  callWithRefs(name: string, args: VbArgumentRef[]): VbValue {
    const info = this.functions.get(name.toLowerCase());
    if (!info) {
      throw new Error(`Undefined function or sub: ${name}`);
    }

    checkArgumentCount(info, args.length);

    const values = args.map((arg) => arg.value);

    let result: VbValue;
    if (info.isUserDefined) {
      result = (info.func as VbUserFunction)(values, args);
    } else {
      result = (info.func as VbFunction)(...values);
    }

    for (let i = 0; i < args.length && i < info.params.length; i++) {
      const param = info.params[i];
      const arg = args[i];
      if (
        param &&
        arg &&
        param.byRef &&
        arg.setValue &&
        values[i] !== arg.value
      ) {
        arg.setValue(values[i]!);
      }
    }

    if (info.isSub) {
      return VbEmpty;
    }

    return result ?? VbEmpty;
  }

  getAll(): IterableIterator<[string, VbFunctionInfo]> {
    return this.functions.entries();
  }

  getUserDefinedFunctions(): Map<string, VbFunctionInfo> {
    const result = new Map<string, VbFunctionInfo>();
    for (const [name, info] of this.functions.entries()) {
      if (info.isUserDefined) {
        result.set(name, info);
      }
    }
    return result;
  }
}
