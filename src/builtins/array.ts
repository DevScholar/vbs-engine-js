import type { VbValue } from "../runtime/index.ts";
import {
  toNumber,
  toString,
  toWholeNumber,
  rejectObjectArgument,
  VbArray,
  createVbArrayFromValues,
  createVbError,
  VbErrorCodes,
} from "../runtime/index.ts";
import { readCompareMode } from "./string.ts";

/** The array a list function was handed, or the error Windows answers with for anything else. */
function readListArgument(value: VbValue, functionName: string): VbArray {
  rejectObjectArgument(value, functionName);
  if (value.type !== "Array") {
    throw createVbError(
      VbErrorCodes.TypeMismatch,
      `Type mismatch: ${functionName}`,
    );
  }
  const arr = value.value as VbArray;
  // Filter reads a list of values, so an array of more than one dimension is a type mismatch there, which a
  // fresh cscript shows for Filter on a two dimensional array. A dynamic array that was never given a size is
  // an empty list rather than an error, so Filter of it is an empty array.
  if (arr.dimensions.length > 1) {
    throw createVbError(
      VbErrorCodes.TypeMismatch,
      `Type mismatch: ${functionName}`,
    );
  }
  return arr;
}

/** The text of an element a list function reads, where a Null counts as no text at all. */
function elementText(value: VbValue): string {
  return value.type === "Null" ? "" : toString(value);
}

export const arrayFunctions = {
  Array: (...args: VbValue[]): VbValue => {
    const arr = createVbArrayFromValues(args);
    return { type: "Array", value: arr };
  },

  // The bounds of something that is not an array are a type mismatch, a Null included: LBound(Null) is error
  // 13 and not the invalid use of Null, and a dimension outside the array is error 9, which getBounds raises.
  // An object is the other way round and answers 91 for a Nothing and 438 for any other.
  LBound: (arrayname: VbValue, dimension?: VbValue): VbValue => {
    rejectObjectArgument(arrayname, "LBound");
    if (arrayname.type !== "Array") {
      throw createVbError(VbErrorCodes.TypeMismatch, "Type mismatch: LBound");
    }
    const arr = arrayname.value as VbArray;
    // The dimension is a whole number, so a fraction is rounded to the nearest one: UBound(a, 2.6) on a two
    // dimensional array is error 9 there, because 2.6 is the third dimension.
    const dim =
      dimension === undefined ? 1 : toWholeNumber(toNumber(dimension));
    const bounds = arr.getBounds(dim);
    return { type: "Long", value: bounds.lower };
  },

  UBound: (arrayname: VbValue, dimension?: VbValue): VbValue => {
    rejectObjectArgument(arrayname, "UBound");
    if (arrayname.type !== "Array") {
      throw createVbError(VbErrorCodes.TypeMismatch, "Type mismatch: UBound");
    }
    const arr = arrayname.value as VbArray;
    const dim =
      dimension === undefined ? 1 : toWholeNumber(toNumber(dimension));
    const bounds = arr.getBounds(dim);
    return { type: "Long", value: bounds.upper };
  },

  IsArray: (expression: VbValue): VbValue => {
    return { type: "Boolean", value: expression.type === "Array" };
  },

  Filter: (
    inputStrings: VbValue,
    value: VbValue,
    include?: VbValue,
    compare?: VbValue,
  ): VbValue => {
    // A Null anywhere among the arguments is the invalid use of Null, which a fresh cscript shows for
    // Filter(Null, "a"). A Null inside the array is not: it counts as no text at all and matches nothing.
    if (
      [inputStrings, value, include, compare].some(
        (argument) => argument !== undefined && argument.type === "Null",
      )
    ) {
      throw createVbError(
        VbErrorCodes.InvalidUseOfNull,
        "Invalid use of Null: Filter",
      );
    }
    const arr = readListArgument(inputStrings, "Filter");
    const compareMode = readCompareMode(compare, "Filter");
    const searchValue = elementText(value);
    const shouldInclude = include ? toNumber(include) !== 0 : true;

    const filtered = arr.toArray().filter((item) => {
      const text = elementText(item);
      const found =
        compareMode === 0
          ? text.includes(searchValue)
          : text.toLowerCase().includes(searchValue.toLowerCase());
      return shouldInclude ? found : !found;
    });

    // What comes back keeps the elements it found, type and all, which a fresh cscript shows for
    // Filter(Array(11, 21), 1) with an Integer element and for Filter(Array(Empty), "") with an Empty one.
    const result = createVbArrayFromValues(filtered);
    return { type: "Array", value: result };
  },
};

export function registerArrayFunctions(context: {
  functionRegistry: {
    register: (
      name: string,
      func: (...args: VbValue[]) => VbValue,
      options?: { isSub: boolean },
    ) => void;
  };
}): void {
  Object.entries(arrayFunctions).forEach(([name, func]) => {
    context.functionRegistry.register(
      name,
      func as (...args: VbValue[]) => VbValue,
    );
  });
}
