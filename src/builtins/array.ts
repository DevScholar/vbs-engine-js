import type { VbValue } from '../runtime/index.ts';
import {
  toNumber,
  toString,
  VbArray,
  createVbArrayFromValues,
  createVbError,
  VbErrorCodes,
} from '../runtime/index.ts';

export const arrayFunctions = {
  Array: (...args: VbValue[]): VbValue => {
    const arr = createVbArrayFromValues(args);
    return { type: 'Array', value: arr };
  },

  // The bounds of something that is not an array are a type mismatch, a Null included: LBound(Null) is error
  // 13 and not the invalid use of Null, and a dimension outside the array is error 9, which getBounds raises.
  LBound: (arrayname: VbValue, dimension?: VbValue): VbValue => {
    if (arrayname.type !== 'Array') {
      throw createVbError(VbErrorCodes.TypeMismatch, 'Type mismatch: LBound');
    }
    const arr = arrayname.value as VbArray;
    const dim = dimension ? Math.floor(toNumber(dimension)) : 1;
    const bounds = arr.getBounds(dim);
    return { type: 'Long', value: bounds.lower };
  },

  UBound: (arrayname: VbValue, dimension?: VbValue): VbValue => {
    if (arrayname.type !== 'Array') {
      throw createVbError(VbErrorCodes.TypeMismatch, 'Type mismatch: UBound');
    }
    const arr = arrayname.value as VbArray;
    const dim = dimension ? Math.floor(toNumber(dimension)) : 1;
    const bounds = arr.getBounds(dim);
    return { type: 'Long', value: bounds.upper };
  },

  IsArray: (expression: VbValue): VbValue => {
    return { type: 'Boolean', value: expression.type === 'Array' };
  },

  Filter: (
    inputStrings: VbValue,
    value: VbValue,
    include?: VbValue,
    _compare?: VbValue
  ): VbValue => {
    void _compare; // Intentionally unused - matches VBScript signature
    // A Null anywhere is the invalid use of Null, which a fresh cscript shows for Filter(Null, "a").
    if (
      [inputStrings, value, include, _compare].some(arg => arg !== undefined && arg.type === 'Null')
    ) {
      throw createVbError(VbErrorCodes.InvalidUseOfNull, 'Invalid use of Null: Filter');
    }
    if (inputStrings.type !== 'Array') {
      throw createVbError(VbErrorCodes.TypeMismatch, 'Type mismatch: Filter');
    }
    const arr = (inputStrings.value as VbArray).toArray();
    // Both sides are read as text, so Filter(Array(1, 2), 1) answers with the element 1.
    const searchValue = toString(value).toLowerCase();
    const shouldInclude = include ? toNumber(include) !== 0 : true;

    const filtered = arr.filter(item => {
      const str = toString(item).toLowerCase();
      const found = str.includes(searchValue);
      return shouldInclude ? found : !found;
    });

    const result = createVbArrayFromValues(filtered);
    return { type: 'Array', value: result };
  },
};

export function registerArrayFunctions(context: {
  functionRegistry: {
    register: (
      name: string,
      func: (...args: VbValue[]) => VbValue,
      options?: { isSub: boolean }
    ) => void;
  };
}): void {
  Object.entries(arrayFunctions).forEach(([name, func]) => {
    context.functionRegistry.register(name, func as (...args: VbValue[]) => VbValue);
  });
}
