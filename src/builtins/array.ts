import type { VbValue } from '../runtime/index.ts';
import { toNumber, VbArray, createVbArrayFromValues } from '../runtime/index.ts';

export const arrayFunctions = {
  Array: (...args: VbValue[]): VbValue => {
    const arr = createVbArrayFromValues(args);
    return { type: 'Array', value: arr };
  },

  LBound: (arrayname: VbValue, dimension?: VbValue): VbValue => {
    if (arrayname.type !== 'Array') {
      throw new Error('Type mismatch: LBound');
    }
    const arr = arrayname.value as VbArray;
    const dim = dimension ? Math.floor(toNumber(dimension)) : 1;
    const bounds = arr.getBounds(dim);
    return { type: 'Long', value: bounds.lower };
  },

  UBound: (arrayname: VbValue, dimension?: VbValue): VbValue => {
    if (arrayname.type !== 'Array') {
      throw new Error('Type mismatch: UBound');
    }
    const arr = arrayname.value as VbArray;
    const dim = dimension ? Math.floor(toNumber(dimension)) : 1;
    const bounds = arr.getBounds(dim);
    return { type: 'Long', value: bounds.upper };
  },

  IsArray: (expression: VbValue): VbValue => {
    return { type: 'Boolean', value: expression.type === 'Array' };
  },

  Filter: (inputStrings: VbValue, value: VbValue, include?: VbValue, _compare?: VbValue): VbValue => {
    void _compare; // Intentionally unused - matches VBScript signature
    if (inputStrings.type !== 'Array') {
      throw new Error('Type mismatch: Filter');
    }
    const arr = (inputStrings.value as VbArray).toArray();
    const searchValue = (value.value as string).toLowerCase();
    const shouldInclude = include ? toNumber(include) !== 0 : true;
    
    const filtered = arr.filter(item => {
      const str = (item.value as string).toLowerCase();
      const found = str.includes(searchValue);
      return shouldInclude ? found : !found;
    });
    
    const result = createVbArrayFromValues(filtered);
    return { type: 'Array', value: result };
  },
};

export function registerArrayFunctions(context: { functionRegistry: { register: (name: string, func: (...args: VbValue[]) => VbValue, options?: { isSub: boolean }) => void } }): void {
  Object.entries(arrayFunctions).forEach(([name, func]) => {
    context.functionRegistry.register(name, func as (...args: VbValue[]) => VbValue);
  });
}
