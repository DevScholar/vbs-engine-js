import type { VbValue } from '../runtime/index.ts';

export function vbToJs(value: VbValue): unknown {
  switch (value.type) {
    case 'Empty':
      return undefined;
    case 'Null':
      return null;
    case 'Boolean':
    case 'Long':
    case 'Double':
    case 'Integer':
    case 'String':
    case 'Date':
      return value.value;
    case 'Array': {
      const arr = value.value as { toArray(): VbValue[] };
      return arr.toArray().map(vbToJs);
    }
    case 'Object':
      return value.value;
    default:
      return value.value;
  }
}

export function jsToVb(value: unknown, thisArg?: unknown): VbValue {
  if (value === undefined) return { type: 'Empty', value: undefined };
  if (value === null) return { type: 'Null', value: null };
  if (typeof value === 'boolean') return { type: 'Boolean', value };
  if (typeof value === 'number') {
    if (Number.isInteger(value) && value >= -2147483648 && value <= 2147483647) {
      return { type: 'Long', value };
    }
    return { type: 'Double', value };
  }
  if (typeof value === 'string') return { type: 'String', value };
  if (value instanceof Date) return { type: 'Date', value };
  if (Array.isArray(value)) {
    return { type: 'Array', value };
  }
  if (typeof value === 'function') {
    return { type: 'Object', value: { type: 'jsfunction', func: value as (...args: unknown[]) => unknown, thisArg: thisArg ?? null } };
  }
  if (typeof value === 'object') {
    return { type: 'Object', value };
  }
  return { type: 'String', value: String(value) };
}
