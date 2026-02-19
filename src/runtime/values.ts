export type VbValueType =
  | 'Empty'
  | 'Null'
  | 'Boolean'
  | 'Integer'
  | 'Long'
  | 'Single'
  | 'Double'
  | 'Currency'
  | 'Date'
  | 'String'
  | 'Object'
  | 'Error'
  | 'Variant'
  | 'Byte'
  | 'Array';

export interface VbEmptyValue {
  type: 'Empty';
  value: undefined;
}

export interface VbNullValue {
  type: 'Null';
  value: null;
}

export interface VbBooleanValue {
  type: 'Boolean';
  value: boolean;
}

export interface VbIntegerValue {
  type: 'Integer';
  value: number;
}

export interface VbLongValue {
  type: 'Long';
  value: number;
}

export interface VbSingleValue {
  type: 'Single';
  value: number;
}

export interface VbDoubleValue {
  type: 'Double';
  value: number;
}

export interface VbCurrencyValue {
  type: 'Currency';
  value: number;
}

export interface VbDateValue {
  type: 'Date';
  value: Date;
}

export interface VbStringValue {
  type: 'String';
  value: string;
}

export interface VbByteValue {
  type: 'Byte';
  value: number;
}

export interface VbArrayValue {
  type: 'Array';
  value: VbValue[];
}

export interface VbObjectValue {
  type: 'Object';
  value: VbObjectValueData | null;
}

export interface VbErrorValue {
  type: 'Error';
  value: number;
}

export interface VbVariantValue {
  type: 'Variant';
  value: unknown;
}

export type VbValue =
  | VbEmptyValue
  | VbNullValue
  | VbBooleanValue
  | VbIntegerValue
  | VbLongValue
  | VbSingleValue
  | VbDoubleValue
  | VbCurrencyValue
  | VbDateValue
  | VbStringValue
  | VbByteValue
  | VbArrayValue
  | VbObjectValue
  | VbErrorValue
  | VbVariantValue;

export interface VbObjectValueData {
  type?: string;
  classInfo?: { name: string };
  properties?: Map<string, VbValue>;
  getProperty?: (name: string) => VbValue;
  setProperty?: (name: string, value: VbValue, isSet?: boolean) => void;
  hasMethod?: (name: string) => boolean;
  getMethod?: (name: string) => { func: (...args: VbValue[]) => VbValue };
  call?: (...args: VbValue[]) => VbValue;
  [key: string]: unknown;
}

export const VbEmpty: VbEmptyValue = { type: 'Empty', value: undefined };
export const VbNull: VbNullValue = { type: 'Null', value: null };
export const VbNothing: VbObjectValue = { type: 'Object', value: null };

export function createVbValue(value: unknown): VbValue {
  if (value === undefined) return VbEmpty;
  if (value === null) return VbNull;
  if (value === Symbol.for('Nothing')) return VbNothing;
  if (typeof value === 'boolean') return { type: 'Boolean', value };
  if (typeof value === 'number') {
    if (Number.isInteger(value)) {
      if (value >= -32768 && value <= 32767) {
        return { type: 'Integer', value };
      }
      return { type: 'Long', value };
    }
    return { type: 'Double', value };
  }
  if (typeof value === 'string') return { type: 'String', value };
  if (value instanceof Date) return { type: 'Date', value };
  if (Array.isArray(value)) return { type: 'Array', value };
  if (typeof value === 'object') return { type: 'Object', value: value as VbObjectValueData };
  return { type: 'Variant', value };
}

export function toBoolean(value: VbValue): boolean {
  if (value.type === 'Boolean') return value.value;
  if (value.type === 'Empty') return false;
  if (value.type === 'Null') throw new Error('Type mismatch: Null cannot be converted to Boolean');
  if (value.type === 'String') {
    const str = value.value;
    if (str === '') return false;
    const num = parseFloat(str);
    if (!isNaN(num)) return num !== 0;
    return str.toLowerCase() !== 'false';
  }
  if (
    value.type === 'Integer' ||
    value.type === 'Long' ||
    value.type === 'Double' ||
    value.type === 'Single' ||
    value.type === 'Byte'
  ) {
    return value.value !== 0;
  }
  return true;
}

export function toNumber(value: VbValue): number {
  if (value.type === 'Empty') return 0;
  if (value.type === 'Null') throw new Error('Type mismatch: Null cannot be converted to Number');
  if (value.type === 'Boolean') return value.value ? -1 : 0;
  if (
    value.type === 'Integer' ||
    value.type === 'Long' ||
    value.type === 'Double' ||
    value.type === 'Single' ||
    value.type === 'Byte' ||
    value.type === 'Currency'
  ) {
    return value.value;
  }
  if (value.type === 'String') {
    const str = value.value.trim();
    if (str === '') return 0;
    const num = parseFloat(str);
    if (isNaN(num)) throw new Error(`Type mismatch: "${str}" cannot be converted to Number`);
    return num;
  }
  if (value.type === 'Date') {
    return value.value.getTime();
  }
  throw new Error(`Type mismatch: ${value.type} cannot be converted to Number`);
}

export function toString(value: VbValue): string {
  if (value.type === 'Empty') return '';
  if (value.type === 'Null') return 'Null';
  if (value.type === 'Boolean') return value.value ? 'True' : 'False';
  if (value.type === 'String') return value.value;
  if (value.type === 'Integer' || value.type === 'Long' || value.type === 'Byte') {
    return String(Math.floor(value.value));
  }
  if (value.type === 'Double' || value.type === 'Single' || value.type === 'Currency') {
    return String(value.value);
  }
  if (value.type === 'Date') {
    return value.value.toLocaleString();
  }
  if (value.type === 'Object') {
    return value.value === null ? 'Nothing' : '[object]';
  }
  if (value.type === 'Array') {
    return '[array]';
  }
  return String(value.value);
}

export function toVbDate(value: VbValue): Date {
  if (value.type === 'Date') return value.value;
  if (value.type === 'String') {
    const str = value.value;
    const date = new Date(str);
    if (isNaN(date.getTime())) {
      throw new Error(`Type mismatch: "${str}" cannot be converted to Date`);
    }
    return date;
  }
  if (value.type === 'Double' || value.type === 'Single') {
    const serial = value.value;
    const baseDate = new Date(1899, 11, 30);
    return new Date(baseDate.getTime() + serial * 86400000);
  }
  throw new Error(`Type mismatch: ${value.type} cannot be converted to Date`);
}

export function isNumeric(value: VbValue): boolean {
  if (
    value.type === 'Integer' ||
    value.type === 'Long' ||
    value.type === 'Double' ||
    value.type === 'Single' ||
    value.type === 'Currency' ||
    value.type === 'Byte'
  ) {
    return true;
  }
  if (value.type === 'String') {
    const str = value.value.trim();
    if (str === '') return false;
    return !isNaN(parseFloat(str));
  }
  if (value.type === 'Boolean') return true;
  if (value.type === 'Empty') return true;
  return false;
}

export function isEmpty(value: VbValue): boolean {
  return value.type === 'Empty';
}

export function isNull(value: VbValue): boolean {
  return value.type === 'Null';
}

export function isObject(value: VbValue): boolean {
  return value.type === 'Object';
}

export function isArray(value: VbValue): boolean {
  return value.type === 'Array';
}

export function getNumericValue(value: VbValue): number {
  switch (value.type) {
    case 'Integer':
    case 'Long':
    case 'Double':
    case 'Single':
    case 'Byte':
    case 'Currency':
      return value.value;
    default:
      return toNumber(value);
  }
}

export function getStringValue(value: VbValue): string {
  if (value.type === 'String') return value.value;
  return toString(value);
}

export function getBooleanValue(value: VbValue): boolean {
  if (value.type === 'Boolean') return value.value;
  return toBoolean(value);
}
