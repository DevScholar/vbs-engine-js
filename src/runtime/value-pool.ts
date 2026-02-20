/**
 * Value Pool - Flyweight Pattern for Common VBScript Values
 * 
 * Caches frequently used VbValue objects to avoid repeated allocations.
 * This significantly reduces GC pressure for common values like:
 * - Small integers (-128 to 127)
 * - Empty, Null, Nothing
 * - Common booleans
 * - Empty string
 */

import type { VbValue, VbLongValue, VbBooleanValue, VbStringValue } from './values.ts';
import { VbEmpty, VbNull, VbNothing } from './values.ts';

// Cache for small integers (common loop counters, indices)
const SMALL_INT_MIN = -128;
const SMALL_INT_MAX = 127;
const smallIntCache: VbLongValue[] = [];

// Initialize small integer cache
for (let i = SMALL_INT_MIN; i <= SMALL_INT_MAX; i++) {
  smallIntCache[i - SMALL_INT_MIN] = { type: 'Long', value: i };
}

// Common boolean values
const VbTrue: VbBooleanValue = { type: 'Boolean', value: true };
const VbFalse: VbBooleanValue = { type: 'Boolean', value: false };

// Empty string
const VbEmptyString: VbStringValue = { type: 'String', value: '' };

// Common string cache (frequently used strings)
const commonStringCache: Map<string, VbStringValue> = new Map();

/**
 * Get a cached VbValue for small integers.
 * Returns cached value if in range, otherwise creates new value.
 */
export function getCachedLong(value: number): VbLongValue {
  if (value >= SMALL_INT_MIN && value <= SMALL_INT_MAX && Number.isInteger(value)) {
    return smallIntCache[value - SMALL_INT_MIN]!;
  }
  return { type: 'Long', value };
}

/**
 * Get a cached boolean value.
 */
export function getCachedBoolean(value: boolean): VbBooleanValue {
  return value ? VbTrue : VbFalse;
}

/**
 * Get a cached string value.
 * Empty string is always cached, other strings may be cached if commonly used.
 */
export function getCachedString(value: string): VbStringValue {
  if (value === '') {
    return VbEmptyString;
  }
  
  // Check common string cache
  const cached = commonStringCache.get(value);
  if (cached) {
    return cached;
  }
  
  // For short strings (likely identifiers), cache them
  if (value.length <= 32) {
    const newValue: VbStringValue = { type: 'String', value };
    commonStringCache.set(value, newValue);
    return newValue;
  }
  
  return { type: 'String', value };
}

/**
 * Pre-cache common strings (keywords, common variable names, etc.)
 */
export function preCacheCommonStrings(strings: string[]): void {
  for (const str of strings) {
    if (!commonStringCache.has(str)) {
      commonStringCache.set(str, { type: 'String', value: str });
    }
  }
}

/**
 * Create a VbValue with caching for common types.
 * This is a drop-in replacement for createVbValue that uses caching.
 */
export function createCachedVbValue(value: unknown): VbValue {
  if (value === undefined) return VbEmpty;
  if (value === null) return VbNull;
  if (value === Symbol.for('Nothing')) return VbNothing;
  if (typeof value === 'boolean') return getCachedBoolean(value);
  if (typeof value === 'number') {
    if (Number.isInteger(value)) {
      if (value >= -32768 && value <= 32767) {
        // Use Long type for integers in VBScript Integer range
        if (value >= SMALL_INT_MIN && value <= SMALL_INT_MAX) {
          return getCachedLong(value);
        }
        return { type: 'Integer', value };
      }
      // Use cached Long for small integers, new Long for larger
      return getCachedLong(value);
    }
    return { type: 'Double', value };
  }
  if (typeof value === 'string') return getCachedString(value);
  if (value instanceof Date) return { type: 'Date', value };
  if (Array.isArray(value)) return { type: 'Array', value };
  if (typeof value === 'object') return { type: 'Object', value };
  return { type: 'Variant', value };
}

// Pre-cache common VBScript keywords and identifiers
const COMMON_VBSCRIPT_STRINGS = [
  'true', 'false', 'empty', 'null', 'nothing',
  'dim', 'redim', 'const', 'public', 'private',
  'sub', 'function', 'class', 'end', 'exit',
  'if', 'then', 'else', 'elseif',
  'for', 'next', 'to', 'step', 'each', 'in',
  'do', 'loop', 'while', 'until', 'wend',
  'select', 'case', 'default',
  'and', 'or', 'not', 'xor', 'eqv', 'imp',
  'new', 'set', 'call', 'byval', 'byref',
  'option', 'explicit', 'on', 'error', 'resume',
  'with', 'property', 'get', 'let',
  'msgbox', 'inputbox',
  'i', 'j', 'k', 'n', 'x', 'y', 'z',
  'count', 'index', 'item', 'key', 'value',
  'name', 'text', 'data', 'result',
  'obj', 'arr', 'str', 'num', 'val',
  'document', 'window', 'console',
];

// Initialize common strings cache
preCacheCommonStrings(COMMON_VBSCRIPT_STRINGS);

export { VbEmpty, VbNull, VbNothing, VbTrue, VbFalse, VbEmptyString };
