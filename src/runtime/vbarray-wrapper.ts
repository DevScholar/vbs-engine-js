/**
 * VBArray - A JavaScript wrapper for VBScript arrays.
 *
 * This class provides compatibility with early Internet Explorer's VBArray object,
 * which was used to access VBScript arrays from JavaScript before direct access
 * was supported in later IE versions.
 *
 * In modern usage, VBScript arrays can be accessed directly from JavaScript,
 * but VBArray is kept for backward compatibility with legacy code.
 *
 * @example
 * ```javascript
 * // Legacy usage (early IE)
 * var vbArray = new VBArray(vbsArray);
 * var jsArray = vbArray.toArray();
 *
 * // Modern usage (direct access)
 * vbsArray[0] = 1;  // Direct index access
 * ```
 */

import type { VbArray } from './vb-array.ts';
import type { VbValue } from './values.ts';

/**
 * Checks if a value is a VbArray instance (internal VBScript array)
 */
function isVbArray(value: unknown): value is VbArray {
  return (
    value !== null &&
    typeof value === 'object' &&
    'getDimensions' in value &&
    'getBounds' in value &&
    'get' in value &&
    typeof (value as VbArray).getDimensions === 'function'
  );
}

/**
 * Converts a VbValue to a JavaScript value
 */
function vbValueToJs(value: VbValue): unknown {
  if (value === null || value === undefined) {
    return value;
  }

  // Handle VbValue type
  if (typeof value === 'object' && 'type' in value) {
    const vbValue = value as VbValue;
    switch (vbValue.type) {
      case 'Empty':
        return undefined;
      case 'Null':
        return null;
      case 'Boolean':
      case 'Integer':
      case 'Long':
      case 'Single':
      case 'Double':
      case 'Currency':
      case 'String':
      case 'Byte':
        return vbValue.value;
      case 'Date':
        return vbValue.value;
      case 'Array':
        // Nested array - wrap in VBArray or convert
        return new VBArray(vbValue.value as VbArray);
      case 'Object':
        return vbValue.value;
      case 'Error':
        return vbValue.value;
      default:
        return vbValue.value;
    }
  }

  return value;
}

/**
 * VBArray class for accessing VBScript arrays from JavaScript.
 *
 * This class mimics the behavior of the VBArray object in Internet Explorer,
 * providing methods to inspect and convert VBScript arrays.
 */
export class VBArray {
  private readonly _vbArray: VbArray;

  /**
   * Creates a new VBArray wrapper around a VBScript array.
   *
   * @param vbArray - The VBScript array to wrap (internal VbArray instance)
   * @throws {TypeError} If the provided value is not a valid VBScript array
   */
  constructor(vbArray: unknown) {
    if (!isVbArray(vbArray)) {
      throw new TypeError(
        'VBArray constructor requires a VBScript array. ' +
          'Ensure you are passing a valid VBScript array from the engine.'
      );
    }
    this._vbArray = vbArray;
  }

  /**
   * Returns the number of dimensions in the array.
   *
   * @returns The number of dimensions (1 for one-dimensional arrays, 2 for two-dimensional, etc.)
   *
   * @example
   * ```javascript
   * var arr = new VBArray(vbsArray);
   * var dims = arr.dimensions();  // Returns 1, 2, etc.
   * ```
   */
  dimensions(): number {
    return this._vbArray.getDimensions();
  }

  /**
   * Returns the lower bound of the specified dimension.
   *
   * VBScript arrays can have any lower bound (not necessarily 0).
   * This method returns the starting index of the specified dimension.
   *
   * @param dimension - The dimension (1-based). Defaults to 1.
   * @returns The lower bound of the specified dimension
   * @throws {RangeError} If the dimension is out of range
   *
   * @example
   * ```javascript
   * var arr = new VBArray(vbsArray);
   * var lower = arr.lbound();     // Lower bound of first dimension
   * var lower2 = arr.lbound(2);   // Lower bound of second dimension
   * ```
   */
  lbound(dimension: number = 1): number {
    if (dimension < 1 || dimension > this._vbArray.getDimensions()) {
      throw new RangeError(`Invalid dimension: ${dimension}`);
    }
    return this._vbArray.getBounds(dimension).lower;
  }

  /**
   * Returns the upper bound of the specified dimension.
   *
   * @param dimension - The dimension (1-based). Defaults to 1.
   * @returns The upper bound of the specified dimension
   * @throws {RangeError} If the dimension is out of range
   *
   * @example
   * ```javascript
   * var arr = new VBArray(vbsArray);
   * var upper = arr.ubound();     // Upper bound of first dimension
   * var upper2 = arr.ubound(2);   // Upper bound of second dimension
   * ```
   */
  ubound(dimension: number = 1): number {
    if (dimension < 1 || dimension > this._vbArray.getDimensions()) {
      throw new RangeError(`Invalid dimension: ${dimension}`);
    }
    return this._vbArray.getBounds(dimension).upper;
  }

  /**
   * Retrieves the value at the specified indices.
   *
   * @param indices - One or more indices (one for each dimension)
   * @returns The value at the specified position, converted to a JavaScript value
   * @throws {RangeError} If the number of indices doesn't match the array dimensions
   * @throws {RangeError} If any index is out of bounds
   *
   * @example
   * ```javascript
   * var arr = new VBArray(vbsArray);
   * var val = arr.getItem(0);           // One-dimensional array
   * var val2 = arr.getItem(0, 1);       // Two-dimensional array
   * ```
   */
  getItem(...indices: number[]): unknown {
    if (indices.length !== this._vbArray.getDimensions()) {
      throw new RangeError(
        `Array has ${this._vbArray.getDimensions()} dimension(s), but ${indices.length} index(es) provided`
      );
    }

    const value = this._vbArray.get(indices);
    return vbValueToJs(value as VbValue);
  }

  /**
   * Converts the VBScript array to a JavaScript array.
   *
   * For multi-dimensional arrays, returns a nested array structure.
   *
   * @returns A JavaScript array containing the same elements
   *
   * @example
   * ```javascript
   * var arr = new VBArray(vbsArray);
   * var jsArray = arr.toArray();  // [1, 2, 3, ...]
   * ```
   */
  toArray(): unknown[] {
    const dims = this._vbArray.getDimensions();

    if (dims === 1) {
      // Simple one-dimensional array
      const bounds = this._vbArray.getBounds(1);
      const result: unknown[] = [];
      for (let i = bounds.lower; i <= bounds.upper; i++) {
        result.push(vbValueToJs(this._vbArray.get([i]) as VbValue));
      }
      return result;
    } else {
      // Multi-dimensional array - create nested structure
      return this._toArrayRecursive(1, []);
    }
  }

  /**
   * Recursively converts multi-dimensional arrays to nested JavaScript arrays.
   */
  private _toArrayRecursive(currentDim: number, indices: number[]): unknown[] {
    const bounds = this._vbArray.getBounds(currentDim);
    const result: unknown[] = [];

    for (let i = bounds.lower; i <= bounds.upper; i++) {
      const newIndices = [...indices, i];
      if (currentDim === this._vbArray.getDimensions()) {
        // Last dimension - get the actual value
        result.push(vbValueToJs(this._vbArray.get(newIndices) as VbValue));
      } else {
        // More dimensions to go - recurse
        result.push(this._toArrayRecursive(currentDim + 1, newIndices));
      }
    }

    return result;
  }
}

/**
 * Type guard to check if a value is a VBArray instance.
 */
export function isVBArray(value: unknown): value is VBArray {
  return value instanceof VBArray;
}

export default VBArray;
