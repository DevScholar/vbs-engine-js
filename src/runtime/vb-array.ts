import type { VbValue } from './values.ts';
import { VbEmpty } from './values.ts';
import { createVbError, VbErrorCodes } from './errors.ts';

/** The error a subscript outside the array raises, which vbscript.dll answers with 9. */
function subscriptOutOfRange(): never {
  throw createVbError(VbErrorCodes.SubscriptOutOfRange, 'Subscript out of range');
}

export class VbArray {
  data: VbValue[];
  dimensions: number[];
  lowerBounds: number[];
  private dynamic: boolean;

  constructor(dimensions: number[], lowerBounds?: number[], dynamic = false) {
    this.dimensions = dimensions;
    this.lowerBounds = lowerBounds ?? dimensions.map(() => 0);
    this.dynamic = dynamic;

    const totalSize = dimensions.reduce((a, b) => a * b, 1);
    this.data = new Array(totalSize).fill(VbEmpty);
  }

  /** A copy of the array, which is what assigning it to another variable makes. */
  clone(): VbArray {
    const copy = new VbArray([], undefined, this.dynamic);
    copy.dimensions = [...this.dimensions];
    copy.lowerBounds = [...this.lowerBounds];
    copy.data = [...this.data];
    return copy;
  }

  /** Whether the array came from Dim a() rather than Dim a(2), which decides what Erase does to it. */
  isDynamic(): boolean {
    return this.dynamic;
  }

  private getIndex(indices: number[]): number {
    if (indices.length !== this.dimensions.length) {
      subscriptOutOfRange();
    }

    let index = 0;
    let multiplier = 1;

    for (let i = this.dimensions.length - 1; i >= 0; i--) {
      const lowerBound = this.lowerBounds[i]!;
      const dimension = this.dimensions[i]!;
      const idx = indices[i]! - lowerBound;
      if (idx < 0 || idx >= dimension) {
        subscriptOutOfRange();
      }
      index += idx * multiplier;
      multiplier *= dimension;
    }

    return index;
  }

  get(indices: number[]): VbValue {
    return this.data[this.getIndex(indices)] ?? VbEmpty;
  }

  set(indices: number[], value: VbValue): void {
    this.data[this.getIndex(indices)] = value;
  }

  getDimensions(): number {
    return this.dimensions.length;
  }

  getBounds(dimension: number): { lower: number; upper: number } {
    if (dimension < 1 || dimension > this.dimensions.length) {
      subscriptOutOfRange();
    }
    const idx = dimension - 1;
    const lower = this.lowerBounds[idx]!;
    const dim = this.dimensions[idx]!;
    return {
      lower,
      upper: lower + dim - 1,
    };
  }

  /**
   * Sizes the array again, from the bounds a script wrote.
   *
   * A bound is the last index, not a count, so ReDim a(4) makes five elements and UBound answers 4, which a
   * fresh cscript shows. Preserve keeps what fits, and it may only change the last dimension: ReDim Preserve
   * on any other one is error 9.
   */
  redim(bounds: number[], preserve: boolean): void {
    const dimensions = bounds.map(bound => bound + 1);
    if (preserve && dimensions.length === this.dimensions.length) {
      const changed = dimensions.findIndex((size, index) => size !== this.dimensions[index]);
      const lastDimension = dimensions.length - 1;
      if (changed >= 0 && changed !== lastDimension) subscriptOutOfRange();
    }
    const newData = new Array(dimensions.reduce((a, b) => a * b, 1)).fill(VbEmpty);

    if (preserve) {
      const minSize = Math.min(this.data.length, newData.length);
      for (let i = 0; i < minSize; i++) {
        newData[i] = this.data[i];
      }
    }

    this.dimensions = dimensions;
    this.lowerBounds = dimensions.map(() => 0);
    this.data = newData;
    this.dynamic = true;
  }

  toArray(): VbValue[] {
    return [...this.data];
  }

  /**
   * Erases the array: a fixed one keeps its bounds and every element becomes Empty, and a dynamic one is
   * freed, so IsArray still answers True while UBound raises error 9. Both measured on a fresh cscript.
   */
  erase(isDynamic: boolean): void {
    if (isDynamic) {
      this.dimensions = [];
      this.lowerBounds = [];
      this.data = [];
      return;
    }
    this.data.fill(VbEmpty);
  }
}

export function createVbArray(bounds: number[]): VbArray {
  // Dim a() makes a dynamic array, which Erase frees and ReDim may size again; Dim a(2) makes a fixed one.
  const dimensions = bounds.map(b => b + 1);
  return new VbArray(dimensions, undefined, bounds.length === 0);
}

export function createVbArrayFromValues(values: VbValue[]): VbArray {
  const arr = new VbArray([values.length]);
  values.forEach((v, i) => arr.set([i], v));
  return arr;
}
