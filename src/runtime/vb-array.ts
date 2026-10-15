import type { VbValue } from './values.ts';
import { VbEmpty } from './values.ts';
import { createVbError, VbErrorCodes } from './errors.ts';

/** The error a subscript outside the array raises, which vbscript.dll answers with 9. */
function subscriptOutOfRange(): never {
  throw createVbError(VbErrorCodes.SubscriptOutOfRange, 'Subscript out of range');
}

/** The number of elements a set of dimensions holds, where no dimensions at all mean no elements. */
function elementCount(dimensions: number[]): number {
  if (dimensions.length === 0) return 0;
  return dimensions.reduce((total, dimension) => total * dimension, 1);
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

    this.data = new Array(elementCount(dimensions)).fill(VbEmpty);
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

  /** Whether the array still has a size, which decides the error of assigning an array to it. */
  isSized(): boolean {
    return this.dimensions.length > 0;
  }

  /**
   * The place an element sits at, with the first subscript moving fastest.
   *
   * A two dimensional array holds a(0, 0), a(1, 0), a(0, 1), a(1, 1) in that order, which the order For Each
   * walks it in shows, and which is why ReDim Preserve keeps every value when only the last dimension grows:
   * the old places are the start of the new ones.
   */
  private getIndex(indices: number[]): number {
    if (indices.length !== this.dimensions.length) {
      subscriptOutOfRange();
    }

    let index = 0;
    let multiplier = 1;

    for (let i = 0; i < this.dimensions.length; i++) {
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
   * fresh cscript shows. Preserve keeps what fits. It may only change the last dimension of an array that
   * already has one: ReDim Preserve on any other one is error 9, and so is Preserve that changes how many
   * dimensions there are. An array with no size yet, Dim a() or one an Erase freed, takes any shape Preserve
   * asks for, which a fresh cscript shows for ReDim Preserve a(2) after Dim a().
   */
  redim(bounds: number[], preserve: boolean): void {
    const dimensions = bounds.map(bound => bound + 1);

    if (preserve && this.dimensions.length > 0) {
      if (dimensions.length !== this.dimensions.length) {
        subscriptOutOfRange();
      }
      const lastDimension = dimensions.length - 1;
      const changedElsewhere = dimensions.some(
        (size, index) => size !== this.dimensions[index] && index !== lastDimension
      );
      if (changedElsewhere) {
        subscriptOutOfRange();
      }
    }

    const newData = new Array(elementCount(dimensions)).fill(VbEmpty);

    if (preserve) {
      const keptCount = Math.min(this.data.length, newData.length);
      for (let index = 0; index < keptCount; index++) {
        newData[index] = this.data[index];
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
  // An array that Array, Split or Filter makes is dynamic, which a fresh cscript shows from every side: Erase
  // frees it, so UBound and a read of any element are both error 9 afterwards, while ReDim may size it again.
  const arr = new VbArray([values.length], undefined, true);
  values.forEach((v, i) => arr.set([i], v));
  return arr;
}
