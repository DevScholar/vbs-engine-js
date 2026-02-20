import type { VbValue } from './values.ts';
import { VbEmpty } from './values.ts';

export class VbArray {
  private data: VbValue[];
  private dimensions: number[];
  private lowerBounds: number[];

  constructor(dimensions: number[], lowerBounds?: number[]) {
    this.dimensions = dimensions;
    this.lowerBounds = lowerBounds ?? dimensions.map(() => 0);
    
    const totalSize = dimensions.reduce((a, b) => a * b, 1);
    this.data = new Array(totalSize).fill(VbEmpty);
  }

  private getIndex(indices: number[]): number {
    if (indices.length !== this.dimensions.length) {
      throw new Error(`Array has ${this.dimensions.length} dimension(s), but ${indices.length} index(es) provided`);
    }

    let index = 0;
    let multiplier = 1;

    for (let i = this.dimensions.length - 1; i >= 0; i--) {
      const lowerBound = this.lowerBounds[i]!;
      const dimension = this.dimensions[i]!;
      const idx = indices[i]! - lowerBound;
      if (idx < 0 || idx >= dimension) {
        throw new Error('Subscript out of range');
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
      throw new Error('Invalid dimension');
    }
    const idx = dimension - 1;
    const lower = this.lowerBounds[idx]!;
    const dim = this.dimensions[idx]!;
    return {
      lower,
      upper: lower + dim - 1,
    };
  }

  redim(dimensions: number[], preserve: boolean): void {
    const newData = new Array(dimensions.reduce((a, b) => a * b, 1)).fill(VbEmpty);
    
    if (preserve) {
      const minSize = Math.min(this.data.length, newData.length);
      for (let i = 0; i < minSize; i++) {
        newData[i] = this.data[i];
      }
    }

    this.dimensions = dimensions;
    this.data = newData;
  }

  toArray(): VbValue[] {
    return [...this.data];
  }

  /**
   * Erase the array - reset all elements to Empty
   * For fixed-size arrays: clears all elements but keeps dimensions
   * For dynamic arrays: would normally deallocate, but we just clear
   */
  erase(): void {
    this.data.fill(VbEmpty);
  }
}

export function createVbArray(bounds: number[]): VbArray {
  const dimensions = bounds.map(b => b + 1);
  return new VbArray(dimensions);
}

export function createVbArrayFromValues(values: VbValue[]): VbArray {
  const arr = new VbArray([values.length]);
  values.forEach((v, i) => arr.set([i], v));
  return arr;
}
