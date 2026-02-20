/**
 * String Interning Pool for VBScript
 * 
 * VBScript is case-insensitive, so we need to normalize strings.
 * This module provides a string interning pool to:
 * 1. Avoid repeated toLowerCase() calls
 * 2. Share identical strings to reduce memory usage
 * 3. Enable fast string comparison using reference equality
 */

export class StringInterner {
  private pool: Map<string, string> = new Map();

  /**
   * Intern a string - returns the canonical (lowercased) version from the pool
   * or creates a new entry if not exists.
   */
  intern(str: string): string {
    const lower = str.toLowerCase();
    const existing = this.pool.get(lower);
    if (existing !== undefined) {
      return existing;
    }
    this.pool.set(lower, lower);
    return lower;
  }

  /**
   * Get the interned version of a string if it exists, otherwise return undefined.
   * Use this when you want to check if a string has been interned without adding it.
   */
  getInterned(str: string): string | undefined {
    return this.pool.get(str.toLowerCase());
  }

  /**
   * Check if a string (or its lowercase version) is already interned.
   */
  has(str: string): boolean {
    return this.pool.has(str.toLowerCase());
  }

  /**
   * Clear the intern pool. Useful for testing or memory-constrained environments.
   */
  clear(): void {
    this.pool.clear();
  }

  /**
   * Get the size of the intern pool.
   */
  get size(): number {
    return this.pool.size;
  }
}

// Global string interner instance
export const globalStringInterner = new StringInterner();

/**
 * Fast string comparison using interned strings.
 * Returns true if the interned versions of both strings are the same reference.
 */
export function internedEquals(a: string, b: string): boolean {
  return globalStringInterner.intern(a) === globalStringInterner.intern(b);
}
