/**
 * Parser Cache - LRU Cache for Parsed AST
 * 
 * Performance improvements:
 * 1. Avoid re-parsing the same code multiple times
 * 2. LRU eviction policy to limit memory usage
 * 3. Hash-based cache key for fast lookup
 */

import type { Program } from '../ast/index.ts';

interface CacheEntry {
  ast: Program;
  lastAccessed: number;
  accessCount: number;
}

export class ParserCache {
  private cache: Map<string, CacheEntry> = new Map();
  private maxSize: number;
  private defaultTTL: number; // Time to live in milliseconds

  constructor(maxSize: number = 100, defaultTTL: number = 5 * 60 * 1000) {
    this.maxSize = maxSize;
    this.defaultTTL = defaultTTL;
  }

  /**
   * Generate a simple hash for the source code.
   * Uses a fast djb2-like hash algorithm.
   */
  private hashSource(source: string): string {
    let hash = 5381;
    for (let i = 0; i < source.length; i++) {
      hash = ((hash << 5) + hash) + source.charCodeAt(i)!;
    }
    return hash.toString(36);
  }

  /**
   * Get cached AST for a source code string.
   * Returns undefined if not found or expired.
   */
  get(source: string): Program | undefined {
    const key = this.hashSource(source);
    const entry = this.cache.get(key);
    
    if (!entry) {
      return undefined;
    }

    // Check if expired
    const now = Date.now();
    if (now - entry.lastAccessed > this.defaultTTL) {
      this.cache.delete(key);
      return undefined;
    }

    // Update access stats
    entry.lastAccessed = now;
    entry.accessCount++;
    
    return entry.ast;
  }

  /**
   * Store parsed AST in cache.
   */
  set(source: string, ast: Program): void {
    // Evict oldest entries if cache is full
    if (this.cache.size >= this.maxSize) {
      this.evictLRU();
    }

    const key = this.hashSource(source);
    const now = Date.now();
    
    this.cache.set(key, {
      ast,
      lastAccessed: now,
      accessCount: 1,
    });
  }

  /**
   * Check if source code is in cache.
   */
  has(source: string): boolean {
    const key = this.hashSource(source);
    const entry = this.cache.get(key);
    
    if (!entry) {
      return false;
    }

    // Check if expired
    const now = Date.now();
    if (now - entry.lastAccessed > this.defaultTTL) {
      this.cache.delete(key);
      return false;
    }

    return true;
  }

  /**
   * Evict least recently used entries.
   */
  private evictLRU(): void {
    let oldestKey: string | null = null;
    let oldestTime = Infinity;

    for (const [key, entry] of this.cache) {
      if (entry.lastAccessed < oldestTime) {
        oldestTime = entry.lastAccessed;
        oldestKey = key;
      }
    }

    if (oldestKey) {
      this.cache.delete(oldestKey);
    }
  }

  /**
   * Clear all cached entries.
   */
  clear(): void {
    this.cache.clear();
  }

  /**
   * Get cache statistics.
   */
  getStats(): {
    size: number;
    maxSize: number;
    totalAccessCount: number;
    averageAccessCount: number;
  } {
    let totalAccessCount = 0;
    for (const entry of this.cache.values()) {
      totalAccessCount += entry.accessCount;
    }

    return {
      size: this.cache.size,
      maxSize: this.maxSize,
      totalAccessCount,
      averageAccessCount: this.cache.size > 0 ? totalAccessCount / this.cache.size : 0,
    };
  }

  /**
   * Clean up expired entries.
   */
  cleanup(): number {
    const now = Date.now();
    let removed = 0;

    for (const [key, entry] of this.cache) {
      if (now - entry.lastAccessed > this.defaultTTL) {
        this.cache.delete(key);
        removed++;
      }
    }

    return removed;
  }
}

// Global parser cache instance
export const globalParserCache = new ParserCache();

/**
 * Cached parse function that wraps the original parser.
 * Use this instead of the original parse() for better performance.
 */
export function parseWithCache(
  source: string,
  parseFn: (source: string) => Program
): Program {
  // Try to get from cache
  const cached = globalParserCache.get(source);
  if (cached) {
    return cached;
  }

  // Parse and cache
  const ast = parseFn(source);
  globalParserCache.set(source, ast);
  return ast;
}
