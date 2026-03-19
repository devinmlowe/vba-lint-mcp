// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { createHash } from 'node:crypto';
import type { ParseResult } from './index.js';
import { logger } from '../logger.js';

/**
 * LRU parse cache keyed on SHA-256 of content.
 *
 * Avoids re-parsing identical content during workspace scans
 * and repeat inspections. This is a performance optimization —
 * the server produces correct results with an empty cache.
 */
export class ParseCache {
  private cache = new Map<string, { result: ParseResult; timestamp: number }>();
  private readonly maxSize: number;

  constructor(maxSize = 50) {
    this.maxSize = maxSize;
  }

  /**
   * Look up a cached parse result by content hash.
   * Returns undefined on cache miss.
   */
  get(content: string): ParseResult | undefined {
    const key = this.hashContent(content);
    const entry = this.cache.get(key);
    if (!entry) {
      return undefined;
    }

    // LRU: move to end (most recently used)
    this.cache.delete(key);
    entry.timestamp = Date.now();
    this.cache.set(key, entry);

    logger.debug({ key: key.slice(0, 12) }, 'Parse cache hit');
    return entry.result;
  }

  /**
   * Store a parse result in the cache.
   * Evicts the least-recently-used entry if at capacity.
   */
  set(content: string, result: ParseResult): void {
    const key = this.hashContent(content);

    // If already present, delete first so re-insert goes to end
    if (this.cache.has(key)) {
      this.cache.delete(key);
    }

    // Evict LRU (first entry in Map iteration order) if at capacity
    if (this.cache.size >= this.maxSize) {
      const firstKey = this.cache.keys().next().value;
      if (firstKey !== undefined) {
        this.cache.delete(firstKey);
        logger.debug({ evictedKey: firstKey.slice(0, 12) }, 'Parse cache LRU eviction');
      }
    }

    this.cache.set(key, { result, timestamp: Date.now() });
  }

  /**
   * Clear all cached entries.
   */
  clear(): void {
    this.cache.clear();
  }

  /**
   * Number of entries currently in the cache.
   */
  get size(): number {
    return this.cache.size;
  }

  /**
   * Compute SHA-256 hash of content for use as cache key.
   */
  private hashContent(content: string): string {
    return createHash('sha256').update(content).digest('hex');
  }
}

/** Singleton parse cache instance for the server. */
export const parseCache = new ParseCache();
