// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { describe, it, expect, beforeEach } from 'vitest';
import { ParseCache } from '../cache.js';
import { parseCode } from '../index.js';

describe('ParseCache', () => {
  let cache: ParseCache;

  beforeEach(() => {
    cache = new ParseCache(3); // Small size for testing eviction
  });

  it('returns undefined on cache miss', () => {
    expect(cache.get('Sub Foo()\nEnd Sub')).toBeUndefined();
  });

  it('returns cached result on cache hit', () => {
    const content = 'Sub Foo()\nEnd Sub\n';
    const result = parseCode(content);
    cache.set(content, result);

    const cached = cache.get(content);
    expect(cached).toBeDefined();
    expect(cached).toBe(result);
  });

  it('cache hit avoids re-parse (same object returned)', () => {
    const content = 'Sub Bar()\n  Dim x As Long\nEnd Sub\n';
    const result = parseCode(content);
    cache.set(content, result);

    // Same content should return the exact same object
    const hit1 = cache.get(content);
    const hit2 = cache.get(content);
    expect(hit1).toBe(result);
    expect(hit2).toBe(result);
  });

  it('different content produces cache miss', () => {
    const content1 = 'Sub Foo()\nEnd Sub\n';
    const content2 = 'Sub Bar()\nEnd Sub\n';
    const result1 = parseCode(content1);
    cache.set(content1, result1);

    expect(cache.get(content2)).toBeUndefined();
  });

  it('evicts LRU entry when exceeding maxSize', () => {
    const contents = [
      'Sub A()\nEnd Sub\n',
      'Sub B()\nEnd Sub\n',
      'Sub C()\nEnd Sub\n',
      'Sub D()\nEnd Sub\n', // This should evict A
    ];

    for (const content of contents) {
      cache.set(content, parseCode(content));
    }

    expect(cache.size).toBe(3);
    // A should have been evicted (LRU)
    expect(cache.get(contents[0])).toBeUndefined();
    // B, C, D should still be present
    expect(cache.get(contents[1])).toBeDefined();
    expect(cache.get(contents[2])).toBeDefined();
    expect(cache.get(contents[3])).toBeDefined();
  });

  it('LRU access refreshes entry position', () => {
    const contents = [
      'Sub A()\nEnd Sub\n',
      'Sub B()\nEnd Sub\n',
      'Sub C()\nEnd Sub\n',
    ];

    for (const content of contents) {
      cache.set(content, parseCode(content));
    }

    // Access A to make it most recently used
    cache.get(contents[0]);

    // Add D — should evict B (now the LRU), not A
    const contentD = 'Sub D()\nEnd Sub\n';
    cache.set(contentD, parseCode(contentD));

    expect(cache.get(contents[0])).toBeDefined(); // A was refreshed
    expect(cache.get(contents[1])).toBeUndefined(); // B was evicted
    expect(cache.get(contents[2])).toBeDefined(); // C still present
    expect(cache.get(contentD)).toBeDefined(); // D is new
  });

  it('clear removes all entries', () => {
    cache.set('Sub A()\nEnd Sub\n', parseCode('Sub A()\nEnd Sub\n'));
    cache.set('Sub B()\nEnd Sub\n', parseCode('Sub B()\nEnd Sub\n'));
    expect(cache.size).toBe(2);

    cache.clear();
    expect(cache.size).toBe(0);
    expect(cache.get('Sub A()\nEnd Sub\n')).toBeUndefined();
  });

  it('content hash invalidates stale cache', () => {
    const contentV1 = 'Sub Foo()\n  Dim x As Long\nEnd Sub\n';
    const contentV2 = 'Sub Foo()\n  Dim x As String\nEnd Sub\n';

    const resultV1 = parseCode(contentV1);
    cache.set(contentV1, resultV1);

    // Same file, different content — should be a cache miss
    expect(cache.get(contentV2)).toBeUndefined();

    // Set V2 and verify it's cached
    const resultV2 = parseCode(contentV2);
    cache.set(contentV2, resultV2);
    expect(cache.get(contentV2)).toBe(resultV2);
  });
});
