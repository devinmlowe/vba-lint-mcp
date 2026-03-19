// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

/**
 * Performance benchmarks for parse and inspection pipelines.
 *
 * Generates a 500-line VBA module programmatically and measures:
 * - Parse time
 * - Full inspection time (Tier A + Tier B)
 *
 * Asserts performance stays within acceptable bounds.
 * Logs actual times for tracking regression over time.
 */

import { describe, it, expect } from 'vitest';
import { parseCode, warmUpParser } from '../../src/parser/index.js';
import { createAllInspections } from '../../src/inspections/registry.js';
import { runInspections } from '../../src/inspections/runner.js';
import { buildSingleModuleFinder } from '../../src/symbols/workspace.js';
import { ParseCache } from '../../src/parser/cache.js';

// Warm up parser (same as server startup) so benchmarks measure
// steady-state performance, not cold-start ATN initialization.
warmUpParser();

/**
 * Generate a ~500-line VBA module with realistic code patterns.
 */
function generateLargeModule(): string {
  const lines: string[] = [
    'Option Explicit',
    '',
    '\'  Auto-generated 500-line VBA module for benchmarking',
    '',
  ];

  // Module-level constants (must come before procedures in VBA)
  for (let i = 1; i <= 10; i++) {
    lines.push(`Private Const CONST_${i} As Long = ${i * 42}`);
  }
  lines.push('');

  // Generate 15 procedures, each ~20-25 lines
  for (let i = 1; i <= 15; i++) {
    lines.push(`Public Sub Procedure${i}()`);
    lines.push(`    Dim result As Long`);
    lines.push(`    Dim temp As String`);
    lines.push(`    Dim counter As Long`);
    lines.push('');

    // Variable assignment
    lines.push(`    result = ${i * 10}`);
    lines.push(`    temp = "Procedure ${i}"`);
    lines.push('');

    // For loop
    lines.push(`    For counter = 1 To ${i + 5}`);
    lines.push(`        result = result + counter`);
    lines.push('    Next counter');
    lines.push('');

    // If/Else
    lines.push(`    If result > ${i * 100} Then`);
    lines.push(`        temp = temp & " - high"`);
    lines.push('    Else');
    lines.push(`        temp = temp & " - low"`);
    lines.push('    End If');
    lines.push('');

    // Select Case
    lines.push(`    Select Case result Mod 3`);
    lines.push('        Case 0');
    lines.push(`            temp = temp & " divisible"`);
    lines.push('        Case 1');
    lines.push(`            temp = temp & " remainder 1"`);
    lines.push('        Case Else');
    lines.push(`            temp = temp & " remainder 2"`);
    lines.push('    End Select');
    lines.push('');

    lines.push('End Sub');
    lines.push('');
  }

  // Generate 3 functions
  for (let i = 1; i <= 3; i++) {
    lines.push(`Public Function Calculate${i}(ByVal input As Long) As Long`);
    lines.push(`    Dim accumulator As Long`);
    lines.push(`    accumulator = input`);
    lines.push('');

    // Do While loop
    lines.push('    Do While accumulator > 0');
    lines.push('        accumulator = accumulator - 1');
    lines.push('    Loop');
    lines.push('');

    lines.push(`    Calculate${i} = accumulator + input`);
    lines.push('End Function');
    lines.push('');
  }

  // Pad to ~500 lines with comment blocks
  while (lines.length < 500) {
    lines.push(`' Line ${lines.length + 1}: benchmark padding comment`);
  }

  return lines.join('\n') + '\n';
}

describe('Performance Benchmarks', () => {
  const largeModule = generateLargeModule();

  it('generated module has ~500 lines', () => {
    const lineCount = largeModule.split('\n').length;
    expect(lineCount).toBeGreaterThanOrEqual(490);
    expect(lineCount).toBeLessThanOrEqual(560);
  });

  it('parse time < 3000ms for 500-line module (after warm-up)', () => {
    const start = Date.now();
    const parseResult = parseCode(largeModule);
    const elapsed = Date.now() - start;

    console.log(`[BENCHMARK] Parse 500-line module: ${elapsed}ms`);

    expect(elapsed).toBeLessThan(3000);
    expect(parseResult.tree).toBeDefined();
    expect(parseResult.errors.length).toBe(0);
  });

  it('full inspection (Tier A + B) < 1000ms for 500-line module', () => {
    const parseResult = parseCode(largeModule);
    const declarationFinder = buildSingleModuleFinder(parseResult);
    const inspections = createAllInspections();

    const start = Date.now();
    const { results, errors } = runInspections(inspections, {
      parseResult,
      declarationFinder,
    }, {
      hasSymbolTable: true,
      hostLibraries: ['excel'],
    });
    const elapsed = Date.now() - start;

    console.log(`[BENCHMARK] Full inspection (Tier A + B) 500-line module: ${elapsed}ms`);
    console.log(`[BENCHMARK] Results: ${results.length} findings, ${errors.length} errors`);

    expect(elapsed).toBeLessThan(1000);
  });

  it('cached parse returns immediately', () => {
    const cache = new ParseCache();

    // Warm the cache
    const parseResult = parseCode(largeModule);
    cache.set(largeModule, parseResult);

    // Measure cache hit time
    const start = Date.now();
    const cached = cache.get(largeModule);
    const elapsed = Date.now() - start;

    console.log(`[BENCHMARK] Cache hit: ${elapsed}ms`);

    expect(cached).toBe(parseResult);
    expect(elapsed).toBeLessThan(50);
  });

  it('end-to-end inspect pipeline < 3000ms for 500-line module', () => {
    const start = Date.now();

    // Full pipeline: parse + symbol table + inspect
    const parseResult = parseCode(largeModule);
    const declarationFinder = buildSingleModuleFinder(parseResult);
    const inspections = createAllInspections();
    const { results } = runInspections(inspections, {
      parseResult,
      declarationFinder,
    }, {
      hasSymbolTable: true,
      hostLibraries: ['excel'],
    });

    const elapsed = Date.now() - start;

    console.log(`[BENCHMARK] End-to-end pipeline 500-line module: ${elapsed}ms (${results.length} findings)`);

    expect(elapsed).toBeLessThan(3000);
  });
});
