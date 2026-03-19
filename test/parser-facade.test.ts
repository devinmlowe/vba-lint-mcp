// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { describe, it, expect } from 'vitest';
import { parseCode, warmUpParser } from '../src/parser/index.js';

describe('Parser Facade', () => {
  it('warmUpParser completes without error', () => {
    expect(() => warmUpParser()).not.toThrow();
  });

  it('parses valid VBA and returns a tree', () => {
    const result = parseCode('Sub Test()\n    Dim x As Long\nEnd Sub\n');
    expect(result.tree).toBeDefined();
    expect(result.errors).toHaveLength(0);
    expect(result.tokens).toBeDefined();
  });

  it('collects parse errors for invalid syntax', () => {
    const result = parseCode('Sub Test(\nEnd Sub\n');
    expect(result.tree).toBeDefined(); // ANTLR4 recovers
    expect(result.errors.length).toBeGreaterThan(0);
    expect(result.errors[0].line).toBeGreaterThan(0);
    expect(result.errors[0].message).toBeTruthy();
  });

  it('rejects input exceeding size limit', () => {
    const bigInput = 'x'.repeat(600_000);
    expect(() => parseCode(bigInput)).toThrow(/exceeds maximum size/);
  });

  it('allows custom max input size', () => {
    const input = 'x'.repeat(100);
    expect(() => parseCode(input, { maxInputSize: 50 })).toThrow(/exceeds maximum size/);
  });

  it('handles empty input', () => {
    const result = parseCode('');
    expect(result.tree).toBeDefined();
  });

  it('handles whitespace-only input', () => {
    const result = parseCode('   \n\n   ');
    expect(result.tree).toBeDefined();
  });

  it('preserves source in result when filePath provided', () => {
    const result = parseCode('Sub Test()\nEnd Sub\n', { filePath: 'Module1.bas' });
    expect(result.source).toBe('Module1.bas');
  });

  it('parses Option Explicit', () => {
    const result = parseCode('Option Explicit\n\nSub Test()\nEnd Sub\n');
    expect(result.tree).toBeDefined();
    expect(result.errors).toHaveLength(0);
  });

  it('parses With block', () => {
    const code = [
      'Sub Test()',
      '    With Sheet1',
      '        .Range("A1").Value = "Hello"',
      '    End With',
      'End Sub',
      '',
    ].join('\n');
    const result = parseCode(code);
    expect(result.tree).toBeDefined();
    expect(result.errors).toHaveLength(0);
  });
});
