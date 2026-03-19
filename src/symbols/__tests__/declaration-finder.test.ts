// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { describe, it, expect } from 'vitest';
import { parseCode } from '../../parser/index.js';
import { buildSingleModuleFinder } from '../workspace.js';
import { DeclarationFinder } from '../declaration-finder.js';

function getFinder(code: string): DeclarationFinder {
  const result = parseCode(code);
  return buildSingleModuleFinder(result);
}

describe('DeclarationFinder', () => {
  describe('findByName', () => {
    it('finds declarations by name (case-insensitive)', () => {
      const finder = getFinder('Sub MySub()\nEnd Sub\n');
      const results = finder.findByName('mysub');
      expect(results.length).toBeGreaterThanOrEqual(1);
      expect(results[0].name).toBe('MySub');
    });
  });

  describe('findByType', () => {
    it('finds all variables', () => {
      const finder = getFinder('Sub Test()\n    Dim a As Long\n    Dim b As String\nEnd Sub\n');
      const vars = finder.findByType('Variable');
      expect(vars).toHaveLength(2);
    });
  });

  describe('findUnused', () => {
    it('finds unused local variables', () => {
      const finder = getFinder('Sub Test()\n    Dim x As Long\n    Dim y As String\n    MsgBox y\nEnd Sub\n');
      const unused = finder.findUnused();
      // x is unused (no references). y has a read reference.
      expect(unused.some(d => d.name === 'x')).toBe(true);
    });

    it('does not include module declaration', () => {
      const finder = getFinder('Sub Test()\nEnd Sub\n');
      const unused = finder.findUnused();
      expect(unused.some(d => d.declarationType === 'Module')).toBe(false);
    });
  });

  describe('findUnusedVariables', () => {
    it('finds variables with zero references', () => {
      const finder = getFinder('Sub Test()\n    Dim x As Long\nEnd Sub\n');
      const unused = finder.findUnusedVariables();
      expect(unused).toHaveLength(1);
      expect(unused[0].name).toBe('x');
    });
  });

  describe('findUnusedParameters', () => {
    it('finds unused parameters', () => {
      const finder = getFinder('Sub Test(x As Long, y As String)\n    MsgBox CStr(x)\nEnd Sub\n');
      const unused = finder.findUnusedParameters();
      expect(unused).toHaveLength(1);
      expect(unused[0].name).toBe('y');
    });
  });

  describe('findNonReturningFunctions', () => {
    it('finds functions that never assign return value', () => {
      const finder = getFinder('Function GetValue() As Long\n    MsgBox "Hello"\nEnd Function\n');
      const nonReturning = finder.findNonReturningFunctions();
      expect(nonReturning).toHaveLength(1);
      expect(nonReturning[0].name).toBe('GetValue');
    });

    it('does not flag functions that assign return value', () => {
      const finder = getFinder('Function GetValue() As Long\n    GetValue = 42\nEnd Function\n');
      const nonReturning = finder.findNonReturningFunctions();
      expect(nonReturning).toHaveLength(0);
    });

    it('does not flag Sub', () => {
      const finder = getFinder('Sub DoWork()\nEnd Sub\n');
      const nonReturning = finder.findNonReturningFunctions();
      expect(nonReturning).toHaveLength(0);
    });
  });

  describe('findReferencesTo', () => {
    it('returns all references to a declaration', () => {
      const finder = getFinder('Sub Test()\n    Dim x As Long\n    x = 42\n    MsgBox x\nEnd Sub\n');
      const x = finder.findByName('x')[0];
      const refs = finder.findReferencesTo(x);
      expect(refs.length).toBeGreaterThanOrEqual(1);
    });
  });
});
