// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { describe, it, expect } from 'vitest';
import { parseCode } from '../../parser/index.js';
import { collectDeclarations } from '../symbol-walker.js';
import { resolveReferences } from '../reference-resolver.js';
import type { Declaration } from '../declaration.js';

function getDeclarationsWithRefs(code: string): Declaration[] {
  const result = parseCode(code);
  const decls = collectDeclarations(result);
  resolveReferences(result, decls);
  return decls;
}

function findByName(decls: Declaration[], name: string): Declaration | undefined {
  return decls.find(d => d.name.toLowerCase() === name.toLowerCase());
}

describe('Reference Resolver', () => {
  describe('Simple name resolution', () => {
    it('resolves variable reference in same scope', () => {
      const code = 'Sub Test()\n    Dim x As Long\n    x = 42\n    MsgBox x\nEnd Sub\n';
      const decls = getDeclarationsWithRefs(code);
      const x = findByName(decls, 'x');
      expect(x).toBeDefined();
      // x is assigned and read
      expect(x!.references.length).toBeGreaterThanOrEqual(1);
    });

    it('marks assignment references correctly', () => {
      const code = 'Sub Test()\n    Dim x As Long\n    x = 42\nEnd Sub\n';
      const decls = getDeclarationsWithRefs(code);
      const x = findByName(decls, 'x');
      expect(x).toBeDefined();
      const assignmentRefs = x!.references.filter(r => r.isAssignment);
      expect(assignmentRefs.length).toBeGreaterThanOrEqual(1);
    });

    it('resolves parameter reference', () => {
      const code = 'Sub Test(x As Long)\n    MsgBox CStr(x)\nEnd Sub\n';
      const decls = getDeclarationsWithRefs(code);
      const x = findByName(decls, 'x');
      expect(x).toBeDefined();
      expect(x!.references.length).toBeGreaterThanOrEqual(1);
    });

    it('resolves module-level variable from procedure', () => {
      const code = 'Dim mValue As Long\nSub Test()\n    mValue = 42\nEnd Sub\n';
      const decls = getDeclarationsWithRefs(code);
      const mValue = findByName(decls, 'mValue');
      expect(mValue).toBeDefined();
      expect(mValue!.references.length).toBeGreaterThanOrEqual(1);
    });
  });

  describe('Function return assignment', () => {
    it('detects function name assignment as return value', () => {
      const code = 'Function GetValue() As Long\n    GetValue = 42\nEnd Function\n';
      const decls = getDeclarationsWithRefs(code);
      const func = findByName(decls, 'GetValue');
      expect(func).toBeDefined();
      expect(func!.declarationType).toBe('Function');
      const assignmentRefs = func!.references.filter(r => r.isAssignment);
      expect(assignmentRefs.length).toBeGreaterThanOrEqual(1);
    });

    it('detects no return assignment for non-returning function', () => {
      const code = 'Function GetValue() As Long\n    MsgBox "Hello"\nEnd Function\n';
      const decls = getDeclarationsWithRefs(code);
      const func = findByName(decls, 'GetValue');
      expect(func).toBeDefined();
      const assignmentRefs = func!.references.filter(r => r.isAssignment);
      expect(assignmentRefs).toHaveLength(0);
    });
  });

  describe('Unused declarations', () => {
    it('variable with no references is unused', () => {
      const code = 'Sub Test()\n    Dim x As Long\nEnd Sub\n';
      const decls = getDeclarationsWithRefs(code);
      const x = findByName(decls, 'x');
      expect(x).toBeDefined();
      expect(x!.references).toHaveLength(0);
    });

    it('parameter with no references is unused', () => {
      const code = 'Sub Test(x As Long)\n    MsgBox "Hello"\nEnd Sub\n';
      const decls = getDeclarationsWithRefs(code);
      const x = findByName(decls, 'x');
      expect(x).toBeDefined();
      expect(x!.references).toHaveLength(0);
    });
  });
});
