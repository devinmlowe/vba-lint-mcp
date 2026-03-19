// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { describe, it, expect } from 'vitest';
import { parseCode } from '../../../parser/index.js';
import type { InspectionContext } from '../../base.js';
import { buildSingleModuleFinder } from '../../../symbols/workspace.js';
import { readFile } from 'node:fs/promises';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createAllInspections, validateRegistry } from '../../registry.js';

const __dirname = dirname(fileURLToPath(import.meta.url));
const fixturesDir = join(__dirname, '..', '__fixtures__');

function runAllTierBInspections(code: string) {
  const parseResult = parseCode(code);
  const declarationFinder = buildSingleModuleFinder(parseResult);
  const context: InspectionContext = { parseResult, declarationFinder };
  const allInspections = createAllInspections();
  const tierB = allInspections.filter(i => i.meta.tier === 'B');

  const results = [];
  for (const inspection of tierB) {
    results.push(...inspection.inspect(context));
  }
  return results;
}

describe('Tier B Interaction Tests', () => {
  it('registry validates with no errors', () => {
    const errors = validateRegistry();
    expect(errors).toHaveLength(0);
  });

  it('clean code fixture produces minimal findings', async () => {
    const code = await readFile(join(fixturesDir, 'clean-code.bas'), 'utf-8');
    const results = runAllTierBInspections(code);
    // Clean code should have very few or no findings
    // The clean-code fixture uses Private, explicit types, etc.
    const ids = results.map(r => r.inspection);
    // Should not have VariableNotUsed, ConstantNotUsed, etc.
    expect(ids).not.toContain('VariableNotUsed');
    expect(ids).not.toContain('ConstantNotUsed');
    expect(ids).not.toContain('NonReturningFunction');
  });

  it('multiple inspections can fire on same code', () => {
    // This variable is unused AND has a short name AND has no explicit type
    const code = 'Sub Test()\n    Dim a\nEnd Sub\n';
    const results = runAllTierBInspections(code);
    const ids = results.map(r => r.inspection);
    expect(ids).toContain('VariableNotUsed');
    expect(ids).toContain('UseMeaningfulName');
    expect(ids).toContain('VariableTypeNotDeclared');
  });

  it('Hungarian + Integer fires together', () => {
    const code = 'Sub Test()\n    Dim intCount As Integer\n    intCount = 1\n    MsgBox CStr(intCount)\nEnd Sub\n';
    const results = runAllTierBInspections(code);
    const ids = results.map(r => r.inspection);
    expect(ids).toContain('HungarianNotation');
    expect(ids).toContain('IntegerDataType');
  });

  it('VariableNotAssigned + UnassignedVariableUsage fire together', () => {
    const code = 'Sub Test()\n    Dim x As Long\n    MsgBox CStr(x)\nEnd Sub\n';
    const results = runAllTierBInspections(code);
    const ids = results.map(r => r.inspection);
    expect(ids).toContain('VariableNotAssigned');
    expect(ids).toContain('UnassignedVariableUsage');
  });

  it('all registered Tier B inspections have unique IDs', () => {
    const allInspections = createAllInspections();
    const tierB = allInspections.filter(i => i.meta.tier === 'B');
    const ids = tierB.map(i => i.meta.id);
    const uniqueIds = new Set(ids);
    expect(uniqueIds.size).toBe(ids.length);
  });

  it('all Tier B inspections return empty when no declarationFinder', () => {
    const code = 'Sub Test()\n    Dim x As Long\nEnd Sub\n';
    const parseResult = parseCode(code);
    const context: InspectionContext = { parseResult };
    const allInspections = createAllInspections();
    const tierB = allInspections.filter(i => i.meta.tier === 'B');

    for (const inspection of tierB) {
      const results = inspection.inspect(context);
      expect(results).toHaveLength(0);
    }
  });

  it('complex module exercises multiple inspections', () => {
    const code = [
      'Dim moduleVar As Long',             // ModuleScopeDimKeyword
      'Public pubField As String',          // EncapsulatePublicField
      '',
      'Private Sub Unused()',               // ProcedureNotUsed
      '    MsgBox "unused"',
      'End Sub',
      '',
      'Sub Main()',
      '    Dim strName As String',          // HungarianNotation
      '    Dim a As Integer',               // UseMeaningfulName + IntegerDataType
      '    Const DEAD_CONST As Long = 1',   // ConstantNotUsed
      '    strName = "hello"',
      '    a = 1',
      '    moduleVar = 42',
      '    MsgBox strName & CStr(a) & CStr(moduleVar)',
      'End Sub',
    ].join('\n');

    const results = runAllTierBInspections(code);
    const ids = results.map(r => r.inspection);

    expect(ids).toContain('ModuleScopeDimKeyword');
    expect(ids).toContain('EncapsulatePublicField');
    expect(ids).toContain('ProcedureNotUsed');
    expect(ids).toContain('HungarianNotation');
    expect(ids).toContain('UseMeaningfulName');
    expect(ids).toContain('IntegerDataType');
    expect(ids).toContain('ConstantNotUsed');
  });
});
