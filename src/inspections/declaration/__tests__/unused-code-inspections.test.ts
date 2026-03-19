// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { describe, it, expect } from 'vitest';
import { parseCode } from '../../../parser/index.js';
import type { InspectionContext } from '../../base.js';
import { buildSingleModuleFinder } from '../../../symbols/workspace.js';
import { readFile } from 'node:fs/promises';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

import { ConstantNotUsedInspection } from '../constant-not-used.js';
import { ProcedureNotUsedInspection } from '../procedure-not-used.js';
import { LineLabelNotUsedInspection } from '../line-label-not-used.js';
import { VariableNotAssignedInspection } from '../variable-not-assigned.js';

const __dirname = dirname(fileURLToPath(import.meta.url));
const fixturesDir = join(__dirname, '..', '__fixtures__');

function inspectCodeWithSymbols<T extends { inspect(ctx: InspectionContext): any }>(
  InspectionClass: new () => T,
  code: string,
) {
  const parseResult = parseCode(code);
  const declarationFinder = buildSingleModuleFinder(parseResult);
  const context: InspectionContext = { parseResult, declarationFinder };
  const inspection = new InspectionClass();
  return inspection.inspect(context);
}

// --- ConstantNotUsed ---
describe('ConstantNotUsedInspection', () => {
  it('detects unused constant', () => {
    const code = 'Sub Test()\n    Const MAX As Long = 100\nEnd Sub\n';
    const results = inspectCodeWithSymbols(ConstantNotUsedInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('ConstantNotUsed');
    expect(results[0].description).toContain('MAX');
  });

  it('does not flag used constant', () => {
    const code = 'Sub Test()\n    Const MAX As Long = 100\n    MsgBox CStr(MAX)\nEnd Sub\n';
    const results = inspectCodeWithSymbols(ConstantNotUsedInspection, code);
    expect(results).toHaveLength(0);
  });

  it('detects from fixture (one unused out of two)', async () => {
    const code = await readFile(join(fixturesDir, 'unused-constant.bas'), 'utf-8');
    const results = inspectCodeWithSymbols(ConstantNotUsedInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].description).toContain('MAX_ITEMS');
  });

  it('has correct metadata', () => {
    const meta = ConstantNotUsedInspection.meta;
    expect(meta.id).toBe('ConstantNotUsed');
    expect(meta.tier).toBe('B');
    expect(meta.category).toBe('CodeQuality');
    expect(meta.defaultSeverity).toBe('warning');
  });

  it('returns empty when no declaration finder', () => {
    const code = 'Sub Test()\n    Const X As Long = 1\nEnd Sub\n';
    const parseResult = parseCode(code);
    const context: InspectionContext = { parseResult };
    const inspection = new ConstantNotUsedInspection();
    expect(inspection.inspect(context)).toHaveLength(0);
  });
});

// --- ProcedureNotUsed ---
describe('ProcedureNotUsedInspection', () => {
  it('detects unused private Sub', () => {
    const code = 'Private Sub Helper()\n    MsgBox "Hello"\nEnd Sub\n';
    const results = inspectCodeWithSymbols(ProcedureNotUsedInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('ProcedureNotUsed');
    expect(results[0].description).toContain('Helper');
  });

  it('does not flag public Sub', () => {
    const code = 'Public Sub DoWork()\n    MsgBox "Hello"\nEnd Sub\n';
    const results = inspectCodeWithSymbols(ProcedureNotUsedInspection, code);
    expect(results).toHaveLength(0);
  });

  it('does not flag event handler (name with underscore)', () => {
    const code = 'Private Sub Worksheet_Change()\n    MsgBox "Changed"\nEnd Sub\n';
    const results = inspectCodeWithSymbols(ProcedureNotUsedInspection, code);
    expect(results).toHaveLength(0);
  });

  it('does not flag called private Sub', () => {
    const code = 'Private Sub Helper()\n    MsgBox "Hello"\nEnd Sub\n\nSub Main()\n    Helper\nEnd Sub\n';
    const results = inspectCodeWithSymbols(ProcedureNotUsedInspection, code);
    expect(results).toHaveLength(0);
  });

  it('detects from fixture (one unused out of two)', async () => {
    const code = await readFile(join(fixturesDir, 'unused-procedure.bas'), 'utf-8');
    const results = inspectCodeWithSymbols(ProcedureNotUsedInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].description).toContain('Unused');
  });

  it('has correct metadata', () => {
    const meta = ProcedureNotUsedInspection.meta;
    expect(meta.id).toBe('ProcedureNotUsed');
    expect(meta.tier).toBe('B');
    expect(meta.defaultSeverity).toBe('suggestion');
  });
});

// --- LineLabelNotUsed ---
describe('LineLabelNotUsedInspection', () => {
  it('detects unused label', () => {
    const code = 'Sub Test()\nMyLabel:\n    MsgBox "Hello"\nEnd Sub\n';
    const results = inspectCodeWithSymbols(LineLabelNotUsedInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('LineLabelNotUsed');
    expect(results[0].description).toContain('MyLabel');
  });

  it('does not flag label referenced by GoTo', () => {
    const code = 'Sub Test()\n    GoTo MyLabel\nMyLabel:\n    MsgBox "Hello"\nEnd Sub\n';
    const results = inspectCodeWithSymbols(LineLabelNotUsedInspection, code);
    expect(results).toHaveLength(0);
  });

  it('has correct metadata', () => {
    const meta = LineLabelNotUsedInspection.meta;
    expect(meta.id).toBe('LineLabelNotUsed');
    expect(meta.tier).toBe('B');
    expect(meta.defaultSeverity).toBe('suggestion');
  });
});

// --- VariableNotAssigned ---
describe('VariableNotAssignedInspection', () => {
  it('detects variable read but never assigned', () => {
    const code = 'Sub Test()\n    Dim x As Long\n    MsgBox CStr(x)\nEnd Sub\n';
    const results = inspectCodeWithSymbols(VariableNotAssignedInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('VariableNotAssigned');
    expect(results[0].description).toContain('x');
  });

  it('does not flag variable that is assigned', () => {
    const code = 'Sub Test()\n    Dim x As Long\n    x = 42\n    MsgBox CStr(x)\nEnd Sub\n';
    const results = inspectCodeWithSymbols(VariableNotAssignedInspection, code);
    expect(results).toHaveLength(0);
  });

  it('does not flag completely unused variable (no reads)', () => {
    const code = 'Sub Test()\n    Dim x As Long\nEnd Sub\n';
    const results = inspectCodeWithSymbols(VariableNotAssignedInspection, code);
    expect(results).toHaveLength(0);
  });

  it('has correct metadata', () => {
    const meta = VariableNotAssignedInspection.meta;
    expect(meta.id).toBe('VariableNotAssigned');
    expect(meta.tier).toBe('B');
    expect(meta.defaultSeverity).toBe('warning');
  });
});
