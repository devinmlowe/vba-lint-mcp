// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { describe, it, expect } from 'vitest';
import { parseCode } from '../../../parser/index.js';
import type { InspectionContext } from '../../base.js';
import { buildSingleModuleFinder } from '../../../symbols/workspace.js';
import { readFile } from 'node:fs/promises';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

import { VariableNotUsedInspection } from '../variable-not-used.js';
import { ParameterNotUsedInspection } from '../parameter-not-used.js';
import { NonReturningFunctionInspection } from '../non-returning-function.js';

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

// --- VariableNotUsed ---
describe('VariableNotUsedInspection', () => {
  it('detects unused variable', () => {
    const code = 'Sub Test()\n    Dim x As Long\nEnd Sub\n';
    const results = inspectCodeWithSymbols(VariableNotUsedInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('VariableNotUsed');
    expect(results[0].description).toContain('x');
  });

  it('does not flag used variable', () => {
    const code = 'Sub Test()\n    Dim x As Long\n    x = 42\n    MsgBox x\nEnd Sub\n';
    const results = inspectCodeWithSymbols(VariableNotUsedInspection, code);
    expect(results).toHaveLength(0);
  });

  it('detects multiple unused variables', () => {
    const code = 'Sub Test()\n    Dim a As Long\n    Dim b As String\n    Dim c As Double\nEnd Sub\n';
    const results = inspectCodeWithSymbols(VariableNotUsedInspection, code);
    expect(results).toHaveLength(3);
  });

  it('does not flag variable that is only assigned', () => {
    // A variable that is assigned but never read should still be flagged
    // as "unused" if it has zero references. But assignment IS a reference.
    const code = 'Sub Test()\n    Dim x As Long\n    x = 42\nEnd Sub\n';
    const results = inspectCodeWithSymbols(VariableNotUsedInspection, code);
    // x has an assignment reference, so it's not completely unused
    expect(results).toHaveLength(0);
  });

  it('detects from fixture', async () => {
    const code = await readFile(join(fixturesDir, 'unused-variable.bas'), 'utf-8');
    const results = inspectCodeWithSymbols(VariableNotUsedInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].description).toContain('x');
  });

  it('no findings for fixture with all used', async () => {
    const code = await readFile(join(fixturesDir, 'all-variables-used.bas'), 'utf-8');
    const results = inspectCodeWithSymbols(VariableNotUsedInspection, code);
    expect(results).toHaveLength(0);
  });

  it('has correct metadata', () => {
    const meta = VariableNotUsedInspection.meta;
    expect(meta.id).toBe('VariableNotUsed');
    expect(meta.tier).toBe('B');
    expect(meta.category).toBe('CodeQuality');
    expect(meta.defaultSeverity).toBe('warning');
  });

  it('returns empty when no declaration finder', () => {
    const code = 'Sub Test()\n    Dim x As Long\nEnd Sub\n';
    const parseResult = parseCode(code);
    const context: InspectionContext = { parseResult };
    const inspection = new VariableNotUsedInspection();
    const results = inspection.inspect(context);
    expect(results).toHaveLength(0);
  });
});

// --- ParameterNotUsed ---
describe('ParameterNotUsedInspection', () => {
  it('detects unused parameter', () => {
    const code = 'Sub Test(x As Long)\n    MsgBox "Hello"\nEnd Sub\n';
    const results = inspectCodeWithSymbols(ParameterNotUsedInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('ParameterNotUsed');
    expect(results[0].description).toContain('x');
  });

  it('does not flag used parameter', () => {
    const code = 'Sub Test(x As Long)\n    MsgBox CStr(x)\nEnd Sub\n';
    const results = inspectCodeWithSymbols(ParameterNotUsedInspection, code);
    expect(results).toHaveLength(0);
  });

  it('detects one unused out of two parameters', async () => {
    const code = await readFile(join(fixturesDir, 'unused-parameter.bas'), 'utf-8');
    const results = inspectCodeWithSymbols(ParameterNotUsedInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].description).toContain('y');
  });

  it('has correct metadata', () => {
    const meta = ParameterNotUsedInspection.meta;
    expect(meta.id).toBe('ParameterNotUsed');
    expect(meta.tier).toBe('B');
    expect(meta.category).toBe('CodeQuality');
    expect(meta.defaultSeverity).toBe('suggestion');
  });
});

// --- NonReturningFunction ---
describe('NonReturningFunctionInspection', () => {
  it('detects function that never assigns return value', () => {
    const code = 'Function GetValue() As Long\n    MsgBox "Hello"\nEnd Function\n';
    const results = inspectCodeWithSymbols(NonReturningFunctionInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('NonReturningFunction');
    expect(results[0].description).toContain('GetValue');
  });

  it('does not flag function that assigns return value', () => {
    const code = 'Function GetValue() As Long\n    GetValue = 42\nEnd Function\n';
    const results = inspectCodeWithSymbols(NonReturningFunctionInspection, code);
    expect(results).toHaveLength(0);
  });

  it('does not flag Sub (not a function)', () => {
    const code = 'Sub DoWork()\n    MsgBox "Hello"\nEnd Sub\n';
    const results = inspectCodeWithSymbols(NonReturningFunctionInspection, code);
    expect(results).toHaveLength(0);
  });

  it('detects from fixture', async () => {
    const code = await readFile(join(fixturesDir, 'non-returning-function.bas'), 'utf-8');
    const results = inspectCodeWithSymbols(NonReturningFunctionInspection, code);
    expect(results).toHaveLength(1);
  });

  it('no findings for returning function fixture', async () => {
    const code = await readFile(join(fixturesDir, 'returning-function.bas'), 'utf-8');
    const results = inspectCodeWithSymbols(NonReturningFunctionInspection, code);
    expect(results).toHaveLength(0);
  });

  it('detects non-returning Property Get', () => {
    const code = 'Property Get Value() As Long\n    MsgBox "Hello"\nEnd Property\n';
    const results = inspectCodeWithSymbols(NonReturningFunctionInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].description).toContain('Property Get');
  });

  it('has correct metadata', () => {
    const meta = NonReturningFunctionInspection.meta;
    expect(meta.id).toBe('NonReturningFunction');
    expect(meta.tier).toBe('B');
    expect(meta.category).toBe('CodeQuality');
    expect(meta.defaultSeverity).toBe('warning');
  });
});
