// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { describe, it, expect } from 'vitest';
import { parseCode } from '../../../parser/index.js';
import type { InspectionContext } from '../../base.js';
import { buildSingleModuleFinder } from '../../../symbols/workspace.js';
import { readFile } from 'node:fs/promises';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

import { ObjectVariableNotSetInspection } from '../object-variable-not-set.js';
import { IntegerDataTypeInspection } from '../integer-data-type.js';
import { VariableTypeNotDeclaredInspection } from '../variable-type-not-declared.js';

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

// --- ObjectVariableNotSet ---
describe('ObjectVariableNotSetInspection', () => {
  it('detects assignment to object variable', () => {
    const code = 'Sub Test()\n    Dim ws As Worksheet\n    ws = ActiveSheet\nEnd Sub\n';
    const results = inspectCodeWithSymbols(ObjectVariableNotSetInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('ObjectVariableNotSet');
    expect(results[0].description).toContain('ws');
  });

  it('detects assignment to Collection variable', () => {
    const code = 'Sub Test()\n    Dim col As Collection\n    col = New Collection\nEnd Sub\n';
    const results = inspectCodeWithSymbols(ObjectVariableNotSetInspection, code);
    expect(results).toHaveLength(1);
  });

  it('does not flag non-object variable assignment', () => {
    const code = 'Sub Test()\n    Dim x As Long\n    x = 42\nEnd Sub\n';
    const results = inspectCodeWithSymbols(ObjectVariableNotSetInspection, code);
    expect(results).toHaveLength(0);
  });

  it('does not flag object variable with no assignment', () => {
    const code = 'Sub Test()\n    Dim ws As Worksheet\n    MsgBox ws.Name\nEnd Sub\n';
    const results = inspectCodeWithSymbols(ObjectVariableNotSetInspection, code);
    expect(results).toHaveLength(0);
  });

  it('has correct metadata', () => {
    const meta = ObjectVariableNotSetInspection.meta;
    expect(meta.id).toBe('ObjectVariableNotSet');
    expect(meta.tier).toBe('B');
    expect(meta.defaultSeverity).toBe('error');
  });
});

// --- IntegerDataType ---
describe('IntegerDataTypeInspection', () => {
  it('detects variable declared As Integer', () => {
    const code = 'Sub Test()\n    Dim count As Integer\n    count = 1\n    MsgBox CStr(count)\nEnd Sub\n';
    const results = inspectCodeWithSymbols(IntegerDataTypeInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('IntegerDataType');
    expect(results[0].description).toContain('count');
  });

  it('does not flag variable declared As Long', () => {
    const code = 'Sub Test()\n    Dim count As Long\n    count = 1\n    MsgBox CStr(count)\nEnd Sub\n';
    const results = inspectCodeWithSymbols(IntegerDataTypeInspection, code);
    expect(results).toHaveLength(0);
  });

  it('detects from fixture', async () => {
    const code = await readFile(join(fixturesDir, 'integer-variables.bas'), 'utf-8');
    const results = inspectCodeWithSymbols(IntegerDataTypeInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].description).toContain('count');
  });

  it('also detects Integer parameters', () => {
    const code = 'Sub Test(x As Integer)\n    MsgBox CStr(x)\nEnd Sub\n';
    const results = inspectCodeWithSymbols(IntegerDataTypeInspection, code);
    expect(results).toHaveLength(1);
  });

  it('has correct metadata', () => {
    const meta = IntegerDataTypeInspection.meta;
    expect(meta.id).toBe('IntegerDataType');
    expect(meta.tier).toBe('B');
    expect(meta.category).toBe('LanguageOpportunities');
    expect(meta.defaultSeverity).toBe('suggestion');
  });
});

// --- VariableTypeNotDeclared ---
describe('VariableTypeNotDeclaredInspection', () => {
  it('detects variable without As clause', () => {
    const code = 'Sub Test()\n    Dim value\n    value = 42\n    MsgBox CStr(value)\nEnd Sub\n';
    const results = inspectCodeWithSymbols(VariableTypeNotDeclaredInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('VariableTypeNotDeclared');
    expect(results[0].description).toContain('value');
  });

  it('does not flag variable with explicit type', () => {
    const code = 'Sub Test()\n    Dim count As Long\n    count = 1\n    MsgBox CStr(count)\nEnd Sub\n';
    const results = inspectCodeWithSymbols(VariableTypeNotDeclaredInspection, code);
    expect(results).toHaveLength(0);
  });

  it('detects from fixture', async () => {
    const code = await readFile(join(fixturesDir, 'implicit-type.bas'), 'utf-8');
    const results = inspectCodeWithSymbols(VariableTypeNotDeclaredInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].description).toContain('value');
  });

  it('has correct metadata', () => {
    const meta = VariableTypeNotDeclaredInspection.meta;
    expect(meta.id).toBe('VariableTypeNotDeclared');
    expect(meta.tier).toBe('B');
    expect(meta.category).toBe('LanguageOpportunities');
    expect(meta.defaultSeverity).toBe('suggestion');
  });
});
