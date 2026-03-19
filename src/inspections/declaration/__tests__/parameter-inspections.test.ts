// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { describe, it, expect } from 'vitest';
import { parseCode } from '../../../parser/index.js';
import type { InspectionContext } from '../../base.js';
import { buildSingleModuleFinder } from '../../../symbols/workspace.js';
import { readFile } from 'node:fs/promises';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

import { ExcessiveParametersInspection } from '../excessive-parameters.js';
import { ParameterCanBeByValInspection } from '../parameter-can-be-byval.js';

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

// --- ExcessiveParameters ---
describe('ExcessiveParametersInspection', () => {
  it('detects procedure with 8 parameters', () => {
    const code = 'Sub TooMany(a As Long, b As Long, c As Long, d As Long, e As Long, f As Long, g As Long, h As Long)\n    MsgBox CStr(a)\nEnd Sub\n';
    const results = inspectCodeWithSymbols(ExcessiveParametersInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('ExcessiveParameters');
    expect(results[0].description).toContain('TooMany');
    expect(results[0].description).toContain('8');
  });

  it('does not flag procedure with 7 parameters', () => {
    const code = 'Sub JustRight(a As Long, b As Long, c As Long, d As Long, e As Long, f As Long, g As Long)\n    MsgBox CStr(a)\nEnd Sub\n';
    const results = inspectCodeWithSymbols(ExcessiveParametersInspection, code);
    expect(results).toHaveLength(0);
  });

  it('does not flag procedure with 3 parameters', () => {
    const code = 'Sub Small(a As Long, b As Long, c As Long)\n    MsgBox CStr(a)\nEnd Sub\n';
    const results = inspectCodeWithSymbols(ExcessiveParametersInspection, code);
    expect(results).toHaveLength(0);
  });

  it('detects from fixture', async () => {
    const code = await readFile(join(fixturesDir, 'excessive-params.bas'), 'utf-8');
    const results = inspectCodeWithSymbols(ExcessiveParametersInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].description).toContain('TooMany');
  });

  it('has correct metadata', () => {
    const meta = ExcessiveParametersInspection.meta;
    expect(meta.id).toBe('ExcessiveParameters');
    expect(meta.tier).toBe('B');
    expect(meta.category).toBe('CodeQuality');
    expect(meta.defaultSeverity).toBe('suggestion');
  });
});

// --- ParameterCanBeByVal ---
describe('ParameterCanBeByValInspection', () => {
  it('detects ByRef parameter never assigned to', () => {
    const code = 'Sub DoWork(ByRef x As Long)\n    MsgBox CStr(x)\nEnd Sub\n';
    const results = inspectCodeWithSymbols(ParameterCanBeByValInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('ParameterCanBeByVal');
    expect(results[0].description).toContain('x');
  });

  it('does not flag ByRef parameter that is assigned', () => {
    const code = 'Sub DoWork(ByRef x As Long)\n    x = 42\nEnd Sub\n';
    const results = inspectCodeWithSymbols(ParameterCanBeByValInspection, code);
    expect(results).toHaveLength(0);
  });

  it('does not flag ByVal parameter', () => {
    const code = 'Sub DoWork(ByVal x As Long)\n    MsgBox CStr(x)\nEnd Sub\n';
    const results = inspectCodeWithSymbols(ParameterCanBeByValInspection, code);
    expect(results).toHaveLength(0);
  });

  it('has correct metadata', () => {
    const meta = ParameterCanBeByValInspection.meta;
    expect(meta.id).toBe('ParameterCanBeByVal');
    expect(meta.tier).toBe('B');
    expect(meta.category).toBe('CodeQuality');
    expect(meta.defaultSeverity).toBe('suggestion');
  });
});
