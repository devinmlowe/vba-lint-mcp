// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { describe, it, expect } from 'vitest';
import { parseCode } from '../../../parser/index.js';
import type { InspectionContext } from '../../base.js';
import { buildSingleModuleFinder } from '../../../symbols/workspace.js';
import { readFile } from 'node:fs/promises';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

import { HungarianNotationInspection } from '../hungarian-notation.js';
import { UseMeaningfulNameInspection } from '../use-meaningful-name.js';
import { UnderscoreInPublicClassModuleMemberInspection } from '../underscore-in-public-class-module-member.js';

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

// --- HungarianNotation ---
describe('HungarianNotationInspection', () => {
  it('detects strName as Hungarian', () => {
    const code = 'Sub Test()\n    Dim strName As String\n    strName = "x"\n    MsgBox strName\nEnd Sub\n';
    const results = inspectCodeWithSymbols(HungarianNotationInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('HungarianNotation');
    expect(results[0].description).toContain('strName');
    expect(results[0].description).toContain('str');
  });

  it('detects intCount as Hungarian', () => {
    const code = 'Sub Test()\n    Dim intCount As Integer\n    intCount = 1\n    MsgBox CStr(intCount)\nEnd Sub\n';
    const results = inspectCodeWithSymbols(HungarianNotationInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].description).toContain('int');
  });

  it('does not flag names without Hungarian prefix', () => {
    const code = 'Sub Test()\n    Dim userName As String\n    userName = "x"\n    MsgBox userName\nEnd Sub\n';
    const results = inspectCodeWithSymbols(HungarianNotationInspection, code);
    expect(results).toHaveLength(0);
  });

  it('does not flag prefix without capital letter after', () => {
    const code = 'Sub Test()\n    Dim string As String\n    string = "x"\n    MsgBox string\nEnd Sub\n';
    const results = inspectCodeWithSymbols(HungarianNotationInspection, code);
    expect(results).toHaveLength(0);
  });

  it('detects multiple Hungarian names from fixture', async () => {
    const code = await readFile(join(fixturesDir, 'hungarian-notation.bas'), 'utf-8');
    const results = inspectCodeWithSymbols(HungarianNotationInspection, code);
    expect(results).toHaveLength(3); // strName, intCount, lngTotal
  });

  it('also checks parameters', () => {
    const code = 'Sub Test(strValue As String)\n    MsgBox strValue\nEnd Sub\n';
    const results = inspectCodeWithSymbols(HungarianNotationInspection, code);
    expect(results).toHaveLength(1);
  });

  it('has correct metadata', () => {
    const meta = HungarianNotationInspection.meta;
    expect(meta.id).toBe('HungarianNotation');
    expect(meta.tier).toBe('B');
    expect(meta.category).toBe('Naming');
    expect(meta.defaultSeverity).toBe('suggestion');
  });
});

// --- UseMeaningfulName ---
describe('UseMeaningfulNameInspection', () => {
  it('detects single-character non-loop variable', () => {
    const code = 'Sub Test()\n    Dim a As Long\n    a = 1\n    MsgBox CStr(a)\nEnd Sub\n';
    const results = inspectCodeWithSymbols(UseMeaningfulNameInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('UseMeaningfulName');
  });

  it('detects two-character variable name', () => {
    const code = 'Sub Test()\n    Dim fn As String\n    fn = "x"\n    MsgBox fn\nEnd Sub\n';
    const results = inspectCodeWithSymbols(UseMeaningfulNameInspection, code);
    expect(results).toHaveLength(1);
  });

  it('allows i as loop variable', () => {
    const code = 'Sub Test()\n    Dim i As Long\n    i = 1\n    MsgBox CStr(i)\nEnd Sub\n';
    const results = inspectCodeWithSymbols(UseMeaningfulNameInspection, code);
    expect(results).toHaveLength(0);
  });

  it('allows j, k, n, x, y, z', () => {
    const names = ['j', 'k', 'n', 'x', 'y', 'z'];
    for (const name of names) {
      const code = `Sub Test()\n    Dim ${name} As Long\n    ${name} = 1\n    MsgBox CStr(${name})\nEnd Sub\n`;
      const results = inspectCodeWithSymbols(UseMeaningfulNameInspection, code);
      expect(results).toHaveLength(0);
    }
  });

  it('does not flag three-character names', () => {
    const code = 'Sub Test()\n    Dim cnt As Long\n    cnt = 1\n    MsgBox CStr(cnt)\nEnd Sub\n';
    const results = inspectCodeWithSymbols(UseMeaningfulNameInspection, code);
    expect(results).toHaveLength(0);
  });

  it('has correct metadata', () => {
    const meta = UseMeaningfulNameInspection.meta;
    expect(meta.id).toBe('UseMeaningfulName');
    expect(meta.tier).toBe('B');
    expect(meta.category).toBe('Naming');
  });
});

// --- UnderscoreInPublicClassModuleMember ---
describe('UnderscoreInPublicClassModuleMemberInspection', () => {
  it('detects public sub with underscore', () => {
    const code = 'Public Sub My_Method()\n    MsgBox "Hello"\nEnd Sub\n';
    const results = inspectCodeWithSymbols(UnderscoreInPublicClassModuleMemberInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('UnderscoreInPublicClassModuleMember');
    expect(results[0].description).toContain('My_Method');
  });

  it('does not flag private sub with underscore', () => {
    const code = 'Private Sub My_Method()\n    MsgBox "Hello"\nEnd Sub\n';
    const results = inspectCodeWithSymbols(UnderscoreInPublicClassModuleMemberInspection, code);
    expect(results).toHaveLength(0);
  });

  it('does not flag public sub without underscore', () => {
    const code = 'Public Sub DoWork()\n    MsgBox "Hello"\nEnd Sub\n';
    const results = inspectCodeWithSymbols(UnderscoreInPublicClassModuleMemberInspection, code);
    expect(results).toHaveLength(0);
  });

  it('has correct metadata', () => {
    const meta = UnderscoreInPublicClassModuleMemberInspection.meta;
    expect(meta.id).toBe('UnderscoreInPublicClassModuleMember');
    expect(meta.tier).toBe('B');
    expect(meta.category).toBe('Naming');
    expect(meta.defaultSeverity).toBe('warning');
  });
});
