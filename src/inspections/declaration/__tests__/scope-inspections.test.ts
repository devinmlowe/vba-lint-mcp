// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { describe, it, expect } from 'vitest';
import { parseCode } from '../../../parser/index.js';
import type { InspectionContext } from '../../base.js';
import { buildSingleModuleFinder } from '../../../symbols/workspace.js';

import { ModuleScopeDimKeywordInspection } from '../module-scope-dim-keyword.js';
import { EncapsulatePublicFieldInspection } from '../encapsulate-public-field.js';
import { MoveFieldCloserToUsageInspection } from '../move-field-closer-to-usage.js';

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

// --- ModuleScopeDimKeyword ---
describe('ModuleScopeDimKeywordInspection', () => {
  it('detects module-level Dim', () => {
    const code = 'Dim moduleVar As Long\n\nSub Test()\n    moduleVar = 42\n    MsgBox CStr(moduleVar)\nEnd Sub\n';
    const results = inspectCodeWithSymbols(ModuleScopeDimKeywordInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('ModuleScopeDimKeyword');
    expect(results[0].description).toContain('moduleVar');
  });

  it('does not flag Private at module level', () => {
    const code = 'Private moduleVar As Long\n\nSub Test()\n    moduleVar = 42\n    MsgBox CStr(moduleVar)\nEnd Sub\n';
    const results = inspectCodeWithSymbols(ModuleScopeDimKeywordInspection, code);
    expect(results).toHaveLength(0);
  });

  it('does not flag Dim inside a procedure', () => {
    const code = 'Sub Test()\n    Dim x As Long\n    x = 1\n    MsgBox CStr(x)\nEnd Sub\n';
    const results = inspectCodeWithSymbols(ModuleScopeDimKeywordInspection, code);
    expect(results).toHaveLength(0);
  });

  it('has correct metadata', () => {
    const meta = ModuleScopeDimKeywordInspection.meta;
    expect(meta.id).toBe('ModuleScopeDimKeyword');
    expect(meta.tier).toBe('B');
    expect(meta.defaultSeverity).toBe('suggestion');
  });
});

// --- EncapsulatePublicField ---
describe('EncapsulatePublicFieldInspection', () => {
  it('detects public module-level variable', () => {
    const code = 'Public Name As String\n\nSub Test()\n    Name = "Hello"\n    MsgBox Name\nEnd Sub\n';
    const results = inspectCodeWithSymbols(EncapsulatePublicFieldInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('EncapsulatePublicField');
    expect(results[0].description).toContain('Name');
  });

  it('does not flag private module-level variable', () => {
    const code = 'Private Name As String\n\nSub Test()\n    Name = "Hello"\n    MsgBox Name\nEnd Sub\n';
    const results = inspectCodeWithSymbols(EncapsulatePublicFieldInspection, code);
    expect(results).toHaveLength(0);
  });

  it('does not flag local variable', () => {
    const code = 'Sub Test()\n    Dim x As Long\n    x = 1\n    MsgBox CStr(x)\nEnd Sub\n';
    const results = inspectCodeWithSymbols(EncapsulatePublicFieldInspection, code);
    expect(results).toHaveLength(0);
  });

  it('has correct metadata', () => {
    const meta = EncapsulatePublicFieldInspection.meta;
    expect(meta.id).toBe('EncapsulatePublicField');
    expect(meta.tier).toBe('B');
    expect(meta.defaultSeverity).toBe('suggestion');
  });
});

// --- MoveFieldCloserToUsage ---
describe('MoveFieldCloserToUsageInspection', () => {
  it('detects module variable used in only one procedure', () => {
    const code = 'Private tempVar As Long\n\nSub OnlyUser()\n    tempVar = 42\n    MsgBox CStr(tempVar)\nEnd Sub\n\nSub Other()\n    MsgBox "other"\nEnd Sub\n';
    const results = inspectCodeWithSymbols(MoveFieldCloserToUsageInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('MoveFieldCloserToUsage');
    expect(results[0].description).toContain('tempVar');
    expect(results[0].description).toContain('OnlyUser');
  });

  it('does not flag module variable used in multiple procedures', () => {
    const code = 'Private counter As Long\n\nSub Increment()\n    counter = counter + 1\nEnd Sub\n\nSub Display()\n    MsgBox CStr(counter)\nEnd Sub\n';
    const results = inspectCodeWithSymbols(MoveFieldCloserToUsageInspection, code);
    expect(results).toHaveLength(0);
  });

  it('does not flag unused module variable', () => {
    const code = 'Private unused As Long\n\nSub Test()\n    MsgBox "Hello"\nEnd Sub\n';
    const results = inspectCodeWithSymbols(MoveFieldCloserToUsageInspection, code);
    expect(results).toHaveLength(0);
  });

  it('has correct metadata', () => {
    const meta = MoveFieldCloserToUsageInspection.meta;
    expect(meta.id).toBe('MoveFieldCloserToUsage');
    expect(meta.tier).toBe('B');
    expect(meta.defaultSeverity).toBe('suggestion');
  });
});
