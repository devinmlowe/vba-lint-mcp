// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { describe, it, expect } from 'vitest';
import { parseCode } from '../../../parser/index.js';
import type { InspectionContext } from '../../base.js';
import { buildSingleModuleFinder } from '../../../symbols/workspace.js';

import { FunctionReturnValueNotUsedInspection } from '../function-return-value-not-used.js';
import { FunctionReturnValueAlwaysDiscardedInspection } from '../function-return-value-always-discarded.js';
import { ProcedureCanBeWrittenAsFunctionInspection } from '../procedure-can-be-written-as-function.js';

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

// --- FunctionReturnValueNotUsed (stub — returns empty for now) ---
describe('FunctionReturnValueNotUsedInspection', () => {
  it('has correct metadata', () => {
    const meta = FunctionReturnValueNotUsedInspection.meta;
    expect(meta.id).toBe('FunctionReturnValueNotUsed');
    expect(meta.tier).toBe('B');
    expect(meta.category).toBe('CodeQuality');
    expect(meta.defaultSeverity).toBe('suggestion');
  });

  it('returns empty for now (needs parse-tree augmentation)', () => {
    const code = 'Function GetValue() As Long\n    GetValue = 42\nEnd Function\n\nSub Test()\n    GetValue\nEnd Sub\n';
    const results = inspectCodeWithSymbols(FunctionReturnValueNotUsedInspection, code);
    // Currently a stub — will be implemented when call-site context is tracked
    expect(results).toHaveLength(0);
  });

  it('returns empty when no declaration finder', () => {
    const code = 'Function GetValue() As Long\n    GetValue = 42\nEnd Function\n';
    const parseResult = parseCode(code);
    const context: InspectionContext = { parseResult };
    const inspection = new FunctionReturnValueNotUsedInspection();
    expect(inspection.inspect(context)).toHaveLength(0);
  });
});

// --- FunctionReturnValueAlwaysDiscarded (stub — returns empty for now) ---
describe('FunctionReturnValueAlwaysDiscardedInspection', () => {
  it('has correct metadata', () => {
    const meta = FunctionReturnValueAlwaysDiscardedInspection.meta;
    expect(meta.id).toBe('FunctionReturnValueAlwaysDiscarded');
    expect(meta.tier).toBe('B');
    expect(meta.category).toBe('CodeQuality');
    expect(meta.defaultSeverity).toBe('suggestion');
  });

  it('returns empty for now (needs call-site context)', () => {
    const code = 'Function DoWork() As Boolean\n    DoWork = True\nEnd Function\n\nSub Caller()\n    DoWork\nEnd Sub\n';
    const results = inspectCodeWithSymbols(FunctionReturnValueAlwaysDiscardedInspection, code);
    expect(results).toHaveLength(0);
  });
});

// --- ProcedureCanBeWrittenAsFunction ---
describe('ProcedureCanBeWrittenAsFunctionInspection', () => {
  it('detects Sub that assigns to a ByRef parameter', () => {
    const code = 'Sub GetValue(ByRef result As Long)\n    result = 42\nEnd Sub\n';
    const results = inspectCodeWithSymbols(ProcedureCanBeWrittenAsFunctionInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('ProcedureCanBeWrittenAsFunction');
    expect(results[0].description).toContain('GetValue');
    expect(results[0].description).toContain('result');
  });

  it('does not flag Sub with no ByRef assignments', () => {
    const code = 'Sub DoWork(ByVal x As Long)\n    MsgBox CStr(x)\nEnd Sub\n';
    const results = inspectCodeWithSymbols(ProcedureCanBeWrittenAsFunctionInspection, code);
    expect(results).toHaveLength(0);
  });

  it('does not flag Sub with multiple ByRef assignments', () => {
    const code = 'Sub SwapValues(ByRef a As Long, ByRef b As Long)\n    a = 1\n    b = 2\nEnd Sub\n';
    const results = inspectCodeWithSymbols(ProcedureCanBeWrittenAsFunctionInspection, code);
    // Multiple assigned ByRef params — not a simple function conversion
    expect(results).toHaveLength(0);
  });

  it('does not flag Function', () => {
    const code = 'Function GetValue() As Long\n    GetValue = 42\nEnd Function\n';
    const results = inspectCodeWithSymbols(ProcedureCanBeWrittenAsFunctionInspection, code);
    expect(results).toHaveLength(0);
  });

  it('has correct metadata', () => {
    const meta = ProcedureCanBeWrittenAsFunctionInspection.meta;
    expect(meta.id).toBe('ProcedureCanBeWrittenAsFunction');
    expect(meta.tier).toBe('B');
    expect(meta.category).toBe('LanguageOpportunities');
    expect(meta.defaultSeverity).toBe('suggestion');
  });
});
