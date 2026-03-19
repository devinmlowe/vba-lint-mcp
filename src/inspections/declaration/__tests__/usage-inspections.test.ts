// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { describe, it, expect } from 'vitest';
import { parseCode } from '../../../parser/index.js';
import type { InspectionContext } from '../../base.js';
import { buildSingleModuleFinder } from '../../../symbols/workspace.js';

import { UnassignedVariableUsageInspection } from '../unassigned-variable-usage.js';

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

// --- UnassignedVariableUsage ---
describe('UnassignedVariableUsageInspection', () => {
  it('detects use of unassigned variable at each usage site', () => {
    const code = 'Sub Test()\n    Dim x As Long\n    MsgBox CStr(x)\nEnd Sub\n';
    const results = inspectCodeWithSymbols(UnassignedVariableUsageInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('UnassignedVariableUsage');
    expect(results[0].description).toContain('x');
  });

  it('detects multiple usage sites of unassigned variable', () => {
    const code = 'Sub Test()\n    Dim x As Long\n    MsgBox CStr(x)\n    MsgBox CStr(x)\nEnd Sub\n';
    const results = inspectCodeWithSymbols(UnassignedVariableUsageInspection, code);
    expect(results).toHaveLength(2);
  });

  it('does not flag variable that is assigned before use', () => {
    const code = 'Sub Test()\n    Dim x As Long\n    x = 42\n    MsgBox CStr(x)\nEnd Sub\n';
    const results = inspectCodeWithSymbols(UnassignedVariableUsageInspection, code);
    expect(results).toHaveLength(0);
  });

  it('does not flag completely unused variable', () => {
    const code = 'Sub Test()\n    Dim x As Long\nEnd Sub\n';
    const results = inspectCodeWithSymbols(UnassignedVariableUsageInspection, code);
    expect(results).toHaveLength(0);
  });

  it('does not flag variable that is only assigned', () => {
    const code = 'Sub Test()\n    Dim x As Long\n    x = 42\nEnd Sub\n';
    const results = inspectCodeWithSymbols(UnassignedVariableUsageInspection, code);
    expect(results).toHaveLength(0);
  });

  it('has correct metadata', () => {
    const meta = UnassignedVariableUsageInspection.meta;
    expect(meta.id).toBe('UnassignedVariableUsage');
    expect(meta.tier).toBe('B');
    expect(meta.category).toBe('CodeQuality');
    expect(meta.defaultSeverity).toBe('warning');
  });

  it('returns empty when no declaration finder', () => {
    const code = 'Sub Test()\n    Dim x As Long\n    MsgBox CStr(x)\nEnd Sub\n';
    const parseResult = parseCode(code);
    const context: InspectionContext = { parseResult };
    const inspection = new UnassignedVariableUsageInspection();
    expect(inspection.inspect(context)).toHaveLength(0);
  });
});
