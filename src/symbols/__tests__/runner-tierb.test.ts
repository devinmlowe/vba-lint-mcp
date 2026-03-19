// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { describe, it, expect } from 'vitest';
import { parseCode } from '../../parser/index.js';
import { runInspections } from '../../inspections/runner.js';
import { createAllInspections } from '../../inspections/registry.js';
import { buildSingleModuleFinder } from '../workspace.js';

describe('Runner Tier B behavior', () => {
  it('skips Tier B inspections when no symbol table', () => {
    const code = 'Sub Test()\n    Dim x As Long\nEnd Sub\n';
    const parseResult = parseCode(code);
    const inspections = createAllInspections();
    const context = { parseResult };

    const { skipped } = runInspections(inspections, context, {
      hasSymbolTable: false,
    });

    // There should be skipped Tier B inspections
    const tierBSkipped = skipped.filter(s => s.reason.includes('Tier B'));
    expect(tierBSkipped.length).toBeGreaterThan(0);
  });

  it('includes Tier B results when symbol table present', () => {
    // Code with an unused variable should produce a VariableNotUsed result
    const code = 'Sub Test()\n    Dim x As Long\nEnd Sub\n';
    const parseResult = parseCode(code);
    const declarationFinder = buildSingleModuleFinder(parseResult);
    const inspections = createAllInspections();
    const context = { parseResult, declarationFinder };

    const { results, skipped } = runInspections(inspections, context, {
      hasSymbolTable: true,
    });

    // No Tier B inspections should be skipped
    const tierBSkipped = skipped.filter(s => s.reason.includes('Tier B'));
    expect(tierBSkipped).toHaveLength(0);

    // Should find the unused variable
    const variableNotUsed = results.filter(r => r.inspection === 'VariableNotUsed');
    expect(variableNotUsed).toHaveLength(1);
  });

  it('reports specific skipped inspection IDs', () => {
    const code = 'Sub Test()\nEnd Sub\n';
    const parseResult = parseCode(code);
    const inspections = createAllInspections();
    const context = { parseResult };

    const { skipped } = runInspections(inspections, context, {
      hasSymbolTable: false,
    });

    const skippedIds = skipped.map(s => s.inspection);
    expect(skippedIds).toContain('VariableNotUsed');
    expect(skippedIds).toContain('ParameterNotUsed');
    expect(skippedIds).toContain('NonReturningFunction');
  });
});
