// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/VariableNotUsedInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { DeclarationInspection, type InspectionContext, type InspectionMetadata } from '../base.js';
import type { InspectionResult } from '../types.js';

/**
 * Detects variables that are declared (Dim) but never referenced.
 *
 * VBA:
 *   Dim x As Long  ' ← declared but never used
 *
 * Rubberduck rationale: Unused variables clutter code and may indicate
 * incomplete implementation or leftover refactoring artifacts.
 */
export class VariableNotUsedInspection extends DeclarationInspection {
  static readonly meta: InspectionMetadata = {
    id: 'VariableNotUsed',
    tier: 'B',
    category: 'CodeQuality',
    defaultSeverity: 'warning',
    name: 'Variable is not used',
    description: 'Variable is declared but never referenced.',
    quickFixDescription: 'Remove the unused variable declaration',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const finder = context.declarationFinder;
    if (!finder) return [];

    const results: InspectionResult[] = [];

    for (const decl of finder.findUnusedVariables()) {
      results.push(
        this.createResult(decl.location, {
          description: `Variable '${decl.name}' is declared but never used.`,
        }),
      );
    }

    return results;
  }
}
