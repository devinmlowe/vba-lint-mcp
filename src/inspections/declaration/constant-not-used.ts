// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/ConstantNotUsedInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { DeclarationInspection, type InspectionContext, type InspectionMetadata } from '../base.js';
import type { InspectionResult } from '../types.js';

/**
 * Detects constants that are declared but never referenced.
 *
 * VBA:
 *   Const MAX_ITEMS As Long = 100  ' ← declared but never used
 *
 * Rubberduck rationale: Unused constants clutter code and may indicate
 * incomplete implementation or leftover refactoring artifacts.
 */
export class ConstantNotUsedInspection extends DeclarationInspection {
  static readonly meta: InspectionMetadata = {
    id: 'ConstantNotUsed',
    tier: 'B',
    category: 'CodeQuality',
    defaultSeverity: 'warning',
    name: 'Constant is not used',
    description: 'Constant is declared but never referenced.',
    quickFixDescription: 'Remove the unused constant declaration',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const finder = context.declarationFinder;
    if (!finder) return [];

    const results: InspectionResult[] = [];

    for (const decl of finder.findByType('Constant')) {
      if (decl.references.length === 0) {
        results.push(
          this.createResult(decl.location, {
            description: `Constant '${decl.name}' is declared but never used.`,
          }),
        );
      }
    }

    return results;
  }
}
