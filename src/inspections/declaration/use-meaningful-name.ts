// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/UseMeaningfulNameInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { DeclarationInspection, type InspectionContext, type InspectionMetadata } from '../base.js';
import type { InspectionResult } from '../types.js';

/** Common single-character loop variable names that are acceptable. */
const ALLOWED_SHORT_NAMES = new Set(['i', 'j', 'k', 'n', 'x', 'y', 'z']);

/**
 * Detects variables with very short (1-2 character) names that are not
 * common loop variables.
 *
 * VBA:
 *   Dim a As Long      ' ← too short, not meaningful
 *   Dim fn As String   ' ← too short
 *
 * Allowed: i, j, k, n, x, y, z (common loop/coordinate variables)
 */
export class UseMeaningfulNameInspection extends DeclarationInspection {
  static readonly meta: InspectionMetadata = {
    id: 'UseMeaningfulName',
    tier: 'B',
    category: 'Naming',
    defaultSeverity: 'suggestion',
    name: 'Use a meaningful name',
    description: 'Variable name is too short to be meaningful.',
    quickFixDescription: 'Rename to a more descriptive name',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const finder = context.declarationFinder;
    if (!finder) return [];

    const results: InspectionResult[] = [];

    for (const decl of finder.getAll()) {
      // Only check variables, constants, and parameters
      if (decl.declarationType !== 'Variable'
        && decl.declarationType !== 'Constant'
        && decl.declarationType !== 'Parameter') continue;

      if (decl.name.length <= 2) {
        // Allow common loop variable names
        if (ALLOWED_SHORT_NAMES.has(decl.name.toLowerCase())) continue;

        results.push(
          this.createResult(decl.location, {
            description: `'${decl.name}' is too short to be a meaningful name.`,
          }),
        );
      }
    }

    return results;
  }
}
