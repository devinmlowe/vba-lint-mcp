// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/LineLabelNotUsedInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { DeclarationInspection, type InspectionContext, type InspectionMetadata } from '../base.js';
import type { InspectionResult } from '../types.js';

/**
 * Detects line labels that are declared but never referenced by GoTo or GoSub.
 *
 * VBA:
 *   Sub Test()
 *     MyLabel:             ' ← no GoTo MyLabel anywhere
 *     MsgBox "Hello"
 *   End Sub
 *
 * Rubberduck rationale: Unused labels clutter code and may indicate
 * leftover error-handling scaffolding.
 */
export class LineLabelNotUsedInspection extends DeclarationInspection {
  static readonly meta: InspectionMetadata = {
    id: 'LineLabelNotUsed',
    tier: 'B',
    category: 'CodeQuality',
    defaultSeverity: 'suggestion',
    name: 'Line label is not used',
    description: 'Line label is declared but never referenced by GoTo or GoSub.',
    quickFixDescription: 'Remove the unused line label',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const finder = context.declarationFinder;
    if (!finder) return [];

    const results: InspectionResult[] = [];

    for (const decl of finder.findByType('LineLabel')) {
      if (decl.references.length === 0) {
        results.push(
          this.createResult(decl.location, {
            description: `Line label '${decl.name}' is never referenced.`,
          }),
        );
      }
    }

    return results;
  }
}
