// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/VariableNotAssignedInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { DeclarationInspection, type InspectionContext, type InspectionMetadata } from '../base.js';
import type { InspectionResult } from '../types.js';

/**
 * Detects variables that are referenced but never assigned a value.
 *
 * VBA:
 *   Sub Test()
 *     Dim x As Long
 *     MsgBox x       ' ← x is read but never assigned
 *   End Sub
 *
 * Rubberduck rationale: A variable that is used but never assigned will
 * always have its default value (0, "", Nothing), which is likely a bug.
 */
export class VariableNotAssignedInspection extends DeclarationInspection {
  static readonly meta: InspectionMetadata = {
    id: 'VariableNotAssigned',
    tier: 'B',
    category: 'CodeQuality',
    defaultSeverity: 'warning',
    name: 'Variable is not assigned',
    description: 'Variable is referenced but never assigned a value.',
    quickFixDescription: 'Add an assignment or remove the variable',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const finder = context.declarationFinder;
    if (!finder) return [];

    const results: InspectionResult[] = [];

    for (const decl of finder.findVariablesNotAssigned()) {
      // Only flag variables that have at least one non-assignment reference
      // (i.e., they are read but never written to)
      const hasReadRef = decl.references.some(r => !r.isAssignment);
      if (hasReadRef) {
        results.push(
          this.createResult(decl.location, {
            description: `Variable '${decl.name}' is used but never assigned a value.`,
          }),
        );
      }
    }

    return results;
  }
}
