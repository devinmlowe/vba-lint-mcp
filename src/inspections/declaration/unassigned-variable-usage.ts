// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/UnassignedVariableUsageInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { DeclarationInspection, type InspectionContext, type InspectionMetadata } from '../base.js';
import type { InspectionResult } from '../types.js';

/**
 * Detects variables that are used (read) before any assignment.
 *
 * VBA:
 *   Sub Test()
 *     Dim x As Long
 *     MsgBox x        ' ← x is read but was never assigned
 *   End Sub
 *
 * This is similar to VariableNotAssigned but flags at each usage site
 * rather than at the declaration. A variable that has zero assignments
 * but has read references is using the default value, which is likely a bug.
 */
export class UnassignedVariableUsageInspection extends DeclarationInspection {
  static readonly meta: InspectionMetadata = {
    id: 'UnassignedVariableUsage',
    tier: 'B',
    category: 'CodeQuality',
    defaultSeverity: 'warning',
    name: 'Unassigned variable is used',
    description: 'Variable is used before being assigned a value.',
    quickFixDescription: 'Assign the variable before use',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const finder = context.declarationFinder;
    if (!finder) return [];

    const results: InspectionResult[] = [];

    for (const decl of finder.findVariablesNotAssigned()) {
      // Flag each read reference to the unassigned variable
      for (const ref of decl.references) {
        if (!ref.isAssignment) {
          results.push(
            this.createResult(ref.location, {
              description: `Variable '${decl.name}' is used but has never been assigned a value.`,
            }),
          );
        }
      }
    }

    return results;
  }
}
