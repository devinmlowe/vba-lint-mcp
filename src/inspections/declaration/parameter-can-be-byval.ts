// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/ParameterCanBeByValInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { DeclarationInspection, type InspectionContext, type InspectionMetadata } from '../base.js';
import type { InspectionResult } from '../types.js';

/**
 * Detects ByRef parameters that are never assigned to, meaning they
 * could safely be passed ByVal.
 *
 * VBA:
 *   Sub DoWork(ByRef x As Long)   ' x is never assigned
 *     MsgBox CStr(x)
 *   End Sub
 *
 * Rubberduck rationale: ByRef parameters that are never assigned to
 * should be ByVal to communicate intent and prevent accidental mutation.
 */
export class ParameterCanBeByValInspection extends DeclarationInspection {
  static readonly meta: InspectionMetadata = {
    id: 'ParameterCanBeByVal',
    tier: 'B',
    category: 'CodeQuality',
    defaultSeverity: 'suggestion',
    name: 'Parameter can be ByVal',
    description: 'ByRef parameter is never assigned to — could be ByVal.',
    quickFixDescription: 'Change parameter to ByVal',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const finder = context.declarationFinder;
    if (!finder) return [];

    const results: InspectionResult[] = [];

    for (const decl of finder.findByType('Parameter')) {
      // Only check ByRef parameters
      if (!decl.isByRef) continue;

      // Check if the parameter is ever assigned to
      const isAssigned = decl.references.some(r => r.isAssignment);
      if (!isAssigned) {
        results.push(
          this.createResult(decl.location, {
            description: `Parameter '${decl.name}' is ByRef but is never assigned — consider ByVal.`,
          }),
        );
      }
    }

    return results;
  }
}
