// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/ProcedureCanBeWrittenAsFunctionInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { DeclarationInspection, type InspectionContext, type InspectionMetadata } from '../base.js';
import type { InspectionResult } from '../types.js';

/**
 * Detects Sub procedures that assign to exactly one ByRef parameter,
 * indicating they could be written as a Function instead.
 *
 * VBA:
 *   Sub GetValue(ByRef result As Long)
 *     result = 42
 *   End Sub
 *
 * Could be rewritten as:
 *   Function GetValue() As Long
 *     GetValue = 42
 *   End Function
 *
 * Rubberduck rationale: Functions are more idiomatic than ByRef output
 * parameters when a procedure produces a single return value.
 */
export class ProcedureCanBeWrittenAsFunctionInspection extends DeclarationInspection {
  static readonly meta: InspectionMetadata = {
    id: 'ProcedureCanBeWrittenAsFunction',
    tier: 'B',
    category: 'LanguageOpportunities',
    defaultSeverity: 'suggestion',
    name: 'Procedure can be written as a Function',
    description: 'Sub assigns to a ByRef parameter — could be a Function instead.',
    quickFixDescription: 'Convert to Function with return value',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const finder = context.declarationFinder;
    if (!finder) return [];

    const results: InspectionResult[] = [];

    for (const decl of finder.findByType('Sub')) {
      // Find parameters that belong to this Sub
      const params = finder.getAll().filter(d =>
        d.declarationType === 'Parameter'
        && d.parentScope === decl,
      );

      // Find ByRef parameters that are assigned to
      const assignedByRefParams = params.filter(p =>
        p.isByRef && p.references.some(r => r.isAssignment),
      );

      // If exactly one ByRef parameter is assigned, suggest converting to Function
      if (assignedByRefParams.length === 1) {
        results.push(
          this.createResult(decl.location, {
            description: `Sub '${decl.name}' assigns to ByRef parameter '${assignedByRefParams[0].name}' — could be a Function.`,
          }),
        );
      }
    }

    return results;
  }
}
