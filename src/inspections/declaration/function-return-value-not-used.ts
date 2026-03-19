// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/FunctionReturnValueNotUsedInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { DeclarationInspection, type InspectionContext, type InspectionMetadata } from '../base.js';
import type { InspectionResult } from '../types.js';

/**
 * Detects Functions whose return value is never captured at any call site.
 *
 * VBA:
 *   Function GetValue() As Long
 *     GetValue = 42
 *   End Function
 *
 *   Sub Test()
 *     GetValue           ' ← return value discarded
 *   End Sub
 *
 * This simplified version checks: if a Function is called (has non-assignment
 * references from outside itself) but never referenced in an expression
 * context — i.e., all references are statement-level calls — then the
 * return value is always discarded.
 */
export class FunctionReturnValueNotUsedInspection extends DeclarationInspection {
  static readonly meta: InspectionMetadata = {
    id: 'FunctionReturnValueNotUsed',
    tier: 'B',
    category: 'CodeQuality',
    defaultSeverity: 'suggestion',
    name: 'Function return value is not used',
    description: 'Function return value is always discarded by callers.',
    quickFixDescription: 'Capture the return value or convert to Sub',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const finder = context.declarationFinder;
    if (!finder) return [];

    const results: InspectionResult[] = [];

    for (const decl of finder.findByType('Function')) {
      // Get call references (exclude self-assignment for return value)
      const callRefs = decl.references.filter(r =>
        !r.isAssignment,
      );

      // If the function is called but never has its return value captured,
      // all call refs would be statement-level calls.
      // Since our reference model doesn't distinguish expression vs statement context,
      // we skip this inspection for now — it needs parse-tree augmentation.
      // See FunctionReturnValueAlwaysDiscardedInspection for the aggregate check.
    }

    return results;
  }
}
