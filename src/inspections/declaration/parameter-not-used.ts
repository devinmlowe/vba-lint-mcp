// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/ParameterNotUsedInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { DeclarationInspection, type InspectionContext, type InspectionMetadata } from '../base.js';
import type { InspectionResult } from '../types.js';

/**
 * Detects parameters that are declared but never referenced in the procedure body.
 *
 * VBA:
 *   Sub Foo(x As Long)  ' ← x is never used
 *     MsgBox "Hello"
 *   End Sub
 *
 * Rubberduck rationale: Unused parameters add unnecessary complexity to
 * calling code and may indicate incomplete implementation.
 */
export class ParameterNotUsedInspection extends DeclarationInspection {
  static readonly meta: InspectionMetadata = {
    id: 'ParameterNotUsed',
    tier: 'B',
    category: 'CodeQuality',
    defaultSeverity: 'suggestion',
    name: 'Parameter is not used',
    description: 'Parameter is declared but never referenced in the procedure body.',
    quickFixDescription: 'Remove the unused parameter',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const finder = context.declarationFinder;
    if (!finder) return [];

    const results: InspectionResult[] = [];

    for (const decl of finder.findUnusedParameters()) {
      results.push(
        this.createResult(decl.location, {
          description: `Parameter '${decl.name}' is not used.`,
        }),
      );
    }

    return results;
  }
}
