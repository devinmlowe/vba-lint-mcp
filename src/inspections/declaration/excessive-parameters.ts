// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/ExcessiveParametersInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { DeclarationInspection, type InspectionContext, type InspectionMetadata } from '../base.js';
import type { InspectionResult } from '../types.js';
import { isProcedureDeclarationType } from '../../symbols/declaration.js';

/** Maximum recommended parameter count. */
const MAX_PARAMETERS = 7;

/**
 * Detects procedures with too many parameters.
 *
 * VBA:
 *   Sub DoWork(a As Long, b As Long, c As Long, d As Long, _
 *              e As Long, f As Long, g As Long, h As Long)
 *   End Sub
 *
 * Rubberduck rationale: Procedures with many parameters are hard to
 * call correctly and may indicate that parameters should be grouped
 * into a Type or class.
 */
export class ExcessiveParametersInspection extends DeclarationInspection {
  static readonly meta: InspectionMetadata = {
    id: 'ExcessiveParameters',
    tier: 'B',
    category: 'CodeQuality',
    defaultSeverity: 'suggestion',
    name: 'Excessive parameters',
    description: 'Procedure has too many parameters.',
    quickFixDescription: 'Refactor to use a Type or reduce parameter count',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const finder = context.declarationFinder;
    if (!finder) return [];

    const results: InspectionResult[] = [];

    for (const decl of finder.getAll()) {
      if (!isProcedureDeclarationType(decl.declarationType)) continue;

      // Count parameters belonging to this procedure
      const paramCount = finder.getAll().filter(d =>
        d.declarationType === 'Parameter'
        && d.parentScope === decl,
      ).length;

      if (paramCount > MAX_PARAMETERS) {
        results.push(
          this.createResult(decl.location, {
            description: `'${decl.name}' has ${paramCount} parameters (maximum recommended is ${MAX_PARAMETERS}).`,
          }),
        );
      }
    }

    return results;
  }
}
