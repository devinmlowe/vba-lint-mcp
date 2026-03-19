// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/IntegerDataTypeInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { DeclarationInspection, type InspectionContext, type InspectionMetadata } from '../base.js';
import type { InspectionResult } from '../types.js';

/**
 * Detects variables, parameters, and constants declared As Integer.
 *
 * VBA:
 *   Dim count As Integer   ' ← should use Long
 *
 * Rubberduck rationale: On 32-bit and 64-bit systems, VBA internally
 * converts Integer to Long for arithmetic. Using Long directly avoids
 * this overhead and prevents overflow for values beyond 32,767.
 */
export class IntegerDataTypeInspection extends DeclarationInspection {
  static readonly meta: InspectionMetadata = {
    id: 'IntegerDataType',
    tier: 'B',
    category: 'LanguageOpportunities',
    defaultSeverity: 'suggestion',
    name: 'Use Long instead of Integer',
    description: 'Integer type is used — Long is preferred in modern VBA.',
    quickFixDescription: 'Change type from Integer to Long',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const finder = context.declarationFinder;
    if (!finder) return [];

    const results: InspectionResult[] = [];

    for (const decl of finder.getAll()) {
      // Check variables, parameters, constants, and function return types
      if (decl.declarationType !== 'Variable'
        && decl.declarationType !== 'Constant'
        && decl.declarationType !== 'Parameter'
        && decl.declarationType !== 'Function') continue;

      if (decl.asTypeName?.toLowerCase() === 'integer') {
        results.push(
          this.createResult(decl.location, {
            description: `'${decl.name}' is declared As Integer — consider using Long instead.`,
          }),
        );
      }
    }

    return results;
  }
}
