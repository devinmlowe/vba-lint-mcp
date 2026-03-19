// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/VariableTypeNotDeclaredInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { DeclarationInspection, type InspectionContext, type InspectionMetadata } from '../base.js';
import type { InspectionResult } from '../types.js';

/**
 * Detects variables and parameters declared without an explicit type
 * (implicit Variant).
 *
 * VBA:
 *   Dim value            ' ← implicit Variant, should be: Dim value As Variant
 *   Sub Test(x)          ' ← implicit Variant parameter
 *
 * Rubberduck rationale: Implicit Variant can hide type errors and makes
 * code harder to understand. Explicitly declare the type, even if Variant
 * is intended.
 */
export class VariableTypeNotDeclaredInspection extends DeclarationInspection {
  static readonly meta: InspectionMetadata = {
    id: 'VariableTypeNotDeclared',
    tier: 'B',
    category: 'LanguageOpportunities',
    defaultSeverity: 'suggestion',
    name: 'Variable type is not declared',
    description: 'Variable is declared without an explicit type (implicit Variant).',
    quickFixDescription: 'Add an explicit As clause',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const finder = context.declarationFinder;
    if (!finder) return [];

    const results: InspectionResult[] = [];

    for (const decl of finder.getAll()) {
      if (decl.declarationType !== 'Variable'
        && decl.declarationType !== 'Constant'
        && decl.declarationType !== 'Parameter') continue;

      if (decl.isImplicitType) {
        results.push(
          this.createResult(decl.location, {
            description: `'${decl.name}' is declared without an explicit type (implicit Variant).`,
          }),
        );
      }
    }

    return results;
  }
}
