// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/EncapsulatePublicFieldInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { DeclarationInspection, type InspectionContext, type InspectionMetadata } from '../base.js';
import type { InspectionResult } from '../types.js';

/**
 * Detects public variables in class modules that should be encapsulated
 * with Property Get/Let/Set procedures.
 *
 * VBA:
 *   Public Name As String   ' ← should use Property Get/Let
 *
 * Rubberduck rationale: Public fields in classes break encapsulation.
 * Using Property procedures allows validation, lazy initialization,
 * and easier refactoring.
 */
export class EncapsulatePublicFieldInspection extends DeclarationInspection {
  static readonly meta: InspectionMetadata = {
    id: 'EncapsulatePublicField',
    tier: 'B',
    category: 'CodeQuality',
    defaultSeverity: 'suggestion',
    name: 'Encapsulate public field',
    description: 'Public variable in a class module should use Property procedures.',
    quickFixDescription: 'Encapsulate with Property Get/Let or Get/Set',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const finder = context.declarationFinder;
    if (!finder) return [];

    const results: InspectionResult[] = [];

    for (const decl of finder.getAll()) {
      if (decl.declarationType !== 'Variable') continue;

      // Only flag public variables at module level
      if (decl.accessibility !== 'Public') continue;
      if (!decl.parentScope || decl.parentScope.declarationType !== 'Module') continue;

      results.push(
        this.createResult(decl.location, {
          description: `Public variable '${decl.name}' should be encapsulated with Property procedures.`,
        }),
      );
    }

    return results;
  }
}
