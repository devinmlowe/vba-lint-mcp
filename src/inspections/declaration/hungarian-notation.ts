// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/HungarianNotationInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { DeclarationInspection, type InspectionContext, type InspectionMetadata } from '../base.js';
import type { InspectionResult } from '../types.js';

/** Common Hungarian notation prefixes used in VBA. */
const HUNGARIAN_PREFIXES = [
  'str', 'int', 'lng', 'bln', 'bol', 'obj', 'col', 'dbl',
  'sng', 'cur', 'dat', 'var', 'arr', 'rng', 'ws', 'wb',
  'byt', 'dec',
];

/**
 * Detects variables using Hungarian notation (type-prefix naming).
 *
 * VBA:
 *   Dim strName As String    ' ← Hungarian prefix "str"
 *   Dim intCount As Integer  ' ← Hungarian prefix "int"
 *
 * Rubberduck rationale: Hungarian notation is a legacy practice that
 * reduces readability. Modern VBA code should use meaningful names
 * and rely on the type declaration for type information.
 */
export class HungarianNotationInspection extends DeclarationInspection {
  static readonly meta: InspectionMetadata = {
    id: 'HungarianNotation',
    tier: 'B',
    category: 'Naming',
    defaultSeverity: 'suggestion',
    name: 'Hungarian notation detected',
    description: 'Variable name uses Hungarian notation prefix.',
    quickFixDescription: 'Rename to remove the type prefix',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const finder = context.declarationFinder;
    if (!finder) return [];

    const results: InspectionResult[] = [];

    for (const decl of finder.getAll()) {
      // Only check variables, constants, and parameters
      if (decl.declarationType !== 'Variable'
        && decl.declarationType !== 'Constant'
        && decl.declarationType !== 'Parameter') continue;

      const nameLower = decl.name.toLowerCase();

      for (const prefix of HUNGARIAN_PREFIXES) {
        // Match prefix followed by an uppercase letter (e.g., strName, intCount)
        if (nameLower.startsWith(prefix) && decl.name.length > prefix.length) {
          const charAfterPrefix = decl.name[prefix.length];
          if (charAfterPrefix === charAfterPrefix.toUpperCase() && charAfterPrefix !== charAfterPrefix.toLowerCase()) {
            results.push(
              this.createResult(decl.location, {
                description: `'${decl.name}' uses Hungarian notation prefix '${prefix}'.`,
              }),
            );
            break; // Don't flag same declaration twice
          }
        }
      }
    }

    return results;
  }
}
