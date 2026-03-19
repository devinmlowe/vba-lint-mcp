// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/NonReturningFunctionInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { DeclarationInspection, type InspectionContext, type InspectionMetadata } from '../base.js';
import type { InspectionResult } from '../types.js';

/**
 * Detects Functions and Property Get procedures where the function name
 * is never assigned, meaning no return value is ever set.
 *
 * VBA:
 *   Function GetValue() As Long
 *     MsgBox "Hello"            ' ← GetValue is never assigned
 *   End Function
 *
 * Rubberduck rationale: A Function that never assigns its return value
 * always returns the default (0, "", Nothing). This is likely a bug —
 * it should either be a Sub or should assign a return value.
 */
export class NonReturningFunctionInspection extends DeclarationInspection {
  static readonly meta: InspectionMetadata = {
    id: 'NonReturningFunction',
    tier: 'B',
    category: 'CodeQuality',
    defaultSeverity: 'warning',
    name: 'Function does not return a value',
    description: 'Function or Property Get never assigns a return value.',
    quickFixDescription: 'Convert to Sub or add a return value assignment',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const finder = context.declarationFinder;
    if (!finder) return [];

    const results: InspectionResult[] = [];

    for (const decl of finder.findNonReturningFunctions()) {
      const kind = decl.declarationType === 'PropertyGet' ? 'Property Get' : 'Function';
      results.push(
        this.createResult(decl.location, {
          description: `${kind} '${decl.name}' does not return a value.`,
        }),
      );
    }

    return results;
  }
}
