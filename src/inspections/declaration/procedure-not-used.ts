// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/ProcedureNotUsedInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { DeclarationInspection, type InspectionContext, type InspectionMetadata } from '../base.js';
import type { InspectionResult } from '../types.js';
import { isProcedureDeclarationType } from '../../symbols/declaration.js';

/**
 * Detects private procedures (Sub/Function) that are never called.
 *
 * VBA:
 *   Private Sub Helper()  ' ← never called from anywhere
 *     MsgBox "Hello"
 *   End Sub
 *
 * Excludes event handlers (names containing '_') and public procedures
 * (which may be called from other modules).
 */
export class ProcedureNotUsedInspection extends DeclarationInspection {
  static readonly meta: InspectionMetadata = {
    id: 'ProcedureNotUsed',
    tier: 'B',
    category: 'CodeQuality',
    defaultSeverity: 'suggestion',
    name: 'Procedure is not used',
    description: 'Private procedure is never called.',
    quickFixDescription: 'Remove the unused procedure',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const finder = context.declarationFinder;
    if (!finder) return [];

    const results: InspectionResult[] = [];

    for (const decl of finder.getAll()) {
      if (!isProcedureDeclarationType(decl.declarationType)) continue;

      // Only flag private procedures — public ones may be called externally
      if (decl.accessibility !== 'Private') continue;

      // Exclude event handlers (e.g., Worksheet_Change, CommandButton1_Click)
      if (decl.name.includes('_')) continue;

      // Check for zero call references (exclude self-assignment references for Functions)
      const callRefs = decl.references.filter(r => !r.isAssignment);
      if (callRefs.length === 0) {
        const kind = decl.declarationType === 'Function' ? 'Function' : 'Sub';
        results.push(
          this.createResult(decl.location, {
            description: `${kind} '${decl.name}' is never called.`,
          }),
        );
      }
    }

    return results;
  }
}
