// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/MoveFieldCloserToUsageInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { DeclarationInspection, type InspectionContext, type InspectionMetadata } from '../base.js';
import type { InspectionResult } from '../types.js';
import { isProcedureDeclarationType } from '../../symbols/declaration.js';

/**
 * Detects module-level variables that are only used within a single procedure.
 *
 * VBA:
 *   Private tempValue As Long  ' ← only used in DoWork
 *
 *   Sub DoWork()
 *     tempValue = 42
 *     MsgBox tempValue
 *   End Sub
 *
 * Rubberduck rationale: Variables should be declared in the narrowest
 * scope possible. If a module-level variable is only used in one
 * procedure, it should be a local variable in that procedure.
 */
export class MoveFieldCloserToUsageInspection extends DeclarationInspection {
  static readonly meta: InspectionMetadata = {
    id: 'MoveFieldCloserToUsage',
    tier: 'B',
    category: 'CodeQuality',
    defaultSeverity: 'suggestion',
    name: 'Move field closer to usage',
    description: 'Module-level variable is only used in one procedure — could be local.',
    quickFixDescription: 'Move to local scope in the using procedure',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const finder = context.declarationFinder;
    if (!finder) return [];

    const results: InspectionResult[] = [];

    // Get all procedures for scope checking
    const procedures = finder.getAll().filter(d => isProcedureDeclarationType(d.declarationType));

    for (const decl of finder.getAll()) {
      if (decl.declarationType !== 'Variable') continue;

      // Only module-level variables
      if (!decl.parentScope || decl.parentScope.declarationType !== 'Module') continue;

      // Must have at least one reference
      if (decl.references.length === 0) continue;

      // Check if all references are within the same procedure
      const usingProcedures = new Set<string>();
      for (const ref of decl.references) {
        // Find which procedure contains this reference
        const containingProc = procedures.find(proc =>
          ref.location.startLine >= proc.location.startLine
          && ref.location.endLine <= proc.location.endLine,
        );
        if (containingProc) {
          usingProcedures.add(containingProc.name);
        } else {
          // Reference outside any procedure — can't move to local
          usingProcedures.add('__module__');
        }
      }

      if (usingProcedures.size === 1 && !usingProcedures.has('__module__')) {
        const [procName] = usingProcedures;
        results.push(
          this.createResult(decl.location, {
            description: `Module-level variable '${decl.name}' is only used in '${procName}' — consider making it local.`,
          }),
        );
      }
    }

    return results;
  }
}
