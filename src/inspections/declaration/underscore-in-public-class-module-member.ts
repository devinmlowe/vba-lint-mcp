// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/UnderscoreInPublicClassModuleMemberInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { DeclarationInspection, type InspectionContext, type InspectionMetadata } from '../base.js';
import type { InspectionResult } from '../types.js';
import { isProcedureDeclarationType } from '../../symbols/declaration.js';

/**
 * Detects public members in class modules whose names contain underscores.
 *
 * VBA uses underscores for interface implementation (e.g., IFoo_Bar) and
 * event handlers (e.g., Worksheet_Change). A public member with an underscore
 * in a class module may conflict with VBA's interface dispatch mechanism.
 *
 * VBA:
 *   Public Sub My_Method()  ' ← underscore in public member name
 *   End Sub
 */
export class UnderscoreInPublicClassModuleMemberInspection extends DeclarationInspection {
  static readonly meta: InspectionMetadata = {
    id: 'UnderscoreInPublicClassModuleMember',
    tier: 'B',
    category: 'Naming',
    defaultSeverity: 'warning',
    name: 'Underscore in public class module member',
    description: 'Public member name contains underscore, which may conflict with VBA interface dispatch.',
    quickFixDescription: 'Rename to remove the underscore',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const finder = context.declarationFinder;
    if (!finder) return [];

    const results: InspectionResult[] = [];

    for (const decl of finder.getAll()) {
      // Only check procedures and variables (public members)
      if (!isProcedureDeclarationType(decl.declarationType)
        && decl.declarationType !== 'Variable'
        && decl.declarationType !== 'Constant') continue;

      // Only flag public members
      if (decl.accessibility !== 'Public' && decl.accessibility !== 'Implicit') continue;

      // Check for underscore in name
      if (decl.name.includes('_')) {
        results.push(
          this.createResult(decl.location, {
            description: `Public member '${decl.name}' contains an underscore.`,
          }),
        );
      }
    }

    return results;
  }
}
