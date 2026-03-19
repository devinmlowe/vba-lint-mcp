// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/IsMissingWithNonArgumentParameterInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { ParseTreeInspection, type InspectionContext, type InspectionMetadata } from '../../base.js';
import type { InspectionResult } from '../../types.js';

/**
 * Detects use of IsMissing() with non-parameter arguments.
 *
 * VBA:
 *   Dim x As Variant
 *   If IsMissing(x) Then  ' ← x is a local variable, not a parameter
 *
 * Note: Full implementation requires symbol resolution (Tier B) to determine
 * if the argument is a parameter or a local variable. This Tier A stub
 * registers the inspection metadata but cannot perform the check.
 *
 * TODO: Implement fully when symbol resolution is available (Tier B).
 */
export class IsMissingWithNonArgumentParameterInspection extends ParseTreeInspection {
  static readonly meta: InspectionMetadata = {
    id: 'IsMissingWithNonArgumentParameter',
    tier: 'A',
    category: 'CodeQuality',
    defaultSeverity: 'warning',
    name: 'IsMissing with non-parameter',
    description: 'IsMissing is called with a non-parameter argument. IsMissing only works with procedure parameters.',
    quickFixDescription: 'Pass a procedure parameter to IsMissing instead',
  };

  inspect(_context: InspectionContext): InspectionResult[] {
    // TODO: Requires symbol resolution to determine if the argument to IsMissing
    // is a procedure parameter or a local variable. Cannot be implemented at Tier A.
    return [];
  }
}
