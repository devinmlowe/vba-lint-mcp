// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/FunctionReturnValueAlwaysDiscardedInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { DeclarationInspection, type InspectionContext, type InspectionMetadata } from '../base.js';
import type { InspectionResult } from '../types.js';

/**
 * Detects Functions whose return value is never used at ANY call site.
 *
 * VBA:
 *   Function DoWork() As Boolean
 *     ' ...
 *     DoWork = True
 *   End Function
 *
 *   Sub Caller1()
 *     DoWork               ' discarded
 *   End Sub
 *   Sub Caller2()
 *     DoWork               ' also discarded
 *   End Sub
 *
 * If a Function sets a return value but no caller ever captures it,
 * the Function might as well be a Sub.
 *
 * Note: This simplified version requires parse-tree context to distinguish
 * statement calls from expression calls. Currently a stub — needs
 * reference augmentation to track call-site context.
 */
export class FunctionReturnValueAlwaysDiscardedInspection extends DeclarationInspection {
  static readonly meta: InspectionMetadata = {
    id: 'FunctionReturnValueAlwaysDiscarded',
    tier: 'B',
    category: 'CodeQuality',
    defaultSeverity: 'suggestion',
    name: 'Function return value is always discarded',
    description: 'Function return value is never used at any call site — consider converting to Sub.',
    quickFixDescription: 'Convert Function to Sub',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const finder = context.declarationFinder;
    if (!finder) return [];

    // This inspection needs call-site context (statement vs expression)
    // which is not yet tracked in the reference model.
    // Returns empty for now — will be implemented when reference model
    // is augmented with call-site context tracking.
    return [];
  }
}
