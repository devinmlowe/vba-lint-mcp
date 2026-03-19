// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/ObjectVariableNotSetInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { DeclarationInspection, type InspectionContext, type InspectionMetadata } from '../base.js';
import type { InspectionResult } from '../types.js';

/** Known VBA object types that require Set for assignment. */
const OBJECT_TYPE_NAMES = new Set([
  'object', 'collection', 'dictionary',
  'workbook', 'worksheet', 'range', 'chart', 'shape',
  'application', 'window', 'pane',
  'recordset', 'connection', 'command',
  'scripting.dictionary', 'scripting.filesystemobject',
  'adodb.recordset', 'adodb.connection',
  'vbcomponent', 'vbproject', 'codemodule',
  'userform', 'control', 'commandbutton', 'textbox', 'label',
  'listbox', 'combobox', 'checkbox', 'optionbutton', 'frame',
]);

/**
 * Detects assignments to object variables that are missing the Set keyword.
 *
 * VBA:
 *   Dim ws As Worksheet
 *   ws = ActiveSheet       ' ← should be: Set ws = ActiveSheet
 *
 * This is a simplified version — it flags any assignment to a variable
 * whose declared type is a known object type.
 */
export class ObjectVariableNotSetInspection extends DeclarationInspection {
  static readonly meta: InspectionMetadata = {
    id: 'ObjectVariableNotSet',
    tier: 'B',
    category: 'CodeQuality',
    defaultSeverity: 'error',
    name: 'Object variable not Set',
    description: 'Assignment to object variable without Set keyword.',
    quickFixDescription: 'Add the Set keyword before the assignment',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const finder = context.declarationFinder;
    if (!finder) return [];

    const results: InspectionResult[] = [];

    for (const decl of finder.getAll()) {
      if (decl.declarationType !== 'Variable') continue;

      // Check if the variable's type is a known object type
      if (!decl.asTypeName) continue;
      if (!OBJECT_TYPE_NAMES.has(decl.asTypeName.toLowerCase())) continue;

      // Check for assignment references (these are potential Set violations)
      for (const ref of decl.references) {
        if (ref.isAssignment) {
          results.push(
            this.createResult(ref.location, {
              description: `'${decl.name}' is an object variable — assignment may require 'Set'.`,
            }),
          );
        }
      }
    }

    return results;
  }
}
