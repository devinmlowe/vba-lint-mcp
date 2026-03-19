// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/ObsoleteTypeHintInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { ParseTreeInspection, type InspectionContext, type InspectionMetadata } from '../../base.js';
import type { InspectionResult } from '../../types.js';
import { VBAParserVisitor } from '../../../parser/generated/grammar/VBAParserVisitor.js';
import type { TypeHintContext } from '../../../parser/generated/grammar/VBAParser.js';

/**
 * Detects use of obsolete type hint characters ($, %, &, !, #, @).
 *
 * VBA:
 *   Dim name$             ' ← obsolete; use `As String` instead
 *   Dim count%            ' ← obsolete; use `As Integer` instead
 */
export class ObsoleteTypeHintInspection extends ParseTreeInspection {
  static readonly meta: InspectionMetadata = {
    id: 'ObsoleteTypeHint',
    tier: 'A',
    category: 'ObsoleteSyntax',
    defaultSeverity: 'suggestion',
    name: 'Obsolete type hint',
    description: 'Type hint characters are obsolete. Use explicit As Type declarations instead.',
    quickFixDescription: 'Replace type hint with explicit As Type clause',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const results: InspectionResult[] = [];
    const visitor = new ObsoleteTypeHintVisitor(this, results);
    visitor.visit(context.parseResult.tree);
    return results;
  }
}

class ObsoleteTypeHintVisitor extends VBAParserVisitor<void> {
  constructor(
    private readonly inspection: ObsoleteTypeHintInspection,
    private readonly results: InspectionResult[],
  ) {
    super();
  }

  override visitTypeHint = (ctx: TypeHintContext): void => {
    const start = ctx.start!;
    const stop = ctx.stop!;
    this.results.push(
      (this.inspection as any).createResult({
        startLine: start.line,
        startColumn: start.column,
        endLine: stop.line,
        endColumn: stop.column + (stop.text?.length ?? 0),
      }),
    );
    this.visitChildren(ctx);
  };
}
