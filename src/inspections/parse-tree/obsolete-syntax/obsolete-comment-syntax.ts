// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/ObsoleteCommentSyntaxInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { ParseTreeInspection, type InspectionContext, type InspectionMetadata } from '../../base.js';
import type { InspectionResult } from '../../types.js';
import { VBAParserVisitor } from '../../../parser/generated/grammar/VBAParserVisitor.js';
import type { RemCommentContext } from '../../../parser/generated/grammar/VBAParser.js';

/**
 * Detects use of the obsolete `Rem` keyword for comments.
 *
 * VBA:
 *   Rem This is a comment  ' ← obsolete; use ' instead
 */
export class ObsoleteCommentSyntaxInspection extends ParseTreeInspection {
  static readonly meta: InspectionMetadata = {
    id: 'ObsoleteCommentSyntax',
    tier: 'A',
    category: 'ObsoleteSyntax',
    defaultSeverity: 'suggestion',
    name: 'Obsolete Rem comment',
    description: "The Rem keyword for comments is obsolete. Use the single-quote (') syntax instead.",
    quickFixDescription: "Replace Rem with '",
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const results: InspectionResult[] = [];
    const visitor = new ObsoleteCommentVisitor(this, results);
    visitor.visit(context.parseResult.tree);
    return results;
  }
}

class ObsoleteCommentVisitor extends VBAParserVisitor<void> {
  constructor(
    private readonly inspection: ObsoleteCommentSyntaxInspection,
    private readonly results: InspectionResult[],
  ) {
    super();
  }

  override visitRemComment = (ctx: RemCommentContext): void => {
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
