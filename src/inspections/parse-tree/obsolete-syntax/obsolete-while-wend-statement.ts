// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/ObsoleteWhileWendStatementInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { ParseTreeInspection, type InspectionContext, type InspectionMetadata } from '../../base.js';
import type { InspectionResult } from '../../types.js';
import { VBAParserVisitor } from '../../../parser/generated/grammar/VBAParserVisitor.js';
import type { WhileWendStmtContext } from '../../../parser/generated/grammar/VBAParser.js';

/**
 * Detects use of the obsolete While...Wend loop construct.
 *
 * VBA:
 *   While condition        ' ← obsolete; use Do While...Loop instead
 *   Wend
 */
export class ObsoleteWhileWendStatementInspection extends ParseTreeInspection {
  static readonly meta: InspectionMetadata = {
    id: 'ObsoleteWhileWendStatement',
    tier: 'A',
    category: 'ObsoleteSyntax',
    defaultSeverity: 'suggestion',
    name: 'Obsolete While...Wend statement',
    description: 'While...Wend is obsolete. Use Do While...Loop instead for better control flow.',
    quickFixDescription: 'Replace While...Wend with Do While...Loop',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const results: InspectionResult[] = [];
    const visitor = new ObsoleteWhileWendVisitor(this, results);
    visitor.visit(context.parseResult.tree);
    return results;
  }
}

class ObsoleteWhileWendVisitor extends VBAParserVisitor<void> {
  constructor(
    private readonly inspection: ObsoleteWhileWendStatementInspection,
    private readonly results: InspectionResult[],
  ) {
    super();
  }

  override visitWhileWendStmt = (ctx: WhileWendStmtContext): void => {
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
