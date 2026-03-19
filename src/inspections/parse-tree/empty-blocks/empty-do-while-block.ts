// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/EmptyDoWhileBlockInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { ParseTreeInspection, type InspectionContext, type InspectionMetadata } from '../../base.js';
import type { InspectionResult } from '../../types.js';
import { VBAParserVisitor } from '../../../parser/generated/grammar/VBAParserVisitor.js';
import type { DoLoopStmtContext } from '../../../parser/generated/grammar/VBAParser.js';

/**
 * Detects Do...Loop blocks with empty bodies.
 *
 * VBA:
 *   Do While condition
 *   Loop                  ' ← empty body
 */
export class EmptyDoWhileBlockInspection extends ParseTreeInspection {
  static readonly meta: InspectionMetadata = {
    id: 'EmptyDoWhileBlock',
    tier: 'A',
    category: 'CodeQuality',
    defaultSeverity: 'warning',
    name: 'Empty Do...Loop block',
    description: 'Do...Loop block is empty and should either contain code or be removed.',
    quickFixDescription: 'Remove the empty Do...Loop block',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const results: InspectionResult[] = [];
    const visitor = new EmptyDoWhileVisitor(this, results);
    visitor.visit(context.parseResult.tree);
    return results;
  }
}

class EmptyDoWhileVisitor extends VBAParserVisitor<void> {
  constructor(
    private readonly inspection: EmptyDoWhileBlockInspection,
    private readonly results: InspectionResult[],
  ) {
    super();
  }

  override visitDoLoopStmt = (ctx: DoLoopStmtContext): void => {
    const block = ctx.block();
    if (block) {
      const blockStmts = block.blockStmt();
      if (blockStmts.length === 0) {
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
      }
    }
    this.visitChildren(ctx);
  };
}
