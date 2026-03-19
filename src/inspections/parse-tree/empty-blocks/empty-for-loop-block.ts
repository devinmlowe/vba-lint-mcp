// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/EmptyForLoopBlockInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { ParseTreeInspection, type InspectionContext, type InspectionMetadata } from '../../base.js';
import type { InspectionResult } from '../../types.js';
import { VBAParserVisitor } from '../../../parser/generated/grammar/VBAParserVisitor.js';
import type { ForNextStmtContext } from '../../../parser/generated/grammar/VBAParser.js';

/**
 * Detects For...Next loops with empty bodies.
 *
 * VBA:
 *   For i = 1 To 10
 *   Next i               ' ← empty body
 */
export class EmptyForLoopBlockInspection extends ParseTreeInspection {
  static readonly meta: InspectionMetadata = {
    id: 'EmptyForLoopBlock',
    tier: 'A',
    category: 'CodeQuality',
    defaultSeverity: 'warning',
    name: 'Empty For...Next loop',
    description: 'For...Next loop is empty and should either contain code or be removed.',
    quickFixDescription: 'Remove the empty For...Next loop',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const results: InspectionResult[] = [];
    const visitor = new EmptyForLoopVisitor(this, results);
    visitor.visit(context.parseResult.tree);
    return results;
  }
}

class EmptyForLoopVisitor extends VBAParserVisitor<void> {
  constructor(
    private readonly inspection: EmptyForLoopBlockInspection,
    private readonly results: InspectionResult[],
  ) {
    super();
  }

  override visitForNextStmt = (ctx: ForNextStmtContext): void => {
    // ForNextStmt uses unterminatedBlock instead of block
    const unterminatedBlock = ctx.unterminatedBlock();
    if (!unterminatedBlock || unterminatedBlock.blockStmt().length === 0) {
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
    this.visitChildren(ctx);
  };
}
