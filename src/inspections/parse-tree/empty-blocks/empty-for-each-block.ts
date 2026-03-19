// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/EmptyForEachBlockInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { ParseTreeInspection, type InspectionContext, type InspectionMetadata } from '../../base.js';
import type { InspectionResult } from '../../types.js';
import { VBAParserVisitor } from '../../../parser/generated/grammar/VBAParserVisitor.js';
import type { ForEachStmtContext } from '../../../parser/generated/grammar/VBAParser.js';

/**
 * Detects For Each...Next loops with empty bodies.
 *
 * VBA:
 *   For Each item In collection
 *   Next item             ' ← empty body
 */
export class EmptyForEachBlockInspection extends ParseTreeInspection {
  static readonly meta: InspectionMetadata = {
    id: 'EmptyForEachBlock',
    tier: 'A',
    category: 'CodeQuality',
    defaultSeverity: 'warning',
    name: 'Empty For Each...Next loop',
    description: 'For Each...Next loop is empty and should either contain code or be removed.',
    quickFixDescription: 'Remove the empty For Each...Next loop',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const results: InspectionResult[] = [];
    const visitor = new EmptyForEachVisitor(this, results);
    visitor.visit(context.parseResult.tree);
    return results;
  }
}

class EmptyForEachVisitor extends VBAParserVisitor<void> {
  constructor(
    private readonly inspection: EmptyForEachBlockInspection,
    private readonly results: InspectionResult[],
  ) {
    super();
  }

  override visitForEachStmt = (ctx: ForEachStmtContext): void => {
    // ForEachStmt uses unterminatedBlock instead of block
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
