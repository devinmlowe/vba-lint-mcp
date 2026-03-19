// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/EmptyElseBlockInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { ParseTreeInspection, type InspectionContext, type InspectionMetadata } from '../../base.js';
import type { InspectionResult } from '../../types.js';
import { VBAParserVisitor } from '../../../parser/generated/grammar/VBAParserVisitor.js';
import type { ElseBlockContext } from '../../../parser/generated/grammar/VBAParser.js';

/**
 * Detects Else blocks with empty bodies.
 *
 * VBA:
 *   If condition Then
 *       MsgBox "Hi"
 *   Else
 *   End If              ' ← empty Else body
 */
export class EmptyElseBlockInspection extends ParseTreeInspection {
  static readonly meta: InspectionMetadata = {
    id: 'EmptyElseBlock',
    tier: 'A',
    category: 'CodeQuality',
    defaultSeverity: 'warning',
    name: 'Empty Else block',
    description: 'Else block is empty and should either contain code or be removed.',
    quickFixDescription: 'Remove the empty Else block',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const results: InspectionResult[] = [];
    const visitor = new EmptyElseBlockVisitor(this, results);
    visitor.visit(context.parseResult.tree);
    return results;
  }
}

class EmptyElseBlockVisitor extends VBAParserVisitor<void> {
  constructor(
    private readonly inspection: EmptyElseBlockInspection,
    private readonly results: InspectionResult[],
  ) {
    super();
  }

  override visitElseBlock = (ctx: ElseBlockContext): void => {
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
