// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/BooleanAssignedInIfElseInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { ParseTreeInspection, type InspectionContext, type InspectionMetadata } from '../../base.js';
import type { InspectionResult } from '../../types.js';
import { VBAParserVisitor } from '../../../parser/generated/grammar/VBAParserVisitor.js';
import type { IfStmtContext } from '../../../parser/generated/grammar/VBAParser.js';

/**
 * Detects If/Else blocks that simply assign True/False to the same variable.
 *
 * VBA:
 *   If condition Then
 *       x = True
 *   Else
 *       x = False
 *   End If
 *   ' Should be: x = condition
 */
export class BooleanAssignedInIfElseInspection extends ParseTreeInspection {
  static readonly meta: InspectionMetadata = {
    id: 'BooleanAssignedInIfElse',
    tier: 'A',
    category: 'CodeQuality',
    defaultSeverity: 'suggestion',
    name: 'Boolean assigned in If/Else',
    description: 'If/Else block assigns True/False to the same variable. Simplify to a direct assignment.',
    quickFixDescription: 'Replace with direct boolean assignment',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const results: InspectionResult[] = [];
    const visitor = new BooleanAssignedVisitor(this, results);
    visitor.visit(context.parseResult.tree);
    return results;
  }
}

class BooleanAssignedVisitor extends VBAParserVisitor<void> {
  constructor(
    private readonly inspection: BooleanAssignedInIfElseInspection,
    private readonly results: InspectionResult[],
  ) {
    super();
  }

  override visitIfStmt = (ctx: IfStmtContext): void => {
    // Check: If block has exactly 1 statement, Else block has exactly 1 statement,
    // no ElseIf blocks, and both statements are assignments of True/False
    const block = ctx.block();
    const elseBlock = ctx.elseBlock();
    const elseIfBlocks = ctx.elseIfBlock();

    if (block && elseBlock && elseIfBlocks.length === 0) {
      const ifStmts = block.blockStmt();
      const elseStmts = elseBlock.block()?.blockStmt() ?? [];

      if (ifStmts.length === 1 && elseStmts.length === 1) {
        const ifText = ifStmts[0].getText().replace(/\s+/g, ' ').trim().toLowerCase();
        const elseText = elseStmts[0].getText().replace(/\s+/g, ' ').trim().toLowerCase();

        // Check if both are assignments (contain '=')
        // and one assigns True while the other assigns False
        const ifMatch = ifText.match(/^(\w+)\s*=\s*(true|false)$/);
        const elseMatch = elseText.match(/^(\w+)\s*=\s*(true|false)$/);

        if (ifMatch && elseMatch) {
          const sameVar = ifMatch[1] === elseMatch[1];
          const oppositeValues = ifMatch[2] !== elseMatch[2];
          if (sameVar && oppositeValues) {
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
      }
    }

    this.visitChildren(ctx);
  };
}
