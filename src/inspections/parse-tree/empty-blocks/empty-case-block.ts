// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/EmptyCaseBlockInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { ParseTreeInspection, type InspectionContext, type InspectionMetadata } from '../../base.js';
import type { InspectionResult } from '../../types.js';
import { VBAParserVisitor } from '../../../parser/generated/grammar/VBAParserVisitor.js';
import type { CaseClauseContext, CaseElseClauseContext } from '../../../parser/generated/grammar/VBAParser.js';

/**
 * Detects Case blocks with empty bodies.
 *
 * VBA:
 *   Select Case x
 *       Case 1           ' ← empty body
 *       Case 2
 *           MsgBox "two"
 *   End Select
 */
export class EmptyCaseBlockInspection extends ParseTreeInspection {
  static readonly meta: InspectionMetadata = {
    id: 'EmptyCaseBlock',
    tier: 'A',
    category: 'CodeQuality',
    defaultSeverity: 'warning',
    name: 'Empty Case block',
    description: 'Case block is empty and should either contain code or be removed.',
    quickFixDescription: 'Remove the empty Case block',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const results: InspectionResult[] = [];
    const visitor = new EmptyCaseBlockVisitor(this, results);
    visitor.visit(context.parseResult.tree);
    return results;
  }
}

class EmptyCaseBlockVisitor extends VBAParserVisitor<void> {
  constructor(
    private readonly inspection: EmptyCaseBlockInspection,
    private readonly results: InspectionResult[],
  ) {
    super();
  }

  override visitCaseClause = (ctx: CaseClauseContext): void => {
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

  override visitCaseElseClause = (ctx: CaseElseClauseContext): void => {
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
