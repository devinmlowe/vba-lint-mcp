// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/EmptyWhileWendBlockInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { ParseTreeInspection, type InspectionContext, type InspectionMetadata } from '../../base.js';
import type { InspectionResult } from '../../types.js';
import { VBAParserVisitor } from '../../../parser/generated/grammar/VBAParserVisitor.js';
import type { WhileWendStmtContext } from '../../../parser/generated/grammar/VBAParser.js';

/**
 * Detects While...Wend loops with empty bodies.
 *
 * VBA:
 *   While condition
 *   Wend                  ' ← empty body
 */
export class EmptyWhileWendBlockInspection extends ParseTreeInspection {
  static readonly meta: InspectionMetadata = {
    id: 'EmptyWhileWendBlock',
    tier: 'A',
    category: 'CodeQuality',
    defaultSeverity: 'warning',
    name: 'Empty While...Wend loop',
    description: 'While...Wend loop is empty and should either contain code or be removed.',
    quickFixDescription: 'Remove the empty While...Wend loop',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const results: InspectionResult[] = [];
    const visitor = new EmptyWhileWendVisitor(this, results);
    visitor.visit(context.parseResult.tree);
    return results;
  }
}

class EmptyWhileWendVisitor extends VBAParserVisitor<void> {
  constructor(
    private readonly inspection: EmptyWhileWendBlockInspection,
    private readonly results: InspectionResult[],
  ) {
    super();
  }

  override visitWhileWendStmt = (ctx: WhileWendStmtContext): void => {
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
