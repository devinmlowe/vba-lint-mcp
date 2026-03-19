// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/EmptyIfBlockInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { ParseTreeInspection, type InspectionContext, type InspectionMetadata } from '../../base.js';
import type { InspectionResult } from '../../types.js';
import { VBAParserVisitor } from '../../../parser/generated/grammar/VBAParserVisitor.js';
import type { IfStmtContext } from '../../../parser/generated/grammar/VBAParser.js';

/**
 * Detects If blocks with empty bodies.
 *
 * VBA:
 *   If condition Then
 *   End If              ' ← empty body
 *
 * Rubberduck rationale: An empty If block is likely unfinished code or a
 * leftover from refactoring. It should either contain logic or be removed.
 */
export class EmptyIfBlockInspection extends ParseTreeInspection {
  static readonly meta: InspectionMetadata = {
    id: 'EmptyIfBlock',
    tier: 'A',
    category: 'CodeQuality',
    defaultSeverity: 'warning',
    name: 'Empty If block',
    description: 'If block is empty and should either contain code or be removed.',
    quickFixDescription: 'Remove the empty If block',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const results: InspectionResult[] = [];
    const tree = context.parseResult.tree;

    // Use a visitor to find all If statements
    const visitor = new EmptyIfBlockVisitor(this, results);
    visitor.visit(tree);

    return results;
  }
}

class EmptyIfBlockVisitor extends VBAParserVisitor<void> {
  constructor(
    private readonly inspection: EmptyIfBlockInspection,
    private readonly results: InspectionResult[],
  ) {
    super();
  }

  override visitIfStmt = (ctx: IfStmtContext): void => {
    // Check if the main If block is empty.
    // ifStmt has a single block() child (the Then block).
    // block : (blockStmt endOfStatement)*
    // An empty block has no blockStmt children.
    const block = ctx.block();
    if (block) {
      const blockStmts = block.blockStmt();
      if (blockStmts.length === 0) {
        const start = ctx.start!;
        const stop = ctx.stop!;
        this.results.push(
          (this.inspection as any).createResult(
            {
              startLine: start.line,
              startColumn: start.column,
              endLine: stop.line,
              endColumn: stop.column + (stop.text?.length ?? 0),
            },
          ),
        );
      }
    }

    // Continue visiting children (nested If statements)
    this.visitChildren(ctx);
  };
}
