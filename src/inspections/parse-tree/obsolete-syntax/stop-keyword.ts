// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/StopKeywordInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { ParseTreeInspection, type InspectionContext, type InspectionMetadata } from '../../base.js';
import type { InspectionResult } from '../../types.js';
import { VBAParserVisitor } from '../../../parser/generated/grammar/VBAParserVisitor.js';
import type { StopStmtContext } from '../../../parser/generated/grammar/VBAParser.js';

/**
 * Detects use of the Stop statement, which is typically a debug leftover.
 *
 * VBA:
 *   Stop                  ' ← debug statement; should be removed
 */
export class StopKeywordInspection extends ParseTreeInspection {
  static readonly meta: InspectionMetadata = {
    id: 'StopKeyword',
    tier: 'A',
    category: 'CodeQuality',
    defaultSeverity: 'warning',
    name: 'Stop keyword',
    description: 'Stop statement halts execution like a breakpoint and should be removed from production code.',
    quickFixDescription: 'Remove the Stop statement',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const results: InspectionResult[] = [];
    const visitor = new StopKeywordVisitor(this, results);
    visitor.visit(context.parseResult.tree);
    return results;
  }
}

class StopKeywordVisitor extends VBAParserVisitor<void> {
  constructor(
    private readonly inspection: StopKeywordInspection,
    private readonly results: InspectionResult[],
  ) {
    super();
  }

  override visitStopStmt = (ctx: StopStmtContext): void => {
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
