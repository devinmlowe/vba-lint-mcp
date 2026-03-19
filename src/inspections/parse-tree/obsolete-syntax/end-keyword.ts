// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/EndKeywordInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { ParseTreeInspection, type InspectionContext, type InspectionMetadata } from '../../base.js';
import type { InspectionResult } from '../../types.js';
import { VBAParserVisitor } from '../../../parser/generated/grammar/VBAParserVisitor.js';
import type { EndStmtContext } from '../../../parser/generated/grammar/VBAParser.js';

/**
 * Detects standalone `End` statements (not End Sub, End Function, etc.).
 *
 * VBA:
 *   End                   ' ← abruptly terminates program; should use proper exit
 */
export class EndKeywordInspection extends ParseTreeInspection {
  static readonly meta: InspectionMetadata = {
    id: 'EndKeyword',
    tier: 'A',
    category: 'CodeQuality',
    defaultSeverity: 'warning',
    name: 'End keyword',
    description: 'Standalone End statement abruptly terminates the program without cleanup.',
    quickFixDescription: 'Replace with proper exit logic',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const results: InspectionResult[] = [];
    const visitor = new EndKeywordVisitor(this, results);
    visitor.visit(context.parseResult.tree);
    return results;
  }
}

class EndKeywordVisitor extends VBAParserVisitor<void> {
  constructor(
    private readonly inspection: EndKeywordInspection,
    private readonly results: InspectionResult[],
  ) {
    super();
  }

  override visitEndStmt = (ctx: EndStmtContext): void => {
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
