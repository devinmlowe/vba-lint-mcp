// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/ObsoleteLetStatementInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { ParseTreeInspection, type InspectionContext, type InspectionMetadata } from '../../base.js';
import type { InspectionResult } from '../../types.js';
import { VBAParserVisitor } from '../../../parser/generated/grammar/VBAParserVisitor.js';
import type { LetStmtContext } from '../../../parser/generated/grammar/VBAParser.js';

/**
 * Detects use of the obsolete `Let` keyword in assignments.
 *
 * VBA:
 *   Let x = 5             ' ← obsolete; should be `x = 5`
 */
export class ObsoleteLetStatementInspection extends ParseTreeInspection {
  static readonly meta: InspectionMetadata = {
    id: 'ObsoleteLetStatement',
    tier: 'A',
    category: 'ObsoleteSyntax',
    defaultSeverity: 'suggestion',
    name: 'Obsolete Let statement',
    description: 'The Let keyword is obsolete and should be removed from value assignments.',
    quickFixDescription: 'Remove the Let keyword',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const results: InspectionResult[] = [];
    const visitor = new ObsoleteLetVisitor(this, results);
    visitor.visit(context.parseResult.tree);
    return results;
  }
}

class ObsoleteLetVisitor extends VBAParserVisitor<void> {
  constructor(
    private readonly inspection: ObsoleteLetStatementInspection,
    private readonly results: InspectionResult[],
  ) {
    super();
  }

  override visitLetStmt = (ctx: LetStmtContext): void => {
    if (ctx.LET()) {
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
