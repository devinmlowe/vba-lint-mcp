// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/ObsoleteCallStatementInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { ParseTreeInspection, type InspectionContext, type InspectionMetadata } from '../../base.js';
import type { InspectionResult } from '../../types.js';
import { VBAParserVisitor } from '../../../parser/generated/grammar/VBAParserVisitor.js';
import type { CallStmtContext } from '../../../parser/generated/grammar/VBAParser.js';

/**
 * Detects use of the obsolete `Call` keyword.
 *
 * VBA:
 *   Call Foo()             ' ← obsolete; should be `Foo`
 */
export class ObsoleteCallStatementInspection extends ParseTreeInspection {
  static readonly meta: InspectionMetadata = {
    id: 'ObsoleteCallStatement',
    tier: 'A',
    category: 'ObsoleteSyntax',
    defaultSeverity: 'suggestion',
    name: 'Obsolete Call statement',
    description: 'The Call keyword is obsolete. Call the procedure directly without it.',
    quickFixDescription: 'Remove the Call keyword',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const results: InspectionResult[] = [];
    const visitor = new ObsoleteCallVisitor(this, results);
    visitor.visit(context.parseResult.tree);
    return results;
  }
}

class ObsoleteCallVisitor extends VBAParserVisitor<void> {
  constructor(
    private readonly inspection: ObsoleteCallStatementInspection,
    private readonly results: InspectionResult[],
  ) {
    super();
  }

  override visitCallStmt = (ctx: CallStmtContext): void => {
    if (ctx.CALL()) {
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
