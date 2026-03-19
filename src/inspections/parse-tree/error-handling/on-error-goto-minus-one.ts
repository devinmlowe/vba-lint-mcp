// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/OnErrorGoToMinusOneInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { ParseTreeInspection, type InspectionContext, type InspectionMetadata } from '../../base.js';
import type { InspectionResult } from '../../types.js';
import { VBAParserVisitor } from '../../../parser/generated/grammar/VBAParserVisitor.js';
import type { OnErrorStmtContext } from '../../../parser/generated/grammar/VBAParser.js';

/**
 * Detects `On Error GoTo -1` which clears the current error object.
 *
 * VBA:
 *   On Error GoTo -1      ' ← clears Err object; often misunderstood
 */
export class OnErrorGoToMinusOneInspection extends ParseTreeInspection {
  static readonly meta: InspectionMetadata = {
    id: 'OnErrorGoToMinusOne',
    tier: 'A',
    category: 'ErrorHandling',
    defaultSeverity: 'warning',
    name: 'On Error GoTo -1',
    description: 'On Error GoTo -1 clears the current error object. Ensure this is intentional.',
    quickFixDescription: 'Review error handling logic',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const results: InspectionResult[] = [];
    const visitor = new OnErrorGoToMinusOneVisitor(this, results);
    visitor.visit(context.parseResult.tree);
    return results;
  }
}

class OnErrorGoToMinusOneVisitor extends VBAParserVisitor<void> {
  constructor(
    private readonly inspection: OnErrorGoToMinusOneInspection,
    private readonly results: InspectionResult[],
  ) {
    super();
  }

  override visitOnErrorStmt = (ctx: OnErrorStmtContext): void => {
    // Check for GoTo -1 pattern
    if (ctx.GOTO()) {
      const expr = ctx.expression();
      if (expr) {
        const text = expr.getText().trim();
        // Check for -1 (could be a unary minus expression)
        if (text === '-1') {
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
    this.visitChildren(ctx);
  };
}
