// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/SheetAccessedUsingStringInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { ParseTreeInspection, type InspectionContext, type InspectionMetadata } from '../../base.js';
import type { InspectionResult } from '../../types.js';
import { VBAParserVisitor } from '../../../parser/generated/grammar/VBAParserVisitor.js';
import type { IndexExprContext } from '../../../parser/generated/grammar/VBAParser.js';

/**
 * Detects Sheets("Name") or Worksheets("Name") access using string literals.
 *
 * VBA:
 *   Sheets("Sheet1")      ' ← fragile; use codename instead
 *
 * Note: Full implementation requires symbol resolution to know the member type.
 * This Tier A implementation detects the pattern syntactically by looking for
 * Sheets/Worksheets followed by a string argument.
 */
export class SheetAccessedUsingStringInspection extends ParseTreeInspection {
  static readonly meta: InspectionMetadata = {
    id: 'SheetAccessedUsingString',
    tier: 'A',
    category: 'Excel',
    defaultSeverity: 'suggestion',
    hostLibraries: ['Excel'],
    name: 'Sheet accessed using string',
    description: 'Accessing sheets by string name is fragile. Use the sheet codename property instead.',
    quickFixDescription: 'Replace with sheet codename',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const results: InspectionResult[] = [];
    const visitor = new SheetStringVisitor(this, results);
    visitor.visit(context.parseResult.tree);
    return results;
  }
}

class SheetStringVisitor extends VBAParserVisitor<void> {
  constructor(
    private readonly inspection: SheetAccessedUsingStringInspection,
    private readonly results: InspectionResult[],
  ) {
    super();
  }

  override visitIndexExpr = (ctx: IndexExprContext): void => {
    const lExpr = ctx.lExpression();
    if (lExpr) {
      const name = lExpr.getText().toLowerCase();
      if (name === 'sheets' || name === 'worksheets') {
        // Check if argument contains a string literal
        const argList = ctx.argumentList();
        if (argList) {
          const text = argList.getText();
          if (text.includes('"')) {
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
