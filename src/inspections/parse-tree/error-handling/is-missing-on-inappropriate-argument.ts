// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/IsMissingOnInappropriateArgumentInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { ParseTreeInspection, type InspectionContext, type InspectionMetadata } from '../../base.js';
import type { InspectionResult } from '../../types.js';
import { VBAParserVisitor } from '../../../parser/generated/grammar/VBAParserVisitor.js';
import type { IndexExprContext } from '../../../parser/generated/grammar/VBAParser.js';

/**
 * Detects use of IsMissing() function.
 *
 * VBA:
 *   If IsMissing(param) Then  ' ← IsMissing only works with Variant Optional params
 *
 * Note: Full implementation would require symbol resolution (Tier B) to check
 * if the argument is actually Optional Variant. This Tier A implementation
 * flags all IsMissing calls as an advisory notice.
 *
 * TODO: Move to Tier B when symbol resolution is available for full validation.
 */
export class IsMissingOnInappropriateArgumentInspection extends ParseTreeInspection {
  static readonly meta: InspectionMetadata = {
    id: 'IsMissingOnInappropriateArgument',
    tier: 'A',
    category: 'CodeQuality',
    defaultSeverity: 'hint',
    name: 'IsMissing usage',
    description: 'IsMissing only works correctly with Optional Variant parameters. Verify the argument type.',
    quickFixDescription: 'Verify the parameter is Optional Variant',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const results: InspectionResult[] = [];
    const visitor = new IsMissingVisitor(this, results);
    visitor.visit(context.parseResult.tree);
    return results;
  }
}

class IsMissingVisitor extends VBAParserVisitor<void> {
  constructor(
    private readonly inspection: IsMissingOnInappropriateArgumentInspection,
    private readonly results: InspectionResult[],
  ) {
    super();
  }

  override visitIndexExpr = (ctx: IndexExprContext): void => {
    // Look for IsMissing(...) calls
    const lExpr = ctx.lExpression();
    if (lExpr) {
      const name = lExpr.getText().toLowerCase();
      if (name === 'ismissing') {
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
