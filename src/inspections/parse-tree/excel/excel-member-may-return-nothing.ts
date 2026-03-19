// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/ExcelMemberMayReturnNothingInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { ParseTreeInspection, type InspectionContext, type InspectionMetadata } from '../../base.js';
import type { InspectionResult } from '../../types.js';
import { VBAParserVisitor } from '../../../parser/generated/grammar/VBAParserVisitor.js';
import type { MemberAccessExprContext } from '../../../parser/generated/grammar/VBAParser.js';

const MAY_RETURN_NOTHING = new Set([
  'find', 'findnext', 'findprevious',
]);

/**
 * Detects use of Excel members that may return Nothing (Find, FindNext, etc.).
 *
 * VBA:
 *   Set cell = Range("A:A").Find("value")
 *   cell.Value = "found"  ' ← cell may be Nothing!
 *
 * Note: Full implementation requires reference tracking (Tier B).
 * This Tier A check flags member access calls to Find/FindNext/FindPrevious.
 */
export class ExcelMemberMayReturnNothingInspection extends ParseTreeInspection {
  static readonly meta: InspectionMetadata = {
    id: 'ExcelMemberMayReturnNothing',
    tier: 'A',
    category: 'Excel',
    defaultSeverity: 'warning',
    hostLibraries: ['Excel'],
    name: 'Excel member may return Nothing',
    description: 'Find/FindNext/FindPrevious may return Nothing. Check the result before using it.',
    quickFixDescription: 'Add a Nothing check after the call',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const results: InspectionResult[] = [];
    const visitor = new MayReturnNothingVisitor(this, results);
    visitor.visit(context.parseResult.tree);
    return results;
  }
}

class MayReturnNothingVisitor extends VBAParserVisitor<void> {
  constructor(
    private readonly inspection: ExcelMemberMayReturnNothingInspection,
    private readonly results: InspectionResult[],
  ) {
    super();
  }

  override visitMemberAccessExpr = (ctx: MemberAccessExprContext): void => {
    // Check the member name (the part after the dot)
    const unrestricted = ctx.unrestrictedIdentifier();
    if (unrestricted) {
      const memberName = unrestricted.getText().toLowerCase();
      if (MAY_RETURN_NOTHING.has(memberName)) {
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
