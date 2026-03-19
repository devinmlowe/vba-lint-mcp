// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/StepOneIsRedundantInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { ParseTreeInspection, type InspectionContext, type InspectionMetadata } from '../../base.js';
import type { InspectionResult } from '../../types.js';
import { VBAParserVisitor } from '../../../parser/generated/grammar/VBAParserVisitor.js';
import type { ForNextStmtContext } from '../../../parser/generated/grammar/VBAParser.js';

/**
 * Detects For loops with an explicit `Step 1`, which is the default.
 *
 * VBA:
 *   For i = 1 To 10 Step 1  ' ← Step 1 is default and redundant
 *   Next i
 */
export class StepOneIsRedundantInspection extends ParseTreeInspection {
  static readonly meta: InspectionMetadata = {
    id: 'StepOneIsRedundant',
    tier: 'A',
    category: 'CodeQuality',
    defaultSeverity: 'hint',
    name: 'Step 1 is redundant',
    description: 'Step 1 is the default for For loops and is redundant.',
    quickFixDescription: 'Remove Step 1',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const results: InspectionResult[] = [];
    const visitor = new StepOneVisitor(this, results);
    visitor.visit(context.parseResult.tree);
    return results;
  }
}

class StepOneVisitor extends VBAParserVisitor<void> {
  constructor(
    private readonly inspection: StepOneIsRedundantInspection,
    private readonly results: InspectionResult[],
  ) {
    super();
  }

  override visitForNextStmt = (ctx: ForNextStmtContext): void => {
    const stepStmt = ctx.stepStmt();
    if (stepStmt) {
      const expr = stepStmt.expression();
      const stepText = expr.getText().trim();
      if (stepText === '1') {
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
