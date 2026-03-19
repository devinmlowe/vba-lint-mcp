// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/StepNotSpecifiedInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { ParseTreeInspection, type InspectionContext, type InspectionMetadata } from '../../base.js';
import type { InspectionResult } from '../../types.js';
import { VBAParserVisitor } from '../../../parser/generated/grammar/VBAParserVisitor.js';
import type { ForNextStmtContext } from '../../../parser/generated/grammar/VBAParser.js';

/**
 * Detects For loops without an explicit Step clause.
 *
 * VBA:
 *   For i = 1 To 10       ' ← implicit Step 1; consider being explicit
 *   Next i
 */
export class StepNotSpecifiedInspection extends ParseTreeInspection {
  static readonly meta: InspectionMetadata = {
    id: 'StepNotSpecified',
    tier: 'A',
    category: 'CodeQuality',
    defaultSeverity: 'hint',
    name: 'Step not specified',
    description: 'For loop does not specify Step. Consider adding an explicit Step clause.',
    quickFixDescription: 'Add Step 1',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const results: InspectionResult[] = [];
    const visitor = new StepNotSpecifiedVisitor(this, results);
    visitor.visit(context.parseResult.tree);
    return results;
  }
}

class StepNotSpecifiedVisitor extends VBAParserVisitor<void> {
  constructor(
    private readonly inspection: StepNotSpecifiedInspection,
    private readonly results: InspectionResult[],
  ) {
    super();
  }

  override visitForNextStmt = (ctx: ForNextStmtContext): void => {
    if (!ctx.stepStmt()) {
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
