// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/OnLocalErrorInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { ParseTreeInspection, type InspectionContext, type InspectionMetadata } from '../../base.js';
import type { InspectionResult } from '../../types.js';
import { VBAParserVisitor } from '../../../parser/generated/grammar/VBAParserVisitor.js';
import type { OnErrorStmtContext } from '../../../parser/generated/grammar/VBAParser.js';

/**
 * Detects use of `On Local Error` which is functionally identical to `On Error`.
 *
 * VBA:
 *   On Local Error Resume Next  ' ← Local is redundant
 */
export class OnLocalErrorInspection extends ParseTreeInspection {
  static readonly meta: InspectionMetadata = {
    id: 'OnLocalError',
    tier: 'A',
    category: 'CodeQuality',
    defaultSeverity: 'suggestion',
    name: 'On Local Error',
    description: 'On Local Error is functionally identical to On Error. The Local keyword is redundant.',
    quickFixDescription: 'Replace with On Error',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const results: InspectionResult[] = [];
    const visitor = new OnLocalErrorVisitor(this, results);
    visitor.visit(context.parseResult.tree);
    return results;
  }
}

class OnLocalErrorVisitor extends VBAParserVisitor<void> {
  constructor(
    private readonly inspection: OnLocalErrorInspection,
    private readonly results: InspectionResult[],
  ) {
    super();
  }

  override visitOnErrorStmt = (ctx: OnErrorStmtContext): void => {
    if (ctx.ON_LOCAL_ERROR()) {
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
