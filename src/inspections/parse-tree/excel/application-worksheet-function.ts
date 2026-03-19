// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/ApplicationWorksheetFunctionInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { ParseTreeInspection, type InspectionContext, type InspectionMetadata } from '../../base.js';
import type { InspectionResult } from '../../types.js';
import { VBAParserVisitor } from '../../../parser/generated/grammar/VBAParserVisitor.js';
import type { MemberAccessExprContext } from '../../../parser/generated/grammar/VBAParser.js';

/**
 * Detects Application.WorksheetFunction.X instead of WorksheetFunction.X
 *
 * VBA:
 *   x = Application.WorksheetFunction.Sum(Range("A1:A10"))
 *   ' Should be: x = WorksheetFunction.Sum(Range("A1:A10"))
 */
export class ApplicationWorksheetFunctionInspection extends ParseTreeInspection {
  static readonly meta: InspectionMetadata = {
    id: 'ApplicationWorksheetFunction',
    tier: 'A',
    category: 'Excel',
    defaultSeverity: 'hint',
    hostLibraries: ['Excel'],
    name: 'Application.WorksheetFunction',
    description: 'Application.WorksheetFunction is redundant. Use WorksheetFunction directly.',
    quickFixDescription: 'Remove the Application. prefix',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const results: InspectionResult[] = [];
    const visitor = new AppWorksheetFuncVisitor(this, results);
    visitor.visit(context.parseResult.tree);
    return results;
  }
}

class AppWorksheetFuncVisitor extends VBAParserVisitor<void> {
  constructor(
    private readonly inspection: ApplicationWorksheetFunctionInspection,
    private readonly results: InspectionResult[],
  ) {
    super();
  }

  override visitMemberAccessExpr = (ctx: MemberAccessExprContext): void => {
    const text = ctx.getText().toLowerCase();
    if (text.startsWith('application.worksheetfunction')) {
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
