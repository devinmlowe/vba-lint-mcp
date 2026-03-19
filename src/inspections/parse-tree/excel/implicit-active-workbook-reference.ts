// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/ImplicitActiveWorkbookReferenceInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { ParseTreeInspection, type InspectionContext, type InspectionMetadata } from '../../base.js';
import type { InspectionResult } from '../../types.js';
import { VBAParserVisitor } from '../../../parser/generated/grammar/VBAParserVisitor.js';
import type { SimpleNameExprContext } from '../../../parser/generated/grammar/VBAParser.js';

const ACTIVE_WORKBOOK_MEMBERS = new Set([
  'sheets', 'worksheets', 'names',
]);

/**
 * Detects unqualified references to Sheets, Worksheets, Names which
 * implicitly refer to the ActiveWorkbook.
 *
 * VBA:
 *   Sheets("Sheet1")      ' ← implicitly ActiveWorkbook.Sheets
 */
export class ImplicitActiveWorkbookReferenceInspection extends ParseTreeInspection {
  static readonly meta: InspectionMetadata = {
    id: 'ImplicitActiveWorkbookReference',
    tier: 'A',
    category: 'Excel',
    defaultSeverity: 'suggestion',
    hostLibraries: ['Excel'],
    name: 'Implicit ActiveWorkbook reference',
    description: 'Unqualified Sheets/Worksheets/Names implicitly refers to ActiveWorkbook.',
    quickFixDescription: 'Qualify with a specific Workbook object',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const results: InspectionResult[] = [];
    const visitor = new ImplicitActiveWorkbookVisitor(this, results);
    visitor.visit(context.parseResult.tree);
    return results;
  }
}

class ImplicitActiveWorkbookVisitor extends VBAParserVisitor<void> {
  constructor(
    private readonly inspection: ImplicitActiveWorkbookReferenceInspection,
    private readonly results: InspectionResult[],
  ) {
    super();
  }

  override visitSimpleNameExpr = (ctx: SimpleNameExprContext): void => {
    const name = ctx.getText().toLowerCase();
    if (ACTIVE_WORKBOOK_MEMBERS.has(name)) {
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
