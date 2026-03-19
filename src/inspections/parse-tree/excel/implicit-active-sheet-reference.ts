// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/ImplicitActiveSheetReferenceInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { ParseTreeInspection, type InspectionContext, type InspectionMetadata } from '../../base.js';
import type { InspectionResult } from '../../types.js';
import { VBAParserVisitor } from '../../../parser/generated/grammar/VBAParserVisitor.js';
import type { SimpleNameExprContext } from '../../../parser/generated/grammar/VBAParser.js';

const ACTIVE_SHEET_MEMBERS = new Set([
  'range', 'cells', 'rows', 'columns',
]);

/**
 * Detects unqualified references to Range, Cells, Rows, Columns which
 * implicitly refer to the ActiveSheet.
 *
 * VBA:
 *   Range("A1").Value = 1  ' ← implicitly ActiveSheet.Range
 *
 * Note: This is a simplified parse-tree check that flags unqualified
 * simple name references. Full qualification detection requires Tier B.
 */
export class ImplicitActiveSheetReferenceInspection extends ParseTreeInspection {
  static readonly meta: InspectionMetadata = {
    id: 'ImplicitActiveSheetReference',
    tier: 'A',
    category: 'Excel',
    defaultSeverity: 'suggestion',
    hostLibraries: ['Excel'],
    name: 'Implicit ActiveSheet reference',
    description: 'Unqualified Range/Cells/Rows/Columns implicitly refers to ActiveSheet.',
    quickFixDescription: 'Qualify with a specific Worksheet object',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const results: InspectionResult[] = [];
    const visitor = new ImplicitActiveSheetVisitor(this, results);
    visitor.visit(context.parseResult.tree);
    return results;
  }
}

class ImplicitActiveSheetVisitor extends VBAParserVisitor<void> {
  constructor(
    private readonly inspection: ImplicitActiveSheetReferenceInspection,
    private readonly results: InspectionResult[],
  ) {
    super();
  }

  override visitSimpleNameExpr = (ctx: SimpleNameExprContext): void => {
    const name = ctx.getText().toLowerCase();
    if (ACTIVE_SHEET_MEMBERS.has(name)) {
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
