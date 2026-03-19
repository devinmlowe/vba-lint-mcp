// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/MultipleDeclarationsInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { ParseTreeInspection, type InspectionContext, type InspectionMetadata } from '../../base.js';
import type { InspectionResult } from '../../types.js';
import { VBAParserVisitor } from '../../../parser/generated/grammar/VBAParserVisitor.js';
import type { VariableStmtContext } from '../../../parser/generated/grammar/VBAParser.js';

/**
 * Detects multiple variable declarations on a single line.
 *
 * VBA:
 *   Dim a, b, c           ' ← should be separate Dim statements
 */
export class MultipleDeclarationsInspection extends ParseTreeInspection {
  static readonly meta: InspectionMetadata = {
    id: 'MultipleDeclarations',
    tier: 'A',
    category: 'CodeQuality',
    defaultSeverity: 'suggestion',
    name: 'Multiple declarations',
    description: 'Multiple variables declared on a single line. Declare each on its own line.',
    quickFixDescription: 'Split into separate declarations',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const results: InspectionResult[] = [];
    const visitor = new MultipleDeclarationsVisitor(this, results);
    visitor.visit(context.parseResult.tree);
    return results;
  }
}

class MultipleDeclarationsVisitor extends VBAParserVisitor<void> {
  constructor(
    private readonly inspection: MultipleDeclarationsInspection,
    private readonly results: InspectionResult[],
  ) {
    super();
  }

  override visitVariableStmt = (ctx: VariableStmtContext): void => {
    const varList = ctx.variableListStmt();
    if (varList) {
      const subStmts = varList.variableSubStmt();
      if (subStmts.length > 1) {
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
