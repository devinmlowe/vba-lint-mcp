// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/SelfAssignedDeclarationInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { ParseTreeInspection, type InspectionContext, type InspectionMetadata } from '../../base.js';
import type { InspectionResult } from '../../types.js';
import { VBAParserVisitor } from '../../../parser/generated/grammar/VBAParserVisitor.js';
import type { VariableStmtContext } from '../../../parser/generated/grammar/VBAParser.js';

/**
 * Detects variable declarations that use `As New` (self-assigned objects).
 *
 * VBA:
 *   Dim obj As New Collection  ' ← auto-instantiation can mask Nothing checks
 *
 * Note: In Rubberduck this checks for `As New` which is the "self-assigned" pattern.
 * True self-assignment (Dim x = x) is not valid VBA syntax.
 */
export class SelfAssignedDeclarationInspection extends ParseTreeInspection {
  static readonly meta: InspectionMetadata = {
    id: 'SelfAssignedDeclaration',
    tier: 'A',
    category: 'CodeQuality',
    defaultSeverity: 'suggestion',
    name: 'Self-assigned declaration',
    description: 'Variable uses As New which auto-instantiates and can mask Nothing checks.',
    quickFixDescription: 'Remove As New and instantiate explicitly with Set',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const results: InspectionResult[] = [];
    const visitor = new SelfAssignedVisitor(this, results);
    visitor.visit(context.parseResult.tree);
    return results;
  }
}

class SelfAssignedVisitor extends VBAParserVisitor<void> {
  constructor(
    private readonly inspection: SelfAssignedDeclarationInspection,
    private readonly results: InspectionResult[],
  ) {
    super();
  }

  override visitVariableStmt = (ctx: VariableStmtContext): void => {
    // Check the text for "As New" pattern
    const text = ctx.getText().replace(/\s+/g, ' ');
    if (/\bAs\s+New\b/i.test(text)) {
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
