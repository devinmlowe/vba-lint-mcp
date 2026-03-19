// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/EmptyStringLiteralInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { ParseTreeInspection, type InspectionContext, type InspectionMetadata } from '../../base.js';
import type { InspectionResult } from '../../types.js';
import { VBAParserVisitor } from '../../../parser/generated/grammar/VBAParserVisitor.js';
import type { LiteralExpressionContext } from '../../../parser/generated/grammar/VBAParser.js';

/**
 * Detects use of empty string literal "" instead of vbNullString.
 *
 * VBA:
 *   If str = "" Then      ' ← use vbNullString for better performance
 *
 * Note: vbNullString is a null pointer, while "" allocates memory for an
 * empty string. Using vbNullString is slightly more performant.
 */
export class EmptyStringLiteralInspection extends ParseTreeInspection {
  static readonly meta: InspectionMetadata = {
    id: 'EmptyStringLiteral',
    tier: 'A',
    category: 'Performance',
    defaultSeverity: 'hint',
    name: 'Empty string literal',
    description: 'Use vbNullString instead of "" for better performance.',
    quickFixDescription: 'Replace "" with vbNullString',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const results: InspectionResult[] = [];
    const visitor = new EmptyStringVisitor(this, results);
    visitor.visit(context.parseResult.tree);
    return results;
  }
}

class EmptyStringVisitor extends VBAParserVisitor<void> {
  constructor(
    private readonly inspection: EmptyStringLiteralInspection,
    private readonly results: InspectionResult[],
  ) {
    super();
  }

  override visitLiteralExpression = (ctx: LiteralExpressionContext): void => {
    const strLiteral = ctx.STRINGLITERAL();
    if (strLiteral) {
      const text = strLiteral.getText();
      if (text === '""') {
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
