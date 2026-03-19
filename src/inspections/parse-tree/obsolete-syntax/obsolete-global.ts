// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/ObsoleteGlobalInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { ParseTreeInspection, type InspectionContext, type InspectionMetadata } from '../../base.js';
import type { InspectionResult } from '../../types.js';
import { VBAParserVisitor } from '../../../parser/generated/grammar/VBAParserVisitor.js';
import type { VisibilityContext } from '../../../parser/generated/grammar/VBAParser.js';

/**
 * Detects use of the obsolete `Global` keyword instead of `Public`.
 *
 * VBA:
 *   Global x As Long      ' ← obsolete; should be `Public x As Long`
 */
export class ObsoleteGlobalInspection extends ParseTreeInspection {
  static readonly meta: InspectionMetadata = {
    id: 'ObsoleteGlobal',
    tier: 'A',
    category: 'ObsoleteSyntax',
    defaultSeverity: 'suggestion',
    name: 'Obsolete Global keyword',
    description: 'The Global keyword is obsolete. Use Public instead.',
    quickFixDescription: 'Replace Global with Public',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const results: InspectionResult[] = [];
    const visitor = new ObsoleteGlobalVisitor(this, results);
    visitor.visit(context.parseResult.tree);
    return results;
  }
}

class ObsoleteGlobalVisitor extends VBAParserVisitor<void> {
  constructor(
    private readonly inspection: ObsoleteGlobalInspection,
    private readonly results: InspectionResult[],
  ) {
    super();
  }

  override visitVisibility = (ctx: VisibilityContext): void => {
    if (ctx.GLOBAL()) {
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
