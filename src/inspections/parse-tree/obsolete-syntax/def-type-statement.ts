// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/DefTypeStatementInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { ParseTreeInspection, type InspectionContext, type InspectionMetadata } from '../../base.js';
import type { InspectionResult } from '../../types.js';
import { VBAParserVisitor } from '../../../parser/generated/grammar/VBAParserVisitor.js';
import type { DefDirectiveContext } from '../../../parser/generated/grammar/VBAParser.js';

/**
 * Detects DefBool, DefInt, DefLng, etc. statements.
 *
 * VBA:
 *   DefInt A-Z            ' ← obsolete; use explicit type declarations
 */
export class DefTypeStatementInspection extends ParseTreeInspection {
  static readonly meta: InspectionMetadata = {
    id: 'DefTypeStatement',
    tier: 'A',
    category: 'ObsoleteSyntax',
    defaultSeverity: 'suggestion',
    name: 'DefType statement',
    description: 'DefType statements (DefBool, DefInt, etc.) implicitly type variables and should be replaced with explicit declarations.',
    quickFixDescription: 'Remove the DefType statement and add explicit types',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const results: InspectionResult[] = [];
    const visitor = new DefTypeVisitor(this, results);
    visitor.visit(context.parseResult.tree);
    return results;
  }
}

class DefTypeVisitor extends VBAParserVisitor<void> {
  constructor(
    private readonly inspection: DefTypeStatementInspection,
    private readonly results: InspectionResult[],
  ) {
    super();
  }

  override visitDefDirective = (ctx: DefDirectiveContext): void => {
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
    this.visitChildren(ctx);
  };
}
