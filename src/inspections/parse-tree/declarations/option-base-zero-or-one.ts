// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/OptionBaseZeroOrOneInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { ParseTreeInspection, type InspectionContext, type InspectionMetadata } from '../../base.js';
import type { InspectionResult } from '../../types.js';
import { VBAParserVisitor } from '../../../parser/generated/grammar/VBAParserVisitor.js';
import type { OptionBaseStmtContext } from '../../../parser/generated/grammar/VBAParser.js';

/**
 * Detects use of `Option Base 0` or `Option Base 1`.
 *
 * Option Base changes the default lower bound of arrays, which can cause
 * confusion and inconsistency.
 */
export class OptionBaseZeroOrOneInspection extends ParseTreeInspection {
  static readonly meta: InspectionMetadata = {
    id: 'OptionBaseZeroOrOne',
    tier: 'A',
    category: 'CodeQuality',
    defaultSeverity: 'hint',
    name: 'Option Base statement',
    description: 'Option Base changes the default array lower bound and can cause confusion.',
    quickFixDescription: 'Remove Option Base and use explicit lower bounds',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const results: InspectionResult[] = [];
    const visitor = new OptionBaseVisitor(this, results);
    visitor.visit(context.parseResult.tree);
    return results;
  }
}

class OptionBaseVisitor extends VBAParserVisitor<void> {
  constructor(
    private readonly inspection: OptionBaseZeroOrOneInspection,
    private readonly results: InspectionResult[],
  ) {
    super();
  }

  override visitOptionBaseStmt = (ctx: OptionBaseStmtContext): void => {
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
