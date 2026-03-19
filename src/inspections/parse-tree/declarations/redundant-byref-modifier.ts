// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/RedundantByRefModifierInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { ParseTreeInspection, type InspectionContext, type InspectionMetadata } from '../../base.js';
import type { InspectionResult } from '../../types.js';
import { VBAParserVisitor } from '../../../parser/generated/grammar/VBAParserVisitor.js';
import type { ArgContext } from '../../../parser/generated/grammar/VBAParser.js';

/**
 * Detects explicit ByRef modifier, which is the default and thus redundant.
 *
 * VBA:
 *   Sub Foo(ByRef x As Long)  ' ← ByRef is default; remove for clarity
 */
export class RedundantByRefModifierInspection extends ParseTreeInspection {
  static readonly meta: InspectionMetadata = {
    id: 'RedundantByRefModifier',
    tier: 'A',
    category: 'CodeQuality',
    defaultSeverity: 'hint',
    name: 'Redundant ByRef modifier',
    description: 'ByRef is the default parameter passing mechanism and is redundant when specified.',
    quickFixDescription: 'Remove the ByRef modifier',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const results: InspectionResult[] = [];
    const visitor = new RedundantByRefVisitor(this, results);
    visitor.visit(context.parseResult.tree);
    return results;
  }
}

class RedundantByRefVisitor extends VBAParserVisitor<void> {
  constructor(
    private readonly inspection: RedundantByRefModifierInspection,
    private readonly results: InspectionResult[],
  ) {
    super();
  }

  override visitArg = (ctx: ArgContext): void => {
    if (ctx.BYREF()) {
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
