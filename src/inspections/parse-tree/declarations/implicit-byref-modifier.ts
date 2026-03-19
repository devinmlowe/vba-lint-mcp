// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/ImplicitByRefModifierInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { ParseTreeInspection, type InspectionContext, type InspectionMetadata } from '../../base.js';
import type { InspectionResult } from '../../types.js';
import { VBAParserVisitor } from '../../../parser/generated/grammar/VBAParserVisitor.js';
import type { ArgContext } from '../../../parser/generated/grammar/VBAParser.js';

/**
 * Detects parameters without explicit ByRef/ByVal modifier.
 *
 * VBA:
 *   Sub Foo(x As Long)    ' ← implicit ByRef; should be explicit
 */
export class ImplicitByRefModifierInspection extends ParseTreeInspection {
  static readonly meta: InspectionMetadata = {
    id: 'ImplicitByRefModifier',
    tier: 'A',
    category: 'CodeQuality',
    defaultSeverity: 'suggestion',
    name: 'Implicit ByRef modifier',
    description: 'Parameter is implicitly passed ByRef. Specify ByRef or ByVal explicitly.',
    quickFixDescription: 'Add explicit ByRef or ByVal',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const results: InspectionResult[] = [];
    const visitor = new ImplicitByRefVisitor(this, results);
    visitor.visit(context.parseResult.tree);
    return results;
  }
}

class ImplicitByRefVisitor extends VBAParserVisitor<void> {
  constructor(
    private readonly inspection: ImplicitByRefModifierInspection,
    private readonly results: InspectionResult[],
  ) {
    super();
  }

  override visitArg = (ctx: ArgContext): void => {
    // Skip ParamArray parameters (always ByRef by definition)
    if (ctx.PARAMARRAY()) {
      this.visitChildren(ctx);
      return;
    }
    if (!ctx.BYVAL() && !ctx.BYREF()) {
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
