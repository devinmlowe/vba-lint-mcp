// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/ImplicitVariantReturnTypeInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { ParseTreeInspection, type InspectionContext, type InspectionMetadata } from '../../base.js';
import type { InspectionResult } from '../../types.js';
import { VBAParserVisitor } from '../../../parser/generated/grammar/VBAParserVisitor.js';
import type {
  FunctionStmtContext,
  PropertyGetStmtContext,
} from '../../../parser/generated/grammar/VBAParser.js';

/**
 * Detects Functions and Property Get without an explicit return type.
 *
 * VBA:
 *   Function Foo()        ' ← returns Variant implicitly; specify As Type
 */
export class ImplicitVariantReturnTypeInspection extends ParseTreeInspection {
  static readonly meta: InspectionMetadata = {
    id: 'ImplicitVariantReturnType',
    tier: 'A',
    category: 'CodeQuality',
    defaultSeverity: 'suggestion',
    name: 'Implicit Variant return type',
    description: 'Function/Property Get returns Variant implicitly. Specify an explicit return type.',
    quickFixDescription: 'Add As Type clause',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const results: InspectionResult[] = [];
    const visitor = new ImplicitReturnVisitor(this, results);
    visitor.visit(context.parseResult.tree);
    return results;
  }
}

class ImplicitReturnVisitor extends VBAParserVisitor<void> {
  constructor(
    private readonly inspection: ImplicitVariantReturnTypeInspection,
    private readonly results: InspectionResult[],
  ) {
    super();
  }

  override visitFunctionStmt = (ctx: FunctionStmtContext): void => {
    if (!ctx.asTypeClause()) {
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

  override visitPropertyGetStmt = (ctx: PropertyGetStmtContext): void => {
    if (!ctx.asTypeClause()) {
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
