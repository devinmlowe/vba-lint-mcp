// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/ImplicitPublicMemberInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { ParseTreeInspection, type InspectionContext, type InspectionMetadata } from '../../base.js';
import type { InspectionResult } from '../../types.js';
import { VBAParserVisitor } from '../../../parser/generated/grammar/VBAParserVisitor.js';
import type {
  SubStmtContext,
  FunctionStmtContext,
  PropertyGetStmtContext,
  PropertySetStmtContext,
  PropertyLetStmtContext,
} from '../../../parser/generated/grammar/VBAParser.js';

/**
 * Detects Sub/Function/Property without explicit Public/Private modifier.
 *
 * VBA:
 *   Sub Foo()             ' ← implicitly Public; should be explicit
 */
export class ImplicitPublicMemberInspection extends ParseTreeInspection {
  static readonly meta: InspectionMetadata = {
    id: 'ImplicitPublicMember',
    tier: 'A',
    category: 'CodeQuality',
    defaultSeverity: 'suggestion',
    name: 'Implicit Public member',
    description: 'Member is implicitly Public. Specify Public or Private explicitly.',
    quickFixDescription: 'Add explicit Public or Private modifier',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const results: InspectionResult[] = [];
    const visitor = new ImplicitPublicVisitor(this, results);
    visitor.visit(context.parseResult.tree);
    return results;
  }
}

class ImplicitPublicVisitor extends VBAParserVisitor<void> {
  constructor(
    private readonly inspection: ImplicitPublicMemberInspection,
    private readonly results: InspectionResult[],
  ) {
    super();
  }

  private checkVisibility(ctx: { visibility(): any; start: any; stop: any }): void {
    if (!ctx.visibility()) {
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

  override visitSubStmt = (ctx: SubStmtContext): void => {
    this.checkVisibility(ctx as any);
    this.visitChildren(ctx);
  };

  override visitFunctionStmt = (ctx: FunctionStmtContext): void => {
    this.checkVisibility(ctx as any);
    this.visitChildren(ctx);
  };

  override visitPropertyGetStmt = (ctx: PropertyGetStmtContext): void => {
    this.checkVisibility(ctx as any);
    this.visitChildren(ctx);
  };

  override visitPropertySetStmt = (ctx: PropertySetStmtContext): void => {
    this.checkVisibility(ctx as any);
    this.visitChildren(ctx);
  };

  override visitPropertyLetStmt = (ctx: PropertyLetStmtContext): void => {
    this.checkVisibility(ctx as any);
    this.visitChildren(ctx);
  };
}
