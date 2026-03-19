// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/EmptyMethodInspection.cs
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
 * Detects Sub/Function/Property procedures with empty bodies.
 *
 * VBA:
 *   Sub DoNothing()
 *   End Sub               ' ← empty body
 */
export class EmptyMethodInspection extends ParseTreeInspection {
  static readonly meta: InspectionMetadata = {
    id: 'EmptyMethod',
    tier: 'A',
    category: 'CodeQuality',
    defaultSeverity: 'suggestion',
    name: 'Empty method',
    description: 'Method body is empty and should either contain code or be removed.',
    quickFixDescription: 'Remove the empty method',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const results: InspectionResult[] = [];
    const visitor = new EmptyMethodVisitor(this, results);
    visitor.visit(context.parseResult.tree);
    return results;
  }
}

class EmptyMethodVisitor extends VBAParserVisitor<void> {
  constructor(
    private readonly inspection: EmptyMethodInspection,
    private readonly results: InspectionResult[],
  ) {
    super();
  }

  private checkBlock(ctx: { block(): { blockStmt(): any[] }; start: any; stop: any }): void {
    const block = ctx.block();
    if (block && block.blockStmt().length === 0) {
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
    this.checkBlock(ctx as any);
    this.visitChildren(ctx);
  };

  override visitFunctionStmt = (ctx: FunctionStmtContext): void => {
    this.checkBlock(ctx as any);
    this.visitChildren(ctx);
  };

  override visitPropertyGetStmt = (ctx: PropertyGetStmtContext): void => {
    this.checkBlock(ctx as any);
    this.visitChildren(ctx);
  };

  override visitPropertySetStmt = (ctx: PropertySetStmtContext): void => {
    this.checkBlock(ctx as any);
    this.visitChildren(ctx);
  };

  override visitPropertyLetStmt = (ctx: PropertyLetStmtContext): void => {
    this.checkBlock(ctx as any);
    this.visitChildren(ctx);
  };
}
