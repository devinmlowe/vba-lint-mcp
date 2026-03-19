// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/UnhandledOnErrorResumeNextInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { ParseTreeInspection, type InspectionContext, type InspectionMetadata } from '../../base.js';
import type { InspectionResult } from '../../types.js';
import { VBAParserVisitor } from '../../../parser/generated/grammar/VBAParserVisitor.js';
import type {
  SubStmtContext,
  FunctionStmtContext,
  OnErrorStmtContext,
} from '../../../parser/generated/grammar/VBAParser.js';

/**
 * Detects `On Error Resume Next` without a corresponding `On Error GoTo 0` reset.
 *
 * VBA:
 *   Sub Test()
 *       On Error Resume Next
 *       ' ... code ...
 *       ' Missing: On Error GoTo 0
 *   End Sub
 */
export class UnhandledOnErrorResumeNextInspection extends ParseTreeInspection {
  static readonly meta: InspectionMetadata = {
    id: 'UnhandledOnErrorResumeNext',
    tier: 'A',
    category: 'ErrorHandling',
    defaultSeverity: 'warning',
    name: 'Unhandled On Error Resume Next',
    description: 'On Error Resume Next is used without a corresponding On Error GoTo 0 to reset error handling.',
    quickFixDescription: 'Add On Error GoTo 0 after the protected section',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const results: InspectionResult[] = [];
    const visitor = new UnhandledErrorVisitor(this, results);
    visitor.visit(context.parseResult.tree);
    return results;
  }
}

class UnhandledErrorVisitor extends VBAParserVisitor<void> {
  constructor(
    private readonly inspection: UnhandledOnErrorResumeNextInspection,
    private readonly results: InspectionResult[],
  ) {
    super();
  }

  private checkMethod(ctx: { block(): { blockStmt(): any[] }; start: any; stop: any }): void {
    const block = ctx.block();
    if (!block) return;

    const stmts = block.blockStmt();
    const resumeNextLocations: Array<{ line: number; column: number; stopLine: number; stopColumn: number }> = [];
    let hasGoTo0 = false;

    for (const stmt of stmts) {
      const mainStmt = stmt.mainBlockStmt();
      if (!mainStmt) continue;
      const text = mainStmt.getText().replace(/\s+/g, ' ').trim().toLowerCase();

      if (text.includes('onerror') || text.includes('on error') || text.includes('onlocalerror') || text.includes('on local error')) {
        if (text.includes('resumenext') || text.includes('resume next')) {
          resumeNextLocations.push({
            line: stmt.start!.line,
            column: stmt.start!.column,
            stopLine: stmt.stop!.line,
            stopColumn: stmt.stop!.column + (stmt.stop!.text?.length ?? 0),
          });
        }
        if (text.includes('goto') && text.includes('0')) {
          hasGoTo0 = true;
        }
      }
    }

    if (!hasGoTo0) {
      for (const loc of resumeNextLocations) {
        this.results.push(
          (this.inspection as any).createResult({
            startLine: loc.line,
            startColumn: loc.column,
            endLine: loc.stopLine,
            endColumn: loc.stopColumn,
          }),
        );
      }
    }
  }

  override visitSubStmt = (ctx: SubStmtContext): void => {
    this.checkMethod(ctx as any);
    this.visitChildren(ctx);
  };

  override visitFunctionStmt = (ctx: FunctionStmtContext): void => {
    this.checkMethod(ctx as any);
    this.visitChildren(ctx);
  };
}
