// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/UnreachableCaseInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { ParseTreeInspection, type InspectionContext, type InspectionMetadata } from '../../base.js';
import type { InspectionResult } from '../../types.js';
import { VBAParserVisitor } from '../../../parser/generated/grammar/VBAParserVisitor.js';
import type { BlockContext } from '../../../parser/generated/grammar/VBAParser.js';

/**
 * Detects code that appears after unconditional exit/termination statements.
 *
 * VBA:
 *   Sub Test()
 *       Exit Sub
 *       MsgBox "unreachable"  ' ← will never execute
 *   End Sub
 *
 * Note: This is a simplified parse-tree level check. It flags statements
 * after Exit Sub/Function/Property, End, GoTo, or Resume within the same block.
 * It does not track GoTo labels or conditional paths.
 */
export class UnreachableCodeInspection extends ParseTreeInspection {
  static readonly meta: InspectionMetadata = {
    id: 'UnreachableCode',
    tier: 'A',
    category: 'CodeQuality',
    defaultSeverity: 'warning',
    name: 'Unreachable code',
    description: 'Code after Exit Sub/Function, End, or GoTo is unreachable.',
    quickFixDescription: 'Remove the unreachable code',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const results: InspectionResult[] = [];
    const visitor = new UnreachableCodeVisitor(this, results);
    visitor.visit(context.parseResult.tree);
    return results;
  }
}

const EXIT_PATTERNS = /^(exit\s+(sub|function|property|do|for)|end|goto\s+)/i;

class UnreachableCodeVisitor extends VBAParserVisitor<void> {
  constructor(
    private readonly inspection: UnreachableCodeInspection,
    private readonly results: InspectionResult[],
  ) {
    super();
  }

  override visitBlock = (ctx: BlockContext): void => {
    const stmts = ctx.blockStmt();
    let foundExit = false;

    for (const stmt of stmts) {
      if (foundExit) {
        // Check if this is a label (labels are reachable via GoTo)
        const mainStmt = stmt.mainBlockStmt();
        const labelDef = stmt.statementLabelDefinition();
        if (labelDef) {
          // Label definitions make subsequent code reachable again
          foundExit = false;
          continue;
        }
        if (mainStmt) {
          const start = stmt.start!;
          const stop = stmt.stop!;
          this.results.push(
            (this.inspection as any).createResult({
              startLine: start.line,
              startColumn: start.column,
              endLine: stop.line,
              endColumn: stop.column + (stop.text?.length ?? 0),
            }),
          );
        }
      } else {
        const mainStmt = stmt.mainBlockStmt();
        if (mainStmt) {
          const text = mainStmt.getText().replace(/\s+/g, ' ').trim();
          if (EXIT_PATTERNS.test(text)) {
            foundExit = true;
          }
        }
      }
    }

    this.visitChildren(ctx);
  };
}
