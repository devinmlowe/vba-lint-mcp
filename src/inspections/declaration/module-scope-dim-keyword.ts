// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/ModuleScopeDimKeywordInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { DeclarationInspection, type InspectionContext, type InspectionMetadata } from '../base.js';
import type { InspectionResult } from '../types.js';
import { ParseTreeWalker } from 'antlr4ng';
import { VBAParserListener } from '../../parser/generated/grammar/VBAParserListener.js';
import type { VariableStmtContext, ModuleBodyContext, ModuleBodyElementContext } from '../../parser/generated/grammar/VBAParser.js';

/**
 * Detects module-level variables declared with Dim instead of Private.
 *
 * VBA:
 *   Dim moduleVar As Long   ' ← at module level, should use Private
 *
 * Rubberduck rationale: At module level, Dim and Private are equivalent,
 * but Private makes the intended scope explicit and is more readable.
 *
 * Note: This inspection uses the parse tree to detect the actual Dim keyword,
 * since the symbol walker normalizes Dim at module level to Private.
 */
export class ModuleScopeDimKeywordInspection extends DeclarationInspection {
  static readonly meta: InspectionMetadata = {
    id: 'ModuleScopeDimKeyword',
    tier: 'B',
    category: 'CodeQuality',
    defaultSeverity: 'suggestion',
    name: 'Use Private instead of Dim at module level',
    description: 'Module-level variable uses Dim instead of Private.',
    quickFixDescription: 'Replace Dim with Private',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const finder = context.declarationFinder;
    if (!finder) return [];

    const results: InspectionResult[] = [];

    // Walk the parse tree to find module-level Dim statements
    const listener = new ModuleDimListener(results, this);
    const walker = ParseTreeWalker.DEFAULT;
    walker.walk(listener, context.parseResult.tree);

    return results;
  }

  /** Expose createResult for the listener */
  makeResult(
    location: { startLine: number; startColumn: number; endLine: number; endColumn: number },
    varName: string,
  ): InspectionResult {
    return this.createResult(location, {
      description: `Module-level variable '${varName}' uses Dim — use Private instead.`,
    });
  }
}

class ModuleDimListener extends VBAParserListener {
  // Track nesting — we only care about module-level (depth 0)
  private insideProcedure = false;

  constructor(
    private readonly results: InspectionResult[],
    private readonly inspection: ModuleScopeDimKeywordInspection,
  ) {
    super();
  }

  override enterVariableStmt = (ctx: VariableStmtContext): void => {
    if (this.insideProcedure) return;

    // Check if this is a Dim statement (no visibility keyword = Dim)
    if (ctx.DIM() && !ctx.visibility()) {
      // Get variable names from sub-statements
      const subStmts = ctx.variableListStmt().variableSubStmt();
      for (const sub of subStmts) {
        const name = sub.identifier().getText();
        const start = ctx.start!;
        const stop = sub.stop ?? sub.start!;
        this.results.push(
          this.inspection.makeResult(
            {
              startLine: start.line,
              startColumn: start.column,
              endLine: stop.line,
              endColumn: stop.column + (stop.text?.length ?? 0),
            },
            name,
          ),
        );
      }
    }
  };

  // Track procedure entry/exit to know if we're at module level
  override enterSubStmt = (): void => { this.insideProcedure = true; };
  override exitSubStmt = (): void => { this.insideProcedure = false; };
  override enterFunctionStmt = (): void => { this.insideProcedure = true; };
  override exitFunctionStmt = (): void => { this.insideProcedure = false; };
  override enterPropertyGetStmt = (): void => { this.insideProcedure = true; };
  override exitPropertyGetStmt = (): void => { this.insideProcedure = false; };
  override enterPropertyLetStmt = (): void => { this.insideProcedure = true; };
  override exitPropertyLetStmt = (): void => { this.insideProcedure = false; };
  override enterPropertySetStmt = (): void => { this.insideProcedure = true; };
  override exitPropertySetStmt = (): void => { this.insideProcedure = false; };
}
