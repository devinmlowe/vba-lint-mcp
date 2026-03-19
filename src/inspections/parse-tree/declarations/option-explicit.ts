// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/OptionExplicitInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { ParseTreeInspection, type InspectionContext, type InspectionMetadata } from '../../base.js';
import type { InspectionResult } from '../../types.js';
import { VBAParserVisitor } from '../../../parser/generated/grammar/VBAParserVisitor.js';
import type {
  ModuleContext,
  OptionExplicitStmtContext,
} from '../../../parser/generated/grammar/VBAParser.js';

/**
 * Detects modules missing `Option Explicit`.
 *
 * Without Option Explicit, VBA allows implicit variable declarations,
 * which can lead to subtle bugs from typos.
 */
export class OptionExplicitInspection extends ParseTreeInspection {
  static readonly meta: InspectionMetadata = {
    id: 'OptionExplicit',
    tier: 'A',
    category: 'CodeQuality',
    defaultSeverity: 'warning',
    name: 'Missing Option Explicit',
    description: 'Module is missing Option Explicit. Variables should be explicitly declared.',
    quickFixDescription: 'Add Option Explicit',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const results: InspectionResult[] = [];
    const visitor = new OptionExplicitVisitor(this, results);
    visitor.visit(context.parseResult.tree);
    return results;
  }
}

class OptionExplicitVisitor extends VBAParserVisitor<void> {
  private hasOptionExplicit = false;

  constructor(
    private readonly inspection: OptionExplicitInspection,
    private readonly results: InspectionResult[],
  ) {
    super();
  }

  override visitOptionExplicitStmt = (_ctx: OptionExplicitStmtContext): void => {
    this.hasOptionExplicit = true;
  };

  override visitModule = (ctx: ModuleContext): void => {
    this.hasOptionExplicit = false;
    this.visitChildren(ctx);
    if (!this.hasOptionExplicit) {
      const start = ctx.start!;
      this.results.push(
        (this.inspection as any).createResult({
          startLine: start.line,
          startColumn: start.column,
          endLine: start.line,
          endColumn: start.column,
        }),
      );
    }
  };
}
