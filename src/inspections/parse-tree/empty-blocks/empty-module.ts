// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/EmptyModuleInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { ParseTreeInspection, type InspectionContext, type InspectionMetadata } from '../../base.js';
import type { InspectionResult } from '../../types.js';
import { VBAParserVisitor } from '../../../parser/generated/grammar/VBAParserVisitor.js';
import type { ModuleContext } from '../../../parser/generated/grammar/VBAParser.js';

/**
 * Detects modules with no declarations or procedures.
 *
 * VBA:
 *   Option Explicit
 *   ' (nothing else)     ' ← empty module
 */
export class EmptyModuleInspection extends ParseTreeInspection {
  static readonly meta: InspectionMetadata = {
    id: 'EmptyModule',
    tier: 'A',
    category: 'CodeQuality',
    defaultSeverity: 'hint',
    name: 'Empty module',
    description: 'Module contains no declarations or procedures and may be unnecessary.',
    quickFixDescription: 'Remove the empty module',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const results: InspectionResult[] = [];
    const visitor = new EmptyModuleVisitor(this, results);
    visitor.visit(context.parseResult.tree);
    return results;
  }
}

class EmptyModuleVisitor extends VBAParserVisitor<void> {
  constructor(
    private readonly inspection: EmptyModuleInspection,
    private readonly results: InspectionResult[],
  ) {
    super();
  }

  override visitModule = (ctx: ModuleContext): void => {
    const declarations = ctx.moduleDeclarations();
    const body = ctx.moduleBody();

    // Check if declarations section has any meaningful elements
    // (moduleOption like Option Explicit doesn't count as "content")
    const declElements = declarations?.moduleDeclarationsElement() ?? [];
    const hasMeaningfulDeclarations = declElements.some(el => {
      // Module options (Option Explicit, Option Compare, etc.) don't count
      if (el.moduleOption()) return false;
      // Attributes don't count
      if (el.attributeStmt()) return false;
      return true;
    });

    // Check if body has any procedures
    const bodyElements = body?.moduleBodyElement() ?? [];
    const hasProcedures = bodyElements.length > 0;

    if (!hasMeaningfulDeclarations && !hasProcedures) {
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

    // No need to visit children for this inspection
  };
}
