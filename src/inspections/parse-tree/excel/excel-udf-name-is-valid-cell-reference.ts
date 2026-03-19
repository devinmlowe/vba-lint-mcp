// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/ExcelUdfNameIsValidCellReferenceInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { ParseTreeInspection, type InspectionContext, type InspectionMetadata } from '../../base.js';
import type { InspectionResult } from '../../types.js';
import { VBAParserVisitor } from '../../../parser/generated/grammar/VBAParserVisitor.js';
import type { FunctionStmtContext } from '../../../parser/generated/grammar/VBAParser.js';

// Pattern matching valid Excel cell references: A1, B2, AA1, XFD1048576, etc.
const CELL_REF_PATTERN = /^[A-Z]{1,3}\d+$/i;

/**
 * Detects public Function names that look like Excel cell references.
 *
 * VBA:
 *   Public Function A1() As Long  ' ← conflicts with cell reference A1
 *
 * Such functions cannot be used as UDFs in Excel formulas because the
 * name is interpreted as a cell reference.
 */
export class ExcelUdfNameIsValidCellReferenceInspection extends ParseTreeInspection {
  static readonly meta: InspectionMetadata = {
    id: 'ExcelUdfNameIsValidCellReference',
    tier: 'A',
    category: 'Excel',
    defaultSeverity: 'warning',
    hostLibraries: ['Excel'],
    name: 'UDF name is a valid cell reference',
    description: 'Function name looks like a cell reference (e.g., A1, B2) and cannot be used as a UDF in Excel formulas.',
    quickFixDescription: 'Rename the function',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const results: InspectionResult[] = [];
    const visitor = new UdfCellRefVisitor(this, results);
    visitor.visit(context.parseResult.tree);
    return results;
  }
}

class UdfCellRefVisitor extends VBAParserVisitor<void> {
  constructor(
    private readonly inspection: ExcelUdfNameIsValidCellReferenceInspection,
    private readonly results: InspectionResult[],
  ) {
    super();
  }

  override visitFunctionStmt = (ctx: FunctionStmtContext): void => {
    // Only check public functions (which can be UDFs)
    const visibility = ctx.visibility();
    const isPrivate = visibility?.PRIVATE();
    if (!isPrivate) {
      const funcName = ctx.functionName();
      if (funcName) {
        const name = funcName.getText();
        if (CELL_REF_PATTERN.test(name)) {
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
    }
    this.visitChildren(ctx);
  };
}
