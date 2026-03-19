// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { describe, it, expect } from 'vitest';
import { parseCode } from '../../../../parser/index.js';
import type { InspectionContext } from '../../../base.js';
import { readFile } from 'node:fs/promises';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

import { ImplicitActiveSheetReferenceInspection } from '../implicit-active-sheet-reference.js';
import { ImplicitActiveWorkbookReferenceInspection } from '../implicit-active-workbook-reference.js';
import { SheetAccessedUsingStringInspection } from '../sheet-accessed-using-string.js';
import { ApplicationWorksheetFunctionInspection } from '../application-worksheet-function.js';
import { ExcelMemberMayReturnNothingInspection } from '../excel-member-may-return-nothing.js';
import { ExcelUdfNameIsValidCellReferenceInspection } from '../excel-udf-name-is-valid-cell-reference.js';

const __dirname = dirname(fileURLToPath(import.meta.url));
const fixturesDir = join(__dirname, '..', '__fixtures__');

function inspectCode<T extends { inspect(ctx: InspectionContext): any }>(
  InspectionClass: new () => T,
  code: string,
) {
  const parseResult = parseCode(code);
  const context: InspectionContext = { parseResult };
  const inspection = new InspectionClass();
  return inspection.inspect(context);
}

// --- ImplicitActiveSheetReference ---
describe('ImplicitActiveSheetReferenceInspection', () => {
  it('detects unqualified Range', () => {
    const code = 'Sub Test()\n    Range("A1").Value = 1\nEnd Sub\n';
    const results = inspectCode(ImplicitActiveSheetReferenceInspection, code);
    expect(results.length).toBeGreaterThanOrEqual(1);
    expect(results[0].inspection).toBe('ImplicitActiveSheetReference');
  });

  it('does not flag qualified Range', () => {
    const code = 'Sub Test()\n    ws.Range("A1").Value = 1\nEnd Sub\n';
    const results = inspectCode(ImplicitActiveSheetReferenceInspection, code);
    // Qualified references use MemberAccessExpr, not SimpleNameExpr
    expect(results).toHaveLength(0);
  });

  it('detects from fixture', async () => {
    const code = await readFile(join(fixturesDir, 'implicit-active-sheet.bas'), 'utf-8');
    const results = inspectCode(ImplicitActiveSheetReferenceInspection, code);
    expect(results.length).toBeGreaterThanOrEqual(1);
  });
});

// --- ImplicitActiveWorkbookReference ---
describe('ImplicitActiveWorkbookReferenceInspection', () => {
  it('detects unqualified Sheets', () => {
    const code = 'Sub Test()\n    Sheets("Sheet1").Range("A1").Value = 1\nEnd Sub\n';
    const results = inspectCode(ImplicitActiveWorkbookReferenceInspection, code);
    expect(results.length).toBeGreaterThanOrEqual(1);
    expect(results[0].inspection).toBe('ImplicitActiveWorkbookReference');
  });

  it('does not flag qualified Sheets', () => {
    const code = 'Sub Test()\n    wb.Sheets("Sheet1").Range("A1").Value = 1\nEnd Sub\n';
    const results = inspectCode(ImplicitActiveWorkbookReferenceInspection, code);
    expect(results).toHaveLength(0);
  });

  it('detects from fixture', async () => {
    const code = await readFile(join(fixturesDir, 'implicit-active-workbook.bas'), 'utf-8');
    const results = inspectCode(ImplicitActiveWorkbookReferenceInspection, code);
    expect(results.length).toBeGreaterThanOrEqual(1);
  });
});

// --- SheetAccessedUsingString ---
describe('SheetAccessedUsingStringInspection', () => {
  it('detects Sheets("Name") pattern', () => {
    const code = 'Sub Test()\n    Sheets("Sheet1").Range("A1").Value = 1\nEnd Sub\n';
    const results = inspectCode(SheetAccessedUsingStringInspection, code);
    expect(results.length).toBeGreaterThanOrEqual(1);
    expect(results[0].inspection).toBe('SheetAccessedUsingString');
  });

  it('does not flag Sheets(index)', () => {
    const code = 'Sub Test()\n    Sheets(1).Range("A1").Value = 1\nEnd Sub\n';
    const results = inspectCode(SheetAccessedUsingStringInspection, code);
    expect(results).toHaveLength(0);
  });

  it('detects from fixture', async () => {
    const code = await readFile(join(fixturesDir, 'sheet-string-access.bas'), 'utf-8');
    const results = inspectCode(SheetAccessedUsingStringInspection, code);
    expect(results.length).toBeGreaterThanOrEqual(1);
  });
});

// --- ApplicationWorksheetFunction ---
describe('ApplicationWorksheetFunctionInspection', () => {
  it('detects Application.WorksheetFunction', () => {
    const code = 'Sub Test()\n    x = Application.WorksheetFunction.Sum(Range("A1:A10"))\nEnd Sub\n';
    const results = inspectCode(ApplicationWorksheetFunctionInspection, code);
    expect(results.length).toBeGreaterThanOrEqual(1);
    expect(results[0].inspection).toBe('ApplicationWorksheetFunction');
  });

  it('does not flag WorksheetFunction alone', () => {
    const code = 'Sub Test()\n    x = WorksheetFunction.Sum(Range("A1:A10"))\nEnd Sub\n';
    const results = inspectCode(ApplicationWorksheetFunctionInspection, code);
    expect(results).toHaveLength(0);
  });
});

// --- ExcelMemberMayReturnNothing ---
describe('ExcelMemberMayReturnNothingInspection', () => {
  it('detects .Find call', () => {
    const code = 'Sub Test()\n    Set cell = Range("A:A").Find("value")\nEnd Sub\n';
    const results = inspectCode(ExcelMemberMayReturnNothingInspection, code);
    expect(results.length).toBeGreaterThanOrEqual(1);
    expect(results[0].inspection).toBe('ExcelMemberMayReturnNothing');
  });

  it('does not flag non-Find members', () => {
    const code = 'Sub Test()\n    x = Range("A1").Value\nEnd Sub\n';
    const results = inspectCode(ExcelMemberMayReturnNothingInspection, code);
    expect(results).toHaveLength(0);
  });
});

// --- ExcelUdfNameIsValidCellReference ---
describe('ExcelUdfNameIsValidCellReferenceInspection', () => {
  it('detects function named like a cell reference', () => {
    const code = 'Public Function A1() As Long\n    A1 = 42\nEnd Function\n';
    const results = inspectCode(ExcelUdfNameIsValidCellReferenceInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('ExcelUdfNameIsValidCellReference');
  });

  it('does not flag normal function names', () => {
    const code = 'Public Function GetValue() As Long\n    GetValue = 42\nEnd Function\n';
    const results = inspectCode(ExcelUdfNameIsValidCellReferenceInspection, code);
    expect(results).toHaveLength(0);
  });

  it('does not flag private functions', () => {
    const code = 'Private Function A1() As Long\n    A1 = 42\nEnd Function\n';
    const results = inspectCode(ExcelUdfNameIsValidCellReferenceInspection, code);
    expect(results).toHaveLength(0);
  });

  it('detects from fixture', async () => {
    const code = await readFile(join(fixturesDir, 'udf-cell-reference.bas'), 'utf-8');
    const results = inspectCode(ExcelUdfNameIsValidCellReferenceInspection, code);
    expect(results).toHaveLength(1);
  });
});
