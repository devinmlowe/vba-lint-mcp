// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { describe, it, expect } from 'vitest';
import { parseCode } from '../../../../parser/index.js';
import type { InspectionContext } from '../../../base.js';
import { readFile } from 'node:fs/promises';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

import { BooleanAssignedInIfElseInspection } from '../boolean-assigned-in-if-else.js';
import { SelfAssignedDeclarationInspection } from '../self-assigned-declaration.js';
import { UnreachableCodeInspection } from '../unreachable-code.js';
import { LineContinuationBetweenKeywordsInspection } from '../line-continuation-between-keywords.js';
import { OnLocalErrorInspection } from '../on-local-error.js';
import { StepNotSpecifiedInspection } from '../step-not-specified.js';
import { StepOneIsRedundantInspection } from '../step-one-is-redundant.js';

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

// --- BooleanAssignedInIfElse ---
describe('BooleanAssignedInIfElseInspection', () => {
  it('detects boolean assignment in If/Else', () => {
    const code = [
      'Sub Test()',
      '    If x > 0 Then',
      '        result = True',
      '    Else',
      '        result = False',
      '    End If',
      'End Sub',
      '',
    ].join('\n');
    const results = inspectCode(BooleanAssignedInIfElseInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('BooleanAssignedInIfElse');
  });

  it('does not flag non-boolean If/Else', () => {
    const code = [
      'Sub Test()',
      '    If x > 0 Then',
      '        result = "positive"',
      '    Else',
      '        result = "non-positive"',
      '    End If',
      'End Sub',
      '',
    ].join('\n');
    const results = inspectCode(BooleanAssignedInIfElseInspection, code);
    expect(results).toHaveLength(0);
  });

  it('detects from fixture', async () => {
    const code = await readFile(join(fixturesDir, 'boolean-in-if-else.bas'), 'utf-8');
    const results = inspectCode(BooleanAssignedInIfElseInspection, code);
    expect(results).toHaveLength(1);
  });
});

// --- SelfAssignedDeclaration ---
describe('SelfAssignedDeclarationInspection', () => {
  it('detects As New declaration', () => {
    const code = 'Sub Test()\n    Dim obj As New Collection\nEnd Sub\n';
    const results = inspectCode(SelfAssignedDeclarationInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('SelfAssignedDeclaration');
  });

  it('does not flag normal declaration', () => {
    const code = 'Sub Test()\n    Dim obj As Collection\nEnd Sub\n';
    const results = inspectCode(SelfAssignedDeclarationInspection, code);
    expect(results).toHaveLength(0);
  });

  it('detects from fixture', async () => {
    const code = await readFile(join(fixturesDir, 'self-assigned.bas'), 'utf-8');
    const results = inspectCode(SelfAssignedDeclarationInspection, code);
    expect(results).toHaveLength(1);
  });
});

// --- UnreachableCode ---
describe('UnreachableCodeInspection', () => {
  it('detects code after Exit Sub', () => {
    const code = 'Sub Test()\n    Exit Sub\n    MsgBox "unreachable"\nEnd Sub\n';
    const results = inspectCode(UnreachableCodeInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('UnreachableCode');
  });

  it('does not flag code without exit', () => {
    const code = 'Sub Test()\n    MsgBox "reachable"\nEnd Sub\n';
    const results = inspectCode(UnreachableCodeInspection, code);
    expect(results).toHaveLength(0);
  });

  it('detects from fixture', async () => {
    const code = await readFile(join(fixturesDir, 'unreachable-code.bas'), 'utf-8');
    const results = inspectCode(UnreachableCodeInspection, code);
    expect(results).toHaveLength(1);
  });
});

// --- LineContinuationBetweenKeywords ---
describe('LineContinuationBetweenKeywordsInspection', () => {
  // This inspection works at token level and is tricky to test with the parser
  // since the parser may handle line continuations internally.
  it('has correct metadata', () => {
    expect(LineContinuationBetweenKeywordsInspection.meta.id).toBe('LineContinuationBetweenKeywords');
    expect(LineContinuationBetweenKeywordsInspection.meta.tier).toBe('A');
  });

  it('does not flag normal code', () => {
    const code = 'Sub Test()\n    MsgBox "Hi"\nEnd Sub\n';
    const results = inspectCode(LineContinuationBetweenKeywordsInspection, code);
    expect(results).toHaveLength(0);
  });
});

// --- OnLocalError ---
describe('OnLocalErrorInspection', () => {
  it('detects On Local Error', () => {
    const code = 'Sub Test()\n    On Local Error Resume Next\nEnd Sub\n';
    const results = inspectCode(OnLocalErrorInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('OnLocalError');
  });

  it('does not flag On Error', () => {
    const code = 'Sub Test()\n    On Error Resume Next\nEnd Sub\n';
    const results = inspectCode(OnLocalErrorInspection, code);
    expect(results).toHaveLength(0);
  });

  it('detects from fixture', async () => {
    const code = await readFile(join(fixturesDir, 'on-local-error.bas'), 'utf-8');
    const results = inspectCode(OnLocalErrorInspection, code);
    expect(results).toHaveLength(1);
  });
});

// --- StepNotSpecified ---
describe('StepNotSpecifiedInspection', () => {
  it('detects For loop without Step', () => {
    const code = 'Sub Test()\n    For i = 1 To 10\n        MsgBox i\n    Next i\nEnd Sub\n';
    const results = inspectCode(StepNotSpecifiedInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('StepNotSpecified');
  });

  it('does not flag For loop with Step', () => {
    const code = 'Sub Test()\n    For i = 1 To 10 Step 2\n        MsgBox i\n    Next i\nEnd Sub\n';
    const results = inspectCode(StepNotSpecifiedInspection, code);
    expect(results).toHaveLength(0);
  });

  it('detects from fixture', async () => {
    const code = await readFile(join(fixturesDir, 'step-not-specified.bas'), 'utf-8');
    const results = inspectCode(StepNotSpecifiedInspection, code);
    expect(results).toHaveLength(1);
  });
});

// --- StepOneIsRedundant ---
describe('StepOneIsRedundantInspection', () => {
  it('detects For loop with Step 1', () => {
    const code = 'Sub Test()\n    For i = 1 To 10 Step 1\n        MsgBox i\n    Next i\nEnd Sub\n';
    const results = inspectCode(StepOneIsRedundantInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('StepOneIsRedundant');
  });

  it('does not flag For loop with Step 2', () => {
    const code = 'Sub Test()\n    For i = 1 To 10 Step 2\n        MsgBox i\n    Next i\nEnd Sub\n';
    const results = inspectCode(StepOneIsRedundantInspection, code);
    expect(results).toHaveLength(0);
  });

  it('does not flag For loop without Step', () => {
    const code = 'Sub Test()\n    For i = 1 To 10\n        MsgBox i\n    Next i\nEnd Sub\n';
    const results = inspectCode(StepOneIsRedundantInspection, code);
    expect(results).toHaveLength(0);
  });

  it('detects from fixture', async () => {
    const code = await readFile(join(fixturesDir, 'step-one-redundant.bas'), 'utf-8');
    const results = inspectCode(StepOneIsRedundantInspection, code);
    expect(results).toHaveLength(1);
  });
});
