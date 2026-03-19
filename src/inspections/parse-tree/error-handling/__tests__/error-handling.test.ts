// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { describe, it, expect } from 'vitest';
import { parseCode } from '../../../../parser/index.js';
import type { InspectionContext } from '../../../base.js';
import { readFile } from 'node:fs/promises';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

import { UnhandledOnErrorResumeNextInspection } from '../unhandled-on-error-resume-next.js';
import { OnErrorGoToMinusOneInspection } from '../on-error-goto-minus-one.js';
import { EmptyStringLiteralInspection } from '../empty-string-literal.js';
import { IsMissingOnInappropriateArgumentInspection } from '../is-missing-on-inappropriate-argument.js';
import { IsMissingWithNonArgumentParameterInspection } from '../is-missing-with-non-argument-parameter.js';

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

// --- UnhandledOnErrorResumeNext ---
describe('UnhandledOnErrorResumeNextInspection', () => {
  it('detects unhandled On Error Resume Next', () => {
    const code = 'Sub Test()\n    On Error Resume Next\n    MsgBox "Hi"\nEnd Sub\n';
    const results = inspectCode(UnhandledOnErrorResumeNextInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('UnhandledOnErrorResumeNext');
  });

  it('does not flag when On Error GoTo 0 is present', () => {
    const code = 'Sub Test()\n    On Error Resume Next\n    MsgBox "Hi"\n    On Error GoTo 0\nEnd Sub\n';
    const results = inspectCode(UnhandledOnErrorResumeNextInspection, code);
    expect(results).toHaveLength(0);
  });

  it('detects from fixture', async () => {
    const code = await readFile(join(fixturesDir, 'unhandled-resume-next.bas'), 'utf-8');
    const results = inspectCode(UnhandledOnErrorResumeNextInspection, code);
    expect(results).toHaveLength(1);
  });

  it('no findings for handled fixture', async () => {
    const code = await readFile(join(fixturesDir, 'handled-resume-next.bas'), 'utf-8');
    const results = inspectCode(UnhandledOnErrorResumeNextInspection, code);
    expect(results).toHaveLength(0);
  });
});

// --- OnErrorGoToMinusOne ---
describe('OnErrorGoToMinusOneInspection', () => {
  it('detects On Error GoTo -1', () => {
    const code = 'Sub Test()\n    On Error GoTo -1\nEnd Sub\n';
    const results = inspectCode(OnErrorGoToMinusOneInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('OnErrorGoToMinusOne');
  });

  it('does not flag On Error GoTo 0', () => {
    const code = 'Sub Test()\n    On Error GoTo 0\nEnd Sub\n';
    const results = inspectCode(OnErrorGoToMinusOneInspection, code);
    expect(results).toHaveLength(0);
  });

  it('detects from fixture', async () => {
    const code = await readFile(join(fixturesDir, 'on-error-goto-minus-one.bas'), 'utf-8');
    const results = inspectCode(OnErrorGoToMinusOneInspection, code);
    expect(results).toHaveLength(1);
  });
});

// --- EmptyStringLiteral ---
describe('EmptyStringLiteralInspection', () => {
  it('detects empty string literal', () => {
    const code = 'Sub Test()\n    If str = "" Then\n        MsgBox "empty"\n    End If\nEnd Sub\n';
    const results = inspectCode(EmptyStringLiteralInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('EmptyStringLiteral');
  });

  it('does not flag non-empty string literal', () => {
    const code = 'Sub Test()\n    If str = "hello" Then\n        MsgBox "hi"\n    End If\nEnd Sub\n';
    const results = inspectCode(EmptyStringLiteralInspection, code);
    expect(results).toHaveLength(0);
  });

  it('detects from fixture', async () => {
    const code = await readFile(join(fixturesDir, 'empty-string-literal.bas'), 'utf-8');
    const results = inspectCode(EmptyStringLiteralInspection, code);
    expect(results).toHaveLength(1);
  });
});

// --- IsMissingOnInappropriateArgument ---
describe('IsMissingOnInappropriateArgumentInspection', () => {
  it('detects IsMissing call', () => {
    const code = 'Sub Test(Optional ByVal x As Variant)\n    If IsMissing(x) Then\n        MsgBox "missing"\n    End If\nEnd Sub\n';
    const results = inspectCode(IsMissingOnInappropriateArgumentInspection, code);
    expect(results.length).toBeGreaterThanOrEqual(1);
    expect(results[0].inspection).toBe('IsMissingOnInappropriateArgument');
  });

  it('does not flag code without IsMissing', () => {
    const code = 'Sub Test()\n    MsgBox "Hi"\nEnd Sub\n';
    const results = inspectCode(IsMissingOnInappropriateArgumentInspection, code);
    expect(results).toHaveLength(0);
  });
});

// --- IsMissingWithNonArgumentParameter ---
describe('IsMissingWithNonArgumentParameterInspection', () => {
  it('has correct metadata', () => {
    expect(IsMissingWithNonArgumentParameterInspection.meta.id).toBe('IsMissingWithNonArgumentParameter');
    expect(IsMissingWithNonArgumentParameterInspection.meta.tier).toBe('A');
  });

  it('returns empty results (requires Tier B)', () => {
    const code = 'Sub Test()\n    Dim x As Variant\n    If IsMissing(x) Then\n    End If\nEnd Sub\n';
    const results = inspectCode(IsMissingWithNonArgumentParameterInspection, code);
    expect(results).toHaveLength(0);
  });
});
