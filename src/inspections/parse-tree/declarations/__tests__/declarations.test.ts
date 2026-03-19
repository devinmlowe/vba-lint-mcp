// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { describe, it, expect } from 'vitest';
import { parseCode } from '../../../../parser/index.js';
import type { InspectionContext } from '../../../base.js';
import { readFile } from 'node:fs/promises';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

import { OptionExplicitInspection } from '../option-explicit.js';
import { OptionBaseZeroOrOneInspection } from '../option-base-zero-or-one.js';
import { MultipleDeclarationsInspection } from '../multiple-declarations.js';
import { ImplicitByRefModifierInspection } from '../implicit-byref-modifier.js';
import { ImplicitPublicMemberInspection } from '../implicit-public-member.js';
import { ImplicitVariantReturnTypeInspection } from '../implicit-variant-return-type.js';
import { RedundantByRefModifierInspection } from '../redundant-byref-modifier.js';

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

// --- OptionExplicit ---
describe('OptionExplicitInspection', () => {
  it('detects missing Option Explicit', () => {
    const code = 'Sub Test()\n    MsgBox "Hi"\nEnd Sub\n';
    const results = inspectCode(OptionExplicitInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('OptionExplicit');
  });

  it('does not flag when Option Explicit present', () => {
    const code = 'Option Explicit\nSub Test()\n    MsgBox "Hi"\nEnd Sub\n';
    const results = inspectCode(OptionExplicitInspection, code);
    expect(results).toHaveLength(0);
  });

  it('detects from fixture', async () => {
    const code = await readFile(join(fixturesDir, 'missing-option-explicit.bas'), 'utf-8');
    const results = inspectCode(OptionExplicitInspection, code);
    expect(results).toHaveLength(1);
  });

  it('no findings for fixture with Option Explicit', async () => {
    const code = await readFile(join(fixturesDir, 'has-option-explicit.bas'), 'utf-8');
    const results = inspectCode(OptionExplicitInspection, code);
    expect(results).toHaveLength(0);
  });
});

// --- OptionBaseZeroOrOne ---
describe('OptionBaseZeroOrOneInspection', () => {
  it('detects Option Base 1', () => {
    const code = 'Option Base 1\nSub Test()\nEnd Sub\n';
    const results = inspectCode(OptionBaseZeroOrOneInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('OptionBaseZeroOrOne');
  });

  it('does not flag modules without Option Base', () => {
    const code = 'Option Explicit\nSub Test()\nEnd Sub\n';
    const results = inspectCode(OptionBaseZeroOrOneInspection, code);
    expect(results).toHaveLength(0);
  });

  it('detects from fixture', async () => {
    const code = await readFile(join(fixturesDir, 'option-base.bas'), 'utf-8');
    const results = inspectCode(OptionBaseZeroOrOneInspection, code);
    expect(results).toHaveLength(1);
  });
});

// --- MultipleDeclarations ---
describe('MultipleDeclarationsInspection', () => {
  it('detects multiple declarations on one line', () => {
    const code = 'Sub Test()\n    Dim a, b, c\nEnd Sub\n';
    const results = inspectCode(MultipleDeclarationsInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('MultipleDeclarations');
  });

  it('does not flag single declaration', () => {
    const code = 'Sub Test()\n    Dim x As Long\nEnd Sub\n';
    const results = inspectCode(MultipleDeclarationsInspection, code);
    expect(results).toHaveLength(0);
  });

  it('detects from fixture', async () => {
    const code = await readFile(join(fixturesDir, 'multiple-declarations.bas'), 'utf-8');
    const results = inspectCode(MultipleDeclarationsInspection, code);
    expect(results).toHaveLength(1);
  });
});

// --- ImplicitByRefModifier ---
describe('ImplicitByRefModifierInspection', () => {
  it('detects implicit ByRef parameter', () => {
    const code = 'Sub Test(x As Long)\n    MsgBox x\nEnd Sub\n';
    const results = inspectCode(ImplicitByRefModifierInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('ImplicitByRefModifier');
  });

  it('does not flag explicit ByVal', () => {
    const code = 'Sub Test(ByVal x As Long)\n    MsgBox x\nEnd Sub\n';
    const results = inspectCode(ImplicitByRefModifierInspection, code);
    expect(results).toHaveLength(0);
  });

  it('does not flag explicit ByRef', () => {
    const code = 'Sub Test(ByRef x As Long)\n    MsgBox x\nEnd Sub\n';
    const results = inspectCode(ImplicitByRefModifierInspection, code);
    expect(results).toHaveLength(0);
  });

  it('detects from fixture', async () => {
    const code = await readFile(join(fixturesDir, 'implicit-byref.bas'), 'utf-8');
    const results = inspectCode(ImplicitByRefModifierInspection, code);
    expect(results).toHaveLength(1);
  });
});

// --- ImplicitPublicMember ---
describe('ImplicitPublicMemberInspection', () => {
  it('detects implicit Public Sub', () => {
    const code = 'Sub Test()\n    MsgBox "Hi"\nEnd Sub\n';
    const results = inspectCode(ImplicitPublicMemberInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('ImplicitPublicMember');
  });

  it('does not flag explicit Public', () => {
    const code = 'Public Sub Test()\n    MsgBox "Hi"\nEnd Sub\n';
    const results = inspectCode(ImplicitPublicMemberInspection, code);
    expect(results).toHaveLength(0);
  });

  it('does not flag explicit Private', () => {
    const code = 'Private Sub Test()\n    MsgBox "Hi"\nEnd Sub\n';
    const results = inspectCode(ImplicitPublicMemberInspection, code);
    expect(results).toHaveLength(0);
  });

  it('detects from fixture', async () => {
    const code = await readFile(join(fixturesDir, 'implicit-public.bas'), 'utf-8');
    const results = inspectCode(ImplicitPublicMemberInspection, code);
    expect(results).toHaveLength(1);
  });
});

// --- ImplicitVariantReturnType ---
describe('ImplicitVariantReturnTypeInspection', () => {
  it('detects Function without return type', () => {
    const code = 'Function GetValue()\n    GetValue = 42\nEnd Function\n';
    const results = inspectCode(ImplicitVariantReturnTypeInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('ImplicitVariantReturnType');
  });

  it('does not flag Function with return type', () => {
    const code = 'Function GetValue() As Long\n    GetValue = 42\nEnd Function\n';
    const results = inspectCode(ImplicitVariantReturnTypeInspection, code);
    expect(results).toHaveLength(0);
  });

  it('detects from fixture', async () => {
    const code = await readFile(join(fixturesDir, 'implicit-variant-return.bas'), 'utf-8');
    const results = inspectCode(ImplicitVariantReturnTypeInspection, code);
    expect(results).toHaveLength(1);
  });
});

// --- RedundantByRefModifier ---
describe('RedundantByRefModifierInspection', () => {
  it('detects explicit ByRef', () => {
    const code = 'Sub Test(ByRef x As Long)\n    MsgBox x\nEnd Sub\n';
    const results = inspectCode(RedundantByRefModifierInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('RedundantByRefModifier');
  });

  it('does not flag ByVal', () => {
    const code = 'Sub Test(ByVal x As Long)\n    MsgBox x\nEnd Sub\n';
    const results = inspectCode(RedundantByRefModifierInspection, code);
    expect(results).toHaveLength(0);
  });

  it('does not flag implicit ByRef', () => {
    const code = 'Sub Test(x As Long)\n    MsgBox x\nEnd Sub\n';
    const results = inspectCode(RedundantByRefModifierInspection, code);
    expect(results).toHaveLength(0);
  });

  it('detects from fixture', async () => {
    const code = await readFile(join(fixturesDir, 'redundant-byref.bas'), 'utf-8');
    const results = inspectCode(RedundantByRefModifierInspection, code);
    expect(results).toHaveLength(1);
  });
});
