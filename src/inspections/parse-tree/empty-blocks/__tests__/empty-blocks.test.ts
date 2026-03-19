// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { describe, it, expect } from 'vitest';
import { parseCode } from '../../../../parser/index.js';
import type { InspectionContext } from '../../../base.js';
import { readFile } from 'node:fs/promises';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

import { EmptyElseBlockInspection } from '../empty-else-block.js';
import { EmptyCaseBlockInspection } from '../empty-case-block.js';
import { EmptyForLoopBlockInspection } from '../empty-for-loop-block.js';
import { EmptyForEachBlockInspection } from '../empty-for-each-block.js';
import { EmptyWhileWendBlockInspection } from '../empty-while-wend-block.js';
import { EmptyDoWhileBlockInspection } from '../empty-do-while-block.js';
import { EmptyMethodInspection } from '../empty-method.js';
import { EmptyModuleInspection } from '../empty-module.js';

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

// --- EmptyElseBlock ---
describe('EmptyElseBlockInspection', () => {
  it('detects empty Else block', () => {
    const code = 'Sub Test()\n    If True Then\n        MsgBox "Hi"\n    Else\n    End If\nEnd Sub\n';
    const results = inspectCode(EmptyElseBlockInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('EmptyElseBlock');
  });

  it('does not flag Else block with content', () => {
    const code = 'Sub Test()\n    If True Then\n        MsgBox "Hi"\n    Else\n        MsgBox "Bye"\n    End If\nEnd Sub\n';
    const results = inspectCode(EmptyElseBlockInspection, code);
    expect(results).toHaveLength(0);
  });

  it('detects from fixture', async () => {
    const code = await readFile(join(fixturesDir, 'empty-else.bas'), 'utf-8');
    const results = inspectCode(EmptyElseBlockInspection, code);
    expect(results).toHaveLength(1);
  });

  it('no findings for non-empty fixture', async () => {
    const code = await readFile(join(fixturesDir, 'non-empty-else.bas'), 'utf-8');
    const results = inspectCode(EmptyElseBlockInspection, code);
    expect(results).toHaveLength(0);
  });

  it('has correct metadata', () => {
    expect(EmptyElseBlockInspection.meta.id).toBe('EmptyElseBlock');
    expect(EmptyElseBlockInspection.meta.tier).toBe('A');
  });
});

// --- EmptyCaseBlock ---
describe('EmptyCaseBlockInspection', () => {
  it('detects empty Case block', () => {
    const code = [
      'Sub Test()',
      '    Select Case x',
      '        Case 1',
      '        Case 2',
      '            MsgBox "two"',
      '    End Select',
      'End Sub',
      '',
    ].join('\n');
    const results = inspectCode(EmptyCaseBlockInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('EmptyCaseBlock');
  });

  it('does not flag Case block with content', () => {
    const code = [
      'Sub Test()',
      '    Select Case x',
      '        Case 1',
      '            MsgBox "one"',
      '        Case 2',
      '            MsgBox "two"',
      '    End Select',
      'End Sub',
      '',
    ].join('\n');
    const results = inspectCode(EmptyCaseBlockInspection, code);
    expect(results).toHaveLength(0);
  });

  it('detects from fixture', async () => {
    const code = await readFile(join(fixturesDir, 'empty-case.bas'), 'utf-8');
    const results = inspectCode(EmptyCaseBlockInspection, code);
    expect(results.length).toBeGreaterThanOrEqual(1);
  });
});

// --- EmptyForLoopBlock ---
describe('EmptyForLoopBlockInspection', () => {
  it('detects empty For loop', () => {
    const code = 'Sub Test()\n    For i = 1 To 10\n    Next i\nEnd Sub\n';
    const results = inspectCode(EmptyForLoopBlockInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('EmptyForLoopBlock');
  });

  it('does not flag For loop with content', () => {
    const code = 'Sub Test()\n    For i = 1 To 10\n        MsgBox i\n    Next i\nEnd Sub\n';
    const results = inspectCode(EmptyForLoopBlockInspection, code);
    expect(results).toHaveLength(0);
  });

  it('detects from fixture', async () => {
    const code = await readFile(join(fixturesDir, 'empty-for-loop.bas'), 'utf-8');
    const results = inspectCode(EmptyForLoopBlockInspection, code);
    expect(results).toHaveLength(1);
  });
});

// --- EmptyForEachBlock ---
describe('EmptyForEachBlockInspection', () => {
  it('detects empty For Each loop', () => {
    const code = 'Sub Test()\n    For Each item In collection\n    Next item\nEnd Sub\n';
    const results = inspectCode(EmptyForEachBlockInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('EmptyForEachBlock');
  });

  it('does not flag For Each loop with content', () => {
    const code = 'Sub Test()\n    For Each item In collection\n        MsgBox item\n    Next item\nEnd Sub\n';
    const results = inspectCode(EmptyForEachBlockInspection, code);
    expect(results).toHaveLength(0);
  });

  it('detects from fixture', async () => {
    const code = await readFile(join(fixturesDir, 'empty-for-each.bas'), 'utf-8');
    const results = inspectCode(EmptyForEachBlockInspection, code);
    expect(results).toHaveLength(1);
  });
});

// --- EmptyWhileWendBlock ---
describe('EmptyWhileWendBlockInspection', () => {
  it('detects empty While...Wend loop', () => {
    const code = 'Sub Test()\n    While True\n    Wend\nEnd Sub\n';
    const results = inspectCode(EmptyWhileWendBlockInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('EmptyWhileWendBlock');
  });

  it('does not flag While...Wend loop with content', () => {
    const code = 'Sub Test()\n    While True\n        MsgBox "loop"\n    Wend\nEnd Sub\n';
    const results = inspectCode(EmptyWhileWendBlockInspection, code);
    expect(results).toHaveLength(0);
  });

  it('detects from fixture', async () => {
    const code = await readFile(join(fixturesDir, 'empty-while-wend.bas'), 'utf-8');
    const results = inspectCode(EmptyWhileWendBlockInspection, code);
    expect(results).toHaveLength(1);
  });
});

// --- EmptyDoWhileBlock ---
describe('EmptyDoWhileBlockInspection', () => {
  it('detects empty Do...Loop block', () => {
    const code = 'Sub Test()\n    Do While True\n    Loop\nEnd Sub\n';
    const results = inspectCode(EmptyDoWhileBlockInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('EmptyDoWhileBlock');
  });

  it('does not flag Do...Loop with content', () => {
    const code = 'Sub Test()\n    Do While True\n        MsgBox "loop"\n    Loop\nEnd Sub\n';
    const results = inspectCode(EmptyDoWhileBlockInspection, code);
    expect(results).toHaveLength(0);
  });

  it('detects from fixture', async () => {
    const code = await readFile(join(fixturesDir, 'empty-do-while.bas'), 'utf-8');
    const results = inspectCode(EmptyDoWhileBlockInspection, code);
    expect(results).toHaveLength(1);
  });
});

// --- EmptyMethod ---
describe('EmptyMethodInspection', () => {
  it('detects empty Sub', () => {
    const code = 'Sub EmptySub()\nEnd Sub\n';
    const results = inspectCode(EmptyMethodInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('EmptyMethod');
  });

  it('detects empty Function', () => {
    const code = 'Function EmptyFunc() As Long\nEnd Function\n';
    const results = inspectCode(EmptyMethodInspection, code);
    expect(results).toHaveLength(1);
  });

  it('does not flag Sub with content', () => {
    const code = 'Sub Test()\n    MsgBox "Hi"\nEnd Sub\n';
    const results = inspectCode(EmptyMethodInspection, code);
    expect(results).toHaveLength(0);
  });

  it('detects from fixture', async () => {
    const code = await readFile(join(fixturesDir, 'empty-method.bas'), 'utf-8');
    const results = inspectCode(EmptyMethodInspection, code);
    expect(results).toHaveLength(1);
  });
});

// --- EmptyModule ---
describe('EmptyModuleInspection', () => {
  it('detects empty module', () => {
    const code = 'Option Explicit\n';
    const results = inspectCode(EmptyModuleInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('EmptyModule');
  });

  it('does not flag module with procedure', () => {
    const code = 'Option Explicit\nSub Test()\n    MsgBox "Hi"\nEnd Sub\n';
    const results = inspectCode(EmptyModuleInspection, code);
    expect(results).toHaveLength(0);
  });

  it('does not flag module with declarations', () => {
    const code = 'Option Explicit\nDim x As Long\n';
    const results = inspectCode(EmptyModuleInspection, code);
    expect(results).toHaveLength(0);
  });

  it('detects from fixture', async () => {
    const code = await readFile(join(fixturesDir, 'empty-module.bas'), 'utf-8');
    const results = inspectCode(EmptyModuleInspection, code);
    expect(results).toHaveLength(1);
  });
});
