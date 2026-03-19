// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { describe, it, expect } from 'vitest';
import { parseCode } from '../../../../parser/index.js';
import { EmptyIfBlockInspection } from '../empty-if-block.js';
import type { InspectionContext } from '../../../base.js';
import { readFile } from 'node:fs/promises';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const fixturesDir = join(__dirname, '..', '__fixtures__');

function inspectCode(code: string) {
  const parseResult = parseCode(code);
  const context: InspectionContext = { parseResult };
  const inspection = new EmptyIfBlockInspection();
  return inspection.inspect(context);
}

describe('EmptyIfBlockInspection', () => {
  it('detects empty If block', () => {
    const code = 'Sub Test()\n    If True Then\n    End If\nEnd Sub\n';
    const results = inspectCode(code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('EmptyIfBlock');
    expect(results[0].severity).toBe('warning');
    expect(results[0].category).toBe('CodeQuality');
    expect(results[0].tier).toBe('A');
    expect(results[0].location.startLine).toBe(2);
  });

  it('does not flag If block with content', () => {
    const code = 'Sub Test()\n    If True Then\n        MsgBox "Hi"\n    End If\nEnd Sub\n';
    const results = inspectCode(code);
    expect(results).toHaveLength(0);
  });

  it('detects empty If from fixture file', async () => {
    const code = await readFile(join(fixturesDir, 'empty-if.bas'), 'utf-8');
    const results = inspectCode(code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('EmptyIfBlock');
  });

  it('no findings for non-empty fixture', async () => {
    const code = await readFile(join(fixturesDir, 'non-empty-if.bas'), 'utf-8');
    const results = inspectCode(code);
    expect(results).toHaveLength(0);
  });

  it('does not flag If with ElseIf body', () => {
    const code = [
      'Sub Test()',
      '    If x > 0 Then',
      '        MsgBox "positive"',
      '    ElseIf x = 0 Then',
      '        MsgBox "zero"',
      '    End If',
      'End Sub',
      '',
    ].join('\n');
    const results = inspectCode(code);
    expect(results).toHaveLength(0);
  });

  it('has correct metadata', () => {
    expect(EmptyIfBlockInspection.meta.id).toBe('EmptyIfBlock');
    expect(EmptyIfBlockInspection.meta.tier).toBe('A');
    expect(EmptyIfBlockInspection.meta.category).toBe('CodeQuality');
    expect(EmptyIfBlockInspection.meta.defaultSeverity).toBe('warning');
    expect(EmptyIfBlockInspection.meta.quickFixDescription).toBeTruthy();
  });
});
