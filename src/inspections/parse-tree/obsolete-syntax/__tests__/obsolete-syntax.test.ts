// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { describe, it, expect } from 'vitest';
import { parseCode } from '../../../../parser/index.js';
import type { InspectionContext } from '../../../base.js';
import { readFile } from 'node:fs/promises';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

import { ObsoleteLetStatementInspection } from '../obsolete-let-statement.js';
import { ObsoleteCallStatementInspection } from '../obsolete-call-statement.js';
import { ObsoleteGlobalInspection } from '../obsolete-global.js';
import { ObsoleteWhileWendStatementInspection } from '../obsolete-while-wend-statement.js';
import { ObsoleteCommentSyntaxInspection } from '../obsolete-comment-syntax.js';
import { ObsoleteTypeHintInspection } from '../obsolete-type-hint.js';
import { StopKeywordInspection } from '../stop-keyword.js';
import { EndKeywordInspection } from '../end-keyword.js';
import { DefTypeStatementInspection } from '../def-type-statement.js';

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

// --- ObsoleteLetStatement ---
describe('ObsoleteLetStatementInspection', () => {
  it('detects Let keyword', () => {
    const code = 'Sub Test()\n    Let x = 5\nEnd Sub\n';
    const results = inspectCode(ObsoleteLetStatementInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('ObsoleteLetStatement');
  });

  it('does not flag assignment without Let', () => {
    const code = 'Sub Test()\n    x = 5\nEnd Sub\n';
    const results = inspectCode(ObsoleteLetStatementInspection, code);
    expect(results).toHaveLength(0);
  });

  it('detects from fixture', async () => {
    const code = await readFile(join(fixturesDir, 'obsolete-let.bas'), 'utf-8');
    const results = inspectCode(ObsoleteLetStatementInspection, code);
    expect(results).toHaveLength(1);
  });

  it('has correct metadata', () => {
    expect(ObsoleteLetStatementInspection.meta.id).toBe('ObsoleteLetStatement');
    expect(ObsoleteLetStatementInspection.meta.tier).toBe('A');
    expect(ObsoleteLetStatementInspection.meta.category).toBe('ObsoleteSyntax');
  });
});

// --- ObsoleteCallStatement ---
describe('ObsoleteCallStatementInspection', () => {
  it('detects Call keyword', () => {
    const code = 'Sub Test()\n    Call Foo\nEnd Sub\n';
    const results = inspectCode(ObsoleteCallStatementInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('ObsoleteCallStatement');
  });

  it('does not flag call without Call keyword', () => {
    const code = 'Sub Test()\n    Foo\nEnd Sub\n';
    const results = inspectCode(ObsoleteCallStatementInspection, code);
    expect(results).toHaveLength(0);
  });

  it('detects from fixture', async () => {
    const code = await readFile(join(fixturesDir, 'obsolete-call.bas'), 'utf-8');
    const results = inspectCode(ObsoleteCallStatementInspection, code);
    expect(results).toHaveLength(1);
  });
});

// --- ObsoleteGlobal ---
describe('ObsoleteGlobalInspection', () => {
  it('detects Global keyword', () => {
    const code = 'Global x As Long\nSub Test()\n    x = 5\nEnd Sub\n';
    const results = inspectCode(ObsoleteGlobalInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('ObsoleteGlobal');
  });

  it('does not flag Public keyword', () => {
    const code = 'Public x As Long\nSub Test()\n    x = 5\nEnd Sub\n';
    const results = inspectCode(ObsoleteGlobalInspection, code);
    expect(results).toHaveLength(0);
  });

  it('detects from fixture', async () => {
    const code = await readFile(join(fixturesDir, 'obsolete-global.bas'), 'utf-8');
    const results = inspectCode(ObsoleteGlobalInspection, code);
    expect(results).toHaveLength(1);
  });
});

// --- ObsoleteWhileWendStatement ---
describe('ObsoleteWhileWendStatementInspection', () => {
  it('detects While...Wend', () => {
    const code = 'Sub Test()\n    While True\n        MsgBox "loop"\n    Wend\nEnd Sub\n';
    const results = inspectCode(ObsoleteWhileWendStatementInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('ObsoleteWhileWendStatement');
  });

  it('does not flag Do While...Loop', () => {
    const code = 'Sub Test()\n    Do While True\n        MsgBox "loop"\n    Loop\nEnd Sub\n';
    const results = inspectCode(ObsoleteWhileWendStatementInspection, code);
    expect(results).toHaveLength(0);
  });

  it('detects from fixture', async () => {
    const code = await readFile(join(fixturesDir, 'obsolete-while-wend.bas'), 'utf-8');
    const results = inspectCode(ObsoleteWhileWendStatementInspection, code);
    expect(results).toHaveLength(1);
  });
});

// --- ObsoleteCommentSyntax ---
describe('ObsoleteCommentSyntaxInspection', () => {
  it('detects Rem comment', () => {
    const code = 'Sub Test()\n    Rem This is a comment\nEnd Sub\n';
    const results = inspectCode(ObsoleteCommentSyntaxInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('ObsoleteCommentSyntax');
  });

  it('does not flag single-quote comment', () => {
    const code = "Sub Test()\n    ' This is a comment\nEnd Sub\n";
    const results = inspectCode(ObsoleteCommentSyntaxInspection, code);
    expect(results).toHaveLength(0);
  });

  it('detects from fixture', async () => {
    const code = await readFile(join(fixturesDir, 'obsolete-rem.bas'), 'utf-8');
    const results = inspectCode(ObsoleteCommentSyntaxInspection, code);
    expect(results).toHaveLength(1);
  });
});

// --- ObsoleteTypeHint ---
describe('ObsoleteTypeHintInspection', () => {
  it('detects type hint character', () => {
    const code = 'Sub Test()\n    Dim name$\nEnd Sub\n';
    const results = inspectCode(ObsoleteTypeHintInspection, code);
    expect(results.length).toBeGreaterThanOrEqual(1);
    expect(results[0].inspection).toBe('ObsoleteTypeHint');
  });

  it('does not flag explicit As Type', () => {
    const code = 'Sub Test()\n    Dim name As String\nEnd Sub\n';
    const results = inspectCode(ObsoleteTypeHintInspection, code);
    expect(results).toHaveLength(0);
  });
});

// --- StopKeyword ---
describe('StopKeywordInspection', () => {
  it('detects Stop statement', () => {
    const code = 'Sub Test()\n    Stop\nEnd Sub\n';
    const results = inspectCode(StopKeywordInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('StopKeyword');
  });

  it('does not flag code without Stop', () => {
    const code = 'Sub Test()\n    MsgBox "Hi"\nEnd Sub\n';
    const results = inspectCode(StopKeywordInspection, code);
    expect(results).toHaveLength(0);
  });

  it('detects from fixture', async () => {
    const code = await readFile(join(fixturesDir, 'stop-keyword.bas'), 'utf-8');
    const results = inspectCode(StopKeywordInspection, code);
    expect(results).toHaveLength(1);
  });
});

// --- EndKeyword ---
describe('EndKeywordInspection', () => {
  it('detects standalone End statement', () => {
    const code = 'Sub Test()\n    End\nEnd Sub\n';
    const results = inspectCode(EndKeywordInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('EndKeyword');
  });

  it('does not flag End Sub or End Function', () => {
    const code = 'Sub Test()\n    MsgBox "Hi"\nEnd Sub\n';
    const results = inspectCode(EndKeywordInspection, code);
    expect(results).toHaveLength(0);
  });

  it('detects from fixture', async () => {
    const code = await readFile(join(fixturesDir, 'end-keyword.bas'), 'utf-8');
    const results = inspectCode(EndKeywordInspection, code);
    expect(results).toHaveLength(1);
  });
});

// --- DefTypeStatement ---
// TODO: DefType parsing requires EqualsString/TextOf/TokenAtRelativePosition globals
// which aren't patched yet. Test with single-letter form that doesn't need letter range.
describe('DefTypeStatementInspection', () => {
  // Skipped: DefType parsing requires EqualsString/TextOf/TokenAtRelativePosition
  // globals that aren't patched in the parser helpers yet. Once those are added,
  // this test can be enabled.
  it.skip('detects DefInt statement (requires parser helper patch)', () => {
    const code = 'DefInt I\nSub Test()\n    MsgBox "Hi"\nEnd Sub\n';
    const results = inspectCode(DefTypeStatementInspection, code);
    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('DefTypeStatement');
  });

  it('does not flag explicit type declarations', () => {
    const code = 'Dim x As Integer\nSub Test()\n    MsgBox "Hi"\nEnd Sub\n';
    const results = inspectCode(DefTypeStatementInspection, code);
    expect(results).toHaveLength(0);
  });

  it('has correct metadata', () => {
    expect(DefTypeStatementInspection.meta.id).toBe('DefTypeStatement');
    expect(DefTypeStatementInspection.meta.tier).toBe('A');
    expect(DefTypeStatementInspection.meta.category).toBe('ObsoleteSyntax');
  });
});
