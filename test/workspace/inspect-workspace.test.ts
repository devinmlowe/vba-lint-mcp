// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { describe, it, expect, beforeEach } from 'vitest';
import path from 'node:path';
import { handleInspectWorkspaceTool } from '../../src/tools/inspect-workspace.js';
import { parseCache } from '../../src/parser/cache.js';

const FIXTURES_DIR = path.resolve(__dirname, '../fixtures/workspace');

describe('vba/inspect-workspace', () => {
  beforeEach(() => {
    parseCache.clear();
  });

  it('finds all VBA files in nested directories', async () => {
    const result = await handleInspectWorkspaceTool({
      path: FIXTURES_DIR,
      hostLibraries: ['excel'],
      limit: 100,
      detailed: true,
    });

    expect(result.isError).toBeUndefined();
    expect(result.structuredContent).toBeDefined();
    // .vbalintignore excludes ignored/ dir, leaving: Module1.bas, Class1.cls, subdir/Helper.bas, Form1.frm
    expect(result.structuredContent.fileCount).toBe(4);
  });

  it('respects .vbalintignore patterns', async () => {
    const result = await handleInspectWorkspaceTool({
      path: FIXTURES_DIR,
      hostLibraries: ['excel'],
      limit: 100,
      detailed: true,
    });

    // The ignored/ directory should be excluded
    const sources = result.structuredContent.results.map((r: { source: string }) => r.source);
    const ignoredFiles = sources.filter((s: string) => s.startsWith('ignored/'));
    expect(ignoredFiles.length).toBe(0);
  });

  it('returns summary mode by default', async () => {
    const result = await handleInspectWorkspaceTool({
      path: FIXTURES_DIR,
      hostLibraries: ['excel'],
      limit: 100,
      detailed: false,
    });

    expect(result.isError).toBeUndefined();
    const structured = result.structuredContent;
    expect(structured.fileCount).toBeDefined();
    expect(structured.totalResults).toBeDefined();
    expect(structured.bySeverity).toBeDefined();
    expect(structured.byFile).toBeDefined();
    expect(structured.topInspections).toBeDefined();
  });

  it('returns detailed mode when requested', async () => {
    const result = await handleInspectWorkspaceTool({
      path: FIXTURES_DIR,
      hostLibraries: ['excel'],
      limit: 100,
      detailed: true,
    });

    expect(result.isError).toBeUndefined();
    const structured = result.structuredContent;
    expect(structured.results).toBeDefined();
    expect(Array.isArray(structured.results)).toBe(true);
    expect(structured.errors).toBeDefined();
    expect(structured.skippedInspections).toBeDefined();
  });

  it('respects limit parameter', async () => {
    const result = await handleInspectWorkspaceTool({
      path: FIXTURES_DIR,
      hostLibraries: ['excel'],
      limit: 2,
      detailed: true,
    });

    expect(result.structuredContent.results.length).toBeLessThanOrEqual(2);
  });

  it('returns error for non-existent path', async () => {
    const result = await handleInspectWorkspaceTool({
      path: '/nonexistent/path/that/does/not/exist',
      hostLibraries: ['excel'],
      limit: 100,
      detailed: false,
    });

    expect(result.isError).toBe(true);
  });

  it('returns error for file path (not directory)', async () => {
    const result = await handleInspectWorkspaceTool({
      path: path.join(FIXTURES_DIR, 'Module1.bas'),
      hostLibraries: ['excel'],
      limit: 100,
      detailed: false,
    });

    expect(result.isError).toBe(true);
  });

  it('handles empty directory gracefully', async () => {
    const { mkdtemp } = await import('node:fs/promises');
    const { tmpdir } = await import('node:os');
    const emptyDir = await mkdtemp(path.join(tmpdir(), 'vba-lint-test-'));

    const result = await handleInspectWorkspaceTool({
      path: emptyDir,
      hostLibraries: ['excel'],
      limit: 100,
      detailed: false,
    });

    expect(result.isError).toBeUndefined();
    expect(result.structuredContent.fileCount).toBe(0);

    // Cleanup
    const { rm } = await import('node:fs/promises');
    await rm(emptyDir, { recursive: true });
  });

  it('populates parse cache during scan', async () => {
    expect(parseCache.size).toBe(0);

    await handleInspectWorkspaceTool({
      path: FIXTURES_DIR,
      hostLibraries: ['excel'],
      limit: 100,
      detailed: false,
    });

    // Cache should have entries for the scanned files
    expect(parseCache.size).toBeGreaterThan(0);
  });

  it('results include source file attribution', async () => {
    const result = await handleInspectWorkspaceTool({
      path: FIXTURES_DIR,
      hostLibraries: ['excel'],
      limit: 100,
      detailed: true,
    });

    // All results should have a source path
    for (const r of result.structuredContent.results) {
      expect(r.source).toBeDefined();
      expect(typeof r.source).toBe('string');
    }
  });

  it('text summary includes file count', async () => {
    const result = await handleInspectWorkspaceTool({
      path: FIXTURES_DIR,
      hostLibraries: ['excel'],
      limit: 100,
      detailed: false,
    });

    const text = result.content[0].text;
    expect(text).toContain('file(s)');
  });
});
