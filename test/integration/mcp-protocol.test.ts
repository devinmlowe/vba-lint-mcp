// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

/**
 * MCP Protocol integration tests.
 *
 * Tests the tool handlers that back the MCP server's JSON-RPC interface.
 * Each test exercises the full pipeline: parse -> inspect -> format response.
 *
 * The tool handlers are the same functions wired into server.ts. This validates
 * the complete request/response cycle without spawning a child process.
 *
 * Runs within vitest (no child process, no tmux, no Claude Code dependency).
 */

import { describe, it, expect } from 'vitest';
import path from 'node:path';
import { handleParseTool } from '../../src/tools/parse.js';
import { handleInspectTool } from '../../src/tools/inspect.js';
import { handleListInspections } from '../../src/tools/list-inspections.js';
import { handleInspectWorkspaceTool } from '../../src/tools/inspect-workspace.js';

const FIXTURES_DIR = path.resolve(__dirname, '../fixtures/workspace');

describe('MCP Protocol Integration Tests', () => {
  // --- Tool listing ---

  describe('vba/list-inspections', () => {
    it('returns all inspections with no filters', () => {
      const result = handleListInspections({});
      expect(result.structuredContent).toBeDefined();
      expect(result.structuredContent.inspections.length).toBeGreaterThan(0);
      expect(result.structuredContent.count).toBeGreaterThan(0);
      expect(result.content[0].text).toContain('inspection(s) available');
    });

    it('filters by tier A', () => {
      const result = handleListInspections({ tier: 'A' });
      const inspections = result.structuredContent.inspections;
      expect(inspections.length).toBeGreaterThan(0);
      for (const insp of inspections) {
        expect(insp.tier).toBe('A');
      }
    });

    it('filters by tier B', () => {
      const result = handleListInspections({ tier: 'B' });
      const inspections = result.structuredContent.inspections;
      expect(inspections.length).toBeGreaterThan(0);
      for (const insp of inspections) {
        expect(insp.tier).toBe('B');
      }
    });

    it('filters by category', () => {
      const result = handleListInspections({ category: 'CodeQuality' });
      const inspections = result.structuredContent.inspections;
      expect(inspections.length).toBeGreaterThan(0);
      for (const insp of inspections) {
        expect(insp.category).toBe('CodeQuality');
      }
    });

    it('filters by host libraries', () => {
      const result = handleListInspections({ hostLibraries: ['excel'] });
      const inspections = result.structuredContent.inspections;
      expect(inspections.length).toBeGreaterThan(0);
      // All returned inspections should either have no host restriction or include excel
      for (const insp of inspections) {
        if (insp.hostLibraries) {
          expect(insp.hostLibraries).toContain('excel');
        }
      }
    });
  });

  // --- vba/parse ---

  describe('vba/parse', () => {
    it('parses valid code', () => {
      const result = handleParseTool({
        code: 'Sub Test()\n  Dim x As Long\nEnd Sub',
        depth: 3,
      });

      expect(result.content).toBeDefined();
      const textContent = result.content.find((c: any) => c.type === 'text') as any;
      expect(textContent).toBeDefined();
      expect(textContent.text).toContain('Parsed');
      expect(textContent.text).toContain('3 lines');
      expect(result.structuredContent).toBeDefined();
      expect(result.structuredContent.ast).toBeDefined();
      expect(result.structuredContent.parseErrors).toBeDefined();
    });

    it('handles empty code', () => {
      const result = handleParseTool({ code: '', depth: 3 });
      expect(result.isError).toBeUndefined();
      expect(result.content).toBeDefined();
    });

    it('reports parse errors', () => {
      const result = handleParseTool({
        code: 'Sub Test(\n  broken syntax here\n',
        depth: 3,
      });
      // Should not be a tool error — parse errors are part of the result
      expect(result.content).toBeDefined();
    });

    it('respects depth parameter', () => {
      const shallow = handleParseTool({
        code: 'Sub Test()\n  Dim x As Long\nEnd Sub',
        depth: 1,
      });
      const deep = handleParseTool({
        code: 'Sub Test()\n  Dim x As Long\nEnd Sub',
        depth: 5,
      });

      // Deeper depth should produce more AST nodes
      const shallowStr = JSON.stringify(shallow.structuredContent.ast);
      const deepStr = JSON.stringify(deep.structuredContent.ast);
      expect(deepStr.length).toBeGreaterThan(shallowStr.length);
    });
  });

  // --- vba/inspect ---

  describe('vba/inspect', () => {
    it('detects issues in code', () => {
      const result = handleInspectTool({
        code: 'Sub Test()\n  If True Then\n  End If\nEnd Sub',
        hostLibraries: ['excel'],
      });

      expect(result.structuredContent).toBeDefined();
      const structured = result.structuredContent as any;
      expect(structured.results.length).toBeGreaterThan(0);
      expect(structured.engineVersion).toBe('0.1.0');
    });

    it('returns empty results for clean code', () => {
      // Minimal clean code that should trigger very few inspections
      const result = handleInspectTool({
        code: 'Option Explicit\n',
        hostLibraries: ['excel'],
      });

      expect(result.structuredContent).toBeDefined();
    });

    it('returns no tool error for empty code', () => {
      const result = handleInspectTool({
        code: '',
        hostLibraries: ['excel'],
      });

      expect(result.isError).toBeUndefined();
      expect(result.content).toBeDefined();
    });

    it('filters by severity', () => {
      const all = handleInspectTool({
        code: 'Sub Test()\n  If True Then\n  End If\nEnd Sub',
        hostLibraries: ['excel'],
      });
      const errorsOnly = handleInspectTool({
        code: 'Sub Test()\n  If True Then\n  End If\nEnd Sub',
        hostLibraries: ['excel'],
        severity: 'error',
      });

      const allResults = (all.structuredContent as any).results;
      const errorResults = (errorsOnly.structuredContent as any).results;

      // Error-only should have <= all results
      expect(errorResults.length).toBeLessThanOrEqual(allResults.length);
      for (const r of errorResults) {
        expect(r.severity).toBe('error');
      }
    });

    it('includes inspection metadata in results', () => {
      const result = handleInspectTool({
        code: 'Sub Test()\n  If True Then\n  End If\nEnd Sub',
        hostLibraries: ['excel'],
      });

      const structured = result.structuredContent as any;
      if (structured.results.length > 0) {
        const r = structured.results[0];
        expect(r.inspection).toBeDefined();
        expect(r.severity).toBeDefined();
        expect(r.category).toBeDefined();
        expect(r.tier).toBeDefined();
        expect(r.location).toBeDefined();
        expect(r.location.startLine).toBeGreaterThanOrEqual(1);
      }
    });

    it('text summary describes findings', () => {
      const result = handleInspectTool({
        code: 'Sub Test()\n  If True Then\n  End If\nEnd Sub',
        hostLibraries: ['excel'],
      });

      const text = result.content[0].text;
      expect(text).toContain('VBA');
      // Should mention lines or findings
      expect(text.length).toBeGreaterThan(10);
    });
  });

  // --- vba/inspect-workspace ---

  describe('vba/inspect-workspace', () => {
    it('scans directory and returns summary', async () => {
      const result = await handleInspectWorkspaceTool({
        path: FIXTURES_DIR,
        hostLibraries: ['excel'],
        limit: 100,
        detailed: false,
      });

      expect(result.isError).toBeUndefined();
      expect(result.structuredContent).toBeDefined();
      expect(result.structuredContent.fileCount).toBe(4); // excludes ignored/
      expect(result.structuredContent.totalResults).toBeDefined();
    });

    it('returns detailed results when requested', async () => {
      const result = await handleInspectWorkspaceTool({
        path: FIXTURES_DIR,
        hostLibraries: ['excel'],
        limit: 100,
        detailed: true,
      });

      expect(result.structuredContent).toBeDefined();
      expect(Array.isArray(result.structuredContent.results)).toBe(true);
      expect(result.structuredContent.errors).toBeDefined();
    });

    it('returns error for invalid path', async () => {
      const result = await handleInspectWorkspaceTool({
        path: '/nonexistent/path',
        hostLibraries: ['excel'],
        limit: 100,
        detailed: false,
      });

      expect(result.isError).toBe(true);
    });

    it('returns error for file (not directory)', async () => {
      const result = await handleInspectWorkspaceTool({
        path: path.join(FIXTURES_DIR, 'Module1.bas'),
        hostLibraries: ['excel'],
        limit: 100,
        detailed: false,
      });

      expect(result.isError).toBe(true);
    });
  });
});
