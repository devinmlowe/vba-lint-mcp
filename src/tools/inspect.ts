// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { z } from 'zod';
import { parseCode } from '../parser/index.js';
import { createAllInspections } from '../inspections/registry.js';
import { runInspections } from '../inspections/runner.js';
import type { InspectionResponse, Severity } from '../inspections/types.js';
import { buildSingleModuleFinder } from '../symbols/workspace.js';

const VERSION = '0.1.0';

export const inspectToolSchema = z.object({
  code: z.string().describe('VBA source code to inspect'),
  hostLibraries: z.array(z.string()).default(['excel']).describe('Host libraries available (e.g., ["excel", "access"])'),
  severity: z.enum(['error', 'warning', 'suggestion', 'hint']).optional().describe('Minimum severity to include'),
  categories: z.array(z.string()).optional().describe('Filter by inspection categories'),
});

export type InspectToolInput = z.infer<typeof inspectToolSchema>;

export function handleInspectTool(input: InspectToolInput) {
  const startTime = Date.now();

  try {
    const parseResult = parseCode(input.code);
    const inspections = createAllInspections();

    // Build symbol table for Tier B inspections
    const declarationFinder = buildSingleModuleFinder(parseResult);
    const context = { parseResult, declarationFinder };

    const { results, errors, skipped } = runInspections(inspections, context, {
      minSeverity: input.severity as Severity | undefined,
      categories: input.categories,
      hostLibraries: input.hostLibraries,
      hasSymbolTable: true,
    });

    const elapsed = Date.now() - startTime;

    const response: InspectionResponse = {
      results,
      errors,
      skippedInspections: skipped,
      parseErrors: parseResult.errors,
      engineVersion: VERSION,
    };

    // Human-readable summary
    const counts = {
      error: results.filter(r => r.severity === 'error').length,
      warning: results.filter(r => r.severity === 'warning').length,
      suggestion: results.filter(r => r.severity === 'suggestion').length,
      hint: results.filter(r => r.severity === 'hint').length,
    };
    const parts: string[] = [];
    if (counts.error) parts.push(`${counts.error} error(s)`);
    if (counts.warning) parts.push(`${counts.warning} warning(s)`);
    if (counts.suggestion) parts.push(`${counts.suggestion} suggestion(s)`);
    if (counts.hint) parts.push(`${counts.hint} hint(s)`);

    const summary = parts.length > 0
      ? `Found ${parts.join(', ')} in ${input.code.split('\n').length} lines of VBA (${elapsed}ms).`
      : `No issues found in ${input.code.split('\n').length} lines of VBA (${elapsed}ms).`;

    const extra: string[] = [];
    if (parseResult.errors.length > 0) {
      extra.push(`${parseResult.errors.length} parse error(s).`);
    }
    if (skipped.length > 0) {
      extra.push(`${skipped.length} inspection(s) skipped (Tier B requires symbol resolution).`);
    }
    if (errors.length > 0) {
      extra.push(`${errors.length} inspection(s) failed to run.`);
    }

    const textSummary = [summary, ...extra].join(' ');

    return {
      content: [{ type: 'text' as const, text: textSummary }],
      structuredContent: response,
    };
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    return {
      content: [{ type: 'text' as const, text: `Inspection error: ${message}` }],
      isError: true,
    };
  }
}
