// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { readdir, stat, realpath } from 'node:fs/promises';
import path from 'node:path';
import { z } from 'zod';
import { parseCode } from '../parser/index.js';
import { parseCache } from '../parser/cache.js';
import { readVBAFile } from '../parser/file-reader.js';
import { createAllInspections } from '../inspections/registry.js';
import { runInspections } from '../inspections/runner.js';
import { WorkspaceContext } from '../symbols/workspace.js';
import { loadIgnorePatterns, filterIgnoredFiles } from './vbalintignore.js';
import { logger } from '../logger.js';
import type { InspectionResult, InspectionError, SkippedInfo, Severity } from '../inspections/types.js';

const VERSION = '0.1.0';

/** VBA file extensions to scan. */
const VBA_EXTENSIONS = new Set(['.bas', '.cls', '.frm']);

/** Maximum files per workspace scan (security control). */
const DEFAULT_MAX_FILES = 500;

export const inspectWorkspaceToolSchema = z.object({
  path: z.string().describe('Directory path to scan for VBA files'),
  hostLibraries: z.array(z.string()).default(['excel']).describe('Host libraries available'),
  severity: z.enum(['error', 'warning', 'suggestion', 'hint']).optional().describe('Minimum severity'),
  categories: z.array(z.string()).optional().describe('Filter by categories'),
  limit: z.number().default(100).describe('Maximum number of results to return'),
  detailed: z.boolean().default(false).describe('Return full results instead of summary'),
});

export type InspectWorkspaceInput = z.infer<typeof inspectWorkspaceToolSchema>;

interface WorkspaceResponse {
  results: InspectionResult[];
  errors: InspectionError[];
  skippedInspections: SkippedInfo[];
  parseErrors: Array<{ source: string; message: string; line: number; column: number }>;
  fileCount: number;
  engineVersion: string;
}

interface WorkspaceSummary {
  fileCount: number;
  totalResults: number;
  bySeverity: Record<string, number>;
  byFile: Array<{ file: string; count: number }>;
  topInspections: Array<{ inspection: string; count: number }>;
  engineVersion: string;
}

/**
 * Handle vba/inspect-workspace tool call.
 *
 * Scans a directory tree for VBA files, builds a workspace context
 * with cross-module symbol resolution, and runs all inspections.
 */
export async function handleInspectWorkspaceTool(input: InspectWorkspaceInput) {
  const startTime = Date.now();

  try {
    // Security: resolve and validate path
    const resolvedPath = path.resolve(input.path);

    // Verify exists and is a directory
    const pathStat = await stat(resolvedPath).catch(() => null);
    if (!pathStat || !pathStat.isDirectory()) {
      return {
        content: [{ type: 'text' as const, text: `Path is not a valid directory: ${resolvedPath}` }],
        isError: true,
      };
    }

    // Resolve symlinks and re-validate
    const realDir = await realpath(resolvedPath);
    const realStat = await stat(realDir);
    if (!realStat.isDirectory()) {
      return {
        content: [{ type: 'text' as const, text: `Resolved path is not a directory: ${realDir}` }],
        isError: true,
      };
    }

    // Discover VBA files
    const entries = await readdir(realDir, { recursive: true, withFileTypes: false });
    const allRelativePaths = (entries as string[])
      .filter(entry => {
        const ext = path.extname(entry).toLowerCase();
        return VBA_EXTENSIONS.has(ext);
      });

    // Apply .vbalintignore patterns
    const ignorePatterns = await loadIgnorePatterns(realDir);
    const filteredPaths = filterIgnoredFiles(allRelativePaths, ignorePatterns);

    // Enforce file count limit
    if (filteredPaths.length > DEFAULT_MAX_FILES) {
      return {
        content: [{
          type: 'text' as const,
          text: `Too many VBA files (${filteredPaths.length}). Maximum is ${DEFAULT_MAX_FILES}. Use .vbalintignore to exclude files.`,
        }],
        isError: true,
      };
    }

    if (filteredPaths.length === 0) {
      return {
        content: [{ type: 'text' as const, text: `No VBA files found in ${resolvedPath}` }],
        structuredContent: {
          results: [],
          errors: [],
          skippedInspections: [],
          parseErrors: [],
          fileCount: 0,
          engineVersion: VERSION,
        },
      };
    }

    // Phase 1: Read and parse all files, build workspace context
    const workspace = new WorkspaceContext();
    const fileParseErrors: WorkspaceResponse['parseErrors'] = [];
    const fileReadErrors: string[] = [];

    for (const relativePath of filteredPaths) {
      const absolutePath = path.join(realDir, relativePath);
      try {
        // Security: resolve symlinks and verify file is within workspace root
        const realFilePath = await realpath(absolutePath);
        if (!realFilePath.startsWith(realDir + path.sep) && realFilePath !== realDir) {
          logger.warn({ file: relativePath, resolved: realFilePath }, 'Skipping symlink outside workspace root');
          continue;
        }

        const content = await readVBAFile(absolutePath);
        const moduleName = path.basename(relativePath, path.extname(relativePath));

        // Check parse cache first
        let parseResult = parseCache.get(content);
        if (!parseResult) {
          parseResult = parseCode(content, { filePath: relativePath });
          parseCache.set(content, parseResult);
        } else {
          // Update source attribution from cache
          parseResult = { ...parseResult, source: relativePath };
        }

        // Collect parse errors with source attribution
        for (const err of parseResult.errors) {
          fileParseErrors.push({
            source: relativePath,
            message: err.message,
            line: err.line,
            column: err.column,
          });
        }

        workspace.addModule(moduleName, parseResult);
      } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        logger.warn({ file: relativePath, error: message }, 'Failed to read/parse VBA file');
        fileReadErrors.push(`${relativePath}: ${message}`);
      }
    }

    // Phase 2: Build cross-module declaration finder
    const declarationFinder = workspace.getDeclarationFinder();

    // Phase 3: Run inspections per module with workspace context
    const allResults: InspectionResult[] = [];
    const allErrors: InspectionError[] = [];
    const allSkipped: SkippedInfo[] = [];
    const seenSkipped = new Set<string>();

    const inspections = createAllInspections();

    // We need to get each module's parseResult to run inspections
    // Re-parse files (from cache) for inspection context
    for (const relativePath of filteredPaths) {
      const absolutePath = path.join(realDir, relativePath);
      try {
        // Security: resolve symlinks and verify file is within workspace root
        const realFilePath = await realpath(absolutePath);
        if (!realFilePath.startsWith(realDir + path.sep) && realFilePath !== realDir) {
          continue; // Already logged in Phase 1
        }

        const content = await readVBAFile(absolutePath);
        let parseResult = parseCache.get(content);
        if (!parseResult) {
          parseResult = parseCode(content, { filePath: relativePath });
          parseCache.set(content, parseResult);
        }

        const context = { parseResult, declarationFinder };
        const { results, errors, skipped } = runInspections(inspections, context, {
          minSeverity: input.severity as Severity | undefined,
          categories: input.categories,
          hostLibraries: input.hostLibraries,
          hasSymbolTable: true,
          source: relativePath,
        });

        allResults.push(...results);
        allErrors.push(...errors);

        // Only report each skipped inspection once
        for (const s of skipped) {
          if (!seenSkipped.has(s.inspection)) {
            seenSkipped.add(s.inspection);
            allSkipped.push(s);
          }
        }
      } catch {
        // Already logged in Phase 1
      }
    }

    // Apply limit
    const limitedResults = allResults.slice(0, input.limit);
    const totalCount = allResults.length;
    const elapsed = Date.now() - startTime;

    if (input.detailed) {
      // Detailed mode: full results array
      const response: WorkspaceResponse = {
        results: limitedResults,
        errors: allErrors,
        skippedInspections: allSkipped,
        parseErrors: fileParseErrors,
        fileCount: filteredPaths.length,
        engineVersion: VERSION,
      };

      const summary = buildTextSummary(limitedResults, totalCount, filteredPaths.length, elapsed, fileReadErrors);
      return {
        content: [{ type: 'text' as const, text: summary }],
        structuredContent: response,
      };
    } else {
      // Summary mode: counts by severity, by file, top inspections
      const summaryData = buildSummaryData(limitedResults, totalCount, filteredPaths.length);

      const summary = buildTextSummary(limitedResults, totalCount, filteredPaths.length, elapsed, fileReadErrors);
      return {
        content: [{ type: 'text' as const, text: summary }],
        structuredContent: summaryData,
      };
    }
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    return {
      content: [{ type: 'text' as const, text: `Workspace inspection error: ${message}` }],
      isError: true,
    };
  }
}

function buildTextSummary(
  results: InspectionResult[],
  totalCount: number,
  fileCount: number,
  elapsed: number,
  fileReadErrors: string[],
): string {
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

  let summary: string;
  if (parts.length > 0) {
    summary = `Found ${parts.join(', ')} across ${fileCount} file(s) (${elapsed}ms).`;
  } else {
    summary = `No issues found across ${fileCount} file(s) (${elapsed}ms).`;
  }

  if (results.length < totalCount) {
    summary += ` Showing ${results.length} of ${totalCount} total results.`;
  }

  const extras: string[] = [];
  if (fileReadErrors.length > 0) {
    extras.push(`${fileReadErrors.length} file(s) could not be read.`);
  }

  return [summary, ...extras].join(' ');
}

function buildSummaryData(
  results: InspectionResult[],
  totalCount: number,
  fileCount: number,
): WorkspaceSummary {
  // Counts by severity
  const bySeverity: Record<string, number> = {};
  for (const r of results) {
    bySeverity[r.severity] = (bySeverity[r.severity] ?? 0) + 1;
  }

  // Counts by file
  const fileMap = new Map<string, number>();
  for (const r of results) {
    const src = r.source ?? '<unknown>';
    fileMap.set(src, (fileMap.get(src) ?? 0) + 1);
  }
  const byFile = Array.from(fileMap.entries())
    .map(([file, count]) => ({ file, count }))
    .sort((a, b) => b.count - a.count);

  // Top inspections
  const inspMap = new Map<string, number>();
  for (const r of results) {
    inspMap.set(r.inspection, (inspMap.get(r.inspection) ?? 0) + 1);
  }
  const topInspections = Array.from(inspMap.entries())
    .map(([inspection, count]) => ({ inspection, count }))
    .sort((a, b) => b.count - a.count)
    .slice(0, 10);

  return {
    fileCount,
    totalResults: totalCount,
    bySeverity,
    byFile,
    topInspections,
    engineVersion: VERSION,
  };
}
