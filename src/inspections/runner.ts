// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import type { InspectionBase, InspectionContext } from './base.js';
import type { InspectionResult, InspectionError, SkippedInfo, Severity } from './types.js';
import { logger } from '../logger.js';

const SEVERITY_ORDER: Record<Severity, number> = {
  error: 0,
  warning: 1,
  suggestion: 2,
  hint: 3,
};

interface RunOptions {
  /** Minimum severity to include in results */
  minSeverity?: Severity;
  /** Filter by categories */
  categories?: string[];
  /** Filter by host libraries */
  hostLibraries?: string[];
  /** Whether symbol resolution is available */
  hasSymbolTable?: boolean;
  /** Source file path (for result attribution) */
  source?: string;
}

export interface RunResult {
  results: InspectionResult[];
  errors: InspectionError[];
  skipped: SkippedInfo[];
}

/**
 * Tiered inspection execution engine.
 *
 * Runs inspections with per-inspection isolation:
 * - Each inspection runs in a try/catch
 * - Tier B inspections are skipped if symbol table not available
 * - Results are filtered by severity, category, and host
 * - Suppressed results (via @Ignore) are marked but included
 */
export function runInspections(
  inspections: InspectionBase[],
  context: InspectionContext,
  options: RunOptions = {},
): RunResult {
  const results: InspectionResult[] = [];
  const errors: InspectionError[] = [];
  const skipped: SkippedInfo[] = [];

  const minSeverityLevel = options.minSeverity
    ? SEVERITY_ORDER[options.minSeverity]
    : SEVERITY_ORDER.hint; // Include all by default

  for (const inspection of inspections) {
    const meta = inspection.meta;

    // Host filtering: skip inspections for hosts not in the active set
    if (meta.hostLibraries && options.hostLibraries) {
      const hasMatchingHost = meta.hostLibraries.some(h =>
        options.hostLibraries!.includes(h),
      );
      if (!hasMatchingHost) continue;
    }

    // Category filtering
    if (options.categories && !options.categories.includes(meta.category)) {
      continue;
    }

    // Tier check: skip Tier B inspections when no symbol table available
    if (meta.tier === 'B' && !options.hasSymbolTable) {
      skipped.push({
        inspection: meta.id,
        reason: `Requires symbol resolution (Tier B); only parse-tree analysis available for inline code`,
      });
      continue;
    }

    // Execute with per-inspection isolation
    try {
      const startTime = Date.now();
      const inspectionResults = inspection.inspect(context);
      const elapsed = Date.now() - startTime;

      if (elapsed > 100) {
        logger.warn({ inspection: meta.id, elapsed }, 'Slow inspection');
      }

      // Apply severity filter and source attribution
      for (const result of inspectionResults) {
        if (options.source) {
          result.source = options.source;
        }

        const severityLevel = SEVERITY_ORDER[result.severity];
        if (severityLevel <= minSeverityLevel) {
          results.push(result);
        }
      }
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      logger.error({ inspection: meta.id, error: message }, 'Inspection failed');
      errors.push({
        inspection: meta.id,
        message,
      });
    }
  }

  return { results, errors, skipped };
}
