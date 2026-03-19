// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

/**
 * Inspection severity levels, matching Rubberduck's CodeInspectionSeverity.
 */
export type Severity = 'error' | 'warning' | 'suggestion' | 'hint';

/**
 * Inspection tier:
 * - A: Parse-tree only — works on any input (string or file)
 * - B: Symbol-aware — requires declaration finder / workspace context
 */
export type InspectionTier = 'A' | 'B';

/**
 * Inspection categories, based on Rubberduck's CodeInspectionType.
 */
export type InspectionCategory =
  | 'CodeQuality'
  | 'LanguageOpportunities'
  | 'Naming'
  | 'Performance'
  | 'Excel'
  | 'MsProject'
  | 'ErrorHandling'
  | 'ObsoleteSyntax';

/**
 * A single inspection result (diagnostic).
 */
export interface InspectionResult {
  /** Inspection ID (e.g., "EmptyIfBlock") */
  inspection: string;
  /** Human-readable description */
  description: string;
  /** Severity level */
  severity: Severity;
  /** Category */
  category: InspectionCategory;
  /** Tier (A=parse-tree, B=symbol-aware) */
  tier: InspectionTier;
  /** Source file path (for workspace mode) */
  source?: string;
  /** Location in source code */
  location: {
    startLine: number;
    startColumn: number;
    endLine: number;
    endColumn: number;
  };
  /** Optional quick fix suggestion */
  quickFix?: {
    description: string;
    edits?: Array<{
      location: {
        startLine: number;
        startColumn: number;
        endLine: number;
        endColumn: number;
      };
      newText: string;
    }>;
  };
  /** Whether this result was suppressed by @Ignore */
  suppressed: boolean;
}

/**
 * Error from an inspection that failed to run.
 */
export interface InspectionError {
  /** Inspection ID that failed */
  inspection: string;
  /** Error message */
  message: string;
}

/**
 * Info about an inspection that was skipped.
 */
export interface SkippedInfo {
  /** Inspection ID */
  inspection: string;
  /** Reason it was skipped */
  reason: string;
}

/**
 * Complete response from an inspection run.
 */
export interface InspectionResponse {
  results: InspectionResult[];
  errors: InspectionError[];
  skippedInspections: SkippedInfo[];
  parseErrors: Array<{ message: string; line: number; column: number }>;
  engineVersion: string;
}
