// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Abstract/InspectionBase.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import type { ParseResult } from '../parser/index.js';
import type { InspectionResult, Severity, InspectionCategory, InspectionTier } from './types.js';
import type { DeclarationFinder } from '../symbols/declaration-finder.js';

/**
 * Context provided to inspections during execution.
 */
export interface InspectionContext {
  /** Parse result with tree and tokens */
  parseResult: ParseResult;
  /** Declaration finder — only available for Tier B inspections */
  declarationFinder?: DeclarationFinder;
}

/**
 * Static metadata for an inspection class.
 * Every concrete inspection must provide these as static properties.
 */
export interface InspectionMetadata {
  /** Unique inspection ID (e.g., "EmptyIfBlock") */
  id: string;
  /** Inspection tier: A (parse-tree) or B (symbol-aware) */
  tier: InspectionTier;
  /** Category */
  category: InspectionCategory;
  /** Default severity */
  defaultSeverity: Severity;
  /** Host libraries this inspection applies to (undefined = all) */
  hostLibraries?: string[];
  /** Human-readable name */
  name: string;
  /** Detailed description / rationale */
  description: string;
  /** Quick fix description (if any) */
  quickFixDescription?: string;
}

/**
 * Abstract base class for all inspections.
 *
 * Follows Rubberduck's Template Method pattern:
 * - Subclasses implement `inspect()` with detection logic
 * - Base class provides metadata access and result construction helpers
 */
export abstract class InspectionBase {
  /** Static metadata — must be overridden by subclasses */
  static readonly meta: InspectionMetadata;

  /** Get metadata from the concrete class */
  get meta(): InspectionMetadata {
    return (this.constructor as typeof InspectionBase).meta;
  }

  /**
   * Run this inspection against the given context.
   * Returns an array of results (may be empty if no issues found).
   */
  abstract inspect(context: InspectionContext): InspectionResult[];

  /**
   * Helper: create an InspectionResult with this inspection's metadata.
   */
  protected createResult(
    location: InspectionResult['location'],
    options?: {
      description?: string;
      severity?: Severity;
      quickFix?: InspectionResult['quickFix'];
      source?: string;
    },
  ): InspectionResult {
    return {
      inspection: this.meta.id,
      description: options?.description ?? this.meta.description,
      severity: options?.severity ?? this.meta.defaultSeverity,
      category: this.meta.category,
      tier: this.meta.tier,
      source: options?.source,
      location,
      quickFix: options?.quickFix,
      suppressed: false,
    };
  }
}

/**
 * Base class for Tier A inspections (parse-tree only).
 * These inspections only need the parse tree — no symbol resolution.
 */
export abstract class ParseTreeInspection extends InspectionBase {
  // Tier A inspections only use context.parseResult
}

/**
 * Base class for Tier B inspections (declaration-based).
 * These require the declaration finder / symbol table.
 */
export abstract class DeclarationInspection extends InspectionBase {
  // Tier B inspections require context.declarationFinder
}

/**
 * Base class for Tier B inspections (reference-based).
 * These require reference resolution + potentially workspace context.
 */
export abstract class ReferenceInspection extends InspectionBase {
  // Tier B inspections require context.declarationFinder + workspace
}
