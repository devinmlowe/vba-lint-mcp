// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { InspectionBase } from './base.js';
import type { InspectionMetadata } from './base.js';

// =================================================================
// Explicit barrel file registration.
// Every inspection class must be imported and added to ALL_INSPECTIONS.
// This provides compile-time safety — missing imports cause build errors.
// =================================================================

// --- Phase 2: Tier A (Parse-Tree) Inspections ---
import { EmptyIfBlockInspection } from './parse-tree/empty-blocks/empty-if-block.js';

/**
 * Master registry of all inspection classes.
 * Order does not matter — the runner handles tiering and filtering.
 */
export const ALL_INSPECTIONS: Array<new () => InspectionBase> = [
  // Tier A: Parse-Tree Inspections
  EmptyIfBlockInspection,
  // More Tier A inspections added in Phase 2
  // Tier B inspections added in Phase 4
];

/**
 * Create instances of all registered inspections.
 */
export function createAllInspections(): InspectionBase[] {
  return ALL_INSPECTIONS.map(Cls => new Cls());
}

/**
 * Get metadata for all registered inspections without instantiating.
 */
export function getAllInspectionMetadata(): InspectionMetadata[] {
  return ALL_INSPECTIONS.map(Cls => (Cls as unknown as typeof InspectionBase).meta);
}

/**
 * Validate the registry at startup.
 * Returns errors for any issues found.
 */
export function validateRegistry(): string[] {
  const errors: string[] = [];
  const ids = new Set<string>();

  for (const Cls of ALL_INSPECTIONS) {
    const meta = (Cls as unknown as typeof InspectionBase).meta;

    if (!meta) {
      errors.push(`Inspection class ${Cls.name} is missing static 'meta' property`);
      continue;
    }

    if (!meta.id) {
      errors.push(`Inspection class ${Cls.name} has empty 'id' in meta`);
      continue;
    }

    if (ids.has(meta.id)) {
      errors.push(`Duplicate inspection ID: ${meta.id}`);
    }
    ids.add(meta.id);

    if (!meta.tier || !['A', 'B'].includes(meta.tier)) {
      errors.push(`Inspection ${meta.id} has invalid tier: ${meta.tier}`);
    }

    if (!meta.category) {
      errors.push(`Inspection ${meta.id} has no category`);
    }

    if (!meta.defaultSeverity) {
      errors.push(`Inspection ${meta.id} has no defaultSeverity`);
    }
  }

  return errors;
}
