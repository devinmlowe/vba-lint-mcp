// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { z } from 'zod';
import { getAllInspectionMetadata } from '../inspections/registry.js';

export const listInspectionsToolSchema = z.object({
  hostLibraries: z.array(z.string()).optional().describe('Filter by host libraries'),
  category: z.string().optional().describe('Filter by category'),
  tier: z.enum(['A', 'B']).optional().describe('Filter by tier (A=parse-tree, B=symbol-aware)'),
});

export type ListInspectionsInput = z.infer<typeof listInspectionsToolSchema>;

export function handleListInspections(input: ListInspectionsInput) {
  let metadata = getAllInspectionMetadata();

  // Apply filters
  if (input.tier) {
    metadata = metadata.filter(m => m.tier === input.tier);
  }

  if (input.category) {
    metadata = metadata.filter(m => m.category === input.category);
  }

  if (input.hostLibraries) {
    metadata = metadata.filter(m =>
      !m.hostLibraries || m.hostLibraries.some(h => input.hostLibraries!.includes(h)),
    );
  }

  const catalog = metadata.map(m => ({
    id: m.id,
    name: m.name,
    tier: m.tier,
    category: m.category,
    defaultSeverity: m.defaultSeverity,
    description: m.description,
    quickFixDescription: m.quickFixDescription,
    hostLibraries: m.hostLibraries,
  }));

  const summary = `${catalog.length} inspection(s) available.`;

  return {
    content: [{ type: 'text' as const, text: summary }],
    structuredContent: { inspections: catalog, count: catalog.length },
  };
}
