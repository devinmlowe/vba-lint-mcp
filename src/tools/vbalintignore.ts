// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { readFile } from 'node:fs/promises';
import path from 'node:path';
import micromatch from 'micromatch';
// micromatch v4 is CJS — handle ESM interop
const mm = (micromatch as unknown as { default?: typeof micromatch }).default ?? micromatch;
import { logger } from '../logger.js';

/**
 * Parse a .vbalintignore file and return a filter function.
 *
 * Format is gitignore-compatible:
 * - One pattern per line
 * - `#` comments and blank lines are ignored
 * - Patterns are matched using micromatch (gitignore-style globs)
 * - Supports patterns like `*.frm`, `tests/`, `** /generated/**`
 */
export async function loadIgnorePatterns(workspaceRoot: string): Promise<string[]> {
  const ignorePath = path.join(workspaceRoot, '.vbalintignore');

  try {
    const content = await readFile(ignorePath, 'utf-8');
    const patterns = parseIgnoreFile(content);
    logger.info({ ignorePath, patternCount: patterns.length }, 'Loaded .vbalintignore');
    return patterns;
  } catch (err) {
    // File not found is expected — not an error
    if ((err as NodeJS.ErrnoException).code === 'ENOENT') {
      return [];
    }
    logger.warn({ ignorePath, error: (err as Error).message }, 'Failed to read .vbalintignore');
    return [];
  }
}

/**
 * Parse the content of a .vbalintignore file into an array of glob patterns.
 * Strips comments and blank lines.
 */
export function parseIgnoreFile(content: string): string[] {
  return content
    .split('\n')
    .map(line => line.trim())
    .filter(line => line.length > 0 && !line.startsWith('#'));
}

/**
 * Filter a list of relative file paths using .vbalintignore patterns.
 * Returns the files that are NOT ignored (i.e., should be inspected).
 */
export function filterIgnoredFiles(
  relativePaths: string[],
  ignorePatterns: string[],
): string[] {
  if (ignorePatterns.length === 0) {
    return relativePaths;
  }

  return mm.not(relativePaths, ignorePatterns, { dot: true });
}
