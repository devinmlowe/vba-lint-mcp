// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

/**
 * VBA module files (.cls, .frm) contain header lines that precede the actual
 * VBA code. These headers include VERSION, BEGIN/END blocks, and Attribute lines.
 *
 * Example .cls header:
 *   VERSION 1.0 CLASS
 *   BEGIN
 *     MultiUse = -1  'True
 *   END
 *   Attribute VB_Name = "MyClass"
 *   Attribute VB_GlobalNameSpace = False
 *
 * The ANTLR4 VBA grammar handles Attribute lines as part of the grammar,
 * but the VERSION/BEGIN/END header block is not part of standard VBA syntax
 * in all grammar configurations.
 *
 * This module provides utilities to detect and handle module headers.
 */

/**
 * Count the number of header lines that should be skipped for line number
 * adjustment. Returns 0 if no header is detected.
 *
 * Note: We do NOT strip headers — the Rubberduck grammar supports parsing
 * them via moduleHeader and moduleAttributes rules. This function is
 * available if we need line offset adjustment in the future.
 */
export function countModuleHeaderLines(code: string): number {
  const lines = code.split('\n');
  let headerLines = 0;

  for (const line of lines) {
    const trimmed = line.trim();
    if (
      trimmed.startsWith('VERSION ') ||
      trimmed === 'BEGIN' ||
      trimmed === 'END' ||
      trimmed.startsWith('MultiUse') ||
      trimmed.startsWith('Attribute ')
    ) {
      headerLines++;
    } else if (trimmed === '' && headerLines > 0) {
      // Blank line within header area
      headerLines++;
    } else {
      // First non-header, non-blank line — stop counting
      break;
    }
  }

  return headerLines;
}

/**
 * Detect if a file is a class module (.cls) based on content.
 */
export function isClassModule(code: string): boolean {
  return code.trimStart().startsWith('VERSION 1.0 CLASS');
}

/**
 * Detect if a file is a form module (.frm) based on content.
 */
export function isFormModule(code: string): boolean {
  const lines = code.split('\n');
  return lines.some(line => line.trim().startsWith('Begin VB.Form'));
}
