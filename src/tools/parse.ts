// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { z } from 'zod';
import { parseCode } from '../parser/index.js';
import type { ParseTree } from 'antlr4ng';

export const parseToolSchema = z.object({
  code: z.string().describe('VBA source code to parse'),
  depth: z.number().min(1).max(10).default(3).describe('Maximum AST depth to return (1-10, default 3)'),
});

export type ParseToolInput = z.infer<typeof parseToolSchema>;

interface ASTNode {
  type: string;
  text?: string;
  line?: number;
  column?: number;
  children?: ASTNode[];
}

/**
 * Convert an ANTLR4 parse tree to a JSON-serializable AST.
 */
function treeToAST(node: ParseTree, maxDepth: number, currentDepth: number = 0): ASTNode {
  const result: ASTNode = {
    type: node.constructor.name,
  };

  // Terminal nodes have text and position
  if ('symbol' in node && node.symbol) {
    const token = node.symbol as { text: string; line: number; column: number };
    result.text = token.text;
    result.line = token.line;
    result.column = token.column;
  }

  // Recurse into children up to maxDepth
  if (currentDepth < maxDepth && node.children && node.children.length > 0) {
    result.children = node.children.map(child =>
      treeToAST(child, maxDepth, currentDepth + 1)
    );
  } else if (node.children && node.children.length > 0) {
    result.children = [{ type: `... (${node.children.length} children, depth limit reached)` }];
  }

  return result;
}

export function handleParseTool(input: ParseToolInput) {
  const startTime = Date.now();

  try {
    const result = parseCode(input.code);
    const ast = treeToAST(result.tree, input.depth);
    const elapsed = Date.now() - startTime;

    const structured = {
      ast,
      parseErrors: result.errors,
      parseTimeMs: elapsed,
    };

    const errorSummary = result.errors.length > 0
      ? ` with ${result.errors.length} parse error(s)`
      : '';

    return {
      content: [
        {
          type: 'text' as const,
          text: `Parsed ${input.code.split('\n').length} lines of VBA in ${elapsed}ms${errorSummary}. AST depth: ${input.depth}.`,
        },
      ],
      structuredContent: structured,
    };
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    return {
      content: [{ type: 'text' as const, text: `Parse error: ${message}` }],
      isError: true,
    };
  }
}
