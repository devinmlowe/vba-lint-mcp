// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { CharStream, CommonTokenStream } from 'antlr4ng';
import { VBALexer } from './generated/grammar/VBALexer.js';
import { VBAParser, type StartRuleContext } from './generated/grammar/VBAParser.js';
import { VBAErrorListener, type ParseError } from './error-listener.js';
import { patchParserHelpers, patchLexerHelpers } from './vba-parser-helpers.js';

export type { ParseError } from './error-listener.js';

const DEFAULT_MAX_INPUT_SIZE = 512 * 1024; // 512KB

export interface ParseOptions {
  /** File path for error attribution */
  filePath?: string;
  /** Override default 512KB input size limit */
  maxInputSize?: number;
}

export interface ParseResult {
  /** Root ANTLR4 parse tree node */
  tree: StartRuleContext;
  /** Token stream for token-level inspection */
  tokens: CommonTokenStream;
  /** Collected parse errors (sanitized) */
  errors: ParseError[];
  /** File path if from file */
  source?: string;
}

/**
 * Parse VBA source code into a parse tree.
 *
 * @param source - VBA source code string
 * @param options - Parse options
 * @returns ParseResult with tree, tokens, and any errors
 * @throws Error if input exceeds size limit
 */
export function parseCode(source: string, options?: ParseOptions): ParseResult {
  const maxSize = options?.maxInputSize ?? DEFAULT_MAX_INPUT_SIZE;

  if (source.length > maxSize) {
    throw new Error(
      `Input exceeds maximum size of ${maxSize} bytes (got ${source.length}). ` +
      `Use maxInputSize option to override.`
    );
  }

  const input = CharStream.fromString(source);
  const lexer = new VBALexer(input);
  patchLexerHelpers(lexer);

  const lexerErrors = new VBAErrorListener();
  lexer.removeErrorListeners();
  lexer.addErrorListener(lexerErrors);

  const tokens = new CommonTokenStream(lexer);

  const parser = new VBAParser(tokens);
  patchParserHelpers(parser);

  const parserErrors = new VBAErrorListener();
  parser.removeErrorListeners();
  parser.addErrorListener(parserErrors);

  const tree = parser.startRule();

  return {
    tree,
    tokens,
    errors: [...lexerErrors.errors, ...parserErrors.errors],
    source: options?.filePath,
  };
}

/**
 * Warm up the ANTLR4 parser by parsing a minimal VBA stub.
 * This initializes ATN serialization and DFA caches.
 * Call once at server startup.
 */
export function warmUpParser(): void {
  parseCode('Sub WarmUp()\nEnd Sub\n');
}

// Re-export generated types that inspections need
export { VBAParser } from './generated/grammar/VBAParser.js';
export { VBALexer } from './generated/grammar/VBALexer.js';
export { VBAParserListener } from './generated/grammar/VBAParserListener.js';
export { VBAParserVisitor } from './generated/grammar/VBAParserVisitor.js';
