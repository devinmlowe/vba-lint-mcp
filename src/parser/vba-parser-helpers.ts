// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.Parsing/Grammar/VBABaseParser.cs, VBABaseLexer.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE
//
// These helper functions were originally methods on VBABaseParser/VBABaseLexer
// (the superClass of the generated parser/lexer). Since we removed the superClass
// option for standalone use, we inject them as global functions that the semantic
// predicates can call.

import { VBAParser } from './generated/grammar/VBAParser.js';
import { VBALexer } from './generated/grammar/VBALexer.js';

/**
 * Patch a VBAParser instance with the helper functions needed by semantic predicates.
 * Must be called after creating the parser but before parsing.
 *
 * The original Rubberduck VBABaseParser used:
 *   TokenTypeAtRelativePosition(i) => _input.La(i)  (1-based lookahead on token stream)
 *   IsTokenType(actual, ...expected) => linear scan for match
 */
export function patchParserHelpers(parser: VBAParser): void {
  const tokenStream = parser.tokenStream;

  // La(i) returns the token TYPE at offset i from current position (1-based).
  // La(1) = current/next token, La(-1) = previous token, etc.
  (globalThis as any).TokenTypeAtRelativePosition = (offset: number): number => {
    return tokenStream.LA(offset);
  };

  (globalThis as any).IsTokenType = (tokenType: number, ...types: number[]): boolean => {
    for (const expected of types) {
      if (tokenType === expected) return true;
    }
    return false;
  };

  // Expose all token type constants as globals.
  // In the C# target, these were inherited from the parser's base class.
  // The generated semantic predicates reference them as bare names (e.g., NEWLINE, IDENTIFIER).
  // We iterate over VBAParser's static numeric properties (token type constants are numbers > 0).
  for (const key of Object.getOwnPropertyNames(VBAParser)) {
    const val = (VBAParser as any)[key];
    if (typeof val === 'number' && val > 0 && key === key.toUpperCase()) {
      (globalThis as any)[key] = val;
    }
  }
}

/**
 * Patch a VBALexer instance with helper functions needed by lexer semantic predicates.
 * Must be called after creating the lexer but before lexing.
 *
 * The original Rubberduck VBABaseLexer used:
 *   CharAtRelativePosition(i) => _input.La(i)  (1-based lookahead on char stream)
 *   IsChar(actual, ...expected) => char comparison
 */
export function patchLexerHelpers(lexer: VBALexer): void {
  const charStream = lexer.inputStream;

  (globalThis as any).CharAtRelativePosition = (offset: number): number => {
    return charStream.LA(offset);
  };

  (globalThis as any).IsChar = (actual: number, ...expected: number[]): boolean => {
    for (const exp of expected) {
      if (actual === exp) return true;
    }
    return false;
  };
}
