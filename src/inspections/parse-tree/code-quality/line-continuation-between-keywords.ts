// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/LineContinuationBetweenKeywordsInspection.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { ParseTreeInspection, type InspectionContext, type InspectionMetadata } from '../../base.js';
import type { InspectionResult } from '../../types.js';
import { VBALexer } from '../../../parser/generated/grammar/VBALexer.js';

/**
 * Detects line continuation characters splitting compound keywords.
 *
 * VBA:
 *   End _
 *   Sub                   ' ← line continuation splits "End Sub"
 *
 * This inspection works at the token level rather than the parse tree,
 * checking for LINE_CONTINUATION tokens between keyword pairs.
 */
export class LineContinuationBetweenKeywordsInspection extends ParseTreeInspection {
  static readonly meta: InspectionMetadata = {
    id: 'LineContinuationBetweenKeywords',
    tier: 'A',
    category: 'CodeQuality',
    defaultSeverity: 'warning',
    name: 'Line continuation between keywords',
    description: 'Line continuation character splits a compound keyword, reducing readability.',
    quickFixDescription: 'Remove the line continuation',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const results: InspectionResult[] = [];
    const tokens = context.parseResult.tokens;
    tokens.fill();
    const allTokens = tokens.getTokens();

    // Look for LINE_CONTINUATION token type
    // The lexer defines it as a hidden channel token
    for (let i = 0; i < allTokens.length - 1; i++) {
      const token = allTokens[i];
      if (token.text === ' _' || token.text === '_') {
        // Check if surrounded by keyword tokens on both sides
        const prevToken = i > 0 ? allTokens[i - 1] : null;
        const nextToken = allTokens[i + 1];

        // Skip whitespace/newlines to find next meaningful token
        let nextIdx = i + 1;
        while (nextIdx < allTokens.length) {
          const t = allTokens[nextIdx];
          if (t.type !== VBALexer.NEWLINE && t.type !== VBALexer.WS) break;
          nextIdx++;
        }

        if (prevToken && nextIdx < allTokens.length) {
          const nextMeaningful = allTokens[nextIdx];
          // Check if this looks like a compound keyword pair
          const prevText = prevToken.text?.toUpperCase() ?? '';
          const nextText = nextMeaningful.text?.toUpperCase() ?? '';

          const compoundKeywords = [
            ['END', 'SUB'], ['END', 'FUNCTION'], ['END', 'IF'],
            ['END', 'SELECT'], ['END', 'PROPERTY'], ['END', 'WITH'],
            ['END', 'ENUM'], ['END', 'TYPE'],
            ['EXIT', 'SUB'], ['EXIT', 'FUNCTION'], ['EXIT', 'FOR'],
            ['EXIT', 'DO'], ['EXIT', 'PROPERTY'],
            ['ON', 'ERROR'],
          ];

          const isCompound = compoundKeywords.some(
            ([first, second]) => prevText === first && nextText === second,
          );

          if (isCompound) {
            results.push(
              this.createResult({
                startLine: token.line,
                startColumn: token.column,
                endLine: token.line,
                endColumn: token.column + (token.text?.length ?? 1),
              }),
            );
          }
        }
      }
    }

    return results;
  }
}
