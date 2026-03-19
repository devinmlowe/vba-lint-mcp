// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { BaseErrorListener, type RecognitionException, type Recognizer, type Token } from 'antlr4ng';

export interface ParseError {
  message: string;
  line: number;
  column: number;
}

/**
 * Custom error listener that collects parse errors into a structured array.
 * Sanitizes error messages to avoid leaking sensitive content from VBA source.
 */
export class VBAErrorListener extends BaseErrorListener {
  readonly errors: ParseError[] = [];

  override syntaxError<T extends Token>(
    _recognizer: Recognizer<T>,
    _offendingSymbol: T | null,
    line: number,
    charPositionInLine: number,
    msg: string,
    _e: RecognitionException | null,
  ): void {
    // Sanitize: truncate token text that might contain sensitive data.
    // ANTLR4 error messages can include source code fragments.
    const sanitized = this.sanitizeMessage(msg);

    this.errors.push({
      message: sanitized,
      line,
      column: charPositionInLine,
    });
  }

  private sanitizeMessage(msg: string): string {
    // Truncate any quoted token text to 50 chars max
    const sanitized = msg.replace(/'[^']{50,}'/g, (match) => {
      return `'${match.slice(1, 51)}...'`;
    });
    // Also truncate the overall message
    if (sanitized.length > 200) {
      return sanitized.slice(0, 200) + '...';
    }
    return sanitized;
  }
}
