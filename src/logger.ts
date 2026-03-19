// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import pino from 'pino';

/**
 * Structured logger that writes to stderr exclusively.
 * stdout is reserved for MCP JSON-RPC protocol messages.
 *
 * Log level configurable via VBA_LINT_LOG_LEVEL env var.
 * Defaults to 'info'. Set to 'silent' to suppress all logging.
 */
export const logger = pino(
  {
    level: process.env.VBA_LINT_LOG_LEVEL ?? 'info',
  },
  pino.destination({ dest: 2, sync: false }), // fd 2 = stderr
);
