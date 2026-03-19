#!/usr/bin/env node
// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { logger } from './logger.js';
import { warmUpParser } from './parser/index.js';
import { handleParseTool } from './tools/parse.js';
import { handleInspectTool } from './tools/inspect.js';
import { handleListInspections } from './tools/list-inspections.js';
import { handleInspectWorkspaceTool } from './tools/inspect-workspace.js';
import { validateRegistry } from './inspections/registry.js';
import { z } from 'zod';

const VERSION = '0.1.0';

async function main() {
  const startTime = Date.now();
  logger.info({ version: VERSION }, 'vba-lint-mcp starting');

  // Validate inspection registry
  const registryErrors = validateRegistry();
  if (registryErrors.length > 0) {
    logger.error({ errors: registryErrors }, 'Inspection registry validation failed');
  }

  // Warm up ANTLR4 parser (initializes ATN/DFA caches)
  warmUpParser();
  const warmUpMs = Date.now() - startTime;
  logger.info({ warmUpMs }, 'Parser warm-up complete');

  const server = new McpServer({
    name: 'vba-lint-mcp',
    version: VERSION,
  });

  // Register vba/parse tool
  server.tool(
    'vba/parse',
    'Parse VBA source code and return the AST (abstract syntax tree). Useful for understanding code structure.',
    {
      code: z.string().describe('VBA source code to parse'),
      depth: z.number().min(1).max(10).default(3).describe('Maximum AST depth to return (1-10, default 3)'),
    },
    async (input) => {
      logger.info({ tool: 'vba/parse', codeLength: input.code.length, depth: input.depth }, 'Tool call');
      return handleParseTool(input);
    },
  );

  // Register vba/inspect tool
  server.tool(
    'vba/inspect',
    'Run VBA code inspections and return diagnostics with severity, location, and suggested fixes.',
    {
      code: z.string().describe('VBA source code to inspect'),
      hostLibraries: z.array(z.string()).default(['excel']).describe('Host libraries available (e.g., ["excel", "access"])'),
      severity: z.enum(['error', 'warning', 'suggestion', 'hint']).optional().describe('Minimum severity to include'),
      categories: z.array(z.string()).optional().describe('Filter by inspection categories'),
    },
    async (input) => {
      logger.info({ tool: 'vba/inspect', codeLength: input.code.length }, 'Tool call');
      return handleInspectTool(input);
    },
  );

  // Register vba/inspect-workspace tool (placeholder — implemented in Phase 5)
  server.tool(
    'vba/inspect-workspace',
    'Scan a directory of VBA files and return aggregated inspection results.',
    {
      path: z.string().describe('Directory path to scan for VBA files'),
      hostLibraries: z.array(z.string()).default(['excel']).describe('Host libraries available'),
      severity: z.enum(['error', 'warning', 'suggestion', 'hint']).optional().describe('Minimum severity'),
      categories: z.array(z.string()).optional().describe('Filter by categories'),
      limit: z.number().default(100).describe('Maximum number of results to return'),
      detailed: z.boolean().default(false).describe('Return full results instead of summary'),
    },
    async (input) => {
      logger.info({ tool: 'vba/inspect-workspace', path: input.path }, 'Tool call');
      return handleInspectWorkspaceTool(input);
    },
  );

  // Register vba/list-inspections tool
  server.tool(
    'vba/list-inspections',
    'List all available VBA inspections with their descriptions, severities, and categories.',
    {
      hostLibraries: z.array(z.string()).optional().describe('Filter by host libraries'),
      category: z.string().optional().describe('Filter by category'),
      tier: z.enum(['A', 'B']).optional().describe('Filter by tier (A=parse-tree, B=symbol-aware)'),
    },
    async (input) => {
      logger.info({ tool: 'vba/list-inspections' }, 'Tool call');
      return handleListInspections(input);
    },
  );

  // Signal handling for graceful shutdown
  const shutdown = () => {
    logger.info('Shutting down gracefully');
    process.exit(0);
  };
  process.on('SIGTERM', shutdown);
  process.on('SIGINT', shutdown);
  process.on('uncaughtException', (err) => {
    logger.error({ err }, 'Uncaught exception — exiting');
    process.exit(1);
  });

  // Start MCP server on stdio
  const transport = new StdioServerTransport();
  await server.connect(transport);

  const totalStartupMs = Date.now() - startTime;
  logger.info({ totalStartupMs }, 'vba-lint-mcp ready');
}

main().catch((err) => {
  logger.error({ err }, 'Fatal error during startup');
  process.exit(1);
});
