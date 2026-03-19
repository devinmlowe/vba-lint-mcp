# SPEC.md — vba-lint-mcp

**Version:** 0.1.0
**Status:** Active development
**License:** GPL-3.0

## Purpose

vba-lint-mcp is an MCP (Model Context Protocol) server that provides VBA static analysis capabilities. It enables AI assistants and development tools to run code inspections, lint checks, and parse tree analysis on VBA source code through a structured JSON-RPC interface.

The inspection logic is derived from [Rubberduck VBA](https://github.com/rubberduck-vba/Rubberduck), an open-source VBIDE add-in. This project translates Rubberduck's proven C# inspection patterns to TypeScript and exposes them via MCP.

## Scope

**In scope:**
- VBA parsing via ANTLR4 grammar (from Rubberduck)
- Tier A inspections: parse-tree analysis (no symbol resolution required)
- Tier B inspections: symbol-aware analysis (declarations, references, cross-module)
- Workspace scanning for multi-file VBA projects
- Configuration via `.vbalintrc.json`
- `@Ignore` annotation support for suppression

**Out of scope:**
- Tier C inspections requiring COM type libraries or VBE runtime
- Conditional compilation preprocessing (`#If`/`#Else`)
- Code formatting or auto-fix application
- Language Server Protocol (LSP) support

## Architecture

```
MCP Tools (server.ts)
    |
    v
Inspection Runner (tiered execution)
    |
    +-- Tier A: Parse-Tree Inspections (43 inspections)
    +-- Tier B: Declaration/Reference Inspections (18 inspections)
    |
    v
Symbol Resolution (2-pass)
    |-- Pass 1: Symbol Walker (collect declarations)
    |-- Pass 2: Reference Resolver (resolve identifiers)
    |
    v
VBA Parser Facade
    |
    v
ANTLR4 Generated Code (antlr4ng)
    |
    v
Grammar Files (VBALexer.g4, VBAParser.g4)
```

## Tool Surface

| Tool | Purpose | Key Parameters |
|---|---|---|
| `vba/inspect` | Run inspections on a code string | `code`, `hostLibraries?`, `severity?`, `categories?` |
| `vba/inspect-workspace` | Scan a directory of VBA files | `path`, `hostLibraries?`, `severity?`, `categories?`, `limit?`, `detailed?` |
| `vba/list-inspections` | List available inspections | `hostLibraries?`, `category?`, `tier?` |
| `vba/parse` | Return AST for a VBA snippet | `code`, `depth?` |

## Inspection Tiers

| Tier | Infrastructure | Count |
|---|---|---|
| A | Parse tree only | 43 |
| B | Parse tree + symbol table | 18 |
| C (out of scope) | COM type library / VBE runtime | ~25-30 |

## Configuration

Optional `.vbalintrc.json` at project root:
- Severity overrides per inspection
- Host library selection (default: `["excel"]`)
- Inspection enable/disable
- `.vbalintignore` glob patterns for workspace scanning

## Dependencies

| Package | Purpose |
|---|---|
| `@modelcontextprotocol/sdk` | MCP server SDK |
| `antlr4ng` | ANTLR4 TypeScript runtime |
| `pino` | Structured logging (stderr) |
| `zod` | Schema validation |
| `micromatch` | Glob matching for `.vbalintignore` |
