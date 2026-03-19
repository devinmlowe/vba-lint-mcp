# vba-lint-mcp — Implementation Plan

**Created:** 2026-03-19
**Revised:** 2026-03-19 (post-NAR rounds 1-3)
**Status:** Approved — NAR findings addressed

---

## 1. Project Overview

Build an MCP server that provides VBA code inspection, linting, and parse tree analysis capabilities. The server exposes tools that Claude Code (or any MCP client) can call during VBA development sessions to get structured diagnostics, severity-filtered results, and quick-fix suggestions.

### 1.1 Design Principles

- **Port, don't invent** — Reuse Rubberduck v2's proven inspection logic, translating C# patterns to TypeScript
- **Tiered execution** — Parse-tree inspections run without symbol resolution; declaration/reference inspections require it. The runner communicates which inspections were skipped and why.
- **Incremental value** — Each phase delivers a usable, testable product
- **Test as you go** — Every inspection has tests asserting ID, location, severity, and false-positive immunity
- **Attribute from day one** — GPL-3.0 license, per-file copyright headers, ATTRIBUTION.md from first derived-code commit
- **Fail gracefully** — Per-inspection isolation, structured error responses, no silent failures

### 1.2 MCP Tool Surface (4 tools)

| Tool | Purpose | Parameters |
|---|---|---|
| `vba/inspect` | Run inspections on a code string | `code`, `hostLibraries?` (string[], default: `["excel"]`), `severity?`, `categories?` |
| `vba/inspect-workspace` | Scan a directory tree | `path`, `hostLibraries?`, `severity?`, `categories?`, `limit?` (default: 100), `detailed?` (default: false) |
| `vba/list-inspections` | List available inspections | `hostLibraries?`, `category?`, `tier?` |
| `vba/parse` | Return AST for a VBA snippet | `code`, `depth?` (default: 3, max: 10) |

**Design decisions (from NAR):**
- `vba/inspect-file` eliminated — Claude Code can read files and pass contents to `vba/inspect`. This removes the entire file-system security surface for single-file inspection. (NAR R2-A2)
- `hostLibraries` replaces `host` — additive array model (e.g., `["excel", "access"]`) instead of single string. Default `["excel"]` since Excel is the dominant VBA host. (NAR R2-A6)
- `vba/inspect-workspace` returns summary by default, full results with `detailed: true`, capped at `limit`. (NAR R3-H6.2)
- `vba/parse` supports `depth` parameter to limit AST serialization size. (NAR R1-L6)

### 1.3 Result Schema

```typescript
interface InspectionResponse {
  results: InspectionResult[];
  errors: InspectionError[];         // Inspections that failed to run
  skippedInspections: SkippedInfo[]; // Inspections skipped due to missing infrastructure
  parseErrors: ParseError[];         // Syntax errors in the VBA code
  engineVersion: string;             // Server version
}

interface InspectionResult {
  inspection: string;        // Inspection ID (e.g., "EmptyIfBlock")
  description: string;       // Human-readable description
  severity: "error" | "warning" | "suggestion" | "hint";
  category: string;          // e.g., "CodeQuality", "Naming", "Obsolete", "Excel"
  tier: "A" | "B";          // A = parse-tree, B = symbol-aware
  source?: string;           // File path (for workspace mode)
  location: {
    startLine: number;
    startColumn: number;
    endLine: number;
    endColumn: number;
  };
  quickFix?: {
    description: string;
    edits?: Array<{          // Supports multi-location fixes (NAR R2-A7)
      location: { startLine: number; startColumn: number; endLine: number; endColumn: number; };
      newText: string;
    }>;
  };
  suppressed: boolean;       // True if @Ignore annotation matched
}

interface InspectionError {
  inspection: string;
  message: string;
}

interface SkippedInfo {
  inspection: string;
  reason: string;  // e.g., "Requires symbol resolution (Tier B); only parse-tree available for inline code"
}

interface ParseError {
  message: string;
  line: number;
  column: number;
}
```

### 1.4 MCP Response Format

Every tool registers a Zod `outputSchema` matching the response shape. Responses use:
- `structuredContent`: typed result object (InspectionResponse or equivalent)
- `content`: single text block with human-readable summary (e.g., "Found 3 warnings, 1 suggestion in 42 lines of VBA")
- Zero results: `structuredContent: { results: [] }` with text "No issues found."
- Parse errors: returned as `parseErrors` in the response, not as MCP-level `isError` (allows partial results from workspace scans)
- Tool-level failures (invalid parameters, server errors): use MCP `isError: true`

### 1.5 Inspection Tier Classification

All inspections are classified into tiers based on required infrastructure:

| Tier | Infrastructure Required | Available In | Count |
|---|---|---|---|
| **A** | Parse tree only | `vba/inspect` (string), `vba/inspect-workspace` | ~55 |
| **B** | Parse tree + symbol table (single-module or cross-module) | `vba/inspect` (string, limited), `vba/inspect-workspace` | ~55 |
| **C (out of scope)** | COM type library, Rubberduck annotations, VBE runtime | Not implementable without COM interop | ~25-30 |

**Tier C inspections (explicitly out of scope):**

Rubberduck-specific (annotation/attribute system):
- MissingAnnotationArgumentInspection
- IllegalAnnotationInspection
- DuplicatedAnnotationInspection
- MissingModuleAnnotationInspection
- AttributeValueOutOfSyncInspection
- MissingAttributeInspection
- ModuleWithoutFolderInspection

COM type library dependent:
- ArgumentWithIncompatibleObjectTypeInspection
- SetAssignmentWithIncompatibleObjectTypeInspection
- ImplicitDefaultMemberAccessInspection (full COM version)
- IndexedDefaultMemberAccessInspection
- RecursiveLetCoercionInspection
- IndexedRecursiveDefaultMemberAccessInspection
- LetCoercionInspection
- ValueRequiredArgumentPassesNothingInspection
- ObjectWhereProcedureIsRequiredInspection
- MemberNotOnInterfaceInspection
- ProcedureRequiredByInterfaceInspection
- ImplementedInterfaceMemberInspection
- SuspiciousPredeclaredInstanceAccessInspection

VBE-context dependent:
- DefaultProjectNameInspection
- PublicControlFieldAccessInspection
- ExcelObjectNameInspection

**Success criteria revised:** 100% of Tier A + Tier B inspections implemented (~110 inspections). Tier C explicitly documented as out of scope with rationale per inspection.

---

## 2. Architecture

### 2.1 Project Structure (co-located tests)

```
vba-lint-mcp/
├── grammar/
│   ├── VBALexer.g4                  # From Rubberduck v2 (commit: pinned)
│   ├── VBAParser.g4                 # From Rubberduck v2 (commit: pinned)
│   └── SOURCE.md                    # Pinned Rubberduck commit hash + upgrade procedure
├── src/
│   ├── server.ts                    # MCP server entry point + signal handling
│   ├── logger.ts                    # Structured logger (pino → stderr)
│   ├── config.ts                    # .vbalintrc.json loading + validation
│   ├── tools/
│   │   ├── inspect.ts               # vba/inspect tool handler
│   │   ├── inspect-workspace.ts     # vba/inspect-workspace tool handler
│   │   ├── list-inspections.ts      # vba/list-inspections tool handler
│   │   └── parse.ts                 # vba/parse tool handler
│   ├── parser/
│   │   ├── index.ts                 # Parser facade (ParseResult type)
│   │   ├── generated/               # antlr4ng-generated TypeScript (gitignored)
│   │   ├── vba-parser-facade.ts     # parseCode() → ParseResult
│   │   ├── error-listener.ts        # Parse error collection (sanitized)
│   │   ├── file-reader.ts           # Encoding detection + BOM stripping + CRLF normalization
│   │   └── module-header.ts         # VBA file header stripping (VERSION, Attribute lines)
│   ├── inspections/
│   │   ├── base.ts                  # Inspection contracts (see Section 2.4)
│   │   ├── registry.ts              # Explicit barrel file registration
│   │   ├── runner.ts                # Tiered execution engine
│   │   ├── types.ts                 # Result types, severity enum, category enum, tier enum
│   │   ├── parse-tree/
│   │   │   ├── empty-blocks/
│   │   │   │   ├── empty-if-block.ts
│   │   │   │   ├── __tests__/
│   │   │   │   │   └── empty-if-block.test.ts
│   │   │   │   └── __fixtures__/
│   │   │   │       ├── empty-if.bas
│   │   │   │       └── non-empty-if.bas
│   │   │   ├── obsolete-syntax/
│   │   │   ├── code-quality/
│   │   │   ├── declarations/
│   │   │   ├── error-handling/
│   │   │   └── excel/
│   │   ├── declaration/
│   │   │   ├── unused/
│   │   │   ├── naming/
│   │   │   ├── types/
│   │   │   ├── scope/
│   │   │   └── functions/
│   │   └── reference/
│   │       ├── access-patterns/
│   │       └── excel/
│   ├── symbols/
│   │   ├── declaration.ts           # Declaration hierarchy (from v2, not v3)
│   │   ├── declaration-finder.ts    # Query interface (multi-pass)
│   │   ├── symbol-walker.ts         # Pass 1: collect declarations
│   │   ├── reference-resolver.ts    # Pass 2: resolve references
│   │   ├── scope.ts                 # Scope chain resolution
│   │   └── workspace.ts             # Multi-module workspace context
│   ├── annotations/
│   │   ├── parser.ts                # @Ignore annotation parsing
│   │   └── types.ts
│   └── resources/
│       ├── en/                      # Key-based string tables (i18n-ready)
│       │   ├── inspection-names.json
│       │   ├── inspection-info.json
│       │   └── quick-fixes.json
│       └── default-config.json
├── test/
│   ├── grammar/                     # Cross-runtime grammar validation (from Rubberduck tests)
│   │   └── grammar-fidelity.test.ts
│   ├── integration/
│   │   └── mcp-protocol.test.ts     # Automated MCP JSON-RPC protocol tests
│   ├── clean-fixtures/              # Files that must produce zero diagnostics
│   │   ├── all-statement-types.bas
│   │   ├── all-control-flow.bas
│   │   ├── excel-correct-patterns.bas
│   │   └── comprehensive-clean.cls
│   └── multi-inspection/            # Interaction / cascading tests
│       └── interaction.test.ts
├── Dockerfile
├── docker-compose.yml
├── package.json
├── tsconfig.json
├── vitest.config.ts
├── PLAN.md
├── SPEC.md
├── README.md
├── CHANGELOG.md
├── CONTRIBUTING.md
├── LICENSE                          # GPL-3.0
└── ATTRIBUTION.md                   # From Phase 1 (not Phase 6)
```

### 2.2 Dependency Graph

```
MCP Tools (server.ts, tools/)
    ↓
Inspection Runner (inspections/runner.ts) ← tiered execution
    ↓
┌─────────────────────────────────────────────────────┐
│ Tier A: Parse-Tree Inspections                       │
│ Tier B-single: Declaration Inspections (one module)  │
│ Tier B-cross: Reference Inspections (workspace)      │
└─────────────────────────────────────────────────────┘
    ↓                ↓                    ↓
Parse Tree     Declaration Finder    Reference Resolver
    ↓                ↓                    ↓
          Symbol Walker (2-pass)
                ↓
          Annotations Parser
                ↓
     VBA Parser Facade (ParseResult)
                ↓
        antlr4ng Generated Code
                ↓
         Grammar Files (.g4)
```

### 2.3 Key Dependencies

| Package | Purpose | Version |
|---|---|---|
| `@modelcontextprotocol/sdk` | MCP server SDK | pinned exact |
| `antlr4ng` | ANTLR4 TypeScript runtime | ^3.0 |
| `antlr4ng-cli` | Grammar → TypeScript generation | ^2.0 (dev) |
| `pino` | Structured logger (stderr) | pinned exact |
| `zod` | Schema validation (MCP outputSchema, config) | pinned exact |
| `micromatch` | .vbalintignore glob matching | pinned exact |
| `vitest` | Test framework | pinned exact |
| `typescript` | Language | ^5.x |
| `tsx` | TypeScript execution | pinned exact |

**ANTLR4 decision (from NAR R1-C4, R2-A1, R3-C5.1):** `antlr4ng` is the only actively maintained TypeScript ANTLR4 runtime. The `antlr4` npm package produces JavaScript (not TypeScript), and `antlr4ts` is abandoned (2021). `antlr4ng-cli` generates native TypeScript. Both runtime and CLI must be pinned to compatible versions.

### 2.4 Core Contracts (Specified)

**Parser Facade:**

```typescript
interface ParseResult {
  tree: VBAModuleContext;     // Root ANTLR4 parse tree node
  tokens: CommonTokenStream;  // For token-level inspection
  errors: ParseError[];       // Collected parse errors (sanitized)
  source?: string;            // File path if from file
}

function parseCode(source: string, options?: ParseOptions): ParseResult;

interface ParseOptions {
  filePath?: string;           // For error attribution
  maxInputSize?: number;       // Override default 512KB limit
}
```

**Inspection Base Classes:**

```typescript
abstract class InspectionBase {
  static readonly id: string;               // e.g., "EmptyIfBlock"
  static readonly tier: "A" | "B";
  static readonly category: InspectionCategory;
  static readonly defaultSeverity: Severity;
  static readonly hostLibraries?: string[];  // undefined = all hosts

  abstract inspect(context: InspectionContext): InspectionResult[];
}

abstract class ParseTreeInspection extends InspectionBase {
  static readonly tier = "A" as const;
  // context.parseResult is always available
}

abstract class DeclarationInspection extends InspectionBase {
  static readonly tier = "B" as const;
  // context.declarationFinder is required
}

abstract class ReferenceInspection extends InspectionBase {
  static readonly tier = "B" as const;
  // context.declarationFinder + context.workspace are required
}

interface InspectionContext {
  parseResult: ParseResult;
  declarationFinder?: DeclarationFinder;  // Available only for Tier B
  workspace?: WorkspaceContext;            // Available only for cross-module
  config: ResolvedConfig;
}
```

**Registry (explicit barrel file, compile-time safe):**

```typescript
// src/inspections/registry.ts
import { EmptyIfBlockInspection } from './parse-tree/empty-blocks/empty-if-block';
// ... all imports explicit

export const ALL_INSPECTIONS: typeof InspectionBase[] = [
  EmptyIfBlockInspection,
  // ... all inspections listed
];

// Startup validation: every registered inspection has matching metadata
export function validateRegistry(): RegistryError[];
```

### 2.5 Error Handling Strategy

| Layer | Behavior |
|---|---|
| **Per-inspection** | Each inspection runs in try/catch. Failures produce an `InspectionError` entry, not a crash. Other inspections continue. |
| **Parser** | Parse failures populate `ParseResult.errors`. Parse tree may be partial (ANTLR4 error recovery). Inspections run on partial trees. |
| **Tool handler** | Invalid parameters return MCP `isError: true` with structured message. |
| **File I/O** | File not found, unreadable, too large → MCP `isError: true`. Encoding detection failures → fall back to UTF-8 with warning in `errors`. |
| **Process** | `SIGTERM`/`SIGINT` → finish current call, close stdio, exit 0. `uncaughtException` → log to stderr, exit 1. |
| **Per-call timeout** | 30-second wall-clock limit via `Promise.race`. Timeout returns partial results collected so far + timeout error. |

### 2.6 Security Controls

| Control | Implementation |
|---|---|
| **Workspace path containment** | `path.resolve()` + verify starts with configured `rootDir`. Reject absolute paths; accept only relative. |
| **Symlink resolution** | `fs.realpath()` before read; re-validate resolved path within rootDir. |
| **Extension allowlist** | Only `.bas`, `.cls`, `.frm` for workspace scanning. |
| **Input size limit** | 512KB max for `vba/inspect` code strings. 512KB per file in workspace. |
| **File count limit** | 500 files per workspace scan (configurable). |
| **Error message sanitization** | Truncate ANTLR4 token text to 50 chars. Strip string literal content. Never include raw source lines — line numbers only. |

### 2.7 Logging Design

- **Channel:** stderr exclusively (stdout is MCP JSON-RPC)
- **Library:** pino (structured JSON)
- **Levels:** ERROR, WARN, INFO, DEBUG (configurable via `VBA_LINT_LOG_LEVEL` env var)
- **What's logged:** server start + version, tool calls (params minus code content), parse errors, inspection errors, timing per call
- **From Phase 1** — not deferred

### 2.8 Configuration

```json
// .vbalintrc.json (workspace root)
{
  "hostLibraries": ["excel"],
  "severity": {
    "ObsoleteLetStatement": "hint",
    "EmptyIfBlock": "error"
  },
  "disabled": ["HungarianNotation"],
  "maxFileSize": 524288,
  "maxFiles": 500
}
```

- Config hierarchy: defaults → `.vbalintrc.json` → per-call parameters
- Validated at startup with Zod schema; malformed entries produce warnings, not crashes
- New inspections are enabled by default (denylist model)

---

## 3. Implementation Phases

### Phase 1: Foundation + Attribution + Contracts (Issues #1-#7)

**Goal:** Working MCP server with parser, core contracts specified, attribution in place, warm-up validated.

| Issue | Title | Description |
|---|---|---|
| #1 | Project scaffold + GPL compliance | package.json, tsconfig, vitest, ESLint, ATTRIBUTION.md, per-file copyright header template, LICENSE verification |
| #2 | ANTLR4 grammar integration | Copy grammars from Rubberduck v2 (pinned commit), generate TypeScript with antlr4ng-cli, verify parser compiles |
| #3 | Grammar validation | Port ≥50 parse fixtures from Rubberduck's `RubberduckTests/Grammar/` to verify cross-runtime fidelity |
| #4 | Parser facade + file reader | `parseCode()` → `ParseResult`, error listener with sanitization, encoding detection (UTF-8 + Windows-1252), BOM stripping, CRLF normalization, module header stripping |
| #5 | MCP server skeleton | Server entry point, 4 tool registrations with Zod schemas, stdio transport, signal handling, pino logger to stderr, ANTLR4 warm-up at startup |
| #6 | `vba/parse` tool | First working tool — returns AST with depth control |
| #7 | Core inspection contracts | Base classes, types, tiered runner skeleton, explicit barrel registry, InspectionContext. Validate with 1 proof-of-concept Tier A inspection. |

**Tests:**
- Grammar fidelity: ≥50 VBA fixtures parse identically to Rubberduck's expected trees (focus: case-insensitive identifiers, line continuations, colon-separated statements, string literals, date literals)
- Parser handles: valid modules, invalid syntax, empty string, binary content, oversized input (>512KB rejected)
- Encoding: UTF-8, UTF-8-BOM, Windows-1252 files all parse correctly
- CRLF: line numbers accurate for CRLF line endings
- Module headers: `.cls` files with VERSION/BEGIN/Attribute headers parse correctly
- MCP protocol: initialize handshake, `vba/parse` with valid/invalid inputs, malformed JSON-RPC
- Server: SIGTERM triggers clean shutdown, warm-up completes before accepting calls
- Cold start benchmark: log startup time, fail if >3 seconds

**README update:** Installation, basic usage, `vba/parse` example, attribution notice, GPL compliance note for consumers.

**Exit criteria:** MCP protocol test passes. Grammar fidelity tests pass. Server starts, warms up, responds to `vba/parse`, shuts down cleanly.

---

### Phase 2: Inspection Framework + Tier A Inspections (Issues #8-#19)

**Goal:** Full inspection framework with all Tier A (parse-tree) inspections, @Ignore support, and the `vba/inspect` tool working.

**Phase 2a: Framework + @Ignore + first batch (Issues #8-#12)**

| Issue | Title | Description |
|---|---|---|
| #8 | Inspection framework finalization | Runner with tiered execution, per-inspection isolation, result construction, timing, partial-result support |
| #9 | @Ignore annotation support | Parse `'@Ignore InspectionName` annotations, apply suppression, set `suppressed: true` in results |
| #10 | Inspection metadata system | Key-based string tables (i18n-ready), startup validation (every registered inspection has metadata, vice versa) |
| #11 | `vba/inspect` tool | Wire runner to MCP tool, severity/category filtering, hostLibraries parameter, MCP response format (structuredContent + text summary) |
| #12 | `vba/list-inspections` tool | Returns catalog with tier, category, severity, description, quick-fix descriptions |

**Phase 2b: All remaining Tier A inspections (Issues #13-#19)**

| Issue | Title | Description |
|---|---|---|
| #13 | Empty block inspections | EmptyIfBlock, EmptyElseBlock, EmptyCaseBlock, EmptyForLoop, EmptyForEach, EmptyWhileWend, EmptyDoWhile, EmptyMethod, EmptyModule (9 inspections) |
| #14 | Obsolete syntax inspections | ObsoleteLet, ObsoleteCall, ObsoleteGlobal, ObsoleteWhileWend, ObsoleteComment, ObsoleteTypeHint, StopKeyword, EndKeyword, DefTypeStatement (9 inspections) |
| #15 | Declaration inspections (parse-tree) | OptionExplicit, OptionBaseZeroOrOne, MultipleDeclarations, ImplicitByRef, ImplicitPublicMember, ImplicitVariantReturn, RedundantByRefModifier (7 inspections) |
| #16 | Code quality inspections (parse-tree) | BooleanAssignedInIfElse, SelfAssignedDeclaration, UnreachableCode, LineContinuationBetweenKeywords, OnLocalError, StepNotSpecified, StepOneIsRedundant (7 inspections) |
| #17 | Error handling inspections | UnhandledOnErrorResumeNext, OnErrorGoToMinusOne, EmptyStringLiteral, IsMissingOnInappropriateArgument, IsMissingWithNonArgumentParameter (5 inspections) |
| #18 | Excel-specific inspections (parse-tree) | ImplicitActiveSheetReference, ImplicitActiveWorkbookReference, SheetAccessedUsingString, ApplicationWorksheetFunction, ExcelMemberMayReturnNothing, ExcelUdfNameIsValidCellReference (6 inspections) |
| #19 | Multi-inspection interaction tests + clean fixtures | ≥10 realistic VBA snippets triggering 3+ inspections, assert exact inspection sets. Clean fixture regression gate (all Tier A inspections, zero findings on all clean fixtures). |

**Test requirements per inspection (NAR R1-C6):**
- Assert `results[0].inspection` (exact ID)
- Assert `results[0].severity`
- Assert `results[0].location.startLine` (at minimum)
- At least one false-positive test (syntactically similar but correct construct)
- At least one Rubberduck-sourced test case where available (translation fidelity)
- @Ignore suppression test for each inspection

**README update:** Full Tier A inspection catalog, `vba/inspect` examples, `vba/list-inspections` usage, .vbalintrc.json configuration guide.

**Exit criteria:** All Tier A inspections pass. Clean fixtures produce zero findings. Multi-inspection interaction tests pass. Performance: single parse + all Tier A inspections < 500ms for 500-line file (first call; <100ms cached).

---

### Phase 3: Symbol Resolution (Issues #20-#25)

**Goal:** Two-pass symbol resolver, declaration finder, workspace context. Validated with proof-of-concept Tier B inspections.

| Issue | Title | Description |
|---|---|---|
| #20 | Declaration model (from v2) | TypeScript port of Rubberduck v2's declaration hierarchy: Module, Sub, Function, Property, Variable, Parameter, Enum, Type, Constant, LineLabel. Uses v2 model (not v3) to match inspection source. |
| #21 | Symbol walker (Pass 1: collect) | Parse tree listener that extracts all declarations. Handles: procedure-level, module-level, block-level scope. Collects but does not resolve forward references. |
| #22 | Reference resolver (Pass 2: resolve) | Resolves identifier references to declarations. Handles: qualified names, `With` block context, `Me` keyword. Marks unresolvable references for external/COM (graceful, not error). |
| #23 | Declaration finder | Query interface: find by name, type, scope; find unused declarations; find references to a declaration. Single-module and cross-module variants. |
| #24 | Workspace context | Multi-module container: `addModule(name, parseResult)`, cross-module Public/Friend visibility, multi-file declaration finder. |
| #25 | Tier B proof-of-concept | Implement 3 Tier B inspections (VariableNotUsed, ParameterNotUsed, NonReturningFunction) to validate the full stack. Verify runner correctly skips Tier B inspections when no symbol table. |

**Multi-pass design (NAR R1-C3):**
- Pass 1 (collect): Walk parse tree, extract all declarations with their scope and visibility. Store in symbol table. Do not attempt to resolve references.
- Pass 2 (resolve): Walk parse tree again, resolve identifier references against the symbol table. Unresolvable references are marked as "external" (likely COM library or unloaded module), not as errors.
- Single-file mode: Both passes run on one module. Cross-module references marked "external."
- Workspace mode: Pass 1 runs on all modules first, then Pass 2 runs with full cross-module visibility.

**Tests:**
- Symbol walker extracts correct declarations from modules with: Subs, Functions, Properties, Enums, Types, module-level variables, procedure-level variables, parameters
- Scope resolution: Public vs Private vs Friend, module-level vs procedure-level vs block-level
- Reference resolution: simple names, qualified names (`Module1.Foo`), `With` blocks, `Me`
- Cross-module: Public procedure in Module1 called from Module2 resolves correctly
- Forward references: Sub A calls Sub B declared later — resolves correctly in pass 2
- Unresolvable references: `Application.Worksheets` marked as external, not error
- Runner skips Tier B inspections for inline code with `skippedInspections` in response

**README update:** Note about Tier B inspections now available, workspace context usage.

**Exit criteria:** 3 proof-of-concept Tier B inspections pass. Runner tiered execution verified. Symbol table correct for multi-module workspace.

---

### Phase 4: All Tier B Inspections (Issues #26-#33)

**Goal:** All remaining Tier B inspections.

| Issue | Title | Description |
|---|---|---|
| #26 | Unused code inspections | VariableNotUsed, ConstantNotUsed, ParameterNotUsed, ProcedureNotUsed, LineLabelNotUsed, VariableNotAssigned, UnassignedVariableUsage (7 inspections — 3 already from #25) |
| #27 | Naming convention inspections | HungarianNotation, UseMeaningfulName, UnderscoreInPublicClassModuleMember (3 inspections) |
| #28 | Type inspections | ObjectVariableNotSet (simplified, without COM type info), IntegerDataType, VariableTypeNotDeclared (3 inspections) |
| #29 | Scope inspections | ModuleScopeDimKeyword, EncapsulatePublicField, MoveFieldCloserToUsage (3 inspections) |
| #30 | Function inspections | NonReturningFunction, FunctionReturnValueNotUsed, FunctionReturnValueAlwaysDiscarded, ProcedureCanBeWrittenAsFunction (4 inspections — 1 already from #25) |
| #31 | Parameter inspections | ExcessiveParameters, ParameterCanBeByVal (2 inspections) |
| #32 | Remaining Tier B inspections | Audit against v2 catalog, implement remaining portable inspections |
| #33 | Tier B interaction tests + clean fixture update | Multi-inspection interactions for Tier B. Clean fixtures updated to include declaration/reference patterns. |

**Tests:** Same per-inspection requirements as Phase 2b. Cross-module tests for inspections that require workspace context.

**README update:** Complete Tier A + Tier B inspection catalog with tier labels.

---

### Phase 5: Workspace Scanning + Integration (Issues #34-#38)

**Goal:** Directory scanning, automated integration tests, performance validation.

| Issue | Title | Description |
|---|---|---|
| #34 | `vba/inspect-workspace` tool | Directory scanning with rootDir containment, symlink resolution, extension filtering, .vbalintignore support (micromatch), file count limit, summary vs detailed mode, parse cache |
| #35 | Parse cache | `ParseCache` keyed on `(filePath, contentHash)`, LRU eviction (50 entries), used by workspace scanner and repeat inspections |
| #36 | .vbalintignore specification | micromatch patterns, workspace root lookup, `#` comments, negation, test fixtures for edge cases |
| #37 | Automated MCP protocol integration tests | Spawn server as child process, send JSON-RPC via stdio, validate responses. Test: initialize, all 4 tools with valid/invalid inputs, concurrent calls, timeout behavior, malformed input. CI-compatible. |
| #38 | Performance benchmarking | Measure: cold start, first-parse 500-line file, cached re-inspect, 50-file workspace scan. Realistic targets: first-parse 300-500ms, cached <50ms, 50-file workspace cold 15-25s, cached <3s. |

**Tests:**
- Workspace scanner: finds all VBA files in nested dirs, respects .vbalintignore, rejects paths outside rootDir, resolves symlinks
- Parse cache: cache hit avoids re-parse, LRU eviction works, content hash invalidates stale cache
- MCP integration: full protocol roundtrip automated in vitest (no tmux, no Claude Code)
- Performance: benchmarks logged, fail if cold start >3s

**README update:** Workspace scanning usage, .vbalintignore format, Claude Code MCP configuration example, performance characteristics.

---

### Phase 6: Packaging + Final Documentation (Issues #39-#43)

**Goal:** Docker image, polished docs, release readiness.

| Issue | Title | Description |
|---|---|---|
| #39 | Dockerfile + docker-compose | Multi-stage build, `node:22-alpine`, `USER node`, read-only volumes, no health check (stdio transport — documented why), image size target <150MB |
| #40 | SPEC.md | SRC-compliant specification |
| #41 | CONTRIBUTING.md + inspection authoring guide | How to add a new inspection: base class, tests, fixtures, metadata, registration |
| #42 | CHANGELOG.md | Version history from v0.1.0 through current |
| #43 | Final README | Complete: install, configure, all tools, full catalog, architecture, Docker, contributing, FAQ, GPL compliance notes |

---

## 4. Testing Strategy

### 4.1 Per-Inspection Tests

Every inspection test asserts:
- `results.length` (expected count)
- `results[0].inspection` (exact inspection ID)
- `results[0].severity` (expected severity)
- `results[0].location.startLine` (exact line)
- At least one false-positive test (similar construct that should NOT trigger)
- At least one Rubberduck-sourced fixture where available (translation fidelity)

### 4.2 Clean Fixture Regression Gate

Multiple clean files covering all VBA constructs:
- `all-statement-types.bas` — Sub, Function, Property, Enum, Type, Declare, Event, WithEvents, Implements
- `all-control-flow.bas` — If/Else, Select Case, For/Next, For Each, Do/Loop, While/Wend, With, GoTo (labeled)
- `excel-correct-patterns.bas` — Qualified ActiveSheet, explicit types, proper error handling
- `comprehensive-clean.cls` — Class module with all patterns

Global regression test: run ALL inspections against ALL clean fixtures; assert zero findings.

### 4.3 Multi-Inspection Interaction Tests

≥10 realistic VBA snippets that trigger 3+ inspections. Assert:
- Exact set of inspections that fire (not just count)
- No contradictory quick-fix suggestions
- Correct location attribution when inspections overlap on same line

### 4.4 Automated Integration Tests (MCP Protocol)

Spawn server as child process, communicate via stdio JSON-RPC:
- Initialize handshake
- Each tool with valid and invalid inputs
- Concurrent tool calls
- Malformed JSON-RPC
- Timeout behavior
- Zero-result responses
- Large result sets

Runs in vitest, CI-compatible, no tmux or Claude Code dependency.

### 4.5 Grammar Fidelity Tests

≥50 VBA fixtures ported from Rubberduck's test suite:
- Case-insensitive identifiers
- Line continuations
- Colon-separated statements
- String literals with embedded quotes
- `#` date literals
- Conditional compilation directives
- Module headers (VERSION, Attribute)

### 4.6 Property-Based / Fuzz Tests

Using `fast-check` or equivalent:
- Generate random VBA-like strings → parse → run inspections → assert no uncaught exceptions
- Edge cases: empty string, only whitespace, only comments, maximum-length lines, deeply nested blocks

### 4.7 Quick-Fix Validation

For inspections with `quickFix.edits`:
- Apply the edit at the reported location
- Re-parse the result
- Re-inspect
- Assert the original inspection no longer fires
- Assert no other inspection is introduced by the fix

---

## 5. Preprocessing Strategy

**Scope for v1:** Conditional compilation (`#If`, `#Const`, `#Else`, `#End If`) is **not supported** in v1. The preprocessor grammar (`VBAPreprocessorParser.g4`) is **not included** in the project.

**Behavior:** Preprocessor directives are treated as comments by the parser. Code inside `#If` blocks is parsed as-is (all branches treated as active). This may produce false positives on code that uses conditional compilation.

**Documentation:** Known limitation documented in README with examples of affected patterns. Future work tracked as a GitHub issue.

---

## 6. Performance Architecture

### 6.1 Parse Cache

```typescript
class ParseCache {
  // Keyed on (filePath, SHA-256 of content)
  // LRU eviction at 50 entries
  // Used by workspace scanner and repeat `vba/inspect` calls on same content
  get(key: string): ParseResult | undefined;
  set(key: string, result: ParseResult): void;
}
```

Designed in Phase 1, implemented in Phase 5. Cache is a performance optimization, not a correctness requirement — server produces correct results after crash-restart with empty cache.

### 6.2 ANTLR4 Warm-Up

At server startup, before accepting tool calls:
```typescript
// Parse a minimal VBA stub to initialize ATN/DFA caches
parseCode("Sub WarmUp()\nEnd Sub");
```

Log startup time. Expected: 200-500ms for grammar initialization.

### 6.3 Realistic Performance Targets

| Scenario | Target | Notes |
|---|---|---|
| Cold start (server startup) | <3s | Grammar ATN initialization |
| First parse, 500-line file | 300-500ms | ANTLR4 TS runtime is 5-15x slower than Java |
| Cached re-inspect, 500-line file | <100ms | Parse cache hit, inspections only |
| 50-file workspace, cold | 15-25s | Acceptable for background scan |
| 50-file workspace, cached | <3s | Cache hits for unchanged files |

---

## 7. GPL-3.0 Compliance

### 7.1 Legal Analysis

Translated inspection logic (C# → TypeScript) constitutes a **derivative work** under copyright law. The entire project is GPL-3.0.

### 7.2 Per-File Copyright Header

Every file derived from Rubberduck includes:
```typescript
// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: [path in Rubberduck repo] (commit: [hash])
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE
```

### 7.3 ATTRIBUTION.md (Phase 1, not Phase 6)

Created in Issue #1. Contains:
- Full Rubberduck copyright notice
- List of derived files with original source paths
- Rubberduck contributor acknowledgment
- Pinned commit hash for grammar and inspection source

### 7.4 Source Availability

Docker images include full source (GPL-3.0 Section 6 compliance). README notes that consumers must comply with GPL-3.0.

### 7.5 Grammar Source Pinning

`grammar/SOURCE.md` contains:
- Exact Rubberduck release tag or commit SHA
- Date of extraction
- Any modifications made to the grammar
- Procedure for upgrading to a newer Rubberduck grammar version

---

## 8. Risk Register

| Risk | Impact | Mitigation |
|---|---|---|
| antlr4ng has undiscovered bugs with VBA grammar | High | Grammar fidelity tests (Phase 1, Issue #3) validate cross-runtime correctness |
| Symbol resolution too complex for practical implementation | High | Phase 3 isolated; Tier A inspections (Phase 2) ship independently. Phase 3 split into collect/resolve passes. |
| ANTLR4 TS runtime performance insufficient | Medium | Mandatory parse cache, realistic targets (not transferred from C# benchmarks), warm-up |
| VBA grammar needs modifications for standalone use | Medium | Grammar fidelity tests against Rubberduck's own test fixtures |
| Translated inspection logic diverges from Rubberduck behavior | Medium | Rubberduck-sourced test fixtures for each inspection; translation fidelity tests |
| Tier C inspections requested by users | Low | Explicitly documented as out of scope with rationale; future work issue |

---

## 9. Commit Strategy

- **Commit after each meaningful unit** — not at end of phase
- **Commit messages** state what changed and why, include plan step context
- **Phase boundary commits** tagged (e.g., `v0.1.0-foundation`, `v0.2.0-tier-a-inspections`)
- **GitHub issues** updated as work progresses (assigned on start, closed with summary on completion)
- **README updated per phase** — specific sections, not just "updated"

---

## 10. Success Criteria

The project is complete when:

1. All 4 MCP tools work correctly
2. 100% of Tier A inspections implemented and passing (~55)
3. 100% of Tier B inspections implemented and passing (~55)
4. Tier C inspections documented as out of scope with per-inspection rationale (~25-30)
5. All tests pass: unit, integration (MCP protocol), grammar fidelity, clean fixture regression, multi-inspection interaction
6. Docker image builds and runs (<150MB, non-root)
7. 6 total NAR rounds completed (3 plan, 3 code) with critical/high findings addressed
8. README accurately reflects all features, tools, inspection catalog (with tier labels), configuration, and Docker usage
9. ATTRIBUTION.md properly credits Rubberduck with per-file copyright headers throughout
10. CHANGELOG.md documents version history
