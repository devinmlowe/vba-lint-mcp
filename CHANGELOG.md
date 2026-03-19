# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## [0.1.0] - 2026-03-19

### Phase 1: Foundation

- Project scaffold: package.json, TypeScript config, vitest, ESLint
- ANTLR4 VBA grammar integration (VBALexer.g4, VBAParser.g4 from Rubberduck v2)
- Parser facade with error collection, encoding detection, BOM stripping, CRLF normalization
- Module header stripping (VERSION, Attribute lines) for exported VBA files
- MCP server with stdio transport and 4 registered tools
- `vba/parse` tool: returns AST with configurable depth (1-10)
- `vba/inspect` tool: runs inspections on inline VBA code strings
- `vba/list-inspections` tool: lists available inspections with filtering
- `vba/inspect-workspace` tool: placeholder registered (implemented in Phase 5)
- Structured logging via pino (stderr)
- ANTLR4 parser warm-up at startup for ATN/DFA cache initialization
- GPL-3.0 license with per-file copyright headers
- ATTRIBUTION.md with Rubberduck VBA acknowledgment and derived works table
- Grammar fidelity tests validating cross-runtime correctness

### Phase 2: Tier A Inspections (43 inspections)

- **Empty Blocks (9):** EmptyIfBlock, EmptyElseBlock, EmptyCaseBlock, EmptyForLoopBlock, EmptyForEachBlock, EmptyWhileWendBlock, EmptyDoWhileBlock, EmptyMethod, EmptyModule
- **Obsolete Syntax (9):** ObsoleteLetStatement, ObsoleteCallStatement, ObsoleteGlobal, ObsoleteWhileWendStatement, ObsoleteCommentSyntax, ObsoleteTypeHint, StopKeyword, EndKeyword, DefTypeStatement
- **Declarations (7):** OptionExplicit, OptionBaseZeroOrOne, MultipleDeclarations, ImplicitByRefModifier, ImplicitPublicMember, ImplicitVariantReturnType, RedundantByRefModifier
- **Code Quality (7):** BooleanAssignedInIfElse, SelfAssignedDeclaration, UnreachableCode, LineContinuationBetweenKeywords, OnLocalError, StepNotSpecified, StepOneIsRedundant
- **Error Handling (5):** UnhandledOnErrorResumeNext, OnErrorGoToMinusOne, EmptyStringLiteral, IsMissingOnInappropriateArgument, IsMissingWithNonArgumentParameter
- **Excel (6):** ImplicitActiveSheetReference, ImplicitActiveWorkbookReference, SheetAccessedUsingString, ApplicationWorksheetFunction, ExcelMemberMayReturnNothing, ExcelUdfNameIsValidCellReference
- Per-inspection unit tests with false-positive coverage
- Clean fixture regression gate (zero diagnostics on well-formed VBA)
- @Ignore annotation support for inspection suppression

### Phase 3: Symbol Resolution

- Declaration model types (variables, constants, procedures, parameters, enums, types, properties)
- Symbol walker (Pass 1): collects declarations from parse tree
- Reference resolver (Pass 2): resolves identifier references to declarations
- Declaration finder query interface for inspections
- Workspace context for multi-module symbol resolution
- Scope chain resolution (module-level, procedure-level, block-level)

### Phase 4: Tier B Inspections (22 inspections)

- **Unused Code (7):** VariableNotUsed, ParameterNotUsed, ConstantNotUsed, ProcedureNotUsed, LineLabelNotUsed, VariableNotAssigned, UnassignedVariableUsage
- **Naming (3):** HungarianNotation, UseMeaningfulName, UnderscoreInPublicClassModuleMember
- **Types (3):** ObjectVariableNotSet, IntegerDataType, VariableTypeNotDeclared
- **Scope (3):** ModuleScopeDimKeyword, EncapsulatePublicField, MoveFieldCloserToUsage
- **Functions (3):** FunctionReturnValueNotUsed, FunctionReturnValueAlwaysDiscarded, ProcedureCanBeWrittenAsFunction, NonReturningFunction
- **Parameters (2):** ExcessiveParameters, ParameterCanBeByVal
- Per-inspection unit tests with false-positive coverage

### Phase 5: Workspace Scanning + Integration Tests

- `vba/inspect-workspace` tool: scans directory trees for .bas/.cls/.frm files
- Cross-module symbol resolution for workspace scanning
- `.vbalintignore` support with micromatch glob patterns
- LRU parse cache keyed on SHA-256 content hash (50-entry eviction)
- Summary mode (default) and detailed mode for workspace results
- Result limiting with configurable cap (default: 100)
- MCP protocol integration tests (JSON-RPC over stdio)
- Performance benchmarks for parse and inspection operations
- Multi-inspection interaction tests

### Phase 6: Packaging + Documentation

- Dockerfile with multi-stage build (node:22-alpine, non-root runtime)
- docker-compose.yml for local usage
- .dockerignore for efficient builds
- SPEC.md (SRC-compliant specification)
- CONTRIBUTING.md with inspection authoring guide
- CHANGELOG.md (this file)
- Final README with full inspection catalog, tool documentation, and configuration guide
