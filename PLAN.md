# vba-lint-mcp — Implementation Plan

**Created:** 2026-03-19
**Status:** Draft — pending NAR review

---

## 1. Project Overview

Build an MCP server that provides VBA code inspection, linting, and parse tree analysis capabilities. The server exposes tools that Claude Code (or any MCP client) can call during VBA development sessions to get structured diagnostics, severity-filtered results, and quick-fix suggestions.

### 1.1 Design Principles

- **Port, don't invent** — Reuse Rubberduck v2's proven inspection logic, translating C# patterns to TypeScript
- **Parse-tree first** — Ship syntax-based inspections before tackling symbol resolution
- **Incremental value** — Each phase delivers a usable, testable product
- **Test as you go** — Every inspection has a corresponding test fixture
- **Attribute properly** — GPL-3.0 license, Rubberduck attribution in code and docs

### 1.2 MCP Tool Surface

| Tool | Purpose | Parameters |
|---|---|---|
| `vba/inspect` | Run inspections on a code string | `code`, `host?` (excel/project/generic), `severity?`, `categories?` |
| `vba/inspect-file` | Inspect a file by path | `path`, `host?`, `severity?`, `categories?` |
| `vba/inspect-workspace` | Scan a directory tree | `path`, `glob?`, `host?`, `severity?`, `categories?` |
| `vba/list-inspections` | List available inspections | `host?`, `category?` |
| `vba/parse` | Return AST for a VBA snippet | `code` |

### 1.3 Result Schema

```typescript
interface InspectionResult {
  inspection: string;        // Inspection ID (e.g., "EmptyIfBlock")
  description: string;       // Human-readable description
  severity: "error" | "warning" | "suggestion" | "hint";
  category: string;          // e.g., "CodeQuality", "Naming", "Obsolete", "Excel"
  location: {
    startLine: number;
    startColumn: number;
    endLine: number;
    endColumn: number;
  };
  quickFix?: {
    description: string;     // e.g., "Replace 'Let x = 5' with 'x = 5'"
    replacement?: string;    // Optional: the actual replacement text
  };
  host?: string;             // If this is a host-specific inspection
}
```

---

## 2. Architecture

### 2.1 Project Structure

```
vba-lint-mcp/
├── grammar/
│   ├── VBALexer.g4                  # ANTLR4 lexer grammar (from Rubberduck v2)
│   ├── VBAParser.g4                 # ANTLR4 parser grammar (from Rubberduck v2)
│   └── VBAPreprocessorParser.g4     # Conditional compilation grammar
├── src/
│   ├── server.ts                    # MCP server entry point
│   ├── tools/
│   │   ├── inspect.ts               # vba/inspect tool handler
│   │   ├── inspect-file.ts          # vba/inspect-file tool handler
│   │   ├── inspect-workspace.ts     # vba/inspect-workspace tool handler
│   │   ├── list-inspections.ts      # vba/list-inspections tool handler
│   │   └── parse.ts                 # vba/parse tool handler
│   ├── parser/
│   │   ├── index.ts                 # Parser facade
│   │   ├── generated/               # ANTLR4-generated TypeScript (gitignored)
│   │   ├── vba-parser-facade.ts     # High-level parse API
│   │   ├── error-listener.ts        # Parse error collection
│   │   └── preprocessor.ts          # Conditional compilation resolver
│   ├── inspections/
│   │   ├── base.ts                  # InspectionBase, ParseTreeInspection, etc.
│   │   ├── registry.ts              # Inspection catalog and discovery
│   │   ├── runner.ts                # Inspection execution engine
│   │   ├── types.ts                 # Result types, severity enum, category enum
│   │   ├── parse-tree/              # Parse-tree inspections (~60)
│   │   │   ├── empty-blocks/
│   │   │   ├── obsolete-syntax/
│   │   │   ├── code-quality/
│   │   │   ├── declarations/
│   │   │   └── excel/
│   │   ├── declaration/             # Declaration-based inspections (~50)
│   │   │   ├── unused/
│   │   │   ├── naming/
│   │   │   ├── types/
│   │   │   └── scope/
│   │   └── reference/               # Reference-based inspections (~30)
│   │       ├── implicit/
│   │       ├── type-safety/
│   │       └── excel/
│   ├── symbols/
│   │   ├── declaration.ts           # Declaration hierarchy (from v3 blueprint)
│   │   ├── declaration-finder.ts    # Symbol query interface
│   │   ├── symbol-walker.ts         # Parse tree → symbol table
│   │   └── scope.ts                 # Scope resolution
│   ├── annotations/
│   │   ├── parser.ts                # @Ignore, @Description annotation parsing
│   │   └── types.ts                 # Annotation type definitions
│   └── resources/
│       ├── inspection-metadata.json # Names, descriptions, quick-fix text (from .resx)
│       └── default-config.json      # Default severity overrides, enabled inspections
├── test/
│   ├── fixtures/                    # VBA test files (.bas, .cls)
│   │   ├── empty-blocks/
│   │   ├── obsolete-syntax/
│   │   ├── naming/
│   │   ├── excel-specific/
│   │   ├── unused-code/
│   │   └── clean/                   # Files that should produce zero diagnostics
│   ├── parser/
│   │   └── parser.test.ts
│   ├── inspections/
│   │   ├── parse-tree.test.ts
│   │   ├── declaration.test.ts
│   │   └── reference.test.ts
│   ├── tools/
│   │   └── tool-handlers.test.ts
│   └── integration/
│       └── mcp-integration.test.ts  # Full MCP protocol tests
├── Dockerfile
├── docker-compose.yml
├── package.json
├── tsconfig.json
├── tsconfig.build.json
├── .eslintrc.json
├── PLAN.md                          # This file
├── SPEC.md                          # SRC specification
├── README.md
├── LICENSE                          # GPL-3.0
└── ATTRIBUTION.md                   # Detailed Rubberduck attribution
```

### 2.2 Dependency Graph

```
MCP Tools (server.ts, tools/)
    ↓
Inspection Runner (inspections/runner.ts)
    ↓
Inspections (inspections/parse-tree/, declaration/, reference/)
    ↓                ↓                    ↓
Parse Tree     Declaration Finder    Reference Tracker
    ↓                ↓                    ↓
            VBA Parser (parser/)
                ↓
        ANTLR4 Generated Code (parser/generated/)
                ↓
            Grammar Files (grammar/)
```

### 2.3 Key Dependencies

| Package | Purpose | Version |
|---|---|---|
| `@modelcontextprotocol/sdk` | MCP server SDK | latest |
| `antlr4` | ANTLR4 TypeScript runtime | ^4.13 |
| `antlr4-tool` or `antlr4ts-cli` | Grammar → TypeScript generation | build-time |
| `vitest` | Test framework | latest |
| `typescript` | Language | ^5.x |
| `tsx` | TypeScript execution | latest |

---

## 3. Implementation Phases

### Phase 1: Foundation (Issues #1-#5)

**Goal:** Working MCP server that can parse VBA and return an AST. No inspections yet.

| Issue | Title | Description |
|---|---|---|
| #1 | Project scaffold | package.json, tsconfig, build scripts, ESLint, vitest config |
| #2 | ANTLR4 grammar integration | Copy grammars from Rubberduck v2, configure antlr4 TS code generation, verify parser builds |
| #3 | Parser facade | High-level API: `parseCode(source: string) → ParseTree`, error listener, preprocessor stub |
| #4 | MCP server skeleton | Server entry point, tool registration for all 5 tools, stdio transport |
| #5 | `vba/parse` tool | First working tool — accepts VBA code, returns AST as structured JSON |

**Tests:**
- Parser correctly parses valid VBA modules (Sub, Function, Property, Class, standard module)
- Parser returns meaningful errors for invalid syntax
- `vba/parse` tool responds with correct MCP protocol format
- MCP server starts and responds to `initialize` handshake

**README update:** Installation instructions, basic usage example with `vba/parse`.

**Exit criteria:** `echo '{"jsonrpc":"2.0","id":1,"method":"tools/call","params":{"name":"vba/parse","arguments":{"code":"Sub Test()\nEnd Sub"}}}' | node dist/server.js` returns a valid AST.

---

### Phase 2: Inspection Framework + Parse-Tree Inspections (Issues #6-#15)

**Goal:** Full inspection framework with all ~60 parse-tree inspections working.

| Issue | Title | Description |
|---|---|---|
| #6 | Inspection framework | Base classes (InspectionBase, ParseTreeInspection), result types, severity/category enums, runner |
| #7 | Inspection registry | Auto-discovery of inspection classes, metadata loading from JSON, host filtering |
| #8 | Empty block inspections | EmptyIfBlock, EmptyElseBlock, EmptyCaseBlock, EmptyForLoop, EmptyWhileLoop, EmptyDoWhile, EmptyMethod, EmptyModule (~8 inspections) |
| #9 | Obsolete syntax inspections | ObsoleteLet, ObsoleteCall, ObsoleteGlobal, ObsoleteWhileWend, ObsoleteComment, ObsoleteTypeHint, StopKeyword, EndKeyword (~8 inspections) |
| #10 | Declaration inspections (parse-tree) | MissingOptionExplicit, MultipleDeclarations, ImplicitByRef, ImplicitPublicMember, ImplicitVariantReturn, ModuleWithoutOption (~8 inspections) |
| #11 | Code quality inspections (parse-tree) | MagicNumber, MultilineFunction, BooleanAssignedInIfElse, AssignmentNotUsed, SelfAssignment, UnreachableCode, GoToUsage (~8 inspections) |
| #12 | Excel-specific inspections (parse-tree) | ImplicitActiveSheetReference, ImplicitActiveWorkbookReference, SheetAccessedByString, ApplicationWorksheetFunction, MemberMayReturnNothing (~6 inspections) |
| #13 | Error handling inspections | UnhandledOnErrorResumeNext, OnErrorGoToMinusOne, EmptyOnErrorBlock, OnErrorWithoutHandler (~4 inspections) |
| #14 | `vba/inspect` and `vba/inspect-file` tools | Wire inspection runner to MCP tools, severity filtering, category filtering, host parameter |
| #15 | `vba/list-inspections` tool | Returns full catalog with descriptions, severities, categories, quick-fix descriptions |

**Inspection metadata:** Extract names, descriptions, rationale, and quick-fix text from Rubberduck v2's `.resx` files into `inspection-metadata.json`.

**Tests:**
- Each inspection has at least one "should trigger" and one "should not trigger" VBA fixture
- Severity filtering works correctly
- Category filtering works correctly
- Host-specific inspections only trigger when host matches
- `vba/list-inspections` returns complete catalog

**README update:** Full tool documentation, example outputs, inspection catalog table, configuration.

**Exit criteria:** All parse-tree inspections pass tests. `vba/inspect` returns correct diagnostics for test VBA files.

---

### Phase 3: Symbol Resolution (Issues #16-#19)

**Goal:** Declaration model and symbol resolver, enabling declaration-based and reference-based inspections.

| Issue | Title | Description |
|---|---|---|
| #16 | Declaration model | TypeScript port of Rubberduck v3's declaration hierarchy: Project, Module, Member, Variable, Parameter, Enum, Type |
| #17 | Symbol walker | Parse tree listener that extracts declarations and builds the symbol table |
| #18 | Declaration finder | Query interface: find by name, type, scope, module; find unused; find references |
| #19 | Scope resolution | Scope chain for variable/procedure visibility, module-level vs procedure-level vs block-level |

**Tests:**
- Symbol walker correctly extracts declarations from complex VBA modules
- Declaration finder queries return correct results
- Scope resolution handles Public/Private/Friend, module-level vs local
- Handles multiple modules in a workspace context

**README update:** Note about symbol-aware inspections now available.

**Exit criteria:** Given a VBA module with functions, variables, and parameters, the symbol walker produces a correct, queryable symbol table.

---

### Phase 4: Declaration & Reference Inspections (Issues #20-#27)

**Goal:** All remaining ~80 inspections from Rubberduck v2's catalog.

| Issue | Title | Description |
|---|---|---|
| #20 | Unused code inspections | VariableNotUsed, ConstantNotUsed, ParameterNotUsed, ProcedureNotUsed, LineLabelNotUsed, VariableNotAssigned (~6 inspections) |
| #21 | Naming convention inspections | HungarianNotation, UseMeaningfulName, DefaultProjectName, NonReturningFunction, ProcedureNameTooLong, ExcessiveParameters (~8 inspections) |
| #22 | Type safety inspections | ObjectVariableNotSet, IncompatibleObjectType, IntegerDataType, VariableTypeNotDeclared, SetAssignmentOnValueType (~6 inspections) |
| #23 | Default member access inspections | ImplicitDefaultMemberAccess, IndexedDefaultMemberAccess, BangNotation, RecursiveDefaultMemberAccess (~5 inspections) |
| #24 | Scope and visibility inspections | ModuleScopeDimKeyword, PublicControlFieldAccess, EncapsulatePublicField (~4 inspections) |
| #25 | Excel reference inspections (advanced) | ExcelUdfNameIsValidCellReference, ApplicationWorksheetFunction, ImplicitActiveSheetReference (reference-based variant) (~4 inspections) |
| #26 | Remaining miscellaneous inspections | All inspections from v2 not covered above — audit against full v2 catalog |
| #27 | @Ignore annotation support | Parse @Ignore annotations, suppress matching inspections in results |

**Tests:**
- Each inspection has positive and negative test fixtures
- @Ignore correctly suppresses inspections
- Workspace-level inspections work across multiple files
- Cross-module reference tracking works

**README update:** Complete inspection catalog with all ~139 inspections listed.

---

### Phase 5: Workspace Scanning & Integration (Issues #28-#31)

**Goal:** Directory scanning, integration testing, and real-world validation.

| Issue | Title | Description |
|---|---|---|
| #28 | `vba/inspect-workspace` tool | Recursively scan directory for .bas/.cls/.frm files, aggregate results, respect .vbalintignore |
| #29 | Integration test suite | Full MCP protocol tests: initialize, tool calls, result validation |
| #30 | Claude Code integration test | Load MCP in Claude Code via tmux, run inspections on test VBA files, validate results |
| #31 | Performance benchmarking | Measure parse + inspect time for single files and workspace of 50+ files, optimize if >2s |

**Tests:**
- Workspace scanner finds all VBA files in nested directories
- .vbalintignore patterns work correctly
- MCP protocol roundtrip works end-to-end
- Claude Code can load and use the MCP server
- Performance: single file <100ms, 50-file workspace <5s

**README update:** Claude Code configuration example, performance characteristics.

---

### Phase 6: Packaging & Documentation (Issues #32-#35)

**Goal:** Docker image, polished docs, and release readiness.

| Issue | Title | Description |
|---|---|---|
| #32 | Dockerfile & docker-compose | Multi-stage build, minimal image, stdio transport, health check |
| #33 | ATTRIBUTION.md | Detailed Rubberduck attribution: contributors, license, specific files derived |
| #34 | SPEC.md | SRC-compliant specification document |
| #35 | Final README | Complete documentation: install, configure, use, inspect catalog, architecture, contributing |

**Tests:**
- Docker image builds successfully
- MCP server works correctly from Docker container
- All existing tests pass in Docker environment

**README update:** Docker usage, complete API reference, contributing guide.

---

## 4. Testing Strategy

### 4.1 Unit Tests (vitest)

Every inspection has a test file with:
- **Positive case:** VBA code that SHOULD trigger the inspection
- **Negative case:** VBA code that should NOT trigger the inspection
- **Edge cases:** Boundary conditions specific to the inspection

Example pattern:
```typescript
describe("EmptyIfBlockInspection", () => {
  it("detects empty If block", () => {
    const code = `Sub Test()\n  If True Then\n  End If\nEnd Sub`;
    const results = inspect(code, [EmptyIfBlockInspection]);
    expect(results).toHaveLength(1);
    expect(results[0].severity).toBe("warning");
  });

  it("ignores If block with content", () => {
    const code = `Sub Test()\n  If True Then\n    x = 1\n  End If\nEnd Sub`;
    const results = inspect(code, [EmptyIfBlockInspection]);
    expect(results).toHaveLength(0);
  });
});
```

### 4.2 VBA Test Fixtures

Stored in `test/fixtures/` as real `.bas` and `.cls` files:

```
test/fixtures/
├── empty-blocks/
│   ├── empty-if.bas           # Triggers EmptyIfBlock
│   ├── empty-else.bas         # Triggers EmptyElseBlock
│   └── non-empty-if.bas       # Should produce no diagnostics
├── obsolete-syntax/
│   ├── obsolete-let.bas       # Triggers ObsoleteLet
│   ├── obsolete-call.bas      # Triggers ObsoleteCall
│   └── modern-syntax.bas      # Should produce no diagnostics
├── excel-specific/
│   ├── implicit-activesheet.bas
│   └── qualified-references.bas
├── clean/
│   └── well-written-module.bas  # Should produce zero diagnostics
└── comprehensive/
    └── kitchen-sink.bas        # Triggers as many inspections as possible
```

### 4.3 Integration Tests (Claude Code via tmux)

Validate the MCP works end-to-end in Claude Code:

1. Start MCP server in a tmux pane
2. Launch Claude Code in another pane with the MCP configured
3. Send VBA code inspection requests
4. Capture and validate responses

### 4.4 Coverage Target

- **Parser:** >90% of grammar rules exercised
- **Inspections:** 100% of inspections have at least one positive + negative test
- **MCP tools:** All 5 tools tested with valid and invalid inputs
- **Integration:** At least one end-to-end MCP roundtrip test

---

## 5. Rubberduck v2 Inspection Catalog

### 5.1 Complete Inspection List (Target: All 139)

**Parse-Tree Inspections (~60) — Phase 2:**

Empty Blocks:
1. EmptyIfBlockInspection
2. EmptyElseBlockInspection
3. EmptyCaseBlockInspection
4. EmptyForLoopBlockInspection
5. EmptyForEachBlockInspection
6. EmptyWhileWendBlockInspection
7. EmptyDoWhileBlockInspection
8. EmptyMethodInspection
9. EmptyModuleInspection

Obsolete Syntax:
10. ObsoleteCallStatementInspection
11. ObsoleteLetStatementInspection
12. ObsoleteGlobalInspection
13. ObsoleteWhileWendStatementInspection
14. ObsoleteCommentSyntaxInspection
15. ObsoleteTypeHintInspection
16. StopKeywordInspection
17. EndKeywordInspection

Declarations (Parse-Tree Level):
18. OptionExplicitInspection
19. OptionBaseZeroOrOneInspection
20. MultipleDeclarationsInspection
21. ImplicitByRefModifierInspection
22. ImplicitPublicMemberInspection
23. ImplicitVariantReturnTypeInspection
24. ModuleWithoutFolderInspection

Code Quality:
25. BooleanAssignedInIfElseInspection
26. AssignmentNotUsedInspection
27. SelfAssignedDeclarationInspection
28. UnreachableCodeInspection
29. MissingAnnotationArgumentInspection
30. RedundantByRefModifierInspection

Error Handling:
31. UnhandledOnErrorResumeNextInspection
32. OnErrorGoToMinusOneInspection
33. EmptyStringLiteralInspection
34. IsMissingOnInappropriateArgumentInspection
35. IsMissingWithNonArgumentParameterInspection

Excel-Specific (Parse-Tree):
36. ImplicitActiveSheetReferenceInspection
37. ImplicitActiveWorkbookReferenceInspection
38. SheetAccessedUsingStringInspection
39. ApplicationWorksheetFunctionInspection
40. ExcelMemberMayReturnNothingInspection
41. ExcelUdfNameIsValidCellReferenceInspection

Miscellaneous Parse-Tree:
42. LineContinuationBetweenKeywordsInspection
43. MissingAttributeInspection
44. IllegalAnnotationInspection
45. DuplicatedAnnotationInspection
46. AttributeValueOutOfSyncInspection
47. MissingModuleAnnotationInspection
48. OnLocalErrorInspection
49. DefTypeStatementInspection
50. StepIsNotSpecifiedInspection
51. StepOneIsRedundantInspection

**Declaration-Based Inspections (~50) — Phase 4:**

Unused Code:
52. VariableNotUsedInspection
53. ConstantNotUsedInspection
54. ParameterNotUsedInspection
55. ProcedureNotUsedInspection
56. LineLabelNotUsedInspection
57. VariableNotAssignedInspection
58. UnassignedVariableUsageInspection

Naming:
59. HungarianNotationInspection
60. UseMeaningfulNameInspection
61. DefaultProjectNameInspection
62. UnderscoreInPublicClassModuleMemberInspection
63. ExcelObjectNameInspection

Types:
64. ObjectVariableNotSetInspection
65. IntegerDataTypeInspection
66. VariableTypeNotDeclaredInspection
67. ArgumentWithIncompatibleObjectTypeInspection
68. SetAssignmentWithIncompatibleObjectTypeInspection

Scope:
69. ModuleScopeDimKeywordInspection
70. EncapsulatePublicFieldInspection
71. PublicControlFieldAccessInspection
72. MoveFieldCloserToUsageInspection

Functions:
73. NonReturningFunctionInspection
74. FunctionReturnValueNotUsedInspection
75. FunctionReturnValueAlwaysDiscardedInspection
76. ProcedureCanBeWrittenAsFunctionInspection

Parameters:
77. ExcessiveParametersInspection
78. ParameterCanBeByValInspection
79. ImplementedInterfaceMemberInspection
80. SuspiciousPredeclaredInstanceAccessInspection

Object Lifecycle:
81. ObjectMemberMayReturnNothingInspection
82. MemberNotOnInterfaceInspection
83. ProcedureRequiredByInterfaceInspection
84. EmptyMethodInspection (declaration variant)

Miscellaneous Declaration:
85-100. (Remaining declaration inspections from v2 catalog — will be fully enumerated during Phase 4 implementation)

**Reference-Based Inspections (~30) — Phase 4:**

Default Member Access:
101. ImplicitDefaultMemberAccessInspection
102. IndexedDefaultMemberAccessInspection
103. UseOfBangNotationInspection
104. IndexedRecursiveDefaultMemberAccessInspection
105. RecursiveLetCoercionInspection
106. LetCoercionInspection

Type Coercion:
107. ValueRequiredArgumentPassesNothingInspection
108. ObjectWhereProcedureIsRequiredInspection
109. ProcedureRequiredInspection

Reference Quality:
110-139. (Remaining reference inspections from v2 catalog — will be fully enumerated during Phase 4 implementation)

### 5.2 Metadata Extraction

For each inspection, extract from Rubberduck v2's `.resx` files:
- **InspectionNames.resx** → `name` field
- **InspectionInfo.resx** → `description` field (rationale)
- **QuickFixes.resx** → `quickFix.description` field
- **CodeInspectionType** attribute → `category` field
- **CodeInspectionSeverity** attribute → default `severity` field

Store as `src/resources/inspection-metadata.json`.

---

## 6. Quality Gates

### 6.1 Pre-Implementation (This Plan)

- [ ] 3 rounds of Non-Advocate Review on this plan
- [ ] Address all critical and high findings
- [ ] Plan approved

### 6.2 Per-Phase

- [ ] All tests pass
- [ ] No TypeScript compiler errors
- [ ] ESLint clean
- [ ] README updated with new features
- [ ] GitHub issues updated (closed with notes)
- [ ] Commit pushed

### 6.3 Post-Implementation

- [ ] 3 rounds of Non-Advocate Review on the code
- [ ] Address all critical and high findings
- [ ] Integration test passes (Claude Code loads MCP, runs inspections)
- [ ] Docker image builds and works
- [ ] Final README review

---

## 7. Risk Register

| Risk | Impact | Mitigation |
|---|---|---|
| ANTLR4 TS target has bugs/limitations | High | Test grammar generation early (Phase 1). Fall back to `antlr4ng` if `antlr4` package has issues |
| Symbol resolution too complex for single pass | Medium | Phase 3 is isolated — parse-tree inspections (Phase 2) ship independently |
| Rubberduck v2 inspection logic doesn't translate cleanly to TS | Medium | Start with simplest inspections, establish patterns, then scale |
| Performance: large workspaces too slow | Low | Lazy parsing, file-level caching, parallel inspection. Benchmark in Phase 5 |
| ANTLR4 grammar needs modifications for standalone use | Medium | Test with diverse VBA samples early. The grammar is well-tested in Rubberduck |

---

## 8. Commit Strategy

- **Commit after each meaningful unit** — not at end of phase
- **Commit messages** state what changed and why
- **Phase boundary commits** tagged (e.g., `v0.1.0-foundation`, `v0.2.0-parse-tree-inspections`)
- **GitHub issues** updated as work progresses (assigned on start, closed with summary on completion)

---

## 9. Attribution Requirements

### 9.1 Files Derived from Rubberduck

| Our File | Derived From | Nature of Derivation |
|---|---|---|
| `grammar/VBALexer.g4` | `Rubberduck.Parsing/Grammar/VBALexer.g4` | Direct copy |
| `grammar/VBAParser.g4` | `Rubberduck.Parsing/Grammar/VBAParser.g4` | Direct copy |
| `src/resources/inspection-metadata.json` | `Rubberduck.Resources/Inspections/*.resx` | Extracted and reformatted |
| `src/inspections/**/*.ts` | `Rubberduck.CodeAnalysis/Inspections/Concrete/*.cs` | Logic translated C# → TypeScript |
| `src/symbols/declaration.ts` | `Rubberduck.Parsing.Legacy/Model/Symbols/Declaration.cs` (v3) | Architecture ported |

### 9.2 Required Attribution

- `LICENSE` — GPL-3.0 (same as Rubberduck)
- `ATTRIBUTION.md` — Detailed list of derived works, Rubberduck contributor acknowledgment
- Source file headers — Brief attribution comment in files derived from Rubberduck
- `README.md` — Attribution section (already present)

---

## 10. Success Criteria

The project is complete when:

1. All 5 MCP tools work correctly
2. ≥130 of 139 Rubberduck v2 inspections are implemented (some may not apply outside VBE context)
3. All inspections have passing tests
4. Claude Code can load the MCP and successfully inspect VBA files
5. Docker image builds and runs
6. 6 total NAR rounds completed (3 plan, 3 code) with findings addressed
7. README accurately reflects all features and usage
8. ATTRIBUTION.md properly credits Rubberduck
