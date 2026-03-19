# vba-lint-mcp

[![License: GPL-3.0](https://img.shields.io/badge/License-GPL--3.0-blue.svg)](https://www.gnu.org/licenses/gpl-3.0)
[![Node.js](https://img.shields.io/badge/Node.js-%3E%3D20-green.svg)](https://nodejs.org/)
[![MCP](https://img.shields.io/badge/MCP-Compatible-purple.svg)](https://modelcontextprotocol.io/)

An MCP server providing VBA code inspections, linting, and parse tree analysis -- 65 inspections ported from Rubberduck VBA.

## Overview

`vba-lint-mcp` exposes VBA static analysis capabilities through the [Model Context Protocol](https://modelcontextprotocol.io/), enabling AI assistants (Claude Code, etc.) and MCP-compatible tools to perform code quality checks on VBA source files.

The inspection logic is derived from the [Rubberduck VBA](https://github.com/rubberduck-vba/Rubberduck) project -- an open-source VBIDE add-in providing code inspections, refactoring, navigation, and unit testing for VBA developers. This project translates Rubberduck's proven C# inspection patterns to TypeScript and makes them accessible via MCP.

**Who is this for?**
- VBA developers using Claude Code or other MCP clients for code assistance
- Teams maintaining Excel/Access VBA projects who want automated code quality checks
- Anyone building MCP-based development toolchains that include VBA support

## Quick Start

### Install

```bash
git clone git@github.com:devinmlowe/vba-lint-mcp.git
cd vba-lint-mcp
npm install
npm run generate-parser   # Requires Java 17+ (JRE)
npm run build
```

### Configure in Claude Code

Add to your MCP settings (`.claude/settings.json` or project settings):

```json
{
  "mcpServers": {
    "vba-lint": {
      "command": "node",
      "args": ["/path/to/vba-lint-mcp/dist/server.js"],
      "env": {
        "VBA_LINT_LOG_LEVEL": "warn"
      }
    }
  }
}
```

### First Inspection

Once configured, Claude Code can inspect VBA code directly:

```
> Inspect this VBA code for issues:
>
> Sub Example()
>     Dim x
>     Let y = 10
>     If True Then
>     End If
> End Sub
```

The server will return diagnostics for missing `Option Explicit`, implicit variant type, obsolete `Let` keyword, and the empty `If` block.

## Tools

### `vba/inspect`

Run inspections on a VBA code string.

| Parameter | Type | Default | Description |
|---|---|---|---|
| `code` | string | (required) | VBA source code to inspect |
| `hostLibraries` | string[] | `["excel"]` | Host libraries available (e.g., `["excel", "access"]`) |
| `severity` | string | (all) | Minimum severity: `"error"`, `"warning"`, `"suggestion"`, `"hint"` |
| `categories` | string[] | (all) | Filter by category (e.g., `["CodeQuality", "Excel"]`) |

**Example output:**
```json
{
  "results": [
    {
      "inspection": "EmptyIfBlock",
      "description": "If block is empty and should either contain code or be removed.",
      "severity": "warning",
      "category": "CodeQuality",
      "tier": "A",
      "location": { "startLine": 3, "startColumn": 4, "endLine": 4, "endColumn": 10 },
      "quickFix": { "description": "Remove the empty If block" },
      "suppressed": false
    }
  ],
  "errors": [],
  "skippedInspections": [],
  "parseErrors": [],
  "engineVersion": "0.1.0"
}
```

### `vba/inspect-workspace`

Scan a directory tree for VBA files (.bas, .cls, .frm) and return aggregated results.

| Parameter | Type | Default | Description |
|---|---|---|---|
| `path` | string | (required) | Directory path to scan |
| `hostLibraries` | string[] | `["excel"]` | Host libraries available |
| `severity` | string | (all) | Minimum severity filter |
| `categories` | string[] | (all) | Category filter |
| `limit` | number | 100 | Maximum results to return |
| `detailed` | boolean | false | Return full results instead of summary |

In summary mode (default), returns counts by file and inspection. In detailed mode, returns full result objects for each finding.

### `vba/list-inspections`

List all available inspections with metadata.

| Parameter | Type | Default | Description |
|---|---|---|---|
| `hostLibraries` | string[] | (all) | Filter by host library |
| `category` | string | (all) | Filter by category |
| `tier` | string | (all) | Filter by tier: `"A"` or `"B"` |

### `vba/parse`

Parse VBA source code and return the AST.

| Parameter | Type | Default | Description |
|---|---|---|---|
| `code` | string | (required) | VBA source code to parse |
| `depth` | number | 3 | Maximum AST depth to return (1-10) |

Useful for understanding code structure. The `depth` parameter limits serialization size for large modules.

## Inspection Catalog

65 inspections across 2 tiers and 8 categories.

### Tier A: Parse-Tree Inspections (43)

These inspections analyze the ANTLR4 parse tree directly and require no symbol resolution. They run on both inline code (`vba/inspect`) and workspace scans.

| ID | Category | Severity | Description |
|---|---|---|---|
| `EmptyIfBlock` | CodeQuality | warning | If block is empty and should either contain code or be removed. |
| `EmptyElseBlock` | CodeQuality | warning | Else block is empty and should either contain code or be removed. |
| `EmptyCaseBlock` | CodeQuality | warning | Case block is empty and should either contain code or be removed. |
| `EmptyForLoopBlock` | CodeQuality | warning | For...Next loop is empty and should either contain code or be removed. |
| `EmptyForEachBlock` | CodeQuality | warning | For Each...Next loop is empty and should either contain code or be removed. |
| `EmptyWhileWendBlock` | CodeQuality | warning | While...Wend loop is empty and should either contain code or be removed. |
| `EmptyDoWhileBlock` | CodeQuality | warning | Do...Loop block is empty and should either contain code or be removed. |
| `EmptyMethod` | CodeQuality | suggestion | Method body is empty and should either contain code or be removed. |
| `EmptyModule` | CodeQuality | hint | Module contains no declarations or procedures and may be unnecessary. |
| `ObsoleteLetStatement` | ObsoleteSyntax | suggestion | The Let keyword is obsolete and should be removed from value assignments. |
| `ObsoleteCallStatement` | ObsoleteSyntax | suggestion | The Call keyword is obsolete. Call the procedure directly without it. |
| `ObsoleteGlobal` | ObsoleteSyntax | suggestion | The Global keyword is obsolete. Use Public instead. |
| `ObsoleteWhileWendStatement` | ObsoleteSyntax | suggestion | While...Wend is obsolete. Use Do While...Loop instead for better control flow. |
| `ObsoleteCommentSyntax` | ObsoleteSyntax | suggestion | The Rem keyword for comments is obsolete. Use the single-quote (') syntax instead. |
| `ObsoleteTypeHint` | ObsoleteSyntax | suggestion | Type hint characters are obsolete. Use explicit As Type declarations instead. |
| `StopKeyword` | CodeQuality | warning | Stop statement halts execution like a breakpoint and should be removed from production code. |
| `EndKeyword` | CodeQuality | warning | Standalone End statement abruptly terminates the program without cleanup. |
| `DefTypeStatement` | ObsoleteSyntax | suggestion | DefType statements (DefBool, DefInt, etc.) implicitly type variables and should be replaced with explicit declarations. |
| `OptionExplicit` | CodeQuality | warning | Module is missing Option Explicit. Variables should be explicitly declared. |
| `OptionBaseZeroOrOne` | CodeQuality | hint | Option Base changes the default array lower bound and can cause confusion. |
| `MultipleDeclarations` | CodeQuality | suggestion | Multiple variables declared on a single line. Declare each on its own line. |
| `ImplicitByRefModifier` | CodeQuality | suggestion | Parameter is implicitly passed ByRef. Specify ByRef or ByVal explicitly. |
| `ImplicitPublicMember` | CodeQuality | suggestion | Member is implicitly Public. Specify Public or Private explicitly. |
| `ImplicitVariantReturnType` | CodeQuality | suggestion | Function/Property Get returns Variant implicitly. Specify an explicit return type. |
| `RedundantByRefModifier` | CodeQuality | hint | ByRef is the default parameter passing mechanism and is redundant when specified. |
| `BooleanAssignedInIfElse` | CodeQuality | suggestion | If/Else block assigns True/False to the same variable. Simplify to a direct assignment. |
| `SelfAssignedDeclaration` | CodeQuality | suggestion | Variable uses As New which auto-instantiates and can mask Nothing checks. |
| `UnreachableCode` | CodeQuality | warning | Code after Exit Sub/Function, End, or GoTo is unreachable. |
| `LineContinuationBetweenKeywords` | CodeQuality | warning | Line continuation character splits a compound keyword, reducing readability. |
| `OnLocalError` | CodeQuality | suggestion | On Local Error is functionally identical to On Error. The Local keyword is redundant. |
| `StepNotSpecified` | CodeQuality | hint | For loop does not specify Step. Consider adding an explicit Step clause. |
| `StepOneIsRedundant` | CodeQuality | hint | Step 1 is the default for For loops and is redundant. |
| `UnhandledOnErrorResumeNext` | ErrorHandling | warning | On Error Resume Next is used without a corresponding On Error GoTo 0 to reset error handling. |
| `OnErrorGoToMinusOne` | ErrorHandling | warning | On Error GoTo -1 clears the current error object. Ensure this is intentional. |
| `EmptyStringLiteral` | Performance | hint | Use vbNullString instead of "" for better performance. |
| `IsMissingOnInappropriateArgument` | CodeQuality | hint | IsMissing only works correctly with Optional Variant parameters. Verify the argument type. |
| `IsMissingWithNonArgumentParameter` | CodeQuality | warning | IsMissing is called with a non-parameter argument. IsMissing only works with procedure parameters. |
| `ImplicitActiveSheetReference` | Excel | suggestion | Unqualified Range/Cells/Rows/Columns implicitly refers to ActiveSheet. |
| `ImplicitActiveWorkbookReference` | Excel | suggestion | Unqualified Sheets/Worksheets/Names implicitly refers to ActiveWorkbook. |
| `SheetAccessedUsingString` | Excel | suggestion | Accessing sheets by string name is fragile. Use the sheet codename property instead. |
| `ApplicationWorksheetFunction` | Excel | hint | Application.WorksheetFunction is redundant. Use WorksheetFunction directly. |
| `ExcelMemberMayReturnNothing` | Excel | warning | Find/FindNext/FindPrevious may return Nothing. Check the result before using it. |
| `ExcelUdfNameIsValidCellReference` | Excel | warning | Function name looks like a cell reference (e.g., A1, B2) and cannot be used as a UDF in Excel formulas. |

### Tier B: Symbol-Aware Inspections (22)

These inspections require symbol resolution (declaration collection + reference resolution). They run on both inline code and workspace scans, with workspace scans providing cross-module resolution.

| ID | Category | Severity | Description |
|---|---|---|---|
| `VariableNotUsed` | CodeQuality | warning | Variable is declared but never referenced. |
| `ParameterNotUsed` | CodeQuality | suggestion | Parameter is declared but never referenced in the procedure body. |
| `NonReturningFunction` | CodeQuality | warning | Function or Property Get never assigns a return value. |
| `ConstantNotUsed` | CodeQuality | warning | Constant is declared but never referenced. |
| `ProcedureNotUsed` | CodeQuality | suggestion | Private procedure is never called. |
| `LineLabelNotUsed` | CodeQuality | suggestion | Line label is declared but never referenced by GoTo or GoSub. |
| `VariableNotAssigned` | CodeQuality | warning | Variable is referenced but never assigned a value. |
| `HungarianNotation` | Naming | suggestion | Variable name uses Hungarian notation prefix. |
| `UseMeaningfulName` | Naming | suggestion | Variable name is too short to be meaningful. |
| `UnderscoreInPublicClassModuleMember` | Naming | warning | Public member name contains underscore, which may conflict with VBA interface dispatch. |
| `ObjectVariableNotSet` | CodeQuality | error | Assignment to object variable without Set keyword. |
| `IntegerDataType` | LanguageOpportunities | suggestion | Integer type is used -- Long is preferred in modern VBA. |
| `VariableTypeNotDeclared` | LanguageOpportunities | suggestion | Variable is declared without an explicit type (implicit Variant). |
| `ModuleScopeDimKeyword` | CodeQuality | suggestion | Module-level variable uses Dim instead of Private. |
| `EncapsulatePublicField` | CodeQuality | suggestion | Public variable in a class module should use Property procedures. |
| `MoveFieldCloserToUsage` | CodeQuality | suggestion | Module-level variable is only used in one procedure -- could be local. |
| `FunctionReturnValueNotUsed` | CodeQuality | suggestion | Function return value is always discarded by callers. |
| `FunctionReturnValueAlwaysDiscarded` | CodeQuality | suggestion | Function return value is never used at any call site -- consider converting to Sub. |
| `ProcedureCanBeWrittenAsFunction` | LanguageOpportunities | suggestion | Sub assigns to a ByRef parameter -- could be a Function instead. |
| `ExcessiveParameters` | CodeQuality | suggestion | Procedure has too many parameters. |
| `ParameterCanBeByVal` | CodeQuality | suggestion | ByRef parameter is never assigned to -- could be ByVal. |
| `UnassignedVariableUsage` | CodeQuality | warning | Variable is used before being assigned a value. |

## Configuration

### `.vbalintrc.json`

Place a `.vbalintrc.json` file in your project root to customize behavior:

```json
{
  "hostLibraries": ["excel"],
  "severity": {
    "EmptyIfBlock": "error",
    "StepNotSpecified": "off"
  }
}
```

### Host Libraries

The `hostLibraries` parameter controls which host-specific inspections run. Default is `["excel"]`. Excel-specific inspections (ImplicitActiveSheetReference, etc.) only run when `"excel"` is included.

### `.vbalintignore`

For workspace scanning, create a `.vbalintignore` file with glob patterns to exclude files:

```
# Ignore generated code
generated/**
*.generated.bas

# Ignore test fixtures
test/**
```

### Inspection Suppression

Use `@Ignore` annotations in VBA comments to suppress specific inspections:

```vba
'@Ignore EmptyIfBlock
If condition Then
End If

'@Ignore EmptyIfBlock, ObsoleteLetStatement
```

## Docker

### Build

```bash
docker build -t vba-lint-mcp .
```

### Run with Docker Compose

```bash
docker compose build
docker compose run --rm vba-lint
```

The MCP server uses stdio transport (stdin/stdout JSON-RPC), so it must be run interactively (`docker compose run`), not detached (`docker compose up`).

### Mount VBA Files for Workspace Scanning

Edit `docker-compose.yml` to mount your VBA project directory:

```yaml
services:
  vba-lint:
    build: .
    stdin_open: true
    volumes:
      - ./my-vba-project:/data:ro
```

Then use the `vba/inspect-workspace` tool with `path: "/data"`.

### Image Size

The runtime image uses `node:22-alpine` with a non-root user. Source files are included for GPL-3.0 compliance.

## Development

### Setup

```bash
npm install
npm run generate-parser   # Requires Java 17+
npm run build
```

### Testing

```bash
npm test                  # All tests
npm run test:watch        # Watch mode
npx vitest run <path>     # Specific test
```

### Adding an Inspection

See [CONTRIBUTING.md](CONTRIBUTING.md) for a step-by-step guide to adding new inspections.

### Project Structure

```
src/
  server.ts              # MCP server entry point
  inspections/
    base.ts              # Inspection contracts
    registry.ts          # Master registration
    runner.ts            # Tiered execution engine
    parse-tree/          # Tier A inspections (by category)
    declaration/         # Tier B inspections
  parser/                # ANTLR4 parser facade
  symbols/               # Symbol resolution (Tier B)
  annotations/           # @Ignore annotation parsing
grammar/
  VBALexer.g4            # ANTLR4 lexer grammar (from Rubberduck)
  VBAParser.g4           # ANTLR4 parser grammar (from Rubberduck)
```

### Known Limitations

- **Conditional compilation** (`#If`, `#Const`, `#Else`) is not preprocessed. All branches are parsed as active code, which may produce false positives in modules using conditional compilation.
- **Tier C inspections** (COM type library dependent, Rubberduck annotation system, VBE runtime) are out of scope. See [PLAN.md](PLAN.md) for the full list.

## Attribution

This project is a derivative work of [Rubberduck VBA](https://github.com/rubberduck-vba/Rubberduck). The VBA grammar and inspection logic are translated from Rubberduck's C# codebase.

**Copyright (C) Rubberduck Contributors** -- Licensed under GPL-3.0.

See [ATTRIBUTION.md](ATTRIBUTION.md) for the complete attribution, derived works table, and contributor acknowledgment.

## License

This project is licensed under the **GNU General Public License v3.0**. See [LICENSE](LICENSE) for the full text.

In accordance with GPL-3.0:
- All derived inspection logic retains attribution to Rubberduck VBA
- Every derived file includes a per-file copyright header
- Docker images include full source code (GPL-3.0 Section 6)
- Consumers of this project must comply with GPL-3.0 terms
