# NAR Round 2 — Consolidated Plan Review

**Reviewer:** Round 2 (Fresh perspective, different reviewer from Round 1)
**Date:** 2026-03-19
**Target:** PLAN.md (Implementation Plan)
**Round:** 2 of 3

---

## Preamble

Round 1 produced 80+ findings across six dimensions. Many were excellent. This review deliberately avoids re-litigating those findings and instead challenges assumptions that Round 1 either accepted or did not probe deeply enough. Where Round 1 findings are referenced, it is to argue they did not go far enough or identified the symptom rather than the root cause.

---

## ARCHITECTURE (PRIMARY)

### A1. [CRITICAL] `antlr4` npm Package Is the Wrong Choice — `antlr4ng` Should Be the Default, Not the Fallback

**Description:** The plan lists `antlr4 ^4.13` as the primary runtime and mentions `antlr4ng` only in the risk register as a fallback. Round 1 flagged this as "unresolved" and recommended a spike. That recommendation was too timid. The situation is not ambiguous — `antlr4ng` is the correct choice and `antlr4` is a known-problematic option.

The `antlr4` npm package (the official ANTLR4 JavaScript runtime) has well-documented issues in TypeScript contexts:
- Its TypeScript typings are incomplete and frequently out of date relative to the JavaScript source.
- The `antlr4ts` package (which `antlr4-tool` / `antlr4ts-cli` targets) is unmaintained since 2021 and stuck at ANTLR 4.9.
- The official `antlr4` JavaScript runtime uses a CommonJS module structure that creates friction with modern ESM-first TypeScript builds.
- Error listener, visitor, and listener interfaces differ significantly between `antlr4` (JS) and `antlr4ng` (native TS), and `antlr4ng` provides actual TypeScript generics on visitor/listener methods.

`antlr4ng` is a maintained, TypeScript-native fork that tracks ANTLR4 releases, provides proper typings, and has become the de facto standard for TypeScript ANTLR4 projects. The plan lists `antlr4-tool` or `antlr4ts-cli` for grammar generation — neither of these generates code compatible with `antlr4ng`. The correct tool is `antlr4ng-cli`.

**Risk:** Choosing `antlr4` (JS runtime) will produce type-unsafe code throughout the parser facade, every inspection visitor, and the symbol walker. The generated code from `antlr4ts-cli` will not work with the `antlr4` runtime. The dependency table is internally inconsistent: the listed generator tools produce code for a different runtime than the listed runtime package.

**Recommendation:** Replace the dependency table entry. Primary runtime: `antlr4ng` (latest). Generator: `antlr4ng-cli`. Remove `antlr4` and `antlr4ts-cli` from the plan entirely. The "spike" recommended by Round 1 is unnecessary — this is a known decision with a clear answer. Document the rationale.

---

### A2. [CRITICAL] The 5-Tool MCP Surface Has a Fundamental Design Flaw: `vba/inspect` and `vba/inspect-file` Are Redundant

**Description:** The plan defines 5 tools. Two of them — `vba/inspect` (accepts a `code` string) and `vba/inspect-file` (accepts a `path`) — do the same thing with different input sources. This is not a matter of convenience; it is a design flaw that doubles the tool handler surface, doubles the test surface, and creates a maintenance obligation to keep two code paths in sync.

From the MCP client's perspective (Claude Code), the client already has the file contents in its context when it decides to call an inspection tool. It can read a file itself. The only value `vba/inspect-file` adds over `vba/inspect` is avoiding a file read in the client — a trivial I/O operation that the client already performs routinely.

Meanwhile, `vba/inspect-file` introduces the entire security attack surface that Round 1 flagged (path traversal, symlink attacks, encoding detection). The `vba/inspect` tool has none of these problems because it operates on a string.

**Risk:** The project maintains two parallel code paths (string-based and file-based) that must produce identical inspection results. Any divergence is a subtle bug. The file-based path requires a security hardening effort (rootDir, path validation, encoding detection, symlink resolution) that the string-based path does not need.

**Recommendation:** Eliminate `vba/inspect-file`. Keep `vba/inspect` (string input) and `vba/inspect-workspace` (directory scan). The workspace tool handles batch file processing; the inspect tool handles single-file/snippet inspection. Claude Code can read the file and pass contents. This reduces the tool surface to 4 tools, eliminates the single-file security surface, and halves the tool handler maintenance. If file-path inspection is later deemed necessary, it can be added as a thin wrapper — but the plan should not start with two redundant tools.

---

### A3. [CRITICAL] The Inspection Runner Will Not Scale to 139 Inspections Without a Categorized Execution Model

**Description:** Round 1 flagged listener-fusion as a performance concern. I am elevating this to a CRITICAL architectural finding because the problem is not just performance — it is correctness. The plan describes three inspection types (parse-tree, declaration, reference) that depend on different infrastructure:
- Parse-tree inspections need only the parse tree.
- Declaration inspections need the parse tree + symbol table.
- Reference inspections need the parse tree + symbol table + reference resolution.

The runner must know which inspections require which infrastructure to avoid: (a) constructing a symbol table when only parse-tree inspections are active, (b) running reference inspections when the symbol table is incomplete (single-file mode with no workspace), (c) returning incorrect "no findings" results for declaration inspections when the symbol table was not built.

The plan describes the runner as a flat list executor. There is no concept of inspection prerequisites, no dependency declaration, and no conditional infrastructure construction. This is not a performance optimization — it is a correctness requirement. A reference inspection that runs without a resolved symbol table will silently produce no findings rather than erroring.

**Risk:** Users call `vba/inspect` on a code string. The runner executes all 139 inspections. Declaration and reference inspections silently return zero findings because there is no workspace context for symbol resolution. The user sees only parse-tree findings and believes their code is clean of the other 80 categories. This is a silent correctness failure.

**Recommendation:** The runner must implement tiered execution: (1) classify each inspection by its required infrastructure tier (parse-tree, declaration, reference); (2) before executing, determine the available infrastructure for the current call context (string input = parse-tree only; single file = parse-tree + local declarations; workspace = all tiers); (3) skip inspections whose required tier exceeds the available infrastructure AND communicate this to the caller via a `skippedInspections` field or equivalent. This is an architectural decision that must be made before any inspections are written.

---

### A4. [HIGH] The Phase Order Is Suboptimal — Phase 3 (Symbol Resolution) Should Be Phase 2

**Description:** The plan orders phases as: (1) Foundation, (2) Parse-Tree Inspections (~51), (3) Symbol Resolution, (4) Declaration/Reference Inspections (~80), (5) Workspace, (6) Packaging. The rationale is "ship syntax-based inspections before tackling symbol resolution" — incremental value delivery.

This sounds reasonable but creates a structural problem. The inspection framework (base classes, runner, registry) is designed in Phase 2 against parse-tree inspections only. When Phase 3 introduces symbol resolution, the framework must be extended to support declaration and reference inspections. When Phase 4 implements those inspections, they will stress the framework in ways that Phase 2 never tested.

A better order: (1) Foundation, (2) Inspection Framework + Symbol Resolution (framework validated against all three tiers with 1-2 inspections per tier), (3) All Parse-Tree Inspections, (4) All Declaration/Reference Inspections, (5) Workspace + Integration, (6) Packaging.

This front-loads the hard architectural work (symbol resolution, multi-tier runner) and validates the framework against all three inspection types before scaling to 139 inspections. The current order risks discovering in Phase 3-4 that the framework designed in Phase 2 cannot accommodate the harder inspection types.

**Risk:** The Phase 2 framework is designed for the easiest inspection type. Phases 3-4 discover it needs substantial rework, requiring changes that touch all 51 already-written parse-tree inspections.

**Recommendation:** Restructure phases: combine the inspection framework design with symbol resolution (both are infrastructure). Validate with a small number of inspections from each tier. Then batch-implement inspections by type. The current order optimizes for early demo-ability at the cost of framework stability.

---

### A5. [HIGH] The Preprocessor Grammar (`VBAPreprocessorParser.g4`) Is Listed But Never Integrated

**Description:** Section 2.1 lists three grammar files: `VBALexer.g4`, `VBAParser.g4`, and `VBAPreprocessorParser.g4`. The preprocessor grammar is never referenced in any phase, any issue, or any component. `parser/preprocessor.ts` is described as a "stub." Round 1 flagged the stub as a maintenance concern. The deeper problem is that the plan includes a grammar file in the project structure that is never generated, never tested, and never used.

If the preprocessor grammar is not going to be used, it should not be in the project — dead grammar files in `grammar/` mislead contributors into thinking preprocessor support exists. If it is going to be used, there must be an issue to generate its TypeScript code, integrate it with the main parser, and test it.

**Risk:** A contributor sees the grammar file, assumes preprocessor support exists, and reports bugs when `#If` blocks produce unexpected results. Or a contributor tries to generate it and discovers it is incompatible with the TypeScript target.

**Recommendation:** Either (a) remove `VBAPreprocessorParser.g4` from the project structure and document that conditional compilation is unsupported, or (b) add explicit issues in Phase 1 or Phase 2 to generate, integrate, and test the preprocessor grammar. Do not ship a grammar file that nothing uses.

---

### A6. [HIGH] The `host` Parameter Design Is Fundamentally Flawed

**Description:** The `host` parameter appears on `vba/inspect`, `vba/inspect-file`, and `vba/inspect-workspace` with values like `excel`, `project`, `generic`. It controls which host-specific inspections run. The plan assumes the caller (Claude Code / the user) knows which host context applies. This assumption is wrong in the general case.

VBA modules do not contain metadata indicating their host application. A `.bas` file exported from Excel looks identical to one exported from Access. The `host` parameter shifts the burden of host detection to the caller, who has no reliable way to determine it. In practice, users will either: (a) not pass `host` (defaulting to generic, missing Excel-specific findings), (b) always pass `excel` (getting false positives for non-Excel code), or (c) guess wrong.

More fundamentally, some VBA code is host-agnostic. A `MsgBox` call works in Excel, Access, and Word. An `ActiveSheet` reference only works in Excel. The current design is binary (one host per call) rather than additive (this code uses these host libraries).

**Risk:** The host parameter creates a false sense of precision. Users who pass the wrong host get misleading results. Users who omit it miss host-specific inspections entirely. The parameter adds complexity to every tool call for marginal benefit.

**Recommendation:** Replace the `host` parameter with a `hostLibraries` array parameter (e.g., `["excel", "access"]`) that specifies which host object libraries are available. Default to `["excel"]` since Excel is the dominant VBA host. Add heuristic host detection to `vba/inspect-file` and `vba/inspect-workspace` based on file content (e.g., presence of `Application.Worksheets` implies Excel). The binary single-host model is too coarse.

---

### A7. [MEDIUM] The `quickFix.replacement` Field Assumes Single-Location Fixes

**Description:** The `InspectionResult` schema provides one `quickFix` object with one optional `replacement` string. Many real-world quick fixes require multiple edits: adding `Option Explicit` requires inserting a line at the top of the module AND potentially declaring variables throughout the body; renaming a procedure requires changes at the declaration site and all call sites; extracting a magic number requires adding a `Const` declaration and replacing the literal.

The Rubberduck v2 quick-fix system supports multi-edit fixes. The plan's schema does not. This is not a future concern — it affects the very first inspections in Phase 2 (e.g., `OptionExplicitInspection` needs a multi-location fix).

**Risk:** Quick fixes are either limited to trivially simple replacements (losing most of the value of the Rubberduck port) or the schema must be redesigned mid-implementation to support multi-edit fixes, breaking all existing quick-fix tests and consumer expectations.

**Recommendation:** Redesign `quickFix` before Phase 2:
```typescript
quickFix?: {
  description: string;
  edits?: Array<{
    location: { startLine: number; startColumn: number; endLine: number; endColumn: number; };
    newText: string;
  }>;
};
```
This is compatible with LSP's `TextEdit` model and supports both single and multi-location fixes.

---

### A8. [MEDIUM] No Consideration of VBA7 vs VBA6 Grammar Differences

**Description:** The plan references "the VBA grammar" as a monolithic artifact. In reality, VBA has evolved: VBA6 (Office 2003 and earlier) and VBA7 (Office 2010+) have differences, most notably the `PtrSafe` keyword for 64-bit Declare statements and `LongPtr`/`LongLong` types. The Rubberduck grammar handles both, but the plan never discusses which VBA version(s) the parser supports, whether version-specific inspections exist, or how version differences affect correctness.

**Risk:** Inspections that flag `PtrSafe` as unnecessary (or its absence as an error) depend on knowing the target VBA version. Without this context, version-specific inspections produce false positives for half of all users.

**Recommendation:** Document explicitly that the parser supports VBA7 syntax (the current standard). If VBA6 compatibility is needed, add a `vbaVersion` parameter or detect it from `#If VBA7` conditional blocks. At minimum, acknowledge the version question in the plan.

---

## TEST QUALITY (PRIMARY)

### T1. [CRITICAL] The "Claude Code via tmux" Integration Test (Issue #30) Is Untestable in CI and Couples to a Specific Client

**Description:** Issue #30 says "Load MCP in Claude Code via tmux, run inspections on test VBA files, validate results." Round 1 flagged this as "not a strategy." I am going further: this test is actively harmful to the project because it couples the test suite to a specific MCP client (Claude Code), a specific execution environment (tmux), and a specific AI model's behavior (which is nondeterministic).

Claude Code is an AI assistant. Its behavior when calling MCP tools depends on the prompt, the model version, the conversation history, and random sampling. A test that "sends VBA code inspection requests" through Claude Code is testing Claude's tool-calling behavior, not the MCP server's correctness. When this test fails, the failure could be in Claude's decision to call the tool, Claude's parameter formatting, Claude's interpretation of the result, tmux's terminal handling, or the MCP server itself. The signal-to-noise ratio is near zero.

Furthermore, this test cannot run in CI. It requires a Claude Code session (authenticated, billed), tmux, and terminal interaction. It is a manual demo dressed up as a test.

**Risk:** Issue #30 consumes development time producing a test that provides no automated quality signal and cannot be included in any CI pipeline. It creates a false sense of coverage — "we tested with Claude Code" — when what was actually tested is one specific interaction on one specific day.

**Recommendation:** Delete Issue #30 entirely. Replace it with a proper MCP protocol conformance test: spawn the server as a child process, send raw JSON-RPC messages over stdio, and validate responses. This tests the actual contract (MCP protocol) without depending on any specific client. The test should be in `mcp-integration.test.ts` (Issue #29) and run in CI. If a Claude Code demo is desired for documentation purposes, make it a README example, not a test.

---

### T2. [CRITICAL] No Test Strategy for the Translation Correctness of Ported Inspections

**Description:** The plan's central premise is "port Rubberduck v2's inspection logic from C# to TypeScript." The test strategy validates each TypeScript inspection against VBA fixtures. What it does not validate is whether the TypeScript implementation is a faithful translation of the C# original.

Translation bugs are the dominant risk in a port project. A C# `switch` statement translated to TypeScript `switch` may behave differently due to fall-through semantics. A C# LINQ query translated to TypeScript array methods may have different evaluation order. A C# `string.Contains` comparison may be case-sensitive where the TypeScript equivalent is not (or vice versa). These bugs produce inspections that work on test fixtures but diverge from Rubberduck's behavior on real-world code.

The plan has no mechanism to detect translation divergence. There is no cross-reference to Rubberduck's own test suite, no import of Rubberduck's test fixtures, and no comparison of results between the C# and TypeScript implementations.

**Risk:** The project ships 139 inspections that behave differently from Rubberduck in subtle ways. Users familiar with Rubberduck expect consistent behavior and report bugs. Diagnosing each bug requires going back to the C# source, understanding the original logic, and comparing it to the TypeScript translation — an expensive per-bug process.

**Recommendation:** For each inspection ported, import at least one test case from Rubberduck's own test suite (`RubberduckTests/Inspections/`). This provides a ground-truth fixture that the original C# implementation is known to handle correctly. If the TypeScript implementation produces different results on the same fixture, the translation is wrong. This is the single highest-value testing investment the project can make.

---

### T3. [HIGH] The Plan Has No Strategy for Testing Inspection Interactions

**Description:** Inspections are tested individually: "EmptyIfBlock" gets a fixture, "ObsoleteLet" gets a fixture. But real VBA code triggers multiple inspections simultaneously. The plan has no tests for:
- Two inspections that fire on overlapping source ranges (e.g., an empty `If` block that also uses obsolete comment syntax inside the empty block).
- An inspection whose quick-fix introduces a violation detected by another inspection (fix `ObsoleteLet` by removing `Let`, which creates an `ImplicitLetCoercion` finding).
- An inspection that should NOT fire when another inspection already covers the same defect at a higher severity.

These interactions are where real-world linters produce confusing output — five findings for one logical problem, or contradictory fix suggestions.

**Risk:** Users see cascading or contradictory findings. They fix one finding and three new ones appear. The tool feels unreliable even when each individual inspection is correct.

**Recommendation:** Add a "multi-inspection interaction" test suite that exercises at least 10 realistic VBA snippets triggering 3+ inspections each. Assert the exact set of inspections that fire (not just count) and verify no contradictory quick-fix suggestions. This suite should be a Phase 2 exit gate, not a Phase 5 afterthought.

---

### T4. [HIGH] No Boundary Testing for the ANTLR4 Grammar Itself

**Description:** The plan treats the grammar as a trusted, pre-validated artifact ("well-tested in Rubberduck"). This is false confidence. The grammar was tested in Rubberduck's Java/C# ANTLR4 runtime. ANTLR4 runtimes can produce different parse trees for the same grammar due to:
- Unicode handling differences between runtimes.
- Case-insensitive matching implementation differences (VBA is case-insensitive; ANTLR4 runtimes differ in how they handle `options { caseInsensitive = true; }`).
- Token channel handling differences.
- Error recovery strategy differences.

The grammar may parse correctly in the C#/Java runtime but fail or produce a different tree in the TypeScript runtime. This is a known ANTLR4 cross-runtime issue.

**Risk:** The grammar parses VBA differently in TypeScript than in C#/Java. Inspections that depend on specific parse tree shapes produce wrong results. This is undetectable by inspection-level tests if the test fixtures happen to avoid the grammar divergence points.

**Recommendation:** Port at least 50 VBA parse-test fixtures from Rubberduck's `RubberduckTests/Grammar/` and verify the TypeScript-generated parser produces equivalent parse trees. Focus on known divergence areas: case-insensitive identifiers, line continuations, colon-separated statements, string literals with embedded quotes, and `#` date literals. This is a Phase 1 exit gate.

---

### T5. [HIGH] The "Clean" Fixture Strategy Is Insufficient

**Description:** The plan includes `test/fixtures/clean/well-written-module.bas` — a single file that should produce zero diagnostics. One clean file is not enough. The clean fixture set must cover every syntactic construct that inspections examine, or new inspections will accidentally fire on "clean" code without any test catching the regression.

**Risk:** A Phase 4 inspection (e.g., `HungarianNotation`) fires on variable names in the clean fixture. No test catches this because the clean fixture was written in Phase 2 before that inspection existed.

**Recommendation:** The clean fixture set should include at minimum: (a) a module using every VBA statement type (Sub, Function, Property Get/Let/Set, Enum, Type, Declare, Event); (b) a module with all common patterns that inspections check (error handling, loops, conditionals, `With` blocks); (c) a module with Excel-host-specific code that is correctly written. Each must produce zero diagnostics when all 139 inspections run. Add a test that runs the full inspection set against every clean fixture as a global regression gate.

---

### T6. [MEDIUM] No Fuzz Testing or Property-Based Testing Strategy

**Description:** The plan's test strategy is entirely example-based: hand-written VBA fixtures with hand-written expected results. For a parser that accepts arbitrary string input, example-based testing is necessary but not sufficient. Property-based testing (e.g., "any valid VBA that parses without error should not crash any inspection") and fuzz testing (e.g., "random byte sequences should not crash the server") are absent.

**Risk:** Edge cases in the ANTLR4 grammar or inspection visitors cause crashes on inputs that no hand-written fixture covers. The parser operates on untrusted input (user code); crash-inducing inputs are a reliability and security concern.

**Recommendation:** Add a property-based test (using `fast-check` or similar): generate random strings that look like VBA (template-based generation), parse them, run inspections, and assert no uncaught exceptions. Add a small fuzz corpus of known-tricky VBA patterns (empty files, files with only comments, files with only preprocessor directives, files with maximum-length lines). These do not need to validate inspection correctness — they only need to confirm no crashes.

---

## MAINTAINABILITY

### M1. [HIGH] The Plan Contains No In-Code Documentation Standard

**Description:** The plan describes README, SPEC.md, ATTRIBUTION.md, and inspection-metadata.json as documentation artifacts. It says nothing about documentation within the source code itself. For a project porting 139 inspections from a different language, in-code documentation is critical:
- Each inspection file should document what Rubberduck C# file it was ported from.
- Each non-obvious translation decision should have a comment explaining why the TypeScript differs from the C#.
- The parser facade should document which ANTLR4 runtime APIs it wraps and why.
- The symbol walker should document its pass strategy.

Round 1's maintainability review focused on structural issues (registry, metadata, test files). It missed that the code itself will be unreadable to a future contributor who does not have the Rubberduck C# source open in a second window.

**Risk:** A developer who inherits this project in two years has no way to understand why an inspection was implemented a particular way without reverse-engineering both the TypeScript code and the original C# source. Bug fixes become guesswork.

**Recommendation:** Add an in-code documentation standard to the plan: (1) every inspection file must have a header comment citing the Rubberduck source file path and commit hash; (2) every non-trivial translation decision must have an inline comment; (3) the parser facade must document its API contract with JSDoc. This is a per-file requirement, not a per-project document.

---

### M2. [HIGH] The `inspection-metadata.json` Approach Creates a Localization Dead End

**Description:** Round 1 recommended moving metadata into inspection classes. I want to challenge both the plan's approach AND Round 1's recommendation by pointing out what neither addressed: the `.resx` files in Rubberduck exist because Rubberduck supports localization (multiple languages for inspection descriptions). The plan extracts English strings into a JSON file and discards the localization infrastructure.

If this project ever needs to support non-English descriptions (e.g., for VBA developers in Japan, Germany, or Brazil — large VBA user populations), the single-language JSON or class-embedded strings will need to be refactored into a proper i18n system.

**Risk:** Low immediate risk, but a localization retrofit after 139 inspections embed English strings in class properties is expensive. The `.resx` extraction discards structure that would have been valuable.

**Recommendation:** Use a key-based metadata approach from the start. Each inspection references a key (e.g., `"EmptyIfBlock.description"`); the key resolves against a locale-specific string table. Ship only English, but the infrastructure supports adding locales without touching inspection code. This is the architecture Rubberduck chose for a reason.

---

### M3. [MEDIUM] No Contribution Guide or Inspection Authoring Guide Planned

**Description:** The plan targets 139 inspections. After the initial implementation, contributors may want to add project-specific or community-contributed inspections. The plan has no provision for an inspection authoring guide — a document explaining how to add a new inspection, what base class to extend, what tests to write, how to register it, and what metadata to provide.

**Risk:** Without a guide, every new inspection requires reading existing inspection source code and inferring the pattern. Inconsistent implementations accumulate.

**Recommendation:** Add an `CONTRIBUTING.md` section or a `docs/authoring-inspections.md` as a Phase 2 deliverable (when the framework stabilizes). Include a template inspection file and a template test file.

---

## SECURITY

### S1. [HIGH] Symlink Following in Workspace Scanner

**Description:** Round 1 flagged path traversal via user-supplied paths. A subtler attack vector exists: symlinks within the workspace directory. If a scanned workspace contains a symlink pointing to `/etc/passwd` or `~/.ssh/id_rsa`, and the file has a `.bas` extension (or the symlink target is read before extension filtering), the server reads the file contents and attempts to parse them. Parse errors will include fragments of the file content in error messages, potentially leaking sensitive data.

Even without extension manipulation, a symlink named `secrets.bas` pointing to a sensitive file will be read, parsed, and its content potentially included in inspection results (e.g., in quick-fix replacement text or error messages that include source excerpts).

**Risk:** Data exfiltration via symlink-to-sensitive-file within a workspace directory that passes the rootDir containment check.

**Recommendation:** Resolve all file paths with `fs.realpath()` before reading and re-validate that the resolved path is still within the rootDir boundary. Add a configuration option to disable symlink following entirely. Test with a fixture containing a symlink to a file outside the workspace.

---

### S2. [HIGH] VBA Code Content Leaking Into MCP Error Responses

**Description:** Round 1 flagged error messages leaking file system paths. A more significant leakage vector exists: VBA code content leaking into error responses. When the ANTLR4 parser fails, the error message includes the offending token text and surrounding context. If the VBA file contains sensitive data (embedded credentials, API keys in string literals — common in legacy VBA codebases), the parse error message will include that data in the MCP response, which is then visible to the AI model and potentially logged.

**Risk:** Parse errors on VBA files containing embedded secrets expose those secrets in MCP responses. The AI model then has the secrets in its context, where they could be inadvertently included in generated code or conversation output.

**Recommendation:** Sanitize ANTLR4 error messages before including them in MCP responses. Truncate token text to a maximum length (e.g., 50 characters). Strip string literal content from error messages. Never include raw source lines in error responses — include only line numbers and generic descriptions.

---

### S3. [MEDIUM] No Content Security Policy for `quickFix.replacement`

**Description:** Round 1 flagged `quickFix.replacement` as "unsanitized code generation." The specific risk that was not articulated: if a quick-fix replacement contains VBA code that, when applied, creates a new security vulnerability in the VBA project. For example, a quick fix that replaces `Dim x As Integer` with `Dim x As Long` is safe. But a quick fix that inserts an `On Error Resume Next` (as part of an error-handling suggestion) suppresses all errors in the procedure, which is itself a known anti-pattern.

**Risk:** Quick-fix suggestions introduce VBA anti-patterns. Users who trust the tool's suggestions degrade their own code.

**Recommendation:** Audit every quick-fix replacement in the metadata extraction to ensure no quick fix introduces a pattern that another inspection would flag. This is a logical consistency check: the fix for inspection A should not trigger inspection B. Add this as a test.

---

## PERFORMANCE

### P1. [HIGH] The Plan Assumes ANTLR4 TypeScript Performance Is Comparable to Java/C# — It Is Not

**Description:** Round 1 flagged baseline justification as missing. I want to be more specific about the magnitude of the problem. Published ANTLR4 benchmarks show the JavaScript/TypeScript runtime is 5-15x slower than the Java runtime for complex grammars. The VBA grammar is complex (~800 rules with case-insensitive matching). Rubberduck's C# parser (using the C# ANTLR4 runtime, which is 2-3x slower than Java) takes measurable time on large VBA modules.

The TypeScript runtime, being 5x slower than C# on a grammar this size, means cold-parse of a 500-line VBA module could easily take 200-500ms — consuming or exceeding the entire 100ms single-file budget before any inspection runs.

The `antlr4ng` runtime (recommended in A1) is faster than the official JS runtime but still significantly slower than Java/C#. The plan's performance targets appear to be derived from Rubberduck's C#-runtime performance, which is not transferable.

**Risk:** The 100ms single-file target is unachievable for non-trivial files without caching. The plan does not acknowledge this, creating a Phase 5 surprise.

**Recommendation:** Revise the performance targets based on a realistic assessment: (a) first-parse of a 500-line file: 300-500ms (acceptable with caching); (b) cached re-inspection: <50ms; (c) 50-file workspace cold parse: 15-25s; (d) 50-file workspace cached: <3s. If these targets are unacceptable, the architecture must include mandatory caching and warm-up, not optional optimization.

---

### P2. [MEDIUM] No Analysis of MCP Protocol Overhead

**Description:** Every tool call incurs MCP protocol overhead: JSON-RPC serialization, stdio pipe I/O, and MCP SDK processing. For workspace scans that return hundreds of results, the serialization of the response JSON may itself be significant. The plan's performance targets (100ms per file) appear to measure only parse + inspect time, not the full round-trip including MCP overhead.

**Risk:** The server meets its internal 100ms target but the client-perceived latency is 200ms+ due to protocol overhead and result serialization.

**Recommendation:** Define performance targets as client-perceived round-trip time, not internal processing time. Measure and document the MCP protocol overhead in Phase 1. If serialization of large result sets is a bottleneck, consider streaming results or paginating workspace scan output.

---

## PRODUCTION READINESS

### PR1. [CRITICAL] No Graceful Shutdown or Signal Handling

**Description:** The plan describes server startup but never mentions shutdown. An MCP server running as a child process of Claude Code receives SIGTERM when Claude Code exits or restarts. If the server is mid-parse of a large workspace scan, an unhandled SIGTERM will produce a hard kill with no cleanup. If the server holds file locks (unlikely but possible with caching), those locks will not be released.

More importantly: if the server has a bug that causes it to hang (infinite loop in an inspection, deadlocked parser), there is no watchdog, no self-diagnostic, and no automatic restart. The plan mentions no supervisor process, no heartbeat, and no liveness check.

**Risk:** A hung MCP server silently disables all VBA inspection capabilities for the Claude Code session. The user has no indication that the tools are unavailable until they try to use one and it times out.

**Recommendation:** Add signal handling in Phase 1: (1) `SIGTERM` and `SIGINT` trigger graceful shutdown (finish current tool call, close stdio, exit); (2) add a per-call timeout (e.g., 30 seconds) after which the server responds with a timeout error rather than hanging indefinitely; (3) document the expected restart behavior (Claude Code restarts the MCP process automatically on crash — verify this assumption).

---

### PR2. [HIGH] No Inspection Catalog Versioning

**Description:** The plan defines success as ">=130 of 139 inspections implemented." But there is no versioning of the inspection catalog itself. When the server reports findings to a client, the client does not know: which inspections were available, which version of the inspection logic ran, or whether the inspection set has changed since the last call. The `vba/list-inspections` tool returns a catalog, but the catalog has no version identifier.

This matters because inspection behavior will change over time (bug fixes, severity adjustments, new inspections). A user who sees different results on the same code after a server update has no way to determine whether the difference is due to code changes or inspection changes.

**Risk:** Result instability across server updates with no mechanism for the client to detect or understand the change.

**Recommendation:** Add an `inspectionCatalogVersion` field to the `vba/list-inspections` response (a semver string or content hash). Add an `engineVersion` field to every tool response. This enables clients to detect when results may differ due to tool updates rather than code changes.

---

### PR3. [HIGH] No Consideration of MCP Server Lifecycle Management

**Description:** The plan assumes the MCP server is a long-lived process. But Claude Code's MCP lifecycle is: spawn on first tool call, keep alive for the session, kill on session end. If the server crashes mid-session, Claude Code may or may not restart it. The plan has no restart resilience strategy, no state recovery, and no assumption documentation about the MCP lifecycle model.

If the server maintains any state (parse cache, symbol table, loaded config), a crash-and-restart loses all state. The first tool call after restart incurs cold-start penalties. If the server was mid-workspace-scan, the scan results are lost.

**Risk:** Crash-restart cycles degrade performance (cold start each time) and lose intermediate results. Users experience inconsistent behavior between fresh sessions and long-running sessions.

**Recommendation:** Design the server to be crash-resilient: (1) no critical state that cannot be reconstructed from inputs; (2) cache is a performance optimization, not a correctness requirement; (3) document the expected lifecycle in SPEC.md. Add a test that verifies the server produces correct results immediately after a cold start with no warm cache.

---

### PR4. [MEDIUM] The Plan Has No Telemetry or Usage Metrics Design

**Description:** After shipping 139 inspections, the team will need to know: which inspections are most commonly triggered, which produce the most false positives (via @Ignore usage), which are never triggered (possibly broken), and how long workspace scans take in practice. The plan has no provision for collecting this data.

**Risk:** Without usage data, inspection quality improvements are guesswork. Broken inspections that never fire are never discovered. Performance regressions in real usage are invisible.

**Recommendation:** Add optional, local-only telemetry: a counter per inspection of how many times it fired, how many times it was suppressed via @Ignore, and aggregate timing data. Store in a local file (not transmitted). Expose via a `vba/stats` tool or include in `vba/list-inspections` output. This is a Phase 5 or Phase 6 feature, but the data collection hooks should be designed into the runner from Phase 2.

---

## Summary Table

| # | ID | Finding | Severity | Focus Area |
|---|-----|---------|----------|------------|
| 1 | A1 | `antlr4` is the wrong package; `antlr4ng` should be primary | CRITICAL | Architecture |
| 2 | A2 | `vba/inspect` and `vba/inspect-file` are redundant; eliminate one | CRITICAL | Architecture |
| 3 | A3 | Runner needs tiered execution model, not flat list | CRITICAL | Architecture |
| 4 | T1 | "Claude Code via tmux" integration test is untestable in CI | CRITICAL | Test Quality |
| 5 | T2 | No strategy for validating translation correctness against Rubberduck originals | CRITICAL | Test Quality |
| 6 | PR1 | No graceful shutdown, signal handling, or call timeout | CRITICAL | Production Readiness |
| 7 | A4 | Phase order should front-load symbol resolution with framework | HIGH | Architecture |
| 8 | A5 | Preprocessor grammar listed but never integrated | HIGH | Architecture |
| 9 | A6 | `host` parameter design is too coarse; binary host model is wrong | HIGH | Architecture |
| 10 | A7 | `quickFix.replacement` assumes single-location fixes | MEDIUM | Architecture |
| 11 | A8 | No VBA6 vs VBA7 grammar difference acknowledgment | MEDIUM | Architecture |
| 12 | T3 | No tests for inspection interactions / cascading findings | HIGH | Test Quality |
| 13 | T4 | No boundary testing for ANTLR4 grammar cross-runtime divergence | HIGH | Test Quality |
| 14 | T5 | Single "clean" fixture is insufficient for regression protection | HIGH | Test Quality |
| 15 | T6 | No fuzz testing or property-based testing strategy | MEDIUM | Test Quality |
| 16 | M1 | No in-code documentation standard for ported inspections | HIGH | Maintainability |
| 17 | M2 | Metadata extraction discards Rubberduck's localization infrastructure | HIGH | Maintainability |
| 18 | M3 | No inspection authoring guide planned | MEDIUM | Maintainability |
| 19 | S1 | Symlink following in workspace scanner enables data exfiltration | HIGH | Security |
| 20 | S2 | VBA code content (including secrets) leaks into error responses | HIGH | Security |
| 21 | S3 | Quick fixes may introduce patterns flagged by other inspections | MEDIUM | Security |
| 22 | P1 | TypeScript ANTLR4 runtime is 5-15x slower than Java; targets unrealistic | HIGH | Performance |
| 23 | P2 | MCP protocol overhead not included in performance targets | MEDIUM | Performance |
| 24 | PR2 | No inspection catalog versioning | HIGH | Production Readiness |
| 25 | PR3 | No crash-restart resilience strategy | HIGH | Production Readiness |
| 26 | PR4 | No telemetry or usage metrics design | MEDIUM | Production Readiness |

---

## Comparison with Round 1

Round 1 identified 80+ findings, many of which were thorough and well-targeted. The following areas were either missed or insufficiently challenged by Round 1:

1. **ANTLR4 runtime selection** — Round 1 said "do a spike." The answer is already known: use `antlr4ng`. The dependency table is also internally inconsistent (generator tool does not match runtime).
2. **Tool surface redundancy** — Round 1 accepted all 5 tools as given. The `vba/inspect-file` tool is redundant with `vba/inspect` and introduces the entire file-system security attack surface.
3. **Tiered inspection execution** — Round 1 mentioned listener-fusion for performance. The deeper issue is correctness: declaration/reference inspections silently return nothing when their required infrastructure is unavailable.
4. **Translation correctness** — Round 1 focused on test coverage patterns but missed the fundamental question: how do you verify a port is faithful to the original?
5. **Cross-runtime grammar divergence** — Round 1 treated the grammar as a trusted artifact. ANTLR4 runtimes produce different parse trees for the same grammar.
6. **Quick-fix schema limitations** — Neither the plan nor Round 1 addressed that the quick-fix model is too simple for multi-edit fixes.
7. **Host parameter design** — Round 1 did not challenge the binary host model.
8. **Server lifecycle management** — Round 1 had no findings about shutdown, crash recovery, or MCP lifecycle.

---

## Overall Assessment

**Fail — revise before implementation begins.**

Six CRITICAL findings must be resolved:
1. The ANTLR4 runtime choice must be changed to `antlr4ng` (A1).
2. The tool surface must be reconsidered — `vba/inspect-file` should be eliminated or justified (A2).
3. The inspection runner must support tiered execution based on available infrastructure (A3).
4. Issue #30 (Claude Code via tmux) must be replaced with automated protocol testing (T1).
5. The testing strategy must include translation fidelity validation against Rubberduck's own test cases (T2).
6. Signal handling and call timeouts must be added to the server design (PR1).

Combined with Round 1's unresolved CRITICAL findings (TBD catalog items, v2/v3 declaration model mismatch, multi-pass symbol resolution, no caching architecture, no error handling strategy, unrestricted file system access, parser DoS), the plan has 13+ CRITICAL findings across two review rounds. The plan is architecturally ambitious and well-intentioned, but it is not ready to execute.
