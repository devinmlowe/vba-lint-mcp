# NAR Round 1 — Architecture Review

**Reviewer:** Architecture Dimension
**Date:** 2026-03-19
**Target:** PLAN.md (Implementation Plan)
**Weight:** PRIMARY

---

## Executive Summary

The plan is structurally sound at the macro level — a layered MCP server with a parse-tree-first delivery strategy is the right decomposition for this problem. However, the plan defers approximately 45 inspection definitions to TBD, misrepresents phase sizes as balanced when they are not, and contains several underspecified architectural joints (registry coupling, symbol resolution multi-pass requirements, ANTLR runtime selection, and the cross-version declaration model) that carry genuine implementation risk. The plan is not ready to execute without resolving the critical and high findings below.

---

## Findings

### [CRITICAL] TBD Items Constitute ~32% of the Inspection Catalog

**Description:** Items 85–100 (16 declaration inspections) and 110–139 (30 reference inspections) are explicitly labeled "will be fully enumerated during Phase 4 implementation." This is not a planning artifact — it means 46 of the 139 target inspections (33%) have no defined scope, no issue ticket, and no effort estimate at plan time.

**Risk:** Phase 4 scope is unknowable. If the un-enumerated inspections include structurally complex cases (e.g., cross-module reference tracking, interface resolution, COM type library queries), Phase 4 could easily be 3x the size currently implied. The "~80 inspections" claim for Phase 4 cannot be validated.

**Recommendation:** Before this plan is approved, enumerate every inspection in the v2 catalog by name. Classify each as parse-tree, declaration, or reference. Assign a complexity estimate (low/medium/high). Collapse TBD blocks entirely. If enumeration reveals Phase 4 is too large, split it.

---

### [CRITICAL] Declaration Model Version Mismatch: v3 Architecture, v2 Inspection Logic

**Description:** Section 9.1 explicitly states that `src/symbols/declaration.ts` is derived from Rubberduck v3's `Declaration.cs`, while all inspections are ported from Rubberduck v2. These are architecturally incompatible. The v3 declaration model was a substantial rewrite that changed how scope, module identity, and member resolution work. Inspections written against the v2 model will have assumptions (about declaration shape, reference resolution APIs, and symbol walker behavior) that do not hold against a v3-derived declaration hierarchy.

**Risk:** Inspections that appear to work against simple test fixtures may silently produce incorrect results (false negatives, wrong locations) against real-world VBA. The mismatch may not surface until Phase 4, after Phase 3 has committed to the v3 model.

**Recommendation:** Make an explicit architectural decision: port the declaration model from v2, or audit every v2 inspection to confirm it is compatible with a v3-derived model before Phase 3 begins. Document the decision and its implications in the plan. Do not leave "port from v3" and "inspections from v2" as simultaneous unchallenged assumptions.

---

### [CRITICAL] Symbol Resolution Requires Multi-Pass; Single-Pass Plan Is Underspecified

**Description:** The plan describes Phase 3 as four issues: declaration model, symbol walker, declaration finder, and scope resolution. This treats symbol resolution as a linear pipeline. VBA symbol resolution is not single-pass: forward references between procedures, cross-module public member visibility, and interface implementation all require at least two passes (collect declarations, then resolve references against the collected set). The plan contains no acknowledgment of this requirement, no design for how multi-pass is coordinated, and no discussion of how incomplete resolution is handled (e.g., a reference to a symbol in a file not included in the workspace scan).

**Risk:** A naive single-pass implementation will produce incorrect results for any non-trivial VBA project. This will be discovered late, during Phase 4, when declaration inspections begin producing wrong outputs.

**Recommendation:** Add an explicit multi-pass design to Phase 3. Define: (1) what a "first pass" collects, (2) what triggers a second pass, (3) how the declaration finder handles unresolved references, and (4) what the symbol walker does when given a single file vs. a full workspace. This is an architectural decision, not an implementation detail.

---

### [HIGH] ANTLR4 Runtime Choice Is Unresolved and Architecturally Significant

**Description:** Section 2.3 lists both `antlr4` (official JS/TS runtime) and `antlr4ts-cli` as options, with `antlr4ng` mentioned only in the risk register as a fallback. These are not interchangeable. `antlr4ts` is unmaintained and targets ANTLR 4.9; `antlr4` (official) has a TypeScript target but the ergonomics differ significantly from `antlr4ts`; `antlr4ng` is actively maintained and Angular-Signal-free as of 4.13 but introduces different visitor/listener conventions. The generated code shape, visitor API surface, and error listener interface differ between all three. This choice affects every inspection, the parser facade, the error listener, and the symbol walker.

**Risk:** Selecting the wrong runtime means rewriting generated bindings and all visitor code partway through Phase 2 or 3. The plan's risk register calls this out but does not force a decision before implementation begins.

**Recommendation:** Make the runtime selection a Phase 1 exit gate. Write a spike: generate the VBA grammar with each candidate runtime, parse a non-trivial VBA file, and confirm the visitor pattern works. Record the decision in the plan. Do not proceed to Phase 2 with this open.

---

### [HIGH] Registry Pattern Creates Tight Coupling Between Metadata JSON and Inspection Classes

**Description:** Section 2.1 shows `inspection-metadata.json` as a resource file containing names, descriptions, quick-fix text, and (implied) category/severity defaults. Section 2.2's issue #7 describes "auto-discovery of inspection classes" with "metadata loading from JSON." The plan does not define the coupling contract: how does the registry match a JSON entry to an inspection class? By class name string? By a static `id` property? If by class name, renaming a class silently breaks the registry. If by a static property, the property becomes a required contract that the plan never specifies.

**Risk:** The registry is the join point between the metadata plane and the implementation plane. Undefined contracts here will produce subtle bugs — inspections that load but produce no metadata, or metadata entries with no corresponding class that silently succeed with empty results.

**Recommendation:** Specify the registry contract explicitly: (1) how inspection classes are discovered (explicit imports, directory scan, or registration call), (2) the required static interface every inspection class must satisfy, (3) what happens when a metadata entry has no corresponding class, and (4) what happens when a class has no metadata entry.

---

### [HIGH] Inspection Class Hierarchy Is Named But Not Specified

**Description:** Issue #6 references `InspectionBase` and `ParseTreeInspection` as base classes, and the directory structure implies a third category (reference-based). The plan never defines what these abstractions contain: what methods are required, what data they receive, what they must return, and how the runner calls them. Without this contract, implementing 60 inspections in Phase 2 before the abstraction is proven will produce inconsistency across the inspection set.

**Risk:** If `InspectionBase` is underspecified and each inspection author (or each implementation session) fills gaps differently, the inspection set will have an inconsistent API surface. Refactoring 60 inspections post-hoc is expensive.

**Recommendation:** Before writing any inspection, write the complete interface specification for `InspectionBase`, `ParseTreeInspection`, `DeclarationInspection`, and `ReferenceInspection`. Define: constructor signature, required static properties (id, name, category, severity), the `inspect(context)` method signature including what `context` contains, and the return type. Validate with one inspection per type before scaling.

---

### [HIGH] Workspace Model Is Architecturally Absent Until Phase 5

**Description:** The plan ships `vba/inspect-workspace` in Phase 5 (issue #28), but cross-module declaration inspections (Phase 4, issues #20–#26) depend on a multi-file workspace model. The plan says Phase 4 tests "work across multiple files" and "cross-module reference tracking works," but the workspace abstraction to support this is not defined until Phase 5. This is a dependency inversion: Phase 4 consumers require what Phase 5 produces.

**Risk:** Either Phase 4 inspections will only work on single files (making them less useful than promised), or the workspace model will be informally constructed during Phase 4 in a way that conflicts with Phase 5's formal implementation.

**Recommendation:** Define the workspace model explicitly in Phase 3 alongside symbol resolution. A workspace is a named set of parse trees + symbol tables. Phase 4 inspections should take a workspace context. Phase 5 adds the MCP tool that creates a workspace from a directory scan. The abstraction must precede its consumers.

---

### [HIGH] "~60 Parse-Tree Inspections" Claim Does Not Match the Enumerated List

**Description:** The plan repeatedly states "~60 parse-tree inspections" in Phase 2. The actual catalog (section 5.1) lists items 1–51 as parse-tree inspections. That is 51, not ~60. The discrepancy is not rounding — it is 15% off. Similarly, "~50 declaration inspections" covers items 52–84 plus "85–100 TBD," which is 33 enumerated + 16 TBD = 49, but several of the enumerated items overlap with parse-tree work (e.g., `EmptyMethodInspection` appears twice, at item 8 and item 84). "~30 reference inspections" covers items 101–109 (9 enumerated) plus "110–139 TBD" (30 TBD). The total enumerated is 109, not 139.

**Risk:** The plan's numeric claims are used to scope phases and write issue descriptions. If the numbers are wrong, phase effort estimates are wrong. Double-counting `EmptyMethodInspection` suggests the classification has not been rigorously reviewed.

**Recommendation:** Audit the full inspection list against the v2 source catalog. Assign each inspection exactly one category. Remove duplicates. Fix all count claims to match the enumerated list. Do not use "~" as a substitute for a correct count.

---

### [MEDIUM] Parser Facade Abstraction Level Is Not Defined

**Description:** Issue #3 and `src/parser/vba-parser-facade.ts` are described with the API `parseCode(source: string) → ParseTree`. This single-sentence description leaves open: what is `ParseTree`? Is it the raw ANTLR `ParseTree` type (exposing ANTLR internals to callers), a wrapped type, or a serializable JSON structure? Does the facade expose the token stream? Does it expose parse errors as part of the return type or via the error listener only? The `vba/parse` tool must serialize the tree to JSON — does the facade produce a serializable tree, or does serialization happen in the tool layer?

**Risk:** If the facade leaks ANTLR internals, all callers (tools, inspections, symbol walker) become coupled to the ANTLR runtime. Swapping runtimes later requires rewriting every caller, not just the facade.

**Recommendation:** Define the facade's output type explicitly before Phase 1 completes. At minimum: `ParseResult { tree: VBAParseTree; errors: ParseError[]; tokens?: TokenStream }` where `VBAParseTree` is a project-defined type that wraps but does not expose the ANTLR concrete type. Inspections receive the wrapped tree; the ANTLR context is available internally but not in the public interface.

---

### [MEDIUM] No Error Propagation Model Defined

**Description:** The plan defines a result schema (`InspectionResult`) but no error schema. What does `vba/inspect` return when: (1) the VBA code fails to parse entirely, (2) a specific inspection throws an uncaught exception, (3) a file path in `vba/inspect-file` does not exist, or (4) the workspace directory is unreadable? MCP tool errors can be returned as tool-level errors (MCP error response) or as structured data within a successful tool response. The plan does not specify which approach is used or whether it is consistent across tools.

**Risk:** Inconsistent error handling across tools will make it difficult for MCP clients (including Claude Code) to handle errors programmatically. A single failing inspection should not prevent all other results from being returned.

**Recommendation:** Define an error propagation policy: (1) parse failures produce a result with `parseErrors` populated and zero inspection results; (2) individual inspection failures are caught, logged, and produce a result entry with `severity: "error"` and a diagnostic message rather than propagating; (3) tool-level failures (file not found, directory unreadable) are returned as MCP tool errors with structured `data`; (4) the result schema is extended with an optional `parseErrors` field.

---

### [MEDIUM] Configuration Architecture Lacks User Override Semantics

**Description:** `default-config.json` is mentioned in the project structure but not specified anywhere in the plan. No section defines: what fields it contains, how it is loaded, whether users can supply a project-level config file (e.g., `.vbalintrc`), how user overrides are merged with defaults, or whether configuration is passed per-call or loaded once at server startup. The `vba/inspect` tool accepts `severity?` and `categories?` parameters, but their relationship to the config file is undefined.

**Risk:** Without a defined override model, configuration will be implemented ad hoc. The relationship between per-call parameters and persistent config will be inconsistent. Users will have no way to persist inspection preferences without modifying a source file.

**Recommendation:** Specify the configuration hierarchy: (1) `default-config.json` provides baseline; (2) a project-level `.vbalintrc.json` (or similar) overrides defaults; (3) per-call parameters override both. Define the config schema: at minimum `{ enabledInspections, disabledInspections, severityOverrides, hostContext }`. Define how and when the config file is loaded (at startup vs. per-call).

---

### [MEDIUM] No Plugin or User-Defined Inspection Architecture

**Description:** The plan has no provision for user-defined inspections. All inspections are hard-coded in `src/inspections/`. There is no documented extension point, no plugin load path, no interface for external inspection packages. Given that this is an MCP server for VBA development, users may have domain-specific naming conventions, forbidden APIs, or organization-specific patterns they want to enforce.

**Risk:** Without an extension point, the server is a closed system. Any customization requires forking the project. This limits adoption for teams with specific needs. Retrofitting a plugin system after 139 inspections are implemented is costly.

**Recommendation:** Define a minimal extension contract now, even if the implementation is deferred. For example: the config file could reference external inspection module paths that the registry loads dynamically. The `InspectionBase` interface (once defined) becomes the plugin contract. A Phase 6 or future phase could implement the plugin loader without changing any existing inspection.

---

### [MEDIUM] Phase 2 Is Disproportionately Large

**Description:** Phase 2 contains 10 issues covering the entire inspection framework plus ~51 parse-tree inspections. Phase 1 is 5 issues with no inspections. The jump in scope from Phase 1 to Phase 2 is substantial. Issue #6 (inspection framework) is the critical dependency for issues #8–#15, meaning all inspection work is blocked on #6. Issue #7 (registry) is a dependency for #14–#15. The phase has a sequential critical path within itself that is not surfaced.

**Risk:** If #6 takes longer than expected or requires revision (likely given the underspecified class hierarchy), the entire Phase 2 inspection set is blocked. A single architectural revision to `InspectionBase` could require touching all ~51 inspections.

**Recommendation:** Split Phase 2: Phase 2a covers issues #6–#7 (framework only, no inspections), with a proof-of-concept inspection per category to validate the framework. Phase 2b adds all parse-tree inspections once the framework is stable. This adds one gate but dramatically reduces the cost of framework iteration.

---

### [MEDIUM] MCP Tool Decomposition Is Missing a Validation/Syntax-Only Tool

**Description:** The 5 tools cover parse, inspect (string), inspect-file, inspect-workspace, and list-inspections. There is no tool for pure syntax validation without full inspection overhead, and no tool for querying the symbol table of a module (which Phase 3 makes possible). MCP clients that want to check syntax before committing VBA code have no lightweight path — they must run the full inspection suite.

**Risk:** `vba/parse` returns an AST, but parsing does not validate that the code is semantically coherent. There is no tool that answers "is this VBA syntactically valid?" without also running inspections. For an IDE-like use case (real-time as-you-type feedback), this gap matters.

**Recommendation:** Evaluate whether `vba/parse` should also return a `valid: boolean` and `errors: ParseError[]` at the top level — this costs nothing architecturally since the parser already produces this. Also consider whether a `vba/symbols` tool (Phase 3+) that returns the symbol table for a module would be useful. Both decisions should be made before Phase 1 finalizes the tool surface.

---

### [MEDIUM] InspectionResult Schema Has Material Omissions

**Description:** The `InspectionResult` schema is missing several fields that will be needed in practice:

1. **`source` field** — In workspace mode, which file did this result come from? Without it, aggregated workspace results cannot be attributed to files.
2. **`ruleUrl` or `docsUrl`** — A link to the Rubberduck documentation for this inspection. Useful for LLM consumers.
3. **`suppressed: boolean`** — Whether this result was suppressed by `@Ignore` (useful for debugging why an expected finding is absent).
4. **`relatedLocations`** — Some inspections (e.g., duplicate declarations) produce findings that reference two locations. The schema supports only one.
5. **`fixable: boolean`** — Whether a quick-fix is available as a machine-applicable replacement, not just a description.

**Risk:** Adding fields post-implementation requires updating every inspection, the runner, all tests, and the MCP tool response contracts. Schema evolution is expensive if not designed upfront.

**Recommendation:** Review the schema against LSP's `Diagnostic` type (which covers this problem domain) and against how Claude Code will consume results. Add at minimum `source` (file path) and `suppressed` before Phase 2 begins. Defer `relatedLocations` to a future version but document the known gap.

---

### [MEDIUM] Preprocessor Is a Stub With No Resolution Plan

**Description:** `src/parser/preprocessor.ts` is described as a "conditional compilation resolver" and issue #3 includes "preprocessor stub." VBA conditional compilation (`#If`, `#Else`, `#Const`) can fundamentally change what code is parsed. A stub that ignores conditional compilation will produce incorrect parse trees for any VBA that uses it. The plan has no subsequent issue to implement real preprocessor resolution, and no phase that addresses this gap.

**Risk:** Any VBA project that uses `#Const` or host-conditional compilation (common in Excel/Access cross-host code) will produce incorrect inspection results. False positives from dead branches and missed inspections in active branches are both possible. This is not a cosmetic issue — it affects correctness.

**Recommendation:** Add an explicit issue in Phase 2 or Phase 3 for preprocessor resolution. At minimum, define the policy: does the server evaluate `#If`/`#Const` literally, resolve them based on the `host` parameter, or emit parse errors for unresolvable conditionals? This decision affects the grammar and the parser facade.

---

### [MEDIUM] Annotation Architecture Is Isolated From Inspection Suppression Lifecycle

**Description:** `@Ignore` annotation support is issue #27, the last issue in Phase 4. However, annotations affect inspection results from Phase 2 onward. Every inspection that runs during Phase 2–Phase 4 development will not be aware of `@Ignore`. Tests written without `@Ignore` awareness may need to be revised once #27 is implemented. The annotation parser (`src/annotations/`) exists in the structure but has no corresponding phase issue until #27.

**Risk:** 26 issues will produce inspection results without suppression support. If a test fixture accidentally includes an `@Ignore` annotation, it will not be suppressed during Phases 2–4, making the test a false negative. More importantly, any integration testing with real VBA code that contains `@Ignore` annotations will produce incorrect results until the very end of Phase 4.

**Recommendation:** Move `@Ignore` parsing into Phase 2 as part of the inspection framework (issue #6 or a new #6a). Even if suppression is only applied to results, not during inspection execution, having the annotation parser available from the start means inspections can be written with suppression in mind and fixtures can correctly test suppressed vs. unsuppressed behavior.

---

### [LOW] Dependency Graph Omits the Annotation Layer

**Description:** Section 2.2's dependency graph does not include `src/annotations/` as a node. Annotations affect both the inspection runner (which must suppress annotated results) and potentially the symbol walker (which may need to read module-level annotations). The graph shows no path from annotations to any other component.

**Risk:** Minor documentation gap, but it signals that annotations were not considered when the architecture was designed, which reinforces the concern about their late addition in issue #27.

**Recommendation:** Add annotations as a node in the dependency graph. Show it flowing into the runner (for suppression) and optionally into the symbol walker.

---

### [LOW] Test Organization Does Not Mirror Inspection Category Structure

**Description:** `test/inspections/` contains three files: `parse-tree.test.ts`, `declaration.test.ts`, and `reference.test.ts`. With ~51 parse-tree inspections across 6 subcategories, a single `parse-tree.test.ts` file will become unwieldy. The plan notes "every inspection has a corresponding test fixture," but the test file structure does not support this — fixtures live in `test/fixtures/` while tests are in flat files.

**Risk:** Low immediate risk, but a single file with 51+ test suites will become hard to navigate and will cause slow test feedback (all or nothing). Adding a new inspection requires finding the right location in a large file rather than creating a new file.

**Recommendation:** Mirror the inspection directory structure in tests: `test/inspections/parse-tree/empty-blocks/`, `test/inspections/parse-tree/obsolete-syntax/`, etc. Each inspection gets its own test file. Vitest's `glob` test discovery handles this automatically.

---

### [LOW] Performance Benchmarking Is Deferred to Phase 5 With No Architectural Preparation

**Description:** Phase 5 issue #31 adds performance benchmarking and notes "optimize if >2s." No caching, lazy parsing, or parallel execution architecture is described. The plan mentions "lazy parsing, file-level caching, parallel inspection" in the risk register but not in any implementation issue.

**Risk:** If caching requires a cache invalidation mechanism, or parallelism requires the parser to be thread-safe (ANTLR parsers are generally not thread-safe across instances), retrofitting these in Phase 5 may require changes to the parser facade and runner that were finalized in Phases 1–2.

**Recommendation:** Add a note to the parser facade design (Phase 1) specifying whether it will be stateless (safe to call concurrently) or stateful (requiring one instance per call). Add a note to the runner design (Phase 2) specifying whether inspections will be called sequentially or concurrently. These design decisions cost nothing to make now and prevent expensive rework in Phase 5.

---

## Summary Table

| # | Finding | Severity | Recommendation |
|---|---------|----------|----------------|
| 1 | TBD items cover 33% of inspection catalog | CRITICAL | Enumerate all inspections before plan approval |
| 2 | Declaration model v3/v2 version mismatch | CRITICAL | Make an explicit compatibility decision before Phase 3 |
| 3 | Symbol resolution multi-pass requirement unaddressed | CRITICAL | Add multi-pass design to Phase 3 spec |
| 4 | ANTLR4 runtime choice unresolved | HIGH | Make runtime selection a Phase 1 exit gate via spike |
| 5 | Registry contract between JSON metadata and classes undefined | HIGH | Specify the coupling contract explicitly |
| 6 | Inspection class hierarchy named but not specified | HIGH | Define full interface before writing any inspection |
| 7 | Workspace model needed in Phase 4, defined in Phase 5 | HIGH | Move workspace model definition to Phase 3 |
| 8 | "~60 parse-tree inspections" is ~51; double-counted entries | HIGH | Audit and fix all count claims; remove duplicates |
| 9 | Parser facade leaks ANTLR internals risk | MEDIUM | Define a wrapper type; insulate callers from ANTLR |
| 10 | No error propagation model defined | MEDIUM | Define a policy for parse errors, inspection exceptions, tool errors |
| 11 | Configuration override semantics undefined | MEDIUM | Specify config hierarchy and schema |
| 12 | No plugin/user-defined inspection architecture | MEDIUM | Define extension contract now, defer implementation |
| 13 | Phase 2 is disproportionately large | MEDIUM | Split into 2a (framework) and 2b (inspections) |
| 14 | Missing syntax validation and symbol query tools | MEDIUM | Add `valid`/`errors` to `vba/parse`; evaluate `vba/symbols` |
| 15 | InspectionResult missing `source`, `suppressed`, `relatedLocations` | MEDIUM | Extend schema before Phase 2 |
| 16 | Preprocessor stub has no resolution plan | MEDIUM | Add explicit issue for preprocessor policy |
| 17 | @Ignore support deferred to last issue of Phase 4 | MEDIUM | Move annotation parsing to Phase 2 framework |
| 18 | Annotation layer absent from dependency graph | LOW | Add to dependency graph |
| 19 | Test file structure does not scale to 139 inspections | LOW | Mirror inspection directory structure in tests |
| 20 | Performance architecture deferred with no preparation | LOW | Document stateless/stateful decisions in Phase 1-2 |

---

## Overall Assessment

**Fail — revise and re-review before implementation begins.**

The plan demonstrates good strategic intent and the phased delivery approach is sound. However, three CRITICAL findings (TBD catalog items, cross-version declaration model mismatch, and unaddressed multi-pass symbol resolution) represent genuine architectural risks that cannot be resolved during implementation without significant rework. Five HIGH findings further undermine confidence in the plan's executability. The plan should not proceed to Phase 1 until all CRITICAL findings are resolved and the HIGH findings related to core contracts (registry, class hierarchy, workspace model) are addressed with at least a documented decision.
