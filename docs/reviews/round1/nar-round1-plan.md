# NAR Round 1 — Consolidated Plan Review

**Reviewer:** Non-Advocate Review (Consolidated)
**Date:** 2026-03-19
**Target:** PLAN.md (Implementation Plan)
**Round:** 1 of 3 (pre-implementation)

---

## Verdict: FAIL — Revise Before Implementation

The plan demonstrates strong strategic intent: a phased MCP server delivering incremental value, a well-chosen upstream source in Rubberduck v2, and a parse-tree-first delivery strategy that is the correct decomposition for this problem. However, the plan has significant gaps across all six dimensions that, taken together, mean implementation would proceed without foundational contracts being defined. The most consequential gaps are concentrated in two areas: (1) architectural contracts that every subsequent line of code depends on are named but not specified, and (2) the testing strategy describes the existence of tests, not their quality.

This review synthesizes findings from six individual dimension reviews (architecture, test-quality, maintainability, security, performance, production-readiness) into a prioritized action plan. Findings are deduplicated and cross-referenced where multiple dimensions surface the same underlying gap.

---

## Critical Findings (Must Fix Before Implementation)

These findings represent structural gaps that cannot be resolved during implementation without significant rework. All must be addressed in the revised plan.

### C1. One-Third of the Inspection Catalog Is Undefined

**Dimensions:** Architecture, Maintainability
**PLAN.md Sections:** 5.1 (items 85-100, 110-139), 3.4 (Phase 4)

Items 85-100 (16 declaration inspections) and 110-139 (30 reference inspections) are explicitly labeled "will be fully enumerated during Phase 4 implementation." This means 46 of 139 target inspections (33%) have no defined scope, no issue ticket, and no complexity estimate. The plan's phase sizing, success criteria (Section 10: ">=130 of 139"), and effort estimates are all built on unknowns.

**Action required:** Enumerate every inspection in Rubberduck v2's `Rubberduck.CodeAnalysis/Inspections/Concrete/` directory by name. Classify each as parse-tree, declaration, or reference. Assign a complexity estimate (low/medium/high). If Phase 4 is too large after enumeration, split it. Remove all "~" approximations and replace with actual counts that match the enumerated list.

**Additionally:** The current enumerated list contains duplicates (EmptyMethodInspection at items 8 and 84; ImplicitActiveSheetReferenceInspection at items 36 and in Issue #25). Resolve whether these are the same inspection in two modes or distinct inspections with distinct IDs.

---

### C2. Declaration Model v3 / Inspection Logic v2 Version Mismatch

**Dimensions:** Architecture
**PLAN.md Sections:** 9.1 (attribution table), 3.3 (Phase 3)

Section 9.1 states `src/symbols/declaration.ts` is derived from Rubberduck v3's `Declaration.cs`, while all inspections are ported from v2. The v3 declaration model was a substantial rewrite with different scope semantics, module identity, and member resolution. Inspections written against v2 assumptions will not necessarily produce correct results against a v3-derived symbol table. This mismatch may not surface until Phase 4, after Phase 3 has committed to the v3 model.

**Action required:** Make an explicit architectural decision and document it in the plan: either port the declaration model from v2 (matching the inspection source), or audit every v2 inspection that touches the declaration model to confirm compatibility with v3 semantics. Do not leave both assumptions unchallenged.

---

### C3. Symbol Resolution Requires Multi-Pass; Plan Assumes Single-Pass

**Dimensions:** Architecture
**PLAN.md Sections:** 3.3 (Phase 3, Issues #16-#19)

Phase 3 describes four sequential issues: declaration model, symbol walker, declaration finder, scope resolution. This implies a linear pipeline, but VBA symbol resolution is not single-pass. Forward references between procedures, cross-module public member visibility, and interface implementation require at least two passes (collect, then resolve). The plan contains no acknowledgment of multi-pass requirements, no design for coordination, and no handling of unresolved references.

**Action required:** Add an explicit multi-pass design to Phase 3. Define: what pass 1 collects, what triggers pass 2, how the declaration finder handles unresolved references, and what happens when analyzing a single file vs. a full workspace.

---

### C4. ANTLR4 TypeScript Runtime Choice Is Unresolved

**Dimensions:** Architecture, Performance
**PLAN.md Sections:** 2.3 (dependencies), 7 (risk register)

Section 2.3 lists `antlr4` and `antlr4ts-cli` without choosing. `antlr4ng` appears only as a fallback in the risk register. These runtimes have different generated code shapes, visitor APIs, and error listener interfaces. The choice affects every inspection, the parser facade, the symbol walker, and all tests. Selecting the wrong one means rewriting all generated bindings mid-project.

**Action required:** Make the runtime selection a Phase 1 exit gate. Write a spike: generate VBA grammar with each candidate, parse a non-trivial file, verify the visitor pattern works. Record the decision in the plan. Do not begin Phase 2 with this open.

---

### C5. Core Architectural Contracts Are Named But Not Specified

**Dimensions:** Architecture, Maintainability
**PLAN.md Sections:** 2.1 (project structure), 3.2 (Phase 2, Issues #6-#7)

Three foundational contracts that every inspection depends on are mentioned by name but never specified:

1. **Inspection class hierarchy** (Issue #6) — `InspectionBase`, `ParseTreeInspection`, `DeclarationInspection`, `ReferenceInspection` are named. No method signatures, no required static properties, no `inspect(context)` definition, no return type specification. 60 inspections will be written against this unspecified contract in Phase 2.

2. **Registry coupling contract** (Issue #7) — How does the registry match a metadata JSON entry to an inspection class? By class name? Static `id` property? What happens when a metadata entry has no class, or vice versa? The "auto-discovery" mechanism is undefined and will fail silently in a compiled TypeScript environment.

3. **Parser facade output type** (Issue #3) — `parseCode(source: string) -> ParseTree` is the entire specification. Is `ParseTree` the raw ANTLR type (coupling all callers to the runtime) or a wrapper? Are parse errors returned in the result or via a separate channel?

**Action required:** Before Phase 2 begins, specify all three contracts completely:
- Inspection base: constructor signature, required static properties (id, name, category, severity), `inspect(context)` method signature with explicit context type, return type.
- Registry: discovery mechanism (explicit barrel file preferred for compile-time safety), static interface contract, failure modes for missing metadata or missing classes.
- Parser facade: define `ParseResult { tree: VBAParseTree; errors: ParseError[]; tokens?: TokenStream }` where `VBAParseTree` wraps but does not expose ANTLR internals.

---

### C6. Testing Strategy Describes Existence, Not Quality

**Dimensions:** Test Quality
**PLAN.md Sections:** 4.1 (unit tests), 4.3 (integration tests), 4.4 (coverage target)

The coverage target of "one positive + one negative test per inspection" is the minimum bar for confirming an inspection exists, not for validating correctness. The example test (Section 4.1) asserts only `results.length` and `severity` — it does not assert the inspection ID, location accuracy, or that the result comes from the right inspection. Location accuracy (startLine, startColumn, endLine, endColumn) is never mentioned as a test requirement, yet it is the primary data that Claude Code uses to navigate to problems.

The integration test strategy (Section 4.3) describes a manual tmux procedure, not an automated suite. `test/integration/mcp-integration.test.ts` appears in the directory structure but has no defined contents.

**Action required:** Redefine the testing contract:
- Every positive test must assert: `results.length`, `results[0].inspection` (exact ID), `results[0].severity`, and `results[0].location.startLine` at minimum.
- Every inspection must have at least one false-positive test (syntactically similar but correct construct, expecting zero results).
- Define `mcp-integration.test.ts` contents: initialize handshake, each tool with valid and invalid inputs, concurrent calls, malformed JSON-RPC. Must run in CI without tmux or Claude Code.
- Add parse error path tests: completely unparseable input, partial recovery, empty string, binary content.

---

### C7. No Error Handling Strategy Across Any Layer

**Dimensions:** Production Readiness, Architecture
**PLAN.md Sections:** 1.3 (result schema), 3.2 (Phase 2)

The plan defines `InspectionResult` but no error schema. Unspecified: what happens when a VBA file fails to parse entirely, when a single inspection throws, when a file path does not exist, or when the workspace directory is unreadable. If a single broken inspection kills all output or crashes the process, the MCP server is unreliable. This is compounded by the MCP protocol having no built-in retry.

**Action required:** Define the error propagation policy:
- Per-inspection isolation: each inspection runs in try/catch; failures produce a structured error entry, not a crash.
- Parse failures return a result with `parseErrors` populated and zero inspection results.
- Tool-level failures (file not found, directory unreadable) use MCP's `isError: true` response.
- Extend the result schema: `{ results: InspectionResult[], errors: { inspection: string, message: string }[], parseErrors?: ParseError[] }`.

---

### C8. Unrestricted File System Read via Path Parameters

**Dimensions:** Security
**PLAN.md Sections:** 1.2 (tool surface), 3.5 (Phase 5, Issue #28)

`vba/inspect-file` and `vba/inspect-workspace` accept unconstrained `path` parameters. No allowlist, no chroot, no cwd-relative restriction. Any MCP client (including a prompt-injected AI session) can read any file readable by the Node.js process user. The `glob` parameter in `vba/inspect-workspace` compounds this with potential directory traversal via `../` in patterns.

**Action required:**
- Define a mandatory `rootDir` configuration value set at server startup.
- Resolve all incoming paths with `path.resolve()` and verify the result starts with `rootDir` before any I/O.
- Reject absolute paths from callers; accept only relative paths resolved against `rootDir`.
- Hardcode allowed file extensions (`*.bas`, `*.cls`, `*.frm`) for workspace scanning. Do not accept arbitrary glob patterns in the initial implementation.

---

### C9. No Resource Limits on Parser or Workspace Scanning

**Dimensions:** Security, Performance, Production Readiness
**PLAN.md Sections:** 2.3, 3.1 (Phase 1), 3.5 (Phase 5)

The ANTLR4 ALL(*) parser has super-linear worst-case complexity. No input size cap, no per-call timeout, no maximum file count for workspace scans. A single crafted input can peg the event loop; a workspace scan of a large directory can exhaust memory.

**Action required:**
- Maximum input size: 512 KB for `vba/inspect` code strings, 512 KB per file for file/workspace tools.
- Maximum file count: 500 files per workspace scan (configurable).
- Per-call timeout using `Promise.race` or worker thread with hard wall-clock limit (10s).
- Wrap parser invocation in try/catch for `RangeError` (stack overflow from deeply nested trees).

---

## High Findings (Should Fix Before Implementation)

These findings represent significant risks that should be addressed in the revised plan. Deferring them is possible but increases the cost of correction.

### H1. Workspace Model Needed in Phase 4, Defined in Phase 5

**Dimensions:** Architecture
**PLAN.md Sections:** 3.4 (Phase 4), 3.5 (Phase 5, Issue #28)

Phase 4 tests require "cross-module reference tracking" and "workspace-level inspections work across multiple files," but the workspace abstraction is not defined until Phase 5's `vba/inspect-workspace`. This is a dependency inversion.

**Recommendation:** Define the workspace model (a named set of parse trees + symbol tables) in Phase 3 alongside symbol resolution. Phase 4 inspections receive a workspace context. Phase 5 adds the MCP tool that creates a workspace from a directory scan.

---

### H2. Phase 2 Is Disproportionately Large with Internal Dependencies

**Dimensions:** Architecture, Maintainability
**PLAN.md Sections:** 3.2 (Phase 2, Issues #6-#15)

Phase 2 contains 10 issues covering the inspection framework plus ~51 parse-tree inspections. Issues #8-#15 are blocked on #6 (framework) and #7 (registry). If the framework design needs revision after the first batch of inspections, all previously written inspections require rework.

**Recommendation:** Split Phase 2: Phase 2a covers Issues #6-#7 (framework + registry) with one proof-of-concept inspection per subcategory. Phase 2b adds all remaining parse-tree inspections once the framework is stable.

---

### H3. InspectionResult Schema Has Material Omissions

**Dimensions:** Architecture, Production Readiness
**PLAN.md Sections:** 1.3 (result schema)

Missing fields that will be needed:
- `source` (file path) — required for workspace mode to attribute results to files.
- `suppressed: boolean` — needed once @Ignore is implemented; useful for debugging.
- `parseErrors` — needed per C7 above.

Adding fields post-implementation requires updating every inspection, the runner, all tests, and the tool response contracts.

**Recommendation:** Add `source` and `suppressed` to the schema before Phase 2 begins.

---

### H4. @Ignore Annotation Support Deferred to Last Issue of Phase 4

**Dimensions:** Architecture, Test Quality
**PLAN.md Sections:** 3.4 (Phase 4, Issue #27)

@Ignore is the user's primary escape hatch for false positives. Deferring it to the last issue of Phase 4 means 26 issues produce results without suppression support. Tests written without @Ignore awareness may need revision. Real VBA code containing @Ignore annotations will produce incorrect results through all of Phases 2-4.

**Recommendation:** Move @Ignore parsing into Phase 2 as part of the inspection framework (Issue #6 or a new issue). Even if suppression is only applied to results post-inspection, having the parser available from the start means tests can cover suppressed vs. unsuppressed behavior.

---

### H5. No Logging or Observability Design

**Dimensions:** Production Readiness
**PLAN.md Sections:** (absent — no section addresses logging)

The server uses stdio for MCP protocol. Any `console.log` to stdout corrupts the JSON-RPC stream. The plan has no logger, no log levels, no structured format, and no designated output channel.

**Recommendation:** Designate stderr as the exclusive log channel from Phase 1. Use a structured logger (e.g., pino). Log: server start with version, tool calls with parameters (excluding code content), parse errors, inspection errors, timing. Make log level configurable via environment variable.

---

### H6. Performance Architecture Deferred to Phase 5

**Dimensions:** Performance
**PLAN.md Sections:** 3.5 (Phase 5, Issue #31), 7 (risk register)

All performance work is deferred to Phase 5 after all 139 inspections are written. Two structural decisions that cannot be cheaply retrofitted are affected:

1. **No caching layer** — Every call re-parses from scratch. No `ParseCache` is defined.
2. **No listener-fusion model** — 139 inspections running as 139 independent tree walks creates O(139 * n) node visits per file. The performance target (100ms) likely cannot accommodate this without listener fusion.

**Recommendation:**
- Design a `ParseCache` component keyed on `(filePath, contentHash)` in Phase 1. Bound cache to 50 entries with LRU eviction.
- Design the runner to support listener fusion: inspections register which grammar contexts they care about; the runner dispatches during a single tree walk. This must be decided before Phase 2 writes 60 inspections to an independent-traversal model.
- Add performance checkpoints at end of Phase 1 (parse time baseline), Phase 2 (runner overhead with 60 inspections), and Phase 3 (symbol walk overhead).

---

### H7. ANTLR4 Cold Start Not Addressed

**Dimensions:** Performance, Production Readiness
**PLAN.md Sections:** 3.1 (Phase 1, Issue #4)

The ANTLR4 TypeScript runtime initializes ATN serialization and DFA cache on first use. For the VBA grammar, this can take 200-500ms. The server will incur this penalty on the first tool call of each session.

**Recommendation:** Add an explicit warm-up step in server startup: parse a minimal VBA stub before accepting tool calls. Document expected startup latency.

---

### H8. Encoding and Line Ending Handling Absent

**Dimensions:** Production Readiness, Test Quality
**PLAN.md Sections:** (absent)

VBA files exported from the VBE can be Windows-1252, UTF-8 with BOM, or UTF-16 LE. CRLF line endings are standard. Node.js `fs.readFile().toString()` assumes UTF-8. Non-ASCII characters in identifiers or comments will be misread. CRLF handling affects line number accuracy in all inspection results.

**Recommendation:** Default to UTF-8 with BOM stripping. Fall back to Windows-1252 for files with bytes in 0x80-0x9F range. Add CRLF and BOM fixtures to the test suite. Test line number accuracy with CRLF files.

---

### H9. Quick-Fix Correctness Is Not Tested

**Dimensions:** Test Quality
**PLAN.md Sections:** 1.3 (quickFix in result schema), 5.2 (metadata extraction)

The plan includes `quickFix.replacement` in results but provides no test strategy for whether the replacement is syntactically valid VBA or actually resolves the flagged issue.

**Recommendation:** For every inspection with a `quickFix.replacement`, add an "apply-and-recheck" test: apply the replacement at the reported location, re-parse, re-inspect, assert zero results for that inspection.

---

### H10. Docker Container Security Posture Unspecified

**Dimensions:** Security, Production Readiness
**PLAN.md Sections:** 3.6 (Phase 6, Issue #32)

The Dockerfile is Phase 6 with one line of description. Default Node.js Docker images run as root. Combined with the path traversal issue (C8), root-level reads inside the container could expose mounted secrets.

**Recommendation:** Specify in the plan: non-root user (`USER node`), Alpine base image (`node:22-alpine`), no privileged mode, read-only volume mounts where possible. Move these requirements to Phase 1 so any Docker-based testing uses the correct posture.

---

## Medium Findings (Recommended Improvements)

These findings represent quality improvements that strengthen the plan but do not block implementation if the critical and high items are addressed.

### M1. Configuration Architecture Undefined

**Sections:** 2.1 (default-config.json), 1.2 (tool parameters)

No specification of how users override defaults. No `.vbalintrc.json` schema, no merge strategy, no load timing. The `severity?` and `categories?` tool parameters have no defined relationship to persistent config.

**Recommendation:** Define config hierarchy before Phase 2: `default-config.json` -> `.vbalintrc.json` (workspace root) -> per-call parameters. Define schema and validation.

---

### M2. .vbalintignore Has No Specification

**Sections:** 3.5 (Phase 5, Issue #28)

One sentence: "respect .vbalintignore." No pattern syntax, no lookup rules, no edge cases, no test strategy.

**Recommendation:** Write a dedicated spec: use minimatch/micromatch, look up from workspace root only, support `#` comments, define negation behavior, add test fixtures.

---

### M3. Preprocessor Is a Permanent Stub with No Resolution Plan

**Sections:** 2.1, 3.1 (Phase 1, Issue #3)

`preprocessor.ts` is described as a "conditional compilation resolver" and a "stub." No phase issue addresses completing it. VBA conditional compilation (`#If`, `#Const`) is common in cross-host code. A stub produces incorrect parse trees for any code using it.

**Recommendation:** Either assign a phase for real implementation, or explicitly scope it out with documented limitations. At minimum, define the stub's behavior (e.g., "all branches inspected as if active") so inspection authors know what to expect.

---

### M4. inspection-metadata.json Sync Problem

**Sections:** 2.1, 5.2

A single JSON file holding metadata for 139 inspections with no defined sync protocol against the inspection classes. No test validates completeness. New inspections can be added without updating the JSON, creating silent gaps.

**Recommendation:** Either move metadata into each inspection class as static properties (eliminating the sync problem), or add a startup-time validation that every registered inspection ID has a corresponding JSON entry. The former is preferred.

---

### M5. Test Organization Does Not Scale

**Sections:** 2.1 (test directory structure)

Three test files (`parse-tree.test.ts`, `declaration.test.ts`, `reference.test.ts`) for 139 inspections. The parse-tree file alone will contain 120+ test cases.

**Recommendation:** Mirror inspection directory structure: `test/inspections/parse-tree/empty-blocks.test.ts`, etc. One test file per inspection or per subcategory.

---

### M6. Dependency Version Pinning

**Sections:** 2.3

`@modelcontextprotocol/sdk` at `latest` and `vitest` at `latest`. The MCP SDK evolves rapidly with breaking changes.

**Recommendation:** Pin all dependencies to exact versions. Use lockfile. Add `npm audit` to CI.

---

### M7. Grammar Maintenance Path Undefined

**Sections:** 9.1

Grammar files are treated as static copies. No mechanism for tracking upstream Rubberduck changes, no pinned commit hash, no upgrade procedure.

**Recommendation:** Pin grammar to a specific Rubberduck commit hash in `grammar/SOURCE.md`. Add regression tests from Rubberduck's test corpus. Document upgrade procedure.

---

### M8. Inspection Boilerplate Unquantified

**Sections:** 3.2 (Phase 2)

Each inspection requires class definition, listener/visitor boilerplate, result construction, metadata lookup, host filtering, and export. If the boilerplate-to-logic ratio is 10:1, the inspection set will be difficult to maintain.

**Recommendation:** Write one complete reference inspection before Phase 2 begins. Measure file size and boilerplate ratio. If ratio exceeds 5:1, redesign the base class. For families of 3+ structurally identical inspections (e.g., empty blocks), design a parameterized pattern rather than separate classes.

---

### M9. No Versioning or Breaking Change Strategy

**Sections:** (absent)

The tool surface (5 tools, parameter names, result schema) has no versioning strategy. No guidance on what constitutes a breaking change.

**Recommendation:** Expose version in MCP `serverInfo`. Define breaking vs. non-breaking changes. Add CHANGELOG.

---

### M10. ATTRIBUTION.md Deferred Past When It Is Needed

**Sections:** 3.6 (Phase 6, Issue #33), 9.2

ATTRIBUTION.md is Phase 6, but Rubberduck-derived code enters the codebase in Phase 2. GPL-3.0 attribution is required from the moment derived code exists.

**Recommendation:** Write ATTRIBUTION.md in Phase 2, not Phase 6.

---

## Low Findings (Nice to Have)

| # | Finding | Section | Recommendation |
|---|---------|---------|----------------|
| L1 | Annotation layer absent from dependency graph | 2.2 | Add annotations node to the graph |
| L2 | No snapshot testing for `vba/parse` AST output | 4.1 | Add snapshot tests for canonical VBA constructs |
| L3 | No performance regression gate in CI | 4.4 | Add wall-clock smoke test for fixed VBA fixture |
| L4 | Error messages may leak filesystem structure | (absent) | Return generic error strings; log details to stderr only |
| L5 | No rate limiting on tool calls | (absent) | Add concurrency limit by Phase 5 |
| L6 | `vba/parse` full AST serialization may be too large | 1.2, 3.1 | Define condensed AST format; support `depth` parameter |
| L7 | Docker health check is not feasible for stdio transport | 3.6 | Drop the health check claim or document why it is absent |
| L8 | No `vba/validate` or `vba/symbols` tool | 1.2 | Evaluate whether `vba/parse` should return `valid: boolean` and `errors[]` at top level |

---

## Cross-Cutting Themes

Three themes emerge across all six dimensions:

### Theme 1: Contracts Are Named, Not Specified

The plan names every major abstraction (InspectionBase, registry, parser facade, workspace model, error handling, config) but specifies none of them. This creates the illusion of architectural completeness while deferring every binding decision to implementation. The risk is not that the implementer will make bad decisions — the risk is that decisions made incrementally across 35 issues will be inconsistent with each other.

### Theme 2: Testing Conflates Coverage with Quality

"One positive + one negative test per inspection" produces a green CI badge. It does not produce a reliable linting tool. Without location accuracy assertions, false-positive tests, quick-fix validation, and error-path coverage, the test suite validates existence, not correctness.

### Theme 3: Performance and Security Are Deferred Past the Point of Cheap Correction

Performance architecture (caching, listener fusion, concurrency) and security controls (path containment, input limits, encoding) are deferred to Phase 5-6 but depend on decisions that should be made in Phase 1-2. Retrofitting these after 139 inspections are written is expensive. The plan should treat these as design constraints, not optimization tasks.

---

## Prioritized Revision Checklist

The following items must be addressed before the plan can pass Round 2 review:

**Must fix (blocks implementation):**
- [ ] C1: Enumerate all 139 inspections; resolve duplicates; fix count claims
- [ ] C2: Decide v2 vs v3 declaration model; document compatibility analysis
- [ ] C3: Add multi-pass symbol resolution design to Phase 3
- [ ] C4: Make ANTLR4 runtime selection a Phase 1 exit gate
- [ ] C5: Specify inspection base class, registry contract, and parser facade output type
- [ ] C6: Redefine test coverage requirements with location, ID, and false-positive assertions
- [ ] C7: Define error propagation policy across all layers
- [ ] C8: Add path containment and file extension restrictions
- [ ] C9: Add input size limits and per-call timeouts

**Should fix (significant risk if ignored):**
- [ ] H1: Move workspace model definition to Phase 3
- [ ] H2: Split Phase 2 into framework (2a) and inspections (2b)
- [ ] H3: Add `source` and `suppressed` fields to InspectionResult
- [ ] H4: Move @Ignore parsing to Phase 2
- [ ] H5: Add logging design (stderr, structured, from Phase 1)
- [ ] H6: Add parse cache design and listener-fusion decision to Phase 1-2
- [ ] H7: Add ANTLR4 warm-up at server startup
- [ ] H8: Add encoding detection and CRLF handling
- [ ] H9: Add quick-fix apply-and-recheck test requirement
- [ ] H10: Specify Docker security posture

---

## Dimension Scores

| Dimension | Weight | Score | Notes |
|-----------|--------|-------|-------|
| Architecture | PRIMARY | FAIL | 3 critical, 5 high — core contracts undefined |
| Test Quality | PRIMARY | FAIL | 4 critical, 6 high — coverage target is not a quality bar |
| Maintainability | SECONDARY | FAIL | 3 critical, 4 high — registry, metadata sync, test org |
| Security | SECONDARY | FAIL | 3 critical, 4 high — path traversal, parser DoS, glob injection |
| Performance | SECONDARY | FAIL | 2 critical, 4 high — no caching, no baseline, deferred to Phase 5 |
| Production Readiness | SECONDARY | FAIL | 3 critical, 5 high — no error handling, no logging, no limits |

---

## Closing Note

The plan's core strategy is sound. The phased delivery, parse-tree-first approach, Rubberduck source material, and MCP tool surface are all well-chosen. The failures are not in the "what" but in the "how specifically" — the plan stops one level short of specification on every major design decision. Addressing the critical and high findings will produce a plan that can be implemented confidently. The majority of findings require adding specification detail, not changing the architecture.
