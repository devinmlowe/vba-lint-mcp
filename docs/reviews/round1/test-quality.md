# NAR Round 1 — Test Quality Review

**Reviewer:** Test Quality Dimension
**Date:** 2026-03-19
**Target:** PLAN.md (Implementation Plan)
**Weight:** PRIMARY

---

## Executive Summary

The plan establishes a recognizable testing skeleton — vitest, fixture files, a positive/negative pair per inspection — but the testing strategy is shallow relative to the complexity of what is being built. The coverage target of "one positive + one negative test per inspection" is the minimum possible bar, not a quality bar, and the plan contains no strategy for false positive prevention, parse error paths, cross-module symbol scenarios, annotation accuracy, quick-fix correctness, or regression protection. Integration testing is described in two sentences that amount to "start a server and send a request," with no validation protocol defined.

---

## Findings

### [CRITICAL] Coverage Target Is Functionally Meaningless

**Description:** Section 4.4 defines the inspection coverage target as "100% of inspections have at least one positive + negative test." One positive case and one negative case is the bare minimum to confirm an inspection exists and does not immediately crash — it says nothing about correctness. For a 139-inspection catalog ported from C# to TypeScript, a single happy-path pair per inspection will routinely miss boundary conditions, off-by-one line numbers, nested construct handling, and interaction effects.

**Risk:** Inspections ship with silent logic errors. A positive test that asserts `results.length === 1` passes even when the result has wrong line numbers, wrong severity, or wrong inspection ID. The coverage metric looks green while the implementation is wrong.

**Recommendation:** Redefine the coverage target to require: (1) at least one positive case asserting correct location (startLine, startColumn, endLine, endColumn), correct severity, and correct inspection ID; (2) at least one negative case for a syntactically similar but correct construct; (3) at least one edge case per inspection (nested, combined, boundary condition). Add a line-location accuracy assertion requirement as a mandatory part of the positive case contract.

---

### [CRITICAL] No False Positive Testing Strategy

**Description:** The plan mentions "negative case: VBA code that should NOT trigger the inspection" exactly once, in the example pattern. There is no systematic treatment of false positive prevention. False positives are the primary user-trust failure mode for a linter. The plan has no concept of: inspections that should be silent when constructs are used correctly together, inspections that interact (e.g., `ObsoleteLet` should not fire inside a comment), or inspections that must be host-aware (should not fire for generic host when targeting Excel).

**Risk:** The tool fires on valid code. Users stop trusting and disable the MCP. This is a reliability failure more damaging than missing a true positive.

**Recommendation:** Add a dedicated "false positive" test category. For every inspection, require at least one test that exercises the closest legitimate construct to the trigger pattern and asserts zero results. For host-specific inspections, require a test that runs the inspection against the wrong host and expects zero results. Treat false positive test count as a tracked metric alongside positive test count.

---

### [CRITICAL] Parse Error Handling Has No Test Strategy

**Description:** The plan mentions an `error-listener.ts` component and states Phase 1 tests include "Parser returns meaningful errors for invalid syntax," but there is no systematic test strategy for how parse errors propagate through the inspection pipeline. Open questions with no test coverage plan: What happens when the parser partially recovers? Do inspections run on a partially-parsed tree? What does `vba/inspect` return when the input is not parseable VBA at all? What about truncated files, binary content in a .bas file, files with BOM markers?

**Risk:** The error path is the most common path for malformed or work-in-progress VBA. Untested error handling will produce confusing or crashing behavior in real use.

**Recommendation:** Add a dedicated `test/parser/error-handling.test.ts` that covers: completely unparseable input, partially valid input (error recovery), empty string, null-like input, binary/non-UTF8 content, oversized input. Specify the contract: inspection runner must either return partial results with an error flag, or return zero results with a structured error — and that contract must be tested.

---

### [CRITICAL] Integration Test Strategy Is Not a Strategy

**Description:** Section 4.3 describes the integration test approach as: "Start MCP server in a tmux pane, launch Claude Code in another pane, send requests, capture responses." This is a manual smoke test procedure, not an integration test suite. The listed `test/integration/mcp-integration.test.ts` file has no described contents in the plan. The Phase 5 Issue #29 says "Full MCP protocol tests: initialize, tool calls, result validation" — these seven words are the entirety of the integration test specification.

**Risk:** The MCP protocol layer is never automatically tested. Protocol-level bugs (malformed JSON-RPC, incorrect tool schema, missing required response fields) are only caught manually or by end users.

**Recommendation:** Define the integration test suite contents explicitly in the plan: (1) `initialize` handshake returns correct server info and tool list; (2) each of the 5 tools called with valid parameters returns a structurally valid response (schema-validated); (3) each tool called with invalid parameters returns a structured error, not a crash; (4) concurrent tool calls do not produce shared-state corruption; (5) the server handles malformed JSON-RPC input gracefully. This suite must run in CI without tmux or Claude Code.

---

### [HIGH] Test File Organization Creates Shared Fixture Ambiguity

**Description:** The fixture structure (`test/fixtures/empty-blocks/empty-if.bas`, etc.) assigns one fixture per inspection trigger, but multiple inspections often fire on the same code. The `kitchen-sink.bas` file is described as "triggers as many inspections as possible" — this file will be a maintenance nightmare. When it starts failing, the failure message will not identify which inspection regressed. Conversely, tests that load fixture files rather than inline strings create implicit coupling: modifying a shared fixture to fix one test can silently break another.

**Risk:** Test failures are hard to diagnose. Fixture changes cause unexpected test breaks. The `clean/well-written-module.bas` file will drift — a future inspection may fire on it without anyone noticing unless the test explicitly checks zero results.

**Recommendation:** Establish a fixture ownership rule: each fixture file is owned by exactly one test. Shared fixtures are forbidden. The `kitchen-sink.bas` concept should be replaced with a workspace-level integration fixture that is only used by workspace scanner tests, not by individual inspection tests. Add a test that loads every fixture in `clean/` and asserts zero total diagnostics — run this as a regression gate.

---

### [HIGH] Inline Code Strings vs External Fixture Files — Decision Not Made

**Description:** The plan uses inline code strings in the unit test example pattern (`const code = \`Sub Test()\\n  If True Then...`) and simultaneously defines an external fixture directory (`test/fixtures/`). The plan never resolves which approach applies where, or why. These two approaches have significant trade-offs that affect test readability, maintainability, and IDE support for VBA syntax.

**Risk:** The team will make inconsistent choices per inspection, resulting in a test suite that mixes approaches without rationale. Inline strings cannot be syntax-highlighted, cannot be opened in an editor with VBA support, and are error-prone when VBA requires specific whitespace. External fixtures require I/O in tests and make the test less self-contained.

**Recommendation:** Make an explicit decision and document it. Recommended approach: use inline strings for unit tests of individual inspections (fast, self-contained, no I/O), and use external `.bas`/`.cls` fixtures for integration tests and workspace scanner tests (realistic file content, real module headers). Document this rule in the test strategy section. Add a linting rule or test helper that validates inline VBA strings are parseable before the inspection runs, to catch typos in test fixtures.

---

### [HIGH] Location Accuracy Is Not Tested

**Description:** The `InspectionResult` schema defines `location: { startLine, startColumn, endLine, endColumn }`. The example unit test only asserts `results.length` and `results[0].severity`. Column and line number accuracy is never mentioned as a test requirement anywhere in the plan. For a tool used by Claude Code to navigate to problems in VBA files, wrong line numbers are a silent usability failure.

**Risk:** Inspections return structurally valid results with wrong line/column data. The tests pass. Claude Code jumps to the wrong line. Users lose confidence in the tool.

**Recommendation:** Mandate that every positive-case test assert the exact expected location. Add a test helper `expectResultAt(result, startLine, startColumn, endLine, endColumn)` that provides a readable failure message. For complex inspections, include at least one test with a multi-line construct to verify the span is computed correctly.

---

### [HIGH] Symbol Resolution Tests Are Underspecified for Cross-Module Scenarios

**Description:** Phase 3 tests mention "Handles multiple modules in a workspace context" as a single bullet point. Symbol resolution across module boundaries is the hardest correctness problem in the entire system. The plan provides no test cases for: a procedure in Module A calling a Public procedure in Module B; a variable declared at module level in a Class and referenced from a standard module; `Friend` visibility across modules in the same project; name collision between a local variable and a module-level declaration; or the order-dependence of multi-pass resolution.

**Risk:** Cross-module symbol resolution ships with silent bugs. Declaration-based and reference-based inspections (Phase 4, ~80 inspections) produce wrong results for any code involving inter-module references — which is most real-world VBA.

**Recommendation:** Define a multi-module test fixture set specifically for Phase 3. At minimum: (1) two-module Public procedure call — both declaration and usage found; (2) module-level variable shadowed by local variable — correct scope resolution; (3) Friend access from same project, blocked from different project; (4) circular reference between modules — no infinite loop; (5) name resolution when the same identifier is declared in two modules — correct disambiguation. Each scenario must have a positive and negative test.

---

### [HIGH] Annotation (@Ignore) Test Strategy Is Absent

**Description:** Issue #27 implements `@Ignore` annotation support in one sentence: "Parse @Ignore annotations, suppress matching inspections in results." The Phase 4 test bullet says "@Ignore correctly suppresses inspections" — singular. There is no specification of: what happens with a malformed annotation; what happens when @Ignore names a non-existent inspection ID; whether @Ignore is case-sensitive; whether it suppresses all instances in a procedure or only the annotated line; whether it suppresses across inherited class members; whether suppressed results are excluded from output entirely or returned with a `suppressed: true` flag.

**Risk:** @Ignore is the user's primary escape hatch. If it silently fails (malformed ID, wrong scope), users cannot suppress false positives. If it over-suppresses (suppresses a whole module when line-level was intended), it hides real bugs.

**Recommendation:** Define the @Ignore contract explicitly before implementation. Tests must cover: exact-match suppression; unknown inspection ID (warn or silently ignore?); malformed annotation syntax; scope of suppression (line vs procedure vs module); multiple @Ignore annotations on one item; @Ignore all (wildcard). Each scenario needs a test.

---

### [HIGH] Quick-Fix Testing Is Not Addressed

**Description:** The `InspectionResult` schema includes a `quickFix` field with `description` and `replacement`. The plan mentions quick-fix text in Section 5.2 (extracting from Rubberduck's `.resx` files) but provides no test strategy for quick-fix correctness. There is no mention of testing: whether the replacement string is syntactically valid VBA; whether applying the replacement actually resolves the inspection; whether the replacement is correctly bounded to the reported location; whether multiple quick-fixes on the same file produce consistent results when applied sequentially.

**Risk:** Quick-fix suggestions corrupt code. Users apply a suggestion and introduce a syntax error. This is a trust-destroying failure mode.

**Recommendation:** For every inspection that provides a `quickFix.replacement`, add a "apply-and-recheck" test: take the original code, apply the replacement at the reported location, re-run the inspection, and assert zero results. This confirms the fix actually resolves what it claims to resolve. Add a test that validates the replacement string is parseable VBA.

---

### [MEDIUM] No Regression Testing Strategy for Shared Infrastructure

**Description:** The inspection runner (`runner.ts`), registry (`registry.ts`), and parser facade (`parser/index.ts`) are shared infrastructure used by all 139 inspections. The plan has no strategy for detecting regressions when these components change. Modifying the runner to fix a bug in one inspection could silently break 50 others. The only protection is running all inspection tests, but there is no explicit gate that requires this before merging changes to shared components.

**Risk:** Infrastructure changes cause silent regressions across many inspections. The test suite catches the regression only if someone runs the full suite — and the plan's commit cadence ("commit after each meaningful unit") suggests partial runs are the norm.

**Recommendation:** Define a CI rule: changes to `parser/`, `inspections/base.ts`, `inspections/runner.ts`, or `inspections/registry.ts` must pass the full inspection test suite, not just the test files adjacent to the changed file. Add a smoke test that instantiates every registered inspection and calls it with a minimal valid input, to catch instantiation/registration failures early.

---

### [MEDIUM] Performance Testing Scope Is Inadequate

**Description:** Phase 5 Issue #31 defines performance benchmarks as "Measure parse + inspect time for single files and workspace of 50+ files, optimize if >2s." The targets (single file <100ms, 50-file workspace <5s) are stated but: (1) the baseline is not defined — 50 files of what size?; (2) there is no benchmark for worst-case inputs (e.g., a single 5,000-line module); (3) memory usage is not addressed; (4) the benchmark is deferred to Phase 5, after all inspections are implemented, leaving no performance signal during development; (5) there is no performance regression gate — a future commit could degrade from 100ms to 800ms without any test failing.

**Risk:** Performance regressions are introduced gradually and undetected until the Phase 5 benchmark, by which point the root cause is buried across many commits. Large real-world VBA projects (common in Excel automation shops) may have modules with thousands of lines.

**Recommendation:** Add a lightweight performance gate in Phase 1 (parser only) and Phase 2 (first set of inspections). Define benchmark inputs explicitly: small (100-line module), medium (500-line module), large (2,000-line module). Add memory high-water-mark tracking. Fail CI if any benchmark regresses more than 2x from baseline. Do not defer all performance work to Phase 5.

---

### [MEDIUM] No Test Strategy for Conditional Compilation / Preprocessor

**Description:** The architecture includes `parser/preprocessor.ts` described as a "Conditional compilation resolver." Phase 1 lists it as a "preprocessor stub." VBA conditional compilation (`#If`, `#Else`, `#End If`, `#Const`) is widely used in real-world VBA for cross-host compatibility. The plan has zero test strategy for how conditional compilation affects inspection results — does the inspector see the resolved tree or the raw tree? Do inspections fire on inactive branches?

**Risk:** Inspections fire on dead code branches (code inside `#If False`), or miss active code because the preprocessor stub discards it. Both are silent failures in real codebases.

**Recommendation:** Define the preprocessor contract before Phase 2 inspections are written. Add tests for: code that varies by `#If` branch — which branch is inspected?; `#Const` defined in the same file vs not defined; nested conditional compilation. At minimum, the stub must have a documented and tested behavior (e.g., "all branches are inspected as if active") so inspection authors know what to expect.

---

### [MEDIUM] Encoding and Line Ending Handling Not Tested

**Description:** Real VBA `.bas` and `.cls` files exported from Excel/Access/VBE can have Windows CRLF line endings, BOM markers (`\uFEFF`), and occasionally Windows-1252 encoding. The plan has no mention of encoding normalization, and no test fixtures that exercise non-UTF8 or non-LF input. Line number reporting is especially sensitive to CRLF handling — a CRLF file parsed with LF assumptions will have all line numbers after line 1 be wrong.

**Risk:** `vba/inspect-file` and `vba/inspect-workspace` (the primary real-world tools) produce wrong line numbers or crash on real exported VBA files.

**Recommendation:** Add explicit tests for: CRLF line endings with correct line number reporting; BOM-prefixed files; a fixture exported from actual Excel VBA (if available). Add encoding normalization to the file-reading path and test it explicitly.

---

### [MEDIUM] Test Isolation — Inspection Registry Shared State

**Description:** The inspection registry (`registry.ts`) likely maintains a singleton or module-level catalog. If tests import and use the registry directly, they may share state across test runs — one test enabling an inspection, another expecting it to be disabled. The plan does not address test isolation for the registry or runner. Vitest runs tests in parallel by default, which can expose shared mutable state.

**Risk:** Tests pass individually but fail under parallel execution. Test order-dependence causes intermittent CI failures that are hard to diagnose.

**Recommendation:** Ensure the inspection registry is either immutable after initialization, or that each test constructs its own registry instance. Document this in the test strategy. Add a vitest configuration note specifying whether tests run in parallel or serial, and why. Run the full suite in a shuffled order at least once to check for order-dependence.

---

### [MEDIUM] No Snapshot or Golden-File Testing for AST Output

**Description:** The `vba/parse` tool returns an AST as structured JSON. The plan has no snapshot or golden-file strategy for AST output. This means: a grammar change that silently alters the AST structure will not be caught by any test (inspection tests may still pass if the inspections use stable node accessors); the shape of the JSON returned to MCP clients is never validated against a known-good baseline.

**Risk:** Grammar updates or parser version bumps silently change the AST structure. Downstream tools (Claude Code using the AST) break without any failing test.

**Recommendation:** Add snapshot tests for the `vba/parse` output of at least 5 canonical VBA snippets (Sub, Function, Class, With block, error handler). Use vitest's `toMatchSnapshot()` or a golden-file approach. These snapshots must be reviewed when updated — they are a contract with consumers.

---

### [LOW] The Example Test Pattern Has a Structural Weakness

**Description:** The example test in Section 4.1 uses `toHaveLength(1)` to assert the inspection fired. This assertion passes even if the inspection fires for the wrong reason (e.g., it fires on a comment line that happens to contain an `If` keyword). The example does not assert `results[0].inspection === "EmptyIfBlock"` — meaning if the runner returned a result from a different inspection, the test would still pass.

**Risk:** Tests pass due to accidental coincidence. An inspection that fires on everything passes its positive test even when the logic is completely wrong.

**Recommendation:** The canonical test pattern must always assert: (1) `results.length`; (2) `results[0].inspection` (exact inspection ID); (3) `results[0].severity`; (4) `results[0].location.startLine` (minimum location check). Update the example in Section 4.1 to reflect this.

---

### [LOW] No Test Strategy for `vba/inspect-workspace` Glob Filtering

**Description:** The `vba/inspect-workspace` tool accepts a `glob?` parameter. The plan mentions `.vbalintignore` pattern testing but does not address: glob pattern edge cases (empty glob, invalid glob, glob that matches nothing, glob that matches non-VBA files); behavior when the directory contains symbolic links or non-readable files; behavior when the path does not exist or is a file rather than a directory.

**Risk:** Workspace scanning fails silently or crashes on atypical directory structures.

**Recommendation:** Add explicit test cases for workspace scanning edge cases: empty directory, directory with no matching files, invalid glob syntax, path that does not exist. These can be tested with temporary directory fixtures created in the test setup.

---

### [LOW] Docker Testing Is Underspecified

**Description:** Phase 6 Docker tests are: "Docker image builds successfully; MCP server works correctly from Docker container; All existing tests pass in Docker environment." "Works correctly" is not a test — it is a wish. There is no plan for: what specific MCP protocol exchange is validated in Docker; how stdio transport is tested in a container; whether the container is tested with `docker run` stdio piping specifically; whether the CI pipeline runs Docker tests or only the host tests.

**Risk:** The Docker image builds but the stdio transport is broken inside the container. This is a real failure mode (PATH issues, missing runtime files, wrong working directory) that "image builds successfully" does not catch.

**Recommendation:** Define a specific Docker integration test script: pipe a known JSON-RPC `initialize` request to `docker run --rm -i vba-lint-mcp` and assert the response matches expected structure. This test must run in CI. Document it explicitly.

---

## Summary Table

| # | Finding | Severity | Recommendation |
|---|---------|----------|----------------|
| 1 | Coverage target is functionally meaningless — one positive/negative pair is not a quality bar | CRITICAL | Require location accuracy, inspection ID, and edge case in every positive test |
| 2 | No systematic false positive testing strategy | CRITICAL | Add dedicated false positive test category; track count alongside positive tests |
| 3 | Parse error handling has no test strategy or contract | CRITICAL | Add `error-handling.test.ts`; define and test the error result contract |
| 4 | Integration test strategy is a manual procedure, not an automated suite | CRITICAL | Define `mcp-integration.test.ts` contents explicitly; run in CI without tmux |
| 5 | Fixture file ownership undefined; `kitchen-sink.bas` is a maintenance trap | HIGH | One fixture per test owner; replace `kitchen-sink.bas` with targeted fixtures |
| 6 | Inline vs external fixture decision never made | HIGH | Inline for unit tests; external for integration/workspace tests; document the rule |
| 7 | Location accuracy (line/column) never asserted in tests | HIGH | Mandate location assertion in every positive test; add `expectResultAt` helper |
| 8 | Cross-module symbol resolution tests underspecified | HIGH | Define multi-module test fixture set with explicit cross-module scenarios |
| 9 | @Ignore annotation test strategy absent; contract undefined | HIGH | Define @Ignore contract; test all scope and error edge cases before implementation |
| 10 | Quick-fix correctness never tested | HIGH | Add apply-and-recheck tests for every inspection with a `replacement` field |
| 11 | No regression gate for shared infrastructure changes | MEDIUM | Require full suite on changes to parser, runner, registry, base classes |
| 12 | Performance testing deferred and underscoped | MEDIUM | Add lightweight perf gates in Phase 1/2; define input sizes; add memory tracking |
| 13 | Preprocessor/conditional compilation has no test strategy | MEDIUM | Define preprocessor contract; add conditional compilation test cases in Phase 1 |
| 14 | Encoding and line ending handling not tested | MEDIUM | Add CRLF and BOM fixture files; test line number accuracy with each |
| 15 | Inspection registry shared state may cause parallel test failures | MEDIUM | Ensure per-test registry instances or immutable registry; document isolation model |
| 16 | No snapshot/golden-file testing for AST output | MEDIUM | Add snapshot tests for `vba/parse` output of canonical VBA constructs |
| 17 | Example test pattern does not assert inspection ID or location | LOW | Update example to assert `results[0].inspection` and `results[0].location.startLine` |
| 18 | Workspace glob filtering edge cases not tested | LOW | Add tests for empty dir, no matches, invalid glob, non-existent path |
| 19 | Docker testing specification is vague | LOW | Define specific stdio pipe test; include in CI; document explicitly |

---

## Overall Assessment

**Fail**

The plan has a testing skeleton but not a testing strategy. Four critical gaps — an inadequate coverage definition, absent false positive testing, no parse error contract, and an integration test section that describes a manual procedure rather than an automated suite — mean the testing plan as written will not provide meaningful quality assurance for a 139-inspection linting engine. The plan cannot proceed to implementation until at minimum the four CRITICAL findings and the HIGH findings on location accuracy, quick-fix testing, and @Ignore contract are resolved. The core problem is that the plan conflates "has tests" with "has good tests" — having one positive and one negative test per inspection will produce a green CI badge on a tool that is unreliable in practice.
