# NAR Round 1 — Maintainability Review

**Reviewer:** Maintainability Dimension
**Date:** 2026-03-19
**Target:** PLAN.md (Implementation Plan)
**Weight:** Secondary (elevated)

---

## Executive Summary

The plan describes a well-intentioned port of Rubberduck v2's inspection catalog, but it defers the hardest maintainability decisions to implementation time in ways that will create compounding technical debt. Three structural weaknesses stand out: an explicit-registration gap that will silently swallow new inspections, a monolithic JSON metadata file whose update process is undefined, and a test organization scheme that collapses 139 inspections into three test files — a design that will become unnavigable within the first phase.

---

## Findings

### [CRITICAL] Inspection Registration Is Described as Auto-Discovery But the Mechanism Is Unspecified

**Description:** Section 2.1 references `registry.ts` as "Inspection catalog and discovery," and Issue #7 calls it "auto-discovery of inspection classes." However, the plan provides no concrete mechanism for how auto-discovery works in a compiled TypeScript/Node.js environment. Filesystem-based glob discovery at runtime is fragile (breaks after bundling, breaks in Docker if the dist structure changes). Import-based registration requires a central import list, which is effectively explicit registration with extra steps. The plan never resolves this tension.

**Risk:** If the discovery mechanism is wrong, new inspections added by a contributor will silently not appear in results — with no compile-time or test-time error to catch the omission. At 139 inspections, one missed registration per category is a realistic failure rate. This is the highest-probability silent failure mode in the entire design.

**Recommendation:** Before Phase 2 begins, the registry pattern must be decided and documented. Options are: (a) explicit barrel file (`inspections/index.ts`) that re-exports all inspection classes — requires one line per inspection, detectable at compile time via the registry's type system; (b) convention-based glob scan at server startup with a startup-time validation step that warns on zero registered inspections; (c) a code-generation step (e.g., a build script that reads the directory and generates the barrel). Whichever is chosen, the mechanism must be stated in the plan, not left to the implementer.

---

### [CRITICAL] Test Organization Collapses 139 Inspections into Three Test Files

**Description:** Section 2.1 shows the entire test suite as three files: `parse-tree.test.ts`, `declaration.test.ts`, and `reference.test.ts`. With ~60 parse-tree inspections and each requiring at minimum two test cases (positive + negative), `parse-tree.test.ts` will contain 120+ test cases in a single file. By Phase 4, `declaration.test.ts` and `reference.test.ts` add another 160+ cases combined.

**Risk:** A 300+ case monolith is not maintainable: locating a specific inspection's tests requires grep or scroll; failures in one category obscure failures in another; code review diffs for adding a single inspection touch a shared file that everyone else is also modifying concurrently. The fixture directory structure (which is per-category) is better organized than the test file structure, which contradicts itself.

**Recommendation:** Mirror the inspection directory structure in the test directory: `test/inspections/parse-tree/empty-blocks/`, `test/inspections/parse-tree/obsolete-syntax/`, etc., with one test file per inspection or at most one test file per subcategory (e.g., `empty-blocks.test.ts`). The plan's fixture organization already implies this structure — the test files should follow it.

---

### [CRITICAL] inspection-metadata.json Is a Single Point of Maintenance Failure with No Update Protocol

**Description:** Section 5.2 describes extracting metadata from Rubberduck v2's `.resx` files into `src/resources/inspection-metadata.json`. This is a one-time manual extraction with no defined process for keeping it current. The JSON file will contain names, descriptions, severity defaults, and quick-fix text for all 139 inspections. The plan does not describe: what happens when the JSON and an inspection class disagree; whether the JSON is the source of truth or the class is; who owns updates when a description is corrected; or how to detect drift between the JSON and the registered inspection set.

**Risk:** Within six months of launch, the JSON will diverge from the implementations. Inspection descriptions shown to users via `vba/list-inspections` will be stale or wrong. There is no test in the plan that validates JSON completeness against the registered inspection set. New inspections added after initial extraction will either be missing from the JSON (silent gap) or require a manual JSON edit that is easily forgotten.

**Recommendation:** Either (a) move metadata into each inspection class as static properties (the class is the single source of truth, no separate file needed), or (b) add a startup-time validation step that asserts every registered inspection ID has a corresponding JSON entry and fails loudly if not. Option (a) eliminates the sync problem entirely and is preferred. The JSON file is an optimization for fast catalog queries that introduces more maintenance burden than it eliminates.

---

### [HIGH] Directory Structure Has a Two-Level Category Taxonomy That Will Drift

**Description:** The plan defines two levels of categorization: top-level inspection type (`parse-tree/`, `declaration/`, `reference/`) and subcategory (`empty-blocks/`, `obsolete-syntax/`, `code-quality/`, etc.). The subcategory assignment is not formally governed — it is implied by the phase issue groupings. There are already ambiguities: `EmptyMethodInspection` appears in both the parse-tree list (Issue #8) and the declaration list (line 506, "declaration variant"), and `ImplicitActiveSheetReferenceInspection` appears in both Issue #12 (parse-tree) and Issue #25 (reference-based variant). These duplicates will create competing implementations in different directories.

**Risk:** Developers adding new inspections will make inconsistent subcategory choices. Over time, `code-quality/` becomes a catch-all. The duplicate inspection names across phases suggest the design has not resolved whether these are the same inspection running in two modes or two separate inspections — a distinction that affects the registry, deduplication logic, and result IDs.

**Recommendation:** Define a formal taxonomy in the plan that governs subcategory assignment, with examples for ambiguous cases. Resolve the `EmptyMethodInspection` duplicate explicitly — either it's one class with a mode flag, or two classes with distinct IDs. The registry must enforce uniqueness on inspection IDs at startup.

---

### [HIGH] Grammar Maintenance Path for Upstream Rubberduck Changes Is Undefined

**Description:** Section 2.1 treats the grammar files as static copies ("Direct copy" in the attribution table). The plan's risk register mentions grammar modifications but only in the context of initial setup, not ongoing maintenance. Rubberduck v2 is an active project — grammar files change when new VBA constructs or edge cases are discovered. The plan establishes no mechanism for tracking upstream grammar changes, no diff strategy, and no test coverage that would reveal regressions from grammar drift.

**Risk:** The grammar will diverge from Rubberduck's over time. VBA files that Rubberduck parses correctly will produce parse errors in vba-lint-mcp. Because the grammar generates code into `parser/generated/` (gitignored), even detecting the version of the grammar in use requires inspecting the source `.g4` files. There is no pinned version, no upstream tracking, and no process defined.

**Recommendation:** Pin the grammar to a specific Rubberduck commit hash recorded in a `grammar/SOURCE.md` or in package.json metadata. Add a regression test suite using VBA samples from the Rubberduck test corpus — if the grammar changes, these tests will break, alerting maintainers. Define an explicit upgrade procedure: pull new grammar, regenerate, run tests, document the commit delta.

---

### [HIGH] Boilerplate Per Inspection Is Unquantified but Structurally High

**Description:** The plan describes a base class hierarchy (`InspectionBase`, `ParseTreeInspection`) but does not show a complete example of what a single inspection file looks like end-to-end. Each inspection needs: class definition extending base, listener or visitor implementation, result construction, metadata lookup (if JSON-based), host filtering logic, and export. For parse-tree inspections that use ANTLR4 listeners, each inspection likely requires entering the listener/visitor pattern boilerplate even when the actual detection logic is two or three lines.

**Risk:** If each inspection is 60-100 lines of file when the core logic is 5-10 lines, contributors face a high friction addition cost. At 139 inspections, total boilerplate volume is significant. Worse, if the boilerplate pattern is established wrong in Phase 2 and discovered wrong in Phase 4, refactoring 60 already-written inspections is a major maintenance event.

**Recommendation:** Before implementing any inspections in Phase 2, write a complete reference implementation of one simple inspection (e.g., `ObsoleteLetStatementInspection`) and measure its file size, line count, and the ratio of boilerplate to logic. If the ratio exceeds 5:1, redesign the base class to absorb more of the pattern. This reference implementation should be locked in before Phase 2 issues are opened.

---

### [HIGH] Vague Inspection Count (Items 85-100 and 110-139 Are Placeholders)

**Description:** Section 5.1 explicitly defers enumeration of inspections 85-100 and 110-139 to Phase 4 implementation: "will be fully enumerated during Phase 4 implementation." This means 54 of the 139 targeted inspections (39%) have no specification yet.

**Risk:** The plan's scope is materially incomplete. Phase 4 work cannot be estimated, prioritized, or reviewed without knowing what these inspections are. When Phase 4 begins, the implementer must simultaneously do discovery work (what inspections exist, how they work in Rubberduck's C# source, whether they require symbol resolution) and implementation work. Discovery failures will result in missed inspections or incorrectly ported logic. The success criterion of "≥130 of 139 inspections" (Section 10) cannot be meaningfully evaluated against a partial catalog.

**Recommendation:** Complete the inspection catalog before Phase 4 begins — ideally as a Phase 3 prerequisite. This can be done by auditing the Rubberduck v2 source at `Rubberduck.CodeAnalysis/Inspections/Concrete/` and producing a complete list with: inspection name, Rubberduck category, required infrastructure (parse-tree vs declaration vs reference), and estimated complexity. This audit should be its own issue and its output should be appended to Section 5.1.

---

### [MEDIUM] Test Fixtures Are Shared Across Inspections with No Ownership Model

**Description:** The fixture files in `test/fixtures/` are organized by category, not by inspection. A file like `empty-if.bas` is presumably used by `EmptyIfBlockInspection` tests. However, the plan makes no statement about whether fixtures are one-to-one with inspections, whether multiple inspections can share a fixture, or who owns fixture files. The `kitchen-sink.bas` fixture in `test/fixtures/comprehensive/` is explicitly designed to trigger many inspections simultaneously.

**Risk:** Shared fixtures create hidden coupling. A change to `empty-if.bas` to improve one test may cause another inspection's test to unexpectedly pass or fail. The `kitchen-sink.bas` fixture will become the first thing contributors add edge cases to because it's the path of least resistance, making it progressively harder to understand what it's testing. When tests break, diagnosing which inspection's expectation was violated becomes a grep exercise.

**Recommendation:** Adopt a one-to-one fixture policy: each inspection gets its own fixture file (or pair of files: `should-trigger.bas` and `should-not-trigger.bas`). Shared fixtures should be limited to integration-level tests, not unit inspection tests. Inline code strings in test files (as shown in the plan's test example) are preferable over file fixtures for unit tests — they keep the test self-contained and eliminate the fixture ownership problem entirely.

---

### [MEDIUM] Dependency Version Strategy Is Pinned to "latest" for Critical Dependencies

**Description:** Section 2.3 lists `@modelcontextprotocol/sdk` as `latest` and `vitest` as `latest`. The MCP SDK is versioned rapidly — the protocol itself has evolved through breaking changes. `tsx` is also pinned to `latest`. Only `antlr4` and `typescript` have version constraints.

**Risk:** A `npm install` performed six months from now will pull different dependency versions than the initial install. If the MCP SDK introduces a breaking change in its tool registration API, all five tool handlers break simultaneously. Because `latest` is resolved at install time, two contributors with different install dates may be running different SDK behavior with no indication in the lockfile diff that the SDK changed.

**Recommendation:** Pin all dependencies to specific versions in package.json (not `latest`, not `^latest`). Use `>=` ranges only where flexibility is explicitly needed. Add a Dependabot or Renovate configuration file to manage version updates as deliberate decisions. The MCP SDK in particular should be pinned to a specific minor version until the protocol stabilizes.

---

### [MEDIUM] Documentation Maintenance Is Spread Across Four Files with Overlapping Content

**Description:** The plan requires maintaining `README.md`, `SPEC.md`, `ATTRIBUTION.md`, and inline source attribution comments. README is updated at every phase boundary (six updates planned). SPEC.md is Phase 6. ATTRIBUTION.md is Phase 6. The inspection catalog appears in both the README (Section 2 Phase 2 exit criteria) and Section 5.1 of the plan itself. There is no single authoritative source for the inspection catalog.

**Risk:** By the time Phase 6 is reached, README has been updated six times incrementally. SPEC.md is written fresh in Phase 6 against code that was written in Phases 1-5. The ATTRIBUTION.md is written in Phase 6 but attribution comments must be present from Phase 2 onward (per Section 9.2). The lag between when attribution is required (immediately) and when ATTRIBUTION.md is written (Phase 6) creates a compliance gap under GPL-3.0.

**Recommendation:** Write ATTRIBUTION.md in Phase 2, not Phase 6, since that is when Rubberduck-derived code first enters the codebase. SPEC.md should be written alongside Phase 1 (define the spec before building to it), not after. The inspection catalog should be maintained as a single JSON or YAML file and generated into both README and `vba/list-inspections` output — eliminating the README-as-catalog pattern that requires manual updates.

---

### [MEDIUM] Code Duplication Risk Across Structurally Similar Inspections

**Description:** The plan groups inspections into families (empty blocks, obsolete syntax, unused code) where the detection pattern is nearly identical across the family. For example, the nine empty-block inspections differ only in which ANTLR4 grammar rule they target. The six unused-code inspections differ only in which declaration type they query. The plan's base class hierarchy (`InspectionBase`, `ParseTreeInspection`) is intended to address this, but the plan provides no detail on how the base classes will be parameterized to handle family-level variation without copy-paste.

**Risk:** Without a concrete parameterization pattern, each of the nine empty-block inspections will be implemented as a separate class with near-duplicate visitor code. After 60 inspections are written this way, the pattern is baked in and refactoring is expensive. The value of base classes is eliminated if each subclass re-implements the same structural pattern with minor variation.

**Recommendation:** Design a parameterized inspection pattern before Phase 2 starts. For empty-block inspections, a single `EmptyBlockInspection` class parameterized by grammar rule context type would be more maintainable than nine separate classes. Show this pattern in the plan with a concrete example. For families of three or more inspections with identical structure, the base class should handle the structural pattern completely; the concrete class should provide only the distinguishing data (rule type, declaration kind, etc.).

---

### [LOW] Phase 6 Docker and Documentation Work Has No Specified Test for Attribution Compliance

**Description:** Issue #33 (ATTRIBUTION.md) and Issue #35 (Final README) are the last two items before release. The quality gate for Phase 6 is "Docker image builds and works" and "all existing tests pass." There is no test or checklist for attribution completeness: no verification that every file derived from Rubberduck has a source file header, no review of whether all Rubberduck contributors are credited, no check that GPL-3.0 requirements (source availability, license text in distributions) are met.

**Risk:** GPL-3.0 compliance failures are a legal risk, not merely a quality risk. Releasing a Docker image without proper source attribution in the image or documentation violates the license. This is unlikely to be litigated but is a reputational risk and a project integrity issue given the explicit attribution goal in the design principles.

**Recommendation:** Add a pre-release attribution checklist to the Phase 6 quality gate: (a) every file in `src/inspections/` has a Rubberduck source attribution header; (b) ATTRIBUTION.md lists all derived files with their upstream paths; (c) the Docker image or its companion documentation references the source repository and license; (d) LICENSE file is included in the Docker image at a known path.

---

### [LOW] Preprocessor Is Stubbed Indefinitely with No Completion Criteria

**Description:** Section 2.1 includes `parser/preprocessor.ts` with the description "Conditional compilation resolver." Phase 1 Issue #3 describes it as a "preprocessor stub." No subsequent phase issue or exit criterion addresses completing the preprocessor. VBA's conditional compilation (`#If`, `#Const`) is not exotic — it appears in any VBA project that uses version guards or platform conditionals.

**Risk:** A stub preprocessor that does not resolve `#If` blocks will produce parse errors or incorrect parse trees for code using conditional compilation. Inspections run against an incorrect tree will produce false positives or miss real issues. The plan never commits to resolving this, creating a permanent known gap with no disclosure to users.

**Recommendation:** Either (a) explicitly scope out conditional compilation support and document the limitation in the README with the specific patterns that are unsupported, or (b) identify a phase in which the preprocessor stub becomes a real implementation. A known, documented limitation is maintainable; an undated stub is not.

---

## Summary Table

| # | Finding | Severity | Recommendation |
|---|---------|----------|----------------|
| 1 | Inspection registration mechanism is unspecified — auto-discovery will silently drop inspections | CRITICAL | Decide and document the registry pattern before Phase 2; prefer explicit barrel file with compile-time enforcement |
| 2 | Test organization collapses 139 inspections into three files | CRITICAL | Mirror the inspection directory structure in the test directory; one test file per inspection or subcategory |
| 3 | inspection-metadata.json has no sync protocol and no drift detection | CRITICAL | Move metadata into inspection classes as static properties, eliminating the separate file entirely |
| 4 | Two-level category taxonomy is not governed; duplicate inspection names across phases | HIGH | Define formal taxonomy; resolve EmptyMethodInspection and ImplicitActiveSheetReferenceInspection duplicates; enforce unique IDs at startup |
| 5 | Grammar maintenance path for upstream Rubberduck changes is undefined | HIGH | Pin grammar to a commit hash; add regression tests from Rubberduck corpus; document upgrade procedure |
| 6 | Per-inspection boilerplate is unquantified and may be structurally excessive | HIGH | Implement one complete reference inspection before Phase 2; measure boilerplate ratio; redesign if >5:1 |
| 7 | 39% of targeted inspections are unspecified placeholders | HIGH | Complete the inspection catalog as a Phase 3 prerequisite deliverable |
| 8 | Test fixtures have no ownership model; kitchen-sink fixture creates hidden coupling | MEDIUM | One-to-one fixture-to-inspection policy; prefer inline code strings for unit tests |
| 9 | Critical dependencies pinned to "latest" | MEDIUM | Pin all dependencies to specific versions; add automated dependency update tooling |
| 10 | Documentation is duplicated across four files; ATTRIBUTION.md deferred past when it's needed | MEDIUM | Write ATTRIBUTION.md in Phase 2; generate inspection catalog from data rather than maintaining in README |
| 11 | Code duplication risk across inspection families without parameterization plan | MEDIUM | Design parameterized base classes before Phase 2; show concrete example in plan |
| 12 | No attribution compliance test in Phase 6 quality gate | LOW | Add pre-release attribution checklist to Phase 6 quality gate |
| 13 | Preprocessor is a permanent stub with no completion criteria or documented limitation | LOW | Either scope out conditional compilation explicitly in docs or assign it to a phase |

---

## Overall Assessment

**Fail — Revise Before Implementation**

Three critical findings must be resolved before Phase 2 begins: the registration mechanism must be specified, the test organization must be redesigned, and the metadata sync problem must be eliminated. Two high findings must also be resolved at plan level: the incomplete inspection catalog (39% undefined) and the undocumented grammar maintenance path. The plan is sound at a high level — the phasing is reasonable and the Rubberduck source material is solid — but it leaves too many implementation-defining decisions to be discovered at implementation time. Those decisions will shape the maintainability of the codebase for its entire life; they should be made deliberately in the plan, not reactively in the code.
