# NAR Round 3 — Final Review: Feasibility and Real-World Usage

**Reviewer:** Round 3 (Final Gate)
**Date:** 2026-03-19
**Target:** PLAN.md (Implementation Plan)
**Focus:** Feasibility, real-world applicability, protocol correctness, legal compliance, toolchain viability

---

## Executive Summary

This is the final quality gate before implementation begins. Rounds 1 and 2 identified structural issues (TBD inspections, error handling, test quality, registration mechanism, performance). This round focuses on whether the project's core premises hold up under scrutiny. They do not, in several critical ways.

The plan's headline claim — "all 139 Rubberduck v2 inspections" — is misleading. Approximately 30-40 of those inspections are either Rubberduck-specific (tied to its annotation system, folder model, or COM interop) or require runtime type information that a standalone parser cannot provide. The ANTLR4 TypeScript runtime selection is specified incorrectly (`antlr4` package + `antlr4ts-cli` — these are incompatible ecosystems). The MCP tool result model is not addressed: the plan defines a TypeScript `InspectionResult` interface but never specifies how it maps to MCP's `content`/`structuredContent` dual-output model. GPL-3.0 compliance analysis is superficial and misses the key legal question of whether translated logic constitutes a derivative work under copyright law.

The plan should not proceed without resolving the 4 critical and 5 high findings below.

---

## Findings

### 1. ARCHITECTURE

#### [CRITICAL] 1.1 — The "139 Inspections" Claim Is Dishonest About Scope

**Description:** The plan targets "all 139 Rubberduck v2 inspections" (Section 5.1, Section 10 success criteria: ">=130 of 139"). A significant portion of these inspections are not universally applicable outside the Rubberduck/VBE host environment. Specific categories of non-portable inspections:

**Rubberduck-specific (annotation/attribute system):**
- `MissingAnnotationArgumentInspection` — Rubberduck's `@Annotation` system is proprietary; raw VBA has no annotations
- `IllegalAnnotationInspection` — same
- `DuplicatedAnnotationInspection` — same
- `MissingModuleAnnotationInspection` — same
- `AttributeValueOutOfSyncInspection` — Rubberduck tracks attribute/annotation sync; meaningless without Rubberduck
- `MissingAttributeInspection` — depends on Rubberduck's expected-attribute model
- `ModuleWithoutFolderInspection` — Rubberduck's virtual folder system; VBA has no folder concept

**COM type library dependent (require runtime type resolution):**
- `ArgumentWithIncompatibleObjectTypeInspection` — needs COM type library to resolve object types
- `SetAssignmentWithIncompatibleObjectTypeInspection` — same
- `ObjectVariableNotSetInspection` (full version) — needs to know which types are reference types vs value types
- `ImplicitDefaultMemberAccessInspection` — needs to resolve default members from COM type libraries
- `IndexedDefaultMemberAccessInspection` — same
- `RecursiveLetCoercionInspection` — same
- `IndexedRecursiveDefaultMemberAccessInspection` — same
- `LetCoercionInspection` — same
- `ValueRequiredArgumentPassesNothingInspection` — needs type info
- `ObjectWhereProcedureIsRequiredInspection` — needs type info
- `MemberNotOnInterfaceInspection` — needs resolved interface definition
- `ProcedureRequiredByInterfaceInspection` — needs resolved interface definition
- `ImplementedInterfaceMemberInspection` — needs resolved interface definition
- `SuspiciousPredeclaredInstanceAccessInspection` — needs class metadata

**VBE-context dependent:**
- `DefaultProjectNameInspection` — VBE project name; standalone files have no project
- `PublicControlFieldAccessInspection` — UserForm control access; requires form context
- `ExcelObjectNameInspection` — requires knowing which objects are Excel worksheet codenames

This is approximately 25-30 inspections that are either impossible or meaningless without Rubberduck's infrastructure or COM type libraries. The real number of universally portable inspections is closer to **100-110**, not 139.

**Impact:** The success criteria (">=130 of 139") sets an impossible target. The plan will either silently produce stub inspections that never fire (inflating the count while providing zero value), or it will fail its own success criteria and appear incomplete despite doing all the useful work.

**Recommendation:** Audit every inspection against three tiers: (A) fully portable with parse tree only; (B) portable with the planned symbol resolver but no COM type info; (C) requires Rubberduck infrastructure or COM type libraries. Publish the tier classification in the plan. Redefine the success criteria as "100% of Tier A + B inspections implemented." Tier C inspections should be explicitly listed as out-of-scope with a rationale, not silently dropped or faked.

---

#### [HIGH] 1.2 — Symbol Resolution Scope Is Drastically Underestimated for Reference-Based Inspections

**Description:** Phase 3 allocates 4 issues (#16-#19) for the entire symbol resolution system. Phase 4 then piles ~80 inspections on top of it. But the reference-based inspections (items 101-139) require capabilities far beyond what a single-file symbol walker provides:

- **Cross-module reference tracking:** `ProcedureNotUsedInspection` must know if a procedure is called from *any other module* in the workspace. This requires multi-file symbol resolution.
- **Default member resolution:** The entire "Default Member Access" category (items 101-106) requires knowing which classes have `Attribute VB_UserMemId = 0` — information that lives in the compiled type library, not the source.
- **Interface resolution:** `MemberNotOnInterfaceInspection`, `ProcedureRequiredByInterfaceInspection`, `ImplementedInterfaceMemberInspection` require resolving `Implements` statements to their interface definitions, which may be in other modules or in referenced COM libraries.

The plan treats "Declaration Finder" (#18) as a single issue. In Rubberduck v2, `DeclarationFinder` is one of the most complex classes in the entire codebase (~2000 lines of C#). It handles qualified name resolution, ambiguous name resolution, default member chaining, and COM interop type resolution. A single issue is not a credible scope for this.

**Recommendation:** Split Phase 3 into two sub-phases: (3a) single-module symbol resolution, (3b) cross-module/workspace symbol resolution. Identify which Phase 4 inspections require (3a) vs (3b) and sequence accordingly. Explicitly scope out COM type library resolution as a future capability.

---

### 2. TEST QUALITY

#### [CRITICAL] 2.1 — "Claude Code Integration Test via tmux" Is Not a Real Test

**Description:** Section 4.3 describes the integration test as:

> 1. Start MCP server in a tmux pane
> 2. Launch Claude Code in another pane with the MCP configured
> 3. Send VBA code inspection requests
> 4. Capture and validate responses

This is a manual procedure, not an automated test. It cannot run in CI. It depends on Claude Code being installed and authenticated. It depends on Claude choosing to call the tool (which is non-deterministic). It depends on tmux being available. "Capture and validate responses" from a tmux pane requires screen-scraping, which is fragile and non-deterministic.

Issue #30 ("Claude Code integration test") is allocated a single issue for what is functionally a manual QA procedure wrapped in automation theater.

**Impact:** The plan claims integration testing as a quality gate (Section 6.3: "Integration test passes — Claude Code loads MCP, runs inspections") but the test cannot be automated or made reproducible.

**Recommendation:** Replace the tmux-based integration test with two concrete, automatable strategies:
1. **MCP protocol integration test (automated, CI-compatible):** Spawn the server as a child process, connect via stdio, send JSON-RPC `tools/call` messages, validate JSON-RPC responses. This is a real integration test that exercises the full stack without requiring Claude Code. The MCP SDK supports programmatic client construction — use it.
2. **Claude Code smoke test (manual, documented):** Write a step-by-step manual test procedure with expected outputs. Label it as a manual acceptance test, not an automated integration test. Do not count it toward automated coverage.

---

#### [HIGH] 2.2 — ">90% Grammar Rules Exercised" Is Unmeasurable as Stated

**Description:** Section 4.4 states the parser coverage target as ">90% of grammar rules exercised." ANTLR4's TypeScript runtime does not provide grammar rule coverage instrumentation. There is no built-in mechanism in any ANTLR4 target to report which parser rules were entered during a parse. Achieving this metric would require:

1. Instrumenting every rule entry in the generated parser (modifying generated code — fragile and breaks on regeneration), or
2. Writing a custom `ParseTreeListener` that records which rule contexts appear in the parse tree after parsing, then comparing against the full rule list from the grammar.

Option (2) is feasible but non-trivial, and the plan does not mention it. The metric cannot be verified without an explicit measurement mechanism.

**Recommendation:** Either (a) define the measurement mechanism — a test utility that walks parsed trees and reports which `ParserRuleContext` subclasses were instantiated, compared against the full set generated from the grammar — or (b) replace the metric with something measurable: "test fixtures collectively cover Sub, Function, Property Get/Set/Let, Class, standard module, Enum, Type, WithEvents, Implements, conditional compilation, and all control flow structures." An enumerated checklist is verifiable; a percentage against an unmeasured denominator is not.

---

### 3. MAINTAINABILITY

#### [HIGH] 3.1 — 400+ Files With No Grouping Strategy Will Become Unnavigable

**Description:** The plan proposes ~139 inspection files, ~139 test files, and ~139+ fixture files. The directory structure (Section 2.1) shows two levels of grouping: `inspections/parse-tree/empty-blocks/`, `inspections/parse-tree/obsolete-syntax/`, etc. But the grouping is inconsistent:

- `parse-tree/` has 5 subdirectories for ~51 inspections (~10 per directory — manageable)
- `declaration/` has 4 subdirectories for ~50 inspections (~12 per directory — manageable)
- `reference/` has 3 subdirectories for ~30 inspections (~10 per directory — manageable)

The real problem is the test/fixture mirror. The plan shows `test/fixtures/` with 7 subdirectories but `test/inspections/` with only 3 test files. Round 1 already flagged this (the 3-file collapse). If the plan adopts per-inspection test files (as Round 1 recommends), the test directory mirrors the source directory — which is correct but means the fixture directory also needs to mirror it, or fixtures need a clear naming convention that maps fixture to inspection without directory structure.

**Recommendation:** Adopt a co-location pattern: each inspection directory contains its own `__tests__/` subdirectory and `__fixtures__/` subdirectory. Example: `src/inspections/parse-tree/empty-blocks/__tests__/empty-if-block.test.ts` and `src/inspections/parse-tree/empty-blocks/__fixtures__/empty-if.bas`. This keeps related files adjacent, avoids a separate `test/` tree that diverges from `src/`, and makes it obvious when an inspection is missing its test. This is a structural decision that must be made before Phase 2, not discovered during implementation.

---

### 4. GPL-3.0 COMPLIANCE

#### [CRITICAL] 4.1 — "Translated Logic" Derivative Work Analysis Is Missing

**Description:** Section 9.1 categorizes the derivation of inspection files as "Logic translated C# -> TypeScript." Section 9.2 lists the required attribution as LICENSE + ATTRIBUTION.md + source file headers + README section. This is necessary but not sufficient, and the plan does not address the core legal question.

**The question:** Is a line-by-line logic translation from C# to TypeScript a derivative work under copyright law?

**The answer is almost certainly yes.** Copyright protects the expression of an algorithm, not just verbatim text. A faithful translation of Rubberduck's inspection logic — preserving the same conditional structure, the same visitor pattern, the same edge case handling — is analogous to translating a novel from English to French. The translation is a derivative work of the original. This means:

1. **The entire project must be GPL-3.0.** The plan gets this right (LICENSE is GPL-3.0). But the implication is that any consumer of this MCP server — including Claude Code configurations that ship it — must comply with GPL-3.0 terms. This is a real constraint the plan should acknowledge.

2. **ATTRIBUTION.md must include specific copyright notices.** GPL-3.0 Section 5 requires that you "keep intact all notices" including copyright notices. The plan says "Detailed list of derived works, Rubberduck contributor acknowledgment" but does not specify that each derived file must retain the original copyright notice. Rubberduck's source files carry `Copyright (C) Rubberduck Contributors`. This notice must appear in every translated inspection file, not just in a central ATTRIBUTION.md.

3. **Source availability.** GPL-3.0 Section 6 requires that when you convey a covered work, you must make the "Corresponding Source" available. For a Docker image (Phase 6), this means the Dockerfile must not strip source, or the image must include a source-availability mechanism.

**What the plan gets wrong:** It treats attribution as a courtesy ("Detailed Rubberduck contributor acknowledgment") rather than a legal obligation with specific requirements. The ATTRIBUTION.md issue (#33) is in Phase 6 — the last phase. Attribution obligations apply from the first commit that contains derived code.

**Recommendation:**
1. Move attribution setup from Phase 6 (#33) to Phase 1 (#1). The LICENSE file, copyright headers, and source file attribution template must exist before any derived code is written.
2. Define the exact header template for derived files. Example: `// Derived from Rubberduck VBA - Copyright (C) Rubberduck Contributors - GPL-3.0`
3. Add a note in the README that consumers of this MCP server via Docker or npm must comply with GPL-3.0. This affects adoption.
4. Acknowledge in the plan that the "logic translated" files are derivative works and that the GPL-3.0 obligation is not merely voluntary attribution but a legal requirement.

---

### 5. ANTLR4 TYPESCRIPT REALITY CHECK

#### [CRITICAL] 5.1 — The Plan Specifies Incompatible ANTLR4 Packages

**Description:** Section 2.3 (Key Dependencies) lists:

| Package | Purpose | Version |
|---|---|---|
| `antlr4` | ANTLR4 TypeScript runtime | ^4.13 |
| `antlr4-tool` or `antlr4ts-cli` | Grammar -> TypeScript generation | build-time |

These are three different, mutually incompatible ecosystems:

1. **`antlr4` (npm)** — The official ANTLR4 JavaScript runtime. Version 4.13.2 (last updated Aug 2024). Generates JavaScript, not TypeScript. The "TypeScript target" is actually JavaScript output with bundled `.d.ts` type declarations. The generated parser code is JavaScript. The code generation tool is the Java-based `antlr4` CLI (requires JRE), not an npm package.

2. **`antlr4ts-cli`** — Part of the `antlr4ts` project, which was a community TypeScript target. Last published: 2021. Project is abandoned. It generates actual TypeScript code but uses a completely different runtime API (`antlr4ts` not `antlr4`). **You cannot use `antlr4ts-cli` to generate code for the `antlr4` runtime. They are incompatible.**

3. **`antlr4ng`** — The actively maintained community fork (last published: March 2025, version 3.0.16). Generates native TypeScript. Has its own runtime (`antlr4ng`). Uses the `antlr4ng-cli` tool for code generation. This is the only actively maintained option that produces actual TypeScript output.

The plan's dependency table is internally contradictory: it pairs the `antlr4` runtime with the `antlr4ts-cli` generator, which will not work. The risk register mentions falling back to `antlr4ng` but treats it as a backup rather than the primary choice.

**Current state of the ecosystem (as of March 2026):**
- `antlr4` (official): JS-only output, requires JRE for generation, type stubs available. Stable but produces JS, not TS.
- `antlr4ts`: Dead. Last release 2021. Do not use.
- `antlr4ng`: Actively maintained, native TS output, no JRE required (WASM-based generator). Best option for a TypeScript project.

**Recommendation:** Replace the dependency table with:

| Package | Purpose | Version |
|---|---|---|
| `antlr4ng` | ANTLR4 TypeScript runtime | ^3.0 |
| `antlr4ng-cli` | Grammar -> TypeScript generation | ^2.0 (build-time) |

Remove all references to `antlr4` and `antlr4ts-cli`. Update the risk register to remove the "fall back to antlr4ng" mitigation — `antlr4ng` should be the primary and only choice. Add a Phase 1 validation step: "Confirm Rubberduck's VBA grammar compiles with `antlr4ng-cli` without modification."

---

### 6. MCP PROTOCOL CORRECTNESS

#### [HIGH] 6.1 — Result Schema Does Not Map to MCP Content Model

**Description:** Section 1.3 defines `InspectionResult` as a TypeScript interface with structured fields (inspection, description, severity, location, quickFix). But MCP tools return `CallToolResult`, which contains:

```typescript
{
  content: ContentBlock[];       // Array of { type: "text", text: string } or { type: "image", ... }
  structuredContent?: object;    // Optional structured JSON (requires outputSchema in tool registration)
}
```

The plan never specifies how `InspectionResult[]` becomes a `CallToolResult`. The options are:

1. **Text-only:** `JSON.stringify()` the results array into a single text content block. Simple but loses structure for programmatic consumers.
2. **Structured output:** Use `structuredContent` with the `InspectionResult[]` as the value, plus a human-readable text summary in `content`. This is the correct approach per the MCP SDK, but requires defining a Zod `outputSchema` at tool registration time.
3. **Multiple text blocks:** One content block per diagnostic. Unusual and not how MCP tools typically work.

The plan also does not address:
- What `content` text looks like for human/LLM consumption (formatted? one-liner per diagnostic? markdown?)
- Whether `structuredContent` is used (it should be — Claude Code can consume structured data)
- How zero-result runs are reported (empty array? success message?)
- How parse errors are reported (as inspection results? as MCP error responses?)

**Impact:** Without this mapping defined, each tool handler implementer will invent their own serialization. The `vba/inspect` tool and `vba/inspect-workspace` tool will likely produce inconsistent output formats.

**Recommendation:** Add a "MCP Response Format" section to the plan that specifies:
1. Every tool registers an `outputSchema` (Zod) matching the `InspectionResult[]` shape.
2. `structuredContent` contains the typed result array.
3. `content` contains a single text block with a human-readable summary (e.g., "Found 3 warnings, 1 suggestion in Module1.bas").
4. Zero results return `structuredContent: { results: [] }` with content text "No issues found."
5. Parse errors are returned as an inspection result with `inspection: "ParseError"`, `severity: "error"`, not as MCP-level errors (which would prevent partial results from other valid files in workspace scans).

---

#### [HIGH] 6.2 — `vba/inspect-workspace` Output Size Is Unbounded

**Description:** The `vba/inspect-workspace` tool scans a directory tree and returns aggregated results. For a workspace with 100 VBA files, each triggering 5-10 inspections, the result could contain 500-1000 inspection results. Serialized as JSON, this could easily exceed 100KB.

MCP does not define a maximum response size, but MCP clients (including Claude Code) have context window limits. A 1000-result JSON blob will consume significant context and may not be useful to an LLM — it cannot meaningfully process 1000 diagnostics in a single response.

**Recommendation:** Add pagination or summarization to `vba/inspect-workspace`:
- Default behavior: return a summary (count by severity, count by file, top 10 most common inspections)
- Optional `detailed: true` parameter: return full results
- Optional `limit` parameter: cap the number of results returned
- Optional `file` parameter: filter to a single file within the workspace (for drill-down)

This is a design decision that affects the tool schema and must be decided before Phase 2, not discovered when someone runs the tool on a real codebase.

---

### 7. ADDITIONAL FINDINGS

#### [MEDIUM] 7.1 — Preprocessor Handling Is a Stub With No Path to Completion

**Description:** Section 2.1 lists `preprocessor.ts` as "Conditional compilation resolver" and Phase 1 Issue #3 calls it a "preprocessor stub." VBA conditional compilation (`#If ... #Then ... #Else ... #End If`) is common in real-world VBA code, especially in Excel add-ins that target multiple Office versions. If the preprocessor is a stub, any code inside `#If` blocks will either (a) not parse correctly, or (b) parse as dead code that triggers false positives from inspections.

Rubberduck v2 has a separate `VBAPreprocessorParser.g4` grammar (listed in Section 2.1) and a multi-pass parse strategy: preprocess first, then parse the resolved code. The plan lists the grammar file but never describes the preprocessing pipeline.

**Impact:** Real-world VBA files with conditional compilation will produce incorrect results — either parse errors or false positive inspections on inactive code branches.

**Recommendation:** Define the preprocessor strategy explicitly: (a) ignore preprocessor directives entirely and accept false positives as a known limitation, documented in README; or (b) implement a two-pass parse (preprocess -> resolve -> parse resolved code) in Phase 1. Option (a) is acceptable for v1 if documented. The current plan's silence is not acceptable.

---

#### [MEDIUM] 7.2 — No VBA File Format Handling (Module Headers, Attributes)

**Description:** Exported VBA files (`.bas`, `.cls`, `.frm`) contain header lines that are not valid VBA syntax:

```
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyClass"
Attribute VB_GlobalNameSpace = False
```

The ANTLR4 VBA grammar from Rubberduck handles these via its attribute grammar rules, but the plan does not mention file format handling. If the parser is fed a raw `.cls` file including its header, the parse may fail or produce unexpected results.

**Recommendation:** Add a file-format normalization step to the parser facade: strip or parse module headers before feeding code to the main grammar. This is a Phase 1 concern, not a Phase 5 discovery.

---

#### [MEDIUM] 7.3 — `host` Parameter Design Is Underspecified

**Description:** Multiple tools accept a `host?` parameter (excel/project/generic). Section 1.2 lists it but never defines what it controls. Questions unanswered:
- Does `host: "excel"` enable Excel-specific inspections, or does it disable non-Excel inspections?
- What does `host: "generic"` mean? All non-host-specific inspections?
- What is `host: "project"`? VBA project as opposed to Excel? Access? Word?
- Is the host parameter per-file or per-workspace?
- What happens when `host` is omitted — are all inspections run, or only generic ones?

**Recommendation:** Define host semantics explicitly. Suggested model: `host` is an additive filter. `generic` (default) runs only universally applicable inspections. `excel` runs generic + Excel-specific. `access` runs generic + Access-specific. Omitting `host` defaults to `generic`. Document this in the tool parameter descriptions.

---

#### [LOW] 7.4 — No Version Pinning Strategy for Grammar Files

**Description:** The plan says "Copy grammars from Rubberduck v2" but does not specify which commit, tag, or release. Rubberduck's grammar has changed across releases. If the grammar is copied from HEAD and Rubberduck later modifies it, there is no record of which version was used.

**Recommendation:** Pin the grammar source to a specific Rubberduck release tag or commit SHA. Record it in ATTRIBUTION.md and as a comment in the grammar files themselves.

---

#### [LOW] 7.5 — Docker Health Check Is Mentioned But Not Designed

**Description:** Issue #32 mentions "health check" for the Docker container, but MCP servers using stdio transport have no HTTP endpoint. A health check for a stdio-based process requires a different approach (e.g., process liveness check, or a `tools/list` ping over stdio).

**Recommendation:** Either implement an HTTP health endpoint alongside stdio (adds complexity), or use a simple process-liveness health check (`CMD ["pgrep", "-f", "server.js"]`). Specify which approach in the plan.

---

## Summary Table

| # | Severity | Area | Finding |
|---|---|---|---|
| 1.1 | CRITICAL | Architecture | "139 inspections" claim is ~25-30 inspections too high; ~30 are non-portable |
| 5.1 | CRITICAL | ANTLR4 | Dependency table specifies incompatible packages; must use antlr4ng |
| 4.1 | CRITICAL | GPL-3.0 | Derivative work analysis missing; attribution deferred to Phase 6 |
| 2.1 | CRITICAL | Test Quality | tmux integration test is manual, not automated, not CI-compatible |
| 1.2 | HIGH | Architecture | Symbol resolution scope underestimated; DeclarationFinder is 1 issue for ~2000 LOC equivalent |
| 2.2 | HIGH | Test Quality | ">90% grammar rules exercised" is unmeasurable without custom tooling |
| 6.1 | HIGH | MCP Protocol | InspectionResult does not map to MCP CallToolResult content model |
| 6.2 | HIGH | MCP Protocol | Workspace results are unbounded; no pagination or summarization |
| 3.1 | HIGH | Maintainability | 400+ files need co-location pattern decided before Phase 2 |
| 7.1 | MEDIUM | Architecture | Preprocessor is a stub with no defined path; will cause false positives |
| 7.2 | MEDIUM | Architecture | VBA file format headers not handled; will cause parse failures |
| 7.3 | MEDIUM | MCP Protocol | `host` parameter semantics undefined |
| 7.4 | LOW | Maintainability | Grammar files not pinned to a Rubberduck release |
| 7.5 | LOW | Architecture | Docker health check incompatible with stdio transport |

---

## Verdict

**NOT READY FOR IMPLEMENTATION.** The 4 critical findings must be resolved before any code is written:

1. **Reclassify inspections into portable tiers** and adjust success criteria to reflect reality.
2. **Fix the ANTLR4 dependency specification** — the current table will fail on the first `npm install`.
3. **Move GPL-3.0 compliance to Phase 1** and define the exact attribution template for derived files.
4. **Replace the tmux integration test** with an automated MCP protocol test using the SDK's client library.

The 5 high findings should be addressed before Phase 2 begins, as they affect tool schema design, test organization, and the symbol resolution architecture.
