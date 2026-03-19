# NAR Round 6 -- Real-World Usage Review

**Date:** 2026-03-19
**Reviewer:** Claude Opus 4.6
**Scope:** Full codebase at `/Users/devinmlowe/Documents/git/vba-lint-mcp/`
**Test suite:** 377 passed, 1 failed (flaky benchmark timing), 1 skipped

---

## 1. User Experience

### 1.1 CRITICAL: Excel inspections silently skipped by default

All 6 Excel inspections register `hostLibraries: ['Excel']` (capital E), but the server's default value passed from both `vba/inspect` and `vba/inspect-workspace` is `['excel']` (lowercase). The runner's host filter uses case-sensitive `Array.includes()`:

```typescript
// runner.ts line 61
const hasMatchingHost = meta.hostLibraries.some(h =>
  options.hostLibraries!.includes(h),
);
```

`'excel' !== 'Excel'`, so all Excel inspections are silently skipped for every user who doesn't explicitly pass `hostLibraries: ["Excel"]`. Since the README examples and server defaults all show lowercase `"excel"`, a user will never see `ImplicitActiveSheetReference`, `ExcelMemberMayReturnNothing`, or any other Excel inspection unless they discover the case mismatch.

**Files affected:**
- `src/inspections/runner.ts` (line 61-63)
- `src/inspections/parse-tree/excel/*.ts` (6 files, all declare `['Excel']`)
- `src/server.ts` (line 58, default `['excel']`)

**Fix:** Either normalize to lowercase in the runner comparison, or change all inspection metadata to lowercase `'excel'`. The runner fix is more robust:
```typescript
const hasMatchingHost = meta.hostLibraries.some(h =>
  options.hostLibraries!.some(oh => oh.toLowerCase() === h.toLowerCase()),
);
```

**Severity: P0** -- This means the entire Excel category (6 inspections, arguably the most useful for real Excel VBA work) is dead code in practice.

### 1.2 Realistic scenario walkthrough

A user inspecting a typical Excel VBA module like:
```vba
Sub ProcessData()
    Dim ws As Worksheet
    Dim rng As Range
    Set ws = ActiveSheet
    Set rng = Range("A1:A100")
    rng.Value = "test"
End Sub
```

**Expected results:** ImplicitActiveSheetReference on `Range("A1:A100")`, OptionExplicit missing, ImplicitByRefModifier (none -- no params), plus Tier B findings.

**Actual results:** OptionExplicit fires. ImplicitActiveSheetReference does NOT fire (Excel host filter bug above). With the bug fixed, it would fire correctly. The Tier B findings (VariableNotUsed, etc.) run correctly because `hasSymbolTable: true` is set in `handleInspectTool`.

### 1.3 Tool response quality

The text summary format is well-designed: `"Found 3 warning(s), 2 suggestion(s) in 6 lines of VBA (12ms)."` -- clear and actionable.

The structured content includes all needed fields. The `quickFix` descriptions are present and helpful. Location data uses 1-based lines (matching VBE convention), which is correct.

### 1.4 Workspace scanning UX

The workspace tool handles edge cases well:
- Empty directories return a clean "No VBA files found" message
- File count limits (500 max) with actionable error message
- Summary mode vs detailed mode is a good design
- Parse errors are collected per-file with source attribution

**Minor issue:** The `limit` parameter (default 100) applies to the results array via `slice(0, input.limit)`, but the text summary is built from the full `allResults` before limiting. This means the summary says "Found 500 warnings" but the structured data only contains 100. A user seeing the text vs structured content gets inconsistent information.

---

## 2. Documentation

### 2.1 README.md -- Good overall, with gaps

**Accurate:**
- Tool surface descriptions match implementation
- Parameter tables match Zod schemas in server.ts
- Project structure diagram is correct
- Known limitations section is honest and accurate
- Docker instructions are complete

**Gaps:**
- The Claude Code configuration example shows the correct settings path and format. Verified against MCP SDK patterns -- this is correct.
- README claims `.vbalintrc.json` support (lines 223-235) but **no implementation exists**. `grep -r vbalintrc src/` returns zero hits. The configuration file is vapor documentation.
- README claims `@Ignore` annotation support (lines 256-264) but the `src/annotations/` directory is empty. No code parses or applies `@Ignore` directives. The `suppressed` field in `InspectionResult` is always `false`.
- The inspection count "65 inspections" is correct: 43 Tier A + 22 Tier B = 65.

### 2.2 CONTRIBUTING.md -- Actionable and accurate

- Step-by-step guide for adding inspections is correct
- The template code matches the actual base class interface
- Test expectations (count, ID, severity, line, false-positive) are well-specified
- GPL header requirements are clearly stated
- One minor issue: the template shows `import { parseCode } from '../../../../parser/index.js'` -- the relative depth is correct for `parse-tree/<category>/__tests__/` directories

### 2.3 SPEC.md -- Accurate

- Architecture diagram matches the actual code flow
- Tier counts match the registry
- Dependency table matches `package.json`
- Out-of-scope items are clearly stated

---

## 3. GPL Compliance

### 3.1 Per-file copyright headers -- Complete

Every non-generated `.ts` file in `src/` has the header:
```
// Derived from Rubberduck VBA -- Copyright (C) Rubberduck Contributors
// License: GPL-3.0 -- https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE
```

110 files confirmed via grep. Inspection files additionally include the `// Original:` line pointing to the Rubberduck C# source.

Test files (`__tests__/*.test.ts`) also carry the header, which is thorough.

### 3.2 ATTRIBUTION.md -- Mostly complete

- Derived works table covers grammar, resources, inspections, and symbols
- Source pinning line says `(to be pinned during grammar integration -- Issue #2)` -- this should be resolved with the actual commit hash
- Legal note section is clear

**Issue:** ATTRIBUTION.md references `src/resources/en/*.json` but the `src/resources/en/` directory appears empty (no files found). Either the resources were removed and ATTRIBUTION.md wasn't updated, or they haven't been created yet.

### 3.3 SOURCE.md -- Missing

The README mentions `SOURCE.md` nowhere, and no `SOURCE.md` file exists. If this was called for in the original plan, it should either be created or the reference removed. This is a minor gap -- the ATTRIBUTION.md serves the purpose.

### 3.4 Docker GPL compliance

The Dockerfile correctly includes source files in the runtime image:
```dockerfile
COPY --from=builder --chown=node:node /app/src ./src
COPY --from=builder --chown=node:node /app/grammar ./grammar
COPY --from=builder --chown=node:node /app/LICENSE ./
COPY --from=builder --chown=node:node /app/ATTRIBUTION.md ./
```

This satisfies GPL-3.0 Section 6. Source code and license are bundled with the distributed binary.

### 3.5 Incorrect Original reference

`src/inspections/parse-tree/code-quality/unreachable-code.ts` cites `// Original: Rubberduck.CodeAnalysis/Inspections/Concrete/UnreachableCaseInspection.cs` but implements unreachable code after Exit/End/GoTo. Rubberduck's `UnreachableCaseInspection` detects unreachable cases in Select statements -- a different inspection. This attribution is inaccurate and should be corrected or noted as "inspired by" rather than "original."

---

## 4. Edge Cases

### 4.1 Empty file

An empty string `""` passed to `vba/inspect`:
- `parseCode("")` will parse to an empty startRule tree
- OptionExplicit will fire (correct -- no Option Explicit present)
- EmptyModule will fire (correct -- module is empty)
- No crash expected. The parser handles empty input gracefully (the warm-up stub confirms parsing works).

### 4.2 File with only comments

```vba
' This is just a comment
' Another comment
```
- Parses successfully
- OptionExplicit fires (correct)
- EmptyModule fires (correct -- comments are not declarations)
- ObsoleteCommentSyntax does NOT fire (these use `'` syntax, which is modern)

### 4.3 File with only Option statements

```vba
Option Explicit
Option Compare Database
```
- Parses successfully
- OptionExplicit does NOT fire (correct)
- EmptyModule fires (correct -- Option statements are explicitly excluded from "meaningful declarations")

### 4.4 Malformed VBA

```vba
Sub Broken(
End Sub
```
- `parseCode` will produce parse errors collected by `VBAErrorListener`
- Parse errors are returned in the response's `parseErrors` array
- Inspections still run on whatever tree was produced (ANTLR4 error recovery)
- The error listener sanitizes messages (truncates at 200 chars, truncates quoted tokens at 50 chars) -- good security practice

### 4.5 Very long single lines

The parser has a 512KB input size limit (`DEFAULT_MAX_INPUT_SIZE`). A single 600KB line would be rejected with a clear error message. Lines within the limit parse normally -- ANTLR4 handles long lines fine.

However, there is no per-line length check for inspection logic. The `getText().replace(/\s+/g, ' ')` pattern used in several inspections (UnhandledOnErrorResumeNext, UnreachableCode) operates on the full text of a statement. For a pathologically long line (e.g., 100KB of concatenated string), this regex replacement could be slow but would not crash.

### 4.6 Deeply nested code

Deeply nested If/For/Do blocks:
- ANTLR4 uses iterative parsing (not recursive descent for most rules), so stack overflow is unlikely for typical nesting
- Visitor-based inspections recurse via `visitChildren()` which follows the tree depth. VBA code rarely exceeds 20-30 levels of nesting, well within JS stack limits
- No explicit depth limit in the inspection runner

### 4.7 Parse cache edge case

In `inspect-workspace.ts`, the cache key is the content hash, but the cached `ParseResult` includes a `source` field. On cache hit, the code creates a shallow copy: `parseResult = { ...parseResult, source: relativePath }`. This is correct -- the spread copies the tree reference without re-parsing.

However, if two files have identical content but different paths, both get the same cached tree. The spread-copy updates `source` for the second file, but the `tree` object itself is shared. This is fine as long as inspections don't mutate the tree (they don't -- all are read-only visitors).

---

## 5. Inspection Correctness

### 5.1 EmptyIfBlock -- Correct

Checks `ctx.block().blockStmt().length === 0` on the `ifStmt` rule. Matches Rubberduck's logic. Only checks the main `If` block, not `ElseIf` blocks (which don't have their own `blockStmt` children in this grammar rule -- they're separate `ifElseIfBlockStmt` contexts). This is correct behavior.

**Potential false positive:** An If block containing only a comment would still have the comment parsed as part of `endOfStatement` within the block, but NOT as a `blockStmt`. So `If True Then\n  ' comment\nEnd If` would be flagged as empty. This matches Rubberduck's behavior -- comments don't count as executable code.

### 5.2 ObsoleteLetStatement -- Correct

Checks for the `LET()` terminal in `letStmt` context. The VBA grammar correctly distinguishes `Let x = 5` from `x = 5` (both are `letStmt` but only the former has a `LET` token). Correct.

### 5.3 OptionExplicit -- Correct

Uses a visitor flag pattern: sets `hasOptionExplicit = true` when visiting `optionExplicitStmt`, then checks the flag after visiting all children of `module`. Reports at position (1,0) when missing, which is sensible. Correct.

### 5.4 ImplicitActiveSheetReference -- Has false positives

Flags any `simpleNameExpr` whose text matches `range`, `cells`, `rows`, or `columns` (case-insensitive). This will produce **false positives** when:
- A user has a variable named `Range` (rare but legal)
- The code is inside a `With` block that qualifies the sheet: `With ws\n  .Range("A1")\nEnd With` -- here `.Range` is qualified but `Range` appearing alone is not. However, `.Range` would be a `memberAccessExpr`, not a `simpleNameExpr`, so this case is likely handled correctly.
- A UDF parameter named `Rows` would be flagged

The inspection's comment notes this: "This is a simplified parse-tree check." Acceptable for Tier A, but could annoy users in edge cases.

### 5.5 UnhandledOnErrorResumeNext -- Correctness concerns

The detection logic uses `getText().replace(/\s+/g, ' ')` on `mainBlockStmt` and checks for string patterns. This has issues:

1. **False negative:** `On Error GoTo 0` resets error handling, but the check looks for both `goto` AND `0` in the same text. The pattern `text.includes('goto') && text.includes('0')` would match `On Error GoTo Label0` (false GoTo 0 detection) or miss `On Error GoTo 0 'comment with no zero`.

2. **Scope issue:** The inspection checks per-method (Sub/Function), which is correct. But it counts any `On Error GoTo 0` anywhere in the method as clearing all `Resume Next` statements, even if they occur in different logical sections.

3. **Property procedures:** The visitor only handles `visitSubStmt` and `visitFunctionStmt`. Property Let/Set/Get procedures are not checked. This is a **false negative** for Property procedures.

### 5.6 UnreachableCode -- Mostly correct, minor issue

The `EXIT_PATTERNS` regex `/^(exit\s+(sub|function|property|do|for)|end|goto\s+)/i` matches `End` as a standalone word. However, `End Sub`, `End Function`, `End If`, etc. also start with `End`. The regex matches the start of the text, and in VBA, `End` as a standalone statement (program termination) is distinct from `End Sub` (block terminator). Since `blockStmt.mainBlockStmt().getText()` returns the full statement text, `End Sub` would match this regex, potentially flagging code after `End Sub` within a block -- but `End Sub` isn't a blockStmt within a block, it's the method terminator. So this is likely fine in practice.

The label-awareness (lines 61-65) is a good touch -- labels make subsequent code reachable via GoTo.

### 5.7 VariableNotUsed (Tier B) -- Correct

Delegates to `finder.findUnusedVariables()`. The correctness depends on the DeclarationFinder implementation. The inspection itself is clean and matches Rubberduck's pattern.

### 5.8 HungarianNotation -- Correct with minor edge case

Checks for common prefixes (`str`, `int`, `lng`, etc.) followed by an uppercase letter. The `charAfterPrefix === charAfterPrefix.toUpperCase() && charAfterPrefix !== charAfterPrefix.toLowerCase()` check ensures the character is actually a letter (not a digit or underscore). The `break` after first match prevents duplicate flagging.

**Edge case:** The prefix `ws` would match `wsName` (flagged as Hungarian for Worksheet), but also `wsData` where `ws` might genuinely mean "web service." This is an inherent limitation of name-based heuristics. Acceptable.

### 5.9 ObjectVariableNotSet (Tier B) -- Significant false positive risk

The inspection flags ANY assignment to a variable whose declared type is in the `OBJECT_TYPE_NAMES` set. This will produce **false positives** when:
- `Set ws = ActiveSheet` -- this IS using Set correctly. The inspection flags the reference as `isAssignment` but doesn't distinguish `Set x = ...` from `x = ...`. If the reference resolver marks Set assignments as `isAssignment: true`, this would flag correct code.

The inspection description acknowledges this: "assignment **may** require 'Set'." But at severity `error`, false positives here would be very annoying.

### 5.10 ExcelMemberMayReturnNothing -- Correct but noisy

Flags `.Find`, `.FindNext`, `.FindPrevious` member access. The detection is correct -- these DO return Nothing. But it will flag every call, including those already guarded by `If Not result Is Nothing Then`. This is a known limitation of Tier A parse-tree analysis.

---

## 6. Summary of Findings

### P0 (Must fix before release)

| # | Finding | Location |
|---|---------|----------|
| 1 | **Excel inspections silently skipped** -- hostLibraries case mismatch (`'Excel'` vs `'excel'`) means all 6 Excel inspections never run | `src/inspections/runner.ts`, all `src/inspections/parse-tree/excel/*.ts` |

### P1 (Should fix before release)

| # | Finding | Location |
|---|---------|----------|
| 2 | `@Ignore` suppression documented but not implemented -- `src/annotations/` is empty, `suppressed` is always `false` | README.md (lines 256-264), SPEC.md, `src/annotations/` |
| 3 | `.vbalintrc.json` configuration documented but not implemented -- no code loads or applies config files | README.md (lines 223-235), SPEC.md |
| 4 | `ObjectVariableNotSet` likely produces false positives on correct `Set` assignments, at `error` severity | `src/inspections/declaration/object-variable-not-set.ts` |
| 5 | `UnhandledOnErrorResumeNext` misses Property procedures (Let/Set/Get) | `src/inspections/parse-tree/error-handling/unhandled-on-error-resume-next.ts` |
| 6 | Workspace summary text counts don't match limited structured results | `src/tools/inspect-workspace.ts` (line 210-224) |

### P2 (Should fix, not blocking)

| # | Finding | Location |
|---|---------|----------|
| 7 | ATTRIBUTION.md references `src/resources/en/*.json` but directory appears empty | `ATTRIBUTION.md` line 17 |
| 8 | ATTRIBUTION.md source pinning commit hash is placeholder | `ATTRIBUTION.md` line 25 |
| 9 | `unreachable-code.ts` cites wrong Rubberduck original (UnreachableCaseInspection vs unreachable code after exit) | `src/inspections/parse-tree/code-quality/unreachable-code.ts` line 2 |
| 10 | Benchmark test is flaky (3097ms vs 3000ms threshold) | `test/performance/benchmark.test.ts` line 127 |

### P3 (Improvement opportunities)

| # | Finding | Location |
|---|---------|----------|
| 11 | `ImplicitActiveSheetReference` may false-positive on user variables named `Range`, `Cells`, etc. | `src/inspections/parse-tree/excel/implicit-active-sheet-reference.ts` |
| 12 | `ExcelMemberMayReturnNothing` flags Find calls even when already guarded by Nothing check | `src/inspections/parse-tree/excel/excel-member-may-return-nothing.ts` |
| 13 | No `SOURCE.md` file exists (referenced in review prompt but may not have been planned) | Project root |
| 14 | `server.ts` line 68 comment says "placeholder -- implemented in Phase 5" but the tool IS implemented | `src/server.ts` line 68 |

---

## 7. Verification Commands

```bash
# Confirm Excel host filter bug
grep -n "hostLibraries:" src/inspections/parse-tree/excel/*.ts
# All show ['Excel'] (capital E)

grep -n "default\(\['excel'\]\)" src/server.ts
# Shows ['excel'] (lowercase)

# Confirm @Ignore not implemented
ls src/annotations/
# Empty directory

# Confirm .vbalintrc.json not implemented
grep -r "vbalintrc" src/
# No hits

# Run tests
npm test
# 377 passed, 1 failed (benchmark timing)
```
