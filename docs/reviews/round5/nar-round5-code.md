# NAR Round 5 -- Adversarial Code Review

**Date:** 2026-03-19
**Reviewer posture:** Adversarial (challenge Round 4 assumptions)
**Scope:** Architecture, test quality, security, performance

---

## 1. ARCHITECTURE

### 1.1 Global Function Patches Are Unsafe Under Concurrency -- CRITICAL

**File:** `src/parser/vba-parser-helpers.ts`

`patchParserHelpers()` and `patchLexerHelpers()` write to `globalThis`:

```ts
(globalThis as any).TokenTypeAtRelativePosition = (offset: number): number => {
  return tokenStream.LA(offset);
};
```

The closures capture a *specific* `tokenStream` / `charStream` instance. If two `parseCode()` calls execute concurrently (e.g., workspace scan with `Promise.all`, or MCP tool calls arriving simultaneously on the same server), the second call overwrites the globals before the first finishes parsing. The first parse then reads tokens from the *wrong* token stream.

**Node.js is single-threaded**, so true parallel execution requires worker threads, which the current code does not use. However:
- Nothing prevents a future refactor from adding `Promise.all` over file parses in workspace scanning.
- The current workspace scanner (`inspect-workspace.ts` lines 128-159) processes files sequentially (`for...of` with `await`), which is safe *today*. But this is an accidental safety property, not an intentional design constraint.
- If ANTLR4's parser yields to the event loop (e.g., during very large file parsing), interleaving is theoretically possible even without explicit parallelism.

**Recommendation:** Replace globalThis patching with a scoped mechanism. Options: (a) Use a `WeakMap<VBAParser, () => number>` registry, (b) patch the parser/lexer *instance* prototype chain, or (c) document the single-threaded constraint with a runtime guard.

**Severity:** High (latent defect, not currently triggered)

### 1.2 Token Constants Overwrite globalThis Permanently

Lines 41-46 of `vba-parser-helpers.ts` iterate over `VBAParser` static properties and set them as globals:

```ts
for (const key of Object.getOwnPropertyNames(VBAParser)) {
  const val = (VBAParser as any)[key];
  if (typeof val === 'number' && val > 0 && key === key.toUpperCase()) {
    (globalThis as any)[key] = val;
  }
}
```

These globals are never cleaned up. This pollutes the global namespace with names like `NEWLINE`, `IDENTIFIER`, `IF`, `THEN` -- common names that could collide with other libraries if this server ever runs in a shared process. The constants also only need to be set once (they are static), but `patchParserHelpers` is called on every parse, redundantly re-writing them each time.

**Recommendation:** Set constants once at module load time. Use a `const tokenConstants = Object.freeze({...})` instead of globals, and modify the grammar's semantic predicates to reference them.

**Severity:** Medium

### 1.3 Three Inspections Are Registered But Non-Functional (Stubs)

The following inspections are in the registry and counted in the "65 inspections" claim, but produce zero results for all input:

| Inspection | File | Issue |
|---|---|---|
| `FunctionReturnValueNotUsed` | `declaration/function-return-value-not-used.ts` | `inspect()` has an empty `for` loop body; always returns `[]`. Comment says "needs parse-tree augmentation." |
| `FunctionReturnValueAlwaysDiscarded` | `declaration/function-return-value-always-discarded.ts` | `inspect()` explicitly `return []` with comment "returns empty for now." |
| `IsMissingOnInappropriateArgument` | `parse-tree/error-handling/is-missing-on-inappropriate-argument.ts` | Flags *all* `IsMissing()` calls unconditionally. The inspection name says "inappropriate argument" but it cannot actually check the argument type (Tier A has no type info). The description was softened to a generic advisory, but it is not doing what Rubberduck's equivalent does. |

The first two are explicitly acknowledged as stubs in their test files (tests only check metadata and confirm empty results). But they inflate the inspection count. `IsMissingOnInappropriateArgument` is more misleading -- it fires, but with false positives for every correct `IsMissing(optionalVariantParam)` usage.

**Recommendation:** Either remove stubs from the registry count, or mark them as `enabled: false` in metadata. For IsMissing, rename to `IsMissingUsageAdvisory` or move to Tier B where it can actually check types.

**Severity:** Medium (misleading capability claims)

### 1.4 UnderscoreInPublicClassModuleMember Cannot Detect Class Modules

This inspection (`declaration/underscore-in-public-class-module-member.ts`) flags public members with underscores in names. The Rubberduck original is scoped to class modules only (hence the name). But the current implementation:

1. Has no access to module type information (the `InspectionContext` does not expose whether the current module is a class, standard, or form module).
2. The `Declaration` interface has no `moduleType` field.
3. `isClassModule()` exists in `parser/module-header.ts` but is never called from any inspection.

**Result:** This inspection fires on *all* modules, not just class modules. A standard module with `Public Sub Worksheet_Change()` (an event handler naming convention) would be falsely flagged. The test confirms this -- it tests against a bare `Public Sub My_Method()` without any class module header context.

**Recommendation:** Either rename to `UnderscoreInPublicMember` (dropping the class-module scope claim), or thread module type through the context so it can be checked.

**Severity:** Medium (false positives, misleading name)

### 1.5 ImplicitActiveSheetReference Has Structural False Positives

`parse-tree/excel/implicit-active-sheet-reference.ts` matches any `SimpleNameExpr` whose text is `Range`, `Cells`, `Rows`, or `Columns`. But VBA allows these as variable names:

```vba
Dim Rows As Long
Rows = 10  ' flagged as "implicit ActiveSheet reference"
```

The inspection cannot distinguish between a local variable named `Rows` and an unqualified member access. The Rubberduck original is Tier B for this reason -- it checks whether the resolved symbol is actually the Excel Range/Cells member. This Tier A downgrade is documented with a comment but the inspection severity is `suggestion`, not `hint`, which feels too aggressive for something with known false positives.

Similarly, `ImplicitActiveWorkbookReference` has the same pattern.

**Severity:** Low-Medium (documented limitation, but severity could be lowered)

### 1.6 Visitor Pattern: `visitChildren` After Results May Double-Count

In `EmptyIfBlockVisitor` (line 76), after checking if a block is empty, the visitor calls `this.visitChildren(ctx)`. This is correct for nested `If` statements. However, the visitor overrides `visitIfStmt` using an arrow function property, which in ANTLR4-ng replaces the default traversal. If `visitChildren` visits the same `IfStmt` context again (due to the tree structure), results could be duplicated. This depends on the ANTLR4-ng visitor contract -- whether `visitChildren` on an `IfStmt` calls `visitIfStmt` for child `IfStmt` nodes. In the standard ANTLR4 visitor, `visitChildren` does call `visit()` on each child, which dispatches to the appropriate `visitX` method. So nested `IfStmt` inside ElseIf blocks would correctly be visited. This appears correct but the pattern is fragile -- any inspection that forgets `visitChildren` silently misses nested constructs.

**Severity:** Low (correct today, fragile pattern)

---

## 2. TEST QUALITY

### 2.1 Five Random Inspections -- Spot Check

**Selected (via registry line-number modular sampling):** EmptyIfBlock (A), UnreachableCode (A), ImplicitActiveSheetReference (A), VariableNotUsed (B), ObjectVariableNotSet (B).

| Inspection | Tests correct thing? | Location asserted? | Negative test? | Fixture test? |
|---|---|---|---|---|
| EmptyIfBlock | Yes -- checks `blockStmt().length === 0` | `startLine` checked (line 31) | Yes (line 34) | Yes (lines 40-45) |
| UnreachableCode | Partially -- tests `Exit Sub` case but NOT `GoTo`, `End`, `Resume`, or `Exit For/Do` | No column assertion | Yes (line 103) | Yes (line 109) |
| ImplicitActiveSheetReference | Yes for happy path, but test for "qualified Range" only checks `ws.Range` | No location assertion at all | Yes (line 40) | Yes (line 47) |
| VariableNotUsed | Yes -- uses `buildSingleModuleFinder` correctly | No location assertion (checks `description.toContain('x')`) | Yes (line 40) | Yes (lines 61-72) |
| ObjectVariableNotSet | Yes for known types, but see 2.3 below | No location assertion | Yes (line 46) | No fixture test |

**Key finding:** Only 1 of 5 inspections has a test that asserts location accuracy beyond `startLine`. No test checks `startColumn`, `endLine`, or `endColumn`. This means location data could be wrong (off-by-one, zero vs one-based confusion) without any test catching it.

### 2.2 Tests for Stub Inspections Are Misleading

`FunctionReturnValueNotUsedInspection` and `FunctionReturnValueAlwaysDiscardedInspection` have tests that *pass* by asserting the result is empty. These tests provide zero coverage of the claimed detection logic -- they test that a no-op returns nothing. If someone removes the stub guard, the tests would *fail* (revealing the lack of implementation), but as written, they create an illusion of coverage.

### 2.3 ObjectVariableNotSet: Hardcoded Type List Is Incomplete and Untestable

The inspection checks against a hardcoded set of 23 type names. But:
- Custom classes are not detected (`Dim obj As MyCustomClass` / `obj = something`).
- The type check is case-insensitive but uses lowercase keys -- correct, but not tested with mixed-case type names like `Worksheet` vs `worksheet`.
- More critically: the inspection flags *all* assignments to object variables, including `Set` assignments. The reference resolver marks both `Let` and `Set` statements as `isAssignment = true`. So `Set ws = ActiveSheet` would be flagged as a false positive -- the `Set` keyword is present but the inspection does not distinguish `Set` from `Let`.

Checking `reference-resolver.ts` line 148-153: `enterSetStmt` creates a reference with `isAssignment: true`. The inspection at `object-variable-not-set.ts` line 57 checks `ref.isAssignment` and flags it. This means the inspection fires on *correct* code (`Set ws = ActiveSheet`) as well as incorrect code (`ws = ActiveSheet`).

**This is a real bug.** The IdentifierReference type needs an `isSetAssignment` field, or the inspection needs to check the parse tree context.

**Severity:** High (false positives on correct code)

### 2.4 No Tests for Workspace Scanner Inspection Flow

The `inspect-workspace.ts` handler has no dedicated test file. The workspace flow involves:
- File discovery, symlink resolution, .vbalintignore filtering
- Parse caching
- Cross-module DeclarationFinder construction
- Per-module inspection with shared finder
- Result aggregation and limiting

None of this is integration-tested. The unit tests for individual inspections all use `buildSingleModuleFinder` (single-module), never `WorkspaceContext.getDeclarationFinder()` (cross-module).

**Severity:** Medium (untested integration path)

### 2.5 Tier B Inspections Do Not Actually Test Symbol Resolution

The user asked whether "proof-of-concept Tier B inspections are actually testing symbol resolution or just parse-tree features."

Examined inspections:
- **VariableNotUsed**: Delegates to `DeclarationFinder.findUnusedVariables()`, which checks `references.length === 0`. This *does* depend on Pass 2 (reference resolution) running. If references were not resolved, every variable would appear unused. So it is genuinely testing symbol resolution.
- **HungarianNotation**: Only checks `decl.name` against a regex -- pure declaration metadata, no references needed. Could be Tier A if declarations were extracted without full resolution.
- **IntegerDataType**: Only checks `decl.asTypeName` -- same as above, pure declaration metadata.
- **EncapsulatePublicField**: Checks `decl.accessibility` and `decl.parentScope` -- declaration metadata only.
- **MoveFieldCloserToUsage**: Actually uses `decl.references` to determine which procedures use a variable -- genuinely tests reference resolution.

**Finding:** At least 3 of the 22 Tier B inspections (HungarianNotation, IntegerDataType, EncapsulatePublicField) only need declaration collection (Pass 1), not reference resolution (Pass 2). They could theoretically be a "Tier A.5" or the tier system could distinguish "needs declarations" from "needs references." This is not a bug but an architectural imprecision.

---

## 3. SECURITY

### 3.1 Workspace Scanner: Symlink Resolution Happens AFTER Extension Filtering -- Partial Vulnerability

The workspace scanner flow in `inspect-workspace.ts`:

1. Line 77: `realpath(resolvedPath)` resolves the *directory* symlink.
2. Line 87: `readdir(realDir, { recursive: true })` lists files.
3. Lines 89-92: Filters by extension (`.bas`, `.cls`, `.frm`).
4. Line 131: `readVBAFile(absolutePath)` reads each file.

The question: if an attacker places a symlink `evil.bas -> /etc/passwd` inside the workspace, will it be read?

**Answer: Yes.** The directory path is resolved via `realpath`, but individual *file* paths are not. `readdir` returns the symlink name (`evil.bas`), the extension filter passes it (`.bas`), and `readVBAFile` calls `readFile(filePath)` which follows symlinks by default.

Mitigating factors:
- The 512KB file size limit in `readVBAFile` prevents reading very large files.
- The file content is parsed as VBA, not returned raw -- ANTLR parse errors would garble the output.
- But parse error messages *do* leak fragments of file content (even after sanitization, up to 50 chars of quoted tokens and 200 chars total per error).
- The MCP tool returns parse errors in the response, so an attacker could read ~200 chars of `/etc/passwd` per parse error.

**Recommendation:** Resolve each file path via `realpath()` and verify it remains under the workspace root before reading. This is a standard symlink-following defense.

**Severity:** High (information disclosure via symlink + parse error leakage)

### 3.2 Parse Error Sanitization Is Insufficient

`error-listener.ts` sanitizes quoted tokens to 50 chars and total message to 200 chars. But:
- 200 characters of `/etc/passwd` content is still meaningful for reconnaissance.
- The regex `/'[^']{50,}'/g` only matches single-quoted strings with 50+ chars. Shorter tokens (e.g., username strings) pass through unsanitized.
- ANTLR4 error messages can include the `getText()` of the offending token even outside quotes, e.g., `"missing EOF at 'root:x:0:0:root...'"`.

**Recommendation:** Strip all source-content fragments from error messages. Replace token text with `<token>` placeholders. Return only the structural error description (expected/found token types, line, column).

**Severity:** Medium (amplifies 3.1)

### 3.3 No Path Traversal Guard on Inline `code` Input

The `inspect` tool (`tools/inspect.ts`) accepts raw VBA code as a string parameter. This is safe -- no file system access. But the `parse` tool (`tools/parse.ts`) also accepts `code` as a string and has no file path parameter, so it is also safe. Good.

The `inspect-workspace` tool *does* accept a path. The `path.resolve()` call on line 65 and `stat()` check on line 68 provide basic validation. However, there is no allowlist or jail -- any directory on the filesystem can be scanned. This is a design decision (MCP tools run with the server's permissions), but worth noting.

**Severity:** Low (by design, but should be documented)

---

## 4. PERFORMANCE

### 4.1 SHA-256 Overhead for Parse Cache

The cache (`parser/cache.ts`) computes `SHA-256` of the full content on every `get()` and `set()` call. For a 100-byte VBA snippet, SHA-256 takes ~1-2 microseconds -- negligible. For the 512KB max file size, it takes ~50-100 microseconds -- still negligible vs. the parse time (hundreds of milliseconds).

However, the cache is checked even for inline `code` input via the `inspect` tool, where caching is unlikely to hit (users rarely send identical code twice). The `handleInspectTool` in `tools/inspect.ts` does *not* use the cache -- it calls `parseCode()` directly. Only `handleInspectWorkspaceTool` uses it. So the overhead concern is moot for inline input.

**Finding:** Cache overhead is acceptable. The SHA-256 choice is fine for content-addressed deduplication. A simpler hash (like xxhash) would be marginally faster but SHA-256 is already in Node.js stdlib.

**Severity:** None (acceptable)

### 4.2 All 65 Inspections Instantiated on Every Call

`createAllInspections()` in `registry.ts` line 208 creates `new Cls()` for all 65 inspection classes on every call. This is called from:
- `handleInspectTool` (every inline inspection)
- `handleInspectWorkspaceTool` (once per workspace scan, then reused for all files)

The inspections are stateless (no constructor args, no mutable fields). Instantiation cost is minimal (65 object allocations with prototype chain setup). However:
- The runner then iterates all 65, applying host/category/tier filters at runtime.
- For an inline inspection with `categories: ['Excel']`, 59 non-Excel inspections are instantiated and then immediately skipped.

**Recommendation:** Consider lazy instantiation or a pre-filtered registry. Not urgent -- the 30ms full-inspection benchmark confirms this is not a bottleneck.

**Severity:** None (measurably irrelevant per benchmarks)

### 4.3 Benchmark Flake: Parse Time Exceeds 3000ms Threshold

The test suite has a failing benchmark: "parse time < 3000ms for 500-line module" measured 3111ms. This is a flaky performance test on what appears to be a cold JIT run. The warm-up call exists (`warmUpParser`) but the benchmark is still sensitive to system load.

More concerning: 3 seconds to parse 500 lines of VBA is quite slow. ANTLR4's JavaScript/TypeScript target is known to be significantly slower than the Java target. For workspace scanning of 500 files, this could mean 25+ minutes of parse time (mitigated by caching identical content, but VBA files are typically unique).

**Severity:** Low (known ANTLR4-TS performance characteristic)

### 4.4 Workspace Scanner Reads Files Twice

In `inspect-workspace.ts`, files are read in Phase 1 (lines 128-159) to build the workspace context, then read *again* in Phase 3 (lines 175-206) to run inspections. The cache prevents re-parsing, but `readVBAFile()` is called twice for each file (disk I/O, encoding detection, BOM stripping, CRLF normalization).

**Recommendation:** Store the content string in Phase 1 and reuse it in Phase 3, or store the parsed modules in a structure that preserves the ParseResult per file.

**Severity:** Low (I/O is fast for small files, but wasteful)

### 4.5 DeclarationFinder Uses Linear Scan for findByType and findUnused

`DeclarationFinder` has a name index (`byNameIndex`) but no type index. `findByType(type)`, `findUnused()`, `findUnusedVariables()`, `findUnusedParameters()`, `findNonReturningFunctions()`, and `findVariablesNotAssigned()` all use `this.declarations.filter(...)`, scanning every declaration. For a 500-file workspace with thousands of declarations, each Tier B inspection that calls these methods triggers a full linear scan.

With 22 Tier B inspections and (say) 5000 declarations, this is `22 * 5000 = 110K` comparisons per file, times 500 files = 55M comparisons. Still likely under a second total, but the O(n*m) scaling is poor.

**Recommendation:** Build type-indexed and scope-indexed lookups in the constructor, similar to `byNameIndex`.

**Severity:** Low (not a bottleneck at current scale)

---

## Summary of Findings by Severity

| # | Finding | Severity | Category |
|---|---|---|---|
| 1 | globalThis patches unsafe under concurrency | High | Architecture |
| 2 | ObjectVariableNotSet flags `Set` assignments (false positives) | High | Architecture/Test |
| 3 | Symlink-following reads arbitrary files via workspace scanner | High | Security |
| 4 | 3 stub inspections inflate capability count | Medium | Architecture |
| 5 | UnderscoreInPublicClassModuleMember fires on all modules | Medium | Architecture |
| 6 | Parse error sanitization leaks content fragments | Medium | Security |
| 7 | No integration tests for workspace scanner | Medium | Test Quality |
| 8 | Token constants pollute globalThis permanently | Medium | Architecture |
| 9 | No location accuracy assertions in any test | Medium | Test Quality |
| 10 | ImplicitActiveSheet false positives on variable names | Low-Medium | Architecture |
| 11 | Workspace scanner reads files twice | Low | Performance |
| 12 | DeclarationFinder linear scans for type queries | Low | Performance |
| 13 | Benchmark flake at 3000ms boundary | Low | Performance |
| 14 | 3 Tier B inspections only need Pass 1 | Low | Architecture |
| 15 | Cache SHA-256 overhead | None | Performance |
| 16 | All 65 inspections instantiated per call | None | Performance |

---

## Recommended Priority Actions

1. **Fix ObjectVariableNotSet** (finding 2): Add `isSetAssignment` to `IdentifierReference` and exclude `Set` assignments from the inspection. This is a real false-positive bug.
2. **Add symlink guard** (finding 3): `realpath()` each file and verify it stays under the workspace root.
3. **Document globalThis constraint** (finding 1): At minimum, add a runtime guard that detects concurrent invocation. Better: refactor away from globals.
4. **Mark stubs as disabled** (finding 4): Remove from registry or add `enabled: false` metadata field.
5. **Add location assertions to tests** (finding 9): At least `startLine` + `startColumn` for each inspection's primary test case.
