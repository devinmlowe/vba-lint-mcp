# Non-Advocate Review: Round 4 (Post-Implementation)

**Project:** vba-lint-mcp
**Scope:** Full codebase review -- 6 dimensions, architecture and test-quality weighted primary
**Date:** 2026-03-19
**Reviewer:** Claude Opus 4.6
**Commit base:** First post-implementation review (65 inspections: 43 Tier A, 22 Tier B)

---

## Executive Summary

The codebase is well-structured with a clean separation between parse-tree (Tier A) and symbol-aware (Tier B) inspections. The two-pass symbol resolution pipeline is architecturally sound. However, there is one critical bug (case-sensitive host library matching silently drops all Excel inspections in production), several medium-severity gaps in test coverage, and some production-readiness concerns.

**Findings by severity:**

| Severity | Count |
|----------|-------|
| CRITICAL | 1     |
| HIGH     | 3     |
| MEDIUM   | 7     |
| LOW      | 5     |

---

## 1. ARCHITECTURE (Primary)

### CRITICAL: Host library case mismatch silently drops Excel inspections

**File:** `/src/inspections/runner.ts:61-63`
**Impact:** All 6 Excel-specific inspections are silently skipped in every production invocation.

The runner performs case-sensitive comparison:
```typescript
const hasMatchingHost = meta.hostLibraries.some(h =>
  options.hostLibraries!.includes(h),
);
```

Inspection metadata declares `hostLibraries: ['Excel']` (capital E), but the MCP tool schema defaults to `['excel']` (lowercase). Since `'excel' !== 'Excel'`, the runner skips all Excel inspections. Tests pass because they call `inspection.inspect()` directly, bypassing the runner.

**Fix:** Normalize to lowercase in the comparison: `options.hostLibraries!.includes(h.toLowerCase())` or normalize both sides.

### MEDIUM: `@Ignore` suppression documented but never implemented

**Files:** `/src/inspections/runner.ts:41`, `/src/inspections/types.ts:65-66`, `/src/inspections/base.ts:85`

The runner's JSDoc claims "Suppressed results (via @Ignore) are marked but included." The `InspectionResult.suppressed` field exists but is hardcoded to `false`. No code parses `@Ignore` annotations from VBA comments. This is either misleading documentation or missing functionality.

### MEDIUM: Workspace scanner reads each file twice

**File:** `/src/tools/inspect-workspace.ts:128-160 and 175-207`

Phase 1 reads and parses all files to build the workspace context. Phase 3 then reads the same files again via `readVBAFile()` to get content for the parse cache lookup. While the parse cache prevents re-parsing, the double filesystem read is unnecessary. The workspace context should store or expose parse results directly.

### LOW: `globalThis` pollution from parser helpers

**File:** `/src/parser/vba-parser-helpers.ts:26-46`

`patchParserHelpers` writes token type constants and helper functions directly to `globalThis`. This is necessary for the ANTLR4 semantic predicates but creates a collision risk if multiple parser instances exist or if integrated into a larger process. The current single-instance MCP server model makes this low risk, but it should be documented as a known constraint.

### LOW: Inspection instances re-created per tool call

**File:** `/src/tools/inspect.ts:27`, `/src/tools/inspect-workspace.ts:171`

`createAllInspections()` instantiates all 65 inspection classes on every tool call. The inspections are stateless (they receive context via `inspect(context)`) and could be singleton instances. Minor allocation overhead in practice but worth noting for workspace scans where it happens per file.

---

## 2. TEST-QUALITY (Primary)

### HIGH: No integration test exercises the runner with host filtering

No test passes `hostLibraries` through the runner to verify that Excel inspections are included/excluded correctly. This is exactly how the CRITICAL case-sensitivity bug went undetected. The `runner-tierb.test.ts` tests tier skipping but never host filtering.

**Recommendation:** Add runner integration tests that:
1. Run all inspections with `hostLibraries: ['excel']` and verify Excel inspections produce results
2. Run with `hostLibraries: ['access']` and verify Excel inspections are skipped
3. Run with `hostLibraries: undefined` and verify Excel inspections are included (current behavior: they are, since both sides must be defined for the filter to activate)

### HIGH: Tests do not assert location coordinates

Most tests assert `results.length` and `results[0].inspection` (the ID). Only the `empty-if-block.test.ts` pioneering test asserts `location.startLine`. No test asserts `startColumn`, `endLine`, or `endColumn`. Location accuracy is critical for quick-fix edits and IDE integration.

**Recommendation:** At minimum, each inspection's primary positive test should assert `startLine` and `endLine`.

### MEDIUM: No test for `handleInspectTool` or `handleInspectWorkspaceTool`

The tool handler functions that compose parsing + symbol resolution + runner + response formatting are untested. This is the integration seam where the host library default, severity filtering, and response shape all come together.

### MEDIUM: No test for `validateRegistry()`

The registry validation runs at startup and catches missing metadata, duplicate IDs, and invalid tiers. It should have a test confirming it returns no errors for the current registry and detects simulated errors.

### LOW: Test helper `inspectCode` is duplicated across 6+ test files

Each test suite declares its own `inspectCode` or `inspectCodeWithSymbols` helper. A shared test utility would reduce duplication and ensure consistent setup.

---

## 3. MAINTAINABILITY

### MEDIUM: Adding a new inspection requires touching 2 files

A new inspection requires: (1) the implementation file, (2) an import + array entry in `registry.ts`. The barrel-file approach provides compile-time safety but `registry.ts` is already 260 lines with 65 inspections. At scale (100+ inspections), this file becomes unwieldy.

**Observation:** This is a conscious trade-off (compile-time safety over convention-based discovery). The current approach is reasonable for the current scale. Consider auto-registration via a build step if the count exceeds ~100.

### LOW: `(this.inspection as any).createResult` pattern in visitors

**Files:** `/src/inspections/parse-tree/empty-blocks/empty-if-block.ts:63`, `/src/inspections/parse-tree/excel/implicit-active-sheet-reference.ts:58`

Visitor classes cast the inspection to `any` to access the `protected createResult` method. This bypasses TypeScript's access control. A cleaner pattern would be a public `createResult` method on the base class or passing a result-factory function to the visitor constructor.

---

## 4. SECURITY

### MEDIUM: Workspace scanner does not restrict path scope

**File:** `/src/tools/inspect-workspace.ts:64-84`

The workspace scanner resolves the path and follows symlinks but does not restrict where the path can point. An MCP client could pass `path: "/"` and the tool would attempt to `readdir` the entire filesystem (limited only by the 500-file cap). There is no allowlist, sandbox boundary, or restriction to prevent scanning sensitive directories.

**Mitigating factors:** The 500-file cap limits blast radius. VBA file extension filtering (`.bas`, `.cls`, `.frm`) limits what content is read. The tool only reads file content, it does not write. The MCP protocol itself provides a trust boundary.

**Recommendation:** Consider requiring the path to be within the MCP client's workspace root if one is provided.

### LOW: Error messages may leak filesystem paths

**File:** `/src/tools/inspect-workspace.ts:70-73`

Error responses include `resolvedPath` which could reveal the server's filesystem structure to the MCP client:
```typescript
text: `Path is not a valid directory: ${resolvedPath}`
```

In a trusted MCP-client relationship this is informational, but worth noting.

---

## 5. PERFORMANCE

### MEDIUM: Parse cache uses SHA-256 hashing for every lookup

**File:** `/src/parser/cache.ts:84-86`

Every `parseCache.get()` call computes a SHA-256 hash of the entire file content. For workspace scans, this means hashing each file twice (Phase 1 and Phase 3). SHA-256 is ~300MB/s on modern hardware, so this is fast for individual files but adds measurable overhead for large workspace scans.

**Recommendation:** Consider caching by file path + mtime for workspace scans, or storing the hash alongside the content during Phase 1.

### LOW: `DeclarationFinder` query methods use linear scans

**File:** `/src/symbols/declaration-finder.ts:46-48`

`findByType()`, `findUnused()`, `findVariablesNotAssigned()`, etc. all iterate the entire declarations array. For single-module analysis this is fine. For workspace scans with thousands of declarations, additional indices (by type, by scope) would help.

---

## 6. PRODUCTION-READINESS

### HIGH: Registry validation errors do not prevent startup

**File:** `/src/server.ts:23-26`

```typescript
const registryErrors = validateRegistry();
if (registryErrors.length > 0) {
  logger.error({ errors: registryErrors }, 'Inspection registry validation failed');
}
```

If the registry has duplicate IDs, missing metadata, or invalid tiers, the server logs an error but continues serving. A client would receive silently incorrect results (e.g., duplicate inspection IDs would cause confusing diagnostics).

**Recommendation:** Either fail startup or deregister the invalid inspections.

### MEDIUM: No `unhandledRejection` handler

**File:** `/src/server.ts:101-111`

The server handles `uncaughtException` and `SIGTERM`/`SIGINT` but not `unhandledRejection`. Async errors in the MCP SDK transport layer or in workspace scanning could crash the process without cleanup.

### LOW: Docker image includes dev dependencies in builder stage

**File:** `/Dockerfile:10`

`npm ci` installs all dependencies including devDependencies in the builder stage. While the runtime stage only copies `node_modules` from the builder (which includes devDependencies), this is actually fine since the runtime stage gets a clean copy. However, the builder could use `npm ci --production` after the build step to verify the production dependency tree is complete, or the runtime stage could do a separate `npm ci --production`.

**Note:** Upon re-reading, the runtime COPY of `node_modules` from builder includes devDependencies in the final image. This bloats the image unnecessarily. A `npm ci --omit=dev` in the runtime stage (or a prune step) would reduce image size.

---

## Summary of Recommended Actions

### Must Fix (before production use)
1. **[CRITICAL]** Fix case-insensitive host library matching in runner
2. **[HIGH]** Add runner integration tests for host filtering
3. **[HIGH]** Fail startup or deregister on registry validation errors

### Should Fix (next iteration)
4. **[HIGH]** Add location coordinate assertions to tests
5. **[MEDIUM]** Add integration tests for tool handlers
6. **[MEDIUM]** Eliminate double file reads in workspace scanner
7. **[MEDIUM]** Add `unhandledRejection` handler
8. **[MEDIUM]** Implement or remove `@Ignore` suppression documentation
9. **[MEDIUM]** Add workspace path scope restriction

### Nice to Have
10. **[LOW]** Extract shared test helpers
11. **[LOW]** Refactor visitor `createResult` access pattern
12. **[LOW]** Optimize parse cache key strategy for workspace scans
13. **[LOW]** Prune devDependencies from Docker runtime image
