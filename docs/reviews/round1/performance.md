# NAR Round 1 — Performance Review

**Reviewer:** Performance Dimension
**Date:** 2026-03-19
**Target:** PLAN.md (Implementation Plan)
**Weight:** Secondary

---

## Executive Summary

The plan defers all performance work to Phase 5 (Issue #31) — a single benchmarking issue positioned after the entire inspection catalog is already built. This is backwards: performance constraints that are discovered late are the most expensive to fix. The 100ms single-file and 5s workspace targets may be achievable, but the plan provides no architectural mechanism to reach them, treating performance as an afterthought rather than a design constraint. Several structural choices — synchronous inspection chaining, no parse-tree caching, cold ANTLR4 initialization, and a full-catalog-by-default execution model — combine to create a high probability that Phase 5 benchmarking reveals a gap that requires significant rework of Phase 2–4 code.

---

## Findings

### [CRITICAL] No Caching Architecture Defined

**Description:** The plan mentions "file-level caching" once, in the risk register, as a mitigation for a risk rated "Low." No caching layer is designed, no cache key strategy is specified (path + mtime? content hash?), no eviction policy exists, and no component owns the cache. The `vba/inspect` tool operates on raw code strings with no identity, making caching impossible for the most common tool call. The `vba/inspect-file` and `vba/inspect-workspace` tools operate on paths but the plan never connects path-based calls to a parse-tree cache.

**Risk:** Every `vba/inspect-file` call re-parses from scratch. In a workspace scan of 100 files, if any tool is called repeatedly (e.g., as an LLM iterates on a file), each call incurs full parse overhead. For large modules (1000+ lines), ANTLR4 parse times for VBA grammar — which is non-trivial due to case-insensitive matching and preprocessor rules — can easily exceed 50ms per file, consuming the entire single-file budget before any inspection runs.

**Recommendation:** Define a `ParseCache` component in the architecture (Section 2.1/2.2) before Phase 2 begins. Key on `(filePath, contentHash)`. The cache should live in the parser facade (`vba-parser-facade.ts`) so all tool paths benefit automatically. Bound cache size (e.g., 50 parse trees max) to prevent unbounded memory growth. This is a Phase 1 design decision, not a Phase 5 optimization.

---

### [CRITICAL] Performance Targets Lack Baseline Justification

**Description:** The targets — "single file <100ms, 50-file workspace <5s" — appear in Phase 5 tests with no supporting analysis. There is no reference to ANTLR4 TypeScript runtime benchmarks for VBA-sized grammars, no analysis of what the 100ms budget actually allocates across parse + symbol walk + 139 inspection passes, and no prior art cited. The VBA grammar is large (VBAParser.g4 from Rubberduck is ~800 rules). ANTLR4's JavaScript/TypeScript runtime is measurably slower than the Java runtime — published benchmarks show 3-10x overhead depending on grammar complexity.

**Risk:** If cold parse of a 500-line VBA module takes 80ms in the ANTLR4 TS runtime (a plausible figure given grammar size), the 100ms target leaves only 20ms for 139 inspection passes, symbol table construction, and result serialization. This is likely infeasible without significant optimization that the plan does not plan for.

**Recommendation:** Before Phase 2, run a grammar generation spike (already in Phase 1 as Issue #2) and immediately benchmark parse time for a representative 200-line, 500-line, and 1000-line VBA module. Use those numbers to set realistic targets and revise the architecture if needed. If parse time alone exceeds 60ms for a 500-line file, the 100ms single-file target must be revised or the caching strategy must be treated as mandatory infrastructure, not optional optimization.

---

### [HIGH] ANTLR4 Cold Start Time Not Addressed

**Description:** The ANTLR4 TypeScript runtime initializes grammar prediction contexts, ATN (Augmented Transition Network) serialization, and DFA cache on first use. For a grammar the size of VBAParser.g4, this initialization can take 200-500ms on first parse. Subsequent parses benefit from the warmed DFA cache. The plan has no warm-up strategy — the server will incur this cold-start penalty on the first tool call of each MCP session.

**Risk:** An LLM's first tool call in a session will be noticeably slow. More importantly, if the MCP server is restarted frequently (e.g., during development or when loaded per-session by Claude Code), cold-start hits every restart. Users who make a single inspection call and close the session experience worst-case latency.

**Recommendation:** In the MCP server startup sequence (`server.ts`), add an explicit warm-up step: parse a minimal VBA stub immediately after the server initializes, before accepting tool calls. This amortizes ATN initialization to server startup time rather than first-call time. Document the expected startup latency in the README so users are not surprised.

---

### [HIGH] Inspection Runner Runs All 139 Inspections by Default

**Description:** The plan describes severity and category filtering as tool parameters, but the runner architecture as described in `runner.ts` and Section 2.2 implies all registered inspections are instantiated and run for every call unless filtered. With 139 inspections, even if each inspection's tree traversal is O(n) in AST nodes, the cumulative traversal cost is 139 separate tree walks. The plan does not describe any listener-fusion or single-pass multi-inspection strategy.

**Risk:** 139 independent AST traversals of a large parse tree is the dominant performance cost for parse-tree inspections. A 1000-node parse tree visited 60 times (Phase 2 inspections alone) creates 60,000 node visits per file. For declaration and reference inspections (Phases 3-4), if those also walk the tree independently, the total climbs to 139 × tree-size operations per file, well beyond what can fit in a 100ms budget for non-trivial files.

**Recommendation:** Design the runner to support a listener-fusion model: group inspections by their grammar rule entry points and compose them into a single `ParseTreeListener` pass. Inspections register which grammar contexts they care about; the runner dispatches to all registered handlers during a single tree walk. This is a significant architectural decision that is much harder to retrofit after 139 inspections are written to an independent-traversal model. Define this in Phase 2's `runner.ts` design.

---

### [HIGH] Symbol Table Construction Cost Is Undefined

**Description:** Phase 3 introduces the symbol walker (`symbol-walker.ts`) — a parse-tree listener that builds a full declaration hierarchy including Projects, Modules, Members, Variables, Parameters, Enums, and Types. The plan does not specify whether this walk is performed for every inspection call, only for declaration/reference inspections, or once per parse-and-cache unit. The relationship between symbol table lifetime and parse tree lifetime is unspecified.

**Risk:** If the symbol walker runs on every `vba/inspect` call (even when only parse-tree inspections are requested), it adds a full tree traversal plus allocation cost for every call. Rubberduck's symbol walker is a known performance bottleneck — their async/background parsing architecture exists specifically because symbol resolution is too slow for synchronous use.

**Recommendation:** Specify explicitly in Phase 3 design: (1) the symbol table is built once per parse and cached alongside the parse tree; (2) declaration/reference inspections declare a dependency on the symbol table so the runner can skip symbol construction when only parse-tree inspections are active; (3) symbol table invalidation is tied to parse tree cache invalidation. This must be decided before Phase 3 begins, not discovered during implementation.

---

### [HIGH] Workspace Scan Is Sequential with No Concurrency Model

**Description:** Issue #28 describes `vba/inspect-workspace` as "recursively scan directory for .bas/.cls/.frm files, aggregate results." There is no mention of parallel file processing. Node.js is single-threaded; without Worker threads or explicit async parallelism, file reading, parsing, and inspection of 100 files occurs sequentially.

**Risk:** For a 100-file workspace where each file takes 100ms (the target), sequential processing yields 10 seconds — double the 5-second target for 50 files, and with no headroom for larger workspaces. The plan's 5-second target for 50 files implicitly requires either: (a) sub-50ms average per file, or (b) concurrent processing, or (c) both. Neither is designed for.

**Recommendation:** Decide before Phase 5 whether workspace scanning uses: (a) Promise.all with bounded concurrency (e.g., `p-limit` for 4-8 concurrent parses), or (b) Node.js Worker threads for true parallelism. Option (a) is simpler but still single-threaded for CPU-bound parsing. Option (b) is more complex but necessary if parse time dominates. At minimum, Issue #28 should specify the concurrency model, not leave it implicit.

---

### [MEDIUM] Memory Consumption from Retained Parse Trees Is Unbounded

**Description:** ANTLR4 parse trees for VBA are memory-intensive. Each grammar rule creates a context object; a 1000-line VBA module with dense statements can produce tens of thousands of tree nodes. If parse trees are cached (as recommended above), the cache has no specified memory bound. A workspace scan of 100 large files could retain 100 parse trees simultaneously.

**Risk:** An MCP server running in a long-lived Claude Code session that processes a large VBA codebase could accumulate several hundred megabytes of retained parse trees with no eviction. Node.js heap growth under these conditions leads to GC pressure, increased GC pause times, and eventually OOM if the workspace is large enough.

**Recommendation:** Specify a maximum cache entry count (e.g., 50 trees) with LRU eviction. For workspace scans, process files in streaming fashion rather than loading all parse trees simultaneously: parse, inspect, collect results, then release the parse tree before moving to the next file. Only retain parse trees for files explicitly held in the session cache.

---

### [MEDIUM] No Incremental Parsing Strategy

**Description:** The plan has no incremental parsing concept. Every call to `vba/inspect` on a modified file re-parses the entire module from scratch. In an LLM coding session, the same file may be inspected dozens of times with minor edits between calls.

**Risk:** For large modules (2000+ lines), re-parsing from scratch on every edit creates a latency floor that scales with file size, not edit size. This is the primary reason Rubberduck uses background/async parsing with incremental invalidation.

**Recommendation:** For Phase 1 scope, this is acceptable — full re-parse on every call is a reasonable starting point. However, the plan should explicitly acknowledge this limitation and flag it as a known constraint for future optimization. The content-hash-based cache (finding #1) partially mitigates this: unchanged files are not re-parsed. What remains is the case where the file changes — incremental parsing is genuinely hard and can be deferred, but should be documented as a known gap.

---

### [MEDIUM] No Async Boundary Between MCP Protocol and Inspection Execution

**Description:** The plan describes `runner.ts` as an inspection execution engine but does not specify whether inspection execution is async. MCP tool calls are async at the protocol level, but if the runner is synchronous internally, it will block the Node.js event loop for the full duration of parse + inspect. For a 300ms workspace scan, this blocks all other MCP tool handling.

**Risk:** If Claude Code sends multiple MCP tool calls in quick succession (e.g., listing inspections while a workspace scan is running), the synchronous runner blocks the event loop, causing queued tool calls to appear hung. Node.js MCP servers rely on the event loop for protocol handling; blocking it degrades responsiveness for the session.

**Recommendation:** The runner should be structured as `async` with at least one yield point per file (e.g., `await setImmediate()` or using async file I/O naturally). For workspace scans, yield between files to allow the event loop to process other tool calls. This is a low-cost change that prevents a category of responsiveness bugs.

---

### [MEDIUM] Phase 5 Benchmarking Is Too Late

**Description:** Issue #31 ("Performance benchmarking") is the last technical issue before packaging. At this point, all 139 inspections are implemented, the runner architecture is fixed, and the symbol walker is built. Discovering that the architecture cannot meet targets at Phase 5 means refactoring code written across 4 phases.

**Risk:** Performance findings at Phase 5 that require architectural changes (e.g., listener fusion, caching, concurrency) invalidate a substantial portion of Phase 2-4 work. This is the highest-cost possible discovery point.

**Recommendation:** Add a performance checkpoint at the end of Phase 1 (parse time baseline), Phase 2 (inspection runner overhead with 60 inspections), and Phase 3 (symbol walk overhead). Gate each phase on meeting a sub-budget (e.g., Phase 2 exit: parse + all parse-tree inspections < 80ms for a 500-line file). This converts Phase 5 benchmarking from a potentially catastrophic discovery into a final validation of an already-verified system.

---

### [LOW] `vba/parse` Tool Returns Full AST as Structured JSON

**Description:** Issue #5 specifies that `vba/parse` returns an AST as structured JSON. ANTLR4 parse trees for VBA modules are large: a 200-line module can produce a tree with 2000+ nodes. Serializing this to JSON creates a large MCP response payload.

**Risk:** For large files, the JSON AST may be too large to be useful as an MCP tool response (some MCP clients have response size limits). Serialization time itself adds latency to what should be a fast parse call. Memory doubles during serialization (parse tree + JSON string simultaneously in heap).

**Recommendation:** Define the JSON AST format as a condensed, human-useful representation rather than a full ANTLR4 tree dump. Include only named rule contexts, skip intermediate single-child rules, and represent token text rather than full token objects. Alternatively, support a `depth` parameter to limit serialization depth for exploratory use. This is a Phase 1 design choice that affects the tool's usability and performance permanently.

---

### [LOW] No Performance Regression Protection in CI

**Description:** The test strategy (Section 4) defines coverage targets but no performance regression tests. There is no mechanism to detect if a new inspection or a change to the runner causes a measurable latency increase.

**Risk:** With 139 inspections written across 4 phases by potentially iterative development, performance regressions accumulate silently. A single poorly-written inspection that does O(n²) tree traversal can degrade all workspace scans without any test catching it.

**Recommendation:** Add a performance smoke test to the vitest suite that measures wall-clock time for a fixed VBA fixture and asserts it completes under a threshold (e.g., 200ms with a generous margin for CI variance). This is not a rigorous benchmark but catches catastrophic regressions. Libraries like `vitest-bench` support this pattern.

---

## Summary Table

| # | Finding | Severity | Recommendation |
|---|---------|----------|----------------|
| 1 | No caching architecture defined | CRITICAL | Design ParseCache in Phase 1; key on content hash; bound to 50 entries |
| 2 | Performance targets lack baseline justification | CRITICAL | Benchmark parse time in Phase 1 spike before committing to 100ms target |
| 3 | ANTLR4 cold start time not addressed | HIGH | Add explicit warm-up parse at server startup |
| 4 | Runner executes all 139 inspections independently | HIGH | Design listener-fusion model before writing inspections |
| 5 | Symbol table construction cost undefined | HIGH | Specify symbol table lifecycle and lazy-construction in Phase 3 design |
| 6 | Workspace scan has no concurrency model | HIGH | Specify Promise.all with bounded concurrency or Worker threads before Phase 5 |
| 7 | Retained parse trees are unbounded in memory | MEDIUM | LRU cache with max entry count; streaming workspace processing |
| 8 | No incremental parsing | MEDIUM | Acknowledge as known limitation; content-hash cache partially mitigates |
| 9 | No async boundary in inspection runner | MEDIUM | Yield between files in workspace scans; keep runner async |
| 10 | Phase 5 benchmarking is too late | MEDIUM | Add performance checkpoints at end of Phases 1, 2, and 3 |
| 11 | `vba/parse` returns full ANTLR4 tree as JSON | LOW | Define condensed AST format; support depth limiting |
| 12 | No performance regression protection in CI | LOW | Add wall-clock smoke test to vitest suite |

---

## Overall Assessment

**Fail** — with a path to Pass with Conditions.

The plan is architecturally incomplete on performance. Two critical findings (#1 and #2) and four high findings (#3, #4, #5, #6) represent gaps that, if unaddressed, make the stated performance targets unreachable without retrofitting substantial architecture after most of the code is written. The root cause is that performance is treated as a Phase 5 concern rather than a Phase 1 constraint. The plan must be revised to: (a) establish a parse-time baseline before any inspections are written, (b) define the caching layer and runner execution model as Phase 1/2 architecture decisions, and (c) add per-phase performance gates. With those additions, the plan is viable — the underlying approach (ANTLR4 + Rubberduck port + TypeScript) is sound if the architecture accounts for the performance characteristics of the runtime.
