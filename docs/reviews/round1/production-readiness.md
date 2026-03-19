# NAR Round 1 — Production Readiness Review

**Reviewer:** Production Readiness Dimension
**Date:** 2026-03-19
**Target:** PLAN.md (Implementation Plan)
**Weight:** Secondary

---

## Executive Summary

The plan describes a well-structured, phased MCP server implementation with good inspection catalog coverage and a sensible dependency graph. However, the plan almost entirely omits operational concerns: there is no error handling strategy, no logging design, no resource limits, no encoding policy, no configuration validation, and the Docker and .vbalintignore sections are one-liners with no substance. The server is designed to be embedded in developer workflows where silent failures or stalls will directly degrade AI assistant reliability, making these gaps significant.

---

## Findings

### [CRITICAL] No Error Handling Strategy Defined

**Description:** The plan describes an inspection runner (`inspections/runner.ts`) and five tool handlers but never specifies what happens when a component throws. If a single inspection throws an unhandled exception, does the runner abort all results? Swallow and continue? Return partial results with an error envelope? The plan is silent. For a service with 139 inspections running per call, any individual inspection failure will be common during development and occasional in production.

**Risk:** A single broken inspection silently kills all output, or crashes the MCP process, or causes the MCP client (Claude Code) to hang awaiting a response that never arrives. The MCP protocol has no built-in retry — a hung tool call degrades the entire session.

**Recommendation:** Define the error handling contract explicitly: (1) per-inspection isolation — each inspection runs in a try/catch; errors are captured as a structured diagnostic with `severity: "error"` and an `inspection: "InternalError"` tag; (2) tool-level error responses use MCP's `isError: true` result field with a structured message; (3) parser failures return a parse-error result rather than throwing to the tool handler.

---

### [CRITICAL] MCP stdio Transport Has No Malformed Input Defense

**Description:** Phase 1 (Issue #4) says "stdio transport" but specifies no handling for malformed JSON, truncated messages, oversized payloads, or clients that disconnect mid-stream. The MCP SDK may handle some of this, but the plan never verifies this assumption or adds a defensive layer.

**Risk:** A malformed JSON-RPC message on stdin can crash the server process. Because the server runs as a child process of Claude Code, a crash silently terminates all MCP capabilities for the session with no user-visible error. There is no restart mechanism specified.

**Recommendation:** (1) Verify what the `@modelcontextprotocol/sdk` stdio transport does with malformed input before Phase 1 ships; (2) add a top-level `process.on('uncaughtException')` handler that logs and exits cleanly rather than hanging; (3) document that the MCP client (Claude Code) handles server restarts and under what conditions.

---

### [CRITICAL] No Resource Limits — Memory, File Size, or Timeout

**Description:** `vba/inspect-workspace` can scan an arbitrary directory tree. `vba/inspect-file` accepts an arbitrary path. There are no stated limits on: file size before parsing, number of files in a workspace scan, total memory the ANTLR4 parse tree may consume, or wall-clock timeout per tool call. The performance target (Phase 5: "50-file workspace < 5s") exists but no enforcement mechanism is defined.

**Risk:** A user running `vba/inspect-workspace` on a large project (hundreds of files, or a path that accidentally traverses node_modules or a network drive) will cause the Node.js process to exhaust memory or stall indefinitely. A large VBA file (e.g., a generated module with thousands of lines) may cause ANTLR4 to consume unbounded memory during parse tree construction.

**Recommendation:** Define and enforce: (1) max file size per `vba/inspect-file` and per file in workspace scan (suggested: 512 KB hard limit, warn at 100 KB); (2) max file count per workspace scan (suggested: 500 files, configurable); (3) per-call timeout with `Promise.race` against an `AbortController` signal; (4) explicit note that the server is single-process and that workspace scans are blocking unless parallelized.

---

### [HIGH] Encoding Handling Is Completely Absent

**Description:** VBA files saved from the Visual Basic Editor (VBE) can be in Windows-1252, UTF-8 with BOM, UTF-16 LE, or platform-default ANSI codepages. Node.js `fs.readFile` defaults to returning a Buffer; calling `.toString()` assumes UTF-8. The plan has no mention of encoding detection or normalization.

**Risk:** Non-ASCII characters in variable names, string literals, or comments (common in European and Asian VBA projects) will be misread, causing parse errors or incorrect inspection results that silently corrupt diagnostics. The `vba/inspect-file` and `vba/inspect-workspace` tools are both affected.

**Recommendation:** (1) Default to UTF-8 with BOM stripping; (2) fall back to Windows-1252 on BOM-less files that contain bytes in the 0x80–0x9F range; (3) accept an optional `encoding` parameter on file tools; (4) log a warning when non-UTF-8 encoding is detected. Consider the `chardet` or `iconv-lite` npm packages.

---

### [HIGH] No Logging or Observability Design

**Description:** The plan has no section on logging. There is no mention of a logger, log levels, structured log format, or where logs go. The MCP server uses stdio for the protocol — any naive `console.log` to stdout will corrupt the JSON-RPC stream.

**Risk:** When something goes wrong in production (wrong inspection results, parse errors, tool timeouts), there is no mechanism to diagnose the failure. Logging to stdout corrupts the MCP protocol. Silent failures look like correct behavior to the client.

**Recommendation:** (1) Designate `stderr` as the exclusive log channel (MCP protocol uses `stdout` only); (2) define log levels: ERROR, WARN, INFO, DEBUG; (3) add a structured logger (e.g., `pino` to stderr) from Phase 1; (4) log: server start with version, each tool call with parameters (excluding code content for size), parse errors, inspection errors, and timing. Make log level configurable via environment variable.

---

### [HIGH] .vbalintignore Is Mentioned Once with No Implementation Spec

**Description:** Issue #28 mentions `.vbalintignore` in one sentence: "respect .vbalintignore." There is no specification of: where the file is looked up (workspace root? each directory? user home?), what pattern syntax is supported (gitignore glob syntax? minimatch? exact paths?), how negation patterns (`!`) are handled, whether it applies to `vba/inspect-file` calls or only workspace scans, or what happens when the file is malformed.

**Risk:** Half-implemented ignore support causes user frustration: patterns that appear to work in one form fail in another, files that should be ignored are scanned, or files that should not be ignored are silently excluded. This is a user-facing feature with high visibility.

**Recommendation:** Write a dedicated spec for `.vbalintignore` before Phase 5 implementation: (1) use `minimatch` or `micromatch` for gitignore-compatible glob syntax; (2) look up from workspace root only; (3) support `#` comments and blank lines; (4) define behavior for negation patterns; (5) add test fixtures covering each edge case; (6) document the format in the README.

---

### [HIGH] No User Configuration Management Strategy

**Description:** The plan mentions `src/resources/default-config.json` for "default severity overrides, enabled inspections" but never describes how users provide their own config: file location, file format, validation, schema, merge strategy with defaults, or error behavior for invalid config. There is no config-loading phase in the architecture.

**Risk:** Without a defined config system, the implementation will make ad hoc decisions that are hard to change later. Users have no way to suppress noisy inspections or adjust severities without modifying the source. This is a day-one usability concern for any real VBA codebase.

**Recommendation:** Design config before Phase 2 ships inspections: (1) support a `.vbalintrc.json` file at workspace root; (2) define a JSON schema with severity overrides and enabled/disabled inspection lists; (3) validate config at startup and emit a warning (not a crash) for invalid entries; (4) document the config format in the README with examples.

---

### [HIGH] Concurrency Model Is Undefined

**Description:** The plan does not address concurrent tool calls. MCP clients can issue multiple tool calls simultaneously. The server is a single Node.js process. If `vba/inspect-workspace` is running and a `vba/inspect` call arrives, they share the event loop. ANTLR4 parse tree operations are synchronous and CPU-bound.

**Risk:** A long workspace scan blocks the event loop, causing other tool calls to queue and appear hung to the client. If the MCP client has a response timeout, queued calls will time out even though the server is healthy.

**Recommendation:** (1) Document that the server processes one workspace scan at a time and subsequent calls queue; (2) consider running ANTLR4 parsing in a worker thread pool for CPU-bound work; (3) add a concurrency limit and return an informative error if the limit is exceeded rather than silently queuing; (4) at minimum, address this in the architecture section even if the initial implementation is single-threaded.

---

### [MEDIUM] Docker Image Specification Is Minimal

**Description:** Issue #32 says "multi-stage build, minimal image, stdio transport, health check" in one line. No base image is specified, no image size target, no security scanning step, no non-root user, no read-only filesystem, and no description of what "health check" means for a stdio-transport process.

**Risk:** The default Node.js Docker image is ~300 MB and includes build tools unnecessary at runtime. A root-running container with no security posture is a liability for teams with container security policies. A health check that doesn't actually validate MCP readiness provides false confidence.

**Recommendation:** (1) Use `node:22-alpine` as the runtime base image; (2) run as a non-root user (`USER node`); (3) add a `HEALTHCHECK` that verifies the process is alive (a `CMD node -e "process.exit(0)"` is meaningless — document that stdio health checks are not feasible and omit the instruction or replace with a liveness probe via a side channel); (4) add a `docker scout` or `trivy` scan to CI; (5) set a size target (< 150 MB).

---

### [MEDIUM] Versioning and Breaking Change Strategy Is Absent

**Description:** The plan defines a tool surface (5 tools, specific parameter names, a result schema) but has no versioning strategy. There is no mention of how the tool surface version is communicated to clients, how breaking changes to the schema are handled, or what constitutes a breaking change.

**Risk:** When new fields are added to `InspectionResult` or parameter names change, existing MCP client configurations break silently. When new inspections are added, users who have automated workflows based on `vba/list-inspections` output will see unexpected new entries.

**Recommendation:** (1) Version the server in `package.json` and expose the version via the MCP `serverInfo` field in the `initialize` response; (2) define what is considered breaking (removing a tool, changing a parameter name, changing a result field type) vs. non-breaking (adding optional result fields, adding new inspections); (3) document the stability guarantee in the README; (4) add a CHANGELOG.

---

### [MEDIUM] Startup Time and Readiness Are Not Addressed

**Description:** ANTLR4 grammar-generated parsers have a known startup cost: the parser tables and lexer DFAs are loaded on first use. If this is deferred to the first tool call, the first `vba/parse` or `vba/inspect` call will be significantly slower than subsequent calls. The plan has no warm-up strategy and no mention of server readiness signaling.

**Risk:** The first tool call in a Claude Code session takes several seconds, which may trigger MCP client timeouts or cause the user to believe the tool is broken. This is a first-impression problem for adoption.

**Recommendation:** (1) Eagerly initialize the ANTLR4 parser during server startup, before the `initialize` response is sent; (2) log startup time at INFO level; (3) benchmark cold-start time as part of Phase 1 exit criteria; (4) if startup exceeds 2 seconds, investigate lazy grammar loading or pre-serialized parse tables.

---

### [MEDIUM] Graceful Degradation for Partial Results Is Unspecified

**Description:** If 3 of 60 inspections fail on a given file, the plan does not say whether the caller receives: (a) only the 57 successful results, (b) an error response with no results, or (c) the 57 results plus error entries for the 3 failed inspections. The result schema has no `errors` field.

**Risk:** Callers cannot distinguish "no issues found" from "inspections failed to run." Silent partial results mislead users into believing their code is clean when the inspection run was incomplete.

**Recommendation:** Add an `errors` array to the top-level tool response (not to `InspectionResult`): `{ results: InspectionResult[], errors: { inspection: string, message: string }[] }`. Always return partial results; never silently swallow inspection failures.

---

### [MEDIUM] Upgrade Path for New Inspections Is Not Defined

**Description:** When new inspections are added in a future release, users with `.vbalintrc.json` configurations that explicitly list enabled inspections will not automatically get them. Users without configs will suddenly see new findings in their codebase after upgrading. Neither scenario is addressed.

**Risk:** Surprise findings after upgrade damage trust. Opt-in configs that don't auto-include new inspections silently omit new coverage, which is the opposite of the intended behavior.

**Recommendation:** (1) Define the default behavior as "all inspections enabled unless explicitly disabled"; (2) document the upgrade policy — new inspections are always on by default; (3) add a `vba/list-inspections` field indicating when each inspection was added (semver); (4) consider an `allowlist` vs. `denylist` model in config so users never miss new inspections unintentionally.

---

### [LOW] ANTLR4 Version Pinning Is Loose

**Description:** The dependency table specifies `antlr4 ^4.13` (runtime) with no specific grammar tool version pinned. ANTLR4 grammar tool and runtime versions must match exactly; a mismatch causes silent parse errors or crashes.

**Risk:** If a developer regenerates grammar with a different ANTLR4 tool version than the runtime expects, generated code silently misbehaves. The `^` semver range in package.json could pull a newer runtime that mismatches a pinned grammar tool.

**Recommendation:** (1) Pin both the ANTLR4 runtime and grammar generation tool to the same exact patch version (e.g., `4.13.1`); (2) add a comment in package.json and the build script noting that these must stay in sync; (3) add a CI check that the generated parser files were produced with the correct tool version.

---

### [LOW] Documentation Completeness for End Users Is Deferred

**Description:** README updates are noted as phase exit criteria, but the complete end-user documentation (install, configure, Claude Code config, inspection catalog, Docker usage) is deferred to Phase 6. Users adopting the tool after Phase 2 or Phase 3 will encounter incomplete documentation.

**Risk:** Early adopters (or the developer themselves using the tool mid-build) have no reference for configuration or full tool usage. Missing config documentation causes misconfiguration that produces wrong results.

**Recommendation:** Create a documentation skeleton in Phase 1 that covers installation and basic usage, and update it incrementally as phases complete — rather than batching all documentation to Phase 6. The README update entries in each phase exit criteria are good; make them more specific (which sections must exist, not just "updated").

---

### [LOW] No `vba/inspect` Input Size Limit

**Description:** The `vba/inspect` tool accepts a `code` string parameter with no stated size limit. MCP tool parameters are passed in JSON over stdio. A caller could pass a very large VBA module as an inline string, causing unbounded memory use in the ANTLR4 parser and in JSON deserialization.

**Risk:** A large code string (> 1 MB) in a JSON-RPC message can cause significant memory pressure and slow parsing. In an AI assistant context, this is less likely but not impossible (e.g., pasting an entire module).

**Recommendation:** Add a maximum `code` string length (suggested: 256 KB) enforced at the tool handler level, with a clear error message. Document the limit.

---

## Summary Table

| # | Finding | Severity | Recommendation |
|---|---------|----------|----------------|
| 1 | No error handling strategy defined | CRITICAL | Per-inspection isolation, MCP `isError` responses, parser error results |
| 2 | MCP stdio transport has no malformed input defense | CRITICAL | Verify SDK behavior; add `uncaughtException` handler; document restart behavior |
| 3 | No resource limits — memory, file size, or timeout | CRITICAL | Define and enforce max file size, max file count, per-call timeout |
| 4 | Encoding handling is completely absent | HIGH | Detect encoding; default UTF-8; fall back to Windows-1252; accept `encoding` param |
| 5 | No logging or observability design | HIGH | Log to stderr only; structured logger from Phase 1; log all tool calls and errors |
| 6 | .vbalintignore is mentioned once with no implementation spec | HIGH | Write a dedicated spec; use micromatch; document and test edge cases |
| 7 | No user configuration management strategy | HIGH | Define `.vbalintrc.json`; validate at startup; document schema |
| 8 | Concurrency model is undefined | HIGH | Document single-threaded model; consider worker threads; add concurrency limit |
| 9 | Docker image specification is minimal | MEDIUM | Alpine base; non-root user; trivy scan in CI; drop the health check claim |
| 10 | Versioning and breaking change strategy is absent | MEDIUM | Expose version in `serverInfo`; define breaking vs. non-breaking; add CHANGELOG |
| 11 | Startup time and readiness are not addressed | MEDIUM | Eager parser init at startup; benchmark cold start in Phase 1 exit criteria |
| 12 | Graceful degradation for partial results is unspecified | MEDIUM | Add `errors` array to tool response; always return partial results |
| 13 | Upgrade path for new inspections is not defined | MEDIUM | All inspections on by default; document upgrade policy; use denylist model |
| 14 | ANTLR4 version pinning is loose | LOW | Pin runtime and tool to identical patch version; add CI version check |
| 15 | Documentation completeness is deferred entirely to Phase 6 | LOW | Build doc skeleton in Phase 1; update incrementally |
| 16 | No `vba/inspect` input size limit | LOW | Enforce max code string length (256 KB); document the limit |

---

## Overall Assessment

**Fail**

Three critical gaps — no error handling contract, no stdio input defense, and no resource limits — are sufficient to fail this plan on production readiness grounds. An MCP server that can crash, hang, or exhaust memory on realistic inputs is not production-ready regardless of how complete the inspection catalog is. The high-severity gaps (encoding, logging, config, concurrency, .vbalintignore) compound the risk. None of these are architectural blockers requiring a plan rewrite; they are design decisions that must be made explicit and tested before the server ships to any user. The plan should be revised to include at minimum: a defined error handling strategy, a logging design, resource limit constants, and an encoding policy. These should be addressed in Phase 1 or early Phase 2, not deferred to Phase 5 or Phase 6.
