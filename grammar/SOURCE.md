# Grammar Source

- **Repository:** https://github.com/rubberduck-vba/Rubberduck
- **Branch:** next
- **Commit:** fae50adab188126a5e7d2a1cefc3328cc18af482
- **Date extracted:** 2026-03-19
- **Files:**
  - VBALexer.g4
  - VBAParser.g4
- **Modifications:**
  - VBALexer.g4: Removed `superClass = VBABaseLexer` and `contextSuperClass = VBABaseParser` options (Rubberduck-specific base classes not needed for standalone use)
  - VBAParser.g4: Removed `superClass = VBABaseParser` and `contextSuperClass = VBABaseParserRuleContext` options (same reason)
  - VBAParser.g4: Changed `unterminatedBlock` rule to use `individualNonEOFEndOfStatement+` instead of `endOfStatement` to fix ANTLR4 error 186 (closure matching EOF in `endOfStatement`)

## Upgrade Procedure

1. Check for grammar changes: `git diff fae50adab188126a5e7d2a1cefc3328cc18af482..[new-hash] -- Rubberduck.Parsing/Grammar/`
2. Copy updated grammar files to `grammar/`
3. Regenerate parser: `npm run generate-parser`
4. Run grammar fidelity tests: `npm test -- --grep grammar`
5. Update this file with new commit hash
