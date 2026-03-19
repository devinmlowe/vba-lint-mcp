# Contributing to vba-lint-mcp

Thank you for your interest in contributing. This project is licensed under GPL-3.0 (inherited from Rubberduck VBA), and all contributions must comply with those terms.

## Prerequisites

- Node.js >= 20
- Java 17+ (JRE) for ANTLR4 parser generation
- Git

## Development Setup

```bash
# Clone the repository
git clone git@github.com:devinmlowe/vba-lint-mcp.git
cd vba-lint-mcp

# Install dependencies
npm install

# Generate the ANTLR4 parser (requires Java)
npm run generate-parser

# Build
npm run build

# Run tests
npm test
```

## Running the Server Locally

```bash
# Development mode (tsx, no build step)
npm run dev

# Production mode
npm run build
node dist/server.js
```

The server communicates via stdio (stdin/stdout JSON-RPC). It is designed to be launched by an MCP client such as Claude Code.

## Adding a New Inspection

This is the most common type of contribution. Follow these steps:

### 1. Choose the Tier

- **Tier A** (parse-tree): Your inspection only needs the ANTLR4 parse tree. Place it under `src/inspections/parse-tree/<category>/`.
- **Tier B** (symbol-aware): Your inspection needs declaration/reference information. Place it under `src/inspections/declaration/`.

### 2. Create the Inspection File

Create a new file in the appropriate directory. Use an existing inspection as a template.

**Example: Tier A inspection** (`src/inspections/parse-tree/code-quality/my-inspection.ts`):

```typescript
// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { ParseTreeInspection, type InspectionContext, type InspectionMetadata } from '../../base.js';
import type { InspectionResult } from '../../types.js';

export class MyInspection extends ParseTreeInspection {
  static readonly meta: InspectionMetadata = {
    id: 'MyInspection',
    tier: 'A',
    category: 'CodeQuality',
    defaultSeverity: 'warning',
    name: 'My inspection name',
    description: 'Human-readable description of what this detects.',
    quickFixDescription: 'Description of the suggested fix',
  };

  inspect(context: InspectionContext): InspectionResult[] {
    const results: InspectionResult[] = [];
    // Walk the parse tree and detect the pattern
    // Use this.createResult() to build results
    return results;
  }
}
```

### 3. Register the Inspection

Add your import and class to `src/inspections/registry.ts`:

```typescript
import { MyInspection } from './parse-tree/code-quality/my-inspection.js';

export const ALL_INSPECTIONS: Array<new () => InspectionBase> = [
  // ... existing inspections ...
  MyInspection,
];
```

### 4. Write Tests

Create a test file alongside your inspection (co-located tests):

**`src/inspections/parse-tree/code-quality/__tests__/my-inspection.test.ts`**:

```typescript
import { describe, it, expect } from 'vitest';
import { MyInspection } from '../my-inspection.js';
import { parseCode } from '../../../../parser/index.js';

describe('MyInspection', () => {
  const inspection = new MyInspection();

  it('detects the problematic pattern', () => {
    const code = `
Sub Example()
    ' VBA code that triggers the inspection
End Sub`;
    const parseResult = parseCode(code);
    const results = inspection.inspect({ parseResult });

    expect(results).toHaveLength(1);
    expect(results[0].inspection).toBe('MyInspection');
    expect(results[0].severity).toBe('warning');
    expect(results[0].location.startLine).toBe(3);
  });

  it('does not flag correct code (false-positive test)', () => {
    const code = `
Sub Example()
    ' VBA code that should NOT trigger
End Sub`;
    const parseResult = parseCode(code);
    const results = inspection.inspect({ parseResult });

    expect(results).toHaveLength(0);
  });
});
```

Every inspection test must assert:
- Correct result count
- Correct inspection ID
- Correct severity
- Correct start line
- At least one false-positive test (similar construct that should NOT trigger)

### 5. Test Fixtures (Optional)

For complex inspections, create `.bas` or `.cls` fixture files:

```
src/inspections/parse-tree/code-quality/__fixtures__/
    my-inspection-positive.bas    # Code that should trigger
    my-inspection-negative.bas    # Code that should not trigger
```

### 6. Verify

```bash
# Run your specific test
npx vitest run src/inspections/parse-tree/code-quality/__tests__/my-inspection.test.ts

# Run all tests (ensure no regressions)
npm test

# Run lint
npm run lint
```

## Running Tests

```bash
# All tests
npm test

# Watch mode
npm run test:watch

# Specific test file
npx vitest run path/to/test.ts

# With coverage
npx vitest run --coverage
```

## Commit Conventions

- Use [Conventional Commits](https://www.conventionalcommits.org/) style
- Prefix with scope: `feat(inspections):`, `fix(parser):`, `test:`, `docs:`, etc.
- Reference the GitHub issue number when applicable: `(#42)`

Examples:
```
feat(inspections): add MyInspection for detecting X (#99)
fix(parser): handle edge case with line continuations in strings
test: add false-positive tests for EmptyIfBlock
docs: update inspection catalog in README
```

## GPL-3.0 Requirements

This project is licensed under GPL-3.0, inherited from Rubberduck VBA. By contributing:

1. **Copyright headers**: Every file derived from Rubberduck must include the attribution header:
   ```typescript
   // Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
   // License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE
   ```

2. **License compatibility**: All contributions must be compatible with GPL-3.0. Do not introduce dependencies with incompatible licenses.

3. **Source availability**: The full source code must remain available to all recipients of the software (including via Docker images).

4. **Attribution**: If your inspection is ported from Rubberduck, include the original C# file path in the header:
   ```typescript
   // Original: Rubberduck.CodeAnalysis/Inspections/Concrete/MyInspection.cs
   ```

## Project Structure

```
src/
  server.ts                 # MCP server entry point
  inspections/
    base.ts                 # Inspection base classes
    registry.ts             # Master registration (all inspections)
    runner.ts               # Tiered execution engine
    types.ts                # Result types, severity, category
    parse-tree/             # Tier A inspections (by category)
    declaration/            # Tier B inspections
  parser/
    index.ts                # Parser facade
    generated/              # ANTLR4 generated code (gitignored)
  symbols/                  # Symbol resolution (Tier B)
  annotations/              # @Ignore annotation parsing
```

## Questions?

Open a GitHub issue for questions about inspection design, grammar modifications, or architecture decisions.
