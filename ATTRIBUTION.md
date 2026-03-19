# Attribution

## Rubberduck VBA

This project is a derivative work of the [Rubberduck VBA](https://github.com/rubberduck-vba/Rubberduck) project.

**Copyright (C) Rubberduck Contributors**

Rubberduck is licensed under the GNU General Public License v3.0. In compliance with GPL-3.0, this project is also licensed under GPL-3.0.

### Derived Works

| File in this project | Original source in Rubberduck | Nature |
|---|---|---|
| `grammar/VBALexer.g4` | `Rubberduck.Parsing/Grammar/VBALexer.g4` | Direct copy |
| `grammar/VBAParser.g4` | `Rubberduck.Parsing/Grammar/VBAParser.g4` | Direct copy |
| `src/resources/en/*.json` | `Rubberduck.Resources/Inspections/*.resx` | Extracted and reformatted |
| `src/inspections/**/*.ts` | `Rubberduck.CodeAnalysis/Inspections/Concrete/*.cs` | Logic translated C# → TypeScript |
| `src/symbols/declaration.ts` | `Rubberduck.Parsing/Model/Symbols/Declaration.cs` | Architecture ported |

### Source Pinning

Grammar and inspection logic derived from Rubberduck v2 repository:
- **Repository:** https://github.com/rubberduck-vba/Rubberduck
- **Commit:** (to be pinned during grammar integration — Issue #2)
- **License:** GPL-3.0

### Rubberduck Contributors

We gratefully acknowledge the work of all Rubberduck contributors. A full list of contributors is available at:
https://github.com/rubberduck-vba/Rubberduck/graphs/contributors

### Legal Note

Translated inspection logic (C# → TypeScript) constitutes a derivative work under copyright law. Every file derived from Rubberduck includes a per-file copyright header:

```
// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: [path in Rubberduck repo] (commit: [hash])
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE
```

Consumers of this project (including via Docker or npm) must comply with GPL-3.0 terms.
