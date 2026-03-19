// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.Parsing/VBA/RubberduckParserState.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import type { ParseResult } from '../parser/index.js';
import type { Declaration } from './declaration.js';
import { DeclarationFinder } from './declaration-finder.js';
import { collectDeclarations } from './symbol-walker.js';
import { resolveReferences } from './reference-resolver.js';

/**
 * Multi-module container for cross-module symbol resolution.
 * For single-file mode, create a workspace with just one module.
 */
export class WorkspaceContext {
  private readonly modules = new Map<string, {
    parseResult: ParseResult;
    declarations: Declaration[];
  }>();

  /**
   * Add a module to the workspace.
   * Runs Pass 1 (declaration collection) immediately.
   * Call getDeclarationFinder() after adding all modules to trigger Pass 2.
   */
  addModule(name: string, parseResult: ParseResult): void {
    const declarations = collectDeclarations(parseResult, name);
    this.modules.set(name, { parseResult, declarations });
  }

  /**
   * Get a cross-module declaration finder.
   * Triggers Pass 2 (reference resolution) for all modules.
   */
  getDeclarationFinder(): DeclarationFinder {
    // Merge all declarations from all modules
    const allDeclarations: Declaration[] = [];
    for (const { declarations } of this.modules.values()) {
      allDeclarations.push(...declarations);
    }

    // Run Pass 2 for each module with the merged declaration set
    for (const { parseResult } of this.modules.values()) {
      resolveReferences(parseResult, allDeclarations);
    }

    return new DeclarationFinder(allDeclarations);
  }

  /**
   * Get the number of modules in the workspace.
   */
  get moduleCount(): number {
    return this.modules.size;
  }
}

/**
 * Convenience: build a DeclarationFinder for a single module.
 */
export function buildSingleModuleFinder(parseResult: ParseResult, moduleName = 'Module1'): DeclarationFinder {
  const declarations = collectDeclarations(parseResult, moduleName);
  resolveReferences(parseResult, declarations);
  return new DeclarationFinder(declarations);
}
