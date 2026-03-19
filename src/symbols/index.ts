// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

export type {
  Declaration,
  DeclarationType,
  Accessibility,
  SourceLocation,
  IdentifierReference,
} from './declaration.js';
export { isProcedureDeclarationType, isReturningDeclarationType } from './declaration.js';
export { collectDeclarations } from './symbol-walker.js';
export { resolveReferences } from './reference-resolver.js';
export { DeclarationFinder } from './declaration-finder.js';
export { WorkspaceContext, buildSingleModuleFinder } from './workspace.js';
