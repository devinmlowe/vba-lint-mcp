// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.Parsing/Model/Symbols/Declaration.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

/**
 * Types of VBA declarations, modeled after Rubberduck's DeclarationType enum.
 */
export type DeclarationType =
  | 'Module'
  | 'Sub'
  | 'Function'
  | 'PropertyGet'
  | 'PropertyLet'
  | 'PropertySet'
  | 'Variable'
  | 'Constant'
  | 'Parameter'
  | 'Enum'
  | 'EnumMember'
  | 'Type'
  | 'TypeMember'
  | 'LineLabel'
  | 'Event';

/**
 * Accessibility levels for declarations, modeled after Rubberduck's Accessibility enum.
 */
export type Accessibility = 'Public' | 'Private' | 'Friend' | 'Global' | 'Implicit';

/**
 * Source location of a declaration or reference.
 */
export interface SourceLocation {
  startLine: number;
  startColumn: number;
  endLine: number;
  endColumn: number;
}

/**
 * A reference to an identifier in source code.
 * Links a usage site back to its resolved declaration.
 */
export interface IdentifierReference {
  /** The identifier name as written in source */
  name: string;
  /** Location of the reference in source code */
  location: SourceLocation;
  /** The resolved declaration (undefined if external/unresolvable) */
  declaration?: Declaration;
  /** True if this reference is an assignment target (lValue) */
  isAssignment: boolean;
}

/**
 * A VBA declaration extracted from source code.
 * Modeled after Rubberduck's Declaration class.
 */
export interface Declaration {
  /** The declared name (e.g., variable name, procedure name) */
  name: string;
  /** What kind of declaration this is */
  declarationType: DeclarationType;
  /** Access modifier */
  accessibility: Accessibility;
  /** The containing scope (module for top-level, procedure for locals) */
  parentScope?: Declaration;
  /** Type name from As clause (e.g., "Long", "String", "Variant") */
  asTypeName?: string;
  /** True if no explicit As Type was specified (implicit Variant) */
  isImplicitType: boolean;
  /** Location of the declaration in source code */
  location: SourceLocation;
  /** All references to this declaration */
  references: IdentifierReference[];

  // --- Parameter-specific ---
  /** True if parameter is ByRef (explicit or implicit) */
  isByRef?: boolean;
  /** True if parameter is Optional */
  isOptional?: boolean;

  // --- Function/variable-specific ---
  /** True if declared as an array */
  isArray?: boolean;
}

/**
 * Helper to determine if a declaration type represents a procedure.
 */
export function isProcedureDeclarationType(type: DeclarationType): boolean {
  return type === 'Sub' || type === 'Function'
    || type === 'PropertyGet' || type === 'PropertyLet' || type === 'PropertySet';
}

/**
 * Helper to determine if a declaration type can have a return value
 * (i.e., the name can be assigned to return a value).
 */
export function isReturningDeclarationType(type: DeclarationType): boolean {
  return type === 'Function' || type === 'PropertyGet';
}
