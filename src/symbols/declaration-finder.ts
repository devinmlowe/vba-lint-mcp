// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.Parsing/VBA/DeclarationCaching/DeclarationFinder.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import type { Declaration, DeclarationType, IdentifierReference } from './declaration.js';
import { isReturningDeclarationType, isProcedureDeclarationType } from './declaration.js';

/**
 * Query interface for the symbol table.
 * Provides efficient lookups and analysis queries over collected declarations.
 */
export class DeclarationFinder {
  private readonly byNameIndex: Map<string, Declaration[]>;

  constructor(private readonly declarations: Declaration[]) {
    // Build case-insensitive name index
    this.byNameIndex = new Map();
    for (const decl of declarations) {
      const key = decl.name.toLowerCase();
      let list = this.byNameIndex.get(key);
      if (!list) {
        list = [];
        this.byNameIndex.set(key, list);
      }
      list.push(decl);
    }
  }

  /**
   * Get all declarations.
   */
  getAll(): Declaration[] {
    return this.declarations;
  }

  /**
   * Find declarations by name (case-insensitive).
   */
  findByName(name: string): Declaration[] {
    return this.byNameIndex.get(name.toLowerCase()) ?? [];
  }

  /**
   * Find declarations by type.
   */
  findByType(type: DeclarationType): Declaration[] {
    return this.declarations.filter(d => d.declarationType === type);
  }

  /**
   * Find declarations with zero non-assignment references.
   * A declaration is "unused" if no one reads from it.
   * Excludes module declarations and certain types.
   */
  findUnused(): Declaration[] {
    return this.declarations.filter(d => {
      // Skip module declarations, line labels, enum members, type members
      if (d.declarationType === 'Module') return false;
      if (d.declarationType === 'LineLabel') return false;
      if (d.declarationType === 'EnumMember') return false;
      if (d.declarationType === 'TypeMember') return false;
      // Skip public procedures (they may be called from other modules)
      if (isProcedureDeclarationType(d.declarationType) && d.accessibility !== 'Private') return false;
      // Skip events
      if (d.declarationType === 'Event') return false;
      // Skip enums and types (they may be used as types)
      if (d.declarationType === 'Enum') return false;
      if (d.declarationType === 'Type') return false;

      // Has no references at all?
      return d.references.length === 0;
    });
  }

  /**
   * Find variables that are never assigned a value.
   */
  findVariablesNotAssigned(): Declaration[] {
    return this.declarations.filter(d => {
      if (d.declarationType !== 'Variable') return false;
      return !d.references.some(r => r.isAssignment);
    });
  }

  /**
   * Find functions/property-gets where the function name is never assigned
   * (meaning the function never sets its return value).
   */
  findNonReturningFunctions(): Declaration[] {
    return this.declarations.filter(d => {
      if (!isReturningDeclarationType(d.declarationType)) return false;
      // Check if any reference to this declaration is an assignment
      // In VBA, you assign the return value by assigning to the function name
      return !d.references.some(r => r.isAssignment);
    });
  }

  /**
   * Find all references to a specific declaration.
   */
  findReferencesTo(declaration: Declaration): IdentifierReference[] {
    return declaration.references;
  }

  /**
   * Find unused variables specifically (Dim'd variable with zero references).
   */
  findUnusedVariables(): Declaration[] {
    return this.declarations.filter(d => {
      if (d.declarationType !== 'Variable') return false;
      return d.references.length === 0;
    });
  }

  /**
   * Find unused parameters (parameters with zero references).
   */
  findUnusedParameters(): Declaration[] {
    return this.declarations.filter(d => {
      if (d.declarationType !== 'Parameter') return false;
      return d.references.length === 0;
    });
  }
}
