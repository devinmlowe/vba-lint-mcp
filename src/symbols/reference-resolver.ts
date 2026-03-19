// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.Parsing/VBA/ReferenceResolving/IdentificationListener.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { ParseTreeWalker } from 'antlr4ng';
import { VBAParserListener } from '../parser/generated/grammar/VBAParserListener.js';
import type {
  SubStmtContext,
  FunctionStmtContext,
  PropertyGetStmtContext,
  PropertyLetStmtContext,
  PropertySetStmtContext,
} from '../parser/generated/grammar/VBAParser.js';
import {
  SimpleNameExprContext,
  LetStmtContext,
  SetStmtContext,
} from '../parser/generated/grammar/VBAParser.js';
import type { ParseResult } from '../parser/index.js';
import type { Declaration, IdentifierReference, SourceLocation } from './declaration.js';
import { isProcedureDeclarationType } from './declaration.js';

/**
 * Build a scope lookup map: for each scope declaration, collect
 * the declarations that belong to it.
 */
function buildScopeMap(declarations: Declaration[]): Map<Declaration | undefined, Declaration[]> {
  const map = new Map<Declaration | undefined, Declaration[]>();
  for (const decl of declarations) {
    const parent = decl.parentScope;
    let list = map.get(parent);
    if (!list) {
      list = [];
      map.set(parent, list);
    }
    list.push(decl);
  }
  return map;
}

/**
 * Find a declaration by name in the given scope, then walk up to parent scopes.
 * VBA name resolution is case-insensitive.
 */
function resolveSimpleName(
  name: string,
  currentScope: Declaration | undefined,
  scopeMap: Map<Declaration | undefined, Declaration[]>,
  allDeclarations: Declaration[],
): Declaration | undefined {
  const lowerName = name.toLowerCase();
  let scope: Declaration | undefined = currentScope;

  while (scope) {
    const children = scopeMap.get(scope);
    if (children) {
      const found = children.find(d => d.name.toLowerCase() === lowerName);
      if (found) return found;
    }
    scope = scope.parentScope;
  }

  // Also check top-level (declarations with no parent scope, e.g., the module itself)
  const topLevel = allDeclarations.filter(d => !d.parentScope);
  return topLevel.find(d => d.name.toLowerCase() === lowerName);
}

/**
 * Pass 2 listener: walks the parse tree to find identifier usages
 * and resolve them to declarations.
 */
class ReferenceResolverListener extends VBAParserListener {
  private currentProcedure: Declaration | undefined;
  private moduleDecl: Declaration | undefined;
  private readonly scopeMap: Map<Declaration | undefined, Declaration[]>;

  // Track assignment context: when entering a LetStmt or SetStmt,
  // the LHS lExpression is an assignment target.
  private inAssignmentLHS = false;

  constructor(
    private readonly declarations: Declaration[],
  ) {
    super();
    this.scopeMap = buildScopeMap(declarations);
    this.moduleDecl = declarations.find(d => d.declarationType === 'Module');
  }

  // --- Track current scope for resolution ---

  override enterSubStmt = (ctx: SubStmtContext): void => {
    const name = ctx.subroutineName().identifier().getText();
    this.currentProcedure = this.findDeclaration(name, 'Sub');
  };

  override exitSubStmt = (_ctx: SubStmtContext): void => {
    this.currentProcedure = undefined;
  };

  override enterFunctionStmt = (ctx: FunctionStmtContext): void => {
    const name = ctx.functionName().identifier().getText();
    this.currentProcedure = this.findDeclaration(name, 'Function');
  };

  override exitFunctionStmt = (_ctx: FunctionStmtContext): void => {
    this.currentProcedure = undefined;
  };

  override enterPropertyGetStmt = (ctx: PropertyGetStmtContext): void => {
    const name = ctx.functionName().identifier().getText();
    this.currentProcedure = this.findDeclaration(name, 'PropertyGet');
  };

  override exitPropertyGetStmt = (_ctx: PropertyGetStmtContext): void => {
    this.currentProcedure = undefined;
  };

  override enterPropertyLetStmt = (ctx: PropertyLetStmtContext): void => {
    const name = ctx.subroutineName().identifier().getText();
    this.currentProcedure = this.findDeclaration(name, 'PropertyLet');
  };

  override exitPropertyLetStmt = (_ctx: PropertyLetStmtContext): void => {
    this.currentProcedure = undefined;
  };

  override enterPropertySetStmt = (ctx: PropertySetStmtContext): void => {
    const name = ctx.subroutineName().identifier().getText();
    this.currentProcedure = this.findDeclaration(name, 'PropertySet');
  };

  override exitPropertySetStmt = (_ctx: PropertySetStmtContext): void => {
    this.currentProcedure = undefined;
  };

  // --- Assignment detection ---

  override enterLetStmt = (ctx: LetStmtContext): void => {
    // The lExpression on the LHS is an assignment target
    const lExpr = ctx.lExpression();
    if (lExpr instanceof SimpleNameExprContext) {
      const name = lExpr.identifier().getText();
      this.addReference(name, locationFromNode(lExpr), true);
    }
    // Note: For MemberAccessExpr on LHS, we skip for now (cross-module)
  };

  override enterSetStmt = (ctx: SetStmtContext): void => {
    const lExpr = ctx.lExpression();
    if (lExpr instanceof SimpleNameExprContext) {
      const name = lExpr.identifier().getText();
      this.addReference(name, locationFromNode(lExpr), true);
    }
  };

  // --- Simple name expressions (reads) ---

  override enterSimpleNameExpr = (ctx: SimpleNameExprContext): void => {
    // Skip if this is the LHS of a let/set statement (handled above as assignment)
    const parent = ctx.parent;
    if (parent && (isLetStmtLHS(parent, ctx) || isSetStmtLHS(parent, ctx))) {
      return;
    }

    const name = ctx.identifier().getText();
    this.addReference(name, locationFromNode(ctx), false);
  };

  // --- Helpers ---

  private findDeclaration(name: string, type: string): Declaration | undefined {
    const lowerName = name.toLowerCase();
    return this.declarations.find(
      d => d.name.toLowerCase() === lowerName && d.declarationType === type,
    );
  }

  private addReference(name: string, location: SourceLocation, isAssignment: boolean): void {
    const scope = this.currentProcedure ?? this.moduleDecl;
    const resolved = resolveSimpleName(name, scope, this.scopeMap, this.declarations);

    const ref: IdentifierReference = {
      name,
      location,
      declaration: resolved,
      isAssignment,
    };

    if (resolved) {
      resolved.references.push(ref);
    }
  }
}

function locationFromNode(ctx: { start: { line: number; column: number } | null; stop: { line: number; column: number; text?: string | null } | null }): SourceLocation {
  const start = ctx.start!;
  const stop = ctx.stop ?? start;
  return {
    startLine: start.line,
    startColumn: start.column,
    endLine: stop.line,
    endColumn: stop.column + (stop.text?.length ?? 0),
  };
}

function isLetStmtLHS(parent: any, ctx: SimpleNameExprContext): boolean {
  return parent instanceof LetStmtContext && parent.lExpression() === ctx;
}

function isSetStmtLHS(parent: any, ctx: SimpleNameExprContext): boolean {
  return parent instanceof SetStmtContext && parent.lExpression() === ctx;
}

/**
 * Resolve references in a set of declarations (Pass 2).
 * Mutates the declarations array by adding references to each declaration.
 *
 * @param parseResult - The parsed VBA module
 * @param declarations - Declarations collected from Pass 1
 */
export function resolveReferences(parseResult: ParseResult, declarations: Declaration[]): void {
  const listener = new ReferenceResolverListener(declarations);
  const walker = ParseTreeWalker.DEFAULT;
  walker.walk(listener, parseResult.tree);
}
