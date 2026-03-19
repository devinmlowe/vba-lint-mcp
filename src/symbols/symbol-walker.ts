// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// Original: Rubberduck.Parsing/VBA/DeclarationResolving/DeclarationSymbolsListener.cs
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { ParseTreeWalker } from 'antlr4ng';
import { VBAParserListener } from '../parser/generated/grammar/VBAParserListener.js';
import type {
  SubStmtContext,
  FunctionStmtContext,
  PropertyGetStmtContext,
  PropertyLetStmtContext,
  PropertySetStmtContext,
  VariableStmtContext,
  VariableSubStmtContext,
  ConstStmtContext,
  ConstSubStmtContext,
  EnumerationStmtContext,
  EnumerationStmt_ConstantContext,
  UdtDeclarationContext,
  UdtMemberContext,
  EventStmtContext,
  ArgContext,
  VisibilityContext,
  AsTypeClauseContext,
  StatementLabelDefinitionContext,
} from '../parser/generated/grammar/VBAParser.js';
import type { ParseResult } from '../parser/index.js';
import type { Declaration, Accessibility, DeclarationType, SourceLocation } from './declaration.js';

/**
 * Extract the text of an AsTypeClause (e.g., "As Long" → "Long").
 */
function extractTypeName(asTypeClause: AsTypeClauseContext | null | undefined): string | undefined {
  if (!asTypeClause) return undefined;
  const typeCtx = asTypeClause.type();
  if (!typeCtx) return undefined;
  return typeCtx.getText();
}

/**
 * Map a VisibilityContext to our Accessibility type.
 */
function mapVisibility(visCtx: VisibilityContext | null | undefined): Accessibility {
  if (!visCtx) return 'Implicit';
  const text = visCtx.getText().toLowerCase();
  if (text === 'public') return 'Public';
  if (text === 'private') return 'Private';
  if (text === 'friend') return 'Friend';
  if (text === 'global') return 'Global';
  return 'Implicit';
}

/**
 * Build a SourceLocation from an ANTLR parse tree context.
 */
function locationFromContext(ctx: { start: { line: number; column: number } | null; stop: { line: number; column: number; text?: string | null } | null }): SourceLocation {
  const start = ctx.start!;
  const stop = ctx.stop ?? start;
  return {
    startLine: start.line,
    startColumn: start.column,
    endLine: stop.line,
    endColumn: stop.column + (stop.text?.length ?? 0),
  };
}

/**
 * Pass 1 listener: walks the parse tree and collects declarations.
 * Tracks scope hierarchy (Module → Procedure → Block) to set parentScope.
 */
class DeclarationCollectorListener extends VBAParserListener {
  readonly declarations: Declaration[] = [];
  private moduleDecl: Declaration | undefined;
  private currentScope: Declaration | undefined;

  // Track the visibility from the parent VariableStmt/ConstStmt context
  private currentVariableVisibility: Accessibility = 'Implicit';
  private isModuleLevelDeclaration = false;

  constructor(private readonly moduleName: string) {
    super();
  }

  // --- Module declaration (created externally, set as root scope) ---

  setModuleDeclaration(decl: Declaration): void {
    this.moduleDecl = decl;
    this.currentScope = decl;
  }

  // --- Sub ---
  override enterSubStmt = (ctx: SubStmtContext): void => {
    const nameCtx = ctx.subroutineName();
    const name = nameCtx.identifier().getText();
    const vis = mapVisibility(ctx.visibility());
    const decl = this.createDeclaration(
      name, 'Sub', vis === 'Implicit' ? 'Public' : vis,
      locationFromContext(ctx),
    );
    this.declarations.push(decl);
    this.currentScope = decl;
  };

  override exitSubStmt = (_ctx: SubStmtContext): void => {
    this.currentScope = this.moduleDecl;
  };

  // --- Function ---
  override enterFunctionStmt = (ctx: FunctionStmtContext): void => {
    const nameCtx = ctx.functionName();
    const name = nameCtx.identifier().getText();
    const vis = mapVisibility(ctx.visibility());
    const asType = extractTypeName(ctx.asTypeClause());
    const decl = this.createDeclaration(
      name, 'Function', vis === 'Implicit' ? 'Public' : vis,
      locationFromContext(ctx),
      asType,
    );
    decl.isImplicitType = !ctx.asTypeClause();
    this.declarations.push(decl);
    this.currentScope = decl;
  };

  override exitFunctionStmt = (_ctx: FunctionStmtContext): void => {
    this.currentScope = this.moduleDecl;
  };

  // --- Property Get ---
  override enterPropertyGetStmt = (ctx: PropertyGetStmtContext): void => {
    const nameCtx = ctx.functionName();
    const name = nameCtx.identifier().getText();
    const vis = mapVisibility(ctx.visibility());
    const asType = extractTypeName(ctx.asTypeClause());
    const decl = this.createDeclaration(
      name, 'PropertyGet', vis === 'Implicit' ? 'Public' : vis,
      locationFromContext(ctx),
      asType,
    );
    decl.isImplicitType = !ctx.asTypeClause();
    this.declarations.push(decl);
    this.currentScope = decl;
  };

  override exitPropertyGetStmt = (_ctx: PropertyGetStmtContext): void => {
    this.currentScope = this.moduleDecl;
  };

  // --- Property Let ---
  override enterPropertyLetStmt = (ctx: PropertyLetStmtContext): void => {
    const nameCtx = ctx.subroutineName();
    const name = nameCtx.identifier().getText();
    const vis = mapVisibility(ctx.visibility());
    const decl = this.createDeclaration(
      name, 'PropertyLet', vis === 'Implicit' ? 'Public' : vis,
      locationFromContext(ctx),
    );
    this.declarations.push(decl);
    this.currentScope = decl;
  };

  override exitPropertyLetStmt = (_ctx: PropertyLetStmtContext): void => {
    this.currentScope = this.moduleDecl;
  };

  // --- Property Set ---
  override enterPropertySetStmt = (ctx: PropertySetStmtContext): void => {
    const nameCtx = ctx.subroutineName();
    const name = nameCtx.identifier().getText();
    const vis = mapVisibility(ctx.visibility());
    const decl = this.createDeclaration(
      name, 'PropertySet', vis === 'Implicit' ? 'Public' : vis,
      locationFromContext(ctx),
    );
    this.declarations.push(decl);
    this.currentScope = decl;
  };

  override exitPropertySetStmt = (_ctx: PropertySetStmtContext): void => {
    this.currentScope = this.moduleDecl;
  };

  // --- Variables (Dim/Static/Public/Private) ---
  // The VariableStmt contains the visibility; VariableSubStmt has individual names.
  override enterVariableStmt = (ctx: VariableStmtContext): void => {
    const vis = mapVisibility(ctx.visibility());
    // Determine if this is module-level based on parent context
    this.isModuleLevelDeclaration = this.currentScope === this.moduleDecl;

    if (vis !== 'Implicit') {
      this.currentVariableVisibility = vis;
    } else if (ctx.DIM()) {
      // Dim at module level = Private, Dim inside procedure = local
      this.currentVariableVisibility = this.isModuleLevelDeclaration ? 'Private' : 'Implicit';
    } else if (ctx.STATIC()) {
      this.currentVariableVisibility = 'Implicit';
    } else {
      this.currentVariableVisibility = 'Implicit';
    }
  };

  override enterVariableSubStmt = (ctx: VariableSubStmtContext): void => {
    const name = ctx.identifier().getText();
    const asType = extractTypeName(ctx.asTypeClause());
    const isArray = ctx.arrayDim() !== null;
    const decl = this.createDeclaration(
      name, 'Variable', this.currentVariableVisibility,
      locationFromContext(ctx),
      asType,
    );
    decl.isArray = isArray;
    this.declarations.push(decl);
  };

  // --- Constants ---
  override enterConstStmt = (ctx: ConstStmtContext): void => {
    const vis = mapVisibility(ctx.visibility());
    this.isModuleLevelDeclaration = this.currentScope === this.moduleDecl;
    if (vis !== 'Implicit') {
      this.currentVariableVisibility = vis;
    } else {
      this.currentVariableVisibility = this.isModuleLevelDeclaration ? 'Private' : 'Implicit';
    }
  };

  override enterConstSubStmt = (ctx: ConstSubStmtContext): void => {
    const name = ctx.identifier().getText();
    const asType = extractTypeName(ctx.asTypeClause());
    const decl = this.createDeclaration(
      name, 'Constant', this.currentVariableVisibility,
      locationFromContext(ctx),
      asType,
    );
    this.declarations.push(decl);
  };

  // --- Enum ---
  override enterEnumerationStmt = (ctx: EnumerationStmtContext): void => {
    const name = ctx.identifier().getText();
    const vis = mapVisibility(ctx.visibility());
    const decl = this.createDeclaration(
      name, 'Enum', vis === 'Implicit' ? 'Public' : vis,
      locationFromContext(ctx),
    );
    this.declarations.push(decl);
    // Enum members are scoped to the enum
    this.currentScope = decl;
  };

  override exitEnumerationStmt = (_ctx: EnumerationStmtContext): void => {
    this.currentScope = this.moduleDecl;
  };

  override enterEnumerationStmt_Constant = (ctx: EnumerationStmt_ConstantContext): void => {
    const name = ctx.identifier().getText();
    const decl = this.createDeclaration(
      name, 'EnumMember', 'Public', // Enum members are always public
      locationFromContext(ctx),
      'Long', // Enum members are always Long
    );
    decl.isImplicitType = false;
    this.declarations.push(decl);
  };

  // --- Type (UDT) ---
  override enterUdtDeclaration = (ctx: UdtDeclarationContext): void => {
    const name = ctx.untypedIdentifier().getText();
    const vis = mapVisibility(ctx.visibility());
    const decl = this.createDeclaration(
      name, 'Type', vis === 'Implicit' ? 'Public' : vis,
      locationFromContext(ctx),
    );
    this.declarations.push(decl);
    this.currentScope = decl;
  };

  override exitUdtDeclaration = (_ctx: UdtDeclarationContext): void => {
    this.currentScope = this.moduleDecl;
  };

  override enterUdtMember = (ctx: UdtMemberContext): void => {
    // UDT members can be either untypedNameMemberDeclaration or reservedNameMemberDeclaration
    const untypedMember = ctx.untypedNameMemberDeclaration();
    const reservedMember = ctx.reservedNameMemberDeclaration();

    let name: string;
    let asType: string | undefined;

    if (untypedMember) {
      name = untypedMember.untypedIdentifier().getText();
      const optArrayClause = untypedMember.optionalArrayClause();
      asType = optArrayClause ? extractTypeName(optArrayClause.asTypeClause()) : undefined;
    } else if (reservedMember) {
      name = reservedMember.unrestrictedIdentifier().getText();
      asType = extractTypeName(reservedMember.asTypeClause());
    } else {
      return;
    }

    const decl = this.createDeclaration(
      name, 'TypeMember', 'Public', // Type members are always public
      locationFromContext(ctx),
      asType,
    );
    this.declarations.push(decl);
  };

  // --- Event ---
  override enterEventStmt = (ctx: EventStmtContext): void => {
    const name = ctx.identifier().getText();
    const vis = mapVisibility(ctx.visibility());
    const decl = this.createDeclaration(
      name, 'Event', vis === 'Implicit' ? 'Public' : vis,
      locationFromContext(ctx),
    );
    this.declarations.push(decl);
  };

  // --- Parameters ---
  override enterArg = (ctx: ArgContext): void => {
    const name = ctx.unrestrictedIdentifier().getText();
    const asType = extractTypeName(ctx.asTypeClause());
    const isByRef = !ctx.BYVAL(); // ByRef is default in VBA
    const isOptional = ctx.OPTIONAL() !== null;

    const decl = this.createDeclaration(
      name, 'Parameter', 'Implicit',
      locationFromContext(ctx),
      asType,
    );
    decl.isByRef = isByRef;
    decl.isOptional = isOptional;
    decl.isArray = ctx.LPAREN() !== null && ctx.RPAREN() !== null;
    this.declarations.push(decl);
  };

  // --- Line Labels ---
  override enterStatementLabelDefinition = (ctx: StatementLabelDefinitionContext): void => {
    const identLabel = ctx.identifierStatementLabel();
    const numberLabel = ctx.standaloneLineNumberLabel();
    const combinedLabel = ctx.combinedLabels();

    let name: string;
    if (identLabel) {
      name = identLabel.legalLabelIdentifier().getText();
    } else if (numberLabel) {
      name = numberLabel.lineNumberLabel().getText();
    } else if (combinedLabel) {
      name = combinedLabel.getText().replace(/:$/, '');
    } else {
      return;
    }

    const decl = this.createDeclaration(
      name, 'LineLabel', 'Private',
      locationFromContext(ctx),
    );
    this.declarations.push(decl);
  };

  // --- Helpers ---
  private createDeclaration(
    name: string,
    declarationType: DeclarationType,
    accessibility: Accessibility,
    location: SourceLocation,
    asTypeName?: string,
  ): Declaration {
    return {
      name,
      declarationType,
      accessibility,
      parentScope: this.currentScope,
      asTypeName,
      isImplicitType: asTypeName === undefined,
      location,
      references: [],
    };
  }
}

/**
 * Collect all declarations from a parse result (Pass 1).
 *
 * @param parseResult - The parsed VBA module
 * @param moduleName - Name of the module (defaults to 'Module1')
 * @returns Array of all declarations found, including a Module declaration
 */
export function collectDeclarations(parseResult: ParseResult, moduleName = 'Module1'): Declaration[] {
  const listener = new DeclarationCollectorListener(moduleName);

  // Create the module declaration as the root scope
  const moduleDecl: Declaration = {
    name: moduleName,
    declarationType: 'Module',
    accessibility: 'Public',
    isImplicitType: true,
    location: { startLine: 1, startColumn: 0, endLine: 1, endColumn: 0 },
    references: [],
  };
  listener.setModuleDeclaration(moduleDecl);
  listener.declarations.push(moduleDecl);

  const walker = ParseTreeWalker.DEFAULT;
  walker.walk(listener, parseResult.tree);

  return listener.declarations;
}
