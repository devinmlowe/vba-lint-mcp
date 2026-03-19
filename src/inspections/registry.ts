// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { InspectionBase } from './base.js';
import type { InspectionMetadata } from './base.js';

// =================================================================
// Explicit barrel file registration.
// Every inspection class must be imported and added to ALL_INSPECTIONS.
// This provides compile-time safety — missing imports cause build errors.
// =================================================================

// --- Tier A: Empty Blocks ---
import { EmptyIfBlockInspection } from './parse-tree/empty-blocks/empty-if-block.js';
import { EmptyElseBlockInspection } from './parse-tree/empty-blocks/empty-else-block.js';
import { EmptyCaseBlockInspection } from './parse-tree/empty-blocks/empty-case-block.js';
import { EmptyForLoopBlockInspection } from './parse-tree/empty-blocks/empty-for-loop-block.js';
import { EmptyForEachBlockInspection } from './parse-tree/empty-blocks/empty-for-each-block.js';
import { EmptyWhileWendBlockInspection } from './parse-tree/empty-blocks/empty-while-wend-block.js';
import { EmptyDoWhileBlockInspection } from './parse-tree/empty-blocks/empty-do-while-block.js';
import { EmptyMethodInspection } from './parse-tree/empty-blocks/empty-method.js';
import { EmptyModuleInspection } from './parse-tree/empty-blocks/empty-module.js';

// --- Tier A: Obsolete Syntax ---
import { ObsoleteLetStatementInspection } from './parse-tree/obsolete-syntax/obsolete-let-statement.js';
import { ObsoleteCallStatementInspection } from './parse-tree/obsolete-syntax/obsolete-call-statement.js';
import { ObsoleteGlobalInspection } from './parse-tree/obsolete-syntax/obsolete-global.js';
import { ObsoleteWhileWendStatementInspection } from './parse-tree/obsolete-syntax/obsolete-while-wend-statement.js';
import { ObsoleteCommentSyntaxInspection } from './parse-tree/obsolete-syntax/obsolete-comment-syntax.js';
import { ObsoleteTypeHintInspection } from './parse-tree/obsolete-syntax/obsolete-type-hint.js';
import { StopKeywordInspection } from './parse-tree/obsolete-syntax/stop-keyword.js';
import { EndKeywordInspection } from './parse-tree/obsolete-syntax/end-keyword.js';
import { DefTypeStatementInspection } from './parse-tree/obsolete-syntax/def-type-statement.js';

// --- Tier A: Declarations ---
import { OptionExplicitInspection } from './parse-tree/declarations/option-explicit.js';
import { OptionBaseZeroOrOneInspection } from './parse-tree/declarations/option-base-zero-or-one.js';
import { MultipleDeclarationsInspection } from './parse-tree/declarations/multiple-declarations.js';
import { ImplicitByRefModifierInspection } from './parse-tree/declarations/implicit-byref-modifier.js';
import { ImplicitPublicMemberInspection } from './parse-tree/declarations/implicit-public-member.js';
import { ImplicitVariantReturnTypeInspection } from './parse-tree/declarations/implicit-variant-return-type.js';
import { RedundantByRefModifierInspection } from './parse-tree/declarations/redundant-byref-modifier.js';

// --- Tier A: Code Quality ---
import { BooleanAssignedInIfElseInspection } from './parse-tree/code-quality/boolean-assigned-in-if-else.js';
import { SelfAssignedDeclarationInspection } from './parse-tree/code-quality/self-assigned-declaration.js';
import { UnreachableCodeInspection } from './parse-tree/code-quality/unreachable-code.js';
import { LineContinuationBetweenKeywordsInspection } from './parse-tree/code-quality/line-continuation-between-keywords.js';
import { OnLocalErrorInspection } from './parse-tree/code-quality/on-local-error.js';
import { StepNotSpecifiedInspection } from './parse-tree/code-quality/step-not-specified.js';
import { StepOneIsRedundantInspection } from './parse-tree/code-quality/step-one-is-redundant.js';

// --- Tier A: Error Handling ---
import { UnhandledOnErrorResumeNextInspection } from './parse-tree/error-handling/unhandled-on-error-resume-next.js';
import { OnErrorGoToMinusOneInspection } from './parse-tree/error-handling/on-error-goto-minus-one.js';
import { EmptyStringLiteralInspection } from './parse-tree/error-handling/empty-string-literal.js';
import { IsMissingOnInappropriateArgumentInspection } from './parse-tree/error-handling/is-missing-on-inappropriate-argument.js';
import { IsMissingWithNonArgumentParameterInspection } from './parse-tree/error-handling/is-missing-with-non-argument-parameter.js';

// --- Tier A: Excel ---
import { ImplicitActiveSheetReferenceInspection } from './parse-tree/excel/implicit-active-sheet-reference.js';
import { ImplicitActiveWorkbookReferenceInspection } from './parse-tree/excel/implicit-active-workbook-reference.js';
import { SheetAccessedUsingStringInspection } from './parse-tree/excel/sheet-accessed-using-string.js';
import { ApplicationWorksheetFunctionInspection } from './parse-tree/excel/application-worksheet-function.js';
import { ExcelMemberMayReturnNothingInspection } from './parse-tree/excel/excel-member-may-return-nothing.js';
import { ExcelUdfNameIsValidCellReferenceInspection } from './parse-tree/excel/excel-udf-name-is-valid-cell-reference.js';

// --- Tier B: Declaration-based ---
import { VariableNotUsedInspection } from './declaration/variable-not-used.js';
import { ParameterNotUsedInspection } from './declaration/parameter-not-used.js';
import { NonReturningFunctionInspection } from './declaration/non-returning-function.js';

// --- Tier B: Unused Code ---
import { ConstantNotUsedInspection } from './declaration/constant-not-used.js';
import { ProcedureNotUsedInspection } from './declaration/procedure-not-used.js';
import { LineLabelNotUsedInspection } from './declaration/line-label-not-used.js';
import { VariableNotAssignedInspection } from './declaration/variable-not-assigned.js';

// --- Tier B: Naming ---
import { HungarianNotationInspection } from './declaration/hungarian-notation.js';
import { UseMeaningfulNameInspection } from './declaration/use-meaningful-name.js';
import { UnderscoreInPublicClassModuleMemberInspection } from './declaration/underscore-in-public-class-module-member.js';

// --- Tier B: Types ---
import { ObjectVariableNotSetInspection } from './declaration/object-variable-not-set.js';
import { IntegerDataTypeInspection } from './declaration/integer-data-type.js';
import { VariableTypeNotDeclaredInspection } from './declaration/variable-type-not-declared.js';

// --- Tier B: Scope ---
import { ModuleScopeDimKeywordInspection } from './declaration/module-scope-dim-keyword.js';
import { EncapsulatePublicFieldInspection } from './declaration/encapsulate-public-field.js';
import { MoveFieldCloserToUsageInspection } from './declaration/move-field-closer-to-usage.js';

// --- Tier B: Functions ---
import { FunctionReturnValueNotUsedInspection } from './declaration/function-return-value-not-used.js';
import { FunctionReturnValueAlwaysDiscardedInspection } from './declaration/function-return-value-always-discarded.js';
import { ProcedureCanBeWrittenAsFunctionInspection } from './declaration/procedure-can-be-written-as-function.js';

// --- Tier B: Parameters ---
import { ExcessiveParametersInspection } from './declaration/excessive-parameters.js';
import { ParameterCanBeByValInspection } from './declaration/parameter-can-be-byval.js';

// --- Tier B: Usage ---
import { UnassignedVariableUsageInspection } from './declaration/unassigned-variable-usage.js';

/**
 * Master registry of all inspection classes.
 * Order does not matter — the runner handles tiering and filtering.
 */
export const ALL_INSPECTIONS: Array<new () => InspectionBase> = [
  // Tier A: Empty Blocks
  EmptyIfBlockInspection,
  EmptyElseBlockInspection,
  EmptyCaseBlockInspection,
  EmptyForLoopBlockInspection,
  EmptyForEachBlockInspection,
  EmptyWhileWendBlockInspection,
  EmptyDoWhileBlockInspection,
  EmptyMethodInspection,
  EmptyModuleInspection,

  // Tier A: Obsolete Syntax
  ObsoleteLetStatementInspection,
  ObsoleteCallStatementInspection,
  ObsoleteGlobalInspection,
  ObsoleteWhileWendStatementInspection,
  ObsoleteCommentSyntaxInspection,
  ObsoleteTypeHintInspection,
  StopKeywordInspection,
  EndKeywordInspection,
  DefTypeStatementInspection,

  // Tier A: Declarations
  OptionExplicitInspection,
  OptionBaseZeroOrOneInspection,
  MultipleDeclarationsInspection,
  ImplicitByRefModifierInspection,
  ImplicitPublicMemberInspection,
  ImplicitVariantReturnTypeInspection,
  RedundantByRefModifierInspection,

  // Tier A: Code Quality
  BooleanAssignedInIfElseInspection,
  SelfAssignedDeclarationInspection,
  UnreachableCodeInspection,
  LineContinuationBetweenKeywordsInspection,
  OnLocalErrorInspection,
  StepNotSpecifiedInspection,
  StepOneIsRedundantInspection,

  // Tier A: Error Handling
  UnhandledOnErrorResumeNextInspection,
  OnErrorGoToMinusOneInspection,
  EmptyStringLiteralInspection,
  IsMissingOnInappropriateArgumentInspection,
  IsMissingWithNonArgumentParameterInspection,

  // Tier A: Excel
  ImplicitActiveSheetReferenceInspection,
  ImplicitActiveWorkbookReferenceInspection,
  SheetAccessedUsingStringInspection,
  ApplicationWorksheetFunctionInspection,
  ExcelMemberMayReturnNothingInspection,
  ExcelUdfNameIsValidCellReferenceInspection,

  // Tier B: Declaration-based
  VariableNotUsedInspection,
  ParameterNotUsedInspection,
  NonReturningFunctionInspection,

  // Tier B: Unused Code
  ConstantNotUsedInspection,
  ProcedureNotUsedInspection,
  LineLabelNotUsedInspection,
  VariableNotAssignedInspection,

  // Tier B: Naming
  HungarianNotationInspection,
  UseMeaningfulNameInspection,
  UnderscoreInPublicClassModuleMemberInspection,

  // Tier B: Types
  ObjectVariableNotSetInspection,
  IntegerDataTypeInspection,
  VariableTypeNotDeclaredInspection,

  // Tier B: Scope
  ModuleScopeDimKeywordInspection,
  EncapsulatePublicFieldInspection,
  MoveFieldCloserToUsageInspection,

  // Tier B: Functions
  FunctionReturnValueNotUsedInspection,
  FunctionReturnValueAlwaysDiscardedInspection,
  ProcedureCanBeWrittenAsFunctionInspection,

  // Tier B: Parameters
  ExcessiveParametersInspection,
  ParameterCanBeByValInspection,

  // Tier B: Usage
  UnassignedVariableUsageInspection,
];

/**
 * Create instances of all registered inspections.
 */
export function createAllInspections(): InspectionBase[] {
  return ALL_INSPECTIONS.map(Cls => new Cls());
}

/**
 * Get metadata for all registered inspections without instantiating.
 */
export function getAllInspectionMetadata(): InspectionMetadata[] {
  return ALL_INSPECTIONS.map(Cls => (Cls as unknown as typeof InspectionBase).meta);
}

/**
 * Validate the registry at startup.
 * Returns errors for any issues found.
 */
export function validateRegistry(): string[] {
  const errors: string[] = [];
  const ids = new Set<string>();

  for (const Cls of ALL_INSPECTIONS) {
    const meta = (Cls as unknown as typeof InspectionBase).meta;

    if (!meta) {
      errors.push(`Inspection class ${Cls.name} is missing static 'meta' property`);
      continue;
    }

    if (!meta.id) {
      errors.push(`Inspection class ${Cls.name} has empty 'id' in meta`);
      continue;
    }

    if (ids.has(meta.id)) {
      errors.push(`Duplicate inspection ID: ${meta.id}`);
    }
    ids.add(meta.id);

    if (!meta.tier || !['A', 'B'].includes(meta.tier)) {
      errors.push(`Inspection ${meta.id} has invalid tier: ${meta.tier}`);
    }

    if (!meta.category) {
      errors.push(`Inspection ${meta.id} has no category`);
    }

    if (!meta.defaultSeverity) {
      errors.push(`Inspection ${meta.id} has no defaultSeverity`);
    }
  }

  return errors;
}
