// Code generated from vba.g4 by ANTLR 4.13.2. DO NOT EDIT.

package antlr // vba
import "github.com/antlr4-go/antlr/v4"

// vbaListener is a complete listener for a parse tree produced by vbaParser.
type vbaListener interface {
	antlr.ParseTreeListener

	// EnterStartRule is called when entering the startRule production.
	EnterStartRule(c *StartRuleContext)

	// EnterModule is called when entering the module production.
	EnterModule(c *ModuleContext)

	// EnterClassFileHeader is called when entering the classFileHeader production.
	EnterClassFileHeader(c *ClassFileHeaderContext)

	// EnterClassVersionIdentification is called when entering the classVersionIdentification production.
	EnterClassVersionIdentification(c *ClassVersionIdentificationContext)

	// EnterClassBeginBlock is called when entering the classBeginBlock production.
	EnterClassBeginBlock(c *ClassBeginBlockContext)

	// EnterBeginBlockConfigElement is called when entering the beginBlockConfigElement production.
	EnterBeginBlockConfigElement(c *BeginBlockConfigElementContext)

	// EnterFormFileHeader is called when entering the formFileHeader production.
	EnterFormFileHeader(c *FormFileHeaderContext)

	// EnterFormVersionIdentification is called when entering the formVersionIdentification production.
	EnterFormVersionIdentification(c *FormVersionIdentificationContext)

	// EnterFormObjectAssign is called when entering the formObjectAssign production.
	EnterFormObjectAssign(c *FormObjectAssignContext)

	// EnterFormBeginBlock is called when entering the formBeginBlock production.
	EnterFormBeginBlock(c *FormBeginBlockContext)

	// EnterBeginPropertyBlock is called when entering the beginPropertyBlock production.
	EnterBeginPropertyBlock(c *BeginPropertyBlockContext)

	// EnterProceduralModule is called when entering the proceduralModule production.
	EnterProceduralModule(c *ProceduralModuleContext)

	// EnterClassModule is called when entering the classModule production.
	EnterClassModule(c *ClassModuleContext)

	// EnterProceduralModuleHeader is called when entering the proceduralModuleHeader production.
	EnterProceduralModuleHeader(c *ProceduralModuleHeaderContext)

	// EnterProceduralModuleAttr is called when entering the proceduralModuleAttr production.
	EnterProceduralModuleAttr(c *ProceduralModuleAttrContext)

	// EnterIgnoredProceduralAttr is called when entering the ignoredProceduralAttr production.
	EnterIgnoredProceduralAttr(c *IgnoredProceduralAttrContext)

	// EnterClassModuleHeader is called when entering the classModuleHeader production.
	EnterClassModuleHeader(c *ClassModuleHeaderContext)

	// EnterClassAttr is called when entering the classAttr production.
	EnterClassAttr(c *ClassAttrContext)

	// EnterIgnoredClassAttr is called when entering the ignoredClassAttr production.
	EnterIgnoredClassAttr(c *IgnoredClassAttrContext)

	// EnterIgnoredAttr is called when entering the ignoredAttr production.
	EnterIgnoredAttr(c *IgnoredAttrContext)

	// EnterNameAttr is called when entering the nameAttr production.
	EnterNameAttr(c *NameAttrContext)

	// EnterProceduralModuleBody is called when entering the proceduralModuleBody production.
	EnterProceduralModuleBody(c *ProceduralModuleBodyContext)

	// EnterClassModuleBody is called when entering the classModuleBody production.
	EnterClassModuleBody(c *ClassModuleBodyContext)

	// EnterUnrestrictedName is called when entering the unrestrictedName production.
	EnterUnrestrictedName(c *UnrestrictedNameContext)

	// EnterName is called when entering the name production.
	EnterName(c *NameContext)

	// EnterUntypedName is called when entering the untypedName production.
	EnterUntypedName(c *UntypedNameContext)

	// EnterProceduralModuleDirectiveElement is called when entering the proceduralModuleDirectiveElement production.
	EnterProceduralModuleDirectiveElement(c *ProceduralModuleDirectiveElementContext)

	// EnterProceduralModuleDeclarationElement is called when entering the proceduralModuleDeclarationElement production.
	EnterProceduralModuleDeclarationElement(c *ProceduralModuleDeclarationElementContext)

	// EnterClassModuleDirectiveElement is called when entering the classModuleDirectiveElement production.
	EnterClassModuleDirectiveElement(c *ClassModuleDirectiveElementContext)

	// EnterClassModuleDeclarationElement is called when entering the classModuleDeclarationElement production.
	EnterClassModuleDeclarationElement(c *ClassModuleDeclarationElementContext)

	// EnterCommonOptionDirective is called when entering the commonOptionDirective production.
	EnterCommonOptionDirective(c *CommonOptionDirectiveContext)

	// EnterOptionCompareDirective is called when entering the optionCompareDirective production.
	EnterOptionCompareDirective(c *OptionCompareDirectiveContext)

	// EnterOptionBaseDirective is called when entering the optionBaseDirective production.
	EnterOptionBaseDirective(c *OptionBaseDirectiveContext)

	// EnterOptionExplicitDirective is called when entering the optionExplicitDirective production.
	EnterOptionExplicitDirective(c *OptionExplicitDirectiveContext)

	// EnterOptionPrivateDirective is called when entering the optionPrivateDirective production.
	EnterOptionPrivateDirective(c *OptionPrivateDirectiveContext)

	// EnterDefDirective is called when entering the defDirective production.
	EnterDefDirective(c *DefDirectiveContext)

	// EnterLetterSpec is called when entering the letterSpec production.
	EnterLetterSpec(c *LetterSpecContext)

	// EnterSingleLetter is called when entering the singleLetter production.
	EnterSingleLetter(c *SingleLetterContext)

	// EnterUniversalLetterRange is called when entering the universalLetterRange production.
	EnterUniversalLetterRange(c *UniversalLetterRangeContext)

	// EnterUpperCaseA is called when entering the upperCaseA production.
	EnterUpperCaseA(c *UpperCaseAContext)

	// EnterUpperCaseZ is called when entering the upperCaseZ production.
	EnterUpperCaseZ(c *UpperCaseZContext)

	// EnterLetterRange is called when entering the letterRange production.
	EnterLetterRange(c *LetterRangeContext)

	// EnterFirstLetter is called when entering the firstLetter production.
	EnterFirstLetter(c *FirstLetterContext)

	// EnterLastLetter is called when entering the lastLetter production.
	EnterLastLetter(c *LastLetterContext)

	// EnterDefType is called when entering the defType production.
	EnterDefType(c *DefTypeContext)

	// EnterCommonModuleDeclarationElement is called when entering the commonModuleDeclarationElement production.
	EnterCommonModuleDeclarationElement(c *CommonModuleDeclarationElementContext)

	// EnterVariableDeclaration is called when entering the variableDeclaration production.
	EnterVariableDeclaration(c *VariableDeclarationContext)

	// EnterVariableHelpAttribute is called when entering the variableHelpAttribute production.
	EnterVariableHelpAttribute(c *VariableHelpAttributeContext)

	// EnterVariableModifier is called when entering the variableModifier production.
	EnterVariableModifier(c *VariableModifierContext)

	// EnterVariableSharedModifier is called when entering the variableSharedModifier production.
	EnterVariableSharedModifier(c *VariableSharedModifierContext)

	// EnterModuleVariableDeclarationList is called when entering the moduleVariableDeclarationList production.
	EnterModuleVariableDeclarationList(c *ModuleVariableDeclarationListContext)

	// EnterVariableDeclarationList is called when entering the variableDeclarationList production.
	EnterVariableDeclarationList(c *VariableDeclarationListContext)

	// EnterVariableDcl is called when entering the variableDcl production.
	EnterVariableDcl(c *VariableDclContext)

	// EnterTypedVariableDcl is called when entering the typedVariableDcl production.
	EnterTypedVariableDcl(c *TypedVariableDclContext)

	// EnterUntypedVariableDcl is called when entering the untypedVariableDcl production.
	EnterUntypedVariableDcl(c *UntypedVariableDclContext)

	// EnterArrayClause is called when entering the arrayClause production.
	EnterArrayClause(c *ArrayClauseContext)

	// EnterAsClause is called when entering the asClause production.
	EnterAsClause(c *AsClauseContext)

	// EnterWitheventsVariableDcl is called when entering the witheventsVariableDcl production.
	EnterWitheventsVariableDcl(c *WitheventsVariableDclContext)

	// EnterClassTypeName is called when entering the classTypeName production.
	EnterClassTypeName(c *ClassTypeNameContext)

	// EnterArrayDim is called when entering the arrayDim production.
	EnterArrayDim(c *ArrayDimContext)

	// EnterBoundsList is called when entering the boundsList production.
	EnterBoundsList(c *BoundsListContext)

	// EnterDimSpec is called when entering the dimSpec production.
	EnterDimSpec(c *DimSpecContext)

	// EnterLowerBound is called when entering the lowerBound production.
	EnterLowerBound(c *LowerBoundContext)

	// EnterUpperBound is called when entering the upperBound production.
	EnterUpperBound(c *UpperBoundContext)

	// EnterAsAutoObject is called when entering the asAutoObject production.
	EnterAsAutoObject(c *AsAutoObjectContext)

	// EnterAsType is called when entering the asType production.
	EnterAsType(c *AsTypeContext)

	// EnterTypeSpec is called when entering the typeSpec production.
	EnterTypeSpec(c *TypeSpecContext)

	// EnterFixedLengthStringSpec is called when entering the fixedLengthStringSpec production.
	EnterFixedLengthStringSpec(c *FixedLengthStringSpecContext)

	// EnterStringLength is called when entering the stringLength production.
	EnterStringLength(c *StringLengthContext)

	// EnterConstantName is called when entering the constantName production.
	EnterConstantName(c *ConstantNameContext)

	// EnterConstDeclaration is called when entering the constDeclaration production.
	EnterConstDeclaration(c *ConstDeclarationContext)

	// EnterConstItemList is called when entering the constItemList production.
	EnterConstItemList(c *ConstItemListContext)

	// EnterConstItem is called when entering the constItem production.
	EnterConstItem(c *ConstItemContext)

	// EnterConstAsClause is called when entering the constAsClause production.
	EnterConstAsClause(c *ConstAsClauseContext)

	// EnterPublicTypeDeclaration is called when entering the publicTypeDeclaration production.
	EnterPublicTypeDeclaration(c *PublicTypeDeclarationContext)

	// EnterPrivateTypeDeclaration is called when entering the privateTypeDeclaration production.
	EnterPrivateTypeDeclaration(c *PrivateTypeDeclarationContext)

	// EnterUdtDeclaration is called when entering the udtDeclaration production.
	EnterUdtDeclaration(c *UdtDeclarationContext)

	// EnterUdtMemberList is called when entering the udtMemberList production.
	EnterUdtMemberList(c *UdtMemberListContext)

	// EnterUdtElement is called when entering the udtElement production.
	EnterUdtElement(c *UdtElementContext)

	// EnterUdtMember is called when entering the udtMember production.
	EnterUdtMember(c *UdtMemberContext)

	// EnterUntypedNameMemberDcl is called when entering the untypedNameMemberDcl production.
	EnterUntypedNameMemberDcl(c *UntypedNameMemberDclContext)

	// EnterReservedNameMemberDcl is called when entering the reservedNameMemberDcl production.
	EnterReservedNameMemberDcl(c *ReservedNameMemberDclContext)

	// EnterOptionalArrayClause is called when entering the optionalArrayClause production.
	EnterOptionalArrayClause(c *OptionalArrayClauseContext)

	// EnterReservedMemberName is called when entering the reservedMemberName production.
	EnterReservedMemberName(c *ReservedMemberNameContext)

	// EnterGlobalEnumDeclaration is called when entering the globalEnumDeclaration production.
	EnterGlobalEnumDeclaration(c *GlobalEnumDeclarationContext)

	// EnterPublicEnumDeclaration is called when entering the publicEnumDeclaration production.
	EnterPublicEnumDeclaration(c *PublicEnumDeclarationContext)

	// EnterPrivateEnumDeclaration is called when entering the privateEnumDeclaration production.
	EnterPrivateEnumDeclaration(c *PrivateEnumDeclarationContext)

	// EnterEnumDeclaration is called when entering the enumDeclaration production.
	EnterEnumDeclaration(c *EnumDeclarationContext)

	// EnterEnumLongptrDeclaration is called when entering the enumLongptrDeclaration production.
	EnterEnumLongptrDeclaration(c *EnumLongptrDeclarationContext)

	// EnterEnumMemberList is called when entering the enumMemberList production.
	EnterEnumMemberList(c *EnumMemberListContext)

	// EnterEnumElement is called when entering the enumElement production.
	EnterEnumElement(c *EnumElementContext)

	// EnterEnumMember is called when entering the enumMember production.
	EnterEnumMember(c *EnumMemberContext)

	// EnterPublicExternalProcedureDeclaration is called when entering the publicExternalProcedureDeclaration production.
	EnterPublicExternalProcedureDeclaration(c *PublicExternalProcedureDeclarationContext)

	// EnterPrivateExternalProcedureDeclaration is called when entering the privateExternalProcedureDeclaration production.
	EnterPrivateExternalProcedureDeclaration(c *PrivateExternalProcedureDeclarationContext)

	// EnterExternalProcDcl is called when entering the externalProcDcl production.
	EnterExternalProcDcl(c *ExternalProcDclContext)

	// EnterExternalSub is called when entering the externalSub production.
	EnterExternalSub(c *ExternalSubContext)

	// EnterExternalFunction is called when entering the externalFunction production.
	EnterExternalFunction(c *ExternalFunctionContext)

	// EnterLibInfo is called when entering the libInfo production.
	EnterLibInfo(c *LibInfoContext)

	// EnterLibClause is called when entering the libClause production.
	EnterLibClause(c *LibClauseContext)

	// EnterAliasClause is called when entering the aliasClause production.
	EnterAliasClause(c *AliasClauseContext)

	// EnterImplementsDirective is called when entering the implementsDirective production.
	EnterImplementsDirective(c *ImplementsDirectiveContext)

	// EnterEventDeclaration is called when entering the eventDeclaration production.
	EnterEventDeclaration(c *EventDeclarationContext)

	// EnterEventParameterList is called when entering the eventParameterList production.
	EnterEventParameterList(c *EventParameterListContext)

	// EnterProceduralModuleCode is called when entering the proceduralModuleCode production.
	EnterProceduralModuleCode(c *ProceduralModuleCodeContext)

	// EnterClassModuleCode is called when entering the classModuleCode production.
	EnterClassModuleCode(c *ClassModuleCodeContext)

	// EnterProceduralModuleCodeElement is called when entering the proceduralModuleCodeElement production.
	EnterProceduralModuleCodeElement(c *ProceduralModuleCodeElementContext)

	// EnterClassModuleCodeElement is called when entering the classModuleCodeElement production.
	EnterClassModuleCodeElement(c *ClassModuleCodeElementContext)

	// EnterCommonModuleCodeElement is called when entering the commonModuleCodeElement production.
	EnterCommonModuleCodeElement(c *CommonModuleCodeElementContext)

	// EnterProcedureDeclaration is called when entering the procedureDeclaration production.
	EnterProcedureDeclaration(c *ProcedureDeclarationContext)

	// EnterSubroutineDeclaration is called when entering the subroutineDeclaration production.
	EnterSubroutineDeclaration(c *SubroutineDeclarationContext)

	// EnterFunctionDeclaration is called when entering the functionDeclaration production.
	EnterFunctionDeclaration(c *FunctionDeclarationContext)

	// EnterPropertyGetDeclaration is called when entering the propertyGetDeclaration production.
	EnterPropertyGetDeclaration(c *PropertyGetDeclarationContext)

	// EnterPropertySetDeclaration is called when entering the propertySetDeclaration production.
	EnterPropertySetDeclaration(c *PropertySetDeclarationContext)

	// EnterEndLabel is called when entering the endLabel production.
	EnterEndLabel(c *EndLabelContext)

	// EnterProcedureTail is called when entering the procedureTail production.
	EnterProcedureTail(c *ProcedureTailContext)

	// EnterProcedureScope is called when entering the procedureScope production.
	EnterProcedureScope(c *ProcedureScopeContext)

	// EnterInitialStatic is called when entering the initialStatic production.
	EnterInitialStatic(c *InitialStaticContext)

	// EnterTrailingStatic is called when entering the trailingStatic production.
	EnterTrailingStatic(c *TrailingStaticContext)

	// EnterSubroutineName is called when entering the subroutineName production.
	EnterSubroutineName(c *SubroutineNameContext)

	// EnterFunctionName is called when entering the functionName production.
	EnterFunctionName(c *FunctionNameContext)

	// EnterPrefixedName is called when entering the prefixedName production.
	EnterPrefixedName(c *PrefixedNameContext)

	// EnterFunctionType is called when entering the functionType production.
	EnterFunctionType(c *FunctionTypeContext)

	// EnterArrayDesignator is called when entering the arrayDesignator production.
	EnterArrayDesignator(c *ArrayDesignatorContext)

	// EnterProcedureParameters is called when entering the procedureParameters production.
	EnterProcedureParameters(c *ProcedureParametersContext)

	// EnterPropertyParameters is called when entering the propertyParameters production.
	EnterPropertyParameters(c *PropertyParametersContext)

	// EnterPropertyParameterList is called when entering the propertyParameterList production.
	EnterPropertyParameterList(c *PropertyParameterListContext)

	// EnterParameterList is called when entering the parameterList production.
	EnterParameterList(c *ParameterListContext)

	// EnterAnyParam is called when entering the anyParam production.
	EnterAnyParam(c *AnyParamContext)

	// EnterPositionalParameters is called when entering the positionalParameters production.
	EnterPositionalParameters(c *PositionalParametersContext)

	// EnterValueParam is called when entering the valueParam production.
	EnterValueParam(c *ValueParamContext)

	// EnterPositionalParam is called when entering the positionalParam production.
	EnterPositionalParam(c *PositionalParamContext)

	// EnterOptionalParam is called when entering the optionalParam production.
	EnterOptionalParam(c *OptionalParamContext)

	// EnterParamArray is called when entering the paramArray production.
	EnterParamArray(c *ParamArrayContext)

	// EnterParamDcl is called when entering the paramDcl production.
	EnterParamDcl(c *ParamDclContext)

	// EnterUntypedNameParamDcl is called when entering the untypedNameParamDcl production.
	EnterUntypedNameParamDcl(c *UntypedNameParamDclContext)

	// EnterTypedNameParamDcl is called when entering the typedNameParamDcl production.
	EnterTypedNameParamDcl(c *TypedNameParamDclContext)

	// EnterOptionalPrefix is called when entering the optionalPrefix production.
	EnterOptionalPrefix(c *OptionalPrefixContext)

	// EnterParameterMechanism is called when entering the parameterMechanism production.
	EnterParameterMechanism(c *ParameterMechanismContext)

	// EnterParameterType is called when entering the parameterType production.
	EnterParameterType(c *ParameterTypeContext)

	// EnterDefaultValue is called when entering the defaultValue production.
	EnterDefaultValue(c *DefaultValueContext)

	// EnterEventHandlerName is called when entering the eventHandlerName production.
	EnterEventHandlerName(c *EventHandlerNameContext)

	// EnterImplementedName is called when entering the implementedName production.
	EnterImplementedName(c *ImplementedNameContext)

	// EnterLifecycleHandlerName is called when entering the lifecycleHandlerName production.
	EnterLifecycleHandlerName(c *LifecycleHandlerNameContext)

	// EnterProcedureBody is called when entering the procedureBody production.
	EnterProcedureBody(c *ProcedureBodyContext)

	// EnterStatementBlock is called when entering the statementBlock production.
	EnterStatementBlock(c *StatementBlockContext)

	// EnterBlockStatement is called when entering the blockStatement production.
	EnterBlockStatement(c *BlockStatementContext)

	// EnterStatement is called when entering the statement production.
	EnterStatement(c *StatementContext)

	// EnterStatementLabelDefinition is called when entering the statementLabelDefinition production.
	EnterStatementLabelDefinition(c *StatementLabelDefinitionContext)

	// EnterStatementLabel is called when entering the statementLabel production.
	EnterStatementLabel(c *StatementLabelContext)

	// EnterStatementLabelList is called when entering the statementLabelList production.
	EnterStatementLabelList(c *StatementLabelListContext)

	// EnterIdentifierStatementLabel is called when entering the identifierStatementLabel production.
	EnterIdentifierStatementLabel(c *IdentifierStatementLabelContext)

	// EnterResetNumberLabel is called when entering the resetNumberLabel production.
	EnterResetNumberLabel(c *ResetNumberLabelContext)

	// EnterLineNumberLabel is called when entering the lineNumberLabel production.
	EnterLineNumberLabel(c *LineNumberLabelContext)

	// EnterRemStatement is called when entering the remStatement production.
	EnterRemStatement(c *RemStatementContext)

	// EnterControlStatement is called when entering the controlStatement production.
	EnterControlStatement(c *ControlStatementContext)

	// EnterControlStatementExceptMultilineIf is called when entering the controlStatementExceptMultilineIf production.
	EnterControlStatementExceptMultilineIf(c *ControlStatementExceptMultilineIfContext)

	// EnterCallStatement is called when entering the callStatement production.
	EnterCallStatement(c *CallStatementContext)

	// EnterWhileStatement is called when entering the whileStatement production.
	EnterWhileStatement(c *WhileStatementContext)

	// EnterForStatement is called when entering the forStatement production.
	EnterForStatement(c *ForStatementContext)

	// EnterForTail is called when entering the forTail production.
	EnterForTail(c *ForTailContext)

	// EnterForClause is called when entering the forClause production.
	EnterForClause(c *ForClauseContext)

	// EnterStartValue is called when entering the startValue production.
	EnterStartValue(c *StartValueContext)

	// EnterEndValue is called when entering the endValue production.
	EnterEndValue(c *EndValueContext)

	// EnterStepClause is called when entering the stepClause production.
	EnterStepClause(c *StepClauseContext)

	// EnterStepIncrement is called when entering the stepIncrement production.
	EnterStepIncrement(c *StepIncrementContext)

	// EnterForEachStatement is called when entering the forEachStatement production.
	EnterForEachStatement(c *ForEachStatementContext)

	// EnterForEachClause is called when entering the forEachClause production.
	EnterForEachClause(c *ForEachClauseContext)

	// EnterCollection is called when entering the collection production.
	EnterCollection(c *CollectionContext)

	// EnterExitForStatement is called when entering the exitForStatement production.
	EnterExitForStatement(c *ExitForStatementContext)

	// EnterDoStatement is called when entering the doStatement production.
	EnterDoStatement(c *DoStatementContext)

	// EnterConditionClause is called when entering the conditionClause production.
	EnterConditionClause(c *ConditionClauseContext)

	// EnterWhileClause is called when entering the whileClause production.
	EnterWhileClause(c *WhileClauseContext)

	// EnterUntilClause is called when entering the untilClause production.
	EnterUntilClause(c *UntilClauseContext)

	// EnterExitDoStatement is called when entering the exitDoStatement production.
	EnterExitDoStatement(c *ExitDoStatementContext)

	// EnterIfStatement is called when entering the ifStatement production.
	EnterIfStatement(c *IfStatementContext)

	// EnterElseIfBlock is called when entering the elseIfBlock production.
	EnterElseIfBlock(c *ElseIfBlockContext)

	// EnterElseBlock is called when entering the elseBlock production.
	EnterElseBlock(c *ElseBlockContext)

	// EnterSingleLineIfStatement is called when entering the singleLineIfStatement production.
	EnterSingleLineIfStatement(c *SingleLineIfStatementContext)

	// EnterIfWithNonEmptyThen is called when entering the ifWithNonEmptyThen production.
	EnterIfWithNonEmptyThen(c *IfWithNonEmptyThenContext)

	// EnterIfWithEmptyThen is called when entering the ifWithEmptyThen production.
	EnterIfWithEmptyThen(c *IfWithEmptyThenContext)

	// EnterSingleLineElseClause is called when entering the singleLineElseClause production.
	EnterSingleLineElseClause(c *SingleLineElseClauseContext)

	// EnterListOrLabel is called when entering the listOrLabel production.
	EnterListOrLabel(c *ListOrLabelContext)

	// EnterSameLineStatement is called when entering the sameLineStatement production.
	EnterSameLineStatement(c *SameLineStatementContext)

	// EnterSelectCaseStatement is called when entering the selectCaseStatement production.
	EnterSelectCaseStatement(c *SelectCaseStatementContext)

	// EnterCaseClause is called when entering the caseClause production.
	EnterCaseClause(c *CaseClauseContext)

	// EnterCaseElseClause is called when entering the caseElseClause production.
	EnterCaseElseClause(c *CaseElseClauseContext)

	// EnterRangeClause is called when entering the rangeClause production.
	EnterRangeClause(c *RangeClauseContext)

	// EnterSelectExpression is called when entering the selectExpression production.
	EnterSelectExpression(c *SelectExpressionContext)

	// EnterComparisonOperator is called when entering the comparisonOperator production.
	EnterComparisonOperator(c *ComparisonOperatorContext)

	// EnterStopStatement is called when entering the stopStatement production.
	EnterStopStatement(c *StopStatementContext)

	// EnterGotoStatement is called when entering the gotoStatement production.
	EnterGotoStatement(c *GotoStatementContext)

	// EnterOnGotoStatement is called when entering the onGotoStatement production.
	EnterOnGotoStatement(c *OnGotoStatementContext)

	// EnterGosubStatement is called when entering the gosubStatement production.
	EnterGosubStatement(c *GosubStatementContext)

	// EnterReturnStatement is called when entering the returnStatement production.
	EnterReturnStatement(c *ReturnStatementContext)

	// EnterOnGosubStatement is called when entering the onGosubStatement production.
	EnterOnGosubStatement(c *OnGosubStatementContext)

	// EnterExitSubStatement is called when entering the exitSubStatement production.
	EnterExitSubStatement(c *ExitSubStatementContext)

	// EnterExitFunctionStatement is called when entering the exitFunctionStatement production.
	EnterExitFunctionStatement(c *ExitFunctionStatementContext)

	// EnterExitPropertyStatement is called when entering the exitPropertyStatement production.
	EnterExitPropertyStatement(c *ExitPropertyStatementContext)

	// EnterRaiseeventStatement is called when entering the raiseeventStatement production.
	EnterRaiseeventStatement(c *RaiseeventStatementContext)

	// EnterEventArgumentList is called when entering the eventArgumentList production.
	EnterEventArgumentList(c *EventArgumentListContext)

	// EnterEventArgument is called when entering the eventArgument production.
	EnterEventArgument(c *EventArgumentContext)

	// EnterWithStatement is called when entering the withStatement production.
	EnterWithStatement(c *WithStatementContext)

	// EnterEndStatement is called when entering the endStatement production.
	EnterEndStatement(c *EndStatementContext)

	// EnterDataManipulationStatement is called when entering the dataManipulationStatement production.
	EnterDataManipulationStatement(c *DataManipulationStatementContext)

	// EnterStaticVariableDeclaration is called when entering the staticVariableDeclaration production.
	EnterStaticVariableDeclaration(c *StaticVariableDeclarationContext)

	// EnterRedimStatement is called when entering the redimStatement production.
	EnterRedimStatement(c *RedimStatementContext)

	// EnterRedimDeclarationList is called when entering the redimDeclarationList production.
	EnterRedimDeclarationList(c *RedimDeclarationListContext)

	// EnterRedimVariableDcl is called when entering the redimVariableDcl production.
	EnterRedimVariableDcl(c *RedimVariableDclContext)

	// EnterRedimTypedVariableDcl is called when entering the redimTypedVariableDcl production.
	EnterRedimTypedVariableDcl(c *RedimTypedVariableDclContext)

	// EnterRedimUntypedDcl is called when entering the redimUntypedDcl production.
	EnterRedimUntypedDcl(c *RedimUntypedDclContext)

	// EnterWithExpressionDcl is called when entering the withExpressionDcl production.
	EnterWithExpressionDcl(c *WithExpressionDclContext)

	// EnterMemberAccessExpressionDcl is called when entering the memberAccessExpressionDcl production.
	EnterMemberAccessExpressionDcl(c *MemberAccessExpressionDclContext)

	// EnterDynamicArrayDim is called when entering the dynamicArrayDim production.
	EnterDynamicArrayDim(c *DynamicArrayDimContext)

	// EnterDynamicBoundsList is called when entering the dynamicBoundsList production.
	EnterDynamicBoundsList(c *DynamicBoundsListContext)

	// EnterDynamicDimSpec is called when entering the dynamicDimSpec production.
	EnterDynamicDimSpec(c *DynamicDimSpecContext)

	// EnterDynamicLowerBound is called when entering the dynamicLowerBound production.
	EnterDynamicLowerBound(c *DynamicLowerBoundContext)

	// EnterDynamicUpperBound is called when entering the dynamicUpperBound production.
	EnterDynamicUpperBound(c *DynamicUpperBoundContext)

	// EnterDynamicArrayClause is called when entering the dynamicArrayClause production.
	EnterDynamicArrayClause(c *DynamicArrayClauseContext)

	// EnterEraseStatement is called when entering the eraseStatement production.
	EnterEraseStatement(c *EraseStatementContext)

	// EnterEraseList is called when entering the eraseList production.
	EnterEraseList(c *EraseListContext)

	// EnterEraseElement is called when entering the eraseElement production.
	EnterEraseElement(c *EraseElementContext)

	// EnterMidStatement is called when entering the midStatement production.
	EnterMidStatement(c *MidStatementContext)

	// EnterModeSpecifier is called when entering the modeSpecifier production.
	EnterModeSpecifier(c *ModeSpecifierContext)

	// EnterStringArgument is called when entering the stringArgument production.
	EnterStringArgument(c *StringArgumentContext)

	// EnterStartMid is called when entering the startMid production.
	EnterStartMid(c *StartMidContext)

	// EnterLength is called when entering the length production.
	EnterLength(c *LengthContext)

	// EnterLsetStatement is called when entering the lsetStatement production.
	EnterLsetStatement(c *LsetStatementContext)

	// EnterRsetStatement is called when entering the rsetStatement production.
	EnterRsetStatement(c *RsetStatementContext)

	// EnterLetStatement is called when entering the letStatement production.
	EnterLetStatement(c *LetStatementContext)

	// EnterSetStatement is called when entering the setStatement production.
	EnterSetStatement(c *SetStatementContext)

	// EnterErrorHandlingStatement is called when entering the errorHandlingStatement production.
	EnterErrorHandlingStatement(c *ErrorHandlingStatementContext)

	// EnterOnErrorStatement is called when entering the onErrorStatement production.
	EnterOnErrorStatement(c *OnErrorStatementContext)

	// EnterErrorBehavior is called when entering the errorBehavior production.
	EnterErrorBehavior(c *ErrorBehaviorContext)

	// EnterResumeStatement is called when entering the resumeStatement production.
	EnterResumeStatement(c *ResumeStatementContext)

	// EnterErrorStatement is called when entering the errorStatement production.
	EnterErrorStatement(c *ErrorStatementContext)

	// EnterErrorNumber is called when entering the errorNumber production.
	EnterErrorNumber(c *ErrorNumberContext)

	// EnterFileStatement is called when entering the fileStatement production.
	EnterFileStatement(c *FileStatementContext)

	// EnterOpenStatement is called when entering the openStatement production.
	EnterOpenStatement(c *OpenStatementContext)

	// EnterPathName is called when entering the pathName production.
	EnterPathName(c *PathNameContext)

	// EnterModeClause is called when entering the modeClause production.
	EnterModeClause(c *ModeClauseContext)

	// EnterModeOpt is called when entering the modeOpt production.
	EnterModeOpt(c *ModeOptContext)

	// EnterAccessClause is called when entering the accessClause production.
	EnterAccessClause(c *AccessClauseContext)

	// EnterAccess is called when entering the access production.
	EnterAccess(c *AccessContext)

	// EnterLock is called when entering the lock production.
	EnterLock(c *LockContext)

	// EnterLenClause is called when entering the lenClause production.
	EnterLenClause(c *LenClauseContext)

	// EnterRecLength is called when entering the recLength production.
	EnterRecLength(c *RecLengthContext)

	// EnterFileNumber is called when entering the fileNumber production.
	EnterFileNumber(c *FileNumberContext)

	// EnterMarkedFileNumber is called when entering the markedFileNumber production.
	EnterMarkedFileNumber(c *MarkedFileNumberContext)

	// EnterUnmarkedFileNumber is called when entering the unmarkedFileNumber production.
	EnterUnmarkedFileNumber(c *UnmarkedFileNumberContext)

	// EnterCloseStatement is called when entering the closeStatement production.
	EnterCloseStatement(c *CloseStatementContext)

	// EnterFileNumberList is called when entering the fileNumberList production.
	EnterFileNumberList(c *FileNumberListContext)

	// EnterSeekStatement is called when entering the seekStatement production.
	EnterSeekStatement(c *SeekStatementContext)

	// EnterPosition is called when entering the position production.
	EnterPosition(c *PositionContext)

	// EnterLockStatement is called when entering the lockStatement production.
	EnterLockStatement(c *LockStatementContext)

	// EnterRecordRange is called when entering the recordRange production.
	EnterRecordRange(c *RecordRangeContext)

	// EnterStartRecordNumber is called when entering the startRecordNumber production.
	EnterStartRecordNumber(c *StartRecordNumberContext)

	// EnterEndRecordNumber is called when entering the endRecordNumber production.
	EnterEndRecordNumber(c *EndRecordNumberContext)

	// EnterUnlockStatement is called when entering the unlockStatement production.
	EnterUnlockStatement(c *UnlockStatementContext)

	// EnterLineInputStatement is called when entering the lineInputStatement production.
	EnterLineInputStatement(c *LineInputStatementContext)

	// EnterVariableName is called when entering the variableName production.
	EnterVariableName(c *VariableNameContext)

	// EnterWidthStatement is called when entering the widthStatement production.
	EnterWidthStatement(c *WidthStatementContext)

	// EnterLineWidth is called when entering the lineWidth production.
	EnterLineWidth(c *LineWidthContext)

	// EnterPrintStatement is called when entering the printStatement production.
	EnterPrintStatement(c *PrintStatementContext)

	// EnterOutputList is called when entering the outputList production.
	EnterOutputList(c *OutputListContext)

	// EnterOutputItem is called when entering the outputItem production.
	EnterOutputItem(c *OutputItemContext)

	// EnterOutputClause is called when entering the outputClause production.
	EnterOutputClause(c *OutputClauseContext)

	// EnterCharPosition is called when entering the charPosition production.
	EnterCharPosition(c *CharPositionContext)

	// EnterOutputExpression is called when entering the outputExpression production.
	EnterOutputExpression(c *OutputExpressionContext)

	// EnterSpcClause is called when entering the spcClause production.
	EnterSpcClause(c *SpcClauseContext)

	// EnterSpcNumber is called when entering the spcNumber production.
	EnterSpcNumber(c *SpcNumberContext)

	// EnterTabClause is called when entering the tabClause production.
	EnterTabClause(c *TabClauseContext)

	// EnterTabNumber is called when entering the tabNumber production.
	EnterTabNumber(c *TabNumberContext)

	// EnterWriteStatement is called when entering the writeStatement production.
	EnterWriteStatement(c *WriteStatementContext)

	// EnterInputStatement is called when entering the inputStatement production.
	EnterInputStatement(c *InputStatementContext)

	// EnterInputList is called when entering the inputList production.
	EnterInputList(c *InputListContext)

	// EnterInputVariable is called when entering the inputVariable production.
	EnterInputVariable(c *InputVariableContext)

	// EnterPutStatement is called when entering the putStatement production.
	EnterPutStatement(c *PutStatementContext)

	// EnterRecordNumber is called when entering the recordNumber production.
	EnterRecordNumber(c *RecordNumberContext)

	// EnterData is called when entering the data production.
	EnterData(c *DataContext)

	// EnterGetStatement is called when entering the getStatement production.
	EnterGetStatement(c *GetStatementContext)

	// EnterVariable is called when entering the variable production.
	EnterVariable(c *VariableContext)

	// EnterAttributeStatement is called when entering the attributeStatement production.
	EnterAttributeStatement(c *AttributeStatementContext)

	// EnterAttributeDescName is called when entering the attributeDescName production.
	EnterAttributeDescName(c *AttributeDescNameContext)

	// EnterAttributeUsrName is called when entering the attributeUsrName production.
	EnterAttributeUsrName(c *AttributeUsrNameContext)

	// EnterDebugStatement is called when entering the debugStatement production.
	EnterDebugStatement(c *DebugStatementContext)

	// EnterDebugArgs is called when entering the debugArgs production.
	EnterDebugArgs(c *DebugArgsContext)

	// EnterDebugSep is called when entering the debugSep production.
	EnterDebugSep(c *DebugSepContext)

	// EnterExpression is called when entering the expression production.
	EnterExpression(c *ExpressionContext)

	// EnterLExpression is called when entering the lExpression production.
	EnterLExpression(c *LExpressionContext)

	// EnterLiteralExpression is called when entering the literalExpression production.
	EnterLiteralExpression(c *LiteralExpressionContext)

	// EnterParenthesizedExpression is called when entering the parenthesizedExpression production.
	EnterParenthesizedExpression(c *ParenthesizedExpressionContext)

	// EnterTypeofIsExpression is called when entering the typeofIsExpression production.
	EnterTypeofIsExpression(c *TypeofIsExpressionContext)

	// EnterNewExpress is called when entering the newExpress production.
	EnterNewExpress(c *NewExpressContext)

	// EnterNotOperatorExpression is called when entering the notOperatorExpression production.
	EnterNotOperatorExpression(c *NotOperatorExpressionContext)

	// EnterUnaryMinusExpression is called when entering the unaryMinusExpression production.
	EnterUnaryMinusExpression(c *UnaryMinusExpressionContext)

	// EnterSimpleNameExpression is called when entering the simpleNameExpression production.
	EnterSimpleNameExpression(c *SimpleNameExpressionContext)

	// EnterInstanceExpression is called when entering the instanceExpression production.
	EnterInstanceExpression(c *InstanceExpressionContext)

	// EnterMemberAccessExpression is called when entering the memberAccessExpression production.
	EnterMemberAccessExpression(c *MemberAccessExpressionContext)

	// EnterIndexExpression is called when entering the indexExpression production.
	EnterIndexExpression(c *IndexExpressionContext)

	// EnterArgumentList is called when entering the argumentList production.
	EnterArgumentList(c *ArgumentListContext)

	// EnterPositionalOrNamedArgumentList is called when entering the positionalOrNamedArgumentList production.
	EnterPositionalOrNamedArgumentList(c *PositionalOrNamedArgumentListContext)

	// EnterNamedArgumentList is called when entering the namedArgumentList production.
	EnterNamedArgumentList(c *NamedArgumentListContext)

	// EnterNamedArgument is called when entering the namedArgument production.
	EnterNamedArgument(c *NamedArgumentContext)

	// EnterArgumentExpression is called when entering the argumentExpression production.
	EnterArgumentExpression(c *ArgumentExpressionContext)

	// EnterDictionaryAccessExpression is called when entering the dictionaryAccessExpression production.
	EnterDictionaryAccessExpression(c *DictionaryAccessExpressionContext)

	// EnterWithExpression is called when entering the withExpression production.
	EnterWithExpression(c *WithExpressionContext)

	// EnterWithMemberAccessExpression is called when entering the withMemberAccessExpression production.
	EnterWithMemberAccessExpression(c *WithMemberAccessExpressionContext)

	// EnterWithDictionaryAccessExpression is called when entering the withDictionaryAccessExpression production.
	EnterWithDictionaryAccessExpression(c *WithDictionaryAccessExpressionContext)

	// EnterConstantExpression is called when entering the constantExpression production.
	EnterConstantExpression(c *ConstantExpressionContext)

	// EnterCcExpression is called when entering the ccExpression production.
	EnterCcExpression(c *CcExpressionContext)

	// EnterBooleanExpression is called when entering the booleanExpression production.
	EnterBooleanExpression(c *BooleanExpressionContext)

	// EnterIntegerExpression is called when entering the integerExpression production.
	EnterIntegerExpression(c *IntegerExpressionContext)

	// EnterVariableExpression is called when entering the variableExpression production.
	EnterVariableExpression(c *VariableExpressionContext)

	// EnterBoundVariableExpression is called when entering the boundVariableExpression production.
	EnterBoundVariableExpression(c *BoundVariableExpressionContext)

	// EnterTypeExpression is called when entering the typeExpression production.
	EnterTypeExpression(c *TypeExpressionContext)

	// EnterDefinedTypeExpression is called when entering the definedTypeExpression production.
	EnterDefinedTypeExpression(c *DefinedTypeExpressionContext)

	// EnterAddressofExpression is called when entering the addressofExpression production.
	EnterAddressofExpression(c *AddressofExpressionContext)

	// EnterProcedurePointerExpression is called when entering the procedurePointerExpression production.
	EnterProcedurePointerExpression(c *ProcedurePointerExpressionContext)

	// EnterWsc is called when entering the wsc production.
	EnterWsc(c *WscContext)

	// EnterEndOfLine is called when entering the endOfLine production.
	EnterEndOfLine(c *EndOfLineContext)

	// EnterUnexpectedEndOfLine is called when entering the unexpectedEndOfLine production.
	EnterUnexpectedEndOfLine(c *UnexpectedEndOfLineContext)

	// EnterWscu is called when entering the wscu production.
	EnterWscu(c *WscuContext)

	// EnterEndOfLineNoWs is called when entering the endOfLineNoWs production.
	EnterEndOfLineNoWs(c *EndOfLineNoWsContext)

	// EnterEndOfStatement is called when entering the endOfStatement production.
	EnterEndOfStatement(c *EndOfStatementContext)

	// EnterEndOfStatementNoWs is called when entering the endOfStatementNoWs production.
	EnterEndOfStatementNoWs(c *EndOfStatementNoWsContext)

	// EnterCommentBody is called when entering the commentBody production.
	EnterCommentBody(c *CommentBodyContext)

	// EnterReservedIdentifier is called when entering the reservedIdentifier production.
	EnterReservedIdentifier(c *ReservedIdentifierContext)

	// EnterAmbiguousIdentifier is called when entering the ambiguousIdentifier production.
	EnterAmbiguousIdentifier(c *AmbiguousIdentifierContext)

	// EnterStatementKeyword is called when entering the statementKeyword production.
	EnterStatementKeyword(c *StatementKeywordContext)

	// EnterRemKeyword is called when entering the remKeyword production.
	EnterRemKeyword(c *RemKeywordContext)

	// EnterMarkerKeyword is called when entering the markerKeyword production.
	EnterMarkerKeyword(c *MarkerKeywordContext)

	// EnterOperatorIdentifier is called when entering the operatorIdentifier production.
	EnterOperatorIdentifier(c *OperatorIdentifierContext)

	// EnterReservedName is called when entering the reservedName production.
	EnterReservedName(c *ReservedNameContext)

	// EnterSpecialForm is called when entering the specialForm production.
	EnterSpecialForm(c *SpecialFormContext)

	// EnterReservedTypeIdentifier is called when entering the reservedTypeIdentifier production.
	EnterReservedTypeIdentifier(c *ReservedTypeIdentifierContext)

	// EnterReservedTypeIdentifierB is called when entering the reservedTypeIdentifierB production.
	EnterReservedTypeIdentifierB(c *ReservedTypeIdentifierBContext)

	// EnterLiteralIdentifier is called when entering the literalIdentifier production.
	EnterLiteralIdentifier(c *LiteralIdentifierContext)

	// EnterBooleanLiteralIdentifier is called when entering the booleanLiteralIdentifier production.
	EnterBooleanLiteralIdentifier(c *BooleanLiteralIdentifierContext)

	// EnterObjectLiteralIdentifier is called when entering the objectLiteralIdentifier production.
	EnterObjectLiteralIdentifier(c *ObjectLiteralIdentifierContext)

	// EnterVariantLiteralIdentifier is called when entering the variantLiteralIdentifier production.
	EnterVariantLiteralIdentifier(c *VariantLiteralIdentifierContext)

	// EnterReservedForImplementationUse is called when entering the reservedForImplementationUse production.
	EnterReservedForImplementationUse(c *ReservedForImplementationUseContext)

	// EnterFutureReserved is called when entering the futureReserved production.
	EnterFutureReserved(c *FutureReservedContext)

	// EnterBuiltinType is called when entering the builtinType production.
	EnterBuiltinType(c *BuiltinTypeContext)

	// EnterTypedName is called when entering the typedName production.
	EnterTypedName(c *TypedNameContext)

	// EnterTypeSuffix is called when entering the typeSuffix production.
	EnterTypeSuffix(c *TypeSuffixContext)

	// EnterAmbiguousKeyword is called when entering the ambiguousKeyword production.
	EnterAmbiguousKeyword(c *AmbiguousKeywordContext)

	// EnterAnyOperator is called when entering the anyOperator production.
	EnterAnyOperator(c *AnyOperatorContext)

	// EnterPowOperator is called when entering the powOperator production.
	EnterPowOperator(c *PowOperatorContext)

	// EnterDivOperator is called when entering the divOperator production.
	EnterDivOperator(c *DivOperatorContext)

	// EnterMultOperator is called when entering the multOperator production.
	EnterMultOperator(c *MultOperatorContext)

	// EnterModOperator is called when entering the modOperator production.
	EnterModOperator(c *ModOperatorContext)

	// EnterPlusOperator is called when entering the plusOperator production.
	EnterPlusOperator(c *PlusOperatorContext)

	// EnterMinusOperator is called when entering the minusOperator production.
	EnterMinusOperator(c *MinusOperatorContext)

	// EnterAmpOperator is called when entering the ampOperator production.
	EnterAmpOperator(c *AmpOperatorContext)

	// EnterIsOperator is called when entering the isOperator production.
	EnterIsOperator(c *IsOperatorContext)

	// EnterLikeOperator is called when entering the likeOperator production.
	EnterLikeOperator(c *LikeOperatorContext)

	// EnterGeqOperator is called when entering the geqOperator production.
	EnterGeqOperator(c *GeqOperatorContext)

	// EnterLeqOperator is called when entering the leqOperator production.
	EnterLeqOperator(c *LeqOperatorContext)

	// EnterGtOperator is called when entering the gtOperator production.
	EnterGtOperator(c *GtOperatorContext)

	// EnterLtOperator is called when entering the ltOperator production.
	EnterLtOperator(c *LtOperatorContext)

	// EnterNeqOperator is called when entering the neqOperator production.
	EnterNeqOperator(c *NeqOperatorContext)

	// EnterEqOperator is called when entering the eqOperator production.
	EnterEqOperator(c *EqOperatorContext)

	// EnterAndOperator is called when entering the andOperator production.
	EnterAndOperator(c *AndOperatorContext)

	// EnterOrOperator is called when entering the orOperator production.
	EnterOrOperator(c *OrOperatorContext)

	// EnterXorOperator is called when entering the xorOperator production.
	EnterXorOperator(c *XorOperatorContext)

	// EnterEqvOperator is called when entering the eqvOperator production.
	EnterEqvOperator(c *EqvOperatorContext)

	// EnterImpOperator is called when entering the impOperator production.
	EnterImpOperator(c *ImpOperatorContext)

	// ExitStartRule is called when exiting the startRule production.
	ExitStartRule(c *StartRuleContext)

	// ExitModule is called when exiting the module production.
	ExitModule(c *ModuleContext)

	// ExitClassFileHeader is called when exiting the classFileHeader production.
	ExitClassFileHeader(c *ClassFileHeaderContext)

	// ExitClassVersionIdentification is called when exiting the classVersionIdentification production.
	ExitClassVersionIdentification(c *ClassVersionIdentificationContext)

	// ExitClassBeginBlock is called when exiting the classBeginBlock production.
	ExitClassBeginBlock(c *ClassBeginBlockContext)

	// ExitBeginBlockConfigElement is called when exiting the beginBlockConfigElement production.
	ExitBeginBlockConfigElement(c *BeginBlockConfigElementContext)

	// ExitFormFileHeader is called when exiting the formFileHeader production.
	ExitFormFileHeader(c *FormFileHeaderContext)

	// ExitFormVersionIdentification is called when exiting the formVersionIdentification production.
	ExitFormVersionIdentification(c *FormVersionIdentificationContext)

	// ExitFormObjectAssign is called when exiting the formObjectAssign production.
	ExitFormObjectAssign(c *FormObjectAssignContext)

	// ExitFormBeginBlock is called when exiting the formBeginBlock production.
	ExitFormBeginBlock(c *FormBeginBlockContext)

	// ExitBeginPropertyBlock is called when exiting the beginPropertyBlock production.
	ExitBeginPropertyBlock(c *BeginPropertyBlockContext)

	// ExitProceduralModule is called when exiting the proceduralModule production.
	ExitProceduralModule(c *ProceduralModuleContext)

	// ExitClassModule is called when exiting the classModule production.
	ExitClassModule(c *ClassModuleContext)

	// ExitProceduralModuleHeader is called when exiting the proceduralModuleHeader production.
	ExitProceduralModuleHeader(c *ProceduralModuleHeaderContext)

	// ExitProceduralModuleAttr is called when exiting the proceduralModuleAttr production.
	ExitProceduralModuleAttr(c *ProceduralModuleAttrContext)

	// ExitIgnoredProceduralAttr is called when exiting the ignoredProceduralAttr production.
	ExitIgnoredProceduralAttr(c *IgnoredProceduralAttrContext)

	// ExitClassModuleHeader is called when exiting the classModuleHeader production.
	ExitClassModuleHeader(c *ClassModuleHeaderContext)

	// ExitClassAttr is called when exiting the classAttr production.
	ExitClassAttr(c *ClassAttrContext)

	// ExitIgnoredClassAttr is called when exiting the ignoredClassAttr production.
	ExitIgnoredClassAttr(c *IgnoredClassAttrContext)

	// ExitIgnoredAttr is called when exiting the ignoredAttr production.
	ExitIgnoredAttr(c *IgnoredAttrContext)

	// ExitNameAttr is called when exiting the nameAttr production.
	ExitNameAttr(c *NameAttrContext)

	// ExitProceduralModuleBody is called when exiting the proceduralModuleBody production.
	ExitProceduralModuleBody(c *ProceduralModuleBodyContext)

	// ExitClassModuleBody is called when exiting the classModuleBody production.
	ExitClassModuleBody(c *ClassModuleBodyContext)

	// ExitUnrestrictedName is called when exiting the unrestrictedName production.
	ExitUnrestrictedName(c *UnrestrictedNameContext)

	// ExitName is called when exiting the name production.
	ExitName(c *NameContext)

	// ExitUntypedName is called when exiting the untypedName production.
	ExitUntypedName(c *UntypedNameContext)

	// ExitProceduralModuleDirectiveElement is called when exiting the proceduralModuleDirectiveElement production.
	ExitProceduralModuleDirectiveElement(c *ProceduralModuleDirectiveElementContext)

	// ExitProceduralModuleDeclarationElement is called when exiting the proceduralModuleDeclarationElement production.
	ExitProceduralModuleDeclarationElement(c *ProceduralModuleDeclarationElementContext)

	// ExitClassModuleDirectiveElement is called when exiting the classModuleDirectiveElement production.
	ExitClassModuleDirectiveElement(c *ClassModuleDirectiveElementContext)

	// ExitClassModuleDeclarationElement is called when exiting the classModuleDeclarationElement production.
	ExitClassModuleDeclarationElement(c *ClassModuleDeclarationElementContext)

	// ExitCommonOptionDirective is called when exiting the commonOptionDirective production.
	ExitCommonOptionDirective(c *CommonOptionDirectiveContext)

	// ExitOptionCompareDirective is called when exiting the optionCompareDirective production.
	ExitOptionCompareDirective(c *OptionCompareDirectiveContext)

	// ExitOptionBaseDirective is called when exiting the optionBaseDirective production.
	ExitOptionBaseDirective(c *OptionBaseDirectiveContext)

	// ExitOptionExplicitDirective is called when exiting the optionExplicitDirective production.
	ExitOptionExplicitDirective(c *OptionExplicitDirectiveContext)

	// ExitOptionPrivateDirective is called when exiting the optionPrivateDirective production.
	ExitOptionPrivateDirective(c *OptionPrivateDirectiveContext)

	// ExitDefDirective is called when exiting the defDirective production.
	ExitDefDirective(c *DefDirectiveContext)

	// ExitLetterSpec is called when exiting the letterSpec production.
	ExitLetterSpec(c *LetterSpecContext)

	// ExitSingleLetter is called when exiting the singleLetter production.
	ExitSingleLetter(c *SingleLetterContext)

	// ExitUniversalLetterRange is called when exiting the universalLetterRange production.
	ExitUniversalLetterRange(c *UniversalLetterRangeContext)

	// ExitUpperCaseA is called when exiting the upperCaseA production.
	ExitUpperCaseA(c *UpperCaseAContext)

	// ExitUpperCaseZ is called when exiting the upperCaseZ production.
	ExitUpperCaseZ(c *UpperCaseZContext)

	// ExitLetterRange is called when exiting the letterRange production.
	ExitLetterRange(c *LetterRangeContext)

	// ExitFirstLetter is called when exiting the firstLetter production.
	ExitFirstLetter(c *FirstLetterContext)

	// ExitLastLetter is called when exiting the lastLetter production.
	ExitLastLetter(c *LastLetterContext)

	// ExitDefType is called when exiting the defType production.
	ExitDefType(c *DefTypeContext)

	// ExitCommonModuleDeclarationElement is called when exiting the commonModuleDeclarationElement production.
	ExitCommonModuleDeclarationElement(c *CommonModuleDeclarationElementContext)

	// ExitVariableDeclaration is called when exiting the variableDeclaration production.
	ExitVariableDeclaration(c *VariableDeclarationContext)

	// ExitVariableHelpAttribute is called when exiting the variableHelpAttribute production.
	ExitVariableHelpAttribute(c *VariableHelpAttributeContext)

	// ExitVariableModifier is called when exiting the variableModifier production.
	ExitVariableModifier(c *VariableModifierContext)

	// ExitVariableSharedModifier is called when exiting the variableSharedModifier production.
	ExitVariableSharedModifier(c *VariableSharedModifierContext)

	// ExitModuleVariableDeclarationList is called when exiting the moduleVariableDeclarationList production.
	ExitModuleVariableDeclarationList(c *ModuleVariableDeclarationListContext)

	// ExitVariableDeclarationList is called when exiting the variableDeclarationList production.
	ExitVariableDeclarationList(c *VariableDeclarationListContext)

	// ExitVariableDcl is called when exiting the variableDcl production.
	ExitVariableDcl(c *VariableDclContext)

	// ExitTypedVariableDcl is called when exiting the typedVariableDcl production.
	ExitTypedVariableDcl(c *TypedVariableDclContext)

	// ExitUntypedVariableDcl is called when exiting the untypedVariableDcl production.
	ExitUntypedVariableDcl(c *UntypedVariableDclContext)

	// ExitArrayClause is called when exiting the arrayClause production.
	ExitArrayClause(c *ArrayClauseContext)

	// ExitAsClause is called when exiting the asClause production.
	ExitAsClause(c *AsClauseContext)

	// ExitWitheventsVariableDcl is called when exiting the witheventsVariableDcl production.
	ExitWitheventsVariableDcl(c *WitheventsVariableDclContext)

	// ExitClassTypeName is called when exiting the classTypeName production.
	ExitClassTypeName(c *ClassTypeNameContext)

	// ExitArrayDim is called when exiting the arrayDim production.
	ExitArrayDim(c *ArrayDimContext)

	// ExitBoundsList is called when exiting the boundsList production.
	ExitBoundsList(c *BoundsListContext)

	// ExitDimSpec is called when exiting the dimSpec production.
	ExitDimSpec(c *DimSpecContext)

	// ExitLowerBound is called when exiting the lowerBound production.
	ExitLowerBound(c *LowerBoundContext)

	// ExitUpperBound is called when exiting the upperBound production.
	ExitUpperBound(c *UpperBoundContext)

	// ExitAsAutoObject is called when exiting the asAutoObject production.
	ExitAsAutoObject(c *AsAutoObjectContext)

	// ExitAsType is called when exiting the asType production.
	ExitAsType(c *AsTypeContext)

	// ExitTypeSpec is called when exiting the typeSpec production.
	ExitTypeSpec(c *TypeSpecContext)

	// ExitFixedLengthStringSpec is called when exiting the fixedLengthStringSpec production.
	ExitFixedLengthStringSpec(c *FixedLengthStringSpecContext)

	// ExitStringLength is called when exiting the stringLength production.
	ExitStringLength(c *StringLengthContext)

	// ExitConstantName is called when exiting the constantName production.
	ExitConstantName(c *ConstantNameContext)

	// ExitConstDeclaration is called when exiting the constDeclaration production.
	ExitConstDeclaration(c *ConstDeclarationContext)

	// ExitConstItemList is called when exiting the constItemList production.
	ExitConstItemList(c *ConstItemListContext)

	// ExitConstItem is called when exiting the constItem production.
	ExitConstItem(c *ConstItemContext)

	// ExitConstAsClause is called when exiting the constAsClause production.
	ExitConstAsClause(c *ConstAsClauseContext)

	// ExitPublicTypeDeclaration is called when exiting the publicTypeDeclaration production.
	ExitPublicTypeDeclaration(c *PublicTypeDeclarationContext)

	// ExitPrivateTypeDeclaration is called when exiting the privateTypeDeclaration production.
	ExitPrivateTypeDeclaration(c *PrivateTypeDeclarationContext)

	// ExitUdtDeclaration is called when exiting the udtDeclaration production.
	ExitUdtDeclaration(c *UdtDeclarationContext)

	// ExitUdtMemberList is called when exiting the udtMemberList production.
	ExitUdtMemberList(c *UdtMemberListContext)

	// ExitUdtElement is called when exiting the udtElement production.
	ExitUdtElement(c *UdtElementContext)

	// ExitUdtMember is called when exiting the udtMember production.
	ExitUdtMember(c *UdtMemberContext)

	// ExitUntypedNameMemberDcl is called when exiting the untypedNameMemberDcl production.
	ExitUntypedNameMemberDcl(c *UntypedNameMemberDclContext)

	// ExitReservedNameMemberDcl is called when exiting the reservedNameMemberDcl production.
	ExitReservedNameMemberDcl(c *ReservedNameMemberDclContext)

	// ExitOptionalArrayClause is called when exiting the optionalArrayClause production.
	ExitOptionalArrayClause(c *OptionalArrayClauseContext)

	// ExitReservedMemberName is called when exiting the reservedMemberName production.
	ExitReservedMemberName(c *ReservedMemberNameContext)

	// ExitGlobalEnumDeclaration is called when exiting the globalEnumDeclaration production.
	ExitGlobalEnumDeclaration(c *GlobalEnumDeclarationContext)

	// ExitPublicEnumDeclaration is called when exiting the publicEnumDeclaration production.
	ExitPublicEnumDeclaration(c *PublicEnumDeclarationContext)

	// ExitPrivateEnumDeclaration is called when exiting the privateEnumDeclaration production.
	ExitPrivateEnumDeclaration(c *PrivateEnumDeclarationContext)

	// ExitEnumDeclaration is called when exiting the enumDeclaration production.
	ExitEnumDeclaration(c *EnumDeclarationContext)

	// ExitEnumLongptrDeclaration is called when exiting the enumLongptrDeclaration production.
	ExitEnumLongptrDeclaration(c *EnumLongptrDeclarationContext)

	// ExitEnumMemberList is called when exiting the enumMemberList production.
	ExitEnumMemberList(c *EnumMemberListContext)

	// ExitEnumElement is called when exiting the enumElement production.
	ExitEnumElement(c *EnumElementContext)

	// ExitEnumMember is called when exiting the enumMember production.
	ExitEnumMember(c *EnumMemberContext)

	// ExitPublicExternalProcedureDeclaration is called when exiting the publicExternalProcedureDeclaration production.
	ExitPublicExternalProcedureDeclaration(c *PublicExternalProcedureDeclarationContext)

	// ExitPrivateExternalProcedureDeclaration is called when exiting the privateExternalProcedureDeclaration production.
	ExitPrivateExternalProcedureDeclaration(c *PrivateExternalProcedureDeclarationContext)

	// ExitExternalProcDcl is called when exiting the externalProcDcl production.
	ExitExternalProcDcl(c *ExternalProcDclContext)

	// ExitExternalSub is called when exiting the externalSub production.
	ExitExternalSub(c *ExternalSubContext)

	// ExitExternalFunction is called when exiting the externalFunction production.
	ExitExternalFunction(c *ExternalFunctionContext)

	// ExitLibInfo is called when exiting the libInfo production.
	ExitLibInfo(c *LibInfoContext)

	// ExitLibClause is called when exiting the libClause production.
	ExitLibClause(c *LibClauseContext)

	// ExitAliasClause is called when exiting the aliasClause production.
	ExitAliasClause(c *AliasClauseContext)

	// ExitImplementsDirective is called when exiting the implementsDirective production.
	ExitImplementsDirective(c *ImplementsDirectiveContext)

	// ExitEventDeclaration is called when exiting the eventDeclaration production.
	ExitEventDeclaration(c *EventDeclarationContext)

	// ExitEventParameterList is called when exiting the eventParameterList production.
	ExitEventParameterList(c *EventParameterListContext)

	// ExitProceduralModuleCode is called when exiting the proceduralModuleCode production.
	ExitProceduralModuleCode(c *ProceduralModuleCodeContext)

	// ExitClassModuleCode is called when exiting the classModuleCode production.
	ExitClassModuleCode(c *ClassModuleCodeContext)

	// ExitProceduralModuleCodeElement is called when exiting the proceduralModuleCodeElement production.
	ExitProceduralModuleCodeElement(c *ProceduralModuleCodeElementContext)

	// ExitClassModuleCodeElement is called when exiting the classModuleCodeElement production.
	ExitClassModuleCodeElement(c *ClassModuleCodeElementContext)

	// ExitCommonModuleCodeElement is called when exiting the commonModuleCodeElement production.
	ExitCommonModuleCodeElement(c *CommonModuleCodeElementContext)

	// ExitProcedureDeclaration is called when exiting the procedureDeclaration production.
	ExitProcedureDeclaration(c *ProcedureDeclarationContext)

	// ExitSubroutineDeclaration is called when exiting the subroutineDeclaration production.
	ExitSubroutineDeclaration(c *SubroutineDeclarationContext)

	// ExitFunctionDeclaration is called when exiting the functionDeclaration production.
	ExitFunctionDeclaration(c *FunctionDeclarationContext)

	// ExitPropertyGetDeclaration is called when exiting the propertyGetDeclaration production.
	ExitPropertyGetDeclaration(c *PropertyGetDeclarationContext)

	// ExitPropertySetDeclaration is called when exiting the propertySetDeclaration production.
	ExitPropertySetDeclaration(c *PropertySetDeclarationContext)

	// ExitEndLabel is called when exiting the endLabel production.
	ExitEndLabel(c *EndLabelContext)

	// ExitProcedureTail is called when exiting the procedureTail production.
	ExitProcedureTail(c *ProcedureTailContext)

	// ExitProcedureScope is called when exiting the procedureScope production.
	ExitProcedureScope(c *ProcedureScopeContext)

	// ExitInitialStatic is called when exiting the initialStatic production.
	ExitInitialStatic(c *InitialStaticContext)

	// ExitTrailingStatic is called when exiting the trailingStatic production.
	ExitTrailingStatic(c *TrailingStaticContext)

	// ExitSubroutineName is called when exiting the subroutineName production.
	ExitSubroutineName(c *SubroutineNameContext)

	// ExitFunctionName is called when exiting the functionName production.
	ExitFunctionName(c *FunctionNameContext)

	// ExitPrefixedName is called when exiting the prefixedName production.
	ExitPrefixedName(c *PrefixedNameContext)

	// ExitFunctionType is called when exiting the functionType production.
	ExitFunctionType(c *FunctionTypeContext)

	// ExitArrayDesignator is called when exiting the arrayDesignator production.
	ExitArrayDesignator(c *ArrayDesignatorContext)

	// ExitProcedureParameters is called when exiting the procedureParameters production.
	ExitProcedureParameters(c *ProcedureParametersContext)

	// ExitPropertyParameters is called when exiting the propertyParameters production.
	ExitPropertyParameters(c *PropertyParametersContext)

	// ExitPropertyParameterList is called when exiting the propertyParameterList production.
	ExitPropertyParameterList(c *PropertyParameterListContext)

	// ExitParameterList is called when exiting the parameterList production.
	ExitParameterList(c *ParameterListContext)

	// ExitAnyParam is called when exiting the anyParam production.
	ExitAnyParam(c *AnyParamContext)

	// ExitPositionalParameters is called when exiting the positionalParameters production.
	ExitPositionalParameters(c *PositionalParametersContext)

	// ExitValueParam is called when exiting the valueParam production.
	ExitValueParam(c *ValueParamContext)

	// ExitPositionalParam is called when exiting the positionalParam production.
	ExitPositionalParam(c *PositionalParamContext)

	// ExitOptionalParam is called when exiting the optionalParam production.
	ExitOptionalParam(c *OptionalParamContext)

	// ExitParamArray is called when exiting the paramArray production.
	ExitParamArray(c *ParamArrayContext)

	// ExitParamDcl is called when exiting the paramDcl production.
	ExitParamDcl(c *ParamDclContext)

	// ExitUntypedNameParamDcl is called when exiting the untypedNameParamDcl production.
	ExitUntypedNameParamDcl(c *UntypedNameParamDclContext)

	// ExitTypedNameParamDcl is called when exiting the typedNameParamDcl production.
	ExitTypedNameParamDcl(c *TypedNameParamDclContext)

	// ExitOptionalPrefix is called when exiting the optionalPrefix production.
	ExitOptionalPrefix(c *OptionalPrefixContext)

	// ExitParameterMechanism is called when exiting the parameterMechanism production.
	ExitParameterMechanism(c *ParameterMechanismContext)

	// ExitParameterType is called when exiting the parameterType production.
	ExitParameterType(c *ParameterTypeContext)

	// ExitDefaultValue is called when exiting the defaultValue production.
	ExitDefaultValue(c *DefaultValueContext)

	// ExitEventHandlerName is called when exiting the eventHandlerName production.
	ExitEventHandlerName(c *EventHandlerNameContext)

	// ExitImplementedName is called when exiting the implementedName production.
	ExitImplementedName(c *ImplementedNameContext)

	// ExitLifecycleHandlerName is called when exiting the lifecycleHandlerName production.
	ExitLifecycleHandlerName(c *LifecycleHandlerNameContext)

	// ExitProcedureBody is called when exiting the procedureBody production.
	ExitProcedureBody(c *ProcedureBodyContext)

	// ExitStatementBlock is called when exiting the statementBlock production.
	ExitStatementBlock(c *StatementBlockContext)

	// ExitBlockStatement is called when exiting the blockStatement production.
	ExitBlockStatement(c *BlockStatementContext)

	// ExitStatement is called when exiting the statement production.
	ExitStatement(c *StatementContext)

	// ExitStatementLabelDefinition is called when exiting the statementLabelDefinition production.
	ExitStatementLabelDefinition(c *StatementLabelDefinitionContext)

	// ExitStatementLabel is called when exiting the statementLabel production.
	ExitStatementLabel(c *StatementLabelContext)

	// ExitStatementLabelList is called when exiting the statementLabelList production.
	ExitStatementLabelList(c *StatementLabelListContext)

	// ExitIdentifierStatementLabel is called when exiting the identifierStatementLabel production.
	ExitIdentifierStatementLabel(c *IdentifierStatementLabelContext)

	// ExitResetNumberLabel is called when exiting the resetNumberLabel production.
	ExitResetNumberLabel(c *ResetNumberLabelContext)

	// ExitLineNumberLabel is called when exiting the lineNumberLabel production.
	ExitLineNumberLabel(c *LineNumberLabelContext)

	// ExitRemStatement is called when exiting the remStatement production.
	ExitRemStatement(c *RemStatementContext)

	// ExitControlStatement is called when exiting the controlStatement production.
	ExitControlStatement(c *ControlStatementContext)

	// ExitControlStatementExceptMultilineIf is called when exiting the controlStatementExceptMultilineIf production.
	ExitControlStatementExceptMultilineIf(c *ControlStatementExceptMultilineIfContext)

	// ExitCallStatement is called when exiting the callStatement production.
	ExitCallStatement(c *CallStatementContext)

	// ExitWhileStatement is called when exiting the whileStatement production.
	ExitWhileStatement(c *WhileStatementContext)

	// ExitForStatement is called when exiting the forStatement production.
	ExitForStatement(c *ForStatementContext)

	// ExitForTail is called when exiting the forTail production.
	ExitForTail(c *ForTailContext)

	// ExitForClause is called when exiting the forClause production.
	ExitForClause(c *ForClauseContext)

	// ExitStartValue is called when exiting the startValue production.
	ExitStartValue(c *StartValueContext)

	// ExitEndValue is called when exiting the endValue production.
	ExitEndValue(c *EndValueContext)

	// ExitStepClause is called when exiting the stepClause production.
	ExitStepClause(c *StepClauseContext)

	// ExitStepIncrement is called when exiting the stepIncrement production.
	ExitStepIncrement(c *StepIncrementContext)

	// ExitForEachStatement is called when exiting the forEachStatement production.
	ExitForEachStatement(c *ForEachStatementContext)

	// ExitForEachClause is called when exiting the forEachClause production.
	ExitForEachClause(c *ForEachClauseContext)

	// ExitCollection is called when exiting the collection production.
	ExitCollection(c *CollectionContext)

	// ExitExitForStatement is called when exiting the exitForStatement production.
	ExitExitForStatement(c *ExitForStatementContext)

	// ExitDoStatement is called when exiting the doStatement production.
	ExitDoStatement(c *DoStatementContext)

	// ExitConditionClause is called when exiting the conditionClause production.
	ExitConditionClause(c *ConditionClauseContext)

	// ExitWhileClause is called when exiting the whileClause production.
	ExitWhileClause(c *WhileClauseContext)

	// ExitUntilClause is called when exiting the untilClause production.
	ExitUntilClause(c *UntilClauseContext)

	// ExitExitDoStatement is called when exiting the exitDoStatement production.
	ExitExitDoStatement(c *ExitDoStatementContext)

	// ExitIfStatement is called when exiting the ifStatement production.
	ExitIfStatement(c *IfStatementContext)

	// ExitElseIfBlock is called when exiting the elseIfBlock production.
	ExitElseIfBlock(c *ElseIfBlockContext)

	// ExitElseBlock is called when exiting the elseBlock production.
	ExitElseBlock(c *ElseBlockContext)

	// ExitSingleLineIfStatement is called when exiting the singleLineIfStatement production.
	ExitSingleLineIfStatement(c *SingleLineIfStatementContext)

	// ExitIfWithNonEmptyThen is called when exiting the ifWithNonEmptyThen production.
	ExitIfWithNonEmptyThen(c *IfWithNonEmptyThenContext)

	// ExitIfWithEmptyThen is called when exiting the ifWithEmptyThen production.
	ExitIfWithEmptyThen(c *IfWithEmptyThenContext)

	// ExitSingleLineElseClause is called when exiting the singleLineElseClause production.
	ExitSingleLineElseClause(c *SingleLineElseClauseContext)

	// ExitListOrLabel is called when exiting the listOrLabel production.
	ExitListOrLabel(c *ListOrLabelContext)

	// ExitSameLineStatement is called when exiting the sameLineStatement production.
	ExitSameLineStatement(c *SameLineStatementContext)

	// ExitSelectCaseStatement is called when exiting the selectCaseStatement production.
	ExitSelectCaseStatement(c *SelectCaseStatementContext)

	// ExitCaseClause is called when exiting the caseClause production.
	ExitCaseClause(c *CaseClauseContext)

	// ExitCaseElseClause is called when exiting the caseElseClause production.
	ExitCaseElseClause(c *CaseElseClauseContext)

	// ExitRangeClause is called when exiting the rangeClause production.
	ExitRangeClause(c *RangeClauseContext)

	// ExitSelectExpression is called when exiting the selectExpression production.
	ExitSelectExpression(c *SelectExpressionContext)

	// ExitComparisonOperator is called when exiting the comparisonOperator production.
	ExitComparisonOperator(c *ComparisonOperatorContext)

	// ExitStopStatement is called when exiting the stopStatement production.
	ExitStopStatement(c *StopStatementContext)

	// ExitGotoStatement is called when exiting the gotoStatement production.
	ExitGotoStatement(c *GotoStatementContext)

	// ExitOnGotoStatement is called when exiting the onGotoStatement production.
	ExitOnGotoStatement(c *OnGotoStatementContext)

	// ExitGosubStatement is called when exiting the gosubStatement production.
	ExitGosubStatement(c *GosubStatementContext)

	// ExitReturnStatement is called when exiting the returnStatement production.
	ExitReturnStatement(c *ReturnStatementContext)

	// ExitOnGosubStatement is called when exiting the onGosubStatement production.
	ExitOnGosubStatement(c *OnGosubStatementContext)

	// ExitExitSubStatement is called when exiting the exitSubStatement production.
	ExitExitSubStatement(c *ExitSubStatementContext)

	// ExitExitFunctionStatement is called when exiting the exitFunctionStatement production.
	ExitExitFunctionStatement(c *ExitFunctionStatementContext)

	// ExitExitPropertyStatement is called when exiting the exitPropertyStatement production.
	ExitExitPropertyStatement(c *ExitPropertyStatementContext)

	// ExitRaiseeventStatement is called when exiting the raiseeventStatement production.
	ExitRaiseeventStatement(c *RaiseeventStatementContext)

	// ExitEventArgumentList is called when exiting the eventArgumentList production.
	ExitEventArgumentList(c *EventArgumentListContext)

	// ExitEventArgument is called when exiting the eventArgument production.
	ExitEventArgument(c *EventArgumentContext)

	// ExitWithStatement is called when exiting the withStatement production.
	ExitWithStatement(c *WithStatementContext)

	// ExitEndStatement is called when exiting the endStatement production.
	ExitEndStatement(c *EndStatementContext)

	// ExitDataManipulationStatement is called when exiting the dataManipulationStatement production.
	ExitDataManipulationStatement(c *DataManipulationStatementContext)

	// ExitStaticVariableDeclaration is called when exiting the staticVariableDeclaration production.
	ExitStaticVariableDeclaration(c *StaticVariableDeclarationContext)

	// ExitRedimStatement is called when exiting the redimStatement production.
	ExitRedimStatement(c *RedimStatementContext)

	// ExitRedimDeclarationList is called when exiting the redimDeclarationList production.
	ExitRedimDeclarationList(c *RedimDeclarationListContext)

	// ExitRedimVariableDcl is called when exiting the redimVariableDcl production.
	ExitRedimVariableDcl(c *RedimVariableDclContext)

	// ExitRedimTypedVariableDcl is called when exiting the redimTypedVariableDcl production.
	ExitRedimTypedVariableDcl(c *RedimTypedVariableDclContext)

	// ExitRedimUntypedDcl is called when exiting the redimUntypedDcl production.
	ExitRedimUntypedDcl(c *RedimUntypedDclContext)

	// ExitWithExpressionDcl is called when exiting the withExpressionDcl production.
	ExitWithExpressionDcl(c *WithExpressionDclContext)

	// ExitMemberAccessExpressionDcl is called when exiting the memberAccessExpressionDcl production.
	ExitMemberAccessExpressionDcl(c *MemberAccessExpressionDclContext)

	// ExitDynamicArrayDim is called when exiting the dynamicArrayDim production.
	ExitDynamicArrayDim(c *DynamicArrayDimContext)

	// ExitDynamicBoundsList is called when exiting the dynamicBoundsList production.
	ExitDynamicBoundsList(c *DynamicBoundsListContext)

	// ExitDynamicDimSpec is called when exiting the dynamicDimSpec production.
	ExitDynamicDimSpec(c *DynamicDimSpecContext)

	// ExitDynamicLowerBound is called when exiting the dynamicLowerBound production.
	ExitDynamicLowerBound(c *DynamicLowerBoundContext)

	// ExitDynamicUpperBound is called when exiting the dynamicUpperBound production.
	ExitDynamicUpperBound(c *DynamicUpperBoundContext)

	// ExitDynamicArrayClause is called when exiting the dynamicArrayClause production.
	ExitDynamicArrayClause(c *DynamicArrayClauseContext)

	// ExitEraseStatement is called when exiting the eraseStatement production.
	ExitEraseStatement(c *EraseStatementContext)

	// ExitEraseList is called when exiting the eraseList production.
	ExitEraseList(c *EraseListContext)

	// ExitEraseElement is called when exiting the eraseElement production.
	ExitEraseElement(c *EraseElementContext)

	// ExitMidStatement is called when exiting the midStatement production.
	ExitMidStatement(c *MidStatementContext)

	// ExitModeSpecifier is called when exiting the modeSpecifier production.
	ExitModeSpecifier(c *ModeSpecifierContext)

	// ExitStringArgument is called when exiting the stringArgument production.
	ExitStringArgument(c *StringArgumentContext)

	// ExitStartMid is called when exiting the startMid production.
	ExitStartMid(c *StartMidContext)

	// ExitLength is called when exiting the length production.
	ExitLength(c *LengthContext)

	// ExitLsetStatement is called when exiting the lsetStatement production.
	ExitLsetStatement(c *LsetStatementContext)

	// ExitRsetStatement is called when exiting the rsetStatement production.
	ExitRsetStatement(c *RsetStatementContext)

	// ExitLetStatement is called when exiting the letStatement production.
	ExitLetStatement(c *LetStatementContext)

	// ExitSetStatement is called when exiting the setStatement production.
	ExitSetStatement(c *SetStatementContext)

	// ExitErrorHandlingStatement is called when exiting the errorHandlingStatement production.
	ExitErrorHandlingStatement(c *ErrorHandlingStatementContext)

	// ExitOnErrorStatement is called when exiting the onErrorStatement production.
	ExitOnErrorStatement(c *OnErrorStatementContext)

	// ExitErrorBehavior is called when exiting the errorBehavior production.
	ExitErrorBehavior(c *ErrorBehaviorContext)

	// ExitResumeStatement is called when exiting the resumeStatement production.
	ExitResumeStatement(c *ResumeStatementContext)

	// ExitErrorStatement is called when exiting the errorStatement production.
	ExitErrorStatement(c *ErrorStatementContext)

	// ExitErrorNumber is called when exiting the errorNumber production.
	ExitErrorNumber(c *ErrorNumberContext)

	// ExitFileStatement is called when exiting the fileStatement production.
	ExitFileStatement(c *FileStatementContext)

	// ExitOpenStatement is called when exiting the openStatement production.
	ExitOpenStatement(c *OpenStatementContext)

	// ExitPathName is called when exiting the pathName production.
	ExitPathName(c *PathNameContext)

	// ExitModeClause is called when exiting the modeClause production.
	ExitModeClause(c *ModeClauseContext)

	// ExitModeOpt is called when exiting the modeOpt production.
	ExitModeOpt(c *ModeOptContext)

	// ExitAccessClause is called when exiting the accessClause production.
	ExitAccessClause(c *AccessClauseContext)

	// ExitAccess is called when exiting the access production.
	ExitAccess(c *AccessContext)

	// ExitLock is called when exiting the lock production.
	ExitLock(c *LockContext)

	// ExitLenClause is called when exiting the lenClause production.
	ExitLenClause(c *LenClauseContext)

	// ExitRecLength is called when exiting the recLength production.
	ExitRecLength(c *RecLengthContext)

	// ExitFileNumber is called when exiting the fileNumber production.
	ExitFileNumber(c *FileNumberContext)

	// ExitMarkedFileNumber is called when exiting the markedFileNumber production.
	ExitMarkedFileNumber(c *MarkedFileNumberContext)

	// ExitUnmarkedFileNumber is called when exiting the unmarkedFileNumber production.
	ExitUnmarkedFileNumber(c *UnmarkedFileNumberContext)

	// ExitCloseStatement is called when exiting the closeStatement production.
	ExitCloseStatement(c *CloseStatementContext)

	// ExitFileNumberList is called when exiting the fileNumberList production.
	ExitFileNumberList(c *FileNumberListContext)

	// ExitSeekStatement is called when exiting the seekStatement production.
	ExitSeekStatement(c *SeekStatementContext)

	// ExitPosition is called when exiting the position production.
	ExitPosition(c *PositionContext)

	// ExitLockStatement is called when exiting the lockStatement production.
	ExitLockStatement(c *LockStatementContext)

	// ExitRecordRange is called when exiting the recordRange production.
	ExitRecordRange(c *RecordRangeContext)

	// ExitStartRecordNumber is called when exiting the startRecordNumber production.
	ExitStartRecordNumber(c *StartRecordNumberContext)

	// ExitEndRecordNumber is called when exiting the endRecordNumber production.
	ExitEndRecordNumber(c *EndRecordNumberContext)

	// ExitUnlockStatement is called when exiting the unlockStatement production.
	ExitUnlockStatement(c *UnlockStatementContext)

	// ExitLineInputStatement is called when exiting the lineInputStatement production.
	ExitLineInputStatement(c *LineInputStatementContext)

	// ExitVariableName is called when exiting the variableName production.
	ExitVariableName(c *VariableNameContext)

	// ExitWidthStatement is called when exiting the widthStatement production.
	ExitWidthStatement(c *WidthStatementContext)

	// ExitLineWidth is called when exiting the lineWidth production.
	ExitLineWidth(c *LineWidthContext)

	// ExitPrintStatement is called when exiting the printStatement production.
	ExitPrintStatement(c *PrintStatementContext)

	// ExitOutputList is called when exiting the outputList production.
	ExitOutputList(c *OutputListContext)

	// ExitOutputItem is called when exiting the outputItem production.
	ExitOutputItem(c *OutputItemContext)

	// ExitOutputClause is called when exiting the outputClause production.
	ExitOutputClause(c *OutputClauseContext)

	// ExitCharPosition is called when exiting the charPosition production.
	ExitCharPosition(c *CharPositionContext)

	// ExitOutputExpression is called when exiting the outputExpression production.
	ExitOutputExpression(c *OutputExpressionContext)

	// ExitSpcClause is called when exiting the spcClause production.
	ExitSpcClause(c *SpcClauseContext)

	// ExitSpcNumber is called when exiting the spcNumber production.
	ExitSpcNumber(c *SpcNumberContext)

	// ExitTabClause is called when exiting the tabClause production.
	ExitTabClause(c *TabClauseContext)

	// ExitTabNumber is called when exiting the tabNumber production.
	ExitTabNumber(c *TabNumberContext)

	// ExitWriteStatement is called when exiting the writeStatement production.
	ExitWriteStatement(c *WriteStatementContext)

	// ExitInputStatement is called when exiting the inputStatement production.
	ExitInputStatement(c *InputStatementContext)

	// ExitInputList is called when exiting the inputList production.
	ExitInputList(c *InputListContext)

	// ExitInputVariable is called when exiting the inputVariable production.
	ExitInputVariable(c *InputVariableContext)

	// ExitPutStatement is called when exiting the putStatement production.
	ExitPutStatement(c *PutStatementContext)

	// ExitRecordNumber is called when exiting the recordNumber production.
	ExitRecordNumber(c *RecordNumberContext)

	// ExitData is called when exiting the data production.
	ExitData(c *DataContext)

	// ExitGetStatement is called when exiting the getStatement production.
	ExitGetStatement(c *GetStatementContext)

	// ExitVariable is called when exiting the variable production.
	ExitVariable(c *VariableContext)

	// ExitAttributeStatement is called when exiting the attributeStatement production.
	ExitAttributeStatement(c *AttributeStatementContext)

	// ExitAttributeDescName is called when exiting the attributeDescName production.
	ExitAttributeDescName(c *AttributeDescNameContext)

	// ExitAttributeUsrName is called when exiting the attributeUsrName production.
	ExitAttributeUsrName(c *AttributeUsrNameContext)

	// ExitDebugStatement is called when exiting the debugStatement production.
	ExitDebugStatement(c *DebugStatementContext)

	// ExitDebugArgs is called when exiting the debugArgs production.
	ExitDebugArgs(c *DebugArgsContext)

	// ExitDebugSep is called when exiting the debugSep production.
	ExitDebugSep(c *DebugSepContext)

	// ExitExpression is called when exiting the expression production.
	ExitExpression(c *ExpressionContext)

	// ExitLExpression is called when exiting the lExpression production.
	ExitLExpression(c *LExpressionContext)

	// ExitLiteralExpression is called when exiting the literalExpression production.
	ExitLiteralExpression(c *LiteralExpressionContext)

	// ExitParenthesizedExpression is called when exiting the parenthesizedExpression production.
	ExitParenthesizedExpression(c *ParenthesizedExpressionContext)

	// ExitTypeofIsExpression is called when exiting the typeofIsExpression production.
	ExitTypeofIsExpression(c *TypeofIsExpressionContext)

	// ExitNewExpress is called when exiting the newExpress production.
	ExitNewExpress(c *NewExpressContext)

	// ExitNotOperatorExpression is called when exiting the notOperatorExpression production.
	ExitNotOperatorExpression(c *NotOperatorExpressionContext)

	// ExitUnaryMinusExpression is called when exiting the unaryMinusExpression production.
	ExitUnaryMinusExpression(c *UnaryMinusExpressionContext)

	// ExitSimpleNameExpression is called when exiting the simpleNameExpression production.
	ExitSimpleNameExpression(c *SimpleNameExpressionContext)

	// ExitInstanceExpression is called when exiting the instanceExpression production.
	ExitInstanceExpression(c *InstanceExpressionContext)

	// ExitMemberAccessExpression is called when exiting the memberAccessExpression production.
	ExitMemberAccessExpression(c *MemberAccessExpressionContext)

	// ExitIndexExpression is called when exiting the indexExpression production.
	ExitIndexExpression(c *IndexExpressionContext)

	// ExitArgumentList is called when exiting the argumentList production.
	ExitArgumentList(c *ArgumentListContext)

	// ExitPositionalOrNamedArgumentList is called when exiting the positionalOrNamedArgumentList production.
	ExitPositionalOrNamedArgumentList(c *PositionalOrNamedArgumentListContext)

	// ExitNamedArgumentList is called when exiting the namedArgumentList production.
	ExitNamedArgumentList(c *NamedArgumentListContext)

	// ExitNamedArgument is called when exiting the namedArgument production.
	ExitNamedArgument(c *NamedArgumentContext)

	// ExitArgumentExpression is called when exiting the argumentExpression production.
	ExitArgumentExpression(c *ArgumentExpressionContext)

	// ExitDictionaryAccessExpression is called when exiting the dictionaryAccessExpression production.
	ExitDictionaryAccessExpression(c *DictionaryAccessExpressionContext)

	// ExitWithExpression is called when exiting the withExpression production.
	ExitWithExpression(c *WithExpressionContext)

	// ExitWithMemberAccessExpression is called when exiting the withMemberAccessExpression production.
	ExitWithMemberAccessExpression(c *WithMemberAccessExpressionContext)

	// ExitWithDictionaryAccessExpression is called when exiting the withDictionaryAccessExpression production.
	ExitWithDictionaryAccessExpression(c *WithDictionaryAccessExpressionContext)

	// ExitConstantExpression is called when exiting the constantExpression production.
	ExitConstantExpression(c *ConstantExpressionContext)

	// ExitCcExpression is called when exiting the ccExpression production.
	ExitCcExpression(c *CcExpressionContext)

	// ExitBooleanExpression is called when exiting the booleanExpression production.
	ExitBooleanExpression(c *BooleanExpressionContext)

	// ExitIntegerExpression is called when exiting the integerExpression production.
	ExitIntegerExpression(c *IntegerExpressionContext)

	// ExitVariableExpression is called when exiting the variableExpression production.
	ExitVariableExpression(c *VariableExpressionContext)

	// ExitBoundVariableExpression is called when exiting the boundVariableExpression production.
	ExitBoundVariableExpression(c *BoundVariableExpressionContext)

	// ExitTypeExpression is called when exiting the typeExpression production.
	ExitTypeExpression(c *TypeExpressionContext)

	// ExitDefinedTypeExpression is called when exiting the definedTypeExpression production.
	ExitDefinedTypeExpression(c *DefinedTypeExpressionContext)

	// ExitAddressofExpression is called when exiting the addressofExpression production.
	ExitAddressofExpression(c *AddressofExpressionContext)

	// ExitProcedurePointerExpression is called when exiting the procedurePointerExpression production.
	ExitProcedurePointerExpression(c *ProcedurePointerExpressionContext)

	// ExitWsc is called when exiting the wsc production.
	ExitWsc(c *WscContext)

	// ExitEndOfLine is called when exiting the endOfLine production.
	ExitEndOfLine(c *EndOfLineContext)

	// ExitUnexpectedEndOfLine is called when exiting the unexpectedEndOfLine production.
	ExitUnexpectedEndOfLine(c *UnexpectedEndOfLineContext)

	// ExitWscu is called when exiting the wscu production.
	ExitWscu(c *WscuContext)

	// ExitEndOfLineNoWs is called when exiting the endOfLineNoWs production.
	ExitEndOfLineNoWs(c *EndOfLineNoWsContext)

	// ExitEndOfStatement is called when exiting the endOfStatement production.
	ExitEndOfStatement(c *EndOfStatementContext)

	// ExitEndOfStatementNoWs is called when exiting the endOfStatementNoWs production.
	ExitEndOfStatementNoWs(c *EndOfStatementNoWsContext)

	// ExitCommentBody is called when exiting the commentBody production.
	ExitCommentBody(c *CommentBodyContext)

	// ExitReservedIdentifier is called when exiting the reservedIdentifier production.
	ExitReservedIdentifier(c *ReservedIdentifierContext)

	// ExitAmbiguousIdentifier is called when exiting the ambiguousIdentifier production.
	ExitAmbiguousIdentifier(c *AmbiguousIdentifierContext)

	// ExitStatementKeyword is called when exiting the statementKeyword production.
	ExitStatementKeyword(c *StatementKeywordContext)

	// ExitRemKeyword is called when exiting the remKeyword production.
	ExitRemKeyword(c *RemKeywordContext)

	// ExitMarkerKeyword is called when exiting the markerKeyword production.
	ExitMarkerKeyword(c *MarkerKeywordContext)

	// ExitOperatorIdentifier is called when exiting the operatorIdentifier production.
	ExitOperatorIdentifier(c *OperatorIdentifierContext)

	// ExitReservedName is called when exiting the reservedName production.
	ExitReservedName(c *ReservedNameContext)

	// ExitSpecialForm is called when exiting the specialForm production.
	ExitSpecialForm(c *SpecialFormContext)

	// ExitReservedTypeIdentifier is called when exiting the reservedTypeIdentifier production.
	ExitReservedTypeIdentifier(c *ReservedTypeIdentifierContext)

	// ExitReservedTypeIdentifierB is called when exiting the reservedTypeIdentifierB production.
	ExitReservedTypeIdentifierB(c *ReservedTypeIdentifierBContext)

	// ExitLiteralIdentifier is called when exiting the literalIdentifier production.
	ExitLiteralIdentifier(c *LiteralIdentifierContext)

	// ExitBooleanLiteralIdentifier is called when exiting the booleanLiteralIdentifier production.
	ExitBooleanLiteralIdentifier(c *BooleanLiteralIdentifierContext)

	// ExitObjectLiteralIdentifier is called when exiting the objectLiteralIdentifier production.
	ExitObjectLiteralIdentifier(c *ObjectLiteralIdentifierContext)

	// ExitVariantLiteralIdentifier is called when exiting the variantLiteralIdentifier production.
	ExitVariantLiteralIdentifier(c *VariantLiteralIdentifierContext)

	// ExitReservedForImplementationUse is called when exiting the reservedForImplementationUse production.
	ExitReservedForImplementationUse(c *ReservedForImplementationUseContext)

	// ExitFutureReserved is called when exiting the futureReserved production.
	ExitFutureReserved(c *FutureReservedContext)

	// ExitBuiltinType is called when exiting the builtinType production.
	ExitBuiltinType(c *BuiltinTypeContext)

	// ExitTypedName is called when exiting the typedName production.
	ExitTypedName(c *TypedNameContext)

	// ExitTypeSuffix is called when exiting the typeSuffix production.
	ExitTypeSuffix(c *TypeSuffixContext)

	// ExitAmbiguousKeyword is called when exiting the ambiguousKeyword production.
	ExitAmbiguousKeyword(c *AmbiguousKeywordContext)

	// ExitAnyOperator is called when exiting the anyOperator production.
	ExitAnyOperator(c *AnyOperatorContext)

	// ExitPowOperator is called when exiting the powOperator production.
	ExitPowOperator(c *PowOperatorContext)

	// ExitDivOperator is called when exiting the divOperator production.
	ExitDivOperator(c *DivOperatorContext)

	// ExitMultOperator is called when exiting the multOperator production.
	ExitMultOperator(c *MultOperatorContext)

	// ExitModOperator is called when exiting the modOperator production.
	ExitModOperator(c *ModOperatorContext)

	// ExitPlusOperator is called when exiting the plusOperator production.
	ExitPlusOperator(c *PlusOperatorContext)

	// ExitMinusOperator is called when exiting the minusOperator production.
	ExitMinusOperator(c *MinusOperatorContext)

	// ExitAmpOperator is called when exiting the ampOperator production.
	ExitAmpOperator(c *AmpOperatorContext)

	// ExitIsOperator is called when exiting the isOperator production.
	ExitIsOperator(c *IsOperatorContext)

	// ExitLikeOperator is called when exiting the likeOperator production.
	ExitLikeOperator(c *LikeOperatorContext)

	// ExitGeqOperator is called when exiting the geqOperator production.
	ExitGeqOperator(c *GeqOperatorContext)

	// ExitLeqOperator is called when exiting the leqOperator production.
	ExitLeqOperator(c *LeqOperatorContext)

	// ExitGtOperator is called when exiting the gtOperator production.
	ExitGtOperator(c *GtOperatorContext)

	// ExitLtOperator is called when exiting the ltOperator production.
	ExitLtOperator(c *LtOperatorContext)

	// ExitNeqOperator is called when exiting the neqOperator production.
	ExitNeqOperator(c *NeqOperatorContext)

	// ExitEqOperator is called when exiting the eqOperator production.
	ExitEqOperator(c *EqOperatorContext)

	// ExitAndOperator is called when exiting the andOperator production.
	ExitAndOperator(c *AndOperatorContext)

	// ExitOrOperator is called when exiting the orOperator production.
	ExitOrOperator(c *OrOperatorContext)

	// ExitXorOperator is called when exiting the xorOperator production.
	ExitXorOperator(c *XorOperatorContext)

	// ExitEqvOperator is called when exiting the eqvOperator production.
	ExitEqvOperator(c *EqvOperatorContext)

	// ExitImpOperator is called when exiting the impOperator production.
	ExitImpOperator(c *ImpOperatorContext)
}
