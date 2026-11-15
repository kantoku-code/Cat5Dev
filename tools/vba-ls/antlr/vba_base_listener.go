// Code generated from vba.g4 by ANTLR 4.13.2. DO NOT EDIT.

package antlr // vba
import "github.com/antlr4-go/antlr/v4"

// BasevbaListener is a complete listener for a parse tree produced by vbaParser.
type BasevbaListener struct{}

var _ vbaListener = &BasevbaListener{}

// VisitTerminal is called when a terminal node is visited.
func (s *BasevbaListener) VisitTerminal(node antlr.TerminalNode) {}

// VisitErrorNode is called when an error node is visited.
func (s *BasevbaListener) VisitErrorNode(node antlr.ErrorNode) {}

// EnterEveryRule is called when any rule is entered.
func (s *BasevbaListener) EnterEveryRule(ctx antlr.ParserRuleContext) {}

// ExitEveryRule is called when any rule is exited.
func (s *BasevbaListener) ExitEveryRule(ctx antlr.ParserRuleContext) {}

// EnterStartRule is called when production startRule is entered.
func (s *BasevbaListener) EnterStartRule(ctx *StartRuleContext) {}

// ExitStartRule is called when production startRule is exited.
func (s *BasevbaListener) ExitStartRule(ctx *StartRuleContext) {}

// EnterModule is called when production module is entered.
func (s *BasevbaListener) EnterModule(ctx *ModuleContext) {}

// ExitModule is called when production module is exited.
func (s *BasevbaListener) ExitModule(ctx *ModuleContext) {}

// EnterClassFileHeader is called when production classFileHeader is entered.
func (s *BasevbaListener) EnterClassFileHeader(ctx *ClassFileHeaderContext) {}

// ExitClassFileHeader is called when production classFileHeader is exited.
func (s *BasevbaListener) ExitClassFileHeader(ctx *ClassFileHeaderContext) {}

// EnterClassVersionIdentification is called when production classVersionIdentification is entered.
func (s *BasevbaListener) EnterClassVersionIdentification(ctx *ClassVersionIdentificationContext) {}

// ExitClassVersionIdentification is called when production classVersionIdentification is exited.
func (s *BasevbaListener) ExitClassVersionIdentification(ctx *ClassVersionIdentificationContext) {}

// EnterClassBeginBlock is called when production classBeginBlock is entered.
func (s *BasevbaListener) EnterClassBeginBlock(ctx *ClassBeginBlockContext) {}

// ExitClassBeginBlock is called when production classBeginBlock is exited.
func (s *BasevbaListener) ExitClassBeginBlock(ctx *ClassBeginBlockContext) {}

// EnterBeginBlockConfigElement is called when production beginBlockConfigElement is entered.
func (s *BasevbaListener) EnterBeginBlockConfigElement(ctx *BeginBlockConfigElementContext) {}

// ExitBeginBlockConfigElement is called when production beginBlockConfigElement is exited.
func (s *BasevbaListener) ExitBeginBlockConfigElement(ctx *BeginBlockConfigElementContext) {}

// EnterFormFileHeader is called when production formFileHeader is entered.
func (s *BasevbaListener) EnterFormFileHeader(ctx *FormFileHeaderContext) {}

// ExitFormFileHeader is called when production formFileHeader is exited.
func (s *BasevbaListener) ExitFormFileHeader(ctx *FormFileHeaderContext) {}

// EnterFormVersionIdentification is called when production formVersionIdentification is entered.
func (s *BasevbaListener) EnterFormVersionIdentification(ctx *FormVersionIdentificationContext) {}

// ExitFormVersionIdentification is called when production formVersionIdentification is exited.
func (s *BasevbaListener) ExitFormVersionIdentification(ctx *FormVersionIdentificationContext) {}

// EnterFormObjectAssign is called when production formObjectAssign is entered.
func (s *BasevbaListener) EnterFormObjectAssign(ctx *FormObjectAssignContext) {}

// ExitFormObjectAssign is called when production formObjectAssign is exited.
func (s *BasevbaListener) ExitFormObjectAssign(ctx *FormObjectAssignContext) {}

// EnterFormBeginBlock is called when production formBeginBlock is entered.
func (s *BasevbaListener) EnterFormBeginBlock(ctx *FormBeginBlockContext) {}

// ExitFormBeginBlock is called when production formBeginBlock is exited.
func (s *BasevbaListener) ExitFormBeginBlock(ctx *FormBeginBlockContext) {}

// EnterBeginPropertyBlock is called when production beginPropertyBlock is entered.
func (s *BasevbaListener) EnterBeginPropertyBlock(ctx *BeginPropertyBlockContext) {}

// ExitBeginPropertyBlock is called when production beginPropertyBlock is exited.
func (s *BasevbaListener) ExitBeginPropertyBlock(ctx *BeginPropertyBlockContext) {}

// EnterProceduralModule is called when production proceduralModule is entered.
func (s *BasevbaListener) EnterProceduralModule(ctx *ProceduralModuleContext) {}

// ExitProceduralModule is called when production proceduralModule is exited.
func (s *BasevbaListener) ExitProceduralModule(ctx *ProceduralModuleContext) {}

// EnterClassModule is called when production classModule is entered.
func (s *BasevbaListener) EnterClassModule(ctx *ClassModuleContext) {}

// ExitClassModule is called when production classModule is exited.
func (s *BasevbaListener) ExitClassModule(ctx *ClassModuleContext) {}

// EnterProceduralModuleHeader is called when production proceduralModuleHeader is entered.
func (s *BasevbaListener) EnterProceduralModuleHeader(ctx *ProceduralModuleHeaderContext) {}

// ExitProceduralModuleHeader is called when production proceduralModuleHeader is exited.
func (s *BasevbaListener) ExitProceduralModuleHeader(ctx *ProceduralModuleHeaderContext) {}

// EnterProceduralModuleAttr is called when production proceduralModuleAttr is entered.
func (s *BasevbaListener) EnterProceduralModuleAttr(ctx *ProceduralModuleAttrContext) {}

// ExitProceduralModuleAttr is called when production proceduralModuleAttr is exited.
func (s *BasevbaListener) ExitProceduralModuleAttr(ctx *ProceduralModuleAttrContext) {}

// EnterIgnoredProceduralAttr is called when production ignoredProceduralAttr is entered.
func (s *BasevbaListener) EnterIgnoredProceduralAttr(ctx *IgnoredProceduralAttrContext) {}

// ExitIgnoredProceduralAttr is called when production ignoredProceduralAttr is exited.
func (s *BasevbaListener) ExitIgnoredProceduralAttr(ctx *IgnoredProceduralAttrContext) {}

// EnterClassModuleHeader is called when production classModuleHeader is entered.
func (s *BasevbaListener) EnterClassModuleHeader(ctx *ClassModuleHeaderContext) {}

// ExitClassModuleHeader is called when production classModuleHeader is exited.
func (s *BasevbaListener) ExitClassModuleHeader(ctx *ClassModuleHeaderContext) {}

// EnterClassAttr is called when production classAttr is entered.
func (s *BasevbaListener) EnterClassAttr(ctx *ClassAttrContext) {}

// ExitClassAttr is called when production classAttr is exited.
func (s *BasevbaListener) ExitClassAttr(ctx *ClassAttrContext) {}

// EnterIgnoredClassAttr is called when production ignoredClassAttr is entered.
func (s *BasevbaListener) EnterIgnoredClassAttr(ctx *IgnoredClassAttrContext) {}

// ExitIgnoredClassAttr is called when production ignoredClassAttr is exited.
func (s *BasevbaListener) ExitIgnoredClassAttr(ctx *IgnoredClassAttrContext) {}

// EnterIgnoredAttr is called when production ignoredAttr is entered.
func (s *BasevbaListener) EnterIgnoredAttr(ctx *IgnoredAttrContext) {}

// ExitIgnoredAttr is called when production ignoredAttr is exited.
func (s *BasevbaListener) ExitIgnoredAttr(ctx *IgnoredAttrContext) {}

// EnterNameAttr is called when production nameAttr is entered.
func (s *BasevbaListener) EnterNameAttr(ctx *NameAttrContext) {}

// ExitNameAttr is called when production nameAttr is exited.
func (s *BasevbaListener) ExitNameAttr(ctx *NameAttrContext) {}

// EnterProceduralModuleBody is called when production proceduralModuleBody is entered.
func (s *BasevbaListener) EnterProceduralModuleBody(ctx *ProceduralModuleBodyContext) {}

// ExitProceduralModuleBody is called when production proceduralModuleBody is exited.
func (s *BasevbaListener) ExitProceduralModuleBody(ctx *ProceduralModuleBodyContext) {}

// EnterClassModuleBody is called when production classModuleBody is entered.
func (s *BasevbaListener) EnterClassModuleBody(ctx *ClassModuleBodyContext) {}

// ExitClassModuleBody is called when production classModuleBody is exited.
func (s *BasevbaListener) ExitClassModuleBody(ctx *ClassModuleBodyContext) {}

// EnterUnrestrictedName is called when production unrestrictedName is entered.
func (s *BasevbaListener) EnterUnrestrictedName(ctx *UnrestrictedNameContext) {}

// ExitUnrestrictedName is called when production unrestrictedName is exited.
func (s *BasevbaListener) ExitUnrestrictedName(ctx *UnrestrictedNameContext) {}

// EnterName is called when production name is entered.
func (s *BasevbaListener) EnterName(ctx *NameContext) {}

// ExitName is called when production name is exited.
func (s *BasevbaListener) ExitName(ctx *NameContext) {}

// EnterUntypedName is called when production untypedName is entered.
func (s *BasevbaListener) EnterUntypedName(ctx *UntypedNameContext) {}

// ExitUntypedName is called when production untypedName is exited.
func (s *BasevbaListener) ExitUntypedName(ctx *UntypedNameContext) {}

// EnterProceduralModuleDirectiveElement is called when production proceduralModuleDirectiveElement is entered.
func (s *BasevbaListener) EnterProceduralModuleDirectiveElement(ctx *ProceduralModuleDirectiveElementContext) {
}

// ExitProceduralModuleDirectiveElement is called when production proceduralModuleDirectiveElement is exited.
func (s *BasevbaListener) ExitProceduralModuleDirectiveElement(ctx *ProceduralModuleDirectiveElementContext) {
}

// EnterProceduralModuleDeclarationElement is called when production proceduralModuleDeclarationElement is entered.
func (s *BasevbaListener) EnterProceduralModuleDeclarationElement(ctx *ProceduralModuleDeclarationElementContext) {
}

// ExitProceduralModuleDeclarationElement is called when production proceduralModuleDeclarationElement is exited.
func (s *BasevbaListener) ExitProceduralModuleDeclarationElement(ctx *ProceduralModuleDeclarationElementContext) {
}

// EnterClassModuleDirectiveElement is called when production classModuleDirectiveElement is entered.
func (s *BasevbaListener) EnterClassModuleDirectiveElement(ctx *ClassModuleDirectiveElementContext) {}

// ExitClassModuleDirectiveElement is called when production classModuleDirectiveElement is exited.
func (s *BasevbaListener) ExitClassModuleDirectiveElement(ctx *ClassModuleDirectiveElementContext) {}

// EnterClassModuleDeclarationElement is called when production classModuleDeclarationElement is entered.
func (s *BasevbaListener) EnterClassModuleDeclarationElement(ctx *ClassModuleDeclarationElementContext) {
}

// ExitClassModuleDeclarationElement is called when production classModuleDeclarationElement is exited.
func (s *BasevbaListener) ExitClassModuleDeclarationElement(ctx *ClassModuleDeclarationElementContext) {
}

// EnterCommonOptionDirective is called when production commonOptionDirective is entered.
func (s *BasevbaListener) EnterCommonOptionDirective(ctx *CommonOptionDirectiveContext) {}

// ExitCommonOptionDirective is called when production commonOptionDirective is exited.
func (s *BasevbaListener) ExitCommonOptionDirective(ctx *CommonOptionDirectiveContext) {}

// EnterOptionCompareDirective is called when production optionCompareDirective is entered.
func (s *BasevbaListener) EnterOptionCompareDirective(ctx *OptionCompareDirectiveContext) {}

// ExitOptionCompareDirective is called when production optionCompareDirective is exited.
func (s *BasevbaListener) ExitOptionCompareDirective(ctx *OptionCompareDirectiveContext) {}

// EnterOptionBaseDirective is called when production optionBaseDirective is entered.
func (s *BasevbaListener) EnterOptionBaseDirective(ctx *OptionBaseDirectiveContext) {}

// ExitOptionBaseDirective is called when production optionBaseDirective is exited.
func (s *BasevbaListener) ExitOptionBaseDirective(ctx *OptionBaseDirectiveContext) {}

// EnterOptionExplicitDirective is called when production optionExplicitDirective is entered.
func (s *BasevbaListener) EnterOptionExplicitDirective(ctx *OptionExplicitDirectiveContext) {}

// ExitOptionExplicitDirective is called when production optionExplicitDirective is exited.
func (s *BasevbaListener) ExitOptionExplicitDirective(ctx *OptionExplicitDirectiveContext) {}

// EnterOptionPrivateDirective is called when production optionPrivateDirective is entered.
func (s *BasevbaListener) EnterOptionPrivateDirective(ctx *OptionPrivateDirectiveContext) {}

// ExitOptionPrivateDirective is called when production optionPrivateDirective is exited.
func (s *BasevbaListener) ExitOptionPrivateDirective(ctx *OptionPrivateDirectiveContext) {}

// EnterDefDirective is called when production defDirective is entered.
func (s *BasevbaListener) EnterDefDirective(ctx *DefDirectiveContext) {}

// ExitDefDirective is called when production defDirective is exited.
func (s *BasevbaListener) ExitDefDirective(ctx *DefDirectiveContext) {}

// EnterLetterSpec is called when production letterSpec is entered.
func (s *BasevbaListener) EnterLetterSpec(ctx *LetterSpecContext) {}

// ExitLetterSpec is called when production letterSpec is exited.
func (s *BasevbaListener) ExitLetterSpec(ctx *LetterSpecContext) {}

// EnterSingleLetter is called when production singleLetter is entered.
func (s *BasevbaListener) EnterSingleLetter(ctx *SingleLetterContext) {}

// ExitSingleLetter is called when production singleLetter is exited.
func (s *BasevbaListener) ExitSingleLetter(ctx *SingleLetterContext) {}

// EnterUniversalLetterRange is called when production universalLetterRange is entered.
func (s *BasevbaListener) EnterUniversalLetterRange(ctx *UniversalLetterRangeContext) {}

// ExitUniversalLetterRange is called when production universalLetterRange is exited.
func (s *BasevbaListener) ExitUniversalLetterRange(ctx *UniversalLetterRangeContext) {}

// EnterUpperCaseA is called when production upperCaseA is entered.
func (s *BasevbaListener) EnterUpperCaseA(ctx *UpperCaseAContext) {}

// ExitUpperCaseA is called when production upperCaseA is exited.
func (s *BasevbaListener) ExitUpperCaseA(ctx *UpperCaseAContext) {}

// EnterUpperCaseZ is called when production upperCaseZ is entered.
func (s *BasevbaListener) EnterUpperCaseZ(ctx *UpperCaseZContext) {}

// ExitUpperCaseZ is called when production upperCaseZ is exited.
func (s *BasevbaListener) ExitUpperCaseZ(ctx *UpperCaseZContext) {}

// EnterLetterRange is called when production letterRange is entered.
func (s *BasevbaListener) EnterLetterRange(ctx *LetterRangeContext) {}

// ExitLetterRange is called when production letterRange is exited.
func (s *BasevbaListener) ExitLetterRange(ctx *LetterRangeContext) {}

// EnterFirstLetter is called when production firstLetter is entered.
func (s *BasevbaListener) EnterFirstLetter(ctx *FirstLetterContext) {}

// ExitFirstLetter is called when production firstLetter is exited.
func (s *BasevbaListener) ExitFirstLetter(ctx *FirstLetterContext) {}

// EnterLastLetter is called when production lastLetter is entered.
func (s *BasevbaListener) EnterLastLetter(ctx *LastLetterContext) {}

// ExitLastLetter is called when production lastLetter is exited.
func (s *BasevbaListener) ExitLastLetter(ctx *LastLetterContext) {}

// EnterDefType is called when production defType is entered.
func (s *BasevbaListener) EnterDefType(ctx *DefTypeContext) {}

// ExitDefType is called when production defType is exited.
func (s *BasevbaListener) ExitDefType(ctx *DefTypeContext) {}

// EnterCommonModuleDeclarationElement is called when production commonModuleDeclarationElement is entered.
func (s *BasevbaListener) EnterCommonModuleDeclarationElement(ctx *CommonModuleDeclarationElementContext) {
}

// ExitCommonModuleDeclarationElement is called when production commonModuleDeclarationElement is exited.
func (s *BasevbaListener) ExitCommonModuleDeclarationElement(ctx *CommonModuleDeclarationElementContext) {
}

// EnterVariableDeclaration is called when production variableDeclaration is entered.
func (s *BasevbaListener) EnterVariableDeclaration(ctx *VariableDeclarationContext) {}

// ExitVariableDeclaration is called when production variableDeclaration is exited.
func (s *BasevbaListener) ExitVariableDeclaration(ctx *VariableDeclarationContext) {}

// EnterVariableHelpAttribute is called when production variableHelpAttribute is entered.
func (s *BasevbaListener) EnterVariableHelpAttribute(ctx *VariableHelpAttributeContext) {}

// ExitVariableHelpAttribute is called when production variableHelpAttribute is exited.
func (s *BasevbaListener) ExitVariableHelpAttribute(ctx *VariableHelpAttributeContext) {}

// EnterVariableModifier is called when production variableModifier is entered.
func (s *BasevbaListener) EnterVariableModifier(ctx *VariableModifierContext) {}

// ExitVariableModifier is called when production variableModifier is exited.
func (s *BasevbaListener) ExitVariableModifier(ctx *VariableModifierContext) {}

// EnterVariableSharedModifier is called when production variableSharedModifier is entered.
func (s *BasevbaListener) EnterVariableSharedModifier(ctx *VariableSharedModifierContext) {}

// ExitVariableSharedModifier is called when production variableSharedModifier is exited.
func (s *BasevbaListener) ExitVariableSharedModifier(ctx *VariableSharedModifierContext) {}

// EnterModuleVariableDeclarationList is called when production moduleVariableDeclarationList is entered.
func (s *BasevbaListener) EnterModuleVariableDeclarationList(ctx *ModuleVariableDeclarationListContext) {
}

// ExitModuleVariableDeclarationList is called when production moduleVariableDeclarationList is exited.
func (s *BasevbaListener) ExitModuleVariableDeclarationList(ctx *ModuleVariableDeclarationListContext) {
}

// EnterVariableDeclarationList is called when production variableDeclarationList is entered.
func (s *BasevbaListener) EnterVariableDeclarationList(ctx *VariableDeclarationListContext) {}

// ExitVariableDeclarationList is called when production variableDeclarationList is exited.
func (s *BasevbaListener) ExitVariableDeclarationList(ctx *VariableDeclarationListContext) {}

// EnterVariableDcl is called when production variableDcl is entered.
func (s *BasevbaListener) EnterVariableDcl(ctx *VariableDclContext) {}

// ExitVariableDcl is called when production variableDcl is exited.
func (s *BasevbaListener) ExitVariableDcl(ctx *VariableDclContext) {}

// EnterTypedVariableDcl is called when production typedVariableDcl is entered.
func (s *BasevbaListener) EnterTypedVariableDcl(ctx *TypedVariableDclContext) {}

// ExitTypedVariableDcl is called when production typedVariableDcl is exited.
func (s *BasevbaListener) ExitTypedVariableDcl(ctx *TypedVariableDclContext) {}

// EnterUntypedVariableDcl is called when production untypedVariableDcl is entered.
func (s *BasevbaListener) EnterUntypedVariableDcl(ctx *UntypedVariableDclContext) {}

// ExitUntypedVariableDcl is called when production untypedVariableDcl is exited.
func (s *BasevbaListener) ExitUntypedVariableDcl(ctx *UntypedVariableDclContext) {}

// EnterArrayClause is called when production arrayClause is entered.
func (s *BasevbaListener) EnterArrayClause(ctx *ArrayClauseContext) {}

// ExitArrayClause is called when production arrayClause is exited.
func (s *BasevbaListener) ExitArrayClause(ctx *ArrayClauseContext) {}

// EnterAsClause is called when production asClause is entered.
func (s *BasevbaListener) EnterAsClause(ctx *AsClauseContext) {}

// ExitAsClause is called when production asClause is exited.
func (s *BasevbaListener) ExitAsClause(ctx *AsClauseContext) {}

// EnterWitheventsVariableDcl is called when production witheventsVariableDcl is entered.
func (s *BasevbaListener) EnterWitheventsVariableDcl(ctx *WitheventsVariableDclContext) {}

// ExitWitheventsVariableDcl is called when production witheventsVariableDcl is exited.
func (s *BasevbaListener) ExitWitheventsVariableDcl(ctx *WitheventsVariableDclContext) {}

// EnterClassTypeName is called when production classTypeName is entered.
func (s *BasevbaListener) EnterClassTypeName(ctx *ClassTypeNameContext) {}

// ExitClassTypeName is called when production classTypeName is exited.
func (s *BasevbaListener) ExitClassTypeName(ctx *ClassTypeNameContext) {}

// EnterArrayDim is called when production arrayDim is entered.
func (s *BasevbaListener) EnterArrayDim(ctx *ArrayDimContext) {}

// ExitArrayDim is called when production arrayDim is exited.
func (s *BasevbaListener) ExitArrayDim(ctx *ArrayDimContext) {}

// EnterBoundsList is called when production boundsList is entered.
func (s *BasevbaListener) EnterBoundsList(ctx *BoundsListContext) {}

// ExitBoundsList is called when production boundsList is exited.
func (s *BasevbaListener) ExitBoundsList(ctx *BoundsListContext) {}

// EnterDimSpec is called when production dimSpec is entered.
func (s *BasevbaListener) EnterDimSpec(ctx *DimSpecContext) {}

// ExitDimSpec is called when production dimSpec is exited.
func (s *BasevbaListener) ExitDimSpec(ctx *DimSpecContext) {}

// EnterLowerBound is called when production lowerBound is entered.
func (s *BasevbaListener) EnterLowerBound(ctx *LowerBoundContext) {}

// ExitLowerBound is called when production lowerBound is exited.
func (s *BasevbaListener) ExitLowerBound(ctx *LowerBoundContext) {}

// EnterUpperBound is called when production upperBound is entered.
func (s *BasevbaListener) EnterUpperBound(ctx *UpperBoundContext) {}

// ExitUpperBound is called when production upperBound is exited.
func (s *BasevbaListener) ExitUpperBound(ctx *UpperBoundContext) {}

// EnterAsAutoObject is called when production asAutoObject is entered.
func (s *BasevbaListener) EnterAsAutoObject(ctx *AsAutoObjectContext) {}

// ExitAsAutoObject is called when production asAutoObject is exited.
func (s *BasevbaListener) ExitAsAutoObject(ctx *AsAutoObjectContext) {}

// EnterAsType is called when production asType is entered.
func (s *BasevbaListener) EnterAsType(ctx *AsTypeContext) {}

// ExitAsType is called when production asType is exited.
func (s *BasevbaListener) ExitAsType(ctx *AsTypeContext) {}

// EnterTypeSpec is called when production typeSpec is entered.
func (s *BasevbaListener) EnterTypeSpec(ctx *TypeSpecContext) {}

// ExitTypeSpec is called when production typeSpec is exited.
func (s *BasevbaListener) ExitTypeSpec(ctx *TypeSpecContext) {}

// EnterFixedLengthStringSpec is called when production fixedLengthStringSpec is entered.
func (s *BasevbaListener) EnterFixedLengthStringSpec(ctx *FixedLengthStringSpecContext) {}

// ExitFixedLengthStringSpec is called when production fixedLengthStringSpec is exited.
func (s *BasevbaListener) ExitFixedLengthStringSpec(ctx *FixedLengthStringSpecContext) {}

// EnterStringLength is called when production stringLength is entered.
func (s *BasevbaListener) EnterStringLength(ctx *StringLengthContext) {}

// ExitStringLength is called when production stringLength is exited.
func (s *BasevbaListener) ExitStringLength(ctx *StringLengthContext) {}

// EnterConstantName is called when production constantName is entered.
func (s *BasevbaListener) EnterConstantName(ctx *ConstantNameContext) {}

// ExitConstantName is called when production constantName is exited.
func (s *BasevbaListener) ExitConstantName(ctx *ConstantNameContext) {}

// EnterConstDeclaration is called when production constDeclaration is entered.
func (s *BasevbaListener) EnterConstDeclaration(ctx *ConstDeclarationContext) {}

// ExitConstDeclaration is called when production constDeclaration is exited.
func (s *BasevbaListener) ExitConstDeclaration(ctx *ConstDeclarationContext) {}

// EnterConstItemList is called when production constItemList is entered.
func (s *BasevbaListener) EnterConstItemList(ctx *ConstItemListContext) {}

// ExitConstItemList is called when production constItemList is exited.
func (s *BasevbaListener) ExitConstItemList(ctx *ConstItemListContext) {}

// EnterConstItem is called when production constItem is entered.
func (s *BasevbaListener) EnterConstItem(ctx *ConstItemContext) {}

// ExitConstItem is called when production constItem is exited.
func (s *BasevbaListener) ExitConstItem(ctx *ConstItemContext) {}

// EnterConstAsClause is called when production constAsClause is entered.
func (s *BasevbaListener) EnterConstAsClause(ctx *ConstAsClauseContext) {}

// ExitConstAsClause is called when production constAsClause is exited.
func (s *BasevbaListener) ExitConstAsClause(ctx *ConstAsClauseContext) {}

// EnterPublicTypeDeclaration is called when production publicTypeDeclaration is entered.
func (s *BasevbaListener) EnterPublicTypeDeclaration(ctx *PublicTypeDeclarationContext) {}

// ExitPublicTypeDeclaration is called when production publicTypeDeclaration is exited.
func (s *BasevbaListener) ExitPublicTypeDeclaration(ctx *PublicTypeDeclarationContext) {}

// EnterPrivateTypeDeclaration is called when production privateTypeDeclaration is entered.
func (s *BasevbaListener) EnterPrivateTypeDeclaration(ctx *PrivateTypeDeclarationContext) {}

// ExitPrivateTypeDeclaration is called when production privateTypeDeclaration is exited.
func (s *BasevbaListener) ExitPrivateTypeDeclaration(ctx *PrivateTypeDeclarationContext) {}

// EnterUdtDeclaration is called when production udtDeclaration is entered.
func (s *BasevbaListener) EnterUdtDeclaration(ctx *UdtDeclarationContext) {}

// ExitUdtDeclaration is called when production udtDeclaration is exited.
func (s *BasevbaListener) ExitUdtDeclaration(ctx *UdtDeclarationContext) {}

// EnterUdtMemberList is called when production udtMemberList is entered.
func (s *BasevbaListener) EnterUdtMemberList(ctx *UdtMemberListContext) {}

// ExitUdtMemberList is called when production udtMemberList is exited.
func (s *BasevbaListener) ExitUdtMemberList(ctx *UdtMemberListContext) {}

// EnterUdtElement is called when production udtElement is entered.
func (s *BasevbaListener) EnterUdtElement(ctx *UdtElementContext) {}

// ExitUdtElement is called when production udtElement is exited.
func (s *BasevbaListener) ExitUdtElement(ctx *UdtElementContext) {}

// EnterUdtMember is called when production udtMember is entered.
func (s *BasevbaListener) EnterUdtMember(ctx *UdtMemberContext) {}

// ExitUdtMember is called when production udtMember is exited.
func (s *BasevbaListener) ExitUdtMember(ctx *UdtMemberContext) {}

// EnterUntypedNameMemberDcl is called when production untypedNameMemberDcl is entered.
func (s *BasevbaListener) EnterUntypedNameMemberDcl(ctx *UntypedNameMemberDclContext) {}

// ExitUntypedNameMemberDcl is called when production untypedNameMemberDcl is exited.
func (s *BasevbaListener) ExitUntypedNameMemberDcl(ctx *UntypedNameMemberDclContext) {}

// EnterReservedNameMemberDcl is called when production reservedNameMemberDcl is entered.
func (s *BasevbaListener) EnterReservedNameMemberDcl(ctx *ReservedNameMemberDclContext) {}

// ExitReservedNameMemberDcl is called when production reservedNameMemberDcl is exited.
func (s *BasevbaListener) ExitReservedNameMemberDcl(ctx *ReservedNameMemberDclContext) {}

// EnterOptionalArrayClause is called when production optionalArrayClause is entered.
func (s *BasevbaListener) EnterOptionalArrayClause(ctx *OptionalArrayClauseContext) {}

// ExitOptionalArrayClause is called when production optionalArrayClause is exited.
func (s *BasevbaListener) ExitOptionalArrayClause(ctx *OptionalArrayClauseContext) {}

// EnterReservedMemberName is called when production reservedMemberName is entered.
func (s *BasevbaListener) EnterReservedMemberName(ctx *ReservedMemberNameContext) {}

// ExitReservedMemberName is called when production reservedMemberName is exited.
func (s *BasevbaListener) ExitReservedMemberName(ctx *ReservedMemberNameContext) {}

// EnterGlobalEnumDeclaration is called when production globalEnumDeclaration is entered.
func (s *BasevbaListener) EnterGlobalEnumDeclaration(ctx *GlobalEnumDeclarationContext) {}

// ExitGlobalEnumDeclaration is called when production globalEnumDeclaration is exited.
func (s *BasevbaListener) ExitGlobalEnumDeclaration(ctx *GlobalEnumDeclarationContext) {}

// EnterPublicEnumDeclaration is called when production publicEnumDeclaration is entered.
func (s *BasevbaListener) EnterPublicEnumDeclaration(ctx *PublicEnumDeclarationContext) {}

// ExitPublicEnumDeclaration is called when production publicEnumDeclaration is exited.
func (s *BasevbaListener) ExitPublicEnumDeclaration(ctx *PublicEnumDeclarationContext) {}

// EnterPrivateEnumDeclaration is called when production privateEnumDeclaration is entered.
func (s *BasevbaListener) EnterPrivateEnumDeclaration(ctx *PrivateEnumDeclarationContext) {}

// ExitPrivateEnumDeclaration is called when production privateEnumDeclaration is exited.
func (s *BasevbaListener) ExitPrivateEnumDeclaration(ctx *PrivateEnumDeclarationContext) {}

// EnterEnumDeclaration is called when production enumDeclaration is entered.
func (s *BasevbaListener) EnterEnumDeclaration(ctx *EnumDeclarationContext) {}

// ExitEnumDeclaration is called when production enumDeclaration is exited.
func (s *BasevbaListener) ExitEnumDeclaration(ctx *EnumDeclarationContext) {}

// EnterEnumLongptrDeclaration is called when production enumLongptrDeclaration is entered.
func (s *BasevbaListener) EnterEnumLongptrDeclaration(ctx *EnumLongptrDeclarationContext) {}

// ExitEnumLongptrDeclaration is called when production enumLongptrDeclaration is exited.
func (s *BasevbaListener) ExitEnumLongptrDeclaration(ctx *EnumLongptrDeclarationContext) {}

// EnterEnumMemberList is called when production enumMemberList is entered.
func (s *BasevbaListener) EnterEnumMemberList(ctx *EnumMemberListContext) {}

// ExitEnumMemberList is called when production enumMemberList is exited.
func (s *BasevbaListener) ExitEnumMemberList(ctx *EnumMemberListContext) {}

// EnterEnumElement is called when production enumElement is entered.
func (s *BasevbaListener) EnterEnumElement(ctx *EnumElementContext) {}

// ExitEnumElement is called when production enumElement is exited.
func (s *BasevbaListener) ExitEnumElement(ctx *EnumElementContext) {}

// EnterEnumMember is called when production enumMember is entered.
func (s *BasevbaListener) EnterEnumMember(ctx *EnumMemberContext) {}

// ExitEnumMember is called when production enumMember is exited.
func (s *BasevbaListener) ExitEnumMember(ctx *EnumMemberContext) {}

// EnterPublicExternalProcedureDeclaration is called when production publicExternalProcedureDeclaration is entered.
func (s *BasevbaListener) EnterPublicExternalProcedureDeclaration(ctx *PublicExternalProcedureDeclarationContext) {
}

// ExitPublicExternalProcedureDeclaration is called when production publicExternalProcedureDeclaration is exited.
func (s *BasevbaListener) ExitPublicExternalProcedureDeclaration(ctx *PublicExternalProcedureDeclarationContext) {
}

// EnterPrivateExternalProcedureDeclaration is called when production privateExternalProcedureDeclaration is entered.
func (s *BasevbaListener) EnterPrivateExternalProcedureDeclaration(ctx *PrivateExternalProcedureDeclarationContext) {
}

// ExitPrivateExternalProcedureDeclaration is called when production privateExternalProcedureDeclaration is exited.
func (s *BasevbaListener) ExitPrivateExternalProcedureDeclaration(ctx *PrivateExternalProcedureDeclarationContext) {
}

// EnterExternalProcDcl is called when production externalProcDcl is entered.
func (s *BasevbaListener) EnterExternalProcDcl(ctx *ExternalProcDclContext) {}

// ExitExternalProcDcl is called when production externalProcDcl is exited.
func (s *BasevbaListener) ExitExternalProcDcl(ctx *ExternalProcDclContext) {}

// EnterExternalSub is called when production externalSub is entered.
func (s *BasevbaListener) EnterExternalSub(ctx *ExternalSubContext) {}

// ExitExternalSub is called when production externalSub is exited.
func (s *BasevbaListener) ExitExternalSub(ctx *ExternalSubContext) {}

// EnterExternalFunction is called when production externalFunction is entered.
func (s *BasevbaListener) EnterExternalFunction(ctx *ExternalFunctionContext) {}

// ExitExternalFunction is called when production externalFunction is exited.
func (s *BasevbaListener) ExitExternalFunction(ctx *ExternalFunctionContext) {}

// EnterLibInfo is called when production libInfo is entered.
func (s *BasevbaListener) EnterLibInfo(ctx *LibInfoContext) {}

// ExitLibInfo is called when production libInfo is exited.
func (s *BasevbaListener) ExitLibInfo(ctx *LibInfoContext) {}

// EnterLibClause is called when production libClause is entered.
func (s *BasevbaListener) EnterLibClause(ctx *LibClauseContext) {}

// ExitLibClause is called when production libClause is exited.
func (s *BasevbaListener) ExitLibClause(ctx *LibClauseContext) {}

// EnterAliasClause is called when production aliasClause is entered.
func (s *BasevbaListener) EnterAliasClause(ctx *AliasClauseContext) {}

// ExitAliasClause is called when production aliasClause is exited.
func (s *BasevbaListener) ExitAliasClause(ctx *AliasClauseContext) {}

// EnterImplementsDirective is called when production implementsDirective is entered.
func (s *BasevbaListener) EnterImplementsDirective(ctx *ImplementsDirectiveContext) {}

// ExitImplementsDirective is called when production implementsDirective is exited.
func (s *BasevbaListener) ExitImplementsDirective(ctx *ImplementsDirectiveContext) {}

// EnterEventDeclaration is called when production eventDeclaration is entered.
func (s *BasevbaListener) EnterEventDeclaration(ctx *EventDeclarationContext) {}

// ExitEventDeclaration is called when production eventDeclaration is exited.
func (s *BasevbaListener) ExitEventDeclaration(ctx *EventDeclarationContext) {}

// EnterEventParameterList is called when production eventParameterList is entered.
func (s *BasevbaListener) EnterEventParameterList(ctx *EventParameterListContext) {}

// ExitEventParameterList is called when production eventParameterList is exited.
func (s *BasevbaListener) ExitEventParameterList(ctx *EventParameterListContext) {}

// EnterProceduralModuleCode is called when production proceduralModuleCode is entered.
func (s *BasevbaListener) EnterProceduralModuleCode(ctx *ProceduralModuleCodeContext) {}

// ExitProceduralModuleCode is called when production proceduralModuleCode is exited.
func (s *BasevbaListener) ExitProceduralModuleCode(ctx *ProceduralModuleCodeContext) {}

// EnterClassModuleCode is called when production classModuleCode is entered.
func (s *BasevbaListener) EnterClassModuleCode(ctx *ClassModuleCodeContext) {}

// ExitClassModuleCode is called when production classModuleCode is exited.
func (s *BasevbaListener) ExitClassModuleCode(ctx *ClassModuleCodeContext) {}

// EnterProceduralModuleCodeElement is called when production proceduralModuleCodeElement is entered.
func (s *BasevbaListener) EnterProceduralModuleCodeElement(ctx *ProceduralModuleCodeElementContext) {}

// ExitProceduralModuleCodeElement is called when production proceduralModuleCodeElement is exited.
func (s *BasevbaListener) ExitProceduralModuleCodeElement(ctx *ProceduralModuleCodeElementContext) {}

// EnterClassModuleCodeElement is called when production classModuleCodeElement is entered.
func (s *BasevbaListener) EnterClassModuleCodeElement(ctx *ClassModuleCodeElementContext) {}

// ExitClassModuleCodeElement is called when production classModuleCodeElement is exited.
func (s *BasevbaListener) ExitClassModuleCodeElement(ctx *ClassModuleCodeElementContext) {}

// EnterCommonModuleCodeElement is called when production commonModuleCodeElement is entered.
func (s *BasevbaListener) EnterCommonModuleCodeElement(ctx *CommonModuleCodeElementContext) {}

// ExitCommonModuleCodeElement is called when production commonModuleCodeElement is exited.
func (s *BasevbaListener) ExitCommonModuleCodeElement(ctx *CommonModuleCodeElementContext) {}

// EnterProcedureDeclaration is called when production procedureDeclaration is entered.
func (s *BasevbaListener) EnterProcedureDeclaration(ctx *ProcedureDeclarationContext) {}

// ExitProcedureDeclaration is called when production procedureDeclaration is exited.
func (s *BasevbaListener) ExitProcedureDeclaration(ctx *ProcedureDeclarationContext) {}

// EnterSubroutineDeclaration is called when production subroutineDeclaration is entered.
func (s *BasevbaListener) EnterSubroutineDeclaration(ctx *SubroutineDeclarationContext) {}

// ExitSubroutineDeclaration is called when production subroutineDeclaration is exited.
func (s *BasevbaListener) ExitSubroutineDeclaration(ctx *SubroutineDeclarationContext) {}

// EnterFunctionDeclaration is called when production functionDeclaration is entered.
func (s *BasevbaListener) EnterFunctionDeclaration(ctx *FunctionDeclarationContext) {}

// ExitFunctionDeclaration is called when production functionDeclaration is exited.
func (s *BasevbaListener) ExitFunctionDeclaration(ctx *FunctionDeclarationContext) {}

// EnterPropertyGetDeclaration is called when production propertyGetDeclaration is entered.
func (s *BasevbaListener) EnterPropertyGetDeclaration(ctx *PropertyGetDeclarationContext) {}

// ExitPropertyGetDeclaration is called when production propertyGetDeclaration is exited.
func (s *BasevbaListener) ExitPropertyGetDeclaration(ctx *PropertyGetDeclarationContext) {}

// EnterPropertySetDeclaration is called when production propertySetDeclaration is entered.
func (s *BasevbaListener) EnterPropertySetDeclaration(ctx *PropertySetDeclarationContext) {}

// ExitPropertySetDeclaration is called when production propertySetDeclaration is exited.
func (s *BasevbaListener) ExitPropertySetDeclaration(ctx *PropertySetDeclarationContext) {}

// EnterEndLabel is called when production endLabel is entered.
func (s *BasevbaListener) EnterEndLabel(ctx *EndLabelContext) {}

// ExitEndLabel is called when production endLabel is exited.
func (s *BasevbaListener) ExitEndLabel(ctx *EndLabelContext) {}

// EnterProcedureTail is called when production procedureTail is entered.
func (s *BasevbaListener) EnterProcedureTail(ctx *ProcedureTailContext) {}

// ExitProcedureTail is called when production procedureTail is exited.
func (s *BasevbaListener) ExitProcedureTail(ctx *ProcedureTailContext) {}

// EnterProcedureScope is called when production procedureScope is entered.
func (s *BasevbaListener) EnterProcedureScope(ctx *ProcedureScopeContext) {}

// ExitProcedureScope is called when production procedureScope is exited.
func (s *BasevbaListener) ExitProcedureScope(ctx *ProcedureScopeContext) {}

// EnterInitialStatic is called when production initialStatic is entered.
func (s *BasevbaListener) EnterInitialStatic(ctx *InitialStaticContext) {}

// ExitInitialStatic is called when production initialStatic is exited.
func (s *BasevbaListener) ExitInitialStatic(ctx *InitialStaticContext) {}

// EnterTrailingStatic is called when production trailingStatic is entered.
func (s *BasevbaListener) EnterTrailingStatic(ctx *TrailingStaticContext) {}

// ExitTrailingStatic is called when production trailingStatic is exited.
func (s *BasevbaListener) ExitTrailingStatic(ctx *TrailingStaticContext) {}

// EnterSubroutineName is called when production subroutineName is entered.
func (s *BasevbaListener) EnterSubroutineName(ctx *SubroutineNameContext) {}

// ExitSubroutineName is called when production subroutineName is exited.
func (s *BasevbaListener) ExitSubroutineName(ctx *SubroutineNameContext) {}

// EnterFunctionName is called when production functionName is entered.
func (s *BasevbaListener) EnterFunctionName(ctx *FunctionNameContext) {}

// ExitFunctionName is called when production functionName is exited.
func (s *BasevbaListener) ExitFunctionName(ctx *FunctionNameContext) {}

// EnterPrefixedName is called when production prefixedName is entered.
func (s *BasevbaListener) EnterPrefixedName(ctx *PrefixedNameContext) {}

// ExitPrefixedName is called when production prefixedName is exited.
func (s *BasevbaListener) ExitPrefixedName(ctx *PrefixedNameContext) {}

// EnterFunctionType is called when production functionType is entered.
func (s *BasevbaListener) EnterFunctionType(ctx *FunctionTypeContext) {}

// ExitFunctionType is called when production functionType is exited.
func (s *BasevbaListener) ExitFunctionType(ctx *FunctionTypeContext) {}

// EnterArrayDesignator is called when production arrayDesignator is entered.
func (s *BasevbaListener) EnterArrayDesignator(ctx *ArrayDesignatorContext) {}

// ExitArrayDesignator is called when production arrayDesignator is exited.
func (s *BasevbaListener) ExitArrayDesignator(ctx *ArrayDesignatorContext) {}

// EnterProcedureParameters is called when production procedureParameters is entered.
func (s *BasevbaListener) EnterProcedureParameters(ctx *ProcedureParametersContext) {}

// ExitProcedureParameters is called when production procedureParameters is exited.
func (s *BasevbaListener) ExitProcedureParameters(ctx *ProcedureParametersContext) {}

// EnterPropertyParameters is called when production propertyParameters is entered.
func (s *BasevbaListener) EnterPropertyParameters(ctx *PropertyParametersContext) {}

// ExitPropertyParameters is called when production propertyParameters is exited.
func (s *BasevbaListener) ExitPropertyParameters(ctx *PropertyParametersContext) {}

// EnterPropertyParameterList is called when production propertyParameterList is entered.
func (s *BasevbaListener) EnterPropertyParameterList(ctx *PropertyParameterListContext) {}

// ExitPropertyParameterList is called when production propertyParameterList is exited.
func (s *BasevbaListener) ExitPropertyParameterList(ctx *PropertyParameterListContext) {}

// EnterParameterList is called when production parameterList is entered.
func (s *BasevbaListener) EnterParameterList(ctx *ParameterListContext) {}

// ExitParameterList is called when production parameterList is exited.
func (s *BasevbaListener) ExitParameterList(ctx *ParameterListContext) {}

// EnterAnyParam is called when production anyParam is entered.
func (s *BasevbaListener) EnterAnyParam(ctx *AnyParamContext) {}

// ExitAnyParam is called when production anyParam is exited.
func (s *BasevbaListener) ExitAnyParam(ctx *AnyParamContext) {}

// EnterPositionalParameters is called when production positionalParameters is entered.
func (s *BasevbaListener) EnterPositionalParameters(ctx *PositionalParametersContext) {}

// ExitPositionalParameters is called when production positionalParameters is exited.
func (s *BasevbaListener) ExitPositionalParameters(ctx *PositionalParametersContext) {}

// EnterValueParam is called when production valueParam is entered.
func (s *BasevbaListener) EnterValueParam(ctx *ValueParamContext) {}

// ExitValueParam is called when production valueParam is exited.
func (s *BasevbaListener) ExitValueParam(ctx *ValueParamContext) {}

// EnterPositionalParam is called when production positionalParam is entered.
func (s *BasevbaListener) EnterPositionalParam(ctx *PositionalParamContext) {}

// ExitPositionalParam is called when production positionalParam is exited.
func (s *BasevbaListener) ExitPositionalParam(ctx *PositionalParamContext) {}

// EnterOptionalParam is called when production optionalParam is entered.
func (s *BasevbaListener) EnterOptionalParam(ctx *OptionalParamContext) {}

// ExitOptionalParam is called when production optionalParam is exited.
func (s *BasevbaListener) ExitOptionalParam(ctx *OptionalParamContext) {}

// EnterParamArray is called when production paramArray is entered.
func (s *BasevbaListener) EnterParamArray(ctx *ParamArrayContext) {}

// ExitParamArray is called when production paramArray is exited.
func (s *BasevbaListener) ExitParamArray(ctx *ParamArrayContext) {}

// EnterParamDcl is called when production paramDcl is entered.
func (s *BasevbaListener) EnterParamDcl(ctx *ParamDclContext) {}

// ExitParamDcl is called when production paramDcl is exited.
func (s *BasevbaListener) ExitParamDcl(ctx *ParamDclContext) {}

// EnterUntypedNameParamDcl is called when production untypedNameParamDcl is entered.
func (s *BasevbaListener) EnterUntypedNameParamDcl(ctx *UntypedNameParamDclContext) {}

// ExitUntypedNameParamDcl is called when production untypedNameParamDcl is exited.
func (s *BasevbaListener) ExitUntypedNameParamDcl(ctx *UntypedNameParamDclContext) {}

// EnterTypedNameParamDcl is called when production typedNameParamDcl is entered.
func (s *BasevbaListener) EnterTypedNameParamDcl(ctx *TypedNameParamDclContext) {}

// ExitTypedNameParamDcl is called when production typedNameParamDcl is exited.
func (s *BasevbaListener) ExitTypedNameParamDcl(ctx *TypedNameParamDclContext) {}

// EnterOptionalPrefix is called when production optionalPrefix is entered.
func (s *BasevbaListener) EnterOptionalPrefix(ctx *OptionalPrefixContext) {}

// ExitOptionalPrefix is called when production optionalPrefix is exited.
func (s *BasevbaListener) ExitOptionalPrefix(ctx *OptionalPrefixContext) {}

// EnterParameterMechanism is called when production parameterMechanism is entered.
func (s *BasevbaListener) EnterParameterMechanism(ctx *ParameterMechanismContext) {}

// ExitParameterMechanism is called when production parameterMechanism is exited.
func (s *BasevbaListener) ExitParameterMechanism(ctx *ParameterMechanismContext) {}

// EnterParameterType is called when production parameterType is entered.
func (s *BasevbaListener) EnterParameterType(ctx *ParameterTypeContext) {}

// ExitParameterType is called when production parameterType is exited.
func (s *BasevbaListener) ExitParameterType(ctx *ParameterTypeContext) {}

// EnterDefaultValue is called when production defaultValue is entered.
func (s *BasevbaListener) EnterDefaultValue(ctx *DefaultValueContext) {}

// ExitDefaultValue is called when production defaultValue is exited.
func (s *BasevbaListener) ExitDefaultValue(ctx *DefaultValueContext) {}

// EnterEventHandlerName is called when production eventHandlerName is entered.
func (s *BasevbaListener) EnterEventHandlerName(ctx *EventHandlerNameContext) {}

// ExitEventHandlerName is called when production eventHandlerName is exited.
func (s *BasevbaListener) ExitEventHandlerName(ctx *EventHandlerNameContext) {}

// EnterImplementedName is called when production implementedName is entered.
func (s *BasevbaListener) EnterImplementedName(ctx *ImplementedNameContext) {}

// ExitImplementedName is called when production implementedName is exited.
func (s *BasevbaListener) ExitImplementedName(ctx *ImplementedNameContext) {}

// EnterLifecycleHandlerName is called when production lifecycleHandlerName is entered.
func (s *BasevbaListener) EnterLifecycleHandlerName(ctx *LifecycleHandlerNameContext) {}

// ExitLifecycleHandlerName is called when production lifecycleHandlerName is exited.
func (s *BasevbaListener) ExitLifecycleHandlerName(ctx *LifecycleHandlerNameContext) {}

// EnterProcedureBody is called when production procedureBody is entered.
func (s *BasevbaListener) EnterProcedureBody(ctx *ProcedureBodyContext) {}

// ExitProcedureBody is called when production procedureBody is exited.
func (s *BasevbaListener) ExitProcedureBody(ctx *ProcedureBodyContext) {}

// EnterStatementBlock is called when production statementBlock is entered.
func (s *BasevbaListener) EnterStatementBlock(ctx *StatementBlockContext) {}

// ExitStatementBlock is called when production statementBlock is exited.
func (s *BasevbaListener) ExitStatementBlock(ctx *StatementBlockContext) {}

// EnterBlockStatement is called when production blockStatement is entered.
func (s *BasevbaListener) EnterBlockStatement(ctx *BlockStatementContext) {}

// ExitBlockStatement is called when production blockStatement is exited.
func (s *BasevbaListener) ExitBlockStatement(ctx *BlockStatementContext) {}

// EnterStatement is called when production statement is entered.
func (s *BasevbaListener) EnterStatement(ctx *StatementContext) {}

// ExitStatement is called when production statement is exited.
func (s *BasevbaListener) ExitStatement(ctx *StatementContext) {}

// EnterStatementLabelDefinition is called when production statementLabelDefinition is entered.
func (s *BasevbaListener) EnterStatementLabelDefinition(ctx *StatementLabelDefinitionContext) {}

// ExitStatementLabelDefinition is called when production statementLabelDefinition is exited.
func (s *BasevbaListener) ExitStatementLabelDefinition(ctx *StatementLabelDefinitionContext) {}

// EnterStatementLabel is called when production statementLabel is entered.
func (s *BasevbaListener) EnterStatementLabel(ctx *StatementLabelContext) {}

// ExitStatementLabel is called when production statementLabel is exited.
func (s *BasevbaListener) ExitStatementLabel(ctx *StatementLabelContext) {}

// EnterStatementLabelList is called when production statementLabelList is entered.
func (s *BasevbaListener) EnterStatementLabelList(ctx *StatementLabelListContext) {}

// ExitStatementLabelList is called when production statementLabelList is exited.
func (s *BasevbaListener) ExitStatementLabelList(ctx *StatementLabelListContext) {}

// EnterIdentifierStatementLabel is called when production identifierStatementLabel is entered.
func (s *BasevbaListener) EnterIdentifierStatementLabel(ctx *IdentifierStatementLabelContext) {}

// ExitIdentifierStatementLabel is called when production identifierStatementLabel is exited.
func (s *BasevbaListener) ExitIdentifierStatementLabel(ctx *IdentifierStatementLabelContext) {}

// EnterResetNumberLabel is called when production resetNumberLabel is entered.
func (s *BasevbaListener) EnterResetNumberLabel(ctx *ResetNumberLabelContext) {}

// ExitResetNumberLabel is called when production resetNumberLabel is exited.
func (s *BasevbaListener) ExitResetNumberLabel(ctx *ResetNumberLabelContext) {}

// EnterLineNumberLabel is called when production lineNumberLabel is entered.
func (s *BasevbaListener) EnterLineNumberLabel(ctx *LineNumberLabelContext) {}

// ExitLineNumberLabel is called when production lineNumberLabel is exited.
func (s *BasevbaListener) ExitLineNumberLabel(ctx *LineNumberLabelContext) {}

// EnterRemStatement is called when production remStatement is entered.
func (s *BasevbaListener) EnterRemStatement(ctx *RemStatementContext) {}

// ExitRemStatement is called when production remStatement is exited.
func (s *BasevbaListener) ExitRemStatement(ctx *RemStatementContext) {}

// EnterControlStatement is called when production controlStatement is entered.
func (s *BasevbaListener) EnterControlStatement(ctx *ControlStatementContext) {}

// ExitControlStatement is called when production controlStatement is exited.
func (s *BasevbaListener) ExitControlStatement(ctx *ControlStatementContext) {}

// EnterControlStatementExceptMultilineIf is called when production controlStatementExceptMultilineIf is entered.
func (s *BasevbaListener) EnterControlStatementExceptMultilineIf(ctx *ControlStatementExceptMultilineIfContext) {
}

// ExitControlStatementExceptMultilineIf is called when production controlStatementExceptMultilineIf is exited.
func (s *BasevbaListener) ExitControlStatementExceptMultilineIf(ctx *ControlStatementExceptMultilineIfContext) {
}

// EnterCallStatement is called when production callStatement is entered.
func (s *BasevbaListener) EnterCallStatement(ctx *CallStatementContext) {}

// ExitCallStatement is called when production callStatement is exited.
func (s *BasevbaListener) ExitCallStatement(ctx *CallStatementContext) {}

// EnterWhileStatement is called when production whileStatement is entered.
func (s *BasevbaListener) EnterWhileStatement(ctx *WhileStatementContext) {}

// ExitWhileStatement is called when production whileStatement is exited.
func (s *BasevbaListener) ExitWhileStatement(ctx *WhileStatementContext) {}

// EnterForStatement is called when production forStatement is entered.
func (s *BasevbaListener) EnterForStatement(ctx *ForStatementContext) {}

// ExitForStatement is called when production forStatement is exited.
func (s *BasevbaListener) ExitForStatement(ctx *ForStatementContext) {}

// EnterForTail is called when production forTail is entered.
func (s *BasevbaListener) EnterForTail(ctx *ForTailContext) {}

// ExitForTail is called when production forTail is exited.
func (s *BasevbaListener) ExitForTail(ctx *ForTailContext) {}

// EnterForClause is called when production forClause is entered.
func (s *BasevbaListener) EnterForClause(ctx *ForClauseContext) {}

// ExitForClause is called when production forClause is exited.
func (s *BasevbaListener) ExitForClause(ctx *ForClauseContext) {}

// EnterStartValue is called when production startValue is entered.
func (s *BasevbaListener) EnterStartValue(ctx *StartValueContext) {}

// ExitStartValue is called when production startValue is exited.
func (s *BasevbaListener) ExitStartValue(ctx *StartValueContext) {}

// EnterEndValue is called when production endValue is entered.
func (s *BasevbaListener) EnterEndValue(ctx *EndValueContext) {}

// ExitEndValue is called when production endValue is exited.
func (s *BasevbaListener) ExitEndValue(ctx *EndValueContext) {}

// EnterStepClause is called when production stepClause is entered.
func (s *BasevbaListener) EnterStepClause(ctx *StepClauseContext) {}

// ExitStepClause is called when production stepClause is exited.
func (s *BasevbaListener) ExitStepClause(ctx *StepClauseContext) {}

// EnterStepIncrement is called when production stepIncrement is entered.
func (s *BasevbaListener) EnterStepIncrement(ctx *StepIncrementContext) {}

// ExitStepIncrement is called when production stepIncrement is exited.
func (s *BasevbaListener) ExitStepIncrement(ctx *StepIncrementContext) {}

// EnterForEachStatement is called when production forEachStatement is entered.
func (s *BasevbaListener) EnterForEachStatement(ctx *ForEachStatementContext) {}

// ExitForEachStatement is called when production forEachStatement is exited.
func (s *BasevbaListener) ExitForEachStatement(ctx *ForEachStatementContext) {}

// EnterForEachClause is called when production forEachClause is entered.
func (s *BasevbaListener) EnterForEachClause(ctx *ForEachClauseContext) {}

// ExitForEachClause is called when production forEachClause is exited.
func (s *BasevbaListener) ExitForEachClause(ctx *ForEachClauseContext) {}

// EnterCollection is called when production collection is entered.
func (s *BasevbaListener) EnterCollection(ctx *CollectionContext) {}

// ExitCollection is called when production collection is exited.
func (s *BasevbaListener) ExitCollection(ctx *CollectionContext) {}

// EnterExitForStatement is called when production exitForStatement is entered.
func (s *BasevbaListener) EnterExitForStatement(ctx *ExitForStatementContext) {}

// ExitExitForStatement is called when production exitForStatement is exited.
func (s *BasevbaListener) ExitExitForStatement(ctx *ExitForStatementContext) {}

// EnterDoStatement is called when production doStatement is entered.
func (s *BasevbaListener) EnterDoStatement(ctx *DoStatementContext) {}

// ExitDoStatement is called when production doStatement is exited.
func (s *BasevbaListener) ExitDoStatement(ctx *DoStatementContext) {}

// EnterConditionClause is called when production conditionClause is entered.
func (s *BasevbaListener) EnterConditionClause(ctx *ConditionClauseContext) {}

// ExitConditionClause is called when production conditionClause is exited.
func (s *BasevbaListener) ExitConditionClause(ctx *ConditionClauseContext) {}

// EnterWhileClause is called when production whileClause is entered.
func (s *BasevbaListener) EnterWhileClause(ctx *WhileClauseContext) {}

// ExitWhileClause is called when production whileClause is exited.
func (s *BasevbaListener) ExitWhileClause(ctx *WhileClauseContext) {}

// EnterUntilClause is called when production untilClause is entered.
func (s *BasevbaListener) EnterUntilClause(ctx *UntilClauseContext) {}

// ExitUntilClause is called when production untilClause is exited.
func (s *BasevbaListener) ExitUntilClause(ctx *UntilClauseContext) {}

// EnterExitDoStatement is called when production exitDoStatement is entered.
func (s *BasevbaListener) EnterExitDoStatement(ctx *ExitDoStatementContext) {}

// ExitExitDoStatement is called when production exitDoStatement is exited.
func (s *BasevbaListener) ExitExitDoStatement(ctx *ExitDoStatementContext) {}

// EnterIfStatement is called when production ifStatement is entered.
func (s *BasevbaListener) EnterIfStatement(ctx *IfStatementContext) {}

// ExitIfStatement is called when production ifStatement is exited.
func (s *BasevbaListener) ExitIfStatement(ctx *IfStatementContext) {}

// EnterElseIfBlock is called when production elseIfBlock is entered.
func (s *BasevbaListener) EnterElseIfBlock(ctx *ElseIfBlockContext) {}

// ExitElseIfBlock is called when production elseIfBlock is exited.
func (s *BasevbaListener) ExitElseIfBlock(ctx *ElseIfBlockContext) {}

// EnterElseBlock is called when production elseBlock is entered.
func (s *BasevbaListener) EnterElseBlock(ctx *ElseBlockContext) {}

// ExitElseBlock is called when production elseBlock is exited.
func (s *BasevbaListener) ExitElseBlock(ctx *ElseBlockContext) {}

// EnterSingleLineIfStatement is called when production singleLineIfStatement is entered.
func (s *BasevbaListener) EnterSingleLineIfStatement(ctx *SingleLineIfStatementContext) {}

// ExitSingleLineIfStatement is called when production singleLineIfStatement is exited.
func (s *BasevbaListener) ExitSingleLineIfStatement(ctx *SingleLineIfStatementContext) {}

// EnterIfWithNonEmptyThen is called when production ifWithNonEmptyThen is entered.
func (s *BasevbaListener) EnterIfWithNonEmptyThen(ctx *IfWithNonEmptyThenContext) {}

// ExitIfWithNonEmptyThen is called when production ifWithNonEmptyThen is exited.
func (s *BasevbaListener) ExitIfWithNonEmptyThen(ctx *IfWithNonEmptyThenContext) {}

// EnterIfWithEmptyThen is called when production ifWithEmptyThen is entered.
func (s *BasevbaListener) EnterIfWithEmptyThen(ctx *IfWithEmptyThenContext) {}

// ExitIfWithEmptyThen is called when production ifWithEmptyThen is exited.
func (s *BasevbaListener) ExitIfWithEmptyThen(ctx *IfWithEmptyThenContext) {}

// EnterSingleLineElseClause is called when production singleLineElseClause is entered.
func (s *BasevbaListener) EnterSingleLineElseClause(ctx *SingleLineElseClauseContext) {}

// ExitSingleLineElseClause is called when production singleLineElseClause is exited.
func (s *BasevbaListener) ExitSingleLineElseClause(ctx *SingleLineElseClauseContext) {}

// EnterListOrLabel is called when production listOrLabel is entered.
func (s *BasevbaListener) EnterListOrLabel(ctx *ListOrLabelContext) {}

// ExitListOrLabel is called when production listOrLabel is exited.
func (s *BasevbaListener) ExitListOrLabel(ctx *ListOrLabelContext) {}

// EnterSameLineStatement is called when production sameLineStatement is entered.
func (s *BasevbaListener) EnterSameLineStatement(ctx *SameLineStatementContext) {}

// ExitSameLineStatement is called when production sameLineStatement is exited.
func (s *BasevbaListener) ExitSameLineStatement(ctx *SameLineStatementContext) {}

// EnterSelectCaseStatement is called when production selectCaseStatement is entered.
func (s *BasevbaListener) EnterSelectCaseStatement(ctx *SelectCaseStatementContext) {}

// ExitSelectCaseStatement is called when production selectCaseStatement is exited.
func (s *BasevbaListener) ExitSelectCaseStatement(ctx *SelectCaseStatementContext) {}

// EnterCaseClause is called when production caseClause is entered.
func (s *BasevbaListener) EnterCaseClause(ctx *CaseClauseContext) {}

// ExitCaseClause is called when production caseClause is exited.
func (s *BasevbaListener) ExitCaseClause(ctx *CaseClauseContext) {}

// EnterCaseElseClause is called when production caseElseClause is entered.
func (s *BasevbaListener) EnterCaseElseClause(ctx *CaseElseClauseContext) {}

// ExitCaseElseClause is called when production caseElseClause is exited.
func (s *BasevbaListener) ExitCaseElseClause(ctx *CaseElseClauseContext) {}

// EnterRangeClause is called when production rangeClause is entered.
func (s *BasevbaListener) EnterRangeClause(ctx *RangeClauseContext) {}

// ExitRangeClause is called when production rangeClause is exited.
func (s *BasevbaListener) ExitRangeClause(ctx *RangeClauseContext) {}

// EnterSelectExpression is called when production selectExpression is entered.
func (s *BasevbaListener) EnterSelectExpression(ctx *SelectExpressionContext) {}

// ExitSelectExpression is called when production selectExpression is exited.
func (s *BasevbaListener) ExitSelectExpression(ctx *SelectExpressionContext) {}

// EnterComparisonOperator is called when production comparisonOperator is entered.
func (s *BasevbaListener) EnterComparisonOperator(ctx *ComparisonOperatorContext) {}

// ExitComparisonOperator is called when production comparisonOperator is exited.
func (s *BasevbaListener) ExitComparisonOperator(ctx *ComparisonOperatorContext) {}

// EnterStopStatement is called when production stopStatement is entered.
func (s *BasevbaListener) EnterStopStatement(ctx *StopStatementContext) {}

// ExitStopStatement is called when production stopStatement is exited.
func (s *BasevbaListener) ExitStopStatement(ctx *StopStatementContext) {}

// EnterGotoStatement is called when production gotoStatement is entered.
func (s *BasevbaListener) EnterGotoStatement(ctx *GotoStatementContext) {}

// ExitGotoStatement is called when production gotoStatement is exited.
func (s *BasevbaListener) ExitGotoStatement(ctx *GotoStatementContext) {}

// EnterOnGotoStatement is called when production onGotoStatement is entered.
func (s *BasevbaListener) EnterOnGotoStatement(ctx *OnGotoStatementContext) {}

// ExitOnGotoStatement is called when production onGotoStatement is exited.
func (s *BasevbaListener) ExitOnGotoStatement(ctx *OnGotoStatementContext) {}

// EnterGosubStatement is called when production gosubStatement is entered.
func (s *BasevbaListener) EnterGosubStatement(ctx *GosubStatementContext) {}

// ExitGosubStatement is called when production gosubStatement is exited.
func (s *BasevbaListener) ExitGosubStatement(ctx *GosubStatementContext) {}

// EnterReturnStatement is called when production returnStatement is entered.
func (s *BasevbaListener) EnterReturnStatement(ctx *ReturnStatementContext) {}

// ExitReturnStatement is called when production returnStatement is exited.
func (s *BasevbaListener) ExitReturnStatement(ctx *ReturnStatementContext) {}

// EnterOnGosubStatement is called when production onGosubStatement is entered.
func (s *BasevbaListener) EnterOnGosubStatement(ctx *OnGosubStatementContext) {}

// ExitOnGosubStatement is called when production onGosubStatement is exited.
func (s *BasevbaListener) ExitOnGosubStatement(ctx *OnGosubStatementContext) {}

// EnterExitSubStatement is called when production exitSubStatement is entered.
func (s *BasevbaListener) EnterExitSubStatement(ctx *ExitSubStatementContext) {}

// ExitExitSubStatement is called when production exitSubStatement is exited.
func (s *BasevbaListener) ExitExitSubStatement(ctx *ExitSubStatementContext) {}

// EnterExitFunctionStatement is called when production exitFunctionStatement is entered.
func (s *BasevbaListener) EnterExitFunctionStatement(ctx *ExitFunctionStatementContext) {}

// ExitExitFunctionStatement is called when production exitFunctionStatement is exited.
func (s *BasevbaListener) ExitExitFunctionStatement(ctx *ExitFunctionStatementContext) {}

// EnterExitPropertyStatement is called when production exitPropertyStatement is entered.
func (s *BasevbaListener) EnterExitPropertyStatement(ctx *ExitPropertyStatementContext) {}

// ExitExitPropertyStatement is called when production exitPropertyStatement is exited.
func (s *BasevbaListener) ExitExitPropertyStatement(ctx *ExitPropertyStatementContext) {}

// EnterRaiseeventStatement is called when production raiseeventStatement is entered.
func (s *BasevbaListener) EnterRaiseeventStatement(ctx *RaiseeventStatementContext) {}

// ExitRaiseeventStatement is called when production raiseeventStatement is exited.
func (s *BasevbaListener) ExitRaiseeventStatement(ctx *RaiseeventStatementContext) {}

// EnterEventArgumentList is called when production eventArgumentList is entered.
func (s *BasevbaListener) EnterEventArgumentList(ctx *EventArgumentListContext) {}

// ExitEventArgumentList is called when production eventArgumentList is exited.
func (s *BasevbaListener) ExitEventArgumentList(ctx *EventArgumentListContext) {}

// EnterEventArgument is called when production eventArgument is entered.
func (s *BasevbaListener) EnterEventArgument(ctx *EventArgumentContext) {}

// ExitEventArgument is called when production eventArgument is exited.
func (s *BasevbaListener) ExitEventArgument(ctx *EventArgumentContext) {}

// EnterWithStatement is called when production withStatement is entered.
func (s *BasevbaListener) EnterWithStatement(ctx *WithStatementContext) {}

// ExitWithStatement is called when production withStatement is exited.
func (s *BasevbaListener) ExitWithStatement(ctx *WithStatementContext) {}

// EnterEndStatement is called when production endStatement is entered.
func (s *BasevbaListener) EnterEndStatement(ctx *EndStatementContext) {}

// ExitEndStatement is called when production endStatement is exited.
func (s *BasevbaListener) ExitEndStatement(ctx *EndStatementContext) {}

// EnterDataManipulationStatement is called when production dataManipulationStatement is entered.
func (s *BasevbaListener) EnterDataManipulationStatement(ctx *DataManipulationStatementContext) {}

// ExitDataManipulationStatement is called when production dataManipulationStatement is exited.
func (s *BasevbaListener) ExitDataManipulationStatement(ctx *DataManipulationStatementContext) {}

// EnterStaticVariableDeclaration is called when production staticVariableDeclaration is entered.
func (s *BasevbaListener) EnterStaticVariableDeclaration(ctx *StaticVariableDeclarationContext) {}

// ExitStaticVariableDeclaration is called when production staticVariableDeclaration is exited.
func (s *BasevbaListener) ExitStaticVariableDeclaration(ctx *StaticVariableDeclarationContext) {}

// EnterRedimStatement is called when production redimStatement is entered.
func (s *BasevbaListener) EnterRedimStatement(ctx *RedimStatementContext) {}

// ExitRedimStatement is called when production redimStatement is exited.
func (s *BasevbaListener) ExitRedimStatement(ctx *RedimStatementContext) {}

// EnterRedimDeclarationList is called when production redimDeclarationList is entered.
func (s *BasevbaListener) EnterRedimDeclarationList(ctx *RedimDeclarationListContext) {}

// ExitRedimDeclarationList is called when production redimDeclarationList is exited.
func (s *BasevbaListener) ExitRedimDeclarationList(ctx *RedimDeclarationListContext) {}

// EnterRedimVariableDcl is called when production redimVariableDcl is entered.
func (s *BasevbaListener) EnterRedimVariableDcl(ctx *RedimVariableDclContext) {}

// ExitRedimVariableDcl is called when production redimVariableDcl is exited.
func (s *BasevbaListener) ExitRedimVariableDcl(ctx *RedimVariableDclContext) {}

// EnterRedimTypedVariableDcl is called when production redimTypedVariableDcl is entered.
func (s *BasevbaListener) EnterRedimTypedVariableDcl(ctx *RedimTypedVariableDclContext) {}

// ExitRedimTypedVariableDcl is called when production redimTypedVariableDcl is exited.
func (s *BasevbaListener) ExitRedimTypedVariableDcl(ctx *RedimTypedVariableDclContext) {}

// EnterRedimUntypedDcl is called when production redimUntypedDcl is entered.
func (s *BasevbaListener) EnterRedimUntypedDcl(ctx *RedimUntypedDclContext) {}

// ExitRedimUntypedDcl is called when production redimUntypedDcl is exited.
func (s *BasevbaListener) ExitRedimUntypedDcl(ctx *RedimUntypedDclContext) {}

// EnterWithExpressionDcl is called when production withExpressionDcl is entered.
func (s *BasevbaListener) EnterWithExpressionDcl(ctx *WithExpressionDclContext) {}

// ExitWithExpressionDcl is called when production withExpressionDcl is exited.
func (s *BasevbaListener) ExitWithExpressionDcl(ctx *WithExpressionDclContext) {}

// EnterMemberAccessExpressionDcl is called when production memberAccessExpressionDcl is entered.
func (s *BasevbaListener) EnterMemberAccessExpressionDcl(ctx *MemberAccessExpressionDclContext) {}

// ExitMemberAccessExpressionDcl is called when production memberAccessExpressionDcl is exited.
func (s *BasevbaListener) ExitMemberAccessExpressionDcl(ctx *MemberAccessExpressionDclContext) {}

// EnterDynamicArrayDim is called when production dynamicArrayDim is entered.
func (s *BasevbaListener) EnterDynamicArrayDim(ctx *DynamicArrayDimContext) {}

// ExitDynamicArrayDim is called when production dynamicArrayDim is exited.
func (s *BasevbaListener) ExitDynamicArrayDim(ctx *DynamicArrayDimContext) {}

// EnterDynamicBoundsList is called when production dynamicBoundsList is entered.
func (s *BasevbaListener) EnterDynamicBoundsList(ctx *DynamicBoundsListContext) {}

// ExitDynamicBoundsList is called when production dynamicBoundsList is exited.
func (s *BasevbaListener) ExitDynamicBoundsList(ctx *DynamicBoundsListContext) {}

// EnterDynamicDimSpec is called when production dynamicDimSpec is entered.
func (s *BasevbaListener) EnterDynamicDimSpec(ctx *DynamicDimSpecContext) {}

// ExitDynamicDimSpec is called when production dynamicDimSpec is exited.
func (s *BasevbaListener) ExitDynamicDimSpec(ctx *DynamicDimSpecContext) {}

// EnterDynamicLowerBound is called when production dynamicLowerBound is entered.
func (s *BasevbaListener) EnterDynamicLowerBound(ctx *DynamicLowerBoundContext) {}

// ExitDynamicLowerBound is called when production dynamicLowerBound is exited.
func (s *BasevbaListener) ExitDynamicLowerBound(ctx *DynamicLowerBoundContext) {}

// EnterDynamicUpperBound is called when production dynamicUpperBound is entered.
func (s *BasevbaListener) EnterDynamicUpperBound(ctx *DynamicUpperBoundContext) {}

// ExitDynamicUpperBound is called when production dynamicUpperBound is exited.
func (s *BasevbaListener) ExitDynamicUpperBound(ctx *DynamicUpperBoundContext) {}

// EnterDynamicArrayClause is called when production dynamicArrayClause is entered.
func (s *BasevbaListener) EnterDynamicArrayClause(ctx *DynamicArrayClauseContext) {}

// ExitDynamicArrayClause is called when production dynamicArrayClause is exited.
func (s *BasevbaListener) ExitDynamicArrayClause(ctx *DynamicArrayClauseContext) {}

// EnterEraseStatement is called when production eraseStatement is entered.
func (s *BasevbaListener) EnterEraseStatement(ctx *EraseStatementContext) {}

// ExitEraseStatement is called when production eraseStatement is exited.
func (s *BasevbaListener) ExitEraseStatement(ctx *EraseStatementContext) {}

// EnterEraseList is called when production eraseList is entered.
func (s *BasevbaListener) EnterEraseList(ctx *EraseListContext) {}

// ExitEraseList is called when production eraseList is exited.
func (s *BasevbaListener) ExitEraseList(ctx *EraseListContext) {}

// EnterEraseElement is called when production eraseElement is entered.
func (s *BasevbaListener) EnterEraseElement(ctx *EraseElementContext) {}

// ExitEraseElement is called when production eraseElement is exited.
func (s *BasevbaListener) ExitEraseElement(ctx *EraseElementContext) {}

// EnterMidStatement is called when production midStatement is entered.
func (s *BasevbaListener) EnterMidStatement(ctx *MidStatementContext) {}

// ExitMidStatement is called when production midStatement is exited.
func (s *BasevbaListener) ExitMidStatement(ctx *MidStatementContext) {}

// EnterModeSpecifier is called when production modeSpecifier is entered.
func (s *BasevbaListener) EnterModeSpecifier(ctx *ModeSpecifierContext) {}

// ExitModeSpecifier is called when production modeSpecifier is exited.
func (s *BasevbaListener) ExitModeSpecifier(ctx *ModeSpecifierContext) {}

// EnterStringArgument is called when production stringArgument is entered.
func (s *BasevbaListener) EnterStringArgument(ctx *StringArgumentContext) {}

// ExitStringArgument is called when production stringArgument is exited.
func (s *BasevbaListener) ExitStringArgument(ctx *StringArgumentContext) {}

// EnterStartMid is called when production startMid is entered.
func (s *BasevbaListener) EnterStartMid(ctx *StartMidContext) {}

// ExitStartMid is called when production startMid is exited.
func (s *BasevbaListener) ExitStartMid(ctx *StartMidContext) {}

// EnterLength is called when production length is entered.
func (s *BasevbaListener) EnterLength(ctx *LengthContext) {}

// ExitLength is called when production length is exited.
func (s *BasevbaListener) ExitLength(ctx *LengthContext) {}

// EnterLsetStatement is called when production lsetStatement is entered.
func (s *BasevbaListener) EnterLsetStatement(ctx *LsetStatementContext) {}

// ExitLsetStatement is called when production lsetStatement is exited.
func (s *BasevbaListener) ExitLsetStatement(ctx *LsetStatementContext) {}

// EnterRsetStatement is called when production rsetStatement is entered.
func (s *BasevbaListener) EnterRsetStatement(ctx *RsetStatementContext) {}

// ExitRsetStatement is called when production rsetStatement is exited.
func (s *BasevbaListener) ExitRsetStatement(ctx *RsetStatementContext) {}

// EnterLetStatement is called when production letStatement is entered.
func (s *BasevbaListener) EnterLetStatement(ctx *LetStatementContext) {}

// ExitLetStatement is called when production letStatement is exited.
func (s *BasevbaListener) ExitLetStatement(ctx *LetStatementContext) {}

// EnterSetStatement is called when production setStatement is entered.
func (s *BasevbaListener) EnterSetStatement(ctx *SetStatementContext) {}

// ExitSetStatement is called when production setStatement is exited.
func (s *BasevbaListener) ExitSetStatement(ctx *SetStatementContext) {}

// EnterErrorHandlingStatement is called when production errorHandlingStatement is entered.
func (s *BasevbaListener) EnterErrorHandlingStatement(ctx *ErrorHandlingStatementContext) {}

// ExitErrorHandlingStatement is called when production errorHandlingStatement is exited.
func (s *BasevbaListener) ExitErrorHandlingStatement(ctx *ErrorHandlingStatementContext) {}

// EnterOnErrorStatement is called when production onErrorStatement is entered.
func (s *BasevbaListener) EnterOnErrorStatement(ctx *OnErrorStatementContext) {}

// ExitOnErrorStatement is called when production onErrorStatement is exited.
func (s *BasevbaListener) ExitOnErrorStatement(ctx *OnErrorStatementContext) {}

// EnterErrorBehavior is called when production errorBehavior is entered.
func (s *BasevbaListener) EnterErrorBehavior(ctx *ErrorBehaviorContext) {}

// ExitErrorBehavior is called when production errorBehavior is exited.
func (s *BasevbaListener) ExitErrorBehavior(ctx *ErrorBehaviorContext) {}

// EnterResumeStatement is called when production resumeStatement is entered.
func (s *BasevbaListener) EnterResumeStatement(ctx *ResumeStatementContext) {}

// ExitResumeStatement is called when production resumeStatement is exited.
func (s *BasevbaListener) ExitResumeStatement(ctx *ResumeStatementContext) {}

// EnterErrorStatement is called when production errorStatement is entered.
func (s *BasevbaListener) EnterErrorStatement(ctx *ErrorStatementContext) {}

// ExitErrorStatement is called when production errorStatement is exited.
func (s *BasevbaListener) ExitErrorStatement(ctx *ErrorStatementContext) {}

// EnterErrorNumber is called when production errorNumber is entered.
func (s *BasevbaListener) EnterErrorNumber(ctx *ErrorNumberContext) {}

// ExitErrorNumber is called when production errorNumber is exited.
func (s *BasevbaListener) ExitErrorNumber(ctx *ErrorNumberContext) {}

// EnterFileStatement is called when production fileStatement is entered.
func (s *BasevbaListener) EnterFileStatement(ctx *FileStatementContext) {}

// ExitFileStatement is called when production fileStatement is exited.
func (s *BasevbaListener) ExitFileStatement(ctx *FileStatementContext) {}

// EnterOpenStatement is called when production openStatement is entered.
func (s *BasevbaListener) EnterOpenStatement(ctx *OpenStatementContext) {}

// ExitOpenStatement is called when production openStatement is exited.
func (s *BasevbaListener) ExitOpenStatement(ctx *OpenStatementContext) {}

// EnterPathName is called when production pathName is entered.
func (s *BasevbaListener) EnterPathName(ctx *PathNameContext) {}

// ExitPathName is called when production pathName is exited.
func (s *BasevbaListener) ExitPathName(ctx *PathNameContext) {}

// EnterModeClause is called when production modeClause is entered.
func (s *BasevbaListener) EnterModeClause(ctx *ModeClauseContext) {}

// ExitModeClause is called when production modeClause is exited.
func (s *BasevbaListener) ExitModeClause(ctx *ModeClauseContext) {}

// EnterModeOpt is called when production modeOpt is entered.
func (s *BasevbaListener) EnterModeOpt(ctx *ModeOptContext) {}

// ExitModeOpt is called when production modeOpt is exited.
func (s *BasevbaListener) ExitModeOpt(ctx *ModeOptContext) {}

// EnterAccessClause is called when production accessClause is entered.
func (s *BasevbaListener) EnterAccessClause(ctx *AccessClauseContext) {}

// ExitAccessClause is called when production accessClause is exited.
func (s *BasevbaListener) ExitAccessClause(ctx *AccessClauseContext) {}

// EnterAccess is called when production access is entered.
func (s *BasevbaListener) EnterAccess(ctx *AccessContext) {}

// ExitAccess is called when production access is exited.
func (s *BasevbaListener) ExitAccess(ctx *AccessContext) {}

// EnterLock is called when production lock is entered.
func (s *BasevbaListener) EnterLock(ctx *LockContext) {}

// ExitLock is called when production lock is exited.
func (s *BasevbaListener) ExitLock(ctx *LockContext) {}

// EnterLenClause is called when production lenClause is entered.
func (s *BasevbaListener) EnterLenClause(ctx *LenClauseContext) {}

// ExitLenClause is called when production lenClause is exited.
func (s *BasevbaListener) ExitLenClause(ctx *LenClauseContext) {}

// EnterRecLength is called when production recLength is entered.
func (s *BasevbaListener) EnterRecLength(ctx *RecLengthContext) {}

// ExitRecLength is called when production recLength is exited.
func (s *BasevbaListener) ExitRecLength(ctx *RecLengthContext) {}

// EnterFileNumber is called when production fileNumber is entered.
func (s *BasevbaListener) EnterFileNumber(ctx *FileNumberContext) {}

// ExitFileNumber is called when production fileNumber is exited.
func (s *BasevbaListener) ExitFileNumber(ctx *FileNumberContext) {}

// EnterMarkedFileNumber is called when production markedFileNumber is entered.
func (s *BasevbaListener) EnterMarkedFileNumber(ctx *MarkedFileNumberContext) {}

// ExitMarkedFileNumber is called when production markedFileNumber is exited.
func (s *BasevbaListener) ExitMarkedFileNumber(ctx *MarkedFileNumberContext) {}

// EnterUnmarkedFileNumber is called when production unmarkedFileNumber is entered.
func (s *BasevbaListener) EnterUnmarkedFileNumber(ctx *UnmarkedFileNumberContext) {}

// ExitUnmarkedFileNumber is called when production unmarkedFileNumber is exited.
func (s *BasevbaListener) ExitUnmarkedFileNumber(ctx *UnmarkedFileNumberContext) {}

// EnterCloseStatement is called when production closeStatement is entered.
func (s *BasevbaListener) EnterCloseStatement(ctx *CloseStatementContext) {}

// ExitCloseStatement is called when production closeStatement is exited.
func (s *BasevbaListener) ExitCloseStatement(ctx *CloseStatementContext) {}

// EnterFileNumberList is called when production fileNumberList is entered.
func (s *BasevbaListener) EnterFileNumberList(ctx *FileNumberListContext) {}

// ExitFileNumberList is called when production fileNumberList is exited.
func (s *BasevbaListener) ExitFileNumberList(ctx *FileNumberListContext) {}

// EnterSeekStatement is called when production seekStatement is entered.
func (s *BasevbaListener) EnterSeekStatement(ctx *SeekStatementContext) {}

// ExitSeekStatement is called when production seekStatement is exited.
func (s *BasevbaListener) ExitSeekStatement(ctx *SeekStatementContext) {}

// EnterPosition is called when production position is entered.
func (s *BasevbaListener) EnterPosition(ctx *PositionContext) {}

// ExitPosition is called when production position is exited.
func (s *BasevbaListener) ExitPosition(ctx *PositionContext) {}

// EnterLockStatement is called when production lockStatement is entered.
func (s *BasevbaListener) EnterLockStatement(ctx *LockStatementContext) {}

// ExitLockStatement is called when production lockStatement is exited.
func (s *BasevbaListener) ExitLockStatement(ctx *LockStatementContext) {}

// EnterRecordRange is called when production recordRange is entered.
func (s *BasevbaListener) EnterRecordRange(ctx *RecordRangeContext) {}

// ExitRecordRange is called when production recordRange is exited.
func (s *BasevbaListener) ExitRecordRange(ctx *RecordRangeContext) {}

// EnterStartRecordNumber is called when production startRecordNumber is entered.
func (s *BasevbaListener) EnterStartRecordNumber(ctx *StartRecordNumberContext) {}

// ExitStartRecordNumber is called when production startRecordNumber is exited.
func (s *BasevbaListener) ExitStartRecordNumber(ctx *StartRecordNumberContext) {}

// EnterEndRecordNumber is called when production endRecordNumber is entered.
func (s *BasevbaListener) EnterEndRecordNumber(ctx *EndRecordNumberContext) {}

// ExitEndRecordNumber is called when production endRecordNumber is exited.
func (s *BasevbaListener) ExitEndRecordNumber(ctx *EndRecordNumberContext) {}

// EnterUnlockStatement is called when production unlockStatement is entered.
func (s *BasevbaListener) EnterUnlockStatement(ctx *UnlockStatementContext) {}

// ExitUnlockStatement is called when production unlockStatement is exited.
func (s *BasevbaListener) ExitUnlockStatement(ctx *UnlockStatementContext) {}

// EnterLineInputStatement is called when production lineInputStatement is entered.
func (s *BasevbaListener) EnterLineInputStatement(ctx *LineInputStatementContext) {}

// ExitLineInputStatement is called when production lineInputStatement is exited.
func (s *BasevbaListener) ExitLineInputStatement(ctx *LineInputStatementContext) {}

// EnterVariableName is called when production variableName is entered.
func (s *BasevbaListener) EnterVariableName(ctx *VariableNameContext) {}

// ExitVariableName is called when production variableName is exited.
func (s *BasevbaListener) ExitVariableName(ctx *VariableNameContext) {}

// EnterWidthStatement is called when production widthStatement is entered.
func (s *BasevbaListener) EnterWidthStatement(ctx *WidthStatementContext) {}

// ExitWidthStatement is called when production widthStatement is exited.
func (s *BasevbaListener) ExitWidthStatement(ctx *WidthStatementContext) {}

// EnterLineWidth is called when production lineWidth is entered.
func (s *BasevbaListener) EnterLineWidth(ctx *LineWidthContext) {}

// ExitLineWidth is called when production lineWidth is exited.
func (s *BasevbaListener) ExitLineWidth(ctx *LineWidthContext) {}

// EnterPrintStatement is called when production printStatement is entered.
func (s *BasevbaListener) EnterPrintStatement(ctx *PrintStatementContext) {}

// ExitPrintStatement is called when production printStatement is exited.
func (s *BasevbaListener) ExitPrintStatement(ctx *PrintStatementContext) {}

// EnterOutputList is called when production outputList is entered.
func (s *BasevbaListener) EnterOutputList(ctx *OutputListContext) {}

// ExitOutputList is called when production outputList is exited.
func (s *BasevbaListener) ExitOutputList(ctx *OutputListContext) {}

// EnterOutputItem is called when production outputItem is entered.
func (s *BasevbaListener) EnterOutputItem(ctx *OutputItemContext) {}

// ExitOutputItem is called when production outputItem is exited.
func (s *BasevbaListener) ExitOutputItem(ctx *OutputItemContext) {}

// EnterOutputClause is called when production outputClause is entered.
func (s *BasevbaListener) EnterOutputClause(ctx *OutputClauseContext) {}

// ExitOutputClause is called when production outputClause is exited.
func (s *BasevbaListener) ExitOutputClause(ctx *OutputClauseContext) {}

// EnterCharPosition is called when production charPosition is entered.
func (s *BasevbaListener) EnterCharPosition(ctx *CharPositionContext) {}

// ExitCharPosition is called when production charPosition is exited.
func (s *BasevbaListener) ExitCharPosition(ctx *CharPositionContext) {}

// EnterOutputExpression is called when production outputExpression is entered.
func (s *BasevbaListener) EnterOutputExpression(ctx *OutputExpressionContext) {}

// ExitOutputExpression is called when production outputExpression is exited.
func (s *BasevbaListener) ExitOutputExpression(ctx *OutputExpressionContext) {}

// EnterSpcClause is called when production spcClause is entered.
func (s *BasevbaListener) EnterSpcClause(ctx *SpcClauseContext) {}

// ExitSpcClause is called when production spcClause is exited.
func (s *BasevbaListener) ExitSpcClause(ctx *SpcClauseContext) {}

// EnterSpcNumber is called when production spcNumber is entered.
func (s *BasevbaListener) EnterSpcNumber(ctx *SpcNumberContext) {}

// ExitSpcNumber is called when production spcNumber is exited.
func (s *BasevbaListener) ExitSpcNumber(ctx *SpcNumberContext) {}

// EnterTabClause is called when production tabClause is entered.
func (s *BasevbaListener) EnterTabClause(ctx *TabClauseContext) {}

// ExitTabClause is called when production tabClause is exited.
func (s *BasevbaListener) ExitTabClause(ctx *TabClauseContext) {}

// EnterTabNumber is called when production tabNumber is entered.
func (s *BasevbaListener) EnterTabNumber(ctx *TabNumberContext) {}

// ExitTabNumber is called when production tabNumber is exited.
func (s *BasevbaListener) ExitTabNumber(ctx *TabNumberContext) {}

// EnterWriteStatement is called when production writeStatement is entered.
func (s *BasevbaListener) EnterWriteStatement(ctx *WriteStatementContext) {}

// ExitWriteStatement is called when production writeStatement is exited.
func (s *BasevbaListener) ExitWriteStatement(ctx *WriteStatementContext) {}

// EnterInputStatement is called when production inputStatement is entered.
func (s *BasevbaListener) EnterInputStatement(ctx *InputStatementContext) {}

// ExitInputStatement is called when production inputStatement is exited.
func (s *BasevbaListener) ExitInputStatement(ctx *InputStatementContext) {}

// EnterInputList is called when production inputList is entered.
func (s *BasevbaListener) EnterInputList(ctx *InputListContext) {}

// ExitInputList is called when production inputList is exited.
func (s *BasevbaListener) ExitInputList(ctx *InputListContext) {}

// EnterInputVariable is called when production inputVariable is entered.
func (s *BasevbaListener) EnterInputVariable(ctx *InputVariableContext) {}

// ExitInputVariable is called when production inputVariable is exited.
func (s *BasevbaListener) ExitInputVariable(ctx *InputVariableContext) {}

// EnterPutStatement is called when production putStatement is entered.
func (s *BasevbaListener) EnterPutStatement(ctx *PutStatementContext) {}

// ExitPutStatement is called when production putStatement is exited.
func (s *BasevbaListener) ExitPutStatement(ctx *PutStatementContext) {}

// EnterRecordNumber is called when production recordNumber is entered.
func (s *BasevbaListener) EnterRecordNumber(ctx *RecordNumberContext) {}

// ExitRecordNumber is called when production recordNumber is exited.
func (s *BasevbaListener) ExitRecordNumber(ctx *RecordNumberContext) {}

// EnterData is called when production data is entered.
func (s *BasevbaListener) EnterData(ctx *DataContext) {}

// ExitData is called when production data is exited.
func (s *BasevbaListener) ExitData(ctx *DataContext) {}

// EnterGetStatement is called when production getStatement is entered.
func (s *BasevbaListener) EnterGetStatement(ctx *GetStatementContext) {}

// ExitGetStatement is called when production getStatement is exited.
func (s *BasevbaListener) ExitGetStatement(ctx *GetStatementContext) {}

// EnterVariable is called when production variable is entered.
func (s *BasevbaListener) EnterVariable(ctx *VariableContext) {}

// ExitVariable is called when production variable is exited.
func (s *BasevbaListener) ExitVariable(ctx *VariableContext) {}

// EnterAttributeStatement is called when production attributeStatement is entered.
func (s *BasevbaListener) EnterAttributeStatement(ctx *AttributeStatementContext) {}

// ExitAttributeStatement is called when production attributeStatement is exited.
func (s *BasevbaListener) ExitAttributeStatement(ctx *AttributeStatementContext) {}

// EnterAttributeDescName is called when production attributeDescName is entered.
func (s *BasevbaListener) EnterAttributeDescName(ctx *AttributeDescNameContext) {}

// ExitAttributeDescName is called when production attributeDescName is exited.
func (s *BasevbaListener) ExitAttributeDescName(ctx *AttributeDescNameContext) {}

// EnterAttributeUsrName is called when production attributeUsrName is entered.
func (s *BasevbaListener) EnterAttributeUsrName(ctx *AttributeUsrNameContext) {}

// ExitAttributeUsrName is called when production attributeUsrName is exited.
func (s *BasevbaListener) ExitAttributeUsrName(ctx *AttributeUsrNameContext) {}

// EnterDebugStatement is called when production debugStatement is entered.
func (s *BasevbaListener) EnterDebugStatement(ctx *DebugStatementContext) {}

// ExitDebugStatement is called when production debugStatement is exited.
func (s *BasevbaListener) ExitDebugStatement(ctx *DebugStatementContext) {}

// EnterDebugArgs is called when production debugArgs is entered.
func (s *BasevbaListener) EnterDebugArgs(ctx *DebugArgsContext) {}

// ExitDebugArgs is called when production debugArgs is exited.
func (s *BasevbaListener) ExitDebugArgs(ctx *DebugArgsContext) {}

// EnterDebugSep is called when production debugSep is entered.
func (s *BasevbaListener) EnterDebugSep(ctx *DebugSepContext) {}

// ExitDebugSep is called when production debugSep is exited.
func (s *BasevbaListener) ExitDebugSep(ctx *DebugSepContext) {}

// EnterExpression is called when production expression is entered.
func (s *BasevbaListener) EnterExpression(ctx *ExpressionContext) {}

// ExitExpression is called when production expression is exited.
func (s *BasevbaListener) ExitExpression(ctx *ExpressionContext) {}

// EnterLExpression is called when production lExpression is entered.
func (s *BasevbaListener) EnterLExpression(ctx *LExpressionContext) {}

// ExitLExpression is called when production lExpression is exited.
func (s *BasevbaListener) ExitLExpression(ctx *LExpressionContext) {}

// EnterLiteralExpression is called when production literalExpression is entered.
func (s *BasevbaListener) EnterLiteralExpression(ctx *LiteralExpressionContext) {}

// ExitLiteralExpression is called when production literalExpression is exited.
func (s *BasevbaListener) ExitLiteralExpression(ctx *LiteralExpressionContext) {}

// EnterParenthesizedExpression is called when production parenthesizedExpression is entered.
func (s *BasevbaListener) EnterParenthesizedExpression(ctx *ParenthesizedExpressionContext) {}

// ExitParenthesizedExpression is called when production parenthesizedExpression is exited.
func (s *BasevbaListener) ExitParenthesizedExpression(ctx *ParenthesizedExpressionContext) {}

// EnterTypeofIsExpression is called when production typeofIsExpression is entered.
func (s *BasevbaListener) EnterTypeofIsExpression(ctx *TypeofIsExpressionContext) {}

// ExitTypeofIsExpression is called when production typeofIsExpression is exited.
func (s *BasevbaListener) ExitTypeofIsExpression(ctx *TypeofIsExpressionContext) {}

// EnterNewExpress is called when production newExpress is entered.
func (s *BasevbaListener) EnterNewExpress(ctx *NewExpressContext) {}

// ExitNewExpress is called when production newExpress is exited.
func (s *BasevbaListener) ExitNewExpress(ctx *NewExpressContext) {}

// EnterNotOperatorExpression is called when production notOperatorExpression is entered.
func (s *BasevbaListener) EnterNotOperatorExpression(ctx *NotOperatorExpressionContext) {}

// ExitNotOperatorExpression is called when production notOperatorExpression is exited.
func (s *BasevbaListener) ExitNotOperatorExpression(ctx *NotOperatorExpressionContext) {}

// EnterUnaryMinusExpression is called when production unaryMinusExpression is entered.
func (s *BasevbaListener) EnterUnaryMinusExpression(ctx *UnaryMinusExpressionContext) {}

// ExitUnaryMinusExpression is called when production unaryMinusExpression is exited.
func (s *BasevbaListener) ExitUnaryMinusExpression(ctx *UnaryMinusExpressionContext) {}

// EnterSimpleNameExpression is called when production simpleNameExpression is entered.
func (s *BasevbaListener) EnterSimpleNameExpression(ctx *SimpleNameExpressionContext) {}

// ExitSimpleNameExpression is called when production simpleNameExpression is exited.
func (s *BasevbaListener) ExitSimpleNameExpression(ctx *SimpleNameExpressionContext) {}

// EnterInstanceExpression is called when production instanceExpression is entered.
func (s *BasevbaListener) EnterInstanceExpression(ctx *InstanceExpressionContext) {}

// ExitInstanceExpression is called when production instanceExpression is exited.
func (s *BasevbaListener) ExitInstanceExpression(ctx *InstanceExpressionContext) {}

// EnterMemberAccessExpression is called when production memberAccessExpression is entered.
func (s *BasevbaListener) EnterMemberAccessExpression(ctx *MemberAccessExpressionContext) {}

// ExitMemberAccessExpression is called when production memberAccessExpression is exited.
func (s *BasevbaListener) ExitMemberAccessExpression(ctx *MemberAccessExpressionContext) {}

// EnterIndexExpression is called when production indexExpression is entered.
func (s *BasevbaListener) EnterIndexExpression(ctx *IndexExpressionContext) {}

// ExitIndexExpression is called when production indexExpression is exited.
func (s *BasevbaListener) ExitIndexExpression(ctx *IndexExpressionContext) {}

// EnterArgumentList is called when production argumentList is entered.
func (s *BasevbaListener) EnterArgumentList(ctx *ArgumentListContext) {}

// ExitArgumentList is called when production argumentList is exited.
func (s *BasevbaListener) ExitArgumentList(ctx *ArgumentListContext) {}

// EnterPositionalOrNamedArgumentList is called when production positionalOrNamedArgumentList is entered.
func (s *BasevbaListener) EnterPositionalOrNamedArgumentList(ctx *PositionalOrNamedArgumentListContext) {
}

// ExitPositionalOrNamedArgumentList is called when production positionalOrNamedArgumentList is exited.
func (s *BasevbaListener) ExitPositionalOrNamedArgumentList(ctx *PositionalOrNamedArgumentListContext) {
}

// EnterNamedArgumentList is called when production namedArgumentList is entered.
func (s *BasevbaListener) EnterNamedArgumentList(ctx *NamedArgumentListContext) {}

// ExitNamedArgumentList is called when production namedArgumentList is exited.
func (s *BasevbaListener) ExitNamedArgumentList(ctx *NamedArgumentListContext) {}

// EnterNamedArgument is called when production namedArgument is entered.
func (s *BasevbaListener) EnterNamedArgument(ctx *NamedArgumentContext) {}

// ExitNamedArgument is called when production namedArgument is exited.
func (s *BasevbaListener) ExitNamedArgument(ctx *NamedArgumentContext) {}

// EnterArgumentExpression is called when production argumentExpression is entered.
func (s *BasevbaListener) EnterArgumentExpression(ctx *ArgumentExpressionContext) {}

// ExitArgumentExpression is called when production argumentExpression is exited.
func (s *BasevbaListener) ExitArgumentExpression(ctx *ArgumentExpressionContext) {}

// EnterDictionaryAccessExpression is called when production dictionaryAccessExpression is entered.
func (s *BasevbaListener) EnterDictionaryAccessExpression(ctx *DictionaryAccessExpressionContext) {}

// ExitDictionaryAccessExpression is called when production dictionaryAccessExpression is exited.
func (s *BasevbaListener) ExitDictionaryAccessExpression(ctx *DictionaryAccessExpressionContext) {}

// EnterWithExpression is called when production withExpression is entered.
func (s *BasevbaListener) EnterWithExpression(ctx *WithExpressionContext) {}

// ExitWithExpression is called when production withExpression is exited.
func (s *BasevbaListener) ExitWithExpression(ctx *WithExpressionContext) {}

// EnterWithMemberAccessExpression is called when production withMemberAccessExpression is entered.
func (s *BasevbaListener) EnterWithMemberAccessExpression(ctx *WithMemberAccessExpressionContext) {}

// ExitWithMemberAccessExpression is called when production withMemberAccessExpression is exited.
func (s *BasevbaListener) ExitWithMemberAccessExpression(ctx *WithMemberAccessExpressionContext) {}

// EnterWithDictionaryAccessExpression is called when production withDictionaryAccessExpression is entered.
func (s *BasevbaListener) EnterWithDictionaryAccessExpression(ctx *WithDictionaryAccessExpressionContext) {
}

// ExitWithDictionaryAccessExpression is called when production withDictionaryAccessExpression is exited.
func (s *BasevbaListener) ExitWithDictionaryAccessExpression(ctx *WithDictionaryAccessExpressionContext) {
}

// EnterConstantExpression is called when production constantExpression is entered.
func (s *BasevbaListener) EnterConstantExpression(ctx *ConstantExpressionContext) {}

// ExitConstantExpression is called when production constantExpression is exited.
func (s *BasevbaListener) ExitConstantExpression(ctx *ConstantExpressionContext) {}

// EnterCcExpression is called when production ccExpression is entered.
func (s *BasevbaListener) EnterCcExpression(ctx *CcExpressionContext) {}

// ExitCcExpression is called when production ccExpression is exited.
func (s *BasevbaListener) ExitCcExpression(ctx *CcExpressionContext) {}

// EnterBooleanExpression is called when production booleanExpression is entered.
func (s *BasevbaListener) EnterBooleanExpression(ctx *BooleanExpressionContext) {}

// ExitBooleanExpression is called when production booleanExpression is exited.
func (s *BasevbaListener) ExitBooleanExpression(ctx *BooleanExpressionContext) {}

// EnterIntegerExpression is called when production integerExpression is entered.
func (s *BasevbaListener) EnterIntegerExpression(ctx *IntegerExpressionContext) {}

// ExitIntegerExpression is called when production integerExpression is exited.
func (s *BasevbaListener) ExitIntegerExpression(ctx *IntegerExpressionContext) {}

// EnterVariableExpression is called when production variableExpression is entered.
func (s *BasevbaListener) EnterVariableExpression(ctx *VariableExpressionContext) {}

// ExitVariableExpression is called when production variableExpression is exited.
func (s *BasevbaListener) ExitVariableExpression(ctx *VariableExpressionContext) {}

// EnterBoundVariableExpression is called when production boundVariableExpression is entered.
func (s *BasevbaListener) EnterBoundVariableExpression(ctx *BoundVariableExpressionContext) {}

// ExitBoundVariableExpression is called when production boundVariableExpression is exited.
func (s *BasevbaListener) ExitBoundVariableExpression(ctx *BoundVariableExpressionContext) {}

// EnterTypeExpression is called when production typeExpression is entered.
func (s *BasevbaListener) EnterTypeExpression(ctx *TypeExpressionContext) {}

// ExitTypeExpression is called when production typeExpression is exited.
func (s *BasevbaListener) ExitTypeExpression(ctx *TypeExpressionContext) {}

// EnterDefinedTypeExpression is called when production definedTypeExpression is entered.
func (s *BasevbaListener) EnterDefinedTypeExpression(ctx *DefinedTypeExpressionContext) {}

// ExitDefinedTypeExpression is called when production definedTypeExpression is exited.
func (s *BasevbaListener) ExitDefinedTypeExpression(ctx *DefinedTypeExpressionContext) {}

// EnterAddressofExpression is called when production addressofExpression is entered.
func (s *BasevbaListener) EnterAddressofExpression(ctx *AddressofExpressionContext) {}

// ExitAddressofExpression is called when production addressofExpression is exited.
func (s *BasevbaListener) ExitAddressofExpression(ctx *AddressofExpressionContext) {}

// EnterProcedurePointerExpression is called when production procedurePointerExpression is entered.
func (s *BasevbaListener) EnterProcedurePointerExpression(ctx *ProcedurePointerExpressionContext) {}

// ExitProcedurePointerExpression is called when production procedurePointerExpression is exited.
func (s *BasevbaListener) ExitProcedurePointerExpression(ctx *ProcedurePointerExpressionContext) {}

// EnterWsc is called when production wsc is entered.
func (s *BasevbaListener) EnterWsc(ctx *WscContext) {}

// ExitWsc is called when production wsc is exited.
func (s *BasevbaListener) ExitWsc(ctx *WscContext) {}

// EnterEndOfLine is called when production endOfLine is entered.
func (s *BasevbaListener) EnterEndOfLine(ctx *EndOfLineContext) {}

// ExitEndOfLine is called when production endOfLine is exited.
func (s *BasevbaListener) ExitEndOfLine(ctx *EndOfLineContext) {}

// EnterUnexpectedEndOfLine is called when production unexpectedEndOfLine is entered.
func (s *BasevbaListener) EnterUnexpectedEndOfLine(ctx *UnexpectedEndOfLineContext) {}

// ExitUnexpectedEndOfLine is called when production unexpectedEndOfLine is exited.
func (s *BasevbaListener) ExitUnexpectedEndOfLine(ctx *UnexpectedEndOfLineContext) {}

// EnterWscu is called when production wscu is entered.
func (s *BasevbaListener) EnterWscu(ctx *WscuContext) {}

// ExitWscu is called when production wscu is exited.
func (s *BasevbaListener) ExitWscu(ctx *WscuContext) {}

// EnterEndOfLineNoWs is called when production endOfLineNoWs is entered.
func (s *BasevbaListener) EnterEndOfLineNoWs(ctx *EndOfLineNoWsContext) {}

// ExitEndOfLineNoWs is called when production endOfLineNoWs is exited.
func (s *BasevbaListener) ExitEndOfLineNoWs(ctx *EndOfLineNoWsContext) {}

// EnterEndOfStatement is called when production endOfStatement is entered.
func (s *BasevbaListener) EnterEndOfStatement(ctx *EndOfStatementContext) {}

// ExitEndOfStatement is called when production endOfStatement is exited.
func (s *BasevbaListener) ExitEndOfStatement(ctx *EndOfStatementContext) {}

// EnterEndOfStatementNoWs is called when production endOfStatementNoWs is entered.
func (s *BasevbaListener) EnterEndOfStatementNoWs(ctx *EndOfStatementNoWsContext) {}

// ExitEndOfStatementNoWs is called when production endOfStatementNoWs is exited.
func (s *BasevbaListener) ExitEndOfStatementNoWs(ctx *EndOfStatementNoWsContext) {}

// EnterCommentBody is called when production commentBody is entered.
func (s *BasevbaListener) EnterCommentBody(ctx *CommentBodyContext) {}

// ExitCommentBody is called when production commentBody is exited.
func (s *BasevbaListener) ExitCommentBody(ctx *CommentBodyContext) {}

// EnterReservedIdentifier is called when production reservedIdentifier is entered.
func (s *BasevbaListener) EnterReservedIdentifier(ctx *ReservedIdentifierContext) {}

// ExitReservedIdentifier is called when production reservedIdentifier is exited.
func (s *BasevbaListener) ExitReservedIdentifier(ctx *ReservedIdentifierContext) {}

// EnterAmbiguousIdentifier is called when production ambiguousIdentifier is entered.
func (s *BasevbaListener) EnterAmbiguousIdentifier(ctx *AmbiguousIdentifierContext) {}

// ExitAmbiguousIdentifier is called when production ambiguousIdentifier is exited.
func (s *BasevbaListener) ExitAmbiguousIdentifier(ctx *AmbiguousIdentifierContext) {}

// EnterStatementKeyword is called when production statementKeyword is entered.
func (s *BasevbaListener) EnterStatementKeyword(ctx *StatementKeywordContext) {}

// ExitStatementKeyword is called when production statementKeyword is exited.
func (s *BasevbaListener) ExitStatementKeyword(ctx *StatementKeywordContext) {}

// EnterRemKeyword is called when production remKeyword is entered.
func (s *BasevbaListener) EnterRemKeyword(ctx *RemKeywordContext) {}

// ExitRemKeyword is called when production remKeyword is exited.
func (s *BasevbaListener) ExitRemKeyword(ctx *RemKeywordContext) {}

// EnterMarkerKeyword is called when production markerKeyword is entered.
func (s *BasevbaListener) EnterMarkerKeyword(ctx *MarkerKeywordContext) {}

// ExitMarkerKeyword is called when production markerKeyword is exited.
func (s *BasevbaListener) ExitMarkerKeyword(ctx *MarkerKeywordContext) {}

// EnterOperatorIdentifier is called when production operatorIdentifier is entered.
func (s *BasevbaListener) EnterOperatorIdentifier(ctx *OperatorIdentifierContext) {}

// ExitOperatorIdentifier is called when production operatorIdentifier is exited.
func (s *BasevbaListener) ExitOperatorIdentifier(ctx *OperatorIdentifierContext) {}

// EnterReservedName is called when production reservedName is entered.
func (s *BasevbaListener) EnterReservedName(ctx *ReservedNameContext) {}

// ExitReservedName is called when production reservedName is exited.
func (s *BasevbaListener) ExitReservedName(ctx *ReservedNameContext) {}

// EnterSpecialForm is called when production specialForm is entered.
func (s *BasevbaListener) EnterSpecialForm(ctx *SpecialFormContext) {}

// ExitSpecialForm is called when production specialForm is exited.
func (s *BasevbaListener) ExitSpecialForm(ctx *SpecialFormContext) {}

// EnterReservedTypeIdentifier is called when production reservedTypeIdentifier is entered.
func (s *BasevbaListener) EnterReservedTypeIdentifier(ctx *ReservedTypeIdentifierContext) {}

// ExitReservedTypeIdentifier is called when production reservedTypeIdentifier is exited.
func (s *BasevbaListener) ExitReservedTypeIdentifier(ctx *ReservedTypeIdentifierContext) {}

// EnterReservedTypeIdentifierB is called when production reservedTypeIdentifierB is entered.
func (s *BasevbaListener) EnterReservedTypeIdentifierB(ctx *ReservedTypeIdentifierBContext) {}

// ExitReservedTypeIdentifierB is called when production reservedTypeIdentifierB is exited.
func (s *BasevbaListener) ExitReservedTypeIdentifierB(ctx *ReservedTypeIdentifierBContext) {}

// EnterLiteralIdentifier is called when production literalIdentifier is entered.
func (s *BasevbaListener) EnterLiteralIdentifier(ctx *LiteralIdentifierContext) {}

// ExitLiteralIdentifier is called when production literalIdentifier is exited.
func (s *BasevbaListener) ExitLiteralIdentifier(ctx *LiteralIdentifierContext) {}

// EnterBooleanLiteralIdentifier is called when production booleanLiteralIdentifier is entered.
func (s *BasevbaListener) EnterBooleanLiteralIdentifier(ctx *BooleanLiteralIdentifierContext) {}

// ExitBooleanLiteralIdentifier is called when production booleanLiteralIdentifier is exited.
func (s *BasevbaListener) ExitBooleanLiteralIdentifier(ctx *BooleanLiteralIdentifierContext) {}

// EnterObjectLiteralIdentifier is called when production objectLiteralIdentifier is entered.
func (s *BasevbaListener) EnterObjectLiteralIdentifier(ctx *ObjectLiteralIdentifierContext) {}

// ExitObjectLiteralIdentifier is called when production objectLiteralIdentifier is exited.
func (s *BasevbaListener) ExitObjectLiteralIdentifier(ctx *ObjectLiteralIdentifierContext) {}

// EnterVariantLiteralIdentifier is called when production variantLiteralIdentifier is entered.
func (s *BasevbaListener) EnterVariantLiteralIdentifier(ctx *VariantLiteralIdentifierContext) {}

// ExitVariantLiteralIdentifier is called when production variantLiteralIdentifier is exited.
func (s *BasevbaListener) ExitVariantLiteralIdentifier(ctx *VariantLiteralIdentifierContext) {}

// EnterReservedForImplementationUse is called when production reservedForImplementationUse is entered.
func (s *BasevbaListener) EnterReservedForImplementationUse(ctx *ReservedForImplementationUseContext) {
}

// ExitReservedForImplementationUse is called when production reservedForImplementationUse is exited.
func (s *BasevbaListener) ExitReservedForImplementationUse(ctx *ReservedForImplementationUseContext) {
}

// EnterFutureReserved is called when production futureReserved is entered.
func (s *BasevbaListener) EnterFutureReserved(ctx *FutureReservedContext) {}

// ExitFutureReserved is called when production futureReserved is exited.
func (s *BasevbaListener) ExitFutureReserved(ctx *FutureReservedContext) {}

// EnterBuiltinType is called when production builtinType is entered.
func (s *BasevbaListener) EnterBuiltinType(ctx *BuiltinTypeContext) {}

// ExitBuiltinType is called when production builtinType is exited.
func (s *BasevbaListener) ExitBuiltinType(ctx *BuiltinTypeContext) {}

// EnterTypedName is called when production typedName is entered.
func (s *BasevbaListener) EnterTypedName(ctx *TypedNameContext) {}

// ExitTypedName is called when production typedName is exited.
func (s *BasevbaListener) ExitTypedName(ctx *TypedNameContext) {}

// EnterTypeSuffix is called when production typeSuffix is entered.
func (s *BasevbaListener) EnterTypeSuffix(ctx *TypeSuffixContext) {}

// ExitTypeSuffix is called when production typeSuffix is exited.
func (s *BasevbaListener) ExitTypeSuffix(ctx *TypeSuffixContext) {}

// EnterAmbiguousKeyword is called when production ambiguousKeyword is entered.
func (s *BasevbaListener) EnterAmbiguousKeyword(ctx *AmbiguousKeywordContext) {}

// ExitAmbiguousKeyword is called when production ambiguousKeyword is exited.
func (s *BasevbaListener) ExitAmbiguousKeyword(ctx *AmbiguousKeywordContext) {}

// EnterAnyOperator is called when production anyOperator is entered.
func (s *BasevbaListener) EnterAnyOperator(ctx *AnyOperatorContext) {}

// ExitAnyOperator is called when production anyOperator is exited.
func (s *BasevbaListener) ExitAnyOperator(ctx *AnyOperatorContext) {}

// EnterPowOperator is called when production powOperator is entered.
func (s *BasevbaListener) EnterPowOperator(ctx *PowOperatorContext) {}

// ExitPowOperator is called when production powOperator is exited.
func (s *BasevbaListener) ExitPowOperator(ctx *PowOperatorContext) {}

// EnterDivOperator is called when production divOperator is entered.
func (s *BasevbaListener) EnterDivOperator(ctx *DivOperatorContext) {}

// ExitDivOperator is called when production divOperator is exited.
func (s *BasevbaListener) ExitDivOperator(ctx *DivOperatorContext) {}

// EnterMultOperator is called when production multOperator is entered.
func (s *BasevbaListener) EnterMultOperator(ctx *MultOperatorContext) {}

// ExitMultOperator is called when production multOperator is exited.
func (s *BasevbaListener) ExitMultOperator(ctx *MultOperatorContext) {}

// EnterModOperator is called when production modOperator is entered.
func (s *BasevbaListener) EnterModOperator(ctx *ModOperatorContext) {}

// ExitModOperator is called when production modOperator is exited.
func (s *BasevbaListener) ExitModOperator(ctx *ModOperatorContext) {}

// EnterPlusOperator is called when production plusOperator is entered.
func (s *BasevbaListener) EnterPlusOperator(ctx *PlusOperatorContext) {}

// ExitPlusOperator is called when production plusOperator is exited.
func (s *BasevbaListener) ExitPlusOperator(ctx *PlusOperatorContext) {}

// EnterMinusOperator is called when production minusOperator is entered.
func (s *BasevbaListener) EnterMinusOperator(ctx *MinusOperatorContext) {}

// ExitMinusOperator is called when production minusOperator is exited.
func (s *BasevbaListener) ExitMinusOperator(ctx *MinusOperatorContext) {}

// EnterAmpOperator is called when production ampOperator is entered.
func (s *BasevbaListener) EnterAmpOperator(ctx *AmpOperatorContext) {}

// ExitAmpOperator is called when production ampOperator is exited.
func (s *BasevbaListener) ExitAmpOperator(ctx *AmpOperatorContext) {}

// EnterIsOperator is called when production isOperator is entered.
func (s *BasevbaListener) EnterIsOperator(ctx *IsOperatorContext) {}

// ExitIsOperator is called when production isOperator is exited.
func (s *BasevbaListener) ExitIsOperator(ctx *IsOperatorContext) {}

// EnterLikeOperator is called when production likeOperator is entered.
func (s *BasevbaListener) EnterLikeOperator(ctx *LikeOperatorContext) {}

// ExitLikeOperator is called when production likeOperator is exited.
func (s *BasevbaListener) ExitLikeOperator(ctx *LikeOperatorContext) {}

// EnterGeqOperator is called when production geqOperator is entered.
func (s *BasevbaListener) EnterGeqOperator(ctx *GeqOperatorContext) {}

// ExitGeqOperator is called when production geqOperator is exited.
func (s *BasevbaListener) ExitGeqOperator(ctx *GeqOperatorContext) {}

// EnterLeqOperator is called when production leqOperator is entered.
func (s *BasevbaListener) EnterLeqOperator(ctx *LeqOperatorContext) {}

// ExitLeqOperator is called when production leqOperator is exited.
func (s *BasevbaListener) ExitLeqOperator(ctx *LeqOperatorContext) {}

// EnterGtOperator is called when production gtOperator is entered.
func (s *BasevbaListener) EnterGtOperator(ctx *GtOperatorContext) {}

// ExitGtOperator is called when production gtOperator is exited.
func (s *BasevbaListener) ExitGtOperator(ctx *GtOperatorContext) {}

// EnterLtOperator is called when production ltOperator is entered.
func (s *BasevbaListener) EnterLtOperator(ctx *LtOperatorContext) {}

// ExitLtOperator is called when production ltOperator is exited.
func (s *BasevbaListener) ExitLtOperator(ctx *LtOperatorContext) {}

// EnterNeqOperator is called when production neqOperator is entered.
func (s *BasevbaListener) EnterNeqOperator(ctx *NeqOperatorContext) {}

// ExitNeqOperator is called when production neqOperator is exited.
func (s *BasevbaListener) ExitNeqOperator(ctx *NeqOperatorContext) {}

// EnterEqOperator is called when production eqOperator is entered.
func (s *BasevbaListener) EnterEqOperator(ctx *EqOperatorContext) {}

// ExitEqOperator is called when production eqOperator is exited.
func (s *BasevbaListener) ExitEqOperator(ctx *EqOperatorContext) {}

// EnterAndOperator is called when production andOperator is entered.
func (s *BasevbaListener) EnterAndOperator(ctx *AndOperatorContext) {}

// ExitAndOperator is called when production andOperator is exited.
func (s *BasevbaListener) ExitAndOperator(ctx *AndOperatorContext) {}

// EnterOrOperator is called when production orOperator is entered.
func (s *BasevbaListener) EnterOrOperator(ctx *OrOperatorContext) {}

// ExitOrOperator is called when production orOperator is exited.
func (s *BasevbaListener) ExitOrOperator(ctx *OrOperatorContext) {}

// EnterXorOperator is called when production xorOperator is entered.
func (s *BasevbaListener) EnterXorOperator(ctx *XorOperatorContext) {}

// ExitXorOperator is called when production xorOperator is exited.
func (s *BasevbaListener) ExitXorOperator(ctx *XorOperatorContext) {}

// EnterEqvOperator is called when production eqvOperator is entered.
func (s *BasevbaListener) EnterEqvOperator(ctx *EqvOperatorContext) {}

// ExitEqvOperator is called when production eqvOperator is exited.
func (s *BasevbaListener) ExitEqvOperator(ctx *EqvOperatorContext) {}

// EnterImpOperator is called when production impOperator is entered.
func (s *BasevbaListener) EnterImpOperator(ctx *ImpOperatorContext) {}

// ExitImpOperator is called when production impOperator is exited.
func (s *BasevbaListener) ExitImpOperator(ctx *ImpOperatorContext) {}
