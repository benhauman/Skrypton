using System;
using Skrypton.CSharpWriter.CodeTranslation.Extensions;
using Skrypton.LegacyParser.CodeBlocks.Basic;
using Skrypton.LegacyParser.Tokens.Basic;

namespace Skrypton.CSharpWriter.CodeTranslation
{
    public static class OutermostScopeKnownTypes
    {
        public static readonly Type[] AllKnownTypes = new Type[] {
     typeof(CommentStatement) // [DataContract(Namespace = "http://vbs")]
    ,typeof(InlineCommentStatement) // [DataContract(Namespace = "http://vbs")]
    ,typeof(SubBlock) // [DataContract(Namespace = "http://vbs")]
    ,typeof(BlankLine) // [DataContract(Namespace = "http://vbs")]
    ,typeof(ValueSettingStatement) // [DataContract(Namespace = "http://vbs")]
    ,typeof(ValueSetTypeOptions) // [DataContract(Namespace = "http://vbs")]
    ,typeof(BuiltInFunctionToken) // [DataContract(Namespace = "http://vbs")]
    ,typeof(OpenBrace) // [DataContract(Namespace = "http://vbs")]
    ,typeof(CloseBrace) // [DataContract(Namespace = "http://vbs")]
    ,typeof(StringToken) // [DataContract(Namespace = "http://vbs")]
    ,typeof(NameToken) // [DataContract(Namespace = "http://vbs")]
    ,typeof(NumericValueToken) // [DataContract(Namespace = "http://vbs")]
    ,typeof(MemberAccessorOrDecimalPointToken) // [DataContract(Namespace = "http://vbs")]
    ,typeof(BuiltInValueToken) // [DataContract(Namespace = "http://vbs")]
    ,typeof(Statement) // [DataContract(Namespace = "http://vbs")]

//    ,typeof(CallStatement) // [DataContract(Namespace = "http://vbs")]
    ,typeof(OperatorToken) // [DataContract(Namespace = "http://vbs")]
    ,typeof(IfBlock) // [DataContract(Namespace = "http://vbs")]
    ,typeof(OptionExplicit) // [DataContract(Namespace = "http://vbs")]
    ,typeof(DimVariable) // [DataContract(Namespace = "http://vbs")]
    ,typeof(BaseDimStatement) // [DataContract(Namespace = "http://vbs")]
    ,typeof(DimStatement) // [DataContract(Namespace = "http://vbs")]
//    ,typeof(LocalDimStatement) // [DataContract(Namespace = "http://vbs")]
    ,typeof(ConstantNonNegativeArrayDimensionDimVariable) // [DataContract(Namespace = "http://vbs")]
    ,typeof(ConstStatement) // [DataContract(Namespace = "http://vbs")]
    ,typeof(FunctionBlock) // [DataContract(Namespace = "http://vbs")]
    ,typeof(ArgumentSeparatorToken) // [DataContract(Namespace = "http://vbs")]
    ,typeof(ForEachBlock) // [DataContract(Namespace = "http://vbs")]
    ,typeof(ForBlock) // [DataContract(Namespace = "http://vbs")]
    ,typeof(ExitStatement) // [DataContract(Namespace = "http://vbs")]
    ,typeof(ComparisonOperatorToken) // [DataContract(Namespace = "http://vbs")]
    ,typeof(LogicalOperatorToken) // [DataContract(Namespace = "http://vbs")]
//    ,typeof(ArithmeticAndStringOperatorToken) // [DataContract(Namespace = "http://vbs")]
    ,typeof(ReDimStatement) // [DataContract(Namespace = "http://vbs")]
    ,typeof(DoBlock) // [DataContract(Namespace = "http://vbs")]
    ,typeof(EraseStatement) // [DataContract(Namespace = "http://vbs")]
    ,typeof(OnErrorResumeNext) // [DataContract(Namespace = "http://vbs")]
    ,typeof(OnErrorGoto0) // [DataContract(Namespace = "http://vbs")]
    ,typeof(WithBlock) // [DataContract(Namespace = "http://vbs")]
    ,typeof(DoNotRenameNameToken) // [DataContract(Namespace = "http://vbs")]
    ,typeof(KeyWordToken) // [DataContract(Namespace = "http://vbs")]
    ,typeof(SelectBlock) // [DataContract(Namespace = "http://vbs")]
    ,typeof(SelectBlock.CaseBlockSegment) // [DataContract(Namespace = "http://vbs")]
    ,typeof(SelectBlock.CaseBlockElseSegment) // [DataContract(Namespace = "http://vbs")]
    ,typeof(SelectBlock.CaseBlockExpressionSegment) // [DataContract(Namespace = "http://vbs")]
    ,typeof(Expression) // [DataContract(Namespace = "http://vbs")]
        ,typeof(OutermostScope) // [DataContract(Namespace = "http://vbs")]

        };
    }
}
