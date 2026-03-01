using System;
using System.Collections.Generic;
using Skrypton.LegacyParser.CodeBlocks;
using Skrypton.LegacyParser.CodeBlocks.Basic;

namespace Skrypton.CSharpWriter.CodeTranslation.BlockTranslators
{
    public static class IHaveNestedContent_Extensions
    {
        /// <summary>
        /// The IHaveNestedContent property AllExecutableBlocks will return executable blocks that are directly contained within a code block and, in the case
        /// of an if or a while block, any executable blocks in conditions. If a full recursive scan of executable blocks is required then this method may be
        /// used.
        /// </summary>
        public static IEnumerable<ICodeBlock> GetAllNestedBlocks(this IHaveNestedContent nestedContentBlock)
        {
            if (nestedContentBlock == null)
                throw new ArgumentNullException(nameof(nestedContentBlock));

            foreach (var codeBlock in nestedContentBlock.AllExecutableBlocks)
            {
                yield return codeBlock;

                var doubleNestedContentBlock = codeBlock as IHaveNestedContent;
                if (doubleNestedContentBlock != null)
                {
                    foreach (var doubleNestedExecutableBlock in doubleNestedContentBlock.GetAllNestedBlocks())
                        yield return doubleNestedExecutableBlock;
                }
            }
        }

        /// <summary>
        /// Does the content within the specified block (that contains nested content) contain a loop with a mismatched EXIT statement (eg. does it contain
        /// a FOR loop with an EXIT DO that does not have a DO loop between the FOR and the EXIT DO) - if so, then the EXIT statment will require multiple
        /// break statements in the translated C# code (one to exit its mismatched containing loop and then another to exit the loop that it targets).
        /// </summary>
        public static bool ContainsLoopThatContainsMismatchedExitThatMustBeHandledAtThisLevel(this IHaveNestedContent nestedContentBlock)
        {
            if (nestedContentBlock == null)
                throw new ArgumentNullException(nameof(nestedContentBlock));

            foreach (ICodeBlock codeBlock in nestedContentBlock.AllExecutableBlocks)
            {
                // If a ForBlock or DoBlock is reached then pass off handling to ContainsMismatchedExitThatMustBeHandledAtThisLevel, specifying an expected
                // exit type consistent with the loop
                if (codeBlock is ForBlock forBlock)
                {
                    if (ContainsMismatchedExitThatMustBeHandledAtThisLevel(forBlock, expectedExitType: ExitStatement.ExitableStatementType.For))
                    {
                        return true;
                    }
                }
                else if (codeBlock is ForEachBlock forEachBlock)
                {
                    if (ContainsMismatchedExitThatMustBeHandledAtThisLevel(forEachBlock, expectedExitType: ExitStatement.ExitableStatementType.For))
                    {
                        return true;
                    }
                }
                else if (codeBlock is DoBlock doBlock)
                {
                    if (ContainsMismatchedExitThatMustBeHandledAtThisLevel(doBlock, expectedExitType: ExitStatement.ExitableStatementType.Do))
                    {
                        return true;
                    }
                }
                else if (codeBlock is ILoopOverNestedContent)
                {
                    // Sanity checking - if there's another looping construct then this method won't be dealing with it properly!
                    throw new ArgumentException("Unexpected ILoopOverNestedContent type: " + codeBlock.GetType());
                }
                else if (codeBlock is IHaveNestedContent doubleNestedContentBlock)
                {
                    if (ContainsLoopThatContainsMismatchedExitThatMustBeHandledAtThisLevel(doubleNestedContentBlock))
                    {
                        return true;
                    }
                }
                else
                {
                    // try next one
                }
            }
            return false;
        }

        private static bool ContainsMismatchedExitThatMustBeHandledAtThisLevel(IHaveNestedContent nestedContentBlock, ExitStatement.ExitableStatementType expectedExitType)
        {
            if (nestedContentBlock == null)
                throw new ArgumentNullException(nameof(nestedContentBlock));
            if (!Enum.IsDefined(typeof(ExitStatement.ExitableStatementType), expectedExitType))
                throw new ArgumentOutOfRangeException(nameof(expectedExitType));

            foreach (var codeBlock in nestedContentBlock.AllExecutableBlocks)
            {
                // If a ForBlock or DoBlock is reached then there can't be a mismatched exit statement that must be handled here - either the For/DoBlock
                // will contain no exit statement or it will contain a mistmatched exit that it must handle itself or it will contain a non-mismatched
                // exit (all of these scenarios indicate that there is no mismatched exit at this level
                if (codeBlock is ILoopOverNestedContent)
                    continue;

                var exitStatement = codeBlock as ExitStatement;
                if (exitStatement != null)
                {
                    if (exitStatement.StatementType != expectedExitType)
                        return true;
                    continue;
                }

                var doubleNestedContentBlock = codeBlock as IHaveNestedContent;
                if (doubleNestedContentBlock != null)
                {
                    if (ContainsMismatchedExitThatMustBeHandledAtThisLevel(doubleNestedContentBlock, expectedExitType))
                        return true;
                }
            }
            return false;
        }
    }
}
