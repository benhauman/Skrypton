using Skrypton.CSharpWriter.CodeTranslation.BlockTranslators;
using Skrypton.CSharpWriter.Lists;
using Skrypton.LegacyParser.CodeBlocks;
using Skrypton.LegacyParser.CodeBlocks.Basic;
using Skrypton.LegacyParser.Tokens.Basic;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Skrypton.CSharpWriter.CodeTranslation.Extensions
{
    internal static class ScopeAccessInformationExtendExtensions
    {
        public static ScopeAccessInformation Extend(
            this ScopeAccessInformation scopeInformation,
            IHaveNestedContent parent,
            IDefineScope scopeDefiningParent,
            CSharpName? parentReturnValueNameIfAny,
            CSharpName? errorRegistrationTokenIfAny,
            IReadOnlyCollection<ICodeBlock> blocksIn)
        {
            if (parent == null)
                throw new ArgumentNullException(nameof(parent));
            if (scopeDefiningParent == null)
                throw new ArgumentNullException(nameof(scopeDefiningParent));
            if (scopeInformation == null)
                throw new ArgumentNullException(nameof(scopeInformation));
            if (blocksIn == null)
                throw new ArgumentNullException(nameof(blocksIn));

            var blocksScopeLocation = scopeDefiningParent.Scope;
            var blocksF = FlattenAllAccessibleBlockLevelCodeBlocks(blocksIn);
            var variables = scopeInformation.Variables.AddRange(
                blocksF
                    .OfType<DimStatement>() // This covers DIM, REDIM, PRIVATE and PUBLIC (they may all be considered the same for these purposes)
                    .SelectMany(d => d.Variables.Select(v => new ScopedNameToken(
                        v.Name.ContentUpperX(),
                        v.Name.LineIndex,
                        blocksScopeLocation
                    )))
                    .ToArray()
            );
            if (scopeDefiningParent != null)
            {
                variables = variables.AddRange(
                    scopeDefiningParent.ExplicitScopeAdditions
                        .Select(v => new ScopedNameToken(
                            v.ContentUpperX(),
                            v.LineIndex,
                            blocksScopeLocation
                        )
                    ).ToArray()
                );
            }
            var constants = scopeInformation.Constants.AddRange(
                blocksF
                    .OfType<ConstStatement>()
                    .SelectMany(c => c.Values.Select(v => new ScopedNameToken(
                        v.Name.ContentUpperX(),
                        v.Name.LineIndex,
                        blocksScopeLocation
                    )))
                    .ToArray()
            );

            return new ScopeAccessInformation(
                parent,
                scopeDefiningParent,
                parentReturnValueNameIfAny,
                errorRegistrationTokenIfAny,
                scopeInformation.DirectedWithReferenceIfAny,
                scopeInformation.ExternalDependencies,
                scopeInformation.Classes.AddRange(
                    blocksF
                        .Where(b => b is ClassBlock)
                        .Cast<ClassBlock>()
                        .Select(c => new ScopedNameToken(c.Name.ContentUpperX(), c.Name.LineIndex, ScopeLocationOptions.OutermostScope)) // These are always OutermostScope
                        .ToArray()
                ),
                scopeInformation.Functions.AddRange(
                    blocksF
                        .Where(b => (b is FunctionBlock) || (b is SubBlock))
                        .Cast<AbstractFunctionBlock>()
                        .Select(b => new ScopedNameToken(b.Name.ContentUpperX(), b.Name.LineIndex, blocksScopeLocation))
                        .ToArray()
                ),
                scopeInformation.Properties.AddRange(
                    blocksF
                        .Where(b => b is PropertyBlock)
                        .Cast<PropertyBlock>()
                        .Select(p => new ScopedNameToken(p.Name.ContentUpperX(), p.Name.LineIndex, ScopeLocationOptions.WithinClass)) // These are always WithinClass
                        .ToArray()
                ),
                constants,
                variables,
                scopeInformation.StructureExitPoints
            );
        }

        private static NonNullImmutableList<ICodeBlock> FlattenAllAccessibleBlockLevelCodeBlocks(IReadOnlyCollection<ICodeBlock> blocks)
        {
            if (blocks == null)
                throw new ArgumentNullException(nameof(blocks));

            var flattenedBlocks = new NonNullImmutableList<ICodeBlock>();
            foreach (var block in blocks)
            {
                flattenedBlocks = flattenedBlocks.Add(block);

                var parentBlock = block as IHaveNestedContent;
                if (parentBlock == null)
                    continue;

                if (parentBlock is IDefineScope)
                {
                    // If this defines scope then we can't expand the current scope by drilling into it - eg. if the current block
                    // is a class then it has nested statements but we can't access them directly (we can't call a function on a
                    // class without calling it on an instance of that class)
                    continue;
                }

                flattenedBlocks = flattenedBlocks.AddRange(
                    FlattenAllAccessibleBlockLevelCodeBlocks(
                        parentBlock.AllExecutableBlocks.ToNonNullImmutableList()
                    )
                );
            }
            return flattenedBlocks;
        }

        /// <summary>
        /// If the parent is scope-defining then both the parent and scopeDefiningParent references will be set to it, this is a convenience method to
        /// save having to specify it explicitly for both
        /// </summary>
        public static ScopeAccessInformation Extend(
            this ScopeAccessInformation scopeInformation,
            IDefineScope parent,
            CSharpName? parentReturnValueNameIfAny,
            CSharpName? errorRegistrationTokenIfAny,
            NonNullImmutableList<ICodeBlock> blocks)
        {
            if (scopeInformation == null)
                throw new ArgumentNullException(nameof(scopeInformation));
            if (parent == null)
                throw new ArgumentNullException(nameof(parent));
            if (blocks == null)
                throw new ArgumentNullException(nameof(blocks));

            return Extend(scopeInformation, parent, parent, parentReturnValueNameIfAny, errorRegistrationTokenIfAny, blocks);
        }

        public static ScopeAccessInformation ExtendExternalDependencies(this ScopeAccessInformation scopeInformation, IReadOnlyCollection<NameToken> externalDependencies)
        {
            if (scopeInformation == null)
                throw new ArgumentNullException(nameof(scopeInformation));
            if (externalDependencies == null)
                throw new ArgumentNullException(nameof(externalDependencies));

            return new ScopeAccessInformation(
                scopeInformation.Parent,
                scopeInformation.ScopeDefiningParent,
                scopeInformation.ParentReturnValueNameIfAny,
                scopeInformation.ErrorRegistrationTokenIfAny,
                scopeInformation.DirectedWithReferenceIfAny,
                scopeInformation.ExternalDependencies.AddRange(externalDependencies),
                scopeInformation.Classes,
                scopeInformation.Functions,
                scopeInformation.Properties,
                scopeInformation.Constants,
                scopeInformation.Variables,
                scopeInformation.StructureExitPoints
            );
        }

        public static ScopeAccessInformation ExtendVariables(this ScopeAccessInformation scopeInformation, IReadOnlyCollection<ScopedNameToken> variables)
        {
            if (scopeInformation == null)
                throw new ArgumentNullException(nameof(scopeInformation));
            if (variables == null)
                throw new ArgumentNullException(nameof(variables));

            return new ScopeAccessInformation(
                scopeInformation.Parent,
                scopeInformation.ScopeDefiningParent,
                scopeInformation.ParentReturnValueNameIfAny,
                scopeInformation.ErrorRegistrationTokenIfAny,
                scopeInformation.DirectedWithReferenceIfAny,
                scopeInformation.ExternalDependencies,
                scopeInformation.Classes,
                scopeInformation.Functions,
                scopeInformation.Properties,
                scopeInformation.Constants,
                scopeInformation.Variables.AddRange(variables),
                scopeInformation.StructureExitPoints
            );
        }

        /// <summary>
        /// If the parent is scope-defining then both the parent and scopeDefiningParent references will be set to it, this is a convenience method to
        /// save having to specify it explicitly for both (for cases where the parent scope - if any - does not have a return value)
        /// </summary>
        public static ScopeAccessInformation Extend(
            this ScopeAccessInformation scopeInformation,
            IDefineScope parent,
            NonNullImmutableList<ICodeBlock> blocks)
        {
            if (scopeInformation == null)
                throw new ArgumentNullException(nameof(scopeInformation));
            if (parent == null)
                throw new ArgumentNullException(nameof(parent));
            if (blocks == null)
                throw new ArgumentNullException(nameof(blocks));

            return Extend(scopeInformation, parent, null, null, blocks);
        }

        public static ScopeAccessInformation AddStructureExitPoints(
            this ScopeAccessInformation scopeInformation,
            CSharpName? structureExitFlagNameIfAny,
            ScopeAccessInformation.ExitableNonScopeDefiningConstructOptions structureExitType)
        {
            if (scopeInformation == null)
                throw new ArgumentNullException(nameof(scopeInformation));
            if (!Enum.IsDefined(typeof(ScopeAccessInformation.ExitableNonScopeDefiningConstructOptions), structureExitType))
                throw new ArgumentOutOfRangeException(nameof(structureExitType));

            return new ScopeAccessInformation(
                scopeInformation.Parent,
                scopeInformation.ScopeDefiningParent,
                scopeInformation.ParentReturnValueNameIfAny,
                scopeInformation.ErrorRegistrationTokenIfAny,
                scopeInformation.DirectedWithReferenceIfAny,
                scopeInformation.ExternalDependencies,
                scopeInformation.Classes,
                scopeInformation.Functions,
                scopeInformation.Properties,
                scopeInformation.Constants,
                scopeInformation.Variables,
                scopeInformation.StructureExitPoints
                    .Add(new ScopeAccessInformation.ExitableNonScopeDefiningConstructDetails(
                        structureExitFlagNameIfAny,
                        structureExitType
                    ))
            );
        }

        public static ScopeAccessInformation SetParent(this ScopeAccessInformation scopeInformation, IHaveNestedContent parent)
        {
            if (scopeInformation == null)
                throw new ArgumentNullException(nameof(scopeInformation));
            if (parent == null)
                throw new ArgumentNullException(nameof(parent));

            if ((parent != scopeInformation.ScopeDefiningParent) && !scopeInformation.ScopeDefiningParent.GetAllNestedBlocks().Contains(parent))
            {
                // The parent must either be the current ScopeDefiningParent or be one of the descendant blocks, otherwise the structure will be invalid
                throw new ArgumentException("The parent must either be the current ScopeDefiningParent or be one of the descendant blocks");
            }
            return new ScopeAccessInformation(
                parent,
                scopeInformation.ScopeDefiningParent,
                scopeInformation.ParentReturnValueNameIfAny,
                scopeInformation.ErrorRegistrationTokenIfAny,
                scopeInformation.DirectedWithReferenceIfAny,
                scopeInformation.ExternalDependencies,
                scopeInformation.Classes,
                scopeInformation.Functions,
                scopeInformation.Properties,
                scopeInformation.Constants,
                scopeInformation.Variables,
                scopeInformation.StructureExitPoints
            );
        }
    }
}