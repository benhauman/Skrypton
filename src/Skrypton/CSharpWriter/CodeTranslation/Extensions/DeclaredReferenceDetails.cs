using Skrypton.LegacyParser.CodeBlocks.Basic;
using System;

namespace Skrypton.CSharpWriter.CodeTranslation.Extensions
{
    public class DeclaredReferenceDetails
    {
        public DeclaredReferenceDetails(ReferenceTypeOptions referenceType, ScopeLocationOptions scopeLocation, string? ownerName = null)
        {
            if (!Enum.IsDefined(typeof(ReferenceTypeOptions), referenceType))
                throw new ArgumentOutOfRangeException(nameof(referenceType));
            if (!Enum.IsDefined(typeof(ScopeLocationOptions), scopeLocation))
                throw new ArgumentOutOfRangeException(nameof(scopeLocation));

            ReferenceType = referenceType;
            ScopeLocation = scopeLocation;
            OwnerName = ownerName;
        }

        public ReferenceTypeOptions ReferenceType { get; private set; }
        public ScopeLocationOptions ScopeLocation { get; private set; }
        public string? OwnerName { get; }
    }
}
