using System;

namespace Skrypton.RuntimeSupport.Attributes
{
    /// <summary>
    /// In order to fully implement VBScript TypeName support, we will need the original names of classes before they were changed for C# generation (for cases
    /// where they WERE changed). Generated classes should be decorated with this attribute to expose that information.
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    public sealed class SourceClassName : Attribute
    {
        public SourceClassName(string name)
        {
            Name = name ?? throw new ArgumentNullException(nameof(name));
        }

        /// <summary>
        /// This will never be null (but this is pretty much the only guarantee we can make due to VBScript's crazy variable name escaping support)
        /// </summary>
        public string Name { get; private set; }
    }
}
