using Skrypton.RuntimeSupport.Attributes;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    /// <summary>
    /// This is an example of the type of class that may be emitted by the translation process, one with a parameter-less default member
    /// </summary>
    [SourceClassName("ExampleDefaultPropertyType")]
#pragma warning disable CS8981 // The type name only contains lower-cased ascii characters. Such names may become reserved for the language.
    public class exampledefaultpropertytype
#pragma warning restore CS8981 // The type name only contains lower-cased ascii characters. Such names may become reserved for the language.
    {
        [IsDefault]
        public object result { get; set; }
    }
}
