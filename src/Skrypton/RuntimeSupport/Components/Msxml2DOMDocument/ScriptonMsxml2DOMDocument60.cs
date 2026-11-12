using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Skrypton.RuntimeSupport.Attributes;
using Skrypton.RuntimeSupport.Implementations;

namespace Skrypton.RuntimeSupport.Components
{
    /// <summary>
    /// The built-in MSXML 6 DOM document (Msxml2.DOMDocument.6.0): the MSXML 3 document, but with TypeName DOMDocument60 and XPath as its
    /// selection language, as with Windows VBScript (issue #108)
    /// </summary>
    [SourceClassName("DOMDocument60")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonMsxml2DOMDocument60 : ScriptonMsxml2DOMDocument
    {
        /// <summary>
        /// The ProgIDs of the MSXML 6 document
        /// </summary>
        public static readonly new IReadOnlyCollection<string> ProgIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Msxml2.DOMDocument.6.0", "Msxml2.FreeThreadedDOMDocument.6.0",
        };

        public ScriptonMsxml2DOMDocument60(IRuntimeHost runtimeHost) : base(runtimeHost)
        {
            setProperty("SelectionLanguage", MsxmlSelection.XPath);
        }
    }
}
