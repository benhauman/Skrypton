// Moved from the test project as-is; it was written without nullable reference type annotations
#nullable disable
// The members are bound by name from scripts (so they must stay instance members) and the names mirror the COM type library
#pragma warning disable CA1822, CA1008, CA1711, CA1716, CA1309, CA1859
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using Skrypton.RuntimeSupport.Implementations;
using Skrypton.RuntimeSupport.Attributes;

//D:\projects.ToDelete\ConsoleApp2\obj\Debug\net10.0\Interop.MSXML2.dll
namespace Skrypton.RuntimeSupport.Components
{
    // test with 'CT125_ClientComputer_Dialog_349_ButtonGeneralInfo_Click'

    [SourceClassName("DOMDocument")] // for TYPENAME(CreateObject("Msxml2.DOMDocument"))
    [ComVisible(true)] // Required because .NET can auto‑implement IDispatch when (1):COM‑visible:true, (2): interface mode:AutoDispatch and (3): DISPID(0) & DISPIDs used
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    // [DefaultMember("Item")]
    internal class ScriptonMsxml2DOMDocument //  see "Msxml2.DOMDocument" from Interop.MSXML2.dll, 'MSXML2.DOMDocument60'
    {
        /// <summary>
        /// The ProgIDs of the MSXML 3 document (TypeName DOMDocument, SelectionLanguage XSLPattern with Windows VBScript - issue #108)
        /// </summary>
        public static readonly IReadOnlyCollection<string> ProgIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Msxml2.DOMDocument", "Msxml2.DOMDocument.3.0", "Msxml2.FreeThreadedDOMDocument", "Msxml2.FreeThreadedDOMDocument.3.0",
            "Microsoft.XMLDOM", "Microsoft.FreeThreadedXMLDOM", "MSXML.DOMDocument",
        };

        private readonly XmlDocument _xmldoc = new XmlDocument();
        private readonly IHostFileSystemHostService _fileSystem;

        /// <summary>
        /// save writes the file through the host's IHostFileSystemHostService (or the process file system if the host has none), like the
        /// FileSystemObject
        /// </summary>
        public ScriptonMsxml2DOMDocument(IRuntimeHost runtimeHost)
        {
            if (runtimeHost == null)
                throw new ArgumentNullException(nameof(runtimeHost));
            _fileSystem = runtimeHost.TryGetRuntimeHostService<IHostFileSystemHostService>() ?? DefaultHostFileSystemHostService.Instance;
        }

        /// <summary>
        /// A new processing instruction (not in the document yet), like MSXML's createProcessingInstruction - for the target "xml" it is the
        /// XML declaration (issue #91)
        /// </summary>
        [DispId(31)]
        public object createProcessingInstruction([In][MarshalAs(UnmanagedType.BStr)] string target, [In][MarshalAs(UnmanagedType.BStr)] string data)
        {
            if (!"xml".Equals(target, StringComparison.Ordinal))
                return ScriptonXmlNodeFactory.Create(_xmldoc.CreateProcessingInstruction(target, data ?? ""));
            string version = XmlDeclarationPseudoAttribute(data, "version") ?? "1.0";
            return ScriptonXmlNodeFactory.Create(_xmldoc.CreateXmlDeclaration(version, XmlDeclarationPseudoAttribute(data, "encoding"), XmlDeclarationPseudoAttribute(data, "standalone")));
        }

        private static string XmlDeclarationPseudoAttribute(string data, string name)
        {
            Match match = Regex.Match(data ?? "", @"\b" + name + @"\s*=\s*(?:""(?<value>[^""]*)""|'(?<value>[^']*)')", RegexOptions.CultureInvariant);
            return match.Success ? match.Groups["value"].Value : null;
        }

        /// <summary>
        /// A new element (not in the document yet), like MSXML's createElement - a qualified name ("dt:CompanyName") keeps its prefix and the
        /// element is written with it (the namespace is the one that an ancestor declares, as with MSXML)
        /// </summary>
        [DispId(21)]
        public object createElement([In][MarshalAs(UnmanagedType.BStr)] string tagName)
        {
            return ScriptonXmlNodeFactory.Create(_xmldoc.CreateElement(tagName));
        }

        /// <summary>
        /// A new attribute (not on an element yet - element.attributes.setNamedItem adds it), like MSXML's createAttribute (issue #121)
        /// </summary>
        [DispId(24)]
        public object createAttribute([In][MarshalAs(UnmanagedType.BStr)] string name) => ScriptonXmlNodeFactory.Create(_xmldoc.CreateAttribute(name));

        /// <summary>
        /// A new CDATA section (its content is written as it is: &lt;![CDATA[..]]&gt;), like MSXML's createCDATASection (issue #121)
        /// </summary>
        [DispId(25)]
        public object createCDATASection([In][MarshalAs(UnmanagedType.BStr)] string data) => ScriptonXmlNodeFactory.Create(_xmldoc.CreateCDataSection(data ?? ""));

        /// <summary>
        /// A new text node, like MSXML's createTextNode
        /// </summary>
        [DispId(23)]
        public object createTextNode([In][MarshalAs(UnmanagedType.BStr)] string data) => ScriptonXmlNodeFactory.Create(_xmldoc.CreateTextNode(data ?? ""));

        [DispId(6)]
        public object firstChild
        {
            [return: MarshalAs(UnmanagedType.IDispatch)] // Nothing if there is no child
            get => ScriptonXmlNodeFactory.Create(_xmldoc.FirstChild);
        }

        /// <summary>
        /// Insert newChild before refChild (at the end if refChild is Nothing) and return it, like MSXML's insertBefore
        /// </summary>
        [DispId(13)]
        public object insertBefore(object newChild, object refChild) => MyXMLDOMNodeBase.InsertBefore(_xmldoc, newChild, refChild);

        /// <summary>
        /// Append newChild (an element or a processing instruction) and return it, like MSXML's appendChild
        /// </summary>
        [DispId(16)]
        public object appendChild(object newChild) => MyXMLDOMNodeBase.InsertBefore(_xmldoc, newChild, null);

        /// <summary>
        /// Write the document to a file, like MSXML's save: as the xml property, but with the XML declaration as the document has it
        /// (with its encoding), in that encoding (UTF-8 without a byte order mark if there is none)
        /// </summary>
        [DispId(64)]
        public void save(object destination)
        {
            if (destination is ScriptDispatchWrapper wrapper)
                destination = wrapper.WrappedObject;
            if (!(destination is string path) || (path.Length == 0))
                throw new ArgumentException($"Unsupported destination of type '{destination?.GetType().FullName ?? "null"}' (a file path is supported)");

            string encodingName = (_xmldoc.FirstChild as XmlDeclaration)?.Encoding;
            Encoding encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
            if (!string.IsNullOrEmpty(encodingName) && !"UTF-8".Equals(encodingName, StringComparison.OrdinalIgnoreCase))
                encoding = Encoding.GetEncoding(encodingName);
            byte[] preamble = (encoding is UTF8Encoding) ? Array.Empty<byte>() : encoding.GetPreamble();
            byte[] content = encoding.GetBytes(MsxmlSerializer.Serialize(_xmldoc, declarationWithEncoding: true));
            using (System.IO.Stream file = _fileSystem.OpenTextFileWrite(path, createIfNotExists: true, overwriteIfExists: true, append: false))
            {
                file.Write(preamble, 0, preamble.Length);
                file.Write(content, 0, content.Length);
            }
        }

        /// <summary>
        /// True by default, as with MSXML (issue #108) - documents are loaded synchronously anyway
        /// </summary>
        [DispId(61)]
        public bool async
        {
            //[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(61)]
            get;
            //[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(61)]
            [param: In]
            set;
        } = true;


        // the other properties of MSXML's DOMDocument - kept, they do not change how Skrypton parses or selects
        private static readonly HashSet<string> OtherProperties = new HashSet<string>(StringComparer.Ordinal)
        {
            "ServerHTTPRequest", "ProhibitDTD", "AllowDocumentFunction", "AllowXsltScript", "NewParser", "ResolveExternals", "ValidateOnParse",
            "MaxElementDepth", "MaxXMLSize", "ForcedResync", "UseInlineSchema", "MultipleErrorMessages", "NormalizeAttributeValues",
        };

        private readonly Dictionary<string, object> _otherProperties = new Dictionary<string, object>(StringComparer.Ordinal);

        /// <summary>
        /// Set a property, like MSXML's setProperty: "SelectionLanguage" ("XSLPattern", the default, or "XPath") and "SelectionNamespaces" change
        /// how selectNodes / selectSingleNode evaluate queries (issue #94); the names and values are case-sensitive
        /// </summary>
        [DispId(204)]
        public void setProperty([In][MarshalAs(UnmanagedType.BStr)] string name, [In][MarshalAs(UnmanagedType.Struct)] object value)
        {
            if (value is ScriptDispatchWrapper wrapper)
                value = wrapper.WrappedObject;
            MsxmlSelection selection = MsxmlSelection.For(_xmldoc);
            if (name == "SelectionLanguage")
                selection.SetLanguage(value);
            else if (name == "SelectionNamespaces")
                selection.SetNamespaces(value);
            else if ((name != null) && OtherProperties.Contains(name))
                _otherProperties[name] = value;
            else
                throw MsxmlSelection.Error("Property name is invalid.");
        }

        /// <summary>
        /// The value of a property set with setProperty (the default of the selection properties if it was not set), like MSXML's getProperty
        /// </summary>
        [DispId(203)]
        public object getProperty([In][MarshalAs(UnmanagedType.BStr)] string name)
        {
            MsxmlSelection selection = MsxmlSelection.For(_xmldoc);
            if (name == "SelectionLanguage")
                return selection.Language;
            if (name == "SelectionNamespaces")
                return selection.Namespaces;
            if ((name != null) && OtherProperties.Contains(name))
                return _otherProperties.TryGetValue(name, out object value) ? value : false;
            throw MsxmlSelection.Error("Property name is invalid.");
        }

        /// <summary>
        /// MSXML validates against the DTD while it parses - Skrypton keeps the value (it does not validate); True by default, as with MSXML (issue #164)
        /// </summary>
        [DispId(65)]
        public bool validateOnParse { get; set; } = true;

        /// <summary>
        /// MSXML resolves external definitions (DTD, entities) while it parses - Skrypton keeps the value (it does not resolve them); True by
        /// default for MSXML 3, False for MSXML 6 (issue #164)
        /// </summary>
        [DispId(66)]
        public bool resolveExternals { get; set; } = true;

        [DispId(67)]
        public bool preserveWhiteSpace
        {
            //[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(67)]
            get;
            //[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(67)]
            [param: In]
            set;
        }

        //[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(58)]
        /// <summary>
        /// Loads the document, like MSXML's DOMDocument.load: the source may be a URL or file path (String), the bytes of an XML document
        /// (a Byte array, as returned by eg. ServerXMLHTTP.responseBody), a Stream or another DOMDocument. As in MSXML, a source that can not be
        /// loaded or parsed does not raise an error - load returns False and the reason is available from parseErrorReason.
        /// </summary>
        public bool load([In][MarshalAs(UnmanagedType.Struct)] object xmlSource)
        {
            return TryLoad(() =>
            {
                switch (xmlSource is ScriptDispatchWrapper wrapper ? wrapper.WrappedObject : xmlSource)
                {
                    case string urlOrPath:
                        _xmldoc.Load(urlOrPath); // XmlDocument.Load(string) accepts file paths and URLs (file:, http:, https:)
                        break;
                    case byte[] bytes:
                        using (var stream = new System.IO.MemoryStream(bytes))
                        {
                            _xmldoc.Load(stream); // the encoding is detected from the BOM / XML declaration (UTF-8 by default)
                        }
                        break;
                    case System.IO.Stream stream:
                        _xmldoc.Load(stream);
                        break;
                    case ScriptonMsxml2DOMDocument otherDocument:
                        _xmldoc.LoadXml(otherDocument._xmldoc.OuterXml);
                        break;
                    case null:
                        throw new ArgumentException("No XML source was specified (the source is Empty or Nothing)");
                    default:
                        throw new ArgumentException($"Unsupported XML source of type '{xmlSource.GetType().FullName}' (a URL or file path, a Byte array, a Stream or a DOMDocument is supported)");
                }
            }, source: "", url: ((xmlSource is ScriptDispatchWrapper urlWrapper ? urlWrapper.WrappedObject : xmlSource) as string) ?? "");
        }

        /// <summary>
        /// Loads the document from a string of XML, like MSXML's DOMDocument.loadXML - returns False (rather than raising an error) if it can not
        /// be parsed
        /// </summary>
        [DispId(55)]
        public bool loadXML([In][MarshalAs(UnmanagedType.BStr)] string bstrXML)
        {
            return TryLoad(() => _xmldoc.LoadXml(bstrXML ?? ""), source: bstrXML ?? "", url: "");
        }

        /// <summary>
        /// The XML of the whole document
        /// </summary>
        [DispId(27)]
        public string xml => MsxmlSerializer.Serialize(_xmldoc, declarationWithEncoding: false); // MSXML ends the document with a line break

        /// <summary>
        /// Why the most recent load / loadXML returned False (empty if it succeeded) - MSXML reports this through parseError.reason
        /// </summary>
        public string parseErrorReason { get; private set; } = "";

        /// <summary>
        /// The error of the most recent load / loadXML (issue #173) - errorCode 0 and an empty reason if it succeeded
        /// </summary>
        [DispId(59)]
        public ScriptonXmlParseError parseError { get; private set; } = ScriptonXmlParseError.None;

        /// <param name="source">the XML text (loadXML), for srcText and filepos - "" for load</param>
        /// <param name="url">the URL or path (load), "" for loadXML</param>
        private bool TryLoad(Action load, string source, string url)
        {
            parseErrorReason = "";
            parseError = ScriptonXmlParseError.None;
            _xmldoc.PreserveWhitespace = preserveWhiteSpace;
            try
            {
                load();
                return true;
            }
            catch (Exception e) when ((e is XmlException) || (e is System.IO.IOException) || (e is UnauthorizedAccessException) || (e is WebException) || (e is ArgumentException) || (e is NotSupportedException) || (e is UriFormatException))
            {
                parseErrorReason = e.Message;
                parseError = ScriptonXmlParseError.From(e, source, url);
                return false;
            }
        }

        //[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(29)]
        [return: MarshalAs(UnmanagedType.Interface)]
        public MyXMLDOMNodeList selectNodes([In][MarshalAs(UnmanagedType.BStr)] string queryString)
        {
            if (string.IsNullOrEmpty(queryString))
                return new MyXMLDOMNodeList(null);
            //try
            //{
            XmlNodeList nodes = MsxmlSelection.SelectNodes(_xmldoc, queryString);
            return new MyXMLDOMNodeList(nodes);
            //}
            //catch
            //{
            //    // MSXML-style behavior: return empty list instead of throwing
            //    return new MyXMLDOMNodeList(null);
            //}
        }

        /// <summary>
        /// The first node that matches the XPath query (Nothing if there is none), like MSXML's DOMDocument.selectSingleNode
        /// </summary>
        [DispId(30)]
        [return: MarshalAs(UnmanagedType.IDispatch)]
        public object selectSingleNode([In][MarshalAs(UnmanagedType.BStr)] string queryString)
        {
            if (string.IsNullOrEmpty(queryString))
                return null;
            XmlNode node = MsxmlSelection.SelectSingleNode(_xmldoc, queryString);
            return ScriptonXmlNodeFactory.Create(node);
        }

        [DispId(40)]
        public MyXMLDOMElement documentElement
        {
            //[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(40)]
            [return: MarshalAs(UnmanagedType.Interface)]
            get
            {
                return (MyXMLDOMElement)ScriptonXmlNodeFactory.Create(_xmldoc.DocumentElement);
            }

            //[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            //[DispId(40)]
            //[param: In]
            //[param: MarshalAs(UnmanagedType.Interface)]
            //set;
        }
    }

    [ComVisible(true)] // Required because .NET can auto‑implement IDispatch when (1):COM‑visible:true, (2): interface mode:AutoDispatch and (3): DISPID(0) & DISPIDs used
    [DefaultMember("Item")] // +[DispId(0)] +[IsDefault]
    [SourceClassName("IXMLDOMSelection")] // TypeName(doc.selectNodes(..)) (as MSXML)
    internal sealed class MyXMLDOMNodeList : IEnumerable
    {
        private readonly XmlNodeList _nodes;

        public MyXMLDOMNodeList(XmlNodeList nodes)
        {
            _nodes = nodes;
            _length = nodes?.Count ?? 0;
        }

        private readonly int _length; // For testing purposes, we can set this to a fixed value or make it configurable

        [DispId(74)]
        public int length
        {
            //[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(74)]
            get => _length;
        }

        //[System.Runtime.CompilerServices.IndexerName("item")]
        [DispId(0)]
        //public MyXMLDOMNode this[[In] int index]
        [IsDefault] // needed for enumeration : dict[currentkey]
        [return: MarshalAs(UnmanagedType.IDispatch)] // Nothing for an index out of range, as with MSXML (issue #179)
        public object item(object key)
        {
            //[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            //[DispId(0)]
            //[return: MarshalAs(UnmanagedType.Interface)]
            //get
            {
                if (_nodes == null)
                    return null;

                // the index of a script is an Integer (Int16) or Long - any number (an unboxing cast of a short to int failed: Type mismatch)
                int index = Convert.ToInt32((key is ScriptDispatchWrapper wrapper) ? wrapper.WrappedObject : key, System.Globalization.CultureInfo.InvariantCulture);
                if (index < 0 || index >= _nodes.Count)
                    return null; // MSXML returns null, not exception

                return ScriptonXmlNodeFactory.Create(_nodes[index]);
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            List<object> resultNodes = new List<object>();
            var xe = _nodes.GetEnumerator();
            if (xe != null)
            {
                while (xe.MoveNext())
                {
                    resultNodes.Add(ScriptonXmlNodeFactory.Create((XmlNode)xe.Current));
                }
            }

            return resultNodes.GetEnumerator();
        }
    }

    internal static class ScriptonXmlNodeFactory
    {
        // one wrapper per node, so that "Is" compares nodes as with MSXML (doc.firstChild Is pi)
        private static readonly ConditionalWeakTable<XmlNode, object> Wrappers = new ConditionalWeakTable<XmlNode, object>();

        /// <summary>
        /// The DOM wrapper for a node - an element is an IXMLDOMElement, a processing instruction (or XML declaration) an
        /// IXMLDOMProcessingInstruction, anything else an IXMLDOMNode (Nothing for null)
        /// </summary>
        public static object Create(XmlNode node)
        {
            if (node == null)
                return null;
            return Wrappers.GetValue(node, n => n switch
            {
                XmlElement element => new MyXMLDOMElement(element),
                XmlProcessingInstruction or XmlDeclaration => new MyXMLDOMProcessingInstruction(n),
                XmlAttribute attribute => new MyXMLDOMAttribute(attribute),
                XmlCDataSection => new MyXMLDOMCDATASection(n),
                XmlText => new MyXMLDOMText(n),
                _ => new MyXMLDOMNode(n),
            });
        }
    }

    [SourceClassName("IXMLDOMProcessingInstruction")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class MyXMLDOMProcessingInstruction : MyXMLDOMNodeBase
    {
        public MyXMLDOMProcessingInstruction(XmlNode node) : base(node)
        {
        }

        public string target => Node.Name;

        public string data => Node.Value;
    }

    /// <summary>
    /// An attribute (issue #121): createAttribute, attributes.getNamedItem, selectSingleNode("@name")
    /// </summary>
    [SourceClassName("IXMLDOMAttribute")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class MyXMLDOMAttribute : MyXMLDOMNodeBase
    {
        private readonly XmlAttribute _attribute;

        public MyXMLDOMAttribute(XmlAttribute attribute) : base(attribute)
        {
            _attribute = attribute;
        }

        public string name => _attribute.Name;

        public string value
        {
            get => _attribute.Value;
            set => _attribute.Value = value ?? "";
        }
    }

    [SourceClassName("IXMLDOMCDATASection")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class MyXMLDOMCDATASection : MyXMLDOMNodeBase
    {
        public MyXMLDOMCDATASection(XmlNode node) : base(node)
        {
        }

        public string data => Node.Value;
    }

    [SourceClassName("IXMLDOMText")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class MyXMLDOMText : MyXMLDOMNodeBase
    {
        public MyXMLDOMText(XmlNode node) : base(node)
        {
        }

        public string data => Node.Value;
    }

    /// <summary>
    /// element.attributes (issue #121): length, item(i) (the default member), getNamedItem (Nothing if there is none), setNamedItem (adds or
    /// replaces the attribute of the name and returns it), removeNamedItem, For Each - in the order the attributes were set
    /// </summary>
    [SourceClassName("IXMLDOMNamedNodeMap")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [DefaultMember("item")]
    internal sealed class MyXMLDOMNamedNodeMap : IEnumerable
    {
        private readonly XmlElement _element;

        public MyXMLDOMNamedNodeMap(XmlElement element)
        {
            _element = element;
        }

        [DispId(74)]
        public int length => _element.Attributes.Count;

        [DispId(0)]
        [IsDefault]
        [return: MarshalAs(UnmanagedType.IDispatch)]
        public object item(int index) => ((index >= 0) && (index < _element.Attributes.Count)) ? ScriptonXmlNodeFactory.Create(_element.Attributes[index]) : null;

        [return: MarshalAs(UnmanagedType.IDispatch)]
        public object getNamedItem(string name) => ScriptonXmlNodeFactory.Create(_element.Attributes[name]);

        [return: MarshalAs(UnmanagedType.IDispatch)]
        public object setNamedItem(object newItem)
        {
            object item = (newItem is ScriptDispatchWrapper wrapper) ? wrapper.WrappedObject : newItem;
            if (!(item is MyXMLDOMNodeBase node) || !(node.Node is XmlAttribute attribute))
                throw MsxmlSelection.Error("The node is not an attribute.");
            _element.Attributes.SetNamedItem(attribute);
            return ScriptonXmlNodeFactory.Create(attribute);
        }

        [return: MarshalAs(UnmanagedType.IDispatch)]
        public object removeNamedItem(string name)
        {
            XmlAttribute attribute = _element.Attributes[name];
            if (attribute == null)
                return null;
            _element.Attributes.Remove(attribute);
            return ScriptonXmlNodeFactory.Create(attribute);
        }

        public IEnumerator GetEnumerator()
        {
            var attributes = new List<object>();
            foreach (XmlAttribute attribute in _element.Attributes)
                attributes.Add(ScriptonXmlNodeFactory.Create(attribute));
            return attributes.GetEnumerator();
        }
    }

    /// <summary>
    /// Writes nodes as MSXML does (its xml property and save): names and attributes as they are in the document - no namespace declarations
    /// are added (.NET adds xmlns="" to a child without a namespace under a default namespace) - empty elements as "&lt;x/&gt;", each node
    /// of the document on a line of its own (CRLF) and the XML declaration without its encoding (save keeps it)
    /// </summary>
    internal static class MsxmlSerializer
    {
        public static string Serialize(XmlDocument document, bool declarationWithEncoding)
        {
            var text = new StringBuilder();
            foreach (XmlNode node in document.ChildNodes)
            {
                if ((node is XmlWhitespace) || (node is XmlSignificantWhitespace))
                    continue;
                if ((node is XmlDeclaration declaration) && declarationWithEncoding)
                    text.Append("<?xml ").Append(declaration.Value).Append("?>");
                else
                    Write(text, node);
                text.Append("\r\n");
            }
            return text.ToString();
        }

        public static string Serialize(XmlNode node)
        {
            if (node is XmlDocument document)
                return Serialize(document, declarationWithEncoding: false);
            var text = new StringBuilder();
            Write(text, node);
            return text.ToString();
        }

        private static void Write(StringBuilder text, XmlNode node)
        {
            switch (node)
            {
                case XmlElement element:
                    text.Append('<').Append(element.Name);
                    foreach (XmlAttribute attribute in element.Attributes)
                        text.Append(' ').Append(attribute.Name).Append("=\"").Append(Escape(attribute.Value, attribute: true)).Append('"');
                    if (!element.HasChildNodes && element.IsEmpty)
                    {
                        text.Append("/>");
                        break;
                    }
                    text.Append('>');
                    foreach (XmlNode child in element.ChildNodes)
                        Write(text, child);
                    text.Append("</").Append(element.Name).Append('>');
                    break;
                case XmlDeclaration declaration:
                    text.Append("<?xml version=\"").Append(declaration.Version).Append('"');
                    if (!string.IsNullOrEmpty(declaration.Standalone))
                        text.Append(" standalone=\"").Append(declaration.Standalone).Append('"');
                    text.Append("?>");
                    break;
                case XmlProcessingInstruction instruction:
                    text.Append("<?").Append(instruction.Target);
                    if (!string.IsNullOrEmpty(instruction.Data))
                        text.Append(' ').Append(instruction.Data);
                    text.Append("?>");
                    break;
                case XmlText or XmlWhitespace or XmlSignificantWhitespace:
                    text.Append(Escape(node.Value, attribute: false));
                    break;
                case XmlAttribute attributeNode:
                    text.Append(attributeNode.Name).Append("=\"").Append(Escape(attributeNode.Value, attribute: true)).Append('"');
                    break;
                default: // CDATA, comments, entity references, document types
                    text.Append(node.OuterXml);
                    break;
            }
        }

        private static string Escape(string value, bool attribute)
        {
            var text = new StringBuilder(value?.Length ?? 0);
            foreach (char c in value ?? "")
            {
                switch (c)
                {
                    case '&': text.Append("&amp;"); break;
                    case '<': text.Append("&lt;"); break;
                    case '>': text.Append("&gt;"); break;
                    case '"' when attribute: text.Append("&quot;"); break;
                    default: text.Append(c); break;
                }
            }
            return text.ToString();
        }
    }

    [SourceClassName("IXMLDOMElement")] // TypeName(doc.documentElement), as MSXML
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class MyXMLDOMElement : MyXMLDOMNodeBase
    {
        private readonly XmlElement _element;

        public MyXMLDOMElement(XmlElement element) : base(element)
        {
            _element = element;
        }

        public string tagName => _element.Name;

        /// <summary>
        /// The value of the attribute, or Null if the element has no such attribute (as MSXML)
        /// </summary>
        public object getAttribute(string name) => _element.HasAttribute(name) ? _element.GetAttribute(name) : (object)DBNull.Value;

        public void setAttribute(string name, object value) => _element.SetAttribute(name, Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture));

        /// <summary>
        /// The attributes of the element (issue #121)
        /// </summary>
        public MyXMLDOMNamedNodeMap attributes => new MyXMLDOMNamedNodeMap(_element);

        private const string DataTypesNamespace = "urn:schemas-microsoft-com:datatypes";

        /// <summary>
        /// The data type of the element (its dt:dt attribute, as MSXML writes it), Null if it has none (issue #170)
        /// </summary>
        public object dataType
        {
            get => _element.HasAttribute("dt", DataTypesNamespace) ? _element.GetAttribute("dt", DataTypesNamespace) : (object)DBNull.Value;
            set
            {
                string type = Convert.ToString((value is ScriptDispatchWrapper wrapper) ? wrapper.WrappedObject : value, System.Globalization.CultureInfo.InvariantCulture) ?? "";
                if (!_element.HasAttribute("xmlns:dt"))
                    _element.SetAttribute("xmlns:dt", DataTypesNamespace);
                XmlAttribute attribute = _element.GetAttributeNode("dt", DataTypesNamespace) ?? _element.SetAttributeNode(_element.OwnerDocument.CreateAttribute("dt", "dt", DataTypesNamespace));
                attribute.Value = type;
            }
        }

        /// <summary>
        /// The value of the element as its data type (issue #170): bin.hex and bin.base64 are the bytes (Byte()) of the text - setting bytes writes
        /// them as lower-case hex or as Base64 in lines of 72 characters, as MSXML does; other elements are their text
        /// </summary>
        public object nodeTypedValue
        {
            get
            {
                string text = _element.InnerText;
                switch (dataType as string)
                {
                    case "bin.hex":
                        string hex = new string(text.Where(c => !char.IsWhiteSpace(c)).ToArray());
                        var bytes = new byte[hex.Length / 2];
                        for (int i = 0; i < bytes.Length; i++)
                            bytes[i] = Convert.ToByte(hex.Substring(2 * i, 2), 16);
                        return bytes;
                    case "bin.base64":
                        return Convert.FromBase64String(new string(text.Where(c => !char.IsWhiteSpace(c)).ToArray()));
                    default:
                        return text;
                }
            }
            set
            {
                object content = (value is ScriptDispatchWrapper wrapper) ? wrapper.WrappedObject : value;
                string type = dataType as string;
                if ((content is byte[] bytes) && (type == "bin.hex"))
                    _element.InnerText = string.Concat(bytes.Select(b => b.ToString("x2", System.Globalization.CultureInfo.InvariantCulture)));
                else if ((content is byte[] base64Bytes) && (type == "bin.base64"))
                    _element.InnerText = Base64Lines(Convert.ToBase64String(base64Bytes));
                else
                    _element.InnerText = Convert.ToString(content, System.Globalization.CultureInfo.InvariantCulture) ?? "";
            }
        }

        private static string Base64Lines(string base64)
        {
            var lines = new System.Collections.Generic.List<string>();
            for (int index = 0; index < base64.Length; index += 72)
                lines.Add(base64.Substring(index, Math.Min(72, base64.Length - index)));
            return string.Join("\n", lines);
        }
    }

    [SourceClassName("IXMLDOMNode")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class MyXMLDOMNode : MyXMLDOMNodeBase
    {
        public MyXMLDOMNode(XmlNode node) : base(node)
        {
        }
    }

    /// <summary>
    /// The IXMLDOMNode members that elements and other nodes share
    /// </summary>
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal abstract class MyXMLDOMNodeBase
    {
        private readonly XmlNode _node;

        protected MyXMLDOMNodeBase(XmlNode node)
        {
            _node = node ?? throw new ArgumentNullException(nameof(node));
        }

        internal XmlNode Node => _node;

        public object firstChild
        {
            [return: MarshalAs(UnmanagedType.IDispatch)] // Nothing if there is no child
            get => ScriptonXmlNodeFactory.Create(_node.FirstChild);
        }

        /// <summary>
        /// Insert newChild before refChild (at the end if refChild is Nothing) and return it, like MSXML's insertBefore
        /// </summary>
        public object insertBefore(object newChild, object refChild) => InsertBefore(_node, newChild, refChild);

        /// <summary>
        /// Append newChild and return it, like MSXML's appendChild
        /// </summary>
        public object appendChild(object newChild) => InsertBefore(_node, newChild, null);

        internal static object InsertBefore(XmlNode parent, object newChild, object refChild)
        {
            XmlNode child = NodeOf(newChild) ?? throw new ArgumentNullException(nameof(newChild));
            // MSXML 3 does not split a CDATA section with "]]>" - appending it is an error (E_INVALIDARG, issue #121)
            if ((child is XmlCDataSection cdata) && (cdata.Value ?? "").Contains("]]>"))
                throw new Skrypton.RuntimeSupport.Exceptions.CustomException(unchecked((int)0x80070057), MsxmlSelection.ErrorSource, "Invalid data for a node of type 'CDATA'.\r\n");
            XmlNode reference = NodeOf(refChild);
            if (reference == null)
                parent.AppendChild(child);
            else
                parent.InsertBefore(child, reference);
            return ScriptonXmlNodeFactory.Create(child);
        }

        private static XmlNode NodeOf(object value)
        {
            if (value is ScriptDispatchWrapper wrapper)
                value = wrapper.WrappedObject;
            return value switch
            {
                null => null,
                DBNull => null,
                MyXMLDOMNodeBase node => node._node,
                _ => throw new ArgumentException($"Not a node of the document: '{value.GetType().FullName}'"),
            };
        }

        public string nodeName => _node.Name;

        /// <summary>
        /// The child nodes (length, item(i), For Each)
        /// </summary>
        public MyXMLDOMNodeList childNodes => new MyXMLDOMNodeList(_node.ChildNodes);

        /// <summary>
        /// The value of an attribute, text or comment node, Null for an element (as MSXML)
        /// </summary>
        public object nodeValue => (_node is XmlElement) ? DBNull.Value : (object)_node.Value;

        public string text
        {
            get => _node.InnerText;
            set => _node.InnerText = value;
        }

        public string xml => MsxmlSerializer.Serialize(_node);

        [DispId(30)]
        [return: MarshalAs(UnmanagedType.IDispatch)] // Nothing if no node matches, as with MSXML (issue #179)
        public object selectSingleNode([In][MarshalAs(UnmanagedType.BStr)] string queryString)
        {
            if (string.IsNullOrEmpty(queryString))
                return null;
            return ScriptonXmlNodeFactory.Create(MsxmlSelection.SelectSingleNode(_node, queryString));
        }

        public MyXMLDOMNodeList selectNodes([In][MarshalAs(UnmanagedType.BStr)] string queryString)
        {
            return new MyXMLDOMNodeList(MsxmlSelection.SelectNodes(_node, queryString));
        }
    }
}
