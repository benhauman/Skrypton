// Moved from the test project as-is; it was written without nullable reference type annotations
#nullable disable
// The members are bound by name from scripts (so they must stay instance members) and the names mirror the COM type library
#pragma warning disable CA1822, CA1008, CA1711, CA1716, CA1309, CA1859
using System;
using System.Collections;
using System.Collections.Generic;
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
    internal sealed class ScriptonMsxml2DOMDocument //  see "Msxml2.DOMDocument" from Interop.MSXML2.dll, 'MSXML2.DOMDocument60'
    {
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
        }


        //[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(204)]
        public void setProperty([In][MarshalAs(UnmanagedType.BStr)] string name, [In][MarshalAs(UnmanagedType.Struct)] object value)
        {
            if ("SelectionLanguage".Equals(name, StringComparison.OrdinalIgnoreCase))
            {
                if (value != null && value is string strValue)
                {
                    if ("XPath".Equals(strValue, StringComparison.OrdinalIgnoreCase))
                    {
                        // Supported value, do nothing
                        return;
                    }
                    else
                    {
                        throw new NotSupportedException($"Unsupported value for SelectionLanguage: {strValue}");
                    }
                }
                else
                {
                    throw new ArgumentException($"Expected a string value for SelectionLanguage, but got: {value?.GetType().Name ?? "null"}");
                }
            }
            throw new NotImplementedException($"name:{name}, value:{value}");
        }

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
            });
        }

        /// <summary>
        /// Loads the document from a string of XML, like MSXML's DOMDocument.loadXML - returns False (rather than raising an error) if it can not
        /// be parsed
        /// </summary>
        [DispId(55)]
        public bool loadXML([In][MarshalAs(UnmanagedType.BStr)] string bstrXML)
        {
            return TryLoad(() => _xmldoc.LoadXml(bstrXML ?? ""));
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

        private bool TryLoad(Action load)
        {
            parseErrorReason = "";
            _xmldoc.PreserveWhitespace = preserveWhiteSpace;
            try
            {
                load();
                return true;
            }
            catch (Exception e) when ((e is XmlException) || (e is System.IO.IOException) || (e is UnauthorizedAccessException) || (e is WebException) || (e is ArgumentException) || (e is NotSupportedException) || (e is UriFormatException))
            {
                parseErrorReason = e.Message;
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
            XmlNodeList nodes = _xmldoc.SelectNodes(queryString);
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
            XmlNode node = _xmldoc.SelectSingleNode(queryString);
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
        public object item(object key)
        {
            //[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            //[DispId(0)]
            //[return: MarshalAs(UnmanagedType.Interface)]
            //get
            {
                if (_nodes == null)
                    return null;

                int index = (int)key;
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
        public object selectSingleNode([In][MarshalAs(UnmanagedType.BStr)] string queryString)
        {
            if (string.IsNullOrEmpty(queryString))
                return null;
            return ScriptonXmlNodeFactory.Create(_node.SelectSingleNode(queryString));
        }

        public MyXMLDOMNodeList selectNodes([In][MarshalAs(UnmanagedType.BStr)] string queryString)
        {
            return new MyXMLDOMNodeList(string.IsNullOrEmpty(queryString) ? null : _node.SelectNodes(queryString));
        }
    }
}
