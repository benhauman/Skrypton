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
using System.Text;
using System.Xml;
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
        public ScriptonMsxml2DOMDocument()
        {
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
        public string xml => _xmldoc.OuterXml;

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
        public MyXMLDOMNode selectSingleNode([In][MarshalAs(UnmanagedType.BStr)] string queryString)
        {
            if (string.IsNullOrEmpty(queryString))
                return null;
            XmlNode node = _xmldoc.SelectSingleNode(queryString);
            return (node == null) ? null : new MyXMLDOMNode(node);
        }

        [DispId(40)]
        public MyXMLDOMElement documentElement
        {
            //[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(40)]
            [return: MarshalAs(UnmanagedType.Interface)]
            get
            {
                XmlElement root = _xmldoc.DocumentElement;
                if (root == null)
                    return null;
                return new MyXMLDOMElement(root);
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

                return new MyXMLDOMNode(_nodes[index]);
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            List<MyXMLDOMNode> resultNodes = new List<MyXMLDOMNode>();
            var xe = _nodes.GetEnumerator();
            if (xe != null)
            {
                while (xe.MoveNext())
                {
                    resultNodes.Add(new MyXMLDOMNode((XmlNode)xe.Current));
                }
            }

            return resultNodes.GetEnumerator();
        }
    }

    [ComVisible(true)] // Required because .NET can auto‑implement IDispatch when (1):COM‑visible:true, (2): interface mode:AutoDispatch and (3): DISPID(0) & DISPIDs used
    internal sealed class MyXMLDOMElement// : IXMLDOMElement
    {
        private readonly XmlElement _element;

        public MyXMLDOMElement(XmlElement element)
        {
            _element = element;
        }

        public string tagNameXZZZ => _element?.Name;

        public string textZZZZ
        {
            get => _element?.InnerText;
            set
            {
                if (_element != null)
                    _element.InnerText = value;
            }
        }

        // Optional: expose underlying node if needed internally
        //public XmlElement UnderlyingElement => _element;

        [DispId(30)]
        public MyXMLDOMNode selectSingleNode([In][MarshalAs(UnmanagedType.BStr)] string queryString)
        {
            if (_element == null || string.IsNullOrEmpty(queryString))
                return null;

            //try
            //{
            XmlNode node = _element.SelectSingleNode(queryString);

            if (node == null)
                return null;

            return new MyXMLDOMNode(node);
            //}
            //catch
            //{
            //    // MSXML-style behavior: invalid XPath returns null
            //    return null;
            //}
        }
    }

    [ComVisible(true)] // Required because .NET can auto‑implement IDispatch when (1):COM‑visible:true, (2): interface mode:AutoDispatch and (3): DISPID(0) & DISPIDs used
    internal sealed class MyXMLDOMNode// : IXMLDOMNode
    {
        private readonly XmlNode _node;

        public MyXMLDOMNode(XmlNode node)
        {
            _node = node;
        }

        public string nodeNameZZZ => _node?.Name;

        public string text
        {
            get => _node?.InnerText;
            set
            {
                if (_node != null)
                    _node.InnerText = value;
            }
        }

        //internal XmlNode UnderlyingNode => _node;
    }
}
