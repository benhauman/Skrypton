using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml;
using Skrypton.RuntimeSupport.Attributes;

//D:\projects.ToDelete\ConsoleApp2\obj\Debug\net10.0\Interop.MSXML2.dll
namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    // test with 'CT125_ClientComputer_Dialog_349_ButtonGeneralInfo_Click'

    [SourceClassName("DOMDocument")] // for TYPENAME(CreateObject("Msxml2.DOMDocument"))
    [ComVisible(true)] // Required because .NET can auto‑implement IDispatch when (1):COM‑visible:true, (2): interface mode:AutoDispatch and (3): DISPID(0) & DISPIDs used
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    // [DefaultMember("Item")]
    internal sealed class MyMsxml2DOMDocument //  see "Msxml2.DOMDocument" from Interop.MSXML2.dll, 'MSXML2.DOMDocument60'
    {
        private readonly XmlDocument _xmldoc = new XmlDocument();
        public MyMsxml2DOMDocument()
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
        public bool load([In][MarshalAs(UnmanagedType.Struct)] object xmlSource)
        {
            if (xmlSource is byte[] byteArray)
            {
                // For testing purposes, we can assume the byte array is UTF-8 encoded XML content
                string xmlContent = Encoding.UTF8.GetString(byteArray);
                // Here you could add logic to parse the XML content if needed for your tests

                using var ms = new System.IO.MemoryStream(byteArray);
                _xmldoc.Load(ms);
                return true; // Simulate successful load
            }
            throw new NotImplementedException($"xmlSource:{xmlSource}");
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
    [DefaultMember("Item")]
    internal sealed class MyXMLDOMNodeList
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
