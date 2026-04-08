using Skrypton.Tests.RuntimeSupport.Implementations;
using System;
using System.Collections.Generic;
using System.Text;
using Skrypton.RuntimeSupport.Implementations;

namespace Skrypton.Tests.RuntimeSupport.Components.HTMLObjectLibrary
{
    internal sealed class MyHTMLDocumentClass : IReflectOnClrType
    {
        private readonly IRuntimeHost _runtimeHost;

        /*
         Add COM Reference  “Microsoft HTML Object Library” => Interop.MSHTML.dll
         */

        public MyHTMLDocumentClass(IRuntimeHost runtimeHost)
        {
            _runtimeHost = runtimeHost ?? throw new ArgumentNullException(nameof(runtimeHost));
        }

        private MyHTMLDocumentWindowClass _parentWindow;
        public object ParentWindow
        {
            get
            {
                if (_parentWindow == null)
                {
                    _parentWindow = new MyHTMLDocumentWindowClass();
                }
                return _parentWindow;
            }
        }
    }

    internal sealed class MyHTMLDocumentWindowClass : IReflectOnClrType
    {
        public MyHTMLDocumentWindowClass()
        {
        }

        private MyHTMLDocumentClipboardClass _clipboard;

        public object ClipboardData
        {
            get
            {
                if (_clipboard == null)
                {
                    _clipboard = new MyHTMLDocumentClipboardClass();
                }

                return _clipboard;
            }
        }
    }

    internal sealed class MyHTMLDocumentClipboardClass : IReflectOnClrType
    {
        public MyHTMLDocumentClipboardClass()
        {
        }

        public void SetData(object p1, object data)
        {
            Console.WriteLine($"[HTML].parentWindow.clipboardData.setData('{p1}', '{data}')");
        }
    }
}
