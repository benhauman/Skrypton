using Skrypton.Tests.RuntimeSupport.Implementations;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using Skrypton.RuntimeSupport.Implementations;

namespace Skrypton.Tests.RuntimeSupport.Components.InternetExplorerApplication
{
    // "InternetExplorer.Application"
    internal sealed class MyInternetExplorerApplicationClass : IReflectOnClrType, IWebBrowser2
    {
        // ProgId: "Microsoft Internet Controls" => 'SHDocVw' (eab22ac0-30c1-11cf-a7eb-0000c05bae0b) => Interop.SHDocVw.dll
        // * shdocvw.idl
        // *  'IWebBrowser2'

        private readonly IRuntimeHost _runtimeHost;

        public MyInternetExplorerApplicationClass(IRuntimeHost runtimeHost)
        {
            _runtimeHost = runtimeHost ?? throw new ArgumentNullException(nameof(runtimeHost));
        }

        [DispId(402)] public bool Visible { get; set; }
        [DispId(104)]
        public void Navigate([In] string url)
        {
            Console.WriteLine($"[IE].Navigate('{url}')");
        }
    }

    internal interface IWebBrowser2 // : IWebBrowserApp : IWebBrowser
    {
        //[DispId(104)]
        //void Navigate([In] string URL, [Optional][In] ref object Flags, [Optional][In] ref object TargetFrameName, [Optional][In] ref object PostData, [Optional][In] object Headers);


    }
}
