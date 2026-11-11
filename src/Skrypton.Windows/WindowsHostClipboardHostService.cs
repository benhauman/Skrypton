using System;
using System.Runtime.InteropServices;
using System.Threading;
using Skrypton.RuntimeSupport.Components;

namespace Skrypton.Windows
{
    /// <summary>
    /// The clipboard of the user (Win32 clipboard, Unicode text) for htmlfile's parentWindow.clipboardData - the formats "Text" and "URL" are text
    /// </summary>
    public sealed class WindowsHostClipboardHostService : IHostClipboardHostService
    {
        private const uint CF_UNICODETEXT = 13;
        private const uint GMEM_MOVEABLE = 0x0002;

        public bool SetText(string format, string text)
        {
            byte[] content = System.Text.Encoding.Unicode.GetBytes((text ?? "") + "\0");
            using (OpenClipboard())
            {
                if (!NativeMethods.EmptyClipboard())
                    return false;
                IntPtr memory = NativeMethods.GlobalAlloc(GMEM_MOVEABLE, (UIntPtr)content.Length);
                if (memory == IntPtr.Zero)
                    return false;
                IntPtr target = NativeMethods.GlobalLock(memory);
                try
                {
                    Marshal.Copy(content, 0, target, content.Length);
                }
                finally
                {
                    NativeMethods.GlobalUnlock(memory);
                }
                if (NativeMethods.SetClipboardData(CF_UNICODETEXT, memory) == IntPtr.Zero)
                {
                    NativeMethods.GlobalFree(memory);
                    return false;
                }
                return true; // the clipboard owns the memory now
            }
        }

        public string? GetText(string format)
        {
            using (OpenClipboard())
            {
                if (!NativeMethods.IsClipboardFormatAvailable(CF_UNICODETEXT))
                    return null;
                IntPtr memory = NativeMethods.GetClipboardData(CF_UNICODETEXT);
                if (memory == IntPtr.Zero)
                    return null;
                IntPtr source = NativeMethods.GlobalLock(memory);
                try
                {
                    return Marshal.PtrToStringUni(source);
                }
                finally
                {
                    NativeMethods.GlobalUnlock(memory);
                }
            }
        }

        public bool Clear(string format)
        {
            using (OpenClipboard())
                return NativeMethods.EmptyClipboard();
        }

        /// <summary>
        /// Open the clipboard - another program may have it open for a moment, so this tries again for a second
        /// </summary>
        private static ClipboardScope OpenClipboard()
        {
            for (int attempt = 0; attempt < 20; attempt++)
            {
                if (NativeMethods.OpenClipboard(IntPtr.Zero))
                    return new ClipboardScope();
                Thread.Sleep(50);
            }
            throw new InvalidOperationException("clipboardData: the clipboard is in use by another program");
        }

        private sealed class ClipboardScope : IDisposable
        {
            public void Dispose() => NativeMethods.CloseClipboard();
        }

        private static class NativeMethods
        {
            [DllImport("user32.dll", SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool OpenClipboard(IntPtr hWndNewOwner);

            [DllImport("user32.dll", SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool CloseClipboard();

            [DllImport("user32.dll", SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool EmptyClipboard();

            [DllImport("user32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool IsClipboardFormatAvailable(uint format);

            [DllImport("user32.dll", SetLastError = true)]
            internal static extern IntPtr GetClipboardData(uint uFormat);

            [DllImport("user32.dll", SetLastError = true)]
            internal static extern IntPtr SetClipboardData(uint uFormat, IntPtr hMem);

            [DllImport("kernel32.dll", SetLastError = true)]
            internal static extern IntPtr GlobalAlloc(uint uFlags, UIntPtr dwBytes);

            [DllImport("kernel32.dll", SetLastError = true)]
            internal static extern IntPtr GlobalLock(IntPtr hMem);

            [DllImport("kernel32.dll", SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool GlobalUnlock(IntPtr hMem);

            [DllImport("kernel32.dll", SetLastError = true)]
            internal static extern IntPtr GlobalFree(IntPtr hMem);
        }
    }
}
