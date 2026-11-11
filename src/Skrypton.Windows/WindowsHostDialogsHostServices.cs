using System;
using System.Runtime.InteropServices;
using System.Threading;
using Skrypton.RuntimeSupport.Implementations;

namespace Skrypton.Windows
{
    /// <summary>
    /// MsgBox as a Windows message box (the buttons, icon and default button flags of MsgBox are those of MessageBox)
    /// </summary>
    public sealed class WindowsHostMessageBoxHostService : IHostMessageBoxHostService
    {
        private const uint MB_SETFOREGROUND = 0x00010000;

        public MessageBoxResult ShowMessageBox(string prompt, MessageBoxButtons buttons, string v2)
        {
            int result = NativeMethods.MessageBoxW(IntPtr.Zero, prompt ?? "", string.IsNullOrEmpty(v2) ? "VBScript" : v2, (uint)buttons | MB_SETFOREGROUND);
            return (MessageBoxResult)result; // the results of MessageBox are those of MsgBox (vbOK 1, vbCancel 2, .. vbNo 7)
        }

        private static class NativeMethods
        {
            [DllImport("user32.dll", CharSet = CharSet.Unicode)]
            internal static extern int MessageBoxW(IntPtr hWnd, string lpText, string lpCaption, uint uType);
        }
    }

    /// <summary>
    /// InputBox as a Windows dialog (prompt, a text box with the default, OK / Cancel) - Cancel returns "", as with VBScript
    /// </summary>
    public sealed class WindowsHostInputBoxHostService : IHostInputBoxHostService
    {
        public string ShowInputBox(string prompt, string title, string defaultText)
        {
            string result = "";
            // a dialog needs a single-threaded apartment
            var thread = new Thread(() => result = Show(prompt ?? "", string.IsNullOrEmpty(title) ? "VBScript" : title, defaultText ?? ""));
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();
            return result;
        }

        private static string Show(string prompt, string title, string defaultText)
        {
            using (var form = new System.Windows.Forms.Form())
            using (var label = new System.Windows.Forms.Label())
            using (var textBox = new System.Windows.Forms.TextBox())
            using (var ok = new System.Windows.Forms.Button())
            using (var cancel = new System.Windows.Forms.Button())
            {
                form.Text = title;
                form.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
                form.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
                form.MinimizeBox = false;
                form.MaximizeBox = false;
                form.ClientSize = new System.Drawing.Size(420, 140);
                form.TopMost = true;

                label.Text = prompt;
                label.SetBounds(12, 12, 310, 80);
                textBox.Text = defaultText;
                textBox.SetBounds(12, 104, 396, 24);
                ok.Text = "OK";
                ok.DialogResult = System.Windows.Forms.DialogResult.OK;
                ok.SetBounds(333, 12, 75, 26);
                cancel.Text = "Cancel";
                cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
                cancel.SetBounds(333, 44, 75, 26);

                form.Controls.AddRange(new System.Windows.Forms.Control[] { label, textBox, ok, cancel });
                form.AcceptButton = ok;
                form.CancelButton = cancel;
                return (form.ShowDialog() == System.Windows.Forms.DialogResult.OK) ? textBox.Text : "";
            }
        }
    }
}
