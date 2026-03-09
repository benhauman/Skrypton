using System;

namespace Skrypton.RuntimeSupport.Implementations
{
    public interface IHostMessageBoxHostService
    {
        MessageBoxResult ShowMessageBox(string prompt, MessageBoxButtons buttons, string v2);
    }
    public enum MessageBoxButtons // https://learn.microsoft.com/de-de/office/vba/language/reference/user-interface-help/msgbox-function
    {
        vbOkOnly = 0,
        //vbOKCancel = 1,
    }
    public enum MessageBoxResult // https://learn.microsoft.com/de-de/office/vba/language/reference/user-interface-help/msgbox-function
    {
        [Obsolete("do not use it.")]
        Unknown = 0,
        vbOK = 1,
    }
}