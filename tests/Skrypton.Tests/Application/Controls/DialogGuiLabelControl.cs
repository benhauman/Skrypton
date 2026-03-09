using System.Runtime.InteropServices;

namespace Skrypton.Tests.Application.Controls
{
    [ComVisible(true)]
    internal sealed class DialogGuiLabelControl : DialogGuiControlBase
    {
        public string Text { get => GetPropertyValueAsT<string>(); set => SetPropertyValueAsT(value); }
    }
}