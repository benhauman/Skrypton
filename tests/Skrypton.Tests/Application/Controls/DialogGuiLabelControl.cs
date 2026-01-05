using System.Runtime.InteropServices;

namespace Skrypton.Tests.Application.Controls
{
    [ComVisible(true)]
    internal sealed class DialogGuiLabelControl : DialogGuiControlBase // <ControlName>HelpLineTextBox</ControlName>
    {
        private string _valueText;
        public string Text
        {
            get => _valueText;
            set => _valueText = value;
        }
    }
}