using System.Runtime.InteropServices;

namespace Skrypton.Tests.Application.Controls
{
    [ComVisible(true)]
    internal sealed class DialogGuiTextControl : DialogGuiControlBase // <ControlName>HelpLineTextBox</ControlName>
    {
        private string _valueText;
        public string Text
        {
            get => RetrieveValueForText();
            set => UpdateValueForText(value);
        }

        private void UpdateValueForText(string value)
        {
            _valueText = value;
        }

        private string RetrieveValueForText()
        {
            return _valueText;
        }
    }
}