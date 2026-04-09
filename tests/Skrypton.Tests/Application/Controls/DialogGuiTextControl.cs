using System.Runtime.InteropServices;

namespace Skrypton.Tests.Application.Controls
{
    [ComVisible(true)]
    public sealed class DialogGuiTextControl : DialogGuiControlBase // <ControlName>HelpLineTextBox</ControlName>
    {
        private string _valueText;
        public string Text { get => RetrieveValueForText(); set => UpdateValueForText(value); }
        public string ToolTip { get => GetPropertyValueAsT<string>(); set => SetPropertyValueAsT(value); }

        public int Left { get => GetPropertyValueAsT<int>(); set => SetPropertyValueAsT(value); }
        public int Width { get => GetPropertyValueAsT<int>(); set => SetPropertyValueAsT(value); }

        private void UpdateValueForText(string value)
        {
            _valueText = value;
        }

        private string RetrieveValueForText()
        {
            return _valueText;
        }

        public string ValidationExpression { get => GetPropertyValueAsT<string>(); set => SetPropertyValueAsT(value); }
    }
}