using System.Runtime.InteropServices;

namespace Skrypton.Tests.Application.Controls
{
    [ComVisible(true)]
    internal sealed class DialogGuiTabPage : DialogGuiControlBase // <ControlName>HelpLineTabPage</ControlName>
    {
        private string _valueCaption;
        public string Caption
        {
            get => _valueCaption;
            set => _valueCaption = value;
        }
    }
}