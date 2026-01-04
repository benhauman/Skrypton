using System;

namespace Skrypton.Tests.Application.Controls
{
    public abstract class DialogGuiControlBase
    {
        public string ControlName { get; private set; }
        internal void InitializeControl(string controlName)
        {
            ControlName = controlName ?? throw new ArgumentNullException(nameof(controlName));
        }


        private ShowControlType _valueShowControl;
        public byte ShowControl // see ShowControlType
        {
            get => (byte)_valueShowControl;
            set => _valueShowControl = (ShowControlType)value;
        }

        private string _valueBackColor;
        public string BackColor
        {
            get => _valueBackColor;
            set => _valueBackColor = value;
        }

        private bool _valueRequestFocus;
        public bool RequestFocus
        {
            get => _valueRequestFocus;
            set => _valueRequestFocus = value;
        }
    }
    public enum ShowControlType
    {
        GuiOnly = 0,
        Always = 1,
        WebOnly = 2,
        Never = 3
    }
}