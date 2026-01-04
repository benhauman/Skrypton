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


        private bool _valueShowControl;
        public bool ShowControl
        {
            get => _valueShowControl;
            set => _valueShowControl = value;
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
}