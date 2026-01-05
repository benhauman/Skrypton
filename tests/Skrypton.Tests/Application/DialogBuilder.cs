using System;
using System.Collections.Generic;
using Skrypton.Tests.Application.Controls;

namespace Skrypton.Tests.Application
{

    public class DialogBuilder
    {
        private readonly Dictionary<string, object> _externalReferences = new Dictionary<string, object>();
        public DialogBase BuildDialog()
        {
            return new DialogBase(_externalReferences);
        }
        private DialogBuilder AddControlCore(string controlName, DialogGuiControlBase c)
        {
            c.InitializeControl(controlName);
            if (_externalReferences.ContainsKey(controlName))
            {
                throw new InvalidOperationException($"controlName:{controlName}");
            }
            _externalReferences.Add(c.ControlName, c);
            return this;
        }
        public DialogBuilder AddTabControl(string controlName)
        {
            return AddControlCore(controlName, new DialogGuiTabPage() { });
        }

        public DialogBuilder AddTextControl(string controlName)
        {
            return AddControlCore(controlName, new DialogGuiTextControl() { });
        }
        public DialogBuilder AddLabelControl(string controlName)
        {
            return AddControlCore(controlName, new DialogGuiLabelControl() { });
        }

        public DialogBuilder AddGroupBox(string controlName)
        {
            return AddControlCore(controlName, new DialogGuiGroupBox() { });
        }
        public DialogBuilder AddButton(string controlName)
        {
            return AddControlCore(controlName, new DialogGuiButtonControl() { });
        }

        internal DialogBuilder AddImageControl(string controlName)
        {
            return AddControlCore(controlName, new DialogGuiImageControl() { });
        }
    }

    public class DialogGuiButtonControl : DialogGuiControlBase
    {
    }

    internal class DialogGuiImageControl : DialogGuiControlBase
    {
    }

    public class DialogGuiGroupBox : DialogGuiControlBase
    {
    }

    public sealed class DialogBase
    {
        public Dictionary<string, object> Controls { get; }

        public DialogBase(Dictionary<string, object> controls)
        {
            Controls = controls ?? throw new ArgumentNullException(nameof(controls));
        }
    }
}