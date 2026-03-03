using System;
using System.Collections.Frozen;
using System.Collections.Generic;
using System.Linq;
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

        internal DialogBuilder AddExternalObject(string objectName, object objectInstance)
        {
            if (_externalReferences.ContainsKey(objectName))
            {
                throw new InvalidOperationException($"objectName:{objectName}");
            }
            _externalReferences.Add(objectName, objectInstance);
            return this;
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
        public IReadOnlyDictionary<string, object> ExternalReferences { get; }

        public DialogBase(IReadOnlyDictionary<string, object> externalReferences)
        {
            ExternalReferences = externalReferences ?? throw new ArgumentNullException(nameof(externalReferences));
        }
    }
}