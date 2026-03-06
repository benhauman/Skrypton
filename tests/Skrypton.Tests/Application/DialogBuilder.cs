using System;
using System.Collections.Frozen;
using System.Collections.Generic;
using System.Linq;
using Skrypton.Tests.Application.Controls;

namespace Skrypton.Tests.Application
{

    public class DialogBuilder
    {
        private readonly IServiceProvider _hostServices;
        private readonly Dictionary<string, object> _externalReferences = new Dictionary<string, object>();

        public DialogBuilder(IServiceProvider hostServices) : this(hostServices, [])
        {
        }
        public DialogBuilder(IServiceProvider hostServices, params DialogGuiControlBase[] controls)
        {
            _hostServices = hostServices ?? throw new ArgumentNullException(nameof(hostServices));
            foreach (var control in controls)
            {
                AddControlCore(control.ID, control);
            }
        }

        public DialogBase BuildDialog()
        {
            return new DialogBase(_hostServices, _externalReferences);
        }
        private DialogBuilder AddControlCore(string controlId, DialogGuiControlBase c)
        {
            c.InitializeControl(controlId);
            if (_externalReferences.ContainsKey(controlId))
            {
                throw new InvalidOperationException($"controlId:{controlId}");
            }
            _externalReferences.Add(c.ID, c);
            return this;
        }
        public DialogBuilder AddTabControl(string controlId)
        {
            return AddControlCore(controlId, new DialogGuiTabPage() { });
        }

        public DialogBuilder AddTextControl(string controlId)
        {
            return AddControlCore(controlId, new DialogGuiTextControl() { });
        }
        public DialogBuilder AddLabelControl(string controlId)
        {
            return AddControlCore(controlId, new DialogGuiLabelControl() { });
        }

        public DialogBuilder AddGroupBox(string controlId)
        {
            return AddControlCore(controlId, new DialogGuiGroupBox() { });
        }
        public DialogBuilder AddButton(string controlId)
        {
            return AddControlCore(controlId, new DialogGuiButtonControl() { });
        }

        internal DialogBuilder AddImageControl(string controlId)
        {
            return AddControlCore(controlId, new DialogGuiImageControl() { });
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

    public sealed class DialogGuiRoot : DialogGuiControlBase
    {

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
        public IServiceProvider HostServices { get; }

        public DialogBase(IServiceProvider hostServices, IReadOnlyDictionary<string, object> externalReferences)
        {
            HostServices = hostServices ?? throw new ArgumentNullException(nameof(hostServices));
            ExternalReferences = externalReferences ?? throw new ArgumentNullException(nameof(externalReferences));
        }
    }
}