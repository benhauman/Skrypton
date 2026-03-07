using System;
using System.Collections.Frozen;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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

        public DialogBase BuildDialog(bool gui = true)
        {
            List<string> scriptNames = new List<string>();
            StringBuilder dialogCode = new StringBuilder();
            if (gui)
            {
                foreach (KeyValuePair<string, string> script in GuiScripts)
                {
                    scriptNames.Add(script.Key);

                    dialogCode.Append($"SUB {script.Key}()");
                    if (script.Value.Length > 0)
                    {
                        if (script.Value[0] != '\n')
                            dialogCode.AppendLine();
                        dialogCode.Append(script.Value);
                        if (script.Value[script.Value.Length - 1] != '\n')
                            dialogCode.AppendLine();
                    }
                    dialogCode.AppendLine($"END SUB");
                }
            }

            return new DialogBase(_hostServices, dialogCode.ToString(), _externalReferences, scriptNames);
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

        public void AddScriptCode(string scriptName, string scriptCode)
        {
            GuiScripts.Add(scriptName, scriptCode);
        }
        public string GetScriptCode(string scriptName)
        {
            return GuiScripts[scriptName];
        }
        public void FixScriptCode(string scriptName, string newCode)
        {
            GuiScripts[scriptName] = newCode;
        }

        private readonly Dictionary<string, string> GuiScripts = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
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
        public string DialogScriptCode { get; }

        public IReadOnlyCollection<string> ScriptNames { get; }

        public DialogBase(IServiceProvider hostServices, string dialogScriptCode, IReadOnlyDictionary<string, object> externalReferences, IReadOnlyCollection<string> scriptNames)
        {
            HostServices = hostServices ?? throw new ArgumentNullException(nameof(hostServices));
            DialogScriptCode = dialogScriptCode ?? throw new ArgumentNullException(nameof(dialogScriptCode));
            ExternalReferences = externalReferences ?? throw new ArgumentNullException(nameof(externalReferences));
            ScriptNames = scriptNames ?? throw new ArgumentNullException(nameof(scriptNames));
        }
    }
}