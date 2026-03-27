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
        private readonly Dictionary<string, DialogExternalReferenceInfo> _externalReferences = new Dictionary<string, DialogExternalReferenceInfo>();
        private readonly DialogGuidModel _dialogModel;


        internal DialogBuilder(IServiceProvider hostServices, DialogGuidModel dialogModel) : this(hostServices, dialogModel, [])
        {
        }
        internal DialogBuilder(IServiceProvider hostServices, DialogGuidModel dialogModel, params DialogGuiControlBase[] controls)
        {
            _hostServices = hostServices ?? throw new ArgumentNullException(nameof(hostServices));
            _dialogModel = dialogModel ?? throw new ArgumentNullException(nameof(dialogModel));
            foreach (var control in controls)
            {
                AddControlCore(control.ID, control);
            }
        }

        private string _globalScriptCode { get; set; } = "";
        public DialogBuilder SetGlobalScriptCode(string dialogGlobalScriptCode)
        {
            _globalScriptCode = dialogGlobalScriptCode ?? throw new ArgumentNullException(nameof(dialogGlobalScriptCode));
            return this;
        }

        public DialogBase BuildDialog(bool gui = true)
        {
            List<string> scriptNames = new List<string>();
            StringBuilder dialogHandlerScriptCodeBuilder = new StringBuilder();
            if (gui)
            {
                foreach (KeyValuePair<string, string> script in GuiScripts)
                {
                    scriptNames.Add(script.Key);

                    dialogHandlerScriptCodeBuilder.Append($"SUB {script.Key}()");
                    if (script.Value.Length > 0)
                    {
                        if (script.Value[0] != '\n')
                            dialogHandlerScriptCodeBuilder.AppendLine();

                        foreach (string line in script.Value.SplitLines())
                        {
                            //dialogHandlerScriptCodeBuilder.Append(script.Value);
                            dialogHandlerScriptCodeBuilder.Append('\t').AppendLine(line);
                        }

                        if (script.Value[script.Value.Length - 1] != '\n')
                            dialogHandlerScriptCodeBuilder.AppendLine();
                    }
                    else
                    {
                        dialogHandlerScriptCodeBuilder.AppendLine();
                    }
                    dialogHandlerScriptCodeBuilder.AppendLine($"END SUB");
                }
            }

            return new DialogBase(_hostServices, _globalScriptCode, dialogHandlerScriptCodeBuilder.ToString(), _externalReferences, scriptNames);
        }
        private DialogBuilder AddControlCore(string controlId, DialogGuiControlBase c)
        {
            c.InitializeControl(_dialogModel, controlId);
            if (_externalReferences.ContainsKey(controlId))
            {
                throw new InvalidOperationException($"controlId:{controlId}");
            }
            _externalReferences.Add(c.ID, new DialogExternalReferenceInfo(c, []));
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

        public DialogBuilder AddExternalObject(string objectName, object objectInstance, params string[] members)
        {
            if (_externalReferences.ContainsKey(objectName))
            {
                throw new InvalidOperationException($"objectName:{objectName}");
            }
            _externalReferences.Add(objectName, new DialogExternalReferenceInfo(objectInstance, members));
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
        public string Caption { get => GetPropertyValueAsT<string>(); set => SetPropertyValueAsT(value); }
    }

    internal class DialogGuiImageControl : DialogGuiControlBase
    {
    }

    public class DialogGuiGroupBox : DialogGuiControlBase
    {
    }

    public sealed class DialogBase
    {
        public IReadOnlyDictionary<string, DialogExternalReferenceInfo> ExternalReferences { get; }
        public IServiceProvider HostServices { get; }
        public string DialogHandlerScriptCode { get; }
        public string DialogGlobalScriptCode { get; }

        public IReadOnlyCollection<string> ScriptNames { get; }

        public DialogBase(IServiceProvider hostServices, string dialogGlobalScriptCode,  string dialogHandlerScriptCode, IReadOnlyDictionary<string, DialogExternalReferenceInfo> externalReferences, IReadOnlyCollection<string> scriptNames)
        {
            HostServices = hostServices ?? throw new ArgumentNullException(nameof(hostServices));
            DialogGlobalScriptCode = dialogGlobalScriptCode ?? throw new ArgumentNullException(nameof(dialogGlobalScriptCode));
            DialogHandlerScriptCode = dialogHandlerScriptCode ?? throw new ArgumentNullException(nameof(dialogHandlerScriptCode));
            ExternalReferences = externalReferences ?? throw new ArgumentNullException(nameof(externalReferences));
            ScriptNames = scriptNames ?? throw new ArgumentNullException(nameof(scriptNames));
        }

        internal string CompleteScriptCode()
        {
            if (string.IsNullOrEmpty(DialogGlobalScriptCode))
                return DialogHandlerScriptCode;
            return new StringBuilder().AppendLine(DialogGlobalScriptCode).AppendLine(DialogHandlerScriptCode).ToString();
        }
    }

    public sealed class DialogExternalReferenceInfo
    {
        public object Instance { get; }
        public string[] Members { get; }

        public bool AddMembers => Members.Length > 0;

        public DialogExternalReferenceInfo(object instance, string[] members)
        {
            Instance = instance ?? throw new ArgumentNullException(nameof(instance));
            Members = members ?? throw new ArgumentNullException(nameof(members));
        }
    }
}