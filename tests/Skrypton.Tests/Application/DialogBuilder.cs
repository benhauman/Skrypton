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
        private readonly Dictionary<string, ScriptExternalReferenceInfo> _externalReferences = new Dictionary<string, ScriptExternalReferenceInfo>();
        private readonly Dictionary<string, DialogGuiControlBase> _controls = new Dictionary<string, DialogGuiControlBase>();
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
                foreach (KeyValuePair<string, ScriptInfo> script in GuiScripts)
                {
                    scriptNames.Add(script.Key);
                    ScriptInfo scriptInfo = script.Value;

                    dialogHandlerScriptCodeBuilder.Append($"SUB {script.Key}()");
                    if (scriptInfo.Code.Length > 0)
                    {
                        if (scriptInfo.Code[0] != '\n')
                            dialogHandlerScriptCodeBuilder.AppendLine();

                        foreach (string line in scriptInfo.Code.SplitLines())
                        {
                            //dialogHandlerScriptCodeBuilder.Append(scriptInfo.Code);
                            dialogHandlerScriptCodeBuilder.Append('\t').AppendLine(line);
                        }

                        if (scriptInfo.Code[scriptInfo.Code.Length - 1] != '\n')
                            dialogHandlerScriptCodeBuilder.AppendLine();
                    }
                    else
                    {
                        dialogHandlerScriptCodeBuilder.AppendLine();
                    }
                    dialogHandlerScriptCodeBuilder.AppendLine($"END SUB");
                }
            }

            return new DialogBase(_hostServices, _globalScriptCode, dialogHandlerScriptCodeBuilder.ToString(), _externalReferences, scriptNames, _controls);
        }
        private DialogBuilder AddControlCore(string controlId, DialogGuiControlBase c)
        {
            c.InitializeControl(_dialogModel, controlId);
            if (_externalReferences.ContainsKey(controlId))
            {
                throw new InvalidOperationException($"controlId:{controlId}");
            }
            _externalReferences.Add(c.ID, new ScriptExternalReferenceInfo(c, []));
            _controls.Add(c.ID, c);
            return this;
        }
        public DialogBuilder AddTabControl(string controlId)
        {
            return AddControlCore(controlId, new DialogGuiTabPage() { });
        }

        public DialogBuilder AddComboBoxControl(string controlId)
        {
            return AddControlCore(controlId, new DialogGuiComboBoxControl() { });
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
            _externalReferences.Add(objectName, new ScriptExternalReferenceInfo(objectInstance, members));
            return this;
        }

        public void AddScriptCode(string scriptName, string scriptCode)
        {
            GuiScripts.Add(scriptName, new ScriptInfo(scriptCode));
        }
        public string GetScriptCode(string scriptName)
        {
            return GuiScripts[scriptName].Code;
        }
        public void FixScriptCode(string scriptName, string newCode)
        {
            GuiScripts[scriptName] = new ScriptInfo(newCode);
        }

        private readonly Dictionary<string, ScriptInfo> GuiScripts = new Dictionary<string, ScriptInfo>(StringComparer.OrdinalIgnoreCase);

        private sealed class ScriptInfo
        {
            public string Code { get; set; }
            public ScriptInfo(string code)
            {
                Code = code ?? throw new ArgumentNullException(nameof(code));
            }
            private readonly Dictionary<string, string> _usedBy = new Dictionary<string, string>();
            public void AddControlEvent(string controlName, string eventName)
            {
                _usedBy.Add(controlName, eventName);
            }
        }
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
        public string Caption { get => GetPropertyValueAsT<string>(); set => SetPropertyValueAsT(value); }
    }

    public sealed class DialogBase
    {
        private readonly IReadOnlyDictionary<string, DialogGuiControlBase> _controls;
        public IReadOnlyDictionary<string, ScriptExternalReferenceInfo> ExternalReferences { get; }
        public IServiceProvider HostServices { get; }
        public string DialogHandlerScriptCode { get; }
        public string DialogGlobalScriptCode { get; }

        public IReadOnlyCollection<string> ScriptNames { get; }

        public DialogBase(IServiceProvider hostServices, string dialogGlobalScriptCode, string dialogHandlerScriptCode, IReadOnlyDictionary<string, ScriptExternalReferenceInfo> externalReferences, IReadOnlyCollection<string> scriptNames, IReadOnlyDictionary<string, DialogGuiControlBase> controls)
        {
            _controls = controls ?? throw new ArgumentNullException(nameof(controls));
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

        public void CollectControlEventScriptNames(Action<DialogGuiControlBase, string, string> collector)
        {
            foreach (var c in _controls.Values)
            {
                c.CollectControlEventScriptNames(collector);
            }
        }
    }

    public sealed class ScriptExternalReferenceInfo
    {
        public object Instance { get; }
        public string[] Members { get; }

        public bool AddMembers => Members.Length > 0;

        public ScriptExternalReferenceInfo(object instance, string[] members)
        {
            Instance = instance ?? throw new ArgumentNullException(nameof(instance));
            Members = members ?? throw new ArgumentNullException(nameof(members));
        }
    }
}