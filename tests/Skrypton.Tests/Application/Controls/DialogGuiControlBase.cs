using Newtonsoft.Json.Bson;
using Skrypton.Tests.RuntimeSupport.Implementations;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection.Metadata.Ecma335;
using System.Runtime.CompilerServices;

namespace Skrypton.Tests.Application.Controls
{
    [DebuggerDisplay("[{GetType().Name}]ID:{ID}")]
    public abstract class DialogGuiControlBase// : IReflectOnClrType
    {
        protected DialogGuiControlBase()
        {

        }
        internal void InitializeControl(DialogGuidModel dialogModel, string id)
        {
            ID = id ?? throw new ArgumentNullException(nameof(id));
            _model = dialogModel ?? throw new ArgumentNullException(nameof(dialogModel));
        }
        private DialogGuidModel _model;
        internal DialogGuidModel model => _model ?? throw new InvalidOperationException("model not set");

        public string ID { get => GetPropertyValueAsT<string>(); private set => SetPropertyValueAsT(value); }
        public bool Disabled { get => GetPropertyValueAsT<bool>(); set => SetPropertyValueAsT(value); }
        public bool UiActive { get => GetPropertyValueAsT<bool>(); set => SetPropertyValueAsT(value); }

        public SymbolName SymbolName { get => GetPropertyValueAsT<SymbolName>(); set => SetPropertyValueAsT(value); }
        public string AttributeKey { get => GetPropertyValueAsT<string>(); set => SetPropertyValueAsT(value); }

        private ShowControlType _valueShowControl;
        public short ShowControl // see ShowControlType
        {
            get => (short)_valueShowControl;
            set => _valueShowControl = (ShowControlType)value;
        }

        public string TextColor { get => GetPropertyValueAsT<string>(); set => SetPropertyValueAsT(value); }
        public string BackColor { get => GetPropertyValueAsT<string>(); set => SetPropertyValueAsT(value); }
        public bool RequestFocus { get => GetPropertyValueAsT<bool>(); set => SetPropertyValueAsT(value); }
        public bool Required { get => GetPropertyValueAsT<bool>(); set => SetPropertyValueAsT(value); }


        private Dictionary<string, object> _properties = new Dictionary<string, object>();

        //public virtual void InitControlProperty(string propertyName, object propertyValue)
        //{
        //    SetPropertyValue(propertyName, propertyValue);
        //}

        //private static readonly HashSet<string> IgnorePropertyMap = new HashSet<string>(StringComparer.Ordinal) { "OnLoad", "OnSave", "OnUpdate" };
        //internal virtual bool ShouldIgnoreValueForProperty(string propertyName) => IgnorePropertyMap.Contains(propertyName);

        protected T GetPropertyValueAsT<T>([CallerMemberName] string propertyName = "")
        {
            return _properties.TryGetValue(propertyName, out object val) ? (T)val : default(T);
        }
        protected void SetPropertyValueAsT<T>(T newValue, [CallerMemberName] string propertyName = "")
        {
            if (!_properties.TryAdd(propertyName, newValue))
            {
                _properties[propertyName] = newValue;
            }
        }

        internal Action<object> ShouldInitValueForProperty(string propertyName)
        {
            var pi = GetType().GetProperty(propertyName);
            return pi == null ? null : (v) =>
            {
                //if (_properties.ContainsKey(pi.Name))
                //    throw new InvalidOperationException();
                _properties.Add(pi.Name, v);
            };
        }

        private static readonly Dictionary<string, Func<DialogGuiControlBase>> ControlFactories = new Dictionary<string, Func<DialogGuiControlBase>>() {
            { "HelpLineDialogControl", () => new DialogGuiRoot() },
            { "HelpLineGroupBox", () => new DialogGuiGroupBox() },
            { "HelpLineTabControl", () => new DialogGuiTabPage() },
            { "HelpLineTextBox", () => new DialogGuiTextControl() },
            { "HelpLineComboBox", () => new DialogGuiComboBoxControl() },
            { "HelpLineSearchButton", () => new DialogGuiSearchButtonControl() },
            { "HelpLineTreeSelControl", () => new HelpLineTreeSelControl()},
            { "HelpLineLabel", () => new DialogGuiLabelControl() },
            { "HelpLineDateTimeControl", () => new DialogGuiDateTimeControl() },
            { "HelpLineTabPage", () => new DialogGuiHelpLineTabPageControl() },
            { "HelpLineTasksControl", () => new DialogGuiHelpLineTasksControl() },
            { "HelpLineTableControl", () => new DialogGuiHelpLineTableControl() },
            { "HelpLineCompound", () => new DialogGuiCompoundControl() },
            { "HelpLineAttachmentControl", () => new DialogGuiAttachmentControl() },
            { "HelpLineTreeComboBox", () => new DialogGuiTreeComboBoxControl() },
            { "HelpLineCheckBox", () => new DialogGuiCheckBoxControl() },
            { "HelpLineButton", () => new DialogGuiButtonControl() },
            { "HelpLineSUControl", () => new DialogGuiSUControl() },
            { "HelpLineComplexText", () => new DialogGuiComplexTextControl() }
        };

        internal static DialogGuiControlBase ControlFactoryCreateDialogControl(string controlTypeName)
        {
            return ControlFactories.TryGetValue(controlTypeName, out var factory)
                ? factory()
                //: new DialogGuiUnknownControl(controlTypeName)
                : throw new NotImplementedException($"ControlTypeName:{controlTypeName}")
                ;

        }
        //protected virtual void SetPropertyValue(string propertyName, object propertyValue)

        //protected virtual void SetPropertyValue(string propertyName, object propertyValue)
        //{
        //    if (propertyName == "ID")
        //    {
        //        string valueID = (string)propertyValue;
        //        if (valueID != ID)
        //            throw new InvalidOperationException($"Invalid property value. Expected:{ID}, actual{valueID}");
        //    }
        //    else
        //    {
        //        //if (ShouldIgnoreValueForProperty(propertyName))
        //        //{
        //        //    // put a breakpoint here
        //        //}
        //        //else
        //        var pi = GetType().GetProperty(propertyName);
        //        pi.SetValue(this, propertyValue);
        //        //{
        //        //    throw new InvalidOperationException($"[{GetType().Name}] Unknown control property name:{propertyName}:{propertyValue}");
        //        //}
        //        //WritePropertyValue(propertyName, propertyValue);
        //    }
        //}
        //protected void WritePropertyValue(string propertyName, object propertyValue)
        //{
        //    _properties.Add(propertyName, propertyValue);
        //}
    }
    public enum ShowControlType
    {
        GuiOnly = 0,
        Always = 1,
        WebOnly = 2,
        Never = 3
    }

    [DebuggerDisplay("{ControlTypeName}")]
    public sealed class DialogGuiUnknownControl : DialogGuiControlBase
    {
        public string ControlTypeName { get; }

        public DialogGuiUnknownControl(string controlTypeName)
        {
            ControlTypeName = controlTypeName;
        }

        //public void SelectTreeItem(object treeItem)
        //{
        //    Console.WriteLine($"[UnknownControl]({ID}).SelectTreeItem({treeItem})");
        //    throw new NotSupportedException($"[UnknownControl:{ControlTypeName}]({ID}).SelectTreeItem({treeItem})");

        //}
        //public void ExpandTreeItem(object treeItem)
        //{
        //    Console.WriteLine($"[UnknownControl]({ID}).SelectTreeItem({treeItem})");
        //    throw new NotSupportedException($"[UnknownControl{ControlTypeName}]({ID}).SelectTreeItem({treeItem})");
        //}
    }
    public sealed class HelpLineTreeSelControl : DialogGuiControlBase
    {
        public void SelectTreeItem(object treeItem)
        {
            Console.WriteLine($"[TreeSel]({ID}).SelectTreeItem({treeItem})");
        }
        public void ExpandTreeItem(object treeItem)
        {
            Console.WriteLine($"[TreeSel]({ID}).SelectTreeItem({treeItem})");
        }
    }

    public sealed class DialogGuiComplexTextControl : DialogGuiControlBase
    {
    }

    public sealed class DialogGuiSUControl : DialogGuiControlBase
    {
    }

    public sealed class DialogGuiCheckBoxControl : DialogGuiControlBase
    {
    }

    public sealed class DialogGuiTreeComboBoxControl : DialogGuiControlBase
    {
    }

    public sealed class DialogGuiAttachmentControl : DialogGuiControlBase
    {
    }

    public sealed class DialogGuiCompoundControl : DialogGuiControlBase
    {
    }
    public sealed class DialogGuiHelpLineTableControl : DialogGuiControlBase
    {
    }

    public sealed class DialogGuiHelpLineTabPageControl : DialogGuiControlBase
    {

    }
    public sealed class DialogGuiHelpLineTasksControl : DialogGuiControlBase
    {

    }
    public sealed class DialogGuiDateTimeControl : DialogGuiControlBase
    {
        public void DeleteContent()
        {
            Console.WriteLine($"[DateTimeControl]({ID}).DeleteContent");
        }
    }

    public sealed class DialogGuiComboBoxControl : DialogGuiControlBase
    {
        public void SelectItem(string value, bool mustEqual = true)
        {
            Console.WriteLine($"[ComboBox]({ID}).SelectItem({value})");
        }

        public int GetCurSel()
        {
            Console.WriteLine($"[ComboBox]({ID}).GetCurSel()");
            return 2;
        }

        public string Text
        {
            get
            {
                var hlobj = model.GetHelpLineObject(SymbolName.Name);
                return (string)hlobj.GetValue(AttributeKey, 0, 0, 0, 0);
                //return "";
            }
            set
            {
                var hlobj = model.GetHelpLineObject(SymbolName.Name);
                hlobj.SetValue(AttributeKey, 0, 0, 0, value);
            }
        }

        public void ResetContent()
        {
            Console.WriteLine($"[ComboBox]({ID}).ResetContent");
        }

        public void AddItem(string item)
        {
            Console.WriteLine($"[ComboBox]({ID}).AddItem('{item}')");
        }
    }

    public sealed class DialogGuiSearchButtonControl : DialogGuiControlBase
    {
        public DialogGuiSearchButtonControl()
        {

        }

        public enum SearchState
        {
            Undefined = 0,
            Design = 1,
            Execute = 2,
            Reset = 3
        }

        public int GetSearchState()
        {
            return (int)SearchState.Execute;
        }

        public string Caption { get => GetPropertyValueAsT<string>(); set => SetPropertyValueAsT(value); }

        public object GetObject(string symbolName, bool search)
        {
            if (search)
                return model.DialogUserControl.GetHelpLineTempObject(symbolName);
            else
                return model.DialogUserControl.GetHelpLineObject(symbolName);
        }

    }

}