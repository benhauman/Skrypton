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

        private ShowControlType _valueShowControl;
        public short ShowControl // see ShowControlType
        {
            get => (short)_valueShowControl;
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
        };

        internal static DialogGuiControlBase ControlFactoryCreateDialogControl(string controlTypeName)
        {
            return ControlFactories.TryGetValue(controlTypeName, out var factory)
                ? factory()
                : new DialogGuiUnknownControl(controlTypeName)
                //: throw new NotImplementedException($"ControlTypeName:{controlTypeName}")
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
    }

    public sealed class DialogGuiComboBoxControl : DialogGuiControlBase
    {
        public void SelectItem(string value, bool mustEqual = true)
        {
            Console.WriteLine($"[ComboBox]({ID}).SelectItem({value})");
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