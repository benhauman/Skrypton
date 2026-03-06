using Newtonsoft.Json.Bson;
using System;
using System.Collections.Generic;
using System.Reflection.Metadata.Ecma335;
using System.Runtime.CompilerServices;

namespace Skrypton.Tests.Application.Controls
{
    public abstract class DialogGuiControlBase
    {
        protected DialogGuiControlBase()
        {

        }
        public string ID { get => GetPropertyValueAsT<string>(); private set => SetPropertyValueAsT(value); }
        internal void InitializeControl(string id)
        {
            ID = id ?? throw new ArgumentNullException(nameof(id));
        }


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
            { "HelpLineTabControl", () => new DialogGuiTabPage() }
        };

        internal static DialogGuiControlBase ControlFactoryCreateDialogControl(string controlTypeName)
        {
            return ControlFactories.TryGetValue(controlTypeName, out var factory)
                ? factory()
                : throw new NotImplementedException($"ControlTypeName:{controlTypeName}");

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
}