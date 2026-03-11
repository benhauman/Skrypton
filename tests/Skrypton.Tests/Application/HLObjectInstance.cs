using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;

namespace Skrypton.Tests.Application
{

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    //[ComDefaultInterface(typeof(IScriptingHelplineObject))]
    internal sealed class HLObjectInstance // see 'EblObjectWrite' // see /Core/Common/Client/RuntimeObject/HelplineObject.cs
    {
        private readonly string _traceName;
        private readonly Dictionary<ObjectValueKey, ObjectValueData> _values = new Dictionary<ObjectValueKey, ObjectValueData>();

        public HLObjectInstance(string traceName = "")
        {
            _traceName = string.IsNullOrEmpty(traceName) ? traceName : $"|{traceName}|";
        }
        private int? _objectId;
        private string _objectDefName;
        public HLObjectInstance InitializeObjectInstance(bool isNew, int? objectId = null, string objectDefName = null)
        {
            IsNew = isNew ? 1 : 0;
            _objectId = objectId;
            _objectDefName = objectDefName;
            return this;
        }

        public int objID()
        {
            return _objectId.HasValue ? _objectId.Value : throw new InvalidOperationException($"{_traceName}Id not set.");
        }
        public new object GetType() // definition name, defid, basetype
        {
            return _objectDefName != null ? _objectDefName : throw new InvalidOperationException($"{_traceName}Type not set.");
        }

        public object IsNew { get; private set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA2204:Literals should be spelled correctly", MessageId = "GetValue")]
        [return: MarshalAs(UnmanagedType.Struct)]
        public object GetValue([In, MarshalAs(UnmanagedType.Struct)] string key, [In] int langid, [In] int ContentID, [In] int suidx, [In] int datatype)
        {
            Console.WriteLine($"{_traceName}GetValue('{key}', langid:{langid}', contentId:{ContentID}, suidx:{suidx}, datatype:{datatype})");
            var vk = new ObjectValueKey((string)key, ContentID, suidx);
            if (_values.TryGetValue(vk, out var ov))
            {
                if (datatype == 0)
                {
                    if (ov.DataRaw == null)
                        return string.Empty;
                    return GetOutputValueText(vk, ov, langid);
                }
                if (datatype == 1)
                {
                    if (ov.DataRaw == null)
                        return null;
                    if (ov.DataRaw is int)
                        return ov.DataRaw;
                }
                throw new NotImplementedException($"GetValue('{key}', langid:{langid}', contentId:{ContentID}, suidx:{suidx}, datatype:{datatype})");
            }
            if (datatype == 0)
                return "";
            throw new InvalidOperationException($"{_traceName}GetValue('{key}', langid:{langid}', contentId:{ContentID}, suidx:{suidx}, datatype:{datatype})");
        }
        public void SetValue([In, MarshalAs(UnmanagedType.Struct)] string key, [In] int langid, [In] int ContentID, [In] int suidx, [In] object newValue)
        {
            Console.WriteLine($"{_traceName}SetValue('{key}', langid:{langid}', contentId:{ContentID}, suidx:{suidx}, newValue ({(newValue?.GetType().Name)}):{newValue})");
            if (_values.TryGetValue(new ObjectValueKey((string)key, ContentID, suidx), out var ov))
            {
                ov.UpdateValue(newValue);
            }
            else
            {
                throw new InvalidOperationException($"{_traceName}SetValue('{key}', langid:{langid}, contentId:{ContentID}, suidx:{suidx}, newValue ({(newValue?.GetType().Name ?? "null")}):{(newValue ?? "null")})");
            }
        }
        public int GetItemCount(int flags, object assocdef)//(0, 130)' not found
        {
            return 0;
        }
        public object HasContent(object attributeKey, int contentid, int suidx)
        {
            Console.WriteLine($"SetValue('{attributeKey}', contentId:{contentid}, suidx:{suidx}");
            if (_values.TryGetValue(new ObjectValueKey((string)attributeKey, contentid, suidx), out var ov))
            {
                return 1; // or count for multiples
            }
            else
            {
                throw new InvalidOperationException($"SetValue('{attributeKey}', contentId:{contentid}, suidx:{suidx}");
            }
        }
        public object IsReadOnly(object key, int suidx) // 1:true, 0:false
        {
            Console.WriteLine($"IsReadOnly('{key}', suidx:{suidx}");
            if (_values.TryGetValue(new ObjectValueKey((string)key, contentId: 0, suidx), out var ov))
            {
                return false;
            }
            else
            {
                throw new InvalidOperationException($"IsReadOnly('{key}', suidx:{suidx}");
            }
        }

        private readonly Dictionary<int, object> _sus = new Dictionary<int, object>();
        public HLObjectInstance RegisterServiceUnitIndex(int suidx)
        {
            _sus.Add(suidx, null);
            return this;
        }
        public int[] GetSvcUnitIndices()
        {
            return _sus.Keys.OrderBy(x => x).ToArray();
        }

        internal HLObjectInstance RegisterValueKey<TValue>(string key, int contentId, int suidx, TValue value)
        {
            _values.Add(new ObjectValueKey(key, contentId, suidx), new ObjectValueData(typeof(TValue)).InitializeValue(value));
            return this;
        }

        private string GetOutputValueText(ObjectValueKey vk, ObjectValueData ov, int langid)
        {
            if (ov.DataType == typeof(string))
                return (string)ov.DataRaw;
            if (ov.DataType == typeof(int))
            {
                return ov.DataRaw == null ? "" : ((int)ov.DataRaw).ToString(CultureInfo.InvariantCulture);
            }
            throw new NotImplementedException($"[{vk.DebugText}]({ov.DataType.Name}):{ov.DataRaw}");
        }

        [DebuggerDisplay("{Key}")]
        private sealed class ObjectValueKey : Tuple<string, int, int>
        {
            public ObjectValueKey(string key, int contentId, int suidx)
                : base(key, contentId, suidx)
            {
                DebugText = $"{key}, c:{contentId}, sux:{suidx}";
            }
            public string Key => Item1;

            internal readonly string DebugText;
        }

        [DebuggerDisplay("{DataType.Name}:{HasValue}:{DataRaw}")]
        private sealed class ObjectValueData
        {
            public ObjectValueData(Type dataType)
            {
                DataType = dataType ?? throw new ArgumentNullException(nameof(dataType));
            }
            public bool HasValue { get; private set; }
            public Type DataType { get; private set; }
            public object DataRaw { get; private set; }

            public ObjectValueData InitializeValue(object value)
            {
                SetData(value);
                return this;
            }
            private void SetData(object value)
            {
                if (value == null)
                {
                    HasValue = false;
                    DataRaw = null;
                }
                else
                {
                    object newValue = value;
                    if (value.GetType() != DataType)
                    {
                        // todo : convert it if needed
                        throw new NotImplementedException($"Expect:{DataType.Name}, actual:{value.GetType().Name}");
                    }
                    HasValue = true;
                    DataRaw = newValue;
                }
            }
            public void UpdateValue(object newValue)
            {
                SetData(newValue);
            }
        }
    }
}