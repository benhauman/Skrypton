using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Skrypton.Tests.Application
{

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    //[ComDefaultInterface(typeof(IScriptingHelplineObject))]
    internal sealed class HLObjectInstance // see 'EblObjectWrite'
    {
        private readonly Dictionary<ObjectValueKey, ObjectValueData> _values = new Dictionary<ObjectValueKey, ObjectValueData>();

        public HLObjectInstance()
        {
        }

        public HLObjectInstance InitializeObjectInstance(bool isNew)
        {
            IsNew = isNew ? 1 : 0;
            return this;
        }

        public object IsNew { get; private set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA2204:Literals should be spelled correctly", MessageId = "GetValue")]
        [return: MarshalAs(UnmanagedType.Struct)]
        public object GetValue([In, MarshalAs(UnmanagedType.Struct)] string key, [In] int langid, [In] int ContentID, [In] int suidx, [In] int datatype)
        {
            Console.WriteLine($"GetValue('{key}', langid:{langid}', contentId:{ContentID}, suidx:{suidx}, datatype:{datatype})");
            if (_values.TryGetValue(new ObjectValueKey((string)key, ContentID, suidx), out var ov))
            {
                if (datatype == 0)
                {
                    if (ov.DataRaw == null)
                        return string.Empty;
                    return GetOutputValueText(ov, langid);
                }
                throw new NotImplementedException($"GetValue('{key}', langid:{langid}', contentId:{ContentID}, suidx:{suidx}, datatype:{datatype})");
            }
            throw new InvalidOperationException($"GetValue('{key}', langid:{langid}', contentId:{ContentID}, suidx:{suidx}, datatype:{datatype})");
            //return null;
        }
        public void SetValue([In, MarshalAs(UnmanagedType.Struct)] string key, [In] int langid, [In] int ContentID, [In] int suidx, [In] object newValue)
        {
            Console.WriteLine($"SetValue('{key}', langid:{langid}', contentId:{ContentID}, suidx:{suidx}, newValue ({(newValue?.GetType().Name)}):{newValue})");
            if (_values.TryGetValue(new ObjectValueKey((string)key, ContentID, suidx), out var ov))
            {
                ov.UpdateValue(newValue);
            }
            else
            {
                throw new InvalidOperationException($"SetValue('{key}', langid:{langid}', contentId:{ContentID}, suidx:{suidx}, newValue ({(newValue?.GetType().Name ?? "null")}):{(newValue ?? "null")})");
            }
        }
        internal HLObjectInstance RegisterValueKey<TValue>(string key, int contentId, int suidx, TValue value)
        {
            _values.Add(new ObjectValueKey(key, contentId, suidx), new ObjectValueData(typeof(TValue)).InitializeValue(value));
            return this;
        }

        private string GetOutputValueText(ObjectValueData ov, int langid)
        {
            if (ov.DataType == typeof(string))
                return (string)ov.DataRaw;
            throw new NotImplementedException($"{ov.DataType.Name}");
        }

        [DebuggerDisplay("{Key}")]
        private sealed class ObjectValueKey : Tuple<string, int, int>
        {
            public ObjectValueKey(string key, int contentId, int suidx)
                : base(key, contentId, suidx)
            {
            }
            public string Key => Item1;
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