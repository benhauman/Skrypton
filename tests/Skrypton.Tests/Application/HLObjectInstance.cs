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
    public sealed class HLObjectInstance // see 'EblObjectWrite' // see /Core/Common/Client/RuntimeObject/HelplineObject.cs
    {
        private readonly string _traceName;
        private readonly Dictionary<ObjectValueKey, ObjectValueData> _values = new Dictionary<ObjectValueKey, ObjectValueData>();

        public HLObjectInstance(string traceName = "")
        {
            _traceName = string.IsNullOrEmpty(traceName) ? traceName : $"|{traceName}|";
        }
        private CultureInfo _culture;
        private int? _objectId;
        private int? _objectDefId;
        private string _objectDefName;
        public HLObjectInstance InitializeObjectInstance(bool isNew, CultureInfo culture, int? objectId = null, int? objectDefId = null, string objectDefName = null)
        {
            IsNew = isNew ? 1 : 0;
            _culture = culture ?? throw new ArgumentNullException(nameof(culture));
            _objectId = objectId;
            _objectDefId = objectDefId;
            _objectDefName = objectDefName;
            return this;
        }

        public int GetID() => objID();

        public int objID()
        {
            return _objectId.HasValue ? _objectId.Value : throw new InvalidOperationException($"{_traceName}Id not set.");
        }
        internal int TestGetObjectDefId()
        {
            return _objectDefId.HasValue ? _objectDefId.Value : throw new InvalidOperationException($"{_traceName}DefId not set.");
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
                else
                {
                    // datatype:2  // Bei Schlagworten den vollständigen Pfad liefern
                    if (ov.DataRaw == null)
                        return null;
                    if (ov.DataRaw is int)
                        return ov.DataRaw;
                    if (ov.DataRaw is string)
                        return ov.DataRaw;
                }
                throw new NotImplementedException($"GetValue('{key}', langid:{langid}', contentId:{ContentID}, suidx:{suidx}, datatype:{datatype}). VT:{ov.DataRaw?.GetType()}, V:{ov.DataRaw}");
            }
            if (datatype == 0)
                return "";
            if (datatype == 1)
                return 1234567;// why not
            if (datatype == 2)
                return 7654321;// why not
            throw new InvalidOperationException($"{_traceName}GetValue('{key}', langid:{langid}', contentId:{ContentID}, suidx:{suidx}, datatype:{datatype})");
        }
        public void SetValue([In, MarshalAs(UnmanagedType.Struct)] string key, [In] int langid, [In] int ContentID, [In] int suidx, [In] object newValue)
        {
            Console.WriteLine($"{_traceName}SetValue('{key}', langid:{langid}', contentId:{ContentID}, suidx:{suidx}, newValue ({(newValue?.GetType().Name)}):{newValue})");
            var vk = new ObjectValueKey((string)key, ContentID, suidx);
            if (_values.TryGetValue(vk, out var ov))
            {
                ov.UpdateValue(newValue);
            }
            else
            {
                _values.Add(vk, new ObjectValueData(newValue?.GetType() ?? typeof(string)).NewValue(newValue));
                //                throw new InvalidOperationException($"{_traceName}SetValue('{key}', langid:{langid}, contentId:{ContentID}, suidx:{suidx}, newValue ({(newValue?.GetType().Name ?? "null")}):{(newValue ?? "null")})");
            }
        }
        public int GetItemCount(int flags, object assocdef)//(0, 130)' not found
        {
            Console.WriteLine($"{_traceName}GetItemCount(flags:{flags}, assocdef:{assocdef})");
            return 0;
        }
        public object[] GetItems(int flags, int nfirst, int nlast, object assocdef)
        {
            Console.WriteLine($"{_traceName}GetItems(flags:{flags}, nfirst:{nfirst}, nlast:{nlast}, assocdef:{assocdef})");
            if (flags == 0x10000)
            {
                if (_associatedItemsA.TryGetValue((string)assocdef, out var items))
                    return items.Values.Select(x => (object)x).ToArray();
                return [];
            }
            else
            {
                object[] hlObjects = [];
                return hlObjects;
            }
        }
        private Dictionary<string, Dictionary<int, HLObjectInstance>> _associatedItemsA = new Dictionary<string, Dictionary<int, HLObjectInstance>>();
        public HLObjectInstance TestRegisterAssociateItemA(string assocdef, HLObjectInstance item)
        {
            if (!_associatedItemsA.TryGetValue(assocdef, out var items))
            {
                items = new Dictionary<int, HLObjectInstance>();
                _associatedItemsA.Add(assocdef, items);
            }
            items.Add(item._objectId!.Value, item);
            return this;
        }
        public void RemoveItem(int flags, object val, object assocdef)
        {
            Console.WriteLine($"{_traceName}SetValue(flags:{flags}, val:{val}, assocdef:{assocdef})");
            //RemoveItemEx(flags, val, 0, assocdef);
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
            var vk = new ObjectValueKey((string)key, contentId: 0, suidx);
            if (_values.TryGetValue(vk, out var ov))
            {
                return false;
            }
            else
            {
                if (vk.AttributePath == "CASEINFO.REACTIONTIME")
                    return false;
                if (vk.AttributePath.StartsWith("HLOBJECTINFO."))
                    return false;
                if (vk.AttributePath.StartsWith("CASEINFO."))
                    return false;
                if (vk.AttributePath.StartsWith("SUINFO."))
                    return false;
                //throw new InvalidOperationException($"IsReadOnly('{key}', suidx:{suidx}");
                return true;
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

        public int GetSvcUnitCount()
        {
            return _sus.Count;
        }
        public int[] GetContentIDs(object key, int suidx)
        {
            Console.WriteLine($"{_traceName}GetContentIDs(key:{key}, suidx:{suidx}')");
            return [];
        }

        public void AddItem(int flags, object val, object assocdef)
        {
            Console.WriteLine($"{_traceName}AddItem(flags:{flags}, val:{val}', assocdef:{assocdef})");
            //AddItemEx(flags, val, 0, assocdef);
        }

        public int[] GetAttachmentKeys(object key, int suidx)
        {
            string attributeKey = (string)key;
            if (attributeKey != "HLOBJECTINFO.ATTACHMENT" && attributeKey != "SUINFO.ATTACHMENT")
                throw new ArgumentException($"Invalid attribute key:{key}");
            int[] attachmentIds = [101111, 102222];
            return attachmentIds;
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
            if (ov.DataType == typeof(short))
            {
                return ov.DataRaw == null ? "" : ((short)ov.DataRaw).ToString(CultureInfo.InvariantCulture);
            }
            if (ov.DataType == typeof(bool))
            {
                return ov.DataRaw == null ? "" : ((bool)ov.DataRaw) ? "1" : "0"; // see 'BitFormatter'
            }
            if (ov.DataType == typeof(DateTime))
            {
                if (ov.DataRaw == null)
                    return "";
                // see 'DateTimeFormatter'
                var dt = (DateTime)ov.DataRaw;
                if (langid == 0)
                    return dt.ToString("yyyy-MM-dd HH:mm:ss");
                return dt.ToString(_culture!);
            }
            throw new NotImplementedException($"[{vk.DebugText}]({ov.DataType.Name}):{ov.DataRaw}");
        }

        [DebuggerDisplay("{AttributePath}")]
        private sealed class ObjectValueKey : Tuple<string, int, int>
        {
            public ObjectValueKey(string key, int contentId, int suidx)
                : base(key, contentId, suidx)
            {
                DebugText = $"{key}, c:{contentId}, sux:{suidx}";
            }
            public string AttributePath => Item1;

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
                        if (DataType == typeof(int) && value is string stringValueInt && stringValueInt.Length == 0)
                        {
                            HasValue = false;
                            DataRaw = null;
                        }
                        else if (DataType == typeof(DateTime) && value is string stringValueDateTime)
                        {
                            // 2026-04-01 22:16:58
                            var dt = DateTime.Parse(stringValueDateTime);
                            HasValue = true;
                            DataRaw = dt;
                        }
                        else
                        {
                            // todo : convert it if needed
                            throw new NotImplementedException($"Expect:{DataType.Name}, actual:{value.GetType().Name}. Value:{value}");
                        }
                    }
                    else
                    {
                        HasValue = true;
                        DataRaw = newValue;
                    }
                }
            }
            public void UpdateValue(object newValue)
            {
                SetData(newValue);
            }
            public ObjectValueData NewValue(object newValue)
            {
                SetData(newValue);
                return this;
            }
        }
    }
}