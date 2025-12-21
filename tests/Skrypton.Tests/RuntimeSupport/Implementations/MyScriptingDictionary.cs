using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using Skrypton.RuntimeSupport.Attributes;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    [SourceClassName("Dictionary")] // for TYPENAME(CreateObject("Scripting.Dictionary"))
    internal sealed class MyScriptingDictionary1  // lubo: Type.GetTypeFromProgID("Scripting.Dictionary")
    {
        //[id(0x00000005), propget]
        private readonly Dictionary<object, object> _inner = new Dictionary<object, object>();
        public int Count => _inner.Count;
        public void Add(object key, object value)
        {
            if (_inner.ContainsKey(key))
                //throw new COMException("Key already exists", unchecked((int)0x800A01C9));
                throw new InvalidOperationException("Key already exists");
            // 0x800A01C9 = VBScript runtime error 457

            _inner.Add(key, value);
        }

    }
    [SourceClassName("Dictionary")] // for TYPENAME(CreateObject("Scripting.Dictionary"))
    [ComVisible(true)] // Required because .NET can auto‑implement IDispatch when (1):COM‑visible:true, (2): interface mode:AutoDispatch and (3): DISPID(0) & DISPIDs used
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [DefaultMember("Item")]
    internal sealed class MyScriptingDictionary : System.Collections.IEnumerable//, IDispatch  // lubo: Type.GetTypeFromProgID("Scripting.Dictionary")
    {
        // Name → DISPID map
        private readonly Dictionary<string, int> _dispIds =
            new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                { "Item",       0 },   // DISPID_VALUE (default)
                { "_NewEnum",  -4 },   // DISPID_NEWENUM
                { "Add",        1 },
                { "Exists",     2 },
                { "Remove",     3 },
                { "RemoveAll",  4 },
                { "Count",      5 },
                { "Keys",       6 },
                { "Items",      7 }
            };

        private const int DISPID_VALUE = 0;
        private const int DISPID_NEWENUM = -4;


        public enum CompareMethod
        {
            BinaryCompare = 0,
            TextCompare = 1
        }
        private Dictionary<string, object> _dict;
        private CompareMethod _compareMode;

        public MyScriptingDictionary()
        {
            _compareMode = CompareMethod.BinaryCompare;
            _dict = CreateDictionary(_compareMode);
        }

        private static Dictionary<string, object> CreateDictionary(CompareMethod mode)
        {
            return new Dictionary<string, object>(
                mode == CompareMethod.TextCompare
                    ? StringComparer.OrdinalIgnoreCase
                    : StringComparer.Ordinal
            );
        }

        // --- CompareMode property (must rebuild dictionary when changed) ---
        public CompareMethod CompareMode
        {
            get => _compareMode;
            set
            {
                if (_compareMode == value)
                    return;

                var newDict = CreateDictionary(value);

                foreach (var kv in _dict)
                    newDict[kv.Key] = kv.Value;

                _dict = newDict;
                _compareMode = value;
            }
        }

        // --- Add(key, value) ---
        public void Add(string key, object value)
        {
            if (_dict.ContainsKey(key))
                throw new ArgumentException("Key already exists");

            _dict[key] = value;
        }

        // --- Exists(key) ---
        public bool Exists(string key) => _dict.ContainsKey(key);

        // --- Item(key) getter/setter (VBScript-style default property) ---
        // --- Default property: Item ---
        [DispId(0)]  // This makes it the default VBScript property
        public object Item(object key)
        {
            string k = key.ToString();

            if (!_dict.TryGetValue(k, out var value))
                throw new KeyNotFoundException("Key not found");

            return value;
        }

        [DispId(0)]
        public void Item(object key, object value)
        {
            string k = key.ToString();
            _dict[k] = value;
        }



        // --- Key(oldKey) setter (rename key) ---
        public void Key(string oldKey, string newKey)
        {
            if (!_dict.ContainsKey(oldKey))
                throw new KeyNotFoundException("Key not found");

            if (_dict.ContainsKey(newKey))
                throw new ArgumentException("New key already exists");

            var value = _dict[oldKey];
            _dict.Remove(oldKey);
            _dict[newKey] = value;
        }

        // --- Remove(key) ---
        public void Remove(string key)
        {
            if (!_dict.Remove(key))
                throw new KeyNotFoundException("Key not found");
        }

        // --- RemoveAll() ---
        public void RemoveAll() => _dict.Clear();

        // --- Count ---
        public int Count => _dict.Count;

        // --- Keys() returns array snapshot ---
        public object[] Keys()
        {
            var arr = new object[_dict.Count];
            int i = 0;
            foreach (var kv in _dict)
                arr[i++] = kv.Key;
            return arr;
        }

        // --- Items() returns array snapshot ---
        public object[] Items()
        {
            var arr = new object[_dict.Count];
            int i = 0;
            foreach (var kv in _dict)
                arr[i++] = kv.Value;
            return arr;
        }

        // --- Enumerator returns KEYS (VBScript behaviour) ---
        public System.Collections.IEnumerator GetEnumerator()
        {
            foreach (var kv in _dict)
                yield return kv.Key;
        }
        // Enumerator (DISPID -4)
        [DispId(-4)]
        public System.Collections.IEnumerator _NewEnum()
        {
            return _dict.Keys.GetEnumerator();
        }

    }

    /*
[
           uuid(EE09B103-97E0-11CF-978F-00A02463E06F),
           dual,
           oleautomation
       ]
       interface IDictionary : IDispatch
       {
           [id(0x00000000), propget] HRESULT Item([in] VARIANT Key, [out, retval] VARIANT* pRet);
           [id(0x00000000), propput] HRESULT Item([in] VARIANT Key, [in] VARIANT newVal);

           [id(0x00000001)] HRESULT Add([in] VARIANT Key, [in] VARIANT Item);
           [id(0x00000002)] HRESULT Exists([in] VARIANT Key, [out, retval] VARIANT_BOOL* pRet);
           [id(0x00000003)] HRESULT Remove([in] VARIANT Key);
           [id(0x00000004)] HRESULT RemoveAll();

           [id(0x00000005), propget] HRESULT Count([out, retval] long* pRet);

           [id(0x00000006), propget] HRESULT Keys([out, retval] VARIANT* pRet);
           [id(0x00000007), propget] HRESULT Items([out, retval] VARIANT* pRet);

           [id(0x00000008), propget] HRESULT CompareMode([out, retval] long* pRet);
           [id(0x00000008), propput] HRESULT CompareMode([in] long newVal);
       };

     */
}