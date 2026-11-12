using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using Skrypton.RuntimeSupport.Attributes;
using Skrypton.RuntimeSupport.Exceptions;

namespace Skrypton.RuntimeSupport.Components
{
    /// <summary>
    /// A managed implementation of the Windows script runtime's (scrrun.dll) Scripting.Dictionary, so that CreateObject("Scripting.Dictionary")
    /// works without COM (see issue #50). It follows scrrun's semantics:
    /// - keys are variants: numeric keys are compared by value (so 1 and CLng(1) are the same key) but a number and a string are different keys
    ///   (1 and "1"), Empty and object references (compared by reference) are valid keys
    /// - CompareMode 0 (binary) or 1 (text, case-insensitive) applies to string keys and may only be changed while the dictionary is empty
    /// - Keys / Items / For Each are in insertion order and renaming a key with Key(old) = new keeps its position
    /// - reading a missing key with Item adds it (with the value Empty)
    /// </summary>
    [SourceClassName("Dictionary")] // TypeName(CreateObject("Scripting.Dictionary")) is "Dictionary"
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
#pragma warning disable CA1010 // Generic IEnumerable<T> - the shape mirrors the COM Scripting.Dictionary (For Each over the keys, as variants)
    internal sealed class ScriptonScriptingDictionary : ScriptonIReflectOnClrType, IEnumerable // IReflect as the other built-ins (issue #159); the DISPIDs are scrrun's
#pragma warning restore CA1010
    {
        public const string ProgId = "Scripting.Dictionary";

        private const string ErrorSource = "Microsoft VBScript runtime error";

        private readonly LinkedList<KeyValuePair<object?, object?>> _entries = new LinkedList<KeyValuePair<object?, object?>>();
        private Dictionary<VariantKey, LinkedListNode<KeyValuePair<object?, object?>>> _index;
        private int _compareMode;

        public ScriptonScriptingDictionary() : base(typeof(ScriptonScriptingDictionary))
        {
            _compareMode = 0;
            _index = new Dictionary<VariantKey, LinkedListNode<KeyValuePair<object?, object?>>>(new VariantKeyComparer(textCompare: false));
        }

        /// <summary>
        /// 0 (vbBinaryCompare) or 1 (vbTextCompare). This may only be changed while the dictionary is empty (otherwise error 5, Invalid procedure
        /// call or argument, is raised, as scrrun does).
        /// </summary>
        [DispId(9)]
        public int CompareMode
        {
            get => _compareMode;
            set
            {
                if ((value != 0) && (value != 1))
                    throw new InvalidProcedureCallOrArgumentException($"CompareMode must be 0 (binary) or 1 (text), not {value}");
                if (value == _compareMode)
                    return;
                if (_entries.Count != 0)
                    throw new InvalidProcedureCallOrArgumentException("CompareMode can not be changed when the dictionary contains items");
                _compareMode = value;
                _index = new Dictionary<VariantKey, LinkedListNode<KeyValuePair<object?, object?>>>(new VariantKeyComparer(textCompare: value == 1));
            }
        }

        [DispId(2)]
        public int Count => _entries.Count;

        [DispId(1)]
        public void Add(object? key, object? item)
        {
            VariantKey variantKey = new VariantKey(key);
            if (_index.ContainsKey(variantKey))
                throw new CustomException(457, ErrorSource, "This key is already associated with an element of this collection");
            _index.Add(variantKey, _entries.AddLast(new KeyValuePair<object?, object?>(key, item)));
        }

        [DispId(3)]
        public bool Exists(object? key) => _index.ContainsKey(new VariantKey(key));

        /// <summary>
        /// The default member, an indexed property (Item, as with scrrun): reading a key that does not exist adds it with the value Empty (as
        /// scrrun does); d(key) = value / d.Item(key) = value adds the key or replaces its value
        /// </summary>
        [DispId(0)]
        [IsDefault]
        public object? this[object? key]
        {
            get
            {
                VariantKey variantKey = new VariantKey(key);
                if (_index.TryGetValue(variantKey, out LinkedListNode<KeyValuePair<object?, object?>>? node))
                    return node.Value.Value;
                _index.Add(variantKey, _entries.AddLast(new KeyValuePair<object?, object?>(key, null)));
                return null;
            }
            set
            {
                VariantKey variantKey = new VariantKey(key);
                if (_index.TryGetValue(variantKey, out LinkedListNode<KeyValuePair<object?, object?>>? node))
                    node.Value = new KeyValuePair<object?, object?>(node.Value.Key, value);
                else
                    _index.Add(variantKey, _entries.AddLast(new KeyValuePair<object?, object?>(key, value)));
            }
        }

        /// <summary>
        /// d.Key(oldKey) = newKey - renames a key, keeping its position and value
        /// </summary>
        [DispId(5)]
        public void Key(object? oldKey, object? newKey)
        {
            VariantKey oldVariantKey = new VariantKey(oldKey);
            if (!_index.TryGetValue(oldVariantKey, out LinkedListNode<KeyValuePair<object?, object?>>? node))
                throw new CustomException(32811, ErrorSource, "Element not found");
            VariantKey newVariantKey = new VariantKey(newKey);
            if (_index.ContainsKey(newVariantKey) && !_index.Comparer.Equals(oldVariantKey, newVariantKey))
                throw new CustomException(457, ErrorSource, "This key is already associated with an element of this collection");

            _index.Remove(oldVariantKey);
            node.Value = new KeyValuePair<object?, object?>(newKey, node.Value.Value);
            _index[newVariantKey] = node;
        }

        [DispId(7)]
        public void Remove(object? key)
        {
            VariantKey variantKey = new VariantKey(key);
            if (!_index.TryGetValue(variantKey, out LinkedListNode<KeyValuePair<object?, object?>>? node))
                throw new CustomException(32811, ErrorSource, "Element not found");
            _index.Remove(variantKey);
            _entries.Remove(node);
        }

        [DispId(8)]
        public void RemoveAll()
        {
            _index.Clear();
            _entries.Clear();
        }

        /// <summary>
        /// The keys, in insertion order
        /// </summary>
        [DispId(6)]
        public object?[] Keys()
        {
            var keys = new object?[_entries.Count];
            int index = 0;
            foreach (KeyValuePair<object?, object?> entry in _entries)
                keys[index++] = entry.Key;
            return keys;
        }

        /// <summary>
        /// The items, in insertion order
        /// </summary>
        [DispId(4)]
        public object?[] Items()
        {
            var items = new object?[_entries.Count];
            int index = 0;
            foreach (KeyValuePair<object?, object?> entry in _entries)
                items[index++] = entry.Value;
            return items;
        }

        /// <summary>
        /// For Each enumerates the keys (a snapshot, in insertion order)
        /// </summary>
        public IEnumerator GetEnumerator() => Keys().GetEnumerator();

        [DispId(-4)] // DISPID_NEWENUM
#pragma warning disable CA1707 // The name is the COM member name
        public IEnumerator _NewEnum() => GetEnumerator();
#pragma warning restore CA1707

        private readonly struct VariantKey
        {
            public VariantKey(object? key)
            {
                // Nothing is a ScriptDispatchWrapper around null - as a key it is the same as any other Nothing
                Value = (key is ScriptDispatchWrapper wrapper) ? (wrapper.WrappedObject ?? NothingKey) : key;
            }

            public object? Value { get; }
        }

        private static readonly object NothingKey = new object();

        private sealed class VariantKeyComparer : IEqualityComparer<VariantKey>
        {
            private readonly bool _textCompare;

            public VariantKeyComparer(bool textCompare)
            {
                _textCompare = textCompare;
            }

            public bool Equals(VariantKey x, VariantKey y)
            {
                object? a = x.Value;
                object? b = y.Value;
                if ((a == null) || (b == null))
                    return (a == null) && (b == null); // Empty is only equal to Empty
                if ((a is string sa) && (b is string sb))
                    return string.Equals(sa, sb, _textCompare ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal);
                if (TryGetNumber(a, out double na) && TryGetNumber(b, out double nb))
                    return na.Equals(nb);
                if ((a is DateTime da) && (b is DateTime db))
                    return da == db;
                if (IsValueKey(a) || IsValueKey(b))
                    return false; // eg. a number and a string are different keys
                return ReferenceEquals(a, b); // object keys are compared by reference
            }

            public int GetHashCode(VariantKey obj)
            {
                object? value = obj.Value;
                if (value == null)
                    return 0;
                if (value is string s)
                    return _textCompare ? StringComparer.OrdinalIgnoreCase.GetHashCode(s) : StringComparer.Ordinal.GetHashCode(s);
                if (TryGetNumber(value, out double number))
                    return number.GetHashCode();
                if (value is DateTime date)
                    return date.GetHashCode();
                return System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(value);
            }

            private static bool IsValueKey(object value) => (value is string) || (value is DateTime) || TryGetNumber(value, out _);

            private static bool TryGetNumber(object value, out double number)
            {
                switch (value)
                {
                    case byte _:
                    case short _:
                    case int _:
                    case long _:
                    case float _:
                    case double _:
                    case decimal _:
                        number = Convert.ToDouble(value, CultureInfo.InvariantCulture);
                        return true;
                    case bool boolean:
                        number = boolean ? -1 : 0;
                        return true;
                    default:
                        number = 0;
                        return false;
                }
            }
        }
    }
}
