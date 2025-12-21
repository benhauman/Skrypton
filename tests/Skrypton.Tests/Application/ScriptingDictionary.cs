using Helpline.Application.ScriptingModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Skrypton.Tests.Application
{
    [ComVisible(true)]
    class ScriptingDictionary
    {
        private bool caseSensitive;
        private Dictionary<string, object> dict;

        sealed class MyEqualityComparer : IEqualityComparer<string>
        {
            internal readonly bool caseSensitive;
            public MyEqualityComparer(bool caseSensitive)
            {
                this.caseSensitive = caseSensitive;
            }
            public bool Equals(string x, string y)
            {
                return string.Equals(x, y, StringComparison.Ordinal);
            }

            public int GetHashCode(string obj)
            {
                return obj == null ? 0 : obj.GetHashCode();
            }
        }
        public ScriptingDictionary()
        {
            this.caseSensitive = true;
        }
        //  the Dictionary object is case sensitive by default.
        // https://www.itprotoday.com/devops-and-software-development/scripting-dictionary-makes-it-easy
        public object CompareMode // http://www.devguru.com/content/technologies/vbscript/dictionary-comparemode.html
        {
            // VBBinaryCompare   - 0 - Binary Comparison (case-sensitive)
            // VBTextCompare     - 1 - Text Comparison
            // VBDataBaseCompare - 2 - Compare information inside database
            // Values greater than 2 can be used to refer to comparisons using specific Locale IDs
            get
            {
                return caseSensitive ? 0 : 1;
            }
            set
            {
                // You will get an error if you try to set CompareMode on a Dictionary that contains items
                int newValue = (int)value;
                if (newValue == 0 || newValue == 1)
                {
                    if (dict != null && dict.Count > 0)
                        throw new NotSupportedException("Cannot set CompareMode on a Dictionary that contains.");

                    bool caseSensitiveNew = newValue == 0;
                    if (caseSensitiveNew != caseSensitive)
                    {
                        caseSensitive = caseSensitiveNew;
                        dict = null;
                    }
                }
                else
                {
                    throw new NotSupportedException("" + value + " " + value.GetType().Name);
                }
            }
        }

        [System.Runtime.InteropServices.DispIdAttribute(DISPIDs.DISPID_VALUE)] // + ComVisibleAttribute!!!
        public object this[string name] // DispId(0:DISPID_VALUE) + ComVisibleAttribute!!!
        {
            get
            {
                return GetItemByName(name);
            }
            set
            {
                SetItemByName(name, value);
            }
        }

        private IDictionary<string, object> EnsureItems()
        {
            if (dict == null)
            {
                dict = new Dictionary<string, object>(new MyEqualityComparer(caseSensitive));
            }
            return dict;
        }

        private object GetItemByName(string name)
        {
            object item;
            if (EnsureItems().TryGetValue(name, out item))
                return item;
            return null;
        }

        private void SetItemByName(string name, object value)
        {
            if (EnsureItems().ContainsKey(name))
            {

            }

            EnsureItems().Add(name, value);
        }
    }
}
