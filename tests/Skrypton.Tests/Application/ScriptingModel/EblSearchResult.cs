using System.Collections;
using System.Collections.Generic;

namespace Helpline.Application.ScriptingModel
{
    class EblSearchResult : System.Collections.IEnumerable // see IHLEblSearchResultDisp
    {

        public int Count {

            get
            {
                return items.Count;
            }
        }

        internal readonly List<object> items = new List<object>();

        internal EblSearchResult AddLoadedItem(object item)
        {
            items.Add(item);
            return this;
        }

        public System.Collections.IEnumerator GetEnumerator()
        {
            return this.items.GetEnumerator();
        }
    }
}
