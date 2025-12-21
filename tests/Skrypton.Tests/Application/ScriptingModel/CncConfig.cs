using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;

namespace Helpline.Application.ScriptingModel
{
    [DebuggerDisplay("{name}")]
    [ComVisible(true)]

    class CncConfigGroup // IHLCncConfigGroup
    {
        private readonly string name;
        public CncConfigGroup(string name)
        {
            this.name = name;
        }

        public CncConfigValue GetValue(string key)
        {
            if (this.values.TryGetValue(((string)key).ToUpperInvariant(), out CncConfigValue value))
                return value;
            throw new InvalidOperationException("Configuration value not found:" + key + ", group:" + this.name);
        }
        internal CncConfigGroup InitValue(string key, Action<CncConfigValue> setup)
        {
            CncConfigValue value = new CncConfigValue();
            this.values.Add(key.ToUpperInvariant(), value);
            setup(value);
            return this;
        }
        public CncConfigGroup GetGroup(string key)
        {
            if (this.subGroups.TryGetValue(key.ToUpperInvariant(), out CncConfigGroup grp))
                return grp;
            throw new InvalidOperationException("Configuration group not found:" + key);
        }

        internal CncConfigGroup AddGroup(string key, Action<CncConfigGroup> setup)
        {
            CncConfigGroup grp = new CncConfigGroup(key);
            setup(grp);

            subGroups.Add(key.ToUpperInvariant(), grp);
            return this;
        }

        public CncConfigGroups Groups
        {
            get
            {
                return new CncConfigGroups(this);
            }
        }

        internal Dictionary<string, CncConfigGroup> subGroups = new Dictionary<string, CncConfigGroup>();
        internal Dictionary<string, CncConfigValue> values = new Dictionary<string, CncConfigValue>();
    }

    [System.Runtime.InteropServices.ComVisible(true)]
    class CncConfigGroups : System.Collections.IEnumerable // IHLCncConfigGroups
    {
        private readonly CncConfigGroup m_spGroup;
        public CncConfigGroups(CncConfigGroup pGroup)
        {
            m_spGroup = pGroup;
        }
        ///[System.Runtime.InteropServices.DispIdAttribute(DISPIDs.DISPID_NEWENUM)]
        ///public object NewEnumerator()
        ///{
        ///    throw new NotImplementedException();
        ///}
        public System.Collections.IEnumerator GetEnumerator()
        {
            return CncConfigEnum.InitFromDictionary(m_spGroup.subGroups.Values.Select(x => (object)x).ToArray());
        }




        // get_Item
        // get_Count
        // get__NewEnum

    }

    class CncConfigEnum
    {

        internal static System.Collections.IEnumerator InitFromDictionary(object[] values)
        {
            // MSSCRX::IDictionary
            return values.GetEnumerator();
        }
    }

    class CncConfigValue
    {
        ///internal string ValueCaseType;
        ///internal string ValueMailAttributeKey;
        ///internal string ValueType;
        internal object m_data;

        public object Data
        {
            get
            {
                return this.m_data;
            }
        }
    }
}
