using System;
using System.Collections.Generic;

namespace Helpline.Application.ScriptingModel
{
    internal static class DISPIDs
    {
        public const int DISPID_VALUE = 0; // see c:\Program Files\Microsoft SDKs\Windows\v6.0A\Include\OAIdl.h
        public const int DISPID_NEWENUM = -4;
    }
    //[System.Reflection.DefaultMemberAttribute("Get")]
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    class CncJob   // IDispatch, IHLCncJob
    {
        private readonly IApplicationTestContext cncTestContext;
        public CncJob(IApplicationTestContext cncTestContext)
        {
            this.cncTestContext = cncTestContext ?? throw new ArgumentNullException(nameof(cncTestContext));
        }

        // [DispIdAttribute(0)]
        private void MyMth0()
        { }


        [System.Runtime.InteropServices.DispIdAttribute(DISPIDs.DISPID_VALUE)] // + ComVisibleAttribute!!!
        public object this[string name] // DispId(0:DISPID_VALUE) + ComVisibleAttribute!!!
        {
            get
            {
                var pi = GetDynamicPropertyByName(name);
                return pi.GetValue(this);
            }
            set
            {
                ///object val = (object)value;
                var pi = GetDynamicPropertyByName(name);
                pi.SetValue(this, value);
            }
        }

        private System.Reflection.PropertyInfo GetDynamicPropertyByName(string name)
        {
            var pi = GetType().GetProperty(name, System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.IgnoreCase);
            if (pi == null)
                throw new InvalidOperationException("HL_E_INVALID_NAME. name[" + name + "]");
            return pi;
        }
        private static object RefNotNull(object value)
        {
            if (value == null)
            {
                throw new NotSupportedException("not initialized");
            }
            return value;
        }

        private object ProcessText { get { return null; } set { SetProcessTextImpl(value); } }
        private object MailRequest { get { return RefNotNull(m_mailRequest); } set { throw new NotSupportedException(); } }
        private object ServerConnection { get { return RefNotNull(this); } set { throw new NotSupportedException(); } }
        private object Config { get { return RefNotNull(m_cfg); } set { throw new NotSupportedException(); } }
        private object Worker { get { return RefNotNull(this); } set { throw new NotSupportedException(); } }


        internal CncMail m_mailRequest;
        internal CncConfigGroup m_cfg;

        private void SetProcessTextImpl(object value)
        {
            Console.WriteLine("CNC-Log:" + (string)value);
        }


        public string NewWorkflowFromMail(string workflowName)
        {
            Console.WriteLine("NewWorkflowFromMail:" + workflowName);
            return null;
        }

        public bool CanExtendWorkflowCase(object objIdentity)
        {
            Console.WriteLine("CanExtendWorkflowCase:" + objIdentity);
            return true;
        }
        public string DoExtendWorkflowCase(object objIdentity)
        {
            Console.WriteLine("DoExtendWorkflowCase:" + objIdentity);
            return null;
        }

        public CncObj GetCaseByReferenceNumber(string referenceNumber)
        {
            var oi = new HLOBJECTID(191, 23822);
            return new CncObj(cncTestContext, oi) { controllerId = CncObj.wfInstanceId, referenceNumber = referenceNumber };
        }

        public bool IsBuiltinCase(object objIdentity)
        {
            CncObj obj = objIdentity as CncObj;
            return string.Equals(obj.controllerId, CncObj.szPNODEWFINSTANCEID, StringComparison.OrdinalIgnoreCase);
        }

        public object find_Cases(string condition, int reserved)
        {
            var oi = new HLOBJECTID(22292, 55555);
            return new EblSearchResult().AddLoadedItem(
                new CncObj(cncTestContext, oi) { controllerId = CncObj.wfInstanceId, referenceNumber = "aaaaaaaaaaaaaaaaa1" }
                );
        }

        public object Find_Persons(string condition, int p2)
        {
            var coll = new MyPersonCollection();
            coll.items.Add(new MyPersonbj { personid = 101, persondefid = 30303 });
            return coll;
        }

        public object CreateScriptEngine()
        {
            return new CncScriptEngine();
        }

        public void ExecuteScript(object engine, object job, string scriptName)
        {
            CncScriptEngine cncEngine = (CncScriptEngine)engine;
            CncJob me = (CncJob)job;
            Console.WriteLine("ExecuteScript:" + scriptName);
        }

    }

    [System.Runtime.InteropServices.ComVisible(true)]
    class MyPersonCollection
    {
        internal readonly List<MyPersonbj> items = new List<MyPersonbj>();

        public int Count
        {
            get
            {
                return items.Count;
            }
        }
    }

    [System.Runtime.InteropServices.ComVisible(true)]
    class MyPersonbj
    {
        internal int personid;
        internal int persondefid;
    }
}
