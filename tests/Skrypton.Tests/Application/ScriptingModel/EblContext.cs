using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Helpline.Application.ScriptingModel
{
    [ComVisible(true)]
    class EblContext
    {
        private readonly ActionContext context;
        private readonly ActionArgs args;
        public EblContext(ActionContext context, ActionArgs args)
        {
            this.context = context;
            this.args = args;
        }
        public void EnableTrace()
        {

        }

        public int GetLocaleID()
        {
            return this.context.LocaleId;
        }

        public int LangIDFromLCID(int lcid)
        {
            return System.Globalization.CultureInfo.GetCultureInfo(lcid).Parent.LCID;
        }

        public object GetCurrentObject()
        {
            var data = args;
            return LoadObject(data.m_oiDefault.ObjectID, data.m_oiDefault.ObjectDefID);
        }


        internal Func<HLOBJECTID, object> LoadObject_Override { get; set; }
        public object LoadObject(int objid, object objectType)
        {
            int defid = (int)objectType; // or name
            var oi = new HLOBJECTID(objid, defid);
            if (LoadObject_Override != null)
                return LoadObject_Override(oi);
            throw new NotImplementedException("objectid:" + oi.ObjectID + ", defid:" + oi.ObjectDefID);
        }
    }

    class ActionArgs
    {
        public HLOBJECTID m_oiDefault { get; set; }
    }

    class HLOBJECTID
    {
        public HLOBJECTID()
        {

        }
        public HLOBJECTID(int objectId,  int objectDefid)
        {
            ObjectID = objectId;
            ObjectDefID = objectDefid;
        }
        internal int ObjectID { get; set; }
        internal int ObjectDefID { get; set; }
    }

    internal class ActionContext
    {
        internal int LocaleId { get; set; }
    }
}
