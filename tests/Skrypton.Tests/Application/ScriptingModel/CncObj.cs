using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Helpline.Application.ScriptingModel
{
    class CncObj : IConnectivityDomainObject // IHLCncHLObj
    {
        internal const string szPNODEWFINSTANCEID = "79F2D5E4-0307-44D3-AF55-51D16604C97B";
        internal const string wfInstanceId = "AAAAAAAA-BBBB-CCCC-DDDD-111166778899";

        internal string controllerId;
        internal string referenceNumber;

        private readonly IApplicationTestContext cncTestContext;
        private readonly HLOBJECTID oi;

        public CncObj(IApplicationTestContext cncTestContext, HLOBJECTID oi)
        {
            this.cncTestContext = cncTestContext ?? throw new ArgumentNullException(nameof(cncTestContext));
            this.oi = oi;
        }

        public void CreateSU()
        {
            Console.WriteLine(">>CreateSU");
        }
        public void AssociatePersons(object persons)
        {
            MyPersonCollection coll = (MyPersonCollection)persons;
        }

        public object GetHLObject()
        {
            return new EblObj(oi);
        }

        public void MergeSUs()
        {
            cncTestContext.HandleMergeSUs(this);
        }
    }

    class EblObj // IHLCncHLObj
    {
        private readonly HLOBJECTID oi;
        internal EblObj(HLOBJECTID oi)
        {
            this.oi = oi;
        }
        public object GetID()
        {
            return oi.ObjectID;
        }

        public object GetValue(object key, object langid, object contentid, object suidx, object datatype)
        {
            // long langid, long contentid, long suidx, long datatype
            //throw new NotImplementedException("" + key + " suidx:" + suidx + " ocid:" + contentid + " datatype:" + datatype + " lang:" + langid);
            return "";
        }
    }
}
