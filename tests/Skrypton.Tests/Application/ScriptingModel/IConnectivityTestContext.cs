using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Helpline.Application.ScriptingModel
{
    interface IConnectivityDomainObject
    {

    }
    interface IApplicationTestContext
    {
        void HandleMergeSUs(IConnectivityDomainObject cncObj);
    }
}
