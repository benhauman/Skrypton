using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Helpline.Application.ScriptingModel
{
    class ApplicationTestContext : IApplicationTestContext
    {
        private ApplicationTestContext()
        {

        }
        public static IApplicationTestContext Create(Action<ApplicationTestContext> setup)
        {
            ApplicationTestContext ctx = new ApplicationTestContext();
            setup(ctx);
            return ctx;
        }
        public Action<IConnectivityDomainObject> HandlerMergeSUs { get; set; }
        public void HandleMergeSUs(IConnectivityDomainObject cncObj)
        {
            if (HandlerMergeSUs == null)
                throw new NotImplementedException("No setup for 'HandlerMergeSUs'.");
            HandlerMergeSUs(cncObj);
        }
    }
}
