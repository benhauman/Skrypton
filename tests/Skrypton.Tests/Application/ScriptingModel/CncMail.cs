using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Helpline.Application.ScriptingModel
{
    class CncMail
    {
        public CncMail()
        {

        }
        public string Subject { get; internal set; }
        public int MailType { get; set; } // public set!!!

        internal string data_From;
        public string SenderMail => data_From;

    }
}
