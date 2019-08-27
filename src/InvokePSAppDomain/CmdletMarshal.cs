using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace InvokePSAppDomain
{
    public class CmdletMarshal : MarshalByRefObject
    {
        public PSMessageQueues MessageQueues { get; set; }

        public override object InitializeLifetimeService()
        {
            return null;
        }

        public void WriteOutput(PSObject psObject)
        {
            MessageQueues.AddOutput(psObject);
        }

        public void WriteVerbose(string message)
        {
            MessageQueues.AddVerbose(message);
        }

        public void WriteError(ErrorRecord record)
        {
            MessageQueues.AddError(record);
        }

        public void WriteProgress(ProgressRecord record)
        {
            MessageQueues.AddProgress(record);
        }

        public void WriteWarning(string message)
        {
            MessageQueues.AddWarning(message);
        }

        public void WriteDebug(string message)
        {
            MessageQueues.AddDebug(message);
        }
    }
}
