using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace InvokePSAppDomain
{
    public class PSMessageQueues
    {
        private readonly ConcurrentQueue<PSObject> outputObjects = new ConcurrentQueue<PSObject>();
        private readonly ConcurrentQueue<ErrorRecord> errorObjects = new ConcurrentQueue<ErrorRecord>();
        private readonly ConcurrentQueue<ProgressRecord> progressObjects = new ConcurrentQueue<ProgressRecord>();
        private readonly ConcurrentQueue<string> warningObjects = new ConcurrentQueue<string>();
        private readonly ConcurrentQueue<string> verboseObjects = new ConcurrentQueue<string>();
        private readonly ConcurrentQueue<string> debugObjects = new ConcurrentQueue<string>();

        public void AddOutput(PSObject psObject)
        {
            outputObjects.Enqueue(psObject);
        }

        public void AddError(ErrorRecord record)
        {
            errorObjects.Enqueue(record);
        }

        public void AddProgress(ProgressRecord record)
        {
            progressObjects.Enqueue(record);
        }

        public void AddWarning(string message)
        {
            warningObjects.Enqueue(message);
        }

        public void AddVerbose(string message)
        {
            verboseObjects.Enqueue(message);
        }

        public void AddDebug(string message)
        {
            debugObjects.Enqueue(message);
        }

        public void DisplayOutputs(Cmdlet cmdlet)
        {
            if(cmdlet == null)
            {
                throw new ArgumentNullException(nameof(cmdlet));
            }

            DisplayOutput(debugObjects, cmdlet.WriteDebug);
            DisplayOutput(progressObjects, cmdlet.WriteProgress);
            DisplayOutput(errorObjects, cmdlet.WriteError);
            DisplayOutput(warningObjects, cmdlet.WriteWarning);
            DisplayOutput(verboseObjects, cmdlet.WriteVerbose);
            DisplayOutput(outputObjects, cmdlet.WriteObject);
        }

        private static void DisplayOutput<T>(ConcurrentQueue<T> queue, Action<T> output) where T : class
        {
            if(queue == null)
            {
                throw new ArgumentNullException(nameof(queue));
            }

            if(output == null)
            {
                throw new ArgumentNullException(nameof(output));
            }

            lock(queue)
            {
                T result = null;
                while(queue.TryDequeue(out result))
                {
                    output(result);
                }
            }
        }
    }
}
