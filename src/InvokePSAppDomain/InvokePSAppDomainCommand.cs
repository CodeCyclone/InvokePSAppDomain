using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Management.Automation;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace InvokePSAppDomain
{
    [Cmdlet(CmdletVerb, CmdletNoun)]
    public class InvokePSAppDomainCommand : Cmdlet
    {
        private const string CmdletVerb = "Invoke";
        private const string CmdletNoun = "PSAppDomain";
        public const string CmdletName = CmdletVerb + "-" + CmdletNoun;

        private PSMessageQueues messageQueues;

        [Alias("Command", "Script")]
        [Parameter(Position = 0, HelpMessage = "Script text to execute in seperate AppDomain")]
        public string[] ScriptText { get; set; }

        public Hashtable Variables { get; set; }

        protected override void BeginProcessing()
        {
            if(ScriptText == null)
            {
                ThrowTerminatingError(new ErrorRecord(new PSArgumentNullException(nameof(ScriptText)), "Null Parameter", ErrorCategory.InvalidArgument, this));
            }

            messageQueues = new PSMessageQueues();
            if(Variables != null)
            {
                RemoveNonSerializableVariables();
            }
        }

        protected override void EndProcessing()
        {
            WriteDebug("Starting " + CmdletName);
            try
            {
                Task task = Task.Factory.StartNew(delegate
                {
                    using(PowerShellAppDomain powerShellAppDomain = new PowerShellAppDomain(messageQueues))
                    {
                        powerShellAppDomain.Variables = Variables;
                        powerShellAppDomain.Execute(ScriptText);
                    }
                });

                while(!task.Wait(100))
                {
                    messageQueues.DisplayOutputs(this);
                }

                messageQueues.DisplayOutputs(this);
            }
            catch (AggregateException ex)
            {
                foreach(var innerException in ex.InnerExceptions)
                {
                    WriteError(new ErrorRecord(innerException, "Aggregate Exception", ErrorCategory.InvalidOperation, this));
                }

                throw;
            }
            finally
            {
                WriteDebug("Stopping " + CmdletName);
            }
        }

        private void RemoveNonSerializableVariables()
        {
            Hashtable hashtable = new Hashtable();
            foreach(object key in Variables.Keys)
            {
                object obj = Variables[key];
                if(obj == null)
                {
                    hashtable.Add(key, obj);
                }
                else
                {
                    Type type = obj.GetType();
                    if(type.IsSerializable || type.GetInterface(typeof(ISerializable).FullName) != null)
                    {
                        hashtable.Add(key, obj);
                    }
                    else
                    {
                        WriteWarning(string.Format(CultureInfo.InvariantCulture, "Removed variable '{0}' from Variables due to not being serializable, type = {1}", new object[2]
                        {
                            key,
                            type.FullName
                        }));
                    }
                }
            }

            Variables = hashtable;
        }
    }
}
