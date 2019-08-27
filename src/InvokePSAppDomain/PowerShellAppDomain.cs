using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Reflection;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;

namespace InvokePSAppDomain
{
    public sealed class PowerShellAppDomain : IDisposable
    {
        private bool disposed;
        private AppDomain scriptDomain;
        private PSMessageQueues messageQueues;
        private readonly string assemblyFullName = typeof(PowerShellExecutor).Assembly.FullName;
        private readonly string typeFullName = typeof(PowerShellExecutor).FullName;

        public Hashtable Variables { get; set; }

        private static string ExecutingDirectory
        {
            get
            {
                string codeBase = Assembly.GetExecutingAssembly().CodeBase;
                UriBuilder uriBuilder = new UriBuilder(codeBase);
                string path = Uri.UnescapeDataString(uriBuilder.Path);
                return Path.GetDirectoryName(path);
            }
        }

        public PowerShellAppDomain(PSMessageQueues messageQueues)
        {
            if(messageQueues == null)
            {
                throw new ArgumentNullException(nameof(messageQueues));
            }

            this.messageQueues = messageQueues;
        }

        public void Execute(params string[] scripts)
        {
            if(scripts == null)
            {
                throw new ArgumentNullException(nameof(scripts));
            }

            if(scriptDomain == null)
            {
                Evidence securityInfo = new Evidence(AppDomain.CurrentDomain.Evidence);
                AppDomainSetup info = new AppDomainSetup
                {
                    ApplicationBase = ExecutingDirectory
                };

                scriptDomain = AppDomain.CreateDomain(string.Format(CultureInfo.InvariantCulture, "{0}:{1}", new object[2]
                {
                    GetType(),
                    Guid.NewGuid()
                }), securityInfo, info);
                messageQueues.AddDebug("AppDomain Created: " + scriptDomain.FriendlyName);
                CmdletMarshal data = new CmdletMarshal
                {
                    MessageQueues = messageQueues
                };
                scriptDomain.SetData("cmdMarshal", data);
            }

            messageQueues.AddDebug("Creating proxy instance of " + assemblyFullName + " :: type: " + typeFullName);
            using(PowerShellExecutor powershellExecutor = (PowerShellExecutor)scriptDomain.CreateInstanceAndUnwrap(assemblyFullName, typeFullName))
            {
                messageQueues.AddDebug("Proxy instance created");
                foreach(string text in scripts)
                {
                    messageQueues.AddDebug("Executing: " + text);
                    try
                    {
                        powershellExecutor.Execute(text, Variables);
                    }
                    catch(Exception exception)
                    {
                        messageQueues.AddError(new ErrorRecord(exception, "Proxy Execute Failed", ErrorCategory.OperationStopped, this));
                        throw;
                    }
                }
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private void Dispose(bool disposing)
        {
            if(!disposed)
            {
                disposed = true;
                if(disposing && scriptDomain != null)
                {
                    AppDomain.Unload(scriptDomain);
                    messageQueues.AddDebug("AppDomain Disposed");
                    scriptDomain = null;
                }
            }
        }
    }
}
