using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Text;
using System.Threading.Tasks;

namespace InvokePSAppDomain
{
    public sealed class PowerShellExecutor : MarshalByRefObject, IDisposable
    {
        private bool disposed;

        private Runspace runspace;

        private CmdletMarshal Cmdlet { get; set; }

        public PowerShellExecutor()
        {
            Cmdlet = (CmdletMarshal)AppDomain.CurrentDomain.GetData("cmdMarshal");
        }

        public override object InitializeLifetimeService()
        {
            return null;
        }

        public void Execute(string commandText, Hashtable variables)
        {
            if(runspace == null)
            {
                InitialSessionState initialSessionState = InitialSessionState.CreateDefault();
                runspace = RunspaceFactory.CreateRunspace(initialSessionState);
                runspace.Open();
                if(variables != null)
                {
                    foreach (object key in variables.Keys)
                    {
                        runspace.SessionStateProxy.SetVariable(key.ToString(), variables[key]);
                    }
                }
            }

            using(PowerShell powerShell = PowerShell.Create())
            {
                powerShell.Runspace = runspace;
                powerShell.Streams.Error.DataAdded += Error_DataAdded;
                powerShell.Streams.Warning.DataAdded += Warning_DataAdded;
                powerShell.Streams.Debug.DataAdded += Debug_DataAdded;
                powerShell.Streams.Verbose.DataAdded += Verbose_DataAdded;
                powerShell.Streams.Progress.DataAdded += Progress_DataAdded;
                using(PSDataCollection<PSObject> PSDataCollection = new PSDataCollection<PSObject>())
                {
                    PSDataCollection.DataAdded += Output_DataAdded;
                    powerShell.Commands.Clear();
                    powerShell.Commands.AddScript(commandText).AddStatement();
                    powerShell.Invoke(null, PSDataCollection, null);
                }
            }
        }

        private void Output_DataAdded(object sender, DataAddedEventArgs e)
        {
            Cmdlet.WriteOutput(((PSDataCollection<PSObject>)sender)[e.Index]);
        }

        private void Progress_DataAdded(object sender, DataAddedEventArgs e)
        {
            Cmdlet.WriteProgress(((PSDataCollection<ProgressRecord>)sender)[e.Index]);
        }

        private void Verbose_DataAdded(object sender, DataAddedEventArgs e)
        {
            Cmdlet.WriteVerbose(((PSDataCollection<VerboseRecord>)sender)[e.Index].ToString());
        }

        private void Debug_DataAdded(object sender, DataAddedEventArgs e)
        {
            Cmdlet.WriteDebug(((PSDataCollection<DebugRecord>)sender)[e.Index].ToString());
        }

        private void Warning_DataAdded(object sender, DataAddedEventArgs e)
        {
            Cmdlet.WriteWarning(((PSDataCollection<WarningRecord>)sender)[e.Index].ToString());
        }

        private void Error_DataAdded(object sender, DataAddedEventArgs e)
        {
            Cmdlet.WriteError(((PSDataCollection<ErrorRecord>)sender)[e.Index]);
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
                if(disposing && runspace != null)
                {
                    runspace.Dispose();
                }
            }
        }
    }
}
