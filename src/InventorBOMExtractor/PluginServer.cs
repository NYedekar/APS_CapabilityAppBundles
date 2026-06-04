using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using Inventor;  // Inventor.ApplicationAddInServer embedded from Inventor.Interop.dll

namespace InventorBOMExtractor
{
    // ClassInterface=None: PluginServer only exposes IApplicationAddInServer via QI.
    // No class interface generated — matches UpdateIPTParam.PluginServer pattern.
    // Automation returns a separate BOMExtractorAutomation (not this), so InventorCoreConsole
    // can call Run on a plain ComVisible object with default AutoDispatch.
    [Guid("00b94922-a4b5-4867-98bc-4e9418b04cfe")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class PluginServer : ApplicationAddInServer
    {
        // Typed as dynamic to match UpdateIPTParam.PluginServer.Automation pattern.
        public dynamic Automation { get; private set; } = null!;

        public void Activate(object addInSiteObject, bool firstTime)
        {
            Automation = new BOMExtractorAutomation();
        }

        public void Deactivate()
        {
            Automation = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public void ExecuteCommand(int commandID) { }
    }
}
