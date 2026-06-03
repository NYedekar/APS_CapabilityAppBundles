using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

namespace InventorBOMExtractor
{
    // ClassInterface.AutoDual exposes Activate/Deactivate/ExecuteCommand/Automation via IDispatch.
    // IApplicationAddInSite is a vtable-only COM interface — accessing it via dynamic throws
    // E_NOINTERFACE. We don't need the server object: Run(doc) works entirely from the doc arg.
    [Guid("00b94922-a4b5-4867-98bc-4e9418b04cfe")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ComVisible(true)]
    public class PluginServer
    {
        public object? Automation { get; private set; }

        public void Activate(object addInSiteObject, bool firstTime)
        {
            Trace.TraceInformation("[InventorBOMExtractor] Activate v"
                + Assembly.GetExecutingAssembly().GetName().Version?.ToString(4));
            try
            {
                Automation = new BOMExtractorAutomation();
                Trace.TraceInformation("[InventorBOMExtractor] Automation ready.");
            }
            catch (Exception ex)
            {
                Trace.TraceError("[InventorBOMExtractor] Activate failed: " + ex);
                try { File.WriteAllText("activate_error.txt", ex.ToString()); } catch { }
            }
        }

        public void Deactivate()
        {
            Trace.TraceInformation("[InventorBOMExtractor] Deactivate");
            Automation = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public void ExecuteCommand(int CommandID) { }
    }
}
