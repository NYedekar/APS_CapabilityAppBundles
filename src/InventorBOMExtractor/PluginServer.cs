using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Inventor;  // Inventor.ApplicationAddInServer embedded from Inventor.Interop.dll

namespace InventorBOMExtractor
{
    [Guid("00b94922-a4b5-4867-98bc-4e9418b04cfe")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComVisible(true)]
    public class PluginServer : ApplicationAddInServer
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

        public void ExecuteCommand(int commandID) { }
    }
}
