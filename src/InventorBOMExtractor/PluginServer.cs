using System;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;

namespace InventorBOMExtractor
{
    // ClassInterface.AutoDual exposes all public methods via IDispatch so Inventor
    // can call Activate/Deactivate/ExecuteCommand/Automation without the interop DLL.
    [Guid("00b94922-a4b5-4867-98bc-4e9418b04cfe")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ComVisible(true)]
    public class PluginServer
    {
        private dynamic? _inventorServer;

        public dynamic? Automation { get; private set; }

        public void Activate(dynamic addInSiteObject, bool firstTime)
        {
            Trace.TraceInformation("[InventorBOMExtractor] Activate v"
                + Assembly.GetExecutingAssembly().GetName().Version?.ToString(4));
            _inventorServer = addInSiteObject.InventorServer;
            Automation = new BOMExtractorAutomation(_inventorServer);
        }

        public void Deactivate()
        {
            Trace.TraceInformation("[InventorBOMExtractor] Deactivate");
            if (_inventorServer != null)
            {
                Marshal.ReleaseComObject(_inventorServer);
                _inventorServer = null;
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public void ExecuteCommand(int CommandID) { }
    }
}
