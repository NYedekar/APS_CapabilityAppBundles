using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

namespace InventorBOMExtractor
{
    // Manually declare IApplicationAddInServer so InventorCoreConsole can find the plugin
    // via QI without needing Autodesk.Inventor.Interop.dll at compile time.
    // GUID 6BA435DD-... is the only GUID embedded in Autodesk.Forge.DesignAutomation.Inventor.Utils.dll
    // and matches the IApplicationAddInServer dispinterface IID from the Inventor type library.
    [Guid("6BA435DD-BBD6-11D4-8DE6-0010B541CAA8")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    [ComVisible(true)]
    public interface IApplicationAddInServer
    {
        void Activate(object addInSiteObject, bool firstTime);
        void Deactivate();
        void ExecuteCommand(int commandID);
        object Automation { get; }
    }

    [Guid("00b94922-a4b5-4867-98bc-4e9418b04cfe")]
    [ClassInterface(ClassInterfaceType.None)]  // expose ONLY IApplicationAddInServer
    [ComVisible(true)]
    public class PluginServer : IApplicationAddInServer
    {
        public object Automation { get; private set; }

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
