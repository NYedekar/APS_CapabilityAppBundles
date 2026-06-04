using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using Inventor;  // Inventor.ApplicationAddInServer embedded from Inventor.Interop.dll

namespace InventorBOMExtractor
{
    // AutoDispatch: the CCW exposes an auto-generated IDispatch class interface that includes
    // ALL public methods (including Run). This ensures GetIDsOfNames("Run") succeeds when
    // InventorCoreConsole calls it on the object returned by Automation.
    // IApplicationAddInServer (for QI by InventorCoreConsole) is still exposed via the
    // explicitly implemented interface — AutoDispatch does not remove explicit interfaces.
    [Guid("00b94922-a4b5-4867-98bc-4e9418b04cfe")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class PluginServer : ApplicationAddInServer
    {
        // Return this object as the automation target so InventorCoreConsole calls
        // Run() on the same CCW it already holds — avoids any CCW creation issues
        // with a separate BOMExtractorAutomation object.
        public object Automation => this;

        public void Activate(object addInSiteObject, bool firstTime)
        {
            try { File.WriteAllText("activate_start.txt", "Activate called at " + DateTime.UtcNow.ToString("o"), new UTF8Encoding(false)); } catch { }
        }

        public void Deactivate()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public void ExecuteCommand(int commandID) { }

        // SMOKE STUB: proves Run() is callable. Writes result.json so smoke test passes.
        public void Run(object doc)
        {
            try { File.WriteAllText("activate_start.txt", "Run() called at " + DateTime.UtcNow.ToString("o"), new UTF8Encoding(false)); } catch { }
            var json = "{\"source\":\"smoke-stub\",\"generatedAt\":\"" + DateTime.UtcNow.ToString("o") + "\",\"errors\":[\"SMOKE_STUB: Run() reached\"]}";
            File.WriteAllText("result.json", json, new UTF8Encoding(false));
        }

        // Some versions of InventorCoreConsole try RunWithArguments first.
        public void RunWithArguments(object doc, object args) => Run(doc);
    }
}
