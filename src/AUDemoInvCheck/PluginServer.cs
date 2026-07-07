// ════════════════════════════════════════════════════════════════════════════
//  APS DA — Inventor plugin entry point (template)
//
//  VERIFIED, shipped green: copied from InventorBOMExtractor's PluginServer.cs
//  (src/InventorBOMExtractor/PluginServer.cs in APS_CapabilityAppBundles).
//
//  PROVEN PATTERN: PluginServer implements Inventor.ApplicationAddInServer
//  (embedded interop type from ../../lib/Inventor/Inventor.Interop.il — see that
//  file for the REAL DISPIDs InventorCoreConsole uses: Activate=0x03001201,
//  Deactivate=0x03001202, ExecuteCommand=0x03001203, get_Automation=0x03001204).
//  Those DISPIDs live on the interop interface, NOT here — do not add DispId
//  attributes to this class; the embedded interop type supplies them.
//
//  ClassInterface(ClassInterfaceType.None): PluginServer only exposes
//  IApplicationAddInServer via QueryInterface, matching the official
//  UpdateIPTParam.PluginServer pattern. Automation returns a separate
//  AUDemoInvCheckAutomation instance (not this), so InventorCoreConsole can call
//  Run on a plain ComVisible object with default AutoDispatch.
//
//  REPLACE: AUDemoInvCheck, {6B3CF117-AB66-4B2E-9A50-5BF5D621CC6B}, AUDemoInvCheckAutomation.
//  {6B3CF117-AB66-4B2E-9A50-5BF5D621CC6B} is a FRESH GUID (run: uuidgen) — it becomes ClassId AND
//  ClientId in App.Inventor.addin AND ProductCode in PackageContents.xml (all
//  three collapse to the SAME GUID in this pattern; do not use three different
//  GUIDs).
// ════════════════════════════════════════════════════════════════════════════
using System;
using System.Runtime.InteropServices;
using Inventor;  // Inventor.ApplicationAddInServer embedded from Inventor.Interop.dll

namespace AUDemoInvCheck
{
    [Guid("{6B3CF117-AB66-4B2E-9A50-5BF5D621CC6B}")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class PluginServer : ApplicationAddInServer
    {
        // Typed as dynamic to match the UpdateIPTParam.PluginServer.Automation pattern.
        public dynamic Automation { get; private set; } = null!;

        public void Activate(object addInSiteObject, bool firstTime)
        {
            Automation = new AUDemoInvCheckAutomation();
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
