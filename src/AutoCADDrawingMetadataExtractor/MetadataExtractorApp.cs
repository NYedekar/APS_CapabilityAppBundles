using Autodesk.AutoCAD.Runtime;
using System;
using System.IO;
using System.Text;

// ── DIAGNOSTIC SMOKE-TEST BUILD (net48 / AutoCAD 2024 engine) ─────────────────
// All in-bundle causes were eliminated on the .NET 8 / +25_0 engine: a correctly
// built, correctly packaged net8.0 assembly was silently never loaded. The only
// remaining variable is the framework/engine itself. This build matches Autodesk's
// proven DA-for-AutoCAD reference stack exactly:
//   * .NET Framework 4.8, AcCoreMgd/AcDbMgd v20 (Copy Local = false)
//   * [assembly: CommandClass(...)] + [assembly: ExtensionApplication(null)]
//     — the official sample passes null, i.e. NO IExtensionApplication.
//   * static [CommandMethod] handlers.
// Still a stub (writes hard-coded result.json, no extractor) so a passing run
// proves the framework/engine was the blocker. If [SMOKE] lines appear, restore
// the full extractor on this exact stack (with Newtonsoft.Json).
// ─────────────────────────────────────────────────────────────────────────────

[assembly: CommandClass(typeof(AutoCADDrawingMetadataExtractor.MetadataExtractorCommands))]
[assembly: ExtensionApplication(null)]

namespace AutoCADDrawingMetadataExtractor
{
    public class MetadataExtractorCommands
    {
        [CommandMethod("EXTRACTDWGMETADATA", CommandFlags.Modal)]
        public static void ExtractDwgMetadata() => WriteSmoke("EXTRACTDWGMETADATA");

        [CommandMethod("EXTRACTALLDRAWINGMETADATA", CommandFlags.Modal)]
        public static void ExtractAllDrawingMetadata() => WriteSmoke("EXTRACTALLDRAWINGMETADATA");

        private static void WriteSmoke(string command)
        {
            Console.WriteLine("[SMOKE] Command " + command + " invoked (net48 build).");
            string json = "{\"smoke\":\"ok\",\"command\":\"" + command + "\",\"framework\":\"net48\"}";
            File.WriteAllText("result.json", json, Encoding.UTF8);
            Console.WriteLine("[SMOKE] result.json written (" + json.Length + " bytes).");
        }
    }
}
