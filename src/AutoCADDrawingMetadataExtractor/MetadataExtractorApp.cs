using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using Newtonsoft.Json;
using System.IO;
using System.Text;

// Proven DA-for-AutoCAD pattern (.NET Framework 4.8, AcCoreMgd/AcDbMgd v20):
// [assembly: CommandClass(...)] + [assembly: ExtensionApplication(null)] with static
// [CommandMethod] handlers — confirmed to load on the AutoCAD 2024 (+24_3) engine.
// JSON via Newtonsoft.Json (bundled in Contents/), since net48 has no in-box JSON.
[assembly: CommandClass(typeof(AutoCADDrawingMetadataExtractor.MetadataExtractorCommands))]
[assembly: ExtensionApplication(null)]

namespace AutoCADDrawingMetadataExtractor
{
    public class MetadataExtractorCommands
    {
        // Settings created locally per call (not a static field) to avoid any type-load
        // surprises in AutoCAD's isolated load context.
        private static JsonSerializerSettings MakeJsonSettings() => new JsonSerializerSettings
        {
            Formatting = Formatting.Indented,
            NullValueHandling = NullValueHandling.Ignore,
        };

        private static Database? ResolveDatabase()
        {
            try
            {
                var doc = Application.DocumentManager.MdiActiveDocument;
                return doc?.Database ?? HostApplicationServices.WorkingDatabase;
            }
            catch
            {
                return HostApplicationServices.WorkingDatabase;
            }
        }

        [CommandMethod("EXTRACTDWGMETADATA", CommandFlags.Modal)]
        public static void ExtractDwgMetadata()
        {
            var db = ResolveDatabase();
            if (db == null)
            {
                System.Console.WriteLine("[MetadataExtractor] ERROR: No active database.");
                return;
            }

            System.Console.WriteLine("[MetadataExtractor] Starting extraction...");

            try
            {
                var extractor = new DwgMetadataExtractor(db);
                var report = extractor.BuildReport();
                string json = JsonConvert.SerializeObject(report, MakeJsonSettings());
                // UTF-8 WITHOUT BOM — Encoding.UTF8 emits a BOM that breaks strict JSON parsers (JS JSON.parse).
                File.WriteAllText("result.json", json, new UTF8Encoding(false));
                System.Console.WriteLine("[MetadataExtractor] Done -- result.json written (" + json.Length + " bytes).");
            }
            catch (System.Exception ex)
            {
                System.Console.WriteLine("[MetadataExtractor] ERROR: " + ex.Message);
                System.Console.WriteLine(ex.StackTrace);
            }
        }

        // Single-pass combined extraction: all 7 metadata sections in one DWG open.
        // Output keys mirror the 7 individual operationIds for drop-in compatibility.
        [CommandMethod("EXTRACTALLDRAWINGMETADATA", CommandFlags.Modal)]
        public static void ExtractAllDrawingMetadata()
        {
            var db = ResolveDatabase();
            if (db == null)
            {
                System.Console.WriteLine("[MetadataExtractor] ERROR: No active database.");
                return;
            }

            System.Console.WriteLine("[MetadataExtractor] Starting combined extraction...");

            try
            {
                var extractor = new DwgMetadataExtractor(db);
                var result = extractor.BuildCombinedReport();
                string json = JsonConvert.SerializeObject(result, MakeJsonSettings());
                // UTF-8 WITHOUT BOM — Encoding.UTF8 emits a BOM that breaks strict JSON parsers (JS JSON.parse).
                File.WriteAllText("result.json", json, new UTF8Encoding(false));
                System.Console.WriteLine("[MetadataExtractor] Done -- result.json written (" + json.Length + " bytes).");
            }
            catch (System.Exception ex)
            {
                System.Console.WriteLine("[MetadataExtractor] ERROR: " + ex.Message);
                System.Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
