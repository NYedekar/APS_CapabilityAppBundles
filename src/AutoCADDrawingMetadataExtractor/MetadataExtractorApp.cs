using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

[assembly: CommandClass(typeof(AutoCADDrawingMetadataExtractor.MetadataExtractorCommands))]
[assembly: ExtensionApplication(typeof(AutoCADDrawingMetadataExtractor.MetadataExtractorApp))]

namespace AutoCADDrawingMetadataExtractor
{
    public class MetadataExtractorApp : IExtensionApplication
    {
        public void Initialize() { System.Console.WriteLine("[MetadataExtractor] Initialize called."); }
        public void Terminate() { }
    }

    public class MetadataExtractorCommands
    {
        private static readonly JsonSerializerOptions JsonOptions = new()
        {
            WriteIndented = true,
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
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
                string json = JsonSerializer.Serialize(report, JsonOptions);
                File.WriteAllText("result.json", json, Encoding.UTF8);
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
                string json = JsonSerializer.Serialize(result, JsonOptions);
                File.WriteAllText("result.json", json, Encoding.UTF8);
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
