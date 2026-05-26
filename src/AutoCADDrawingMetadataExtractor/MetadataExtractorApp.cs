using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using Newtonsoft.Json;
using System.IO;
using System.Text;

[assembly: CommandClass(typeof(AutoCADDrawingMetadataExtractor.MetadataExtractorCommands))]
[assembly: ExtensionApplication(typeof(AutoCADDrawingMetadataExtractor.MetadataExtractorApp))]

namespace AutoCADDrawingMetadataExtractor
{
    public class MetadataExtractorApp : IExtensionApplication
    {
        public void Initialize() { }
        public void Terminate() { }
    }

    public class MetadataExtractorCommands
    {
        [CommandMethod("EXTRACTDWGMETADATA", CommandFlags.Modal)]
        public static void ExtractDwgMetadata()
        {
            Database? db;
            try
            {
                var doc = Application.DocumentManager.MdiActiveDocument;
                db = doc?.Database ?? HostApplicationServices.WorkingDatabase;
            }
            catch
            {
                db = HostApplicationServices.WorkingDatabase;
            }

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

                var jsonSettings = new JsonSerializerSettings
                {
                    Formatting = Formatting.Indented,
                    NullValueHandling = NullValueHandling.Ignore,
                };

                string json = JsonConvert.SerializeObject(report, jsonSettings);
                File.WriteAllText("result.json", json, Encoding.UTF8);

                System.Console.WriteLine($"[MetadataExtractor] Done — result.json written ({json.Length} bytes).");
            }
            catch (System.Exception ex)
            {
                // Qualify System.Exception explicitly — Autodesk.AutoCAD.Runtime also defines Exception.
                System.Console.WriteLine($"[MetadataExtractor] ERROR: {ex.Message}\n{ex.StackTrace}");
            }
        }
    }
}
