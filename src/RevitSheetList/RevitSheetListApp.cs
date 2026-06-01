// Scaffolded from templates/revit/DBApplication.cs.
// Capability: export every sheet (number, name, id) + a count to result.json.
// Proven headless pattern: IExternalDBApplication + DesignAutomationBridge.
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Newtonsoft.Json;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using DesignAutomationFramework;

namespace RevitSheetList
{
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    public class RevitSheetListApp : IExternalDBApplication
    {
        public ExternalDBApplicationResult OnStartup(ControlledApplication app)
        {
            DesignAutomationBridge.DesignAutomationReadyEvent += OnDesignAutomationReady;
            return ExternalDBApplicationResult.Succeeded;
        }

        public ExternalDBApplicationResult OnShutdown(ControlledApplication app)
            => ExternalDBApplicationResult.Succeeded;

        private void OnDesignAutomationReady(object sender, DesignAutomationReadyEventArgs e)
        {
            e.Succeeded = true;
            try
            {
                var doc = e.DesignAutomationData.RevitDoc
                    ?? throw new InvalidOperationException("Revit document could not be opened.");

                Console.WriteLine($"[RevitSheetList] Processing: {doc.Title}");

                var sheets = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewSheet))
                    .Cast<ViewSheet>()
                    .OrderBy(s => s.SheetNumber)
                    .Select(s => new SheetRow
                    {
                        Number = s.SheetNumber,
                        Name   = s.Name,
                        // ElementId.IntegerValue (int) was REMOVED in Revit 2024+ — use .Value (long).
                        // It compiles against the 2024 stubs but throws "Method not found" on the 2026 runtime.
                        Id     = s.Id.Value,
                    })
                    .ToList();

                var report = new SheetReport
                {
                    ExtractedAt = DateTime.UtcNow.ToString("o"),
                    Title       = doc.Title,
                    SheetCount  = sheets.Count,
                    Sheets      = sheets,
                };

                string json = JsonConvert.SerializeObject(report, Formatting.Indented);
                // UTF-8 WITHOUT BOM — Encoding.UTF8 emits a BOM that breaks strict JSON parsers.
                File.WriteAllText("result.json", json, new UTF8Encoding(false));

                Console.WriteLine($"[RevitSheetList] Done — {sheets.Count} sheets → result.json.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[RevitSheetList] ERROR: {ex.Message}");
                var err = new { ok = false, error = ex.Message, stack = ex.StackTrace };
                File.WriteAllText("result.json",
                    JsonConvert.SerializeObject(err, Formatting.Indented),
                    new UTF8Encoding(false));
            }
        }
    }

    internal class SheetReport
    {
        public string ExtractedAt { get; set; } = "";
        public string Title { get; set; } = "";
        public int SheetCount { get; set; }
        public List<SheetRow> Sheets { get; set; } = new List<SheetRow>();
    }

    internal class SheetRow
    {
        public string Number { get; set; } = "";
        public string Name { get; set; } = "";
        public long Id { get; set; }   // ElementId.Value is Int64 in Revit 2024+
    }
}
