// Scaffolded from templates/revit/DBApplication.cs.
// Capability: export every sheet (number, name, id) as RFC 4180 CSV to result.csv.
// Proven headless pattern: IExternalDBApplication + DesignAutomationBridge.
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
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

                // Tabular output: header row + one row per sheet, ordered by sheet number.
                var sb = new StringBuilder();
                sb.Append("Number,Name,Id\r\n");
                foreach (var s in sheets)
                    sb.Append($"{Csv(s.Number)},{Csv(s.Name)},{s.Id}\r\n");

                // UTF-8 WITHOUT BOM — Encoding.UTF8 emits a BOM that trips strict parsers.
                File.WriteAllText("result.csv", sb.ToString(), new UTF8Encoding(false));

                Console.WriteLine($"[RevitSheetList] Done — {sheets.Count} sheets → result.csv.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[RevitSheetList] ERROR: {ex.Message}");
                // Still emit result.csv so the caller always gets a parseable artifact.
                File.WriteAllText("result.csv",
                    "error\r\n" + Csv(ex.Message) + "\r\n",
                    new UTF8Encoding(false));
            }
        }

        // RFC 4180 field escaping: quote when the value contains a comma, quote, CR or LF;
        // embedded quotes are doubled.
        private static string Csv(string s)
        {
            if (s == null) s = "";
            if (s.IndexOf(',') >= 0 || s.IndexOf('"') >= 0 || s.IndexOf('\n') >= 0 || s.IndexOf('\r') >= 0)
                s = "\"" + s.Replace("\"", "\"\"") + "\"";
            return s;
        }
    }

    internal class SheetRow
    {
        public string Number { get; set; } = "";
        public string Name { get; set; } = "";
        public long Id { get; set; }   // ElementId.Value is Int64 in Revit 2024+
    }
}
