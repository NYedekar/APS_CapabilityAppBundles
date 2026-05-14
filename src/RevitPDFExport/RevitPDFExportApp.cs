using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using DesignAutomationFramework;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;

namespace RevitPDFExport
{
    // ─── Input model ────────────────────────────────────────────────────────────

    /// <summary>
    /// Deserialized from params.json provided by the WorkItem caller.
    /// </summary>
    public class ExportParams
    {
        /// <summary>"sheets" or "views"</summary>
        public string Operation { get; set; } = "sheets";

        /// <summary>
        /// For operation="sheets": sheet numbers to export (e.g. ["A1","A2"]).
        /// Empty list = export ALL sheets.
        /// </summary>
        public List<string> SheetNumbers { get; set; } = new List<string>();

        /// <summary>
        /// For operation="views": exact view names to export.
        /// Empty list = export ALL printable model views.
        /// </summary>
        public List<string> ViewNames { get; set; } = new List<string>();

        /// <summary>
        /// true  → all sheets/views combined into a single result.pdf
        /// false → one PDF per sheet/view, all bundled into result.zip
        /// </summary>
        public bool Combine { get; set; } = false;

        /// <summary>
        /// Paper size override: "Default" | "A4" | "A3" | "A2" | "A1" | "A0" |
        ///                      "Letter" | "Tabloid"
        /// "Default" respects each sheet's own paper size (recommended for sheets).
        /// </summary>
        public string PaperSize { get; set; } = "Default";
    }

    // ─── Output model ───────────────────────────────────────────────────────────

    public class ExportResult
    {
        public string Operation { get; set; } = string.Empty;
        public bool Success { get; set; }
        public int ExportedCount { get; set; }
        public List<string> ExportedItems { get; set; } = new List<string>();
        public List<string> SkippedItems { get; set; } = new List<string>();
        public List<string> Errors { get; set; } = new List<string>();
        public string OutputFile { get; set; } = string.Empty;
        public string Timestamp { get; set; } = string.Empty;
    }

    // ─── Plugin entry point ─────────────────────────────────────────────────────

    /// <summary>
    /// APS Design Automation plugin for Revit PDF export.
    ///
    /// CRITICAL RULES (from hard-won experience):
    ///   - Always implement IExternalDBApplication — NOT IExternalApplication
    ///   - Register DesignAutomationReadyEvent in OnStartup — the ONLY entry point DA uses
    ///   - Always set e.Succeeded = true before your logic; set false in catch
    ///   - Write output files to working directory — DA picks them up automatically
    ///   - Use Console.WriteLine for logging — it appears in the DA report
    ///   - Zero UI code — no TaskDialog, no WPF, nothing
    ///   - ElementId.IntegerValue is deprecated in Revit 2024+ — use ElementId.Value
    ///   - net48: no tuple deconstruction on KeyValuePair — use .Key / .Value
    /// </summary>
    [Autodesk.Revit.Attributes.Regeneration(
        Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    public class RevitPDFExportApp : IExternalDBApplication
    {
        // ── Lifecycle ────────────────────────────────────────────────────────────

        public ExternalDBApplicationResult OnStartup(ControlledApplication app)
        {
            DesignAutomationBridge.DesignAutomationReadyEvent += OnReady;
            return ExternalDBApplicationResult.Succeeded;
        }

        public ExternalDBApplicationResult OnShutdown(ControlledApplication app)
            => ExternalDBApplicationResult.Succeeded;

        // ── Main handler ─────────────────────────────────────────────────────────

        private static void OnReady(object sender, DesignAutomationReadyEventArgs e)
        {
            // MUST set true before logic; only set false on unrecoverable error
            e.Succeeded = true;

            var result = new ExportResult
            {
                Timestamp = DateTime.UtcNow.ToString("o")
            };

            try
            {
                var doc = e.DesignAutomationData.RevitDoc
                    ?? throw new InvalidOperationException("RevitDoc is null — document failed to open");

                Console.WriteLine($"[RevitPDFExport] Document opened: {doc.Title}");

                var exportParams = ReadParams();
                result.Operation = exportParams.Operation;

                Console.WriteLine($"[RevitPDFExport] Operation : {exportParams.Operation}");
                Console.WriteLine($"[RevitPDFExport] Combine   : {exportParams.Combine}");
                Console.WriteLine($"[RevitPDFExport] PaperSize : {exportParams.PaperSize}");

                // Create a clean temp folder for the exported PDFs
                var workDir = Directory.GetCurrentDirectory();
                var pdfDir  = Path.Combine(workDir, "pdf_output");
                if (Directory.Exists(pdfDir)) Directory.Delete(pdfDir, true);
                Directory.CreateDirectory(pdfDir);

                // Dispatch to the correct operation
                switch (exportParams.Operation.ToLowerInvariant())
                {
                    case "sheets":
                        ExportSheets(doc, exportParams, pdfDir, result);
                        break;
                    case "views":
                        ExportViews(doc, exportParams, pdfDir, result);
                        break;
                    default:
                        throw new InvalidOperationException(
                            $"Unknown operation '{exportParams.Operation}'. " +
                            "Valid values: \"sheets\", \"views\".");
                }

                // Zip the exported PDFs into result.zip
                var zipPath = Path.Combine(workDir, "result.zip");
                if (File.Exists(zipPath)) File.Delete(zipPath);
                ZipFile.CreateFromDirectory(pdfDir, zipPath);
                result.OutputFile = "result.zip";
                result.Success    = true;

                Console.WriteLine(
                    $"[RevitPDFExport] Done — {result.ExportedCount} item(s) → result.zip");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[RevitPDFExport] ERROR: {ex}");
                result.Success = false;
                result.Errors.Add(ex.Message);
                e.Succeeded = false;
            }
            finally
            {
                // Always write result.json — even on failure, so the caller can inspect
                var jsonSettings = new JsonSerializerSettings
                {
                    ContractResolver  = new CamelCasePropertyNamesContractResolver(),
                    Formatting        = Formatting.Indented
                };
                File.WriteAllText("result.json", JsonConvert.SerializeObject(result, jsonSettings));
                Console.WriteLine("[RevitPDFExport] result.json written");
            }
        }

        // ── Operation: export sheets ─────────────────────────────────────────────

        private static void ExportSheets(
            Document doc,
            ExportParams p,
            string outputDir,
            ExportResult result)
        {
            // Collect all non-placeholder sheets
            var allSheets = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .Where(s => !s.IsPlaceholder)
                .OrderBy(s => s.SheetNumber)
                .ToList();

            Console.WriteLine($"[RevitPDFExport] Total sheets in model: {allSheets.Count}");

            // Apply sheet-number filter if provided
            List<ViewSheet> toExport;
            if (p.SheetNumbers != null && p.SheetNumbers.Count > 0)
            {
                var filter = new HashSet<string>(p.SheetNumbers, StringComparer.OrdinalIgnoreCase);
                toExport = allSheets.Where(s => filter.Contains(s.SheetNumber)).ToList();
                var missing = p.SheetNumbers
                    .Where(n => !allSheets.Any(
                        s => string.Equals(s.SheetNumber, n, StringComparison.OrdinalIgnoreCase)))
                    .ToList();
                foreach (var m in missing)
                    result.SkippedItems.Add($"Sheet not found: {m}");

                Console.WriteLine(
                    $"[RevitPDFExport] Filtered to {toExport.Count} of {allSheets.Count} sheets");
            }
            else
            {
                toExport = allSheets;
                Console.WriteLine($"[RevitPDFExport] Exporting all {toExport.Count} sheets");
            }

            if (toExport.Count == 0)
            {
                result.Errors.Add("No sheets matched the export criteria");
                return;
            }

            var viewIds = toExport.Select(s => s.Id).ToList();

            var options = BuildPDFOptions(p, baseFileName: "sheets");
            bool ok = doc.Export(outputDir, viewIds, options);

            if (!ok)
            {
                // Revit returns false when ALL views fail; partial success still returns true
                throw new InvalidOperationException(
                    "Document.Export returned false for sheets. " +
                    "Ensure sheets contain placed views and the model is not workshared-out.");
            }

            result.ExportedCount = toExport.Count;
            result.ExportedItems = toExport
                .Select(s => $"{s.SheetNumber} – {s.Name}")
                .ToList();
        }

        // ── Operation: export views ──────────────────────────────────────────────

        private static void ExportViews(
            Document doc,
            ExportParams p,
            string outputDir,
            ExportResult result)
        {
            // View types that Revit can export to PDF headlessly
            // (excludes Schedules, Legends, DraftingViews with no geometry,
            //  and internal view types that can't be printed)
            var exportableTypes = new HashSet<ViewType>
            {
                ViewType.FloorPlan,
                ViewType.CeilingPlan,
                ViewType.Elevation,
                ViewType.Section,
                ViewType.ThreeD,
                ViewType.AreaPlan,
                ViewType.EngineeringPlan,
                ViewType.Detail,
                ViewType.DraftingView
            };

            var allViews = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v =>
                    exportableTypes.Contains(v.ViewType) &&
                    !v.IsTemplate &&
                    v.CanBePrinted)
                .OrderBy(v => v.ViewType.ToString())
                .ThenBy(v => v.Name)
                .ToList();

            Console.WriteLine($"[RevitPDFExport] Total printable views in model: {allViews.Count}");

            // Apply view-name filter if provided
            List<View> toExport;
            if (p.ViewNames != null && p.ViewNames.Count > 0)
            {
                var filter = new HashSet<string>(p.ViewNames, StringComparer.OrdinalIgnoreCase);
                toExport = allViews.Where(v => filter.Contains(v.Name)).ToList();
                var missing = p.ViewNames
                    .Where(n => !allViews.Any(
                        v => string.Equals(v.Name, n, StringComparison.OrdinalIgnoreCase)))
                    .ToList();
                foreach (var m in missing)
                    result.SkippedItems.Add($"View not found: {m}");

                Console.WriteLine(
                    $"[RevitPDFExport] Filtered to {toExport.Count} of {allViews.Count} views");
            }
            else
            {
                toExport = allViews;
                Console.WriteLine($"[RevitPDFExport] Exporting all {toExport.Count} views");
            }

            if (toExport.Count == 0)
            {
                result.Errors.Add("No views matched the export criteria");
                return;
            }

            var viewIds = toExport.Select(v => v.Id).ToList();

            var options = BuildPDFOptions(p, baseFileName: "views");
            bool ok = doc.Export(outputDir, viewIds, options);

            if (!ok)
            {
                throw new InvalidOperationException(
                    "Document.Export returned false for views. " +
                    "Ensure views have visible elements and are not empty.");
            }

            result.ExportedCount = toExport.Count;
            result.ExportedItems = toExport
                .Select(v => $"[{v.ViewType}] {v.Name}")
                .ToList();
        }

        // ── PDF options builder ──────────────────────────────────────────────────

        private static PDFExportOptions BuildPDFOptions(ExportParams p, string baseFileName)
        {
            var options = new PDFExportOptions
            {
                // When Combine=true  → single file named "{baseFileName}.pdf"
                // When Combine=false → one file per view, named by Revit from view name
                FileName        = baseFileName,
                Combine         = p.Combine,

                // Always fit content to the page — avoids clipping
                ZoomType        = ZoomType.FitPage,

                // Full colour output; switch to GrayScale if file size is a concern
                ColorDepth      = ColorDepthType.Color,

                // High raster quality for embedded images / raster regions
                RasterQuality   = RasterQualityType.High,

                // Clean up print output — remove non-printed annotation elements
                HideScopeBoxes          = true,
                HideCropBoundaries      = true,
                HideReferencePlane      = true,
                HideUnreferencedViewTags = true,

                // "Default" uses each sheet's own paper size (best for sheets)
                // Override with A4/A3/Letter when exporting standalone views
                PaperFormat     = MapPaperFormat(p.PaperSize)
            };

            Console.WriteLine(
                $"[RevitPDFExport] PDFExportOptions → Combine={options.Combine}, " +
                $"PaperFormat={options.PaperFormat}, RasterQuality={options.RasterQuality}");

            return options;
        }

        private static ExportPaperFormat MapPaperFormat(string? paperSize)
        {
            switch ((paperSize ?? "Default").ToUpperInvariant())
            {
                case "A0":     return ExportPaperFormat.ISO_A0;
                case "A1":     return ExportPaperFormat.ISO_A1;
                case "A2":     return ExportPaperFormat.ISO_A2;
                case "A3":     return ExportPaperFormat.ISO_A3;
                case "A4":     return ExportPaperFormat.ISO_A4;
                case "LETTER": return ExportPaperFormat.ANSI_A_Letter;
                case "TABLOID":
                case "11X17":  return ExportPaperFormat.ANSI_B_Tabloid;
                default:       return ExportPaperFormat.Default;
            }
        }

        // ── Param loader ─────────────────────────────────────────────────────────

        private static ExportParams ReadParams()
        {
            const string fileName = "params.json";
            var path = Path.Combine(Directory.GetCurrentDirectory(), fileName);

            if (!File.Exists(path))
            {
                Console.WriteLine(
                    $"[RevitPDFExport] {fileName} not found — using defaults " +
                    "(operation: sheets, combine: false, paperSize: Default)");
                return new ExportParams();
            }

            try
            {
                var json   = File.ReadAllText(path);
                var parsed = JsonConvert.DeserializeObject<ExportParams>(json)
                             ?? new ExportParams();
                Console.WriteLine($"[RevitPDFExport] Loaded {fileName}: {json}");
                return parsed;
            }
            catch (Exception ex)
            {
                Console.WriteLine(
                    $"[RevitPDFExport] Failed to parse {fileName}: {ex.Message} — using defaults");
                return new ExportParams();
            }
        }
    }
}
