using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using DesignAutomationFramework;
using Newtonsoft.Json;
using System;
using System.IO;
using System.Text;

namespace RevitIFCExport
{
    // ─── Input model ────────────────────────────────────────────────────────────

    public class ExportParams
    {
        /// <summary>"IFC2x3" | "IFC2x3CV2" | "IFC4"</summary>
        public string IfcVersion { get; set; } = "IFC2x3CV2";

        /// <summary>Export base quantities (QTO) to IFC property sets.</summary>
        public bool ExportBaseQuantities { get; set; } = false;

        /// <summary>Split walls and columns at floor levels.</summary>
        public bool WallAndColumnSplitting { get; set; } = false;

        /// <summary>Space boundary level: 0 = none, 1 = 1st level, 2 = 2nd level.</summary>
        public int SpaceBoundaryLevel { get; set; } = 0;
    }

    // ─── Output model ───────────────────────────────────────────────────────────

    public class ExportResult
    {
        public bool Success { get; set; }
        public string IfcVersion { get; set; } = string.Empty;
        public string DocumentTitle { get; set; } = string.Empty;
        public string OutputFile { get; set; } = string.Empty;
        public long FileSizeBytes { get; set; }
        public string Timestamp { get; set; } = string.Empty;
        public string? Error { get; set; }
    }

    // ─── Plugin entry point ─────────────────────────────────────────────────────

    /// <summary>
    /// APS Design Automation plugin for Revit → IFC export.
    ///
    /// CRITICAL RULES (from hard-won experience):
    ///   - Implement IExternalDBApplication, NOT IExternalApplication
    ///   - Register DesignAutomationReadyEvent in OnStartup — the ONLY entry point DA uses
    ///   - Always set e.Succeeded = true before logic; set false on unrecoverable error
    ///   - Write output files to working directory — DA picks them up as output args
    ///   - Use Console.WriteLine for logging — appears in the DA report
    ///   - Zero UI code — no TaskDialog, no WPF, nothing
    ///   - ElementId.IntegerValue deprecated in Revit 2024+ — use ElementId.Value
    ///   - net48: no tuple deconstruction on KeyValuePair — use .Key / .Value
    ///   - result.json must be written with new UTF8Encoding(false) — NO BOM
    /// </summary>
    [Autodesk.Revit.Attributes.Regeneration(
        Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    public class RevitIFCExportApp : IExternalDBApplication
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
            e.Succeeded = true;

            var result = new ExportResult { Timestamp = DateTime.UtcNow.ToString("o") };

            try
            {
                var doc = e.DesignAutomationData.RevitDoc
                    ?? throw new InvalidOperationException("RevitDoc is null — document failed to open");

                Console.WriteLine($"[RevitIFCExport] Document opened: {doc.Title}");
                result.DocumentTitle = doc.Title;

                var p = ReadParams();
                var ifcVersion = ParseVersion(p.IfcVersion);

                Console.WriteLine($"[RevitIFCExport] IFC version           : {p.IfcVersion} ({ifcVersion})");
                Console.WriteLine($"[RevitIFCExport] ExportBaseQuantities  : {p.ExportBaseQuantities}");
                Console.WriteLine($"[RevitIFCExport] WallAndColumnSplitting: {p.WallAndColumnSplitting}");
                Console.WriteLine($"[RevitIFCExport] SpaceBoundaryLevel    : {p.SpaceBoundaryLevel}");

                var workDir   = Directory.GetCurrentDirectory();
                var exportDir = Path.Combine(workDir, "ifc_output");
                if (Directory.Exists(exportDir)) Directory.Delete(exportDir, true);
                Directory.CreateDirectory(exportDir);

                var options = new IFCExportOptions
                {
                    FileVersion            = ifcVersion,
                    ExportBaseQuantities   = p.ExportBaseQuantities,
                    WallAndColumnSplitting = p.WallAndColumnSplitting,
                    SpaceBoundaryLevel     = p.SpaceBoundaryLevel,
                };

                // IFCExportOptions.Export requires an open transaction — wrap and commit.
                bool ok;
                using (var tx = new Transaction(doc, "ExportIFC"))
                {
                    tx.Start();
                    ok = doc.Export(exportDir, "result", options);
                    tx.Commit();
                }

                if (!ok)
                    throw new InvalidOperationException(
                        "Document.Export returned false — IFC export failed. " +
                        "Ensure the model is not corrupt and Revit can process it.");

                // Revit writes {name}.ifc in the export folder
                var ifcSrc  = Path.Combine(exportDir, "result.ifc");
                if (!File.Exists(ifcSrc))
                    throw new InvalidOperationException(
                        $"IFC file not found at expected path: {ifcSrc}");

                // Copy to working dir so DA picks it up as the 'ifcFile' output argument
                var ifcDest = Path.Combine(workDir, "result.ifc");
                File.Copy(ifcSrc, ifcDest, overwrite: true);

                result.Success       = true;
                result.IfcVersion    = p.IfcVersion;
                result.OutputFile    = "result.ifc";
                result.FileSizeBytes = new FileInfo(ifcDest).Length;

                Console.WriteLine(
                    $"[RevitIFCExport] Done — result.ifc ({result.FileSizeBytes:N0} bytes)");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[RevitIFCExport] ERROR: {ex}");
                result.Success = false;
                result.Error   = ex.Message;
                e.Succeeded    = false;
            }
            finally
            {
                // Always write result.json — even on failure — with NO BOM
                File.WriteAllText(
                    "result.json",
                    JsonConvert.SerializeObject(result, Formatting.Indented),
                    new UTF8Encoding(false));
                Console.WriteLine("[RevitIFCExport] result.json written");
            }
        }

        // ── IFC version parser ───────────────────────────────────────────────────

        private static IFCVersion ParseVersion(string? version)
        {
            switch ((version ?? "IFC2x3CV2").ToUpperInvariant())
            {
                case "IFC4":
                case "IFC4RV":
                    return IFCVersion.IFC4;
                case "IFC2X3":
                    return IFCVersion.IFC2x3;
                default:
                    // IFC2x3CV2 — Coordination View 2.0 (most broadly accepted format)
                    return IFCVersion.IFC2x3CV2;
            }
        }

        // ── Param loader ─────────────────────────────────────────────────────────

        private static ExportParams ReadParams()
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), "params.json");
            if (!File.Exists(path))
            {
                Console.WriteLine("[RevitIFCExport] params.json not found — using defaults");
                return new ExportParams();
            }

            try
            {
                var json   = File.ReadAllText(path);
                var parsed = JsonConvert.DeserializeObject<ExportParams>(json) ?? new ExportParams();
                Console.WriteLine($"[RevitIFCExport] Loaded params.json: {json}");
                return parsed;
            }
            catch (Exception ex)
            {
                Console.WriteLine(
                    $"[RevitIFCExport] Failed to parse params.json: {ex.Message} — using defaults");
                return new ExportParams();
            }
        }
    }
}
