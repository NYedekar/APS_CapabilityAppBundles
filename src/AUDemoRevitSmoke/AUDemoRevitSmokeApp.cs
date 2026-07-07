// ════════════════════════════════════════════════════════════════════════════
//  APS DA — Revit plugin entry point (template)
//
//  PROVEN PATTERN (loads headlessly on the Revit DA engine):
//    IExternalDBApplication  (NOT IExternalApplication — that's the UI variant)
//    + subscribe to DesignAutomationBridge.DesignAutomationReadyEvent
//
//  The FullClassName in the .addin MUST be AUDemoRevitSmoke.AUDemoRevitSmokeApp exactly.
//
//  net48 TRAP: write result.json with `new UTF8Encoding(false)` — Encoding.UTF8
//  emits a BOM that breaks strict JSON parsers downstream.
//
//  DA CONTEXT — what the worker does for you automatically:
//    - Detaches the model from central (worksharing). Never add WorksharingMode or
//      SetWorksharingMode calls — both are absent from the 2024 stubs and not needed.
//    - Provides the RevitDoc via DesignAutomationReadyEventArgs.
//
//  REPLACE: AUDemoRevitSmoke, AUDemoRevitSmokeApp, and the processing body.
// ════════════════════════════════════════════════════════════════════════════
using System;
using System.IO;
using System.Text;
using Newtonsoft.Json;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using DesignAutomationFramework;

namespace AUDemoRevitSmoke
{
    // ── Optional: params.json contract ──────────────────────────────────────
    // If your activity has a "params" get-arg, the MCP uploads the user's config
    // as params.json. Deserialise it here. Remove if not needed.
    // internal class RunParams
    // {
    //     [JsonProperty("inputMode")] public string? InputMode { get; set; }
    //     // add your fields here
    // }
    // ────────────────────────────────────────────────────────────────────────

    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    public class AUDemoRevitSmokeApp : IExternalDBApplication
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
                // Diagnostic: log all files visible in the working directory.
                // Useful when debugging missing inputs or unexpected file names.
                var wdFiles = Directory.GetFiles(Directory.GetCurrentDirectory());
                Console.WriteLine($"[AUDemoRevitSmokeApp] Working dir files: {string.Join(", ", System.Linq.Enumerable.Select(wdFiles, f => Path.GetFileName(f)))}");

                var doc = e.DesignAutomationData.RevitDoc
                    ?? throw new InvalidOperationException("Revit document could not be opened.");

                Console.WriteLine($"[AUDemoRevitSmokeApp] Processing: {doc.Title}");

                // ── Optional: read params.json ───────────────────────────────
                // RunParams? runParams = null;
                // if (File.Exists("params.json"))
                //     runParams = JsonConvert.DeserializeObject<RunParams>(File.ReadAllText("params.json"));
                // ────────────────────────────────────────────────────────────

                // ── your extraction / processing logic ──────────────────────

                // If modifying the model, wrap ALL changes in a Transaction:
                //
                // using (var tx = new Transaction(doc, "AUDemoRevitSmokeApp"))
                // {
                //     tx.Start();
                //
                //     // Parameter setting: prefer SetValueString for Double/Integer —
                //     // it respects project units. Fall back to p.Set(parsed) only
                //     // if SetValueString returns false.
                //     //
                //     // var p = elem.LookupParameter("ParamName");
                //     // if (p != null && !p.IsReadOnly)
                //     //     if (!p.SetValueString(value)) p.Set(double.Parse(value));
                //
                //     tx.Commit();
                // }

                var result = new { ok = true, title = doc.Title, framework = "net48" };
                // ────────────────────────────────────────────────────────────

                // ── Optional: save modified model ────────────────────────────
                // If the activity outputs a modified .rvt, save it before writing result.json:
                // doc.SaveAs(Path.GetFullPath("result.rvt"), new SaveAsOptions { OverwriteExistingFile = true });
                // ────────────────────────────────────────────────────────────

                // UTF-8 WITHOUT BOM — required for downstream JSON.parse to succeed.
                File.WriteAllText("result.json",
                    JsonConvert.SerializeObject(result, Formatting.Indented),
                    new UTF8Encoding(false));

                Console.WriteLine("[AUDemoRevitSmokeApp] Done — result.json written.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[AUDemoRevitSmokeApp] ERROR: {ex.Message}");
                var err = new { ok = false, error = ex.Message, stack = ex.StackTrace };
                File.WriteAllText("result.json",
                    JsonConvert.SerializeObject(err, Formatting.Indented),
                    new UTF8Encoding(false));
                // e.Succeeded stays true so the workitem completes and uploads result.json
                // (a diagnosable artifact) rather than reporting a hard worker failure.
            }
        }
    }
}
