// ════════════════════════════════════════════════════════════════════════════
//  APS DA — Inventor automation class (SMOKE-STUB template)
//
//  Base pattern VERIFIED, shipped green: adapted from InventorBOMExtractor's
//  BOMExtractorAutomation.cs (src/InventorBOMExtractor/BOMExtractorAutomation.cs
//  in APS_CapabilityAppBundles). That real file walks the assembly occurrence
//  tree to build a BOM. This template REPLACES that logic with a minimal
//  smoke stub that only proves the plugin loads and Run() is reached —
//  it does NOT do real Inventor-document processing.
//
//  FOR THE REAL RECIPE (walking ComponentDefinition/Occurrences, late-bound
//  Type.InvokeMember helpers for Item/Prop/SetProp because C# `dynamic` does
//  NOT reliably marshal args to a COM Item(VARIANT) member, PropertySets
//  lookups, MassProperties, etc.) SEE reference/appbuilder-guide.md §9.
//
//  REPLACE: AUDemoInvCheck, AUDemoInvCheckAutomation, and the body of Run() with
//  your real extraction/processing logic once the smoke stub is green.
// ════════════════════════════════════════════════════════════════════════════
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace AUDemoInvCheck
{
    // Matches UpdateIPTParam.SampleAutomation pattern: [ComVisible(true)] only
    // (default AutoDispatch class interface). InventorCoreConsole calls Run(doc) on it.
    [ComVisible(true)]
    public class AUDemoInvCheckAutomation
    {
        // Inventor DocumentTypeEnum (verified: a .iam assembly reports 12291):
        //   kPartDocumentObject = 12290, kAssemblyDocumentObject = 12291, kDrawingDocumentObject = 12292
        private const int kAssemblyDocumentObject = 12291;

        // Set before each COM call so result.json errors pinpoint the failing step.
        private string _stage = "init";

        public AUDemoInvCheckAutomation() { }

        // Entry point invoked by InventorCoreConsole (DISPID 0x03001204 -> Automation -> Run).
        public void Run(object doc)
        {
            RunReal(doc);
        }

        // ---- Late-bound COM helpers (Type.InvokeMember) --------------------------------
        // C# `dynamic` does NOT reliably marshal args to a COM Item(VARIANT) member
        // (throws E_INVALIDARG). InvokeMember marshals string/int -> VARIANT correctly.
        // Keep these even in the smoke stub — the real recipe (guide §9) needs them
        // the moment you replace the stub body below.
        private static object Prop(object o, string name)
            => o.GetType().InvokeMember(name, BindingFlags.GetProperty, null, o, null);

        private static void SetProp(object o, string name, object val)
            => o.GetType().InvokeMember(name, BindingFlags.SetProperty, null, o, new[] { val });

        private static object Item(object collection, object index)
            => collection.GetType().InvokeMember("Item",
                   BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, collection, new[] { index });

        private static string Str(object o) => o?.ToString() ?? string.Empty;
        // --------------------------------------------------------------------------------

        private void RunReal(object doc)
        {
            var report = new SmokeReport { Smoke = "ok", GeneratedAt = DateTime.UtcNow.ToString("o") };
            try
            {
                _stage = "doc.FullFileName";
                report.Source = Str(Prop(doc, "FullFileName"));
                _stage = "doc.DocumentType";
                int docType = (int)Prop(doc, "DocumentType");
                report.DocumentType = docType;

                // ── your real extraction / processing logic goes here ──────────────────
                // See reference/appbuilder-guide.md §9 for the BOM-walking recipe
                // (ComponentDefinition -> Occurrences -> recurse into SubOccurrences),
                // which is the pattern this stub replaces.
                // ────────────────────────────────────────────────────────────────────────
            }
            catch (Exception ex)
            {
                // InvokeMember wraps COM errors in TargetInvocationException — unwrap to the root.
                var inner = ex;
                while (inner.InnerException != null) inner = inner.InnerException;
                report.Errors.Add($"[stage={_stage}] {inner.GetType().Name}: {inner.Message} (HRESULT=0x{inner.HResult:X8})");
            }
            finally
            {
                WriteResult(report);
            }
        }

        private static readonly JsonSerializerSettings JsonSettings = new JsonSerializerSettings
        {
            ContractResolver = new CamelCasePropertyNamesContractResolver(),
            Formatting = Formatting.Indented,
        };

        private static void WriteResult(SmokeReport report)
        {
            string json = JsonConvert.SerializeObject(report, JsonSettings);
            // UTF8Encoding(false) = no BOM — required for downstream JSON.parse to succeed.
            File.WriteAllText("result.json", json, new UTF8Encoding(false));
        }
    }

    internal class SmokeReport
    {
        public string Smoke { get; set; } = "ok";
        public string GeneratedAt { get; set; } = string.Empty;
        public string Source { get; set; } = string.Empty;
        public int DocumentType { get; set; }
        public System.Collections.Generic.List<string> Errors { get; set; } = new System.Collections.Generic.List<string>();
    }
}
