using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace InventorBOMExtractor
{
    // Matches UpdateIPTParam.SampleAutomation pattern exactly:
    //   [ComVisible(true)] only — no [ClassInterface], no [Guid], no explicit IDispatch interface.
    // Default ClassInterface = AutoDispatch, which is how the official sample exposes Run.
    [ComVisible(true)]
    public class BOMExtractorAutomation
    {
        // Inventor DocumentTypeEnum (verified empirically: a .iam assembly reports 12291):
        //   kUnknownDocumentObject  = 12289 (0x3001)
        //   kPartDocumentObject     = 12290 (0x3002)
        //   kAssemblyDocumentObject = 12291 (0x3003)   <-- assemblies
        //   kDrawingDocumentObject  = 12292 (0x3004)
        private const int kAssemblyDocumentObject = 12291;

        public BOMExtractorAutomation() { }

        // Entry point invoked by InventorCoreConsole (DISPID 0x03001204 -> Automation -> Run).
        public void Run(object doc)
        {
            RunReal(doc);
        }

        // Set before each COM call so the result.json error pinpoints the exact failing step.
        private string _stage = "init";

        // Real BOM extraction.
        private void RunReal(object doc)
        {
            var report = new BOMReport { GeneratedAt = DateTime.UtcNow.ToString("o") };
            try
            {
                dynamic d = doc;
                _stage = "doc.FullFileName";
                report.Source = (string)d.FullFileName;
                _stage = "doc.DocumentType";
                int docType = (int)d.DocumentType;
                if (docType != kAssemblyDocumentObject)
                {
                    report.Errors.Add($"Input is not an assembly (.iam). DocumentType={docType}");
                }
                else
                {
                    int total = 0;
                    report.TopLevelRows = ExtractBOM(d, report, ref total);
                    report.TotalComponents = total;
                }
            }
            catch (Exception ex)
            {
                report.Errors.Add($"[stage={_stage}] {ex.GetType().Name}: {ex.Message}");
            }
            finally
            {
                WriteResult(report);
            }
        }

        private List<BOMRow> ExtractBOM(dynamic asmDoc, BOMReport report, ref int total)
        {
            var rows = new List<BOMRow>();
            _stage = "asmDoc.ComponentDefinition";
            dynamic compDef = asmDoc.ComponentDefinition;
            _stage = "compDef.BOM";
            dynamic bom = compDef.BOM;
            _stage = "bom.StructuredViewEnabled=true";
            bom.StructuredViewEnabled = true;
            _stage = "bom.BOMViews";
            dynamic views = bom.BOMViews;

            // Late-bound COM: passing a managed string to Item(VARIANT) via `dynamic` does NOT
            // marshal correctly (E_INVALIDARG). Iterate by integer index and match on Title.
            _stage = "views.Count";
            int count = (int)views.Count;
            report.Errors.Add($"DIAG: BOMViews.Count={count}");
            dynamic view = null;
            for (int i = 1; i <= count; i++)
            {
                _stage = $"views.Item({i})";
                dynamic v = views.Item(i);
                _stage = $"views.Item({i}).Title";
                string title;
                try { title = (string)v.Title; } catch { title = "(no Title)"; }
                report.Errors.Add($"DIAG: BOMView[{i}].Title='{title}'");
                if (view == null && title.IndexOf("Structured", StringComparison.OrdinalIgnoreCase) >= 0)
                    view = v;
            }
            // Fallback: first view if no title matched (en-US structured view is usually index 1).
            if (view == null && count >= 1) { _stage = "views.Item(1) fallback"; view = views.Item(1); }
            if (view == null) return rows;
            _stage = "view.BOMRows";
            WalkRows(view.BOMRows, rows, ref total);
            return rows;
        }

        private void WalkRows(dynamic rowEnum, List<BOMRow> target, ref int total)
        {
            foreach (dynamic row in rowEnum)
            {
                total++;
                var entry = new BOMRow
                {
                    ItemNumber = SafeString(() => (string)row.ItemNumber),
                    Quantity   = SafeDouble(() => (double)row.ItemQuantity),
                };
                try
                {
                    // .Item(...) not [..] — COM default-member binding (see ExtractBOM note).
                    dynamic compDef = row.ComponentDefinitions.Item(1);
                    dynamic compDoc = compDef.Document;
                    entry.IsAssembly = (int)compDoc.DocumentType == kAssemblyDocumentObject;
                    dynamic propSets = compDoc.PropertySets;
                    dynamic trackingSet = propSets.Item("Design Tracking Properties");
                    entry.PartNumber  = SafeProp(trackingSet, "Part Number");
                    entry.Description = SafeProp(trackingSet, "Description");
                    entry.Material    = SafeProp(trackingSet, "Material");
                    entry.Mass        = SafeProp(trackingSet, "Mass");
                    string unit       = SafeProp(trackingSet, "Unit Quantity");
                    entry.Unit        = string.IsNullOrWhiteSpace(unit) ? "ea" : unit;
                }
                catch { }
                try
                {
                    dynamic childRows = row.ChildRows;
                    if (childRows != null)
                        WalkRows(childRows, entry.ChildRows, ref total);
                }
                catch { }
                target.Add(entry);
            }
        }

        private static string SafeProp(dynamic propSet, string name)
        {
            try { return propSet.Item(name).Value?.ToString() ?? string.Empty; }
            catch { return string.Empty; }
        }

        private static string SafeString(Func<string> fn)
        {
            try { return fn() ?? string.Empty; } catch { return string.Empty; }
        }

        private static double SafeDouble(Func<double> fn)
        {
            try { return fn(); } catch { return 0; }
        }

        private static readonly JsonSerializerSettings JsonSettings = new JsonSerializerSettings
        {
            // camelCase: source, generatedAt, totalComponents, topLevelRows, errors, itemNumber, …
            ContractResolver = new CamelCasePropertyNamesContractResolver(),
            Formatting = Formatting.Indented,
        };

        private static void WriteResult(BOMReport report)
        {
            string json = JsonConvert.SerializeObject(report, JsonSettings);
            File.WriteAllText("result.json", json, new UTF8Encoding(false));
        }
    }
}
