using Autodesk.Forge.DesignAutomation.Inventor.Utils;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace InventorBOMExtractor
{
    // All Inventor API calls are late-bound via dynamic (IDispatch).
    // No autodesk.inventor.interop.dll required at compile time.
    // AutoDispatch: pure IDispatch interface — no type library generation (AutoDual TLB
    // generation can silently drop methods with dynamic/object parameters).
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class BOMExtractorAutomation
    {
        // No server reference needed — all BOM extraction works directly from the doc argument.
        public BOMExtractorAutomation() { }

        // Called by Inventor DA after Activate(). Parameter type must be object (not dynamic)
        // so the ClassInterface IDispatch correctly exposes Run via GetIDsOfNames.
        public void Run(object doc)
        {
            // SMOKE STUB: prove Run() is called before attempting any Inventor API calls.
            // If result.json appears in output with smoke=true, Run() works end-to-end.
            // Remove this stub once confirmed and restore real BOM extraction.
            File.WriteAllText("activate_debug.txt", "[InventorBOMExtractor] Run() was called at " + DateTime.UtcNow.ToString("o"), new UTF8Encoding(false));
            var smokeReport = new BOMReport
            {
                GeneratedAt = DateTime.UtcNow.ToString("o"),
                Source = "smoke-stub",
                TotalComponents = 0,
                TopLevelRows = new List<BOMRow>(),
            };
            smokeReport.Errors.Add("SMOKE_STUB: Run() reached — replace with real BOM logic");
            WriteResult(smokeReport);
            Trace.TraceInformation("[InventorBOMExtractor] SMOKE STUB: result.json written. Run() confirmed callable.");
            return;
            // ── real BOM extraction below (unreachable until smoke passes) ──
#pragma warning disable CS0162
            var report = new BOMReport { GeneratedAt = DateTime.UtcNow.ToString("o") };
            try
            {
                using (new HeartBeat())
                {
                    dynamic dynDoc = doc;
                    report.Source = (string)dynDoc.FullFileName;
                    Trace.TraceInformation("[InventorBOMExtractor] Run — doc={0}", report.Source);

                    // DocumentTypeEnum.kAssemblyDocumentObject == 12292 (0x3004)
                    int docType = (int)dynDoc.DocumentType;
                    if (docType != 12292)
                    {
                        report.Errors.Add($"Input is not an assembly (.iam). DocumentType={docType}");
                        Trace.TraceWarning("[InventorBOMExtractor] Not an assembly.");
                    }
                    else
                    {
                        int total = 0;
                        report.TopLevelRows = ExtractBOM(dynDoc, ref total);
                        report.TotalComponents = total;
                        Trace.TraceInformation("[InventorBOMExtractor] BOM: {0} top-level, {1} total", report.TopLevelRows.Count, total);
                    }
                }
            }
            catch (Exception ex)
            {
                Trace.TraceError("[InventorBOMExtractor] Exception: {0}", ex);
                report.Errors.Add(ex.Message);
            }
            finally
            {
                WriteResult(report);
            }
#pragma warning restore CS0162
        }

        private List<BOMRow> ExtractBOM(dynamic asmDoc, ref int total)
        {
            var rows = new List<BOMRow>();
            dynamic compDef = asmDoc.ComponentDefinition;
            dynamic bom = compDef.BOM;
            bom.StructuredViewEnabled = true;
            bom.StructuredViewFirstLevelOnly = false;

            dynamic views = bom.BOMViews;
            dynamic view = views["Structured"];
            if (view == null)
            {
                Trace.TraceWarning("[InventorBOMExtractor] 'Structured' BOM view not found.");
                return rows;
            }

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
                    dynamic compDef = row.ComponentDefinitions[1];
                    dynamic compDoc = compDef.Document;

                    // DocumentTypeEnum.kAssemblyDocumentObject == 12292
                    entry.IsAssembly = (int)compDoc.DocumentType == 12292;

                    dynamic propSets = compDoc.PropertySets;
                    dynamic trackingSet = propSets["Design Tracking Properties"];

                    entry.PartNumber  = SafeProp(trackingSet, "Part Number");
                    entry.Description = SafeProp(trackingSet, "Description");
                    entry.Material    = SafeProp(trackingSet, "Material");
                    entry.Mass        = SafeProp(trackingSet, "Mass");
                    string unit       = SafeProp(trackingSet, "Unit Quantity");
                    entry.Unit        = string.IsNullOrWhiteSpace(unit) ? "ea" : unit;
                }
                catch (Exception ex)
                {
                    Trace.TraceWarning("[InventorBOMExtractor] Row {0} property read: {1}", entry.ItemNumber, ex.Message);
                }

                try
                {
                    dynamic childRows = row.ChildRows;
                    if (childRows != null)
                        WalkRows(childRows, entry.ChildRows, ref total);
                }
                catch { /* no child rows */ }

                target.Add(entry);
            }
        }

        private static string SafeProp(dynamic propSet, string name)
        {
            try { return propSet[name].Value?.ToString() ?? string.Empty; }
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

        private static void WriteResult(BOMReport report)
        {
            string json = JsonConvert.SerializeObject(report, Formatting.Indented);
            File.WriteAllText("result.json", json, new UTF8Encoding(false));
            Trace.TraceInformation("[InventorBOMExtractor] result.json written ({0} chars).", json.Length);
        }
    }
}
