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
    // Explicit IDispatch interface for the automation object.
    // Using an explicit interface (rather than ClassInterface.AutoDispatch) guarantees that
    // IDispatch.GetIDsOfNames("Run") and "RunWithArguments" resolve correctly. AutoDispatch
    // auto-generation can silently exclude methods with certain parameter types on net48.
    [ComVisible(true)]
    [Guid("A7F8C3D2-E5B1-4F6A-9D3E-2C4B8A7F1E09")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IInventorAutomation
    {
        [DispId(1)]
        void Run(object doc);

        [DispId(2)]
        void RunWithArguments(object doc, object args);
    }

    // ClassInterfaceType.None: CCW exposes only IInventorAutomation (IDispatch).
    // No auto-generated class interface — IDispatch is IInventorAutomation directly.
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class BOMExtractorAutomation : IInventorAutomation
    {
        // No server reference needed — all BOM extraction works directly from the doc argument.
        public BOMExtractorAutomation() { }

        public void RunWithArguments(object doc, object args) => Run(doc);

        // SMOKE STUB: proves Run() is reachable from InventorCoreConsole.
        // Once result.json appears in output with source="smoke-stub", replace with RunReal().
        public void Run(object doc)
        {
            File.WriteAllText("activate_debug.txt",
                "[InventorBOMExtractor] Run() called at " + DateTime.UtcNow.ToString("o"),
                new UTF8Encoding(false));
            var report = new BOMReport
            {
                GeneratedAt = DateTime.UtcNow.ToString("o"),
                Source = "smoke-stub",
            };
            report.Errors.Add("SMOKE_STUB: Run() reached");
            WriteResult(report);
        }

        // Real BOM extraction — not called until smoke stub is confirmed green.
        private void RunReal(object doc)
        {
            var report = new BOMReport { GeneratedAt = DateTime.UtcNow.ToString("o") };
            try
            {
                using (new HeartBeat())
                {
                    dynamic d = doc;
                    report.Source = (string)d.FullFileName;
                    int docType = (int)d.DocumentType;
                    if (docType != 12292)
                    {
                        report.Errors.Add($"Input is not an assembly (.iam). DocumentType={docType}");
                    }
                    else
                    {
                        int total = 0;
                        report.TopLevelRows = ExtractBOM(d, ref total);
                        report.TotalComponents = total;
                    }
                }
            }
            catch (Exception ex)
            {
                report.Errors.Add(ex.Message);
            }
            finally
            {
                WriteResult(report);
            }
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
