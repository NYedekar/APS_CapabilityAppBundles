using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace InventorBOMExtractor
{
    // Matches UpdateIPTParam.SampleAutomation pattern: [ComVisible(true)] only
    // (default AutoDispatch class interface). InventorCoreConsole calls Run(doc) on it.
    [ComVisible(true)]
    public class BOMExtractorAutomation
    {
        // Inventor DocumentTypeEnum (verified: a .iam assembly reports 12291):
        //   kPartDocumentObject = 12290, kAssemblyDocumentObject = 12291, kDrawingDocumentObject = 12292
        private const int kAssemblyDocumentObject = 12291;

        // Set before each COM call so result.json errors pinpoint the failing step.
        private string _stage = "init";

        public BOMExtractorAutomation() { }

        // Entry point invoked by InventorCoreConsole (DISPID 0x03001204 -> Automation -> Run).
        public void Run(object doc)
        {
            RunReal(doc);
        }

        // ---- Late-bound COM helpers (Type.InvokeMember) --------------------------------
        // C# `dynamic` does NOT reliably marshal args to a COM Item(VARIANT) member
        // (throws E_INVALIDARG). InvokeMember marshals string/int -> VARIANT correctly.
        private static object Prop(object o, string name)
            => o.GetType().InvokeMember(name, BindingFlags.GetProperty, null, o, null);

        private static void SetProp(object o, string name, object val)
            => o.GetType().InvokeMember(name, BindingFlags.SetProperty, null, o, new[] { val });

        private static object Item(object collection, object index)
            => collection.GetType().InvokeMember("Item",
                   BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, collection, new[] { index });

        private static string Str(object o) => o?.ToString() ?? string.Empty;

        private static double ToDouble(object o)
        {
            try { return Convert.ToDouble(o); } catch { return 0; }
        }
        // --------------------------------------------------------------------------------

        private void RunReal(object doc)
        {
            var report = new BOMReport { GeneratedAt = DateTime.UtcNow.ToString("o") };
            try
            {
                _stage = "doc.FullFileName";
                report.Source = Str(Prop(doc, "FullFileName"));
                _stage = "doc.DocumentType";
                int docType = (int)Prop(doc, "DocumentType");
                if (docType != kAssemblyDocumentObject)
                {
                    report.Errors.Add($"Input is not an assembly (.iam). DocumentType={docType}");
                }
                else
                {
                    int total = 0;
                    report.TopLevelRows = ExtractBOM(doc, ref total);
                    report.TotalComponents = total;
                }
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

        private List<BOMRow> ExtractBOM(object asmDoc, ref int total)
        {
            var rows = new List<BOMRow>();
            _stage = "ComponentDefinition";
            object compDef = Prop(asmDoc, "ComponentDefinition");
            _stage = "BOM";
            object bom = Prop(compDef, "BOM");
            _stage = "StructuredViewFirstLevelOnly=false";
            try { SetProp(bom, "StructuredViewFirstLevelOnly", false); } catch { }
            _stage = "StructuredViewEnabled=true";
            SetProp(bom, "StructuredViewEnabled", true);
            _stage = "BOMViews";
            object views = Prop(bom, "BOMViews");
            // BOMViews.Item("Structured") is the documented access; InvokeMember marshals the
            // string to a VARIANT properly (unlike dynamic).
            _stage = "BOMViews.Item(\"Structured\")";
            object view = Item(views, "Structured");
            _stage = "view.BOMRows";
            object bomRows = Prop(view, "BOMRows");
            WalkRows(bomRows, rows, ref total);
            return rows;
        }

        private void WalkRows(object bomRows, List<BOMRow> target, ref int total)
        {
            _stage = "BOMRows.Count";
            int count = (int)Prop(bomRows, "Count");
            for (int i = 1; i <= count; i++)
            {
                total++;
                _stage = $"BOMRows.Item({i})";
                object row = Item(bomRows, i);

                var entry = new BOMRow
                {
                    ItemNumber = SafeStr(() => Str(Prop(row, "ItemNumber"))),
                    Quantity   = SafeNum(() => ToDouble(Prop(row, "ItemQuantity"))),
                };

                try
                {
                    object compDefs = Prop(row, "ComponentDefinitions");
                    object cd = Item(compDefs, 1);

                    // Virtual components have no Document — read PropertySets off the def itself.
                    object propHolder = cd;
                    try { object docObj = Prop(cd, "Document"); if (docObj != null) propHolder = docObj; }
                    catch { propHolder = cd; }

                    try { entry.IsAssembly = (int)Prop(propHolder, "DocumentType") == kAssemblyDocumentObject; }
                    catch { }

                    object propSets = Prop(propHolder, "PropertySets");
                    object dtp = Item(propSets, "Design Tracking Properties");
                    entry.PartNumber  = PropVal(dtp, "Part Number");
                    entry.Description = PropVal(dtp, "Description");
                    entry.Material    = PropVal(dtp, "Material");
                    string unit       = PropVal(dtp, "Unit Quantity");
                    entry.Unit        = string.IsNullOrWhiteSpace(unit) ? "ea" : unit;

                    // Numeric mass from computed mass properties (not the iProperty string).
                    try { object mp = Prop(cd, "MassProperties"); entry.Mass = Str(Prop(mp, "Mass")); }
                    catch { }
                }
                catch { }

                // ChildRows can be null (parts-only views / leaf rows).
                try
                {
                    object childRows = Prop(row, "ChildRows");
                    if (childRows != null)
                        WalkRows(childRows, entry.ChildRows, ref total);
                }
                catch { }

                target.Add(entry);
            }
        }

        private static string PropVal(object propSet, string name)
        {
            try { return Str(Prop(Item(propSet, name), "Value")); }
            catch { return string.Empty; }
        }

        private static string SafeStr(Func<string> fn)
        {
            try { return fn() ?? string.Empty; } catch { return string.Empty; }
        }

        private static double SafeNum(Func<double> fn)
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
