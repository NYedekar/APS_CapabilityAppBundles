using System;
using System.Collections.Generic;
using System.Globalization;
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

        // Set before each COM call so result.csv errors pinpoint the failing step.
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

        // The structured BOM-view API (BOMViews.Item("Structured").BOMRows) throws E_FAIL under
        // the headless Inventor Server used by Design Automation. The assembly OCCURRENCE tree
        // is fully available (verified: 63 occurrences resolved), so we build the BOM from it:
        // walk occurrences, roll up identical part numbers into a quantity per level, recurse
        // into sub-assemblies for hierarchy.
        private List<BOMRow> ExtractBOM(object asmDoc, ref int total)
        {
            var rows = new List<BOMRow>();
            _stage = "ComponentDefinition";
            object compDef = Prop(asmDoc, "ComponentDefinition");
            _stage = "Occurrences";
            object occs = Prop(compDef, "Occurrences");
            WalkOccurrences(occs, rows, ref total);
            return rows;
        }

        private void WalkOccurrences(object occs, List<BOMRow> target, ref int total)
        {
            int count = (int)Prop(occs, "Count");
            var order = new List<string>();
            var rowByKey = new Dictionary<string, BOMRow>();
            var firstOccByKey = new Dictionary<string, object>();

            for (int i = 1; i <= count; i++)
            {
                object occ;
                try { occ = Item(occs, i); } catch { continue; }

                var entry = new BOMRow { Quantity = 1, Unit = "ea" };
                object def = null;
                try { def = Prop(occ, "Definition"); } catch { }
                if (def != null) FillProps(def, entry);

                // Roll up identical parts (by part number; fall back to occurrence name).
                string key = !string.IsNullOrWhiteSpace(entry.PartNumber)
                    ? entry.PartNumber
                    : SafeStr(() => Str(Prop(occ, "Name")));

                if (rowByKey.TryGetValue(key, out var existing))
                {
                    existing.Quantity += 1;
                }
                else
                {
                    rowByKey[key] = entry;
                    firstOccByKey[key] = occ;
                    order.Add(key);
                }
            }

            int item = 0;
            foreach (var key in order)
            {
                item++;
                total++;
                var entry = rowByKey[key];
                entry.ItemNumber = item.ToString();

                if (entry.IsAssembly)
                {
                    try
                    {
                        object sub = Prop(firstOccByKey[key], "SubOccurrences");
                        if (sub != null && (int)Prop(sub, "Count") > 0)
                            WalkOccurrences(sub, entry.ChildRows, ref total);
                    }
                    catch { }
                }
                target.Add(entry);
            }
        }

        // Fill part metadata from a ComponentDefinition (handles virtual components w/o Document).
        private void FillProps(object def, BOMRow entry)
        {
            object propHolder = def;
            try { object docObj = Prop(def, "Document"); if (docObj != null) propHolder = docObj; }
            catch { propHolder = def; }

            try { entry.IsAssembly = (int)Prop(propHolder, "DocumentType") == kAssemblyDocumentObject; }
            catch { }

            try
            {
                object propSets = Prop(propHolder, "PropertySets");
                object dtp = Item(propSets, "Design Tracking Properties");
                entry.PartNumber  = PropVal(dtp, "Part Number");
                entry.Description = PropVal(dtp, "Description");
                entry.Material    = PropVal(dtp, "Material");
            }
            catch { }

            // Numeric mass from computed mass properties (not the "Mass" iProperty string).
            try { object mp = Prop(def, "MassProperties"); entry.Mass = Str(Prop(mp, "Mass")); }
            catch { }
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

        // The BOM tree is flattened into a single CSV: each row carries a Level (0 = top level,
        // incremented per sub-assembly depth) so the hierarchy is preserved in a tabular form.
        private static void WriteResult(BOMReport report)
        {
            var sb = new StringBuilder();
            if (report.TopLevelRows.Count == 0 && report.Errors.Count > 0)
            {
                // Extraction failed — emit the diagnostics as a single-column CSV.
                sb.Append("error\r\n");
                foreach (var e in report.Errors)
                    sb.Append(Csv(e) + "\r\n");
            }
            else
            {
                sb.Append("Level,ItemNumber,PartNumber,Description,Quantity,Unit,Material,Mass,IsAssembly\r\n");
                foreach (var row in report.TopLevelRows)
                    AppendRow(sb, row, 0);
            }
            // UTF-8 WITHOUT BOM — keeps strict parsers happy.
            File.WriteAllText("result.csv", sb.ToString(), new UTF8Encoding(false));
        }

        private static void AppendRow(StringBuilder sb, BOMRow row, int level)
        {
            sb.Append(string.Join(",",
                level.ToString(CultureInfo.InvariantCulture),
                Csv(row.ItemNumber),
                Csv(row.PartNumber),
                Csv(row.Description),
                row.Quantity.ToString(CultureInfo.InvariantCulture),
                Csv(row.Unit),
                Csv(row.Material),
                Csv(row.Mass),
                row.IsAssembly ? "true" : "false"));
            sb.Append("\r\n");
            foreach (var child in row.ChildRows)
                AppendRow(sb, child, level + 1);
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
}
