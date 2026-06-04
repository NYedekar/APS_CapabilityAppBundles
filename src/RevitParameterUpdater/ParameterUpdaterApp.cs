// ════════════════════════════════════════════════════════════════════════════
//  RevitParameterUpdater — APS Design Automation (Revit R2026, net48)
//
//  Updates element parameters in a Revit model from a CSV, XLSX, or JSON input.
//
//  INPUTS (placed in working dir by DA before revitcoreconsole runs):
//    input.rvt         — the Revit model to modify
//    params_input.dat  — CSV / XLSX / JSON with change requests
//    params.json       — control JSON: { "inputMode": "full_file|delta_file|text_input" }
//
//  OUTPUTS:
//    result.rvt        — modified model (always written; unchanged if 0 changes)
//    result.json       — { ok, summary: { total, applied, skipped, errors }, changes: [...] }
//
//  INPUT MODES:
//    delta_file  — every row in params_input is a change to apply
//    text_input  — same as delta_file; params_input is a JSON array from chat
//    full_file   — params_input contains ALL elements; only rows whose
//                  NewValue differs from the current model value are applied
//
//  COLUMN FORMAT (CSV / XLSX):
//    ElementId (optional)  ElementName  Category (optional)  Parameter  NewValue
//
//  JSON format (text_input):
//    [{"elementName":"...","parameter":"...","value":"...","elementId":"...","category":"..."}]
//
//  NET48 TRAPS (all guarded here):
//    - ElementId.Value (long), NOT .IntegerValue (int) — removed in Revit 2024+
//    - new UTF8Encoding(false) for result.json — Encoding.UTF8 emits BOM
//    - No Dictionary.GetValueOrDefault — use TryGetValue
//    - SetValueString used for Double/Integer; falls back to parsed Set()
// ════════════════════════════════════════════════════════════════════════════

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using ExcelDataReader;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using DesignAutomationFramework;

namespace RevitParameterUpdater
{
    // ─── Entry point ─────────────────────────────────────────────────────────

    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    public class ParameterUpdaterApp : IExternalDBApplication
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
            UpdateResult result;

            try
            {
                var doc = e.DesignAutomationData.RevitDoc
                    ?? throw new InvalidOperationException("Revit document could not be opened.");

                Console.WriteLine($"[RevitParameterUpdater] Processing: {doc.Title}");

                // ── 1. Read control params + optional inline changes ─────────
                string inputMode = "delta_file";
                List<ChangeRequest>? inlineChanges = null;

                if (File.Exists("params.json"))
                {
                    try
                    {
                        var paramsText = File.ReadAllText("params.json", Encoding.UTF8);
                        var ctrl = JsonConvert.DeserializeObject<RunParams>(paramsText);
                        if (ctrl?.InputMode != null)
                            inputMode = ctrl.InputMode.ToLowerInvariant().Trim();
                        if (ctrl?.Changes != null)
                        {
                            inlineChanges = InputParser.ParseJson(ctrl.Changes.ToString());
                            Console.WriteLine($"[RevitParameterUpdater] Inline changes in params.json: {inlineChanges.Count}");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[RevitParameterUpdater] WARNING: could not parse params.json — {ex.Message}; defaulting to delta_file");
                    }
                }
                Console.WriteLine($"[RevitParameterUpdater] inputMode={inputMode}");

                // ── 2. Parse change requests ─────────────────────────────────
                // Diagnostic: log what's actually in the working directory.
                var wdFiles = Directory.GetFiles(Directory.GetCurrentDirectory());
                Console.WriteLine($"[RevitParameterUpdater] Working dir files: {string.Join(", ", wdFiles.Select(f => Path.GetFileName(f)))}");

                List<ChangeRequest> requests;
                if (File.Exists("params_input.dat"))
                {
                    var fSize = new FileInfo("params_input.dat").Length;
                    var magic = new byte[Math.Min(8, (int)fSize)];
                    using (var fs = File.OpenRead("params_input.dat"))
                        fs.Read(magic, 0, magic.Length);
                    Console.WriteLine($"[RevitParameterUpdater] params_input.dat: {fSize} bytes  magic={BitConverter.ToString(magic)}");

                    requests = InputParser.Parse("params_input.dat");
                    Console.WriteLine($"[RevitParameterUpdater] Parsed {requests.Count} change request(s) from params_input.dat");

                    // If the file parsed to 0 rows (e.g. wrong content type) fall back to inline changes.
                    if (requests.Count == 0 && inlineChanges?.Count > 0)
                    {
                        Console.WriteLine($"[RevitParameterUpdater] params_input.dat yielded 0 rows — using inline changes from params.json");
                        requests = inlineChanges;
                    }
                }
                else if (inlineChanges?.Count > 0)
                {
                    Console.WriteLine($"[RevitParameterUpdater] params_input.dat absent — using {inlineChanges.Count} inline change(s) from params.json");
                    requests = inlineChanges;
                }
                else
                {
                    throw new InvalidOperationException(
                        "No change requests: params_input.dat is absent/empty and params.json has no 'changes' array.");
                }
                Console.WriteLine($"[RevitParameterUpdater] Total change requests to process: {requests.Count}");

                // ── 3. Apply changes ────────────────────────────────────────
                var updater = new Updater(doc);
                var changes = updater.Apply(requests, inputMode);

                // ── 4. Save the model ───────────────────────────────────────
                // DA always opens models detached from central, so IsWorkshared is false.
                // SaveAsOptions alone is sufficient; WorksharingMode is not in the 2024 stubs.
                string outPath = Path.GetFullPath("result.rvt");
                doc.SaveAs(outPath, new SaveAsOptions { OverwriteExistingFile = true });
                Console.WriteLine($"[RevitParameterUpdater] Saved modified model → result.rvt");

                // ── 5. Build result ─────────────────────────────────────────
                result = new UpdateResult
                {
                    Ok      = true,
                    Summary = new ResultSummary
                    {
                        Total   = changes.Count,
                        Applied = changes.Count(c => c.Status == "applied"),
                        Skipped = changes.Count(c => c.Status == "skipped"),
                        Errors  = changes.Count(c => c.Status == "error"),
                    },
                    Changes = changes,
                };
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[RevitParameterUpdater] FATAL: {ex.Message}\n{ex.StackTrace}");
                result = new UpdateResult
                {
                    Ok    = false,
                    Error = ex.Message,
                    Summary = new ResultSummary(),
                    Changes = new List<ChangeResult>(),
                };
                // e.Succeeded stays true — DA still uploads result.json so we can diagnose
            }

            string json = JsonConvert.SerializeObject(result, Formatting.Indented);
            File.WriteAllText("result.json", json, new UTF8Encoding(false)); // no BOM
            Console.WriteLine($"[RevitParameterUpdater] Done. Applied={result.Summary.Applied} Skipped={result.Summary.Skipped} Errors={result.Summary.Errors}");
        }
    }

    // ─── Control params ───────────────────────────────────────────────────────

    internal class RunParams
    {
        [JsonProperty("inputMode")]
        public string? InputMode { get; set; }

        // Optional inline change list — lets agents pass changes directly in params.json
        // without needing a separate paramsInput file.
        // Format: [{"elementId":"...","elementName":"...","category":"...","parameter":"...","value":"..."}]
        [JsonProperty("changes")]
        public JToken? Changes { get; set; }
    }

    // ─── Data models ─────────────────────────────────────────────────────────

    internal class ChangeRequest
    {
        public string? ElementId   { get; set; }
        public string  ElementName { get; set; } = "";
        public string? Category    { get; set; }
        public string  Parameter   { get; set; } = "";
        public string  NewValue    { get; set; } = "";
    }

    public class ChangeResult
    {
        public string  ElementId   { get; set; } = "";
        public string  ElementName { get; set; } = "";
        public string  Parameter   { get; set; } = "";
        public string  OldValue    { get; set; } = "";
        public string  NewValue    { get; set; } = "";
        public string  Status      { get; set; } = "";  // "applied" | "skipped" | "error"
        public string? Reason      { get; set; }
    }

    public class ResultSummary
    {
        public int Total   { get; set; }
        public int Applied { get; set; }
        public int Skipped { get; set; }
        public int Errors  { get; set; }
    }

    public class UpdateResult
    {
        public bool         Ok      { get; set; }
        public string?      Error   { get; set; }
        public ResultSummary Summary { get; set; } = new ResultSummary();
        public List<ChangeResult> Changes { get; set; } = new List<ChangeResult>();
    }

    // ─── Input parser ─────────────────────────────────────────────────────────
    // Detects format from magic bytes:
    //   PK\x03\x04 → XLSX (ZIP container)
    //   '[' or '{' → JSON array
    //   everything else → CSV

    internal static class InputParser
    {
        internal static List<ChangeRequest> Parse(string path)
        {
            // Detect format via magic bytes
            byte[] sig = new byte[4];
            using (var fs = File.OpenRead(path))
                fs.Read(sig, 0, 4);

            if (sig[0] == 0x50 && sig[1] == 0x4B && sig[2] == 0x03 && sig[3] == 0x04)
                return ParseXlsx(path);

            string text = File.ReadAllText(path, Encoding.UTF8).TrimStart('﻿').TrimStart();
            if (text.StartsWith("[") || text.StartsWith("{"))
                return ParseJson(text);

            return ParseCsv(text);
        }

        // ── XLSX ──────────────────────────────────────────────────────────────

        private static List<ChangeRequest> ParseXlsx(string path)
        {
            var results = new List<ChangeRequest>();
            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                var ds = reader.AsDataSet(new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration { UseHeaderRow = true }
                });
                if (ds.Tables.Count == 0) return results;
                var table = ds.Tables[0];

                foreach (DataRow row in table.Rows)
                {
                    var req = RowToRequest(
                        col => GetCell(row, table, col));
                    if (req != null) results.Add(req);
                }
            }
            return results;
        }

        private static string GetCell(DataRow row, DataTable table, string colName)
        {
            foreach (DataColumn col in table.Columns)
            {
                if (string.Equals(col.ColumnName?.Trim(), colName, StringComparison.OrdinalIgnoreCase))
                    return row[col]?.ToString()?.Trim() ?? "";
            }
            return "";
        }

        // ── CSV ───────────────────────────────────────────────────────────────

        private static List<ChangeRequest> ParseCsv(string text)
        {
            var results = new List<ChangeRequest>();
            var lines = text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            if (lines.Length == 0) return results;

            // Parse header row
            var headers = SplitCsvRow(lines[0])
                .Select(h => h.Trim())
                .ToList();

            for (int i = 1; i < lines.Length; i++)
            {
                var cells = SplitCsvRow(lines[i]);
                string Get(string name)
                {
                    int idx = headers.FindIndex(h => string.Equals(h, name, StringComparison.OrdinalIgnoreCase));
                    return idx >= 0 && idx < cells.Count ? cells[idx].Trim() : "";
                }
                var req = RowToRequest(Get);
                if (req != null) results.Add(req);
            }
            return results;
        }

        private static List<string> SplitCsvRow(string line)
        {
            var fields = new List<string>();
            var sb = new StringBuilder();
            bool inQuote = false;
            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (c == '"')
                {
                    if (inQuote && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        sb.Append('"');
                        i++;
                    }
                    else
                    {
                        inQuote = !inQuote;
                    }
                }
                else if (c == ',' && !inQuote)
                {
                    fields.Add(sb.ToString());
                    sb.Clear();
                }
                else
                {
                    sb.Append(c);
                }
            }
            fields.Add(sb.ToString());
            return fields;
        }

        // ── JSON ──────────────────────────────────────────────────────────────

        internal static List<ChangeRequest> ParseJson(string text)
        {
            var results = new List<ChangeRequest>();

            // Support both array and single-object input
            text = text.Trim();
            if (text.StartsWith("{"))
                text = "[" + text + "]";

            var items = JsonConvert.DeserializeObject<List<JsonChangeRow>>(text);
            if (items == null) return results;

            foreach (var item in items)
            {
                if (string.IsNullOrWhiteSpace(item.Parameter)) continue;
                results.Add(new ChangeRequest
                {
                    ElementId   = item.ElementId?.Trim(),
                    ElementName = (item.ElementName ?? "").Trim(),
                    Category    = item.Category?.Trim(),
                    Parameter   = item.Parameter.Trim(),
                    NewValue    = (item.Value ?? item.NewValue ?? "").Trim(),
                });
            }
            return results;
        }

        // ── Shared row → ChangeRequest mapper ────────────────────────────────

        private static ChangeRequest? RowToRequest(Func<string, string> get)
        {
            string param = get("Parameter");
            if (string.IsNullOrWhiteSpace(param)) return null;

            // Accept "ElementName" or "Name" in the header
            string name = get("ElementName");
            if (string.IsNullOrWhiteSpace(name)) name = get("Name");

            return new ChangeRequest
            {
                ElementId   = NullIfEmpty(get("ElementId")),
                ElementName = name.Trim(),
                Category    = NullIfEmpty(get("Category")),
                Parameter   = param.Trim(),
                NewValue    = get("NewValue").Trim(),
            };
        }

        private static string? NullIfEmpty(string s)
            => string.IsNullOrWhiteSpace(s) ? null : s.Trim();

        // ── JSON row shape ─────────────────────────────────────────────────────
        private class JsonChangeRow
        {
            [JsonProperty("elementId")]   public string? ElementId   { get; set; }
            [JsonProperty("elementName")] public string? ElementName { get; set; }
            [JsonProperty("category")]    public string? Category    { get; set; }
            [JsonProperty("parameter")]   public string? Parameter   { get; set; }
            [JsonProperty("value")]       public string? Value       { get; set; }
            [JsonProperty("newValue")]    public string? NewValue    { get; set; }
        }
    }

    // ─── Updater ──────────────────────────────────────────────────────────────

    internal class Updater
    {
        private readonly Document _doc;

        // Name lookup map built once: elem.Name (lower) → list of (element, categoryName)
        // Only built lazily when a name-based lookup is needed.
        private Dictionary<string, List<(Element elem, string cat)>>? _nameMap;

        internal Updater(Document doc) => _doc = doc;

        internal List<ChangeResult> Apply(List<ChangeRequest> requests, string inputMode)
        {
            var results = new List<ChangeResult>();
            if (requests.Count == 0) return results;

            // Build name map if any request lacks an ElementId
            if (requests.Any(r => string.IsNullOrEmpty(r.ElementId)))
                _nameMap = BuildNameMap();

            using (var tx = new Transaction(_doc, "RevitParameterUpdater: batch update"))
            {
                tx.Start();
                foreach (var req in requests)
                    results.Add(ApplyOne(req, inputMode));
                tx.Commit();
            }
            return results;
        }

        private ChangeResult ApplyOne(ChangeRequest req, string inputMode)
        {
            var res = new ChangeResult
            {
                ElementName = req.ElementName,
                Parameter   = req.Parameter,
                NewValue    = req.NewValue,
            };

            // ── 1. Find element ───────────────────────────────────────────
            Element? elem = FindElement(req);
            if (elem == null)
            {
                res.Status = "error";
                res.Reason = $"element not found (name='{req.ElementName}', id='{req.ElementId}', category='{req.Category}')";
                Console.WriteLine($"[RevitParameterUpdater] {res.Status}: {res.Reason}");
                return res;
            }

            res.ElementId   = elem.Id.Value.ToString();
            res.ElementName = !string.IsNullOrEmpty(elem.Name) ? elem.Name : req.ElementName;

            // ── 2. Find parameter ─────────────────────────────────────────
            Parameter? param = elem.LookupParameter(req.Parameter);
            res.OldValue = param != null ? GetParamValue(param) : "";

            // ── 3. Delta check (full_file) ────────────────────────────────
            if (inputMode == "full_file" && param != null && !param.IsReadOnly)
            {
                bool same = string.Equals(res.OldValue.Trim(), req.NewValue.Trim(),
                    StringComparison.OrdinalIgnoreCase);
                if (same)
                {
                    res.Status = "skipped";
                    res.Reason = "value already matches (full_file delta: no change)";
                    return res;
                }
            }

            // ── 4a. Primary path: standard Revit parameter ────────────────
            if (param != null && !param.IsReadOnly)
            {
                if (!SetParamValue(param, req.NewValue, out string setReason))
                {
                    res.Status = "error";
                    res.Reason = setReason;
                    Console.WriteLine($"[RevitParameterUpdater] error on element {res.ElementId} param '{req.Parameter}': {setReason}");
                    return res;
                }
                res.NewValue = GetParamValue(param);
                res.Status   = "applied";
                Console.WriteLine($"[RevitParameterUpdater] applied: elem={res.ElementId} param='{req.Parameter}' '{res.OldValue}' → '{res.NewValue}'");
                return res;
            }

            // ── 4b. Fallback: compound structure (walls, roofs, floors, ceilings) ──
            // Handles thickness/layer-width parameters that are not accessible via
            // LookupParameter because they live in the type's CompoundStructure.
            if (TryApplyToCompoundStructure(elem, req.Parameter, req.NewValue,
                    out string csOld, out string csNew, out string csReason))
            {
                res.OldValue = csOld;
                res.NewValue = csNew;
                res.Status   = "applied";
                Console.WriteLine($"[RevitParameterUpdater] applied (compound): elem={res.ElementId} param='{req.Parameter}' '{csOld}' → '{csNew}'");
                return res;
            }

            // ── 4c. Fallback: FamilySymbol type parameter ─────────────────────
            // Reaches type-level parameters not directly on the instance
            // (door panel width, window frame, curtain panel thickness, etc.).
            if (TryApplyToFamilyType(elem, req.Parameter, req.NewValue,
                    out string ftOld, out string ftNew, out string ftReason))
            {
                res.OldValue = ftOld;
                res.NewValue = ftNew;
                res.Status   = "applied";
                Console.WriteLine($"[RevitParameterUpdater] applied (family type): elem={res.ElementId} param='{req.Parameter}' '{ftOld}' → '{ftNew}'");
                return res;
            }

            // All three paths failed — report the most useful error
            res.Status = "error";
            if (param == null)
            {
                var details = new List<string>();
                if (!string.IsNullOrEmpty(csReason)) details.Add($"compound: {csReason}");
                if (!string.IsNullOrEmpty(ftReason)) details.Add($"family type: {ftReason}");
                res.Reason = details.Count > 0
                    ? $"'{req.Parameter}' not found. {string.Join(" | ", details)}"
                    : $"parameter '{req.Parameter}' not found on element {res.ElementId} (tried instance, compound structure, family type)";
            }
            else
            {
                var details = new List<string>();
                if (!string.IsNullOrEmpty(csReason)) details.Add($"compound: {csReason}");
                if (!string.IsNullOrEmpty(ftReason)) details.Add($"family type: {ftReason}");
                res.Reason = $"parameter '{req.Parameter}' is read-only on element {res.ElementId}" +
                             (details.Count > 0 ? $"; {string.Join(" | ", details)}" : "");
            }
            Console.WriteLine($"[RevitParameterUpdater] {res.Status}: {res.Reason}");
            return res;
        }

        // Maps layer-function keywords (case-insensitive substring match on param name) to
        // Revit MaterialFunctionAssignment values. More-specific terms listed first so
        // "Finish 1" matches before the bare "Finish" catch-all.
        private static readonly (string keyword, MaterialFunctionAssignment func)[] s_layerFuncKeywords =
        {
            ("finish 1",    MaterialFunctionAssignment.Finish1),
            ("finish1",     MaterialFunctionAssignment.Finish1),
            ("finish 2",    MaterialFunctionAssignment.Finish2),
            ("finish2",     MaterialFunctionAssignment.Finish2),
            ("substrate 1", MaterialFunctionAssignment.Substrate1),
            ("substrate1",  MaterialFunctionAssignment.Substrate1),
            ("substrate 2", MaterialFunctionAssignment.Substrate2),
            ("substrate2",  MaterialFunctionAssignment.Substrate2),
            ("structural",  MaterialFunctionAssignment.Structure),
            ("structure",   MaterialFunctionAssignment.Structure),
            ("membrane",    MaterialFunctionAssignment.Membrane),
            ("thermal",     MaterialFunctionAssignment.ThermalOrAir),
            ("substrate",   MaterialFunctionAssignment.Substrate1),
            ("finish",      MaterialFunctionAssignment.Finish1),
            ("air",         MaterialFunctionAssignment.ThermalOrAir),
        };

        // ── Compound structure path ───────────────────────────────────────────
        // Supports: "Thickness", "Layer N Thickness", "Layer N Width" (1-indexed N),
        // and function-keyword forms: "Structural Layer Thickness",
        // "Finish Layer Thickness", "Substrate Layer Thickness", "Membrane Thickness", etc.
        // For single-layer types, any of the above targets the only layer.
        // For multi-layer types without an N or keyword, falls back to the structural layer.

        private bool TryApplyToCompoundStructure(Element elem, string paramName, string newValue,
            out string oldValue, out string newValueResult, out string reason)
        {
            oldValue = ""; newValueResult = ""; reason = "";

            if (elem is not HostObject) return false;

            var type = _doc.GetElement(elem.GetTypeId()) as HostObjAttributes;
            if (type == null) { reason = "element type is not a compound host type"; return false; }

            CompoundStructure? cs = type.GetCompoundStructure();
            if (cs == null) { reason = "element type has no compound structure"; return false; }

            int layerCount = cs.LayerCount;
            if (layerCount == 0) { reason = "compound structure has no layers"; return false; }

            // Determine target layer index
            int targetIdx = -1;
            var layerMatch = System.Text.RegularExpressions.Regex.Match(
                paramName, @"layer\s*(\d+)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);

            if (layerMatch.Success)
            {
                int n = int.Parse(layerMatch.Groups[1].Value) - 1; // 1-indexed → 0-indexed
                if (n < 0 || n >= layerCount)
                {
                    reason = $"Layer {n + 1} out of range — this compound type has {layerCount} layer(s)";
                    return false;
                }
                targetIdx = n;
            }
            else if (layerCount == 1)
            {
                targetIdx = 0;
            }
            else
            {
                // Multi-layer: check for a function keyword in the param name first
                // (e.g. "Structural Layer Thickness" → Structure, "Finish Layer" → Finish1).
                // Falls back to the structural layer when no keyword matches.
                string paramLower = paramName.ToLowerInvariant();
                MaterialFunctionAssignment? targetFunc = null;
                foreach (var pair in s_layerFuncKeywords)
                {
                    if (paramLower.Contains(pair.keyword))
                    {
                        targetFunc = pair.func;
                        break;
                    }
                }

                var funcToSearch = targetFunc ?? MaterialFunctionAssignment.Structure;
                for (int i = 0; i < layerCount; i++)
                {
                    if (cs.GetLayerFunction(i) == funcToSearch)
                    {
                        targetIdx = i;
                        break;
                    }
                }
                if (targetIdx < 0)
                {
                    var descs = new List<string>();
                    for (int i = 0; i < layerCount; i++)
                        descs.Add($"Layer {i + 1} ({cs.GetLayerFunction(i)}, {cs.GetLayerWidth(i) * 304.8:F1} mm)");
                    string hint = targetFunc.HasValue
                        ? $"No layer with function '{funcToSearch}' found."
                        : "No structural layer found — use an explicit keyword.";
                    reason = $"{hint} Use 'Layer N Thickness' (1-indexed) or a function keyword " +
                             $"('Structural', 'Finish', 'Finish 1', 'Finish 2', 'Substrate', 'Membrane', 'Thermal'). " +
                             $"Available layers: {string.Join("; ", descs)}";
                    return false;
                }
            }

            double oldWidthFeet = cs.GetLayerWidth(targetIdx);
            oldValue = $"{oldWidthFeet * 304.8:F2} mm";

            if (!TryParseLengthToFeet(newValue, out double newWidthFeet))
            {
                reason = $"Cannot parse '{newValue}' as a length. Use formats like '300 mm', '0.984 ft', or a plain number in project mm.";
                return false;
            }

            try
            {
                cs.SetLayerWidth(targetIdx, newWidthFeet);
                type.SetCompoundStructure(cs);
                newValueResult = $"{newWidthFeet * 304.8:F2} mm";
                return true;
            }
            catch (Exception ex)
            {
                reason = $"SetCompoundStructure failed: {ex.Message}";
                return false;
            }
        }

        // Parse length strings to decimal feet (Revit internal unit).
        // Supports: "300 mm", "30 cm", "0.984 ft", "0.984'", or bare number (assumed mm).
        private static bool TryParseLengthToFeet(string value, out double feet)
        {
            feet = 0;
            value = value.Trim();

            double num;
            if (value.EndsWith("mm", StringComparison.OrdinalIgnoreCase))
            {
                if (!double.TryParse(value.Substring(0, value.Length - 2).Trim(), System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out num)) return false;
                feet = num / 304.8; return true;
            }
            if (value.EndsWith("cm", StringComparison.OrdinalIgnoreCase))
            {
                if (!double.TryParse(value.Substring(0, value.Length - 2).Trim(), System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out num)) return false;
                feet = num / 30.48; return true;
            }
            if (value.EndsWith("ft", StringComparison.OrdinalIgnoreCase))
            {
                if (!double.TryParse(value.Substring(0, value.Length - 2).Trim(), System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out num)) return false;
                feet = num; return true;
            }
            if (value.EndsWith("'"))
            {
                if (!double.TryParse(value.Substring(0, value.Length - 1).Trim(), System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out num)) return false;
                feet = num; return true;
            }
            // Bare number — assume project units are mm (most Revit projects outside the US)
            if (double.TryParse(value, System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out num))
            {
                feet = num / 304.8; return true;
            }
            return false;
        }

        // ── FamilySymbol type parameter path ─────────────────────────────────
        // Edits a type-level parameter on the element's FamilySymbol.
        // This reaches parameters that live on the family type rather than the instance —
        // e.g. door panel width, window height constraint, curtain panel thickness,
        // fixture model depth. Editing the type affects ALL instances of that type,
        // so the result is written back to the type in the model.

        private bool TryApplyToFamilyType(Element elem, string paramName, string newValue,
            out string oldValue, out string newValueResult, out string reason)
        {
            oldValue = ""; newValueResult = ""; reason = "";

            var fi = elem as FamilyInstance;
            if (fi == null) return false;

            FamilySymbol? symbol = fi.Symbol;
            if (symbol == null) { reason = "FamilyInstance has no associated FamilySymbol"; return false; }

            Parameter? param = symbol.LookupParameter(paramName);
            if (param == null)
            {
                reason = $"'{paramName}' not found on FamilySymbol '{symbol.Name}' (family: '{symbol.FamilyName}')";
                return false;
            }
            if (param.IsReadOnly)
            {
                reason = $"'{paramName}' is read-only on FamilySymbol '{symbol.Name}'";
                return false;
            }

            oldValue = GetParamValue(param);
            if (!SetParamValue(param, newValue, out reason)) return false;
            newValueResult = GetParamValue(param);
            return true;
        }

        // ── Element lookup ────────────────────────────────────────────────────

        private Element? FindElement(ChangeRequest req)
        {
            // By ElementId (fastest, most reliable)
            if (!string.IsNullOrEmpty(req.ElementId) && long.TryParse(req.ElementId, out long id))
                return _doc.GetElement(new ElementId(id));

            // By name (+ optional category filter) using pre-built map
            string nameLower = req.ElementName.ToLowerInvariant();
            if (_nameMap == null || !_nameMap.TryGetValue(nameLower, out var candidates))
                return null;

            if (string.IsNullOrEmpty(req.Category))
                return candidates.Count > 0 ? candidates[0].elem : null;

            foreach (var (elem, cat) in candidates)
            {
                if (string.Equals(cat, req.Category, StringComparison.OrdinalIgnoreCase))
                    return elem;
            }
            return null;
        }

        private Dictionary<string, List<(Element, string)>> BuildNameMap()
        {
            Console.WriteLine("[RevitParameterUpdater] Building element name map...");
            var map = new Dictionary<string, List<(Element, string)>>(StringComparer.OrdinalIgnoreCase);
            var collector = new FilteredElementCollector(_doc).WhereElementIsNotElementType();
            foreach (Element elem in collector)
            {
                string name = elem.Name ?? "";
                if (string.IsNullOrEmpty(name)) continue;
                string key = name.ToLowerInvariant();
                if (!map.TryGetValue(key, out var list))
                {
                    list = new List<(Element, string)>();
                    map[key] = list;
                }
                list.Add((elem, elem.Category?.Name ?? ""));
            }
            Console.WriteLine($"[RevitParameterUpdater] Name map built — {map.Count} unique names.");
            return map;
        }

        // ── Parameter value helpers ───────────────────────────────────────────

        private static string GetParamValue(Parameter p)
        {
            try
            {
                if (!p.HasValue) return "";
                return p.StorageType switch
                {
                    StorageType.String    => p.AsString() ?? "",
                    StorageType.Double    => p.AsValueString() ?? p.AsDouble().ToString("G10"),
                    StorageType.Integer   => p.AsValueString() ?? p.AsInteger().ToString(),
                    StorageType.ElementId => p.AsElementId() == ElementId.InvalidElementId
                                              ? "" : (p.AsValueString() ?? p.AsElementId().Value.ToString()),
                    _                    => "",
                };
            }
            catch { return ""; }
        }

        // Instance method (not static) so it can call ResolveElementIdByName via _doc.
        private bool SetParamValue(Parameter p, string value, out string reason)
        {
            reason = "";
            if (p.IsReadOnly) { reason = "parameter is read-only"; return false; }

            try
            {
                switch (p.StorageType)
                {
                    case StorageType.String:
                        p.Set(value);
                        return true;

                    case StorageType.Double:
                        // SetValueString respects project units (e.g. "4200 mm", "13'-9\"").
                        // Fall back to raw double parse if it refuses.
                        if (p.SetValueString(value)) return true;
                        if (double.TryParse(value,
                            System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture,
                            out double d))
                        { p.Set(d); return true; }
                        reason = $"cannot parse '{value}' as a unit string or number for this parameter";
                        return false;

                    case StorageType.Integer:
                        // SetValueString handles Yes/No, enum labels, etc.
                        if (p.SetValueString(value)) return true;
                        if (int.TryParse(value, out int i)) { p.Set(i); return true; }
                        reason = $"cannot parse '{value}' as an integer or enum for this parameter";
                        return false;

                    case StorageType.ElementId:
                        // Try numeric first (viewer sometimes sends raw ElementId integer).
                        if (long.TryParse(value, out long eid))
                        { p.Set(new ElementId(eid)); return true; }
                        // Resolve by name: Material → Level → Phase → broad fallback.
                        // Covers cases where the viewer sends a display name
                        // (e.g. "Wood - Cherry" for a material, "Level 1" for a base level).
                        var resolved = ResolveElementIdByName(value);
                        if (resolved != null) { p.Set(resolved.Id); return true; }
                        reason = $"cannot resolve '{value}' as an ElementId — expected a numeric ID " +
                                 $"or a valid element name (Material, Level, Phase, or other element in this model)";
                        return false;

                    default:
                        reason = $"unsupported StorageType: {p.StorageType}";
                        return false;
                }
            }
            catch (Exception ex)
            {
                reason = ex.Message;
                return false;
            }
        }

        // Resolve an element by display name for ElementId parameters.
        // Searches high-frequency types first (Material, Level, Phase) then falls back
        // to a broad sweep across all instances then element types.
        private Element? ResolveElementIdByName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return null;

            // Priority types — cover the vast majority of named ElementId parameters
            Type[] priorityTypes = { typeof(Material), typeof(Level), typeof(Phase) };
            foreach (var t in priorityTypes)
            {
                var hit = new FilteredElementCollector(_doc)
                    .OfClass(t)
                    .FirstOrDefault(e => e.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
                if (hit != null) return hit;
            }

            // Broad fallback: instances first (rooms, grids, etc.), then element types
            var inst = new FilteredElementCollector(_doc)
                .WhereElementIsNotElementType()
                .FirstOrDefault(e => !string.IsNullOrEmpty(e.Name) &&
                                     e.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
            if (inst != null) return inst;

            return new FilteredElementCollector(_doc)
                .WhereElementIsElementType()
                .FirstOrDefault(e => !string.IsNullOrEmpty(e.Name) &&
                                     e.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
        }
    }
}
