// ════════════════════════════════════════════════════════════════════════════
//  AutoCADTitleBlockUpdater — APS Design Automation (AutoCAD 2024, +24_3, net48)
//
//  Reads and updates title-block attributes (Revision, Date, Sheet Number, Project,
//  Drawn By, ...) across paper-space layouts, then saves the modified drawing.
//  WRITE-BACK bundle modeled on RevitParameterUpdater (open ForWrite, apply in a
//  transaction, Commit, save the file out).
//
//  SINGLE VERB, TWO MODES (DA is headless — no mid-run prompts). The bundle self-
//  selects the mode from its inputs on every call:
//    MODE A — EXTRACT (no "changes" payload, or "mode":"extract"): discover every
//             title-block insert across in-scope layouts and emit its schema. No write.
//             result.json carries needsInput:true so the caller collects values and reruns.
//    MODE B — UPDATE (a "changes" payload IS present): map provided values onto
//             extracted attribute tags (case-insensitive), open matching
//             AttributeReferences ForWrite, set TextString, AdjustAlignment, skip
//             constant attributes, Commit, SaveAs result.dwg.
//
//  INPUTS (placed in working dir by DA before accoreconsole runs):
//    input.dwg     — the drawing (opened via accoreconsole /i, becomes active document)
//    params.json   — OPTIONAL control JSON, chat/inline path:
//                      { "mode":"extract|update", "titleBlockName":"*",
//                        "layoutScope":"all" | ["A-101","A-102"],
//                        "changes":{ "REVISION":"C", "DATE":"2026-07-07" } }
//    changes.dat   — OPTIONAL uploaded file (webapp JSON same shape, or CSV). Detected
//                    by magic byte then extension. Inline params.json.changes preferred.
//
//  OUTPUTS:
//    result.json   — schema (Mode A) or update summary (Mode B); also the smoke gate file
//    result.dwg    — modified drawing (Mode B only)
//
//  net48 TRAPS (all guarded here):
//    - No System.Text.Json — use Newtonsoft (ships in Contents/).
//    - No Dictionary.GetValueOrDefault — use TryGetValue.
//    - Write result.json with new UTF8Encoding(false) — Encoding.UTF8 emits a BOM.
//
//  VERB APSTITLEBLOCK is APS-prefixed to dodge blocked/built-in AutoCAD commands.
// ════════════════════════════════════════════════════════════════════════════
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Autodesk.AutoCAD.ApplicationServices.Core;   // Application (core console)
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

[assembly: CommandClass(typeof(AutoCADTitleBlockUpdater.Commands))]
[assembly: ExtensionApplication(null)]

namespace AutoCADTitleBlockUpdater
{
    public class Commands
    {
        // Verb must match <Command Global="APSTITLEBLOCK"> in PackageContents.xml and
        // the .scr the activity feeds accoreconsole. APS-prefixed to dodge built-ins.
        [CommandMethod("APSTITLEBLOCK", CommandFlags.Modal)]
        public static void RunTitleBlock()
        {
            string mode = "extract";
            try
            {
                var wdFiles = Directory.GetFiles(Directory.GetCurrentDirectory());
                Console.WriteLine($"[APSTITLEBLOCK] Working dir files: {string.Join(", ", wdFiles.Select(Path.GetFileName))}");

                // In accoreconsole the input DWG opened via /i is the active document.
                var doc = Application.DocumentManager.MdiActiveDocument;
                var db  = doc?.Database ?? HostApplicationServices.WorkingDatabase;

                // ── 1. Read control params + normalize the changes payload ──────
                var input = InputReader.Read();
                Console.WriteLine($"[APSTITLEBLOCK] titleBlockName='{input.TitleBlockName}' " +
                    $"layoutScope='{input.LayoutScopeText}' changes={input.Changes.Count} " +
                    $"source={input.Source ?? "(none)"} forceExtract={input.ForceExtract}");

                // Mode self-selection: a changes payload present (and not forced to extract) → UPDATE.
                bool update = !input.ForceExtract && input.Changes.Count > 0;
                mode = update ? "update" : "extract";
                Console.WriteLine($"[APSTITLEBLOCK] mode={mode}");

                object result;
                if (update)
                {
                    result = TitleBlockEngine.Update(db, input);
                    // SaveAs current DWG version — the write-back output.
                    string outPath = Path.GetFullPath("result.dwg");
                    db.SaveAs(outPath, DwgVersion.Current);
                    Console.WriteLine($"[APSTITLEBLOCK] Saved modified drawing → result.dwg");
                }
                else
                {
                    result = TitleBlockEngine.Extract(db, input);
                }

                WriteJson("result.json", result);
                Console.WriteLine("[APSTITLEBLOCK] Done.");
            }
            catch (System.Exception ex)
            {
                Console.WriteLine($"[APSTITLEBLOCK] ERROR: {ex.Message}\n{ex.StackTrace}");
                WriteJson("result.json", new Dictionary<string, object>
                {
                    ["ok"]    = false,
                    ["mode"]  = mode,
                    ["error"] = ex.Message,
                    ["stack"] = ex.StackTrace ?? "",
                });
            }
        }

        // UTF-8 WITHOUT BOM — required for downstream JSON.parse to succeed.
        private static void WriteJson(string path, object payload)
        {
            File.WriteAllText(path,
                JsonConvert.SerializeObject(payload, Formatting.Indented),
                new UTF8Encoding(false));
        }
    }

    // ─── Normalized run input ───────────────────────────────────────────────────
    internal class RunInput
    {
        public bool ForceExtract { get; set; }              // "mode":"extract" was supplied
        public string TitleBlockName { get; set; } = "*";   // "*" = any block carrying the tags
        public HashSet<string>? LayoutScope { get; set; }   // null = all layouts
        public string LayoutScopeText { get; set; } = "all";
        // Case-insensitive tag → value.
        public Dictionary<string, string> Changes { get; } =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        public string? Source { get; set; }                 // "chat" | "json" | "csv"

        public bool LayoutInScope(string layoutName) =>
            LayoutScope == null || LayoutScope.Contains(layoutName);

        public bool BlockMatches(string blockName) =>
            TitleBlockName == "*" || string.IsNullOrEmpty(TitleBlockName) ||
            string.Equals(blockName, TitleBlockName, StringComparison.OrdinalIgnoreCase);
    }

    // ─── Input reader ───────────────────────────────────────────────────────────
    // Sources, all normalized to RunInput. Precedence: the explicit inline config in
    // params.json ("chat") is preferred over an uploaded changes.dat file.
    //   1) CHAT  : params.json inline { titleBlockName, layoutScope, changes:{...} }
    //   2) JSON  : changes.dat, magic byte '{' — same shape or a bare changes object
    //   3) CSV   : changes.dat — two-column (tag,value) or header-of-tags + one value row
    internal static class InputReader
    {
        internal static RunInput Read()
        {
            var input = new RunInput();

            // ── 1. params.json (control + inline changes) ───────────────────────
            JObject? control = null;
            if (File.Exists("params.json"))
            {
                try { control = JObject.Parse(File.ReadAllText("params.json", Encoding.UTF8).TrimStart('﻿')); }
                catch (System.Exception ex)
                {
                    Console.WriteLine($"[APSTITLEBLOCK] WARNING: could not parse params.json — {ex.Message}");
                }
            }

            if (control != null)
            {
                ApplyScope(input, control);
                var mode = (string?)control["mode"];
                if (!string.IsNullOrWhiteSpace(mode) &&
                    mode!.Trim().Equals("extract", StringComparison.OrdinalIgnoreCase))
                    input.ForceExtract = true;

                if (control["changes"] is JObject inlineChanges && inlineChanges.Count > 0)
                {
                    MergeChanges(input.Changes, inlineChanges);
                    input.Source = "chat";
                }
            }

            // ── 2/3. changes.dat file — only when no inline changes were supplied ─
            if (input.Changes.Count == 0 && File.Exists("changes.dat"))
            {
                var fi = new FileInfo("changes.dat");
                if (fi.Length > 0)
                {
                    byte[] sig = new byte[4];
                    using (var fs = File.OpenRead("changes.dat")) fs.Read(sig, 0, Math.Min(4, (int)fi.Length));
                    string text = File.ReadAllText("changes.dat", Encoding.UTF8).TrimStart('﻿').TrimStart();

                    if (text.StartsWith("{") || text.StartsWith("["))
                    {
                        ParseJsonFile(input, text);
                        input.Source = "json";
                    }
                    else
                    {
                        ParseCsv(input, text);
                        input.Source = "csv";
                    }
                }
            }

            return input;
        }

        private static void ApplyScope(RunInput input, JObject o)
        {
            var tbn = (string?)o["titleBlockName"];
            if (!string.IsNullOrWhiteSpace(tbn)) input.TitleBlockName = tbn!.Trim();

            var scope = o["layoutScope"];
            if (scope != null) SetLayoutScope(input, scope);
        }

        private static void SetLayoutScope(RunInput input, JToken scope)
        {
            if (scope.Type == JTokenType.Array)
            {
                var names = scope.Select(t => (string?)t)
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .Select(s => s!.Trim());
                var set = new HashSet<string>(names, StringComparer.OrdinalIgnoreCase);
                if (set.Count > 0)
                {
                    input.LayoutScope = set;
                    input.LayoutScopeText = string.Join(",", set);
                }
            }
            else
            {
                var s = (string?)scope;
                if (!string.IsNullOrWhiteSpace(s) &&
                    !s!.Trim().Equals("all", StringComparison.OrdinalIgnoreCase))
                {
                    input.LayoutScope = new HashSet<string>(
                        new[] { s.Trim() }, StringComparer.OrdinalIgnoreCase);
                    input.LayoutScopeText = s.Trim();
                }
            }
        }

        // A webapp JSON may carry the full config shape (with a "changes" object and
        // optional scope), or be a bare { tag: value } object.
        private static void ParseJsonFile(RunInput input, string text)
        {
            var tok = JToken.Parse(text);
            if (tok is JObject obj)
            {
                if (obj["changes"] is JObject changes)
                {
                    ApplyScope(input, obj);
                    MergeChanges(input.Changes, changes);
                }
                else
                {
                    MergeChanges(input.Changes, obj);
                }
            }
        }

        // CSV: two-column (tag,value per row) OR header-row-of-tags + one value row.
        // Optional titleBlock / layout columns scope the run (first non-empty wins as
        // the run-wide scope; per-row scoping is not modeled — see registry contract).
        private static void ParseCsv(RunInput input, string text)
        {
            var lines = text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            if (lines.Length == 0) return;

            var row0 = SplitCsvRow(lines[0]).Select(c => c.Trim()).ToList();

            bool tagValueHeader = row0.Count == 2 &&
                row0[0].Equals("tag", StringComparison.OrdinalIgnoreCase) &&
                row0[1].Equals("value", StringComparison.OrdinalIgnoreCase);

            // Header-of-tags form only when the first row clearly names scope/tag columns
            // across >2 columns; otherwise treat two-column data as tag,value rows.
            bool headerOfTags = row0.Count > 2 && lines.Length >= 2;

            if (tagValueHeader)
            {
                for (int i = 1; i < lines.Length; i++)
                    AddTwoCol(input, SplitCsvRow(lines[i]));
                return;
            }

            if (!headerOfTags)
            {
                // Two-column tag,value rows (no header, or arbitrary header we skip if non-data).
                foreach (var line in lines)
                    AddTwoCol(input, SplitCsvRow(line));
                return;
            }

            // Header-row-of-tags + one value row (with optional titleBlock/layout columns).
            var valueRow = SplitCsvRow(lines[1]);
            for (int c = 0; c < row0.Count && c < valueRow.Count; c++)
            {
                string tag = row0[c];
                string val = valueRow[c].Trim();
                if (string.IsNullOrWhiteSpace(tag)) continue;

                if (tag.Equals("titleBlock", StringComparison.OrdinalIgnoreCase) ||
                    tag.Equals("titleBlockName", StringComparison.OrdinalIgnoreCase) ||
                    tag.Equals("block", StringComparison.OrdinalIgnoreCase))
                {
                    if (!string.IsNullOrWhiteSpace(val)) input.TitleBlockName = val;
                    continue;
                }
                if (tag.Equals("layout", StringComparison.OrdinalIgnoreCase) ||
                    tag.Equals("layoutScope", StringComparison.OrdinalIgnoreCase))
                {
                    if (!string.IsNullOrWhiteSpace(val)) SetLayoutScope(input, new JValue(val));
                    continue;
                }
                input.Changes[tag] = val;
            }
        }

        private static void AddTwoCol(RunInput input, List<string> cells)
        {
            if (cells.Count < 2 || string.IsNullOrWhiteSpace(cells[0])) return;
            string tag = cells[0].Trim();
            // Skip a stray header row like "tag,value".
            if (tag.Equals("tag", StringComparison.OrdinalIgnoreCase) &&
                cells[1].Trim().Equals("value", StringComparison.OrdinalIgnoreCase)) return;
            input.Changes[tag] = cells[1].Trim();
        }

        private static void MergeChanges(Dictionary<string, string> into, JObject src)
        {
            foreach (var p in src.Properties())
            {
                if (string.IsNullOrWhiteSpace(p.Name)) continue;
                // JValue.ToString() yields the raw scalar text (string content, number,
                // or "True"/"False"); Null → empty string.
                into[p.Name.Trim()] = p.Value == null || p.Value.Type == JTokenType.Null
                    ? "" : p.Value.ToString();
            }
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
                    if (inQuote && i + 1 < line.Length && line[i + 1] == '"') { sb.Append('"'); i++; }
                    else inQuote = !inQuote;
                }
                else if (c == ',' && !inQuote) { fields.Add(sb.ToString()); sb.Clear(); }
                else sb.Append(c);
            }
            fields.Add(sb.ToString());
            return fields;
        }
    }

    // ─── Title-block discovery + update engine ──────────────────────────────────
    internal static class TitleBlockEngine
    {
        // A discovered attribute reference: everything needed for both schema and write.
        private class FoundAttr
        {
            public string Layout = "";
            public string Block = "";
            public ObjectId AttrId;         // Null for constant attrs (definition-only)
            public string Tag = "";
            public string Prompt = "";
            public string CurrentValue = "";
            public bool Constant;
        }

        // ── Shared discovery walk: layouts → BlockReference inserts → attributes ──
        // Used by BOTH modes so extract and update see exactly the same title blocks.
        private static List<FoundAttr> Discover(Transaction tr, Database db, RunInput input)
        {
            var found = new List<FoundAttr>();
            var defCache = new Dictionary<ObjectId, Dictionary<string, DefAttr>>();

            var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

            // Order layouts by tab order for stable, human-friendly output.
            var layouts = new List<KeyValuePair<string, Layout>>();
            foreach (DBDictionaryEntry entry in layoutDict)
            {
                var lo = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                layouts.Add(new KeyValuePair<string, Layout>(entry.Key, lo));
            }

            foreach (var kvp in layouts.OrderBy(l => l.Value.TabOrder))
            {
                string name = kvp.Key;
                if (!input.LayoutInScope(name)) continue;

                var btr = (BlockTableRecord)tr.GetObject(kvp.Value.BlockTableRecordId, OpenMode.ForRead);
                foreach (ObjectId entId in btr)
                {
                    if (entId.ObjectClass.DxfName != "INSERT") continue;
                    var br = tr.GetObject(entId, OpenMode.ForRead) as BlockReference;
                    if (br == null) continue;

                    // Effective block name (dynamic-block aware).
                    string blockName = EffectiveBlockName(tr, br);
                    if (!input.BlockMatches(blockName)) continue;

                    // Build tag → (prompt, constant) from the block definition (cached).
                    ObjectId defBtrId = br.BlockTableRecord;
                    if (!defCache.TryGetValue(defBtrId, out var defInfo))
                    {
                        defInfo = ReadDefinitionAttrs(tr, defBtrId);
                        defCache[defBtrId] = defInfo;
                    }

                    if (br.AttributeCollection.Count == 0 && defInfo.Count == 0) continue;

                    // Variable (editable) attributes live on the insert's AttributeCollection.
                    var seenTags = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach (ObjectId arId in br.AttributeCollection)
                    {
                        var ar = tr.GetObject(arId, OpenMode.ForRead) as AttributeReference;
                        if (ar == null) continue;
                        string tag = ar.Tag ?? "";
                        seenTags.Add(tag);
                        string prompt = defInfo.TryGetValue(tag, out var d) && !string.IsNullOrEmpty(d.Prompt)
                            ? d.Prompt : tag;
                        found.Add(new FoundAttr
                        {
                            Layout = name, Block = blockName, AttrId = arId,
                            Tag = tag, Prompt = prompt,
                            CurrentValue = ar.TextString ?? "",
                            Constant = ar.IsConstant,
                        });
                    }

                    // Constant attributes are baked into the definition (not in the
                    // AttributeCollection). Report them so the caller sees the full schema.
                    foreach (var kv in defInfo)
                    {
                        if (!kv.Value.Constant || seenTags.Contains(kv.Key)) continue;
                        found.Add(new FoundAttr
                        {
                            Layout = name, Block = blockName, AttrId = ObjectId.Null,
                            Tag = kv.Key,
                            Prompt = string.IsNullOrEmpty(kv.Value.Prompt) ? kv.Key : kv.Value.Prompt,
                            CurrentValue = "", Constant = true,
                        });
                    }
                }
            }
            return found;
        }

        private static string EffectiveBlockName(Transaction tr, BlockReference br)
        {
            try
            {
                if (br.IsDynamicBlock)
                {
                    var dyn = (BlockTableRecord)tr.GetObject(br.DynamicBlockTableRecord, OpenMode.ForRead);
                    return dyn.Name;
                }
            }
            catch { /* fall through to br.Name */ }
            return br.Name;
        }

        private struct DefAttr
        {
            public string Prompt;
            public bool Constant;
        }

        private static Dictionary<string, DefAttr> ReadDefinitionAttrs(Transaction tr, ObjectId btrId)
        {
            var map = new Dictionary<string, DefAttr>(StringComparer.OrdinalIgnoreCase);
            var btr = tr.GetObject(btrId, OpenMode.ForRead) as BlockTableRecord;
            if (btr == null) return map;
            foreach (ObjectId id in btr)
            {
                if (id.ObjectClass.DxfName != "ATTDEF") continue;
                var ad = tr.GetObject(id, OpenMode.ForRead) as AttributeDefinition;
                if (ad == null || string.IsNullOrEmpty(ad.Tag)) continue;
                map[ad.Tag] = new DefAttr { Prompt = ad.Prompt ?? "", Constant = ad.Constant };
            }
            return map;
        }

        // ── MODE A — EXTRACT ──────────────────────────────────────────────────────
        internal static object Extract(Database db, RunInput input)
        {
            var titleBlocks = new List<object>();
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var found = Discover(tr, db, input);
                // Group by (layout, block) preserving discovery order.
                foreach (var grp in found
                    .GroupBy(f => f.Layout + " " + f.Block))
                {
                    var first = grp.First();
                    var attributes = grp.Select(a => (object)new Dictionary<string, object>
                    {
                        ["tag"]          = a.Tag,
                        ["prompt"]       = a.Prompt,
                        ["currentValue"] = a.CurrentValue,
                        ["constant"]     = a.Constant,
                    }).ToList();

                    titleBlocks.Add(new Dictionary<string, object>
                    {
                        ["block"]      = first.Block,
                        ["layout"]     = first.Layout,
                        ["attributes"] = attributes,
                    });
                }
                tr.Commit();
            }

            Console.WriteLine($"[APSTITLEBLOCK] Extract found {titleBlocks.Count} title block(s).");
            return new Dictionary<string, object>
            {
                ["ok"]          = true,
                ["mode"]        = "extract",
                ["needsInput"]  = true,
                ["rerunVerb"]   = "APSTITLEBLOCK",
                ["titleBlocks"] = titleBlocks,
            };
        }

        // ── MODE B — UPDATE ─────────────────────────────────────────────────────
        internal static object Update(Database db, RunInput input)
        {
            var changesApplied = new List<object>();
            var matchedKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var found = Discover(tr, db, input);
                foreach (var attr in found)
                {
                    if (attr.Constant || attr.AttrId.IsNull) continue;   // skip constant attributes
                    if (!input.Changes.TryGetValue(attr.Tag, out var newValue)) continue;

                    matchedKeys.Add(attr.Tag);
                    var ar = tr.GetObject(attr.AttrId, OpenMode.ForWrite) as AttributeReference;
                    if (ar == null) continue;
                    if (ar.IsConstant) continue;                          // belt-and-suspenders

                    string oldValue = ar.TextString ?? "";
                    ar.TextString = newValue;
                    ar.AdjustAlignment(db);

                    changesApplied.Add(new Dictionary<string, object>
                    {
                        ["layout"]   = attr.Layout,
                        ["block"]    = attr.Block,
                        ["tag"]      = attr.Tag,
                        ["oldValue"] = oldValue,
                        ["newValue"] = newValue,
                    });
                    Console.WriteLine($"[APSTITLEBLOCK] applied: layout='{attr.Layout}' block='{attr.Block}' " +
                        $"tag='{attr.Tag}' '{oldValue}' → '{newValue}'");
                }
                tr.Commit();
            }

            var unmapped = input.Changes.Keys
                .Where(k => !matchedKeys.Contains(k))
                .OrderBy(k => k, StringComparer.OrdinalIgnoreCase)
                .Cast<object>().ToList();

            Console.WriteLine($"[APSTITLEBLOCK] Update applied {changesApplied.Count} change(s); " +
                $"unmapped keys: {unmapped.Count}");

            return new Dictionary<string, object>
            {
                ["ok"]             = true,
                ["mode"]           = "update",
                ["updatedCount"]   = changesApplied.Count,
                ["source"]         = input.Source ?? "chat",
                ["changesApplied"] = changesApplied,
                ["unmapped"]       = unmapped,
            };
        }
    }
}
