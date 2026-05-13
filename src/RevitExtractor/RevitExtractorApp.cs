using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Newtonsoft.Json;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using DesignAutomationFramework;

namespace RevitExtractor
{
    // ═════════════════════════════════════════════════════════════════════════
    // ENTRY POINT
    // ═════════════════════════════════════════════════════════════════════════

    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    public class RevitExtractorApp : IExternalDBApplication
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
                var doc = e.DesignAutomationData.RevitDoc
                    ?? throw new InvalidOperationException("Revit document could not be opened.");

                Console.WriteLine($"[RevitExtractor] Processing: {doc.Title}");

                var extractor = new Extractor(doc);
                var report    = extractor.BuildReport();

                Writer.WriteJson(report);
                Writer.WriteCsv(report);

                Console.WriteLine("[RevitExtractor] Done — result.json + result.csv written.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[RevitExtractor] ERROR: {ex.Message}\n{ex.StackTrace}");
                e.Succeeded = false;
            }
        }
    }

    // ═════════════════════════════════════════════════════════════════════════
    // EXTRACTOR  — all data-gathering logic lives here
    // ═════════════════════════════════════════════════════════════════════════

    internal class Extractor
    {
        private readonly Document _doc;

        // Cache type parameters by TypeId so we only read each type once.
        // Large models can have hundreds of identical door/window types.
        private readonly Dictionary<ElementId, Dictionary<string, string>> _typeCache
            = new Dictionary<ElementId, Dictionary<string, string>>();

        // Categories to walk.  Add or remove rows here to change scope.
        private static readonly (string Label, BuiltInCategory BIC)[] TARGET_CATEGORIES =
        {
            ("Walls",         BuiltInCategory.OST_Walls),
            ("Floors",        BuiltInCategory.OST_Floors),
            ("Ceilings",      BuiltInCategory.OST_Ceilings),
            ("Roofs",         BuiltInCategory.OST_Roofs),
            ("Doors",         BuiltInCategory.OST_Doors),
            ("Windows",       BuiltInCategory.OST_Windows),
            ("Rooms",         BuiltInCategory.OST_Rooms),
            ("Stairs",        BuiltInCategory.OST_Stairs),
            ("Columns",       BuiltInCategory.OST_Columns),
            ("StructColumns", BuiltInCategory.OST_StructuralColumns),
            ("StructFraming", BuiltInCategory.OST_StructuralFraming),
            ("Furniture",     BuiltInCategory.OST_Furniture),
            ("Plumbing",      BuiltInCategory.OST_PlumbingFixtures),
            ("MechEquip",     BuiltInCategory.OST_MechanicalEquipment),
            ("ElecEquip",     BuiltInCategory.OST_ElectricalEquipment),
            ("Grids",         BuiltInCategory.OST_Grids),
        };

        internal Extractor(Document doc) => _doc = doc;

        // ── Top-level report ──────────────────────────────────────────────

        internal ModelReport BuildReport()
        {
            Console.WriteLine("[RevitExtractor] Extracting project info + levels...");
            var report = new ModelReport
            {
                ExtractedAt  = DateTime.UtcNow.ToString("o"),
                ProjectInfo  = GetProjectInfo(),
                ElementCounts = GetElementCounts(),
                Levels       = GetLevels(),
                Warnings     = GetWarnings(),
                Categories   = new Dictionary<string, List<ElementData>>(),
            };

            foreach (var (label, bic) in TARGET_CATEGORIES)
            {
                Console.WriteLine($"[RevitExtractor] Extracting {label}...");
                var elements = ExtractCategory(bic);
                if (elements.Count > 0)
                    report.Categories[label] = elements;
            }

            return report;
        }

        // ── Project info ──────────────────────────────────────────────────

        private ProjectInfoData GetProjectInfo()
        {
            var pi = _doc.ProjectInformation;
            return new ProjectInfoData
            {
                Name             = pi.Name,
                Number           = pi.Number,
                ClientName       = pi.ClientName,
                Address          = pi.Address,
                BuildingName     = pi.BuildingName,
                Status           = pi.Status,
                OrganizationName = pi.OrganizationName,
            };
        }

        // ── Element counts ────────────────────────────────────────────────

        private Dictionary<string, int> GetElementCounts()
        {
            var counts = new Dictionary<string, int>();
            foreach (var (label, bic) in TARGET_CATEGORIES)
            {
                counts[label] = new FilteredElementCollector(_doc)
                    .OfCategory(bic)
                    .WhereElementIsNotElementType()
                    .GetElementCount();
            }
            // Also count Sheets + Views separately (not in TARGET_CATEGORIES)
            counts["Sheets"] = new FilteredElementCollector(_doc)
                .OfCategory(BuiltInCategory.OST_Sheets)
                .WhereElementIsNotElementType().GetElementCount();
            counts["Views"] = new FilteredElementCollector(_doc)
                .OfCategory(BuiltInCategory.OST_Views)
                .WhereElementIsNotElementType().GetElementCount();
            return counts;
        }

        // ── Levels ────────────────────────────────────────────────────────

        private List<LevelData> GetLevels()
        {
            return new FilteredElementCollector(_doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .OrderBy(l => l.Elevation)
                .Select(l => new LevelData
                {
                    Name      = l.Name,
                    ElevationM = Math.Round(l.Elevation * 0.3048, 3),
                })
                .ToList();
        }

        // ── Warnings ──────────────────────────────────────────────────────

        private List<string> GetWarnings()
            => _doc.GetWarnings()
                   .Take(50)
                   .Select(w => w.GetDescriptionText())
                   .ToList();

        // ═════════════════════════════════════════════════════════════════
        // PER-ELEMENT PARAMETER EXTRACTION
        // ═════════════════════════════════════════════════════════════════

        private List<ElementData> ExtractCategory(BuiltInCategory bic)
        {
            var results = new List<ElementData>();

            var elements = new FilteredElementCollector(_doc)
                .OfCategory(bic)
                .WhereElementIsNotElementType()
                .ToElements();

            foreach (var elem in elements)
            {
                try
                {
                    results.Add(new ElementData
                    {
                        ElementId          = elem.Id.IntegerValue.ToString(),
                        FamilyType         = GetFamilyTypeName(elem),
                        InstanceParameters = ReadParameters(elem),
                        TypeParameters     = ReadTypeParameters(elem),
                    });
                }
                catch (Exception ex)
                {
                    // Log and continue — one bad element shouldn't abort the whole category
                    Console.WriteLine($"[RevitExtractor] Skipped element {elem.Id}: {ex.Message}");
                }
            }

            return results;
        }

        // ── Instance parameters ───────────────────────────────────────────

        /// <summary>
        /// Reads every instance parameter on an element.
        /// Duplicate names: first occurrence wins (Revit can surface the same
        /// built-in under multiple internal IDs in rare cases).
        /// </summary>
        private static Dictionary<string, string> ReadParameters(Element elem)
        {
            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (Parameter p in elem.Parameters)
            {
                var name = SafeParamName(p);
                if (name == null || dict.ContainsKey(name)) continue;
                dict[name] = GetParamValue(p);
            }

            return dict;
        }

        // ── Type parameters (cached) ──────────────────────────────────────

        /// <summary>
        /// Reads every parameter on the element's type.
        /// Results are cached by TypeId — identical type, one read.
        /// </summary>
        private Dictionary<string, string> ReadTypeParameters(Element elem)
        {
            var typeId = elem.GetTypeId();
            if (typeId == null || typeId == ElementId.InvalidElementId)
                return new Dictionary<string, string>();

            if (_typeCache.TryGetValue(typeId, out var cached))
                return cached;

            var typeElem = _doc.GetElement(typeId);
            var dict     = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            if (typeElem != null)
            {
                foreach (Parameter p in typeElem.Parameters)
                {
                    var name = SafeParamName(p);
                    if (name == null || dict.ContainsKey(name)) continue;
                    dict[name] = GetParamValue(p);
                }
            }

            _typeCache[typeId] = dict;
            return dict;
        }

        // ── Helpers ───────────────────────────────────────────────────────

        private static string? SafeParamName(Parameter p)
        {
            try
            {
                var name = p?.Definition?.Name;
                return string.IsNullOrWhiteSpace(name) ? null : name.Trim();
            }
            catch { return null; }
        }

        /// <summary>
        /// Reads a parameter value as a display string.
        ///
        /// StorageType mapping:
        ///   String    → AsString()
        ///   Double    → AsValueString()  ← formatted with units (e.g. "4200 mm")
        ///   Integer   → AsValueString()  ← enum labels where applicable
        ///   ElementId → AsValueString()  ← resolved name where possible
        ///
        /// Returns "" for no-value, InvalidElementId, or any read error.
        /// </summary>
        private static string GetParamValue(Parameter p)
        {
            try
            {
                if (p == null || !p.HasValue) return "";

                return p.StorageType switch
                {
                    StorageType.String =>
                        p.AsString() ?? "",

                    StorageType.Double =>
                        // AsValueString() gives formatted value with project units (e.g. "5400 mm")
                        p.AsValueString() ?? p.AsDouble().ToString("G10"),

                    StorageType.Integer =>
                        // AsValueString() returns Yes/No, enum label, etc. where Revit knows the type
                        p.AsValueString() ?? p.AsInteger().ToString(),

                    StorageType.ElementId =>
                        p.AsElementId() == ElementId.InvalidElementId
                            ? ""
                            : (p.AsValueString() ?? p.AsElementId().IntegerValue.ToString()),

                    _ => ""
                };
            }
            catch { return ""; }
        }

        private static string GetFamilyTypeName(Element elem)
        {
            try
            {
                if (elem is FamilyInstance fi)
                    return $"{fi.Symbol?.FamilyName} : {fi.Symbol?.Name}";
                return elem.Name ?? "";
            }
            catch { return ""; }
        }
    }

    // ═════════════════════════════════════════════════════════════════════════
    // WRITER  — JSON and CSV output
    // ═════════════════════════════════════════════════════════════════════════

    internal static class Writer
    {
        // ── JSON ──────────────────────────────────────────────────────────────

      internal static void WriteJson(ModelReport report)
{
    var settings = new JsonSerializerSettings
    {
        Formatting        = Formatting.Indented,
        NullValueHandling = NullValueHandling.Ignore,
    };
    File.WriteAllText("result.json", JsonConvert.SerializeObject(report, settings));
    Console.WriteLine("[RevitExtractor] result.json written");
}

        // ── CSV ───────────────────────────────────────────────────────────────
        //
        // Layout: one row per element, columns are:
        //
        //   Category | ElementId | FamilyType
        //   | <all unique instance param names across all categories>
        //   | Type_<all unique type param names across all categories>
        //
        // Instance and type columns are in separate blocks, type names are
        // prefixed with "Type_" to avoid collisions with instance param names.
        //
        // Two passes:
        //   Pass 1 — collect union of all instance + type param names
        //   Pass 2 — write rows, filling "" for absent params
        // ─────────────────────────────────────────────────────────────────────

        internal static void WriteCsv(ModelReport report)
        {
            // ── Pass 1: collect all unique param names ────────────────────
            var instanceCols = new LinkedList<string>(); // preserves insertion order
            var instanceSet  = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var typeCols     = new LinkedList<string>();
            var typeSet      = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var elements in report.Categories.Values)
            {
                foreach (var elem in elements)
                {
                    foreach (var k in elem.InstanceParameters.Keys)
                        if (instanceSet.Add(k)) instanceCols.AddLast(k);

                    foreach (var k in elem.TypeParameters.Keys)
                        if (typeSet.Add(k)) typeCols.AddLast(k);
                }
            }

            var instanceColList = instanceCols.ToList();
            var typeColList     = typeCols.ToList();

            // ── Pass 2: write CSV ─────────────────────────────────────────
            var sb = new StringBuilder();

            // Header
            var header = new List<string> { "Category", "ElementId", "FamilyType" };
            header.AddRange(instanceColList);
            header.AddRange(typeColList.Select(c => "Type_" + c));
            sb.AppendLine(string.Join(",", header.Select(CsvEscape)));

            // Rows
            foreach (var (category, elements) in report.Categories)
            {
                foreach (var elem in elements)
                {
                    var row = new List<string>
                    {
                        category,
                        elem.ElementId,
                        elem.FamilyType ?? "",
                    };

                    foreach (var col in instanceColList)
                        row.Add(elem.InstanceParameters.TryGetValue(col, out var iv) ? iv : "");

                    foreach (var col in typeColList)
                        row.Add(elem.TypeParameters.TryGetValue(col, out var tv) ? tv : "");

                    sb.AppendLine(string.Join(",", row.Select(CsvEscape)));
                }
            }

            File.WriteAllText("result.csv", sb.ToString(), Encoding.UTF8);
            Console.WriteLine($"[RevitExtractor] result.csv written " +
                $"({instanceColList.Count} instance cols, {typeColList.Count} type cols)");
        }

        /// <summary>
        /// RFC 4180 CSV escaping: wrap in quotes if value contains comma, quote, or newline.
        /// </summary>
        private static string CsvEscape(string value)
        {
            if (value == null) return "";
            if (value.IndexOfAny(new[] { ',', '"', '\n', '\r' }) < 0) return value;
            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }
    }

    // ═════════════════════════════════════════════════════════════════════════
    // DATA MODELS
    // ═════════════════════════════════════════════════════════════════════════

    public class ModelReport
    {
        public string ExtractedAt { get; set; } = "";
        public ProjectInfoData?  ProjectInfo   { get; set; }
        public Dictionary<string, int>?  ElementCounts { get; set; }
        public List<LevelData>?  Levels        { get; set; }
        public List<string>?     Warnings      { get; set; }

        /// <summary>
        /// Key = category label (e.g. "Walls"), Value = list of elements with full params.
        /// </summary>
        public Dictionary<string, List<ElementData>> Categories { get; set; }
            = new Dictionary<string, List<ElementData>>();
    }

    public class ProjectInfoData
    {
        public string? Name             { get; set; }
        public string? Number           { get; set; }
        public string? ClientName       { get; set; }
        public string? Address          { get; set; }
        public string? BuildingName     { get; set; }
        public string? Status           { get; set; }
        public string? OrganizationName { get; set; }
    }

    public class LevelData
    {
        public string? Name      { get; set; }
        public double  ElevationM { get; set; }
    }

    public class ElementData
    {
        /// <summary>Revit ElementId as string — stable cross-session reference.</summary>
        public string ElementId { get; set; } = "";

        /// <summary>"FamilyName : TypeName" for family instances; element Name for system families.</summary>
        public string? FamilyType { get; set; }

        /// <summary>All instance parameters. Missing/null values are stored as "".</summary>
        public Dictionary<string, string> InstanceParameters { get; set; }
            = new Dictionary<string, string>();

        /// <summary>All type parameters from the element's type. Cached — identical types share one read.</summary>
        public Dictionary<string, string> TypeParameters { get; set; }
            = new Dictionary<string, string>();
    }
}
