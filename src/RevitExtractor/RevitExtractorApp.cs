using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Newtonsoft.Json;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
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

        // The last (most recent) project phase — used to resolve which room a
        // family instance sits in. Room bindings are phase-specific, so we look
        // them up against the model's final phase.
        private readonly Phase _lastPhase;

        // All placed, bounded rooms — cached once for the geometric point-in-room
        // fallback. Many OOTB families (furniture, plumbing) have no Room
        // Calculation Point, so FamilyInstance.Room returns null even when the
        // instance clearly sits inside a room; IsPointInRoom recovers those.
        private readonly IList<Room> _placedRooms;

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

        internal Extractor(Document doc)
        {
            _doc         = doc;
            _lastPhase   = GetLastPhase(doc);
            _placedRooms = GetPlacedRooms(doc);
        }

        private static Phase GetLastPhase(Document doc)
        {
            try
            {
                var phases = doc.Phases;
                if (phases == null || phases.Size == 0) return null;
                return phases.get_Item(phases.Size - 1) as Phase;
            }
            catch { return null; }
        }

        private static IList<Room> GetPlacedRooms(Document doc)
        {
            try
            {
                return new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Rooms)
                    .WhereElementIsNotElementType()
                    .OfType<Room>()
                    .Where(r => r.Area > 0)   // placed + bounded only (unplaced rooms have Area 0)
                    .ToList();
            }
            catch { return new List<Room>(); }
        }

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
                    var (roomName, roomNumber) = GetInstanceRoom(elem);
                    results.Add(new ElementData
                    {
                        ElementId          = elem.Id.Value.ToString(),
                        FamilyType         = GetFamilyTypeName(elem),
                        RoomName           = roomName,
                        RoomNumber         = roomNumber,
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
                            : (p.AsValueString() ?? p.AsElementId().Value.ToString()),

                    _ => ""
                };
            }
            catch { return ""; }
        }

        // ── Room placement (family instances only) ────────────────────────

        /// <summary>
        /// Resolves the room a family instance occupies, returning (name, number).
        ///
        /// Strategy — most reliable first:
        ///   1. Phase-based get_Room(lastPhase)  ← geometry-accurate, phase-aware
        ///   2. Phase-based From/To room         ← doors, windows (openings)
        ///   3. Phase-less .Room / .FromRoom / .ToRoom fallbacks
        ///   4. Geometric point-in-room test     ← recovers families with no
        ///      Room Calculation Point (most OOTB furniture, plumbing, fixtures)
        ///
        /// System families (walls, floors, etc.) are not FamilyInstances and
        /// return ("", ""). Any API throw is swallowed — a missing room must
        /// never abort the element.
        /// </summary>
        private (string Name, string Number) GetInstanceRoom(Element elem)
        {
            try
            {
                if (!(elem is FamilyInstance fi)) return ("", "");

                Room room = null;

                if (_lastPhase != null)
                {
                    try { room = fi.get_Room(_lastPhase); } catch { }
                    if (room == null) { try { room = fi.get_FromRoom(_lastPhase); } catch { } }
                    if (room == null) { try { room = fi.get_ToRoom(_lastPhase);   } catch { } }
                }

                if (room == null) { try { room = fi.Room;     } catch { } }
                if (room == null) { try { room = fi.FromRoom; } catch { } }
                if (room == null) { try { room = fi.ToRoom;   } catch { } }

                // Geometric fallback: test the instance's location point against
                // each placed room. Recovers families with no Room Calculation Point.
                //
                // Floor-standing furniture has its location point ON the floor
                // plane (Z = level elevation), i.e. right on the room's lower
                // boundary, where IsPointInRoom is unreliable. Nudge the test
                // point up ~2 ft so it sits clearly inside the room volume.
                if (room == null && _placedRooms.Count > 0
                    && fi.Location is LocationPoint loc)
                {
                    const double NUDGE_FT = 2.0;   // Revit internal units are feet
                    var basePt   = loc.Point;
                    var raisedPt = new XYZ(basePt.X, basePt.Y, basePt.Z + NUDGE_FT);
                    foreach (var r in _placedRooms)
                    {
                        try
                        {
                            if (r.IsPointInRoom(raisedPt) || r.IsPointInRoom(basePt))
                            {
                                room = r;
                                break;
                            }
                        }
                        catch { }
                    }
                }

                if (room == null) return ("", "");

                var name = room.get_Parameter(BuiltInParameter.ROOM_NAME)?.AsString()
                           ?? room.Name ?? "";
                var number = room.Number ?? "";
                return (name.Trim(), number.Trim());
            }
            catch { return ("", ""); }
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
            var header = new List<string> { "Category", "ElementId", "FamilyType", "RoomName", "RoomNumber" };
            header.AddRange(instanceColList);
            header.AddRange(typeColList.Select(c => "Type_" + c));
            sb.AppendLine(string.Join(",", header.Select(CsvEscape)));

            // Rows
            foreach (var kvp in report.Categories)
            {
                var category = kvp.Key;
                var elements = kvp.Value;
                foreach (var elem in elements)
                {
                    var row = new List<string>
                    {
                        category,
                        elem.ElementId,
                        elem.FamilyType ?? "",
                        elem.RoomName ?? "",
                        elem.RoomNumber ?? "",
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

        /// <summary>Name of the room this family instance sits in. "" if none/not applicable.</summary>
        public string? RoomName { get; set; }

        /// <summary>Number of the room this family instance sits in. "" if none/not applicable.</summary>
        public string? RoomNumber { get; set; }

        /// <summary>All instance parameters. Missing/null values are stored as "".</summary>
        public Dictionary<string, string> InstanceParameters { get; set; }
            = new Dictionary<string, string>();

        /// <summary>All type parameters from the element's type. Cached — identical types share one read.</summary>
        public Dictionary<string, string> TypeParameters { get; set; }
            = new Dictionary<string, string>();
    }
}
