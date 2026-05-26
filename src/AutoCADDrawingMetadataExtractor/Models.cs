using System.Collections.Generic;

namespace AutoCADDrawingMetadataExtractor
{
    public class DrawingMetadataReport
    {
        public string ExtractedAt { get; set; } = "";
        public SummaryInfoData? SummaryInfo { get; set; }
        public DrawingSettingsData? DrawingSettings { get; set; }
        public EntityCountsData? EntityCounts { get; set; }
        public List<LayerData> LayerTable { get; set; } = new();
        public List<LayoutData> Layouts { get; set; } = new();
        public List<BlockDefinitionData> BlockDefinitions { get; set; } = new();
        public List<LinetypeData> Linetypes { get; set; } = new();
        public List<TextStyleData> TextStyles { get; set; } = new();
        public List<string> DimStyles { get; set; } = new();
        public List<string> NamedViews { get; set; } = new();
        public List<string> UcsTable { get; set; } = new();
    }

    // ── Combined result shape for ExtractAllDrawingMetadata ──────────────────
    // Keys mirror the 7 individual operationIds so callers can use either path.

    public class CombinedMetadataResult
    {
        public string ExtractedAt { get; set; } = "";
        public DrawingStatisticsResult DrawingStatistics { get; set; } = new();
        public SymbolTableInventoryResult SymbolTableInventory { get; set; } = new();
        public List<LayoutData> LayoutData { get; set; } = new();
        public SummaryInfoData? ProjectCustomProperties { get; set; }
        public DrawingHistoryData? DrawingHistory { get; set; }
        public List<LayerData> LayerTable { get; set; } = new();
        public List<XrefData> XrefList { get; set; } = new();
    }

    public class DrawingStatisticsResult
    {
        public int TotalModelSpaceEntities { get; set; }
        public Dictionary<string, int> ByEntityType { get; set; } = new();
        public Dictionary<string, int> ByLayer { get; set; } = new();
        public int LayerCount { get; set; }
        public int BlockDefinitionCount { get; set; }
        public int XrefCount { get; set; }
        public string? InsertionUnits { get; set; }
        public string? LinearUnits { get; set; }
        public Point2dData? ExtentsMin { get; set; }
        public Point2dData? ExtentsMax { get; set; }
        public Point2dData? LimitsMin { get; set; }
        public Point2dData? LimitsMax { get; set; }
    }

    public class SymbolTableInventoryResult
    {
        public List<LinetypeData> Linetypes { get; set; } = new();
        public List<TextStyleData> TextStyles { get; set; } = new();
        public List<string> DimStyles { get; set; } = new();
        public List<string> NamedViews { get; set; } = new();
        public List<string> UcsTable { get; set; } = new();
    }

    public class DrawingHistoryData
    {
        public string? LastSavedBy { get; set; }
        public string? RevisionNumber { get; set; }
        public string? TotalEditingTime { get; set; }
    }

    public class XrefData
    {
        public string Name { get; set; } = "";
        public string? Path { get; set; }
        public string? XrefStatus { get; set; }
    }

    // ── Shared primitive models ───────────────────────────────────────────────

    public class SummaryInfoData
    {
        public string? Title { get; set; }
        public string? Subject { get; set; }
        public string? Author { get; set; }
        public string? Keywords { get; set; }
        public string? Comments { get; set; }
        public string? LastSavedBy { get; set; }
        public string? RevisionNumber { get; set; }
        // CustomProperties not available in AutoCAD.NET.Core — requires the full Windows package.
    }

    public class DrawingSettingsData
    {
        public string? InsertionUnits { get; set; }
        public string? LinearUnits { get; set; }
        public string? AngularUnits { get; set; }
        public string? Measurement { get; set; }
        public Point2dData? ExtentsMin { get; set; }
        public Point2dData? ExtentsMax { get; set; }
        public Point2dData? LimitsMin { get; set; }
        public Point2dData? LimitsMax { get; set; }
        public Point3dData? InsertionBase { get; set; }
    }

    public class Point2dData
    {
        public double X { get; set; }
        public double Y { get; set; }
    }

    public class Point3dData
    {
        public double X { get; set; }
        public double Y { get; set; }
        public double Z { get; set; }
    }

    public class EntityCountsData
    {
        public int TotalModelSpaceEntities { get; set; }
        public Dictionary<string, int> ByEntityType { get; set; } = new();
        public Dictionary<string, int> ByLayer { get; set; } = new();
    }

    public class LayerData
    {
        public string Name { get; set; } = "";
        public bool IsOff { get; set; }
        public bool IsFrozen { get; set; }
        public bool IsLocked { get; set; }
        public bool IsPlottable { get; set; }
        public int ColorIndex { get; set; }
        public string? TrueColor { get; set; }
        public string? Linetype { get; set; }
        public string? LineWeight { get; set; }
        public string? Description { get; set; }
    }

    public class LayoutData
    {
        public string Name { get; set; } = "";
        public int TabOrder { get; set; }
        public bool IsModelSpace { get; set; }
        public string? PlotterName { get; set; }
        public string? PaperSize { get; set; }
        public string? PlotPaperUnits { get; set; }
        public string? PlotRotation { get; set; }
        public int ViewportCount { get; set; }
    }

    public class BlockDefinitionData
    {
        public string Name { get; set; } = "";
        public int EntityCount { get; set; }
        public List<string> AttributeTags { get; set; } = new();
        public bool IsXref { get; set; }
        public string? XrefPath { get; set; }
        public string? XrefStatus { get; set; }
        public bool IsLayout { get; set; }
        public bool IsAnonymous { get; set; }
    }

    public class LinetypeData
    {
        public string Name { get; set; } = "";
        public string? Description { get; set; }
    }

    public class TextStyleData
    {
        public string Name { get; set; } = "";
        public string? FileName { get; set; }
        public string? BigFontFileName { get; set; }
        public double TextSize { get; set; }
        public double XScale { get; set; }
    }
}
