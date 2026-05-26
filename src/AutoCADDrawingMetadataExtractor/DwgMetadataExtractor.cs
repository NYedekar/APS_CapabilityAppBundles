using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;

namespace AutoCADDrawingMetadataExtractor
{
    internal class DwgMetadataExtractor
    {
        private readonly Database _db;

        internal DwgMetadataExtractor(Database db) => _db = db;

        // ── Full flat report (used by EXTRACTDWGMETADATA) ─────────────────────

        internal DrawingMetadataReport BuildReport()
        {
            var report = new DrawingMetadataReport
            {
                ExtractedAt = DateTime.UtcNow.ToString("o"),
            };

            Console.WriteLine("[MetadataExtractor] Extracting summary info...");
            report.SummaryInfo = GetSummaryInfo();

            Console.WriteLine("[MetadataExtractor] Extracting drawing settings...");
            report.DrawingSettings = GetDrawingSettings();

            // StartTransaction() is the correct .NET API — StartReadOnlyTransaction() does not exist.
            using var tr = _db.TransactionManager.StartTransaction();

            Console.WriteLine("[MetadataExtractor] Extracting layer table...");
            report.LayerTable = GetLayerTable(tr);

            Console.WriteLine("[MetadataExtractor] Extracting layouts...");
            report.Layouts = GetLayouts(tr);

            Console.WriteLine("[MetadataExtractor] Extracting block definitions...");
            report.BlockDefinitions = GetBlockDefinitions(tr);

            Console.WriteLine("[MetadataExtractor] Extracting entity counts...");
            report.EntityCounts = GetEntityCounts(tr);

            Console.WriteLine("[MetadataExtractor] Extracting symbol tables...");
            GetSymbolTables(tr, report);

            return report;
            // Transaction disposed here without Commit — correct for read-only use.
        }

        // ── Combined operationId-keyed report (used by EXTRACTALLDRAWINGMETADATA) ──
        // Single DWG open/close pass; output keys mirror the 7 individual operationIds.

        internal CombinedMetadataResult BuildCombinedReport()
        {
            var result = new CombinedMetadataResult
            {
                ExtractedAt = DateTime.UtcNow.ToString("o"),
            };

            Console.WriteLine("[MetadataExtractor] Building combined report (single pass)...");

            // ProjectCustomProperties + DrawingHistory both come from SummaryInfo
            var summaryInfo = GetSummaryInfo();
            result.ProjectCustomProperties = summaryInfo;
            result.DrawingHistory = new DrawingHistoryData
            {
                LastSavedBy    = summaryInfo.LastSavedBy,
                RevisionNumber = summaryInfo.RevisionNumber,
            };
            try
            {
                result.DrawingHistory.TotalEditingTime = _db.TotalEditingTime.ToString(@"hh\:mm\:ss");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[MetadataExtractor] TotalEditingTime unavailable: {ex.Message}");
            }

            var settings = GetDrawingSettings();

            using var tr = _db.TransactionManager.StartTransaction();

            // LayerTable
            result.LayerTable = GetLayerTable(tr);

            // LayoutData
            result.LayoutData = GetLayouts(tr);

            // BlockDefinitions → XrefList + counts for DrawingStatistics
            var blockDefs = GetBlockDefinitions(tr);
            result.XrefList = blockDefs
                .Where(b => b.IsXref)
                .Select(b => new XrefData { Name = b.Name, Path = b.XrefPath, XrefStatus = b.XrefStatus })
                .ToList();

            // EntityCounts for DrawingStatistics
            var entityCounts = GetEntityCounts(tr);
            result.DrawingStatistics = new DrawingStatisticsResult
            {
                TotalModelSpaceEntities = entityCounts.TotalModelSpaceEntities,
                ByEntityType            = entityCounts.ByEntityType,
                ByLayer                 = entityCounts.ByLayer,
                LayerCount              = result.LayerTable.Count,
                BlockDefinitionCount    = blockDefs.Count(b => !b.IsLayout && !b.IsAnonymous && !b.IsXref),
                XrefCount               = result.XrefList.Count,
                InsertionUnits          = settings.InsertionUnits,
                LinearUnits             = settings.LinearUnits,
                ExtentsMin              = settings.ExtentsMin,
                ExtentsMax              = settings.ExtentsMax,
                LimitsMin               = settings.LimitsMin,
                LimitsMax               = settings.LimitsMax,
            };

            // SymbolTableInventory
            var tempReport = new DrawingMetadataReport();
            GetSymbolTables(tr, tempReport);
            result.SymbolTableInventory = new SymbolTableInventoryResult
            {
                Linetypes   = tempReport.Linetypes,
                TextStyles  = tempReport.TextStyles,
                DimStyles   = tempReport.DimStyles,
                NamedViews  = tempReport.NamedViews,
                UcsTable    = tempReport.UcsTable,
            };

            Console.WriteLine($"[MetadataExtractor] Combined report built: " +
                $"{result.LayerTable.Count} layers, {result.LayoutData.Count} layouts, " +
                $"{result.XrefList.Count} xrefs, {entityCounts.TotalModelSpaceEntities} entities.");

            return result;
        }

        // ── Summary Info ──────────────────────────────────────────────────────

        private SummaryInfoData GetSummaryInfo()
        {
            var info = _db.SummaryInfo;
            var data = new SummaryInfoData
            {
                Title          = info.Title,
                Subject        = info.Subject,
                Author         = info.Author,
                Keywords       = info.Keywords,
                Comments       = info.Comments,
                LastSavedBy    = info.LastSavedBy,
                RevisionNumber = info.RevisionNumber,
            };

            return data;
        }

        // ── Drawing Settings ──────────────────────────────────────────────────

        private DrawingSettingsData GetDrawingSettings()
        {
            return new DrawingSettingsData
            {
                InsertionUnits = _db.Insunits.ToString(),
                LinearUnits    = LinearUnitsName(_db.Lunits),
                AngularUnits   = AngularUnitsName(_db.Aunits),
                Measurement    = _db.Measurement.ToString(),
                ExtentsMin     = Pt2(_db.Extmin),
                ExtentsMax     = Pt2(_db.Extmax),
                LimitsMin      = Pt2(_db.Limmin),   // Point2d
                LimitsMax      = Pt2(_db.Limmax),   // Point2d
                InsertionBase  = Pt3(_db.Insbase),
            };
        }

        // ── Layer Table ───────────────────────────────────────────────────────

        private List<LayerData> GetLayerTable(Transaction tr)
        {
            var result = new List<LayerData>();
            var table = (LayerTable)tr.GetObject(_db.LayerTableId, OpenMode.ForRead);

            foreach (ObjectId id in table)
            {
                try
                {
                    var layer = (LayerTableRecord)tr.GetObject(id, OpenMode.ForRead);
                    var data = new LayerData
                    {
                        Name        = layer.Name,
                        IsOff       = layer.IsOff,
                        IsFrozen    = layer.IsFrozen,
                        IsLocked    = layer.IsLocked,
                        IsPlottable = layer.IsPlottable,
                        ColorIndex  = layer.Color.ColorIndex,
                        Linetype    = GetLinetypeName(tr, layer.LinetypeObjectId),
                        LineWeight  = layer.LineWeight.ToString(),
                        Description = layer.Description,
                    };

                    if (layer.Color.ColorMethod == ColorMethod.ByColor)
                        data.TrueColor = $"#{layer.Color.Red:X2}{layer.Color.Green:X2}{layer.Color.Blue:X2}";

                    result.Add(data);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[MetadataExtractor] Skipped layer {id}: {ex.Message}");
                }
            }

            return result;
        }

        // ── Layouts ───────────────────────────────────────────────────────────

        private List<LayoutData> GetLayouts(Transaction tr)
        {
            var result = new List<LayoutData>();
            var layoutDict = (DBDictionary)tr.GetObject(_db.LayoutDictionaryId, OpenMode.ForRead);

            foreach (DBDictionaryEntry entry in layoutDict)
            {
                try
                {
                    var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);

                    // GetViewports() does not exist in .NET API.
                    // Count viewports by iterating the layout's own BTR.
                    int vpCount = 0;
                    if (!layout.BlockTableRecordId.IsNull)
                    {
                        var layoutBtr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);
                        foreach (ObjectId entId in layoutBtr)
                        {
                            try
                            {
                                // Viewport.Number == 1 is the paper-space pseudo-viewport; skip it.
                                if (tr.GetObject(entId, OpenMode.ForRead) is Viewport vp && vp.Number != 1)
                                    vpCount++;
                            }
                            catch { }
                        }
                    }

                    result.Add(new LayoutData
                    {
                        Name           = layout.LayoutName,
                        TabOrder       = layout.TabOrder,
                        IsModelSpace   = layout.ModelType,
                        PlotterName    = layout.PlotConfigurationName,
                        PaperSize      = layout.CanonicalMediaName,
                        PlotPaperUnits = layout.PlotPaperUnits.ToString(),
                        PlotRotation   = layout.PlotRotation.ToString(),
                        ViewportCount  = vpCount,
                    });
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[MetadataExtractor] Skipped layout {entry.Key}: {ex.Message}");
                }
            }

            return result.OrderBy(l => l.TabOrder).ToList();
        }

        // ── Block Definitions ─────────────────────────────────────────────────

        private List<BlockDefinitionData> GetBlockDefinitions(Transaction tr)
        {
            var result = new List<BlockDefinitionData>();
            var blockTable = (BlockTable)tr.GetObject(_db.BlockTableId, OpenMode.ForRead);

            foreach (ObjectId id in blockTable)
            {
                try
                {
                    var btr = (BlockTableRecord)tr.GetObject(id, OpenMode.ForRead);
                    var data = new BlockDefinitionData
                    {
                        Name        = btr.Name,
                        IsXref      = btr.IsFromExternalReference,
                        IsLayout    = btr.IsLayout,
                        IsAnonymous = btr.IsAnonymous,
                        XrefPath    = btr.IsFromExternalReference ? btr.PathName : null,
                        XrefStatus  = btr.IsFromExternalReference ? btr.XrefStatus.ToString() : null,
                    };

                    int count = 0;
                    foreach (ObjectId entId in btr)
                    {
                        count++;
                        try
                        {
                            if (tr.GetObject(entId, OpenMode.ForRead) is AttributeDefinition attDef && !attDef.Constant)
                                data.AttributeTags.Add(attDef.Tag);
                        }
                        catch { }
                    }
                    data.EntityCount = count;

                    result.Add(data);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[MetadataExtractor] Skipped block {id}: {ex.Message}");
                }
            }

            return result;
        }

        // ── Entity Counts ─────────────────────────────────────────────────────

        private EntityCountsData GetEntityCounts(Transaction tr)
        {
            var data = new EntityCountsData();
            var modelSpaceId = SymbolUtilityServices.GetBlockModelSpaceId(_db);
            var modelBtr = (BlockTableRecord)tr.GetObject(modelSpaceId, OpenMode.ForRead);

            foreach (ObjectId entId in modelBtr)
            {
                try
                {
                    var ent = (Entity)tr.GetObject(entId, OpenMode.ForRead);
                    data.TotalModelSpaceEntities++;

                    string typeName = ent.GetType().Name;
                    data.ByEntityType[typeName] = data.ByEntityType.GetValueOrDefault(typeName) + 1;

                    string layerName = ent.Layer ?? "0";
                    data.ByLayer[layerName] = data.ByLayer.GetValueOrDefault(layerName) + 1;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[MetadataExtractor] Skipped entity {entId}: {ex.Message}");
                }
            }

            return data;
        }

        // ── Symbol Tables ─────────────────────────────────────────────────────

        private void GetSymbolTables(Transaction tr, DrawingMetadataReport report)
        {
            var ltTable = (LinetypeTable)tr.GetObject(_db.LinetypeTableId, OpenMode.ForRead);
            foreach (ObjectId id in ltTable)
            {
                try
                {
                    var lt = (LinetypeTableRecord)tr.GetObject(id, OpenMode.ForRead);
                    report.Linetypes.Add(new LinetypeData { Name = lt.Name, Description = lt.AsciiDescription });
                }
                catch { }
            }

            var tsTable = (TextStyleTable)tr.GetObject(_db.TextStyleTableId, OpenMode.ForRead);
            foreach (ObjectId id in tsTable)
            {
                try
                {
                    var ts = (TextStyleTableRecord)tr.GetObject(id, OpenMode.ForRead);
                    report.TextStyles.Add(new TextStyleData
                    {
                        Name            = ts.Name,
                        FileName        = ts.FileName,
                        BigFontFileName = ts.BigFontFileName,
                        TextSize        = ts.TextSize,
                        XScale          = ts.XScale,
                    });
                }
                catch { }
            }

            var dsTable = (DimStyleTable)tr.GetObject(_db.DimStyleTableId, OpenMode.ForRead);
            foreach (ObjectId id in dsTable)
            {
                try { report.DimStyles.Add(((DimStyleTableRecord)tr.GetObject(id, OpenMode.ForRead)).Name); }
                catch { }
            }

            var viewTable = (ViewTable)tr.GetObject(_db.ViewTableId, OpenMode.ForRead);
            foreach (ObjectId id in viewTable)
            {
                try { report.NamedViews.Add(((ViewTableRecord)tr.GetObject(id, OpenMode.ForRead)).Name); }
                catch { }
            }

            var ucsTable = (UcsTable)tr.GetObject(_db.UcsTableId, OpenMode.ForRead);
            foreach (ObjectId id in ucsTable)
            {
                try { report.UcsTable.Add(((UcsTableRecord)tr.GetObject(id, OpenMode.ForRead)).Name); }
                catch { }
            }
        }

        // ── Helpers ───────────────────────────────────────────────────────────

        private string GetLinetypeName(Transaction tr, ObjectId ltId)
        {
            try
            {
                if (ltId.IsNull) return "";
                return ((LinetypeTableRecord)tr.GetObject(ltId, OpenMode.ForRead)).Name;
            }
            catch { return ""; }
        }

        // Limmin/Limmax are Point2d; Extmin/Extmax/Insbase are Point3d.
        private static Point2dData Pt2(Point2d p) => new() { X = Math.Round(p.X, 6), Y = Math.Round(p.Y, 6) };
        private static Point2dData Pt2(Point3d p) => new() { X = Math.Round(p.X, 6), Y = Math.Round(p.Y, 6) };
        private static Point3dData Pt3(Point3d p) => new() { X = Math.Round(p.X, 6), Y = Math.Round(p.Y, 6), Z = Math.Round(p.Z, 6) };

        private static string LinearUnitsName(int code) => code switch
        {
            1 => "Scientific", 2 => "Decimal", 3 => "Engineering",
            4 => "Architectural", 5 => "Fractional", _ => code.ToString(),
        };

        private static string AngularUnitsName(int code) => code switch
        {
            0 => "Degrees", 1 => "DegreesMinutes", 2 => "DegreesMinutesSeconds",
            3 => "Gradians", 4 => "Radians", _ => code.ToString(),
        };
    }
}
