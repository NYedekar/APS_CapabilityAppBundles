using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.DatabaseServices;

namespace CADStandardsChecker
{
    internal static class ComplianceChecker
    {
        internal static ComplianceReport Run(Database db, RunParams? runParams, string? dwsPath, string fileName)
        {
            var report = new ComplianceReport
            {
                CheckedAt = DateTime.UtcNow.ToString("o"),
                File      = fileName,
            };

            // ── DWS-based checks (compare drawing objects against standards file) ──
            if (!string.IsNullOrEmpty(dwsPath) && File.Exists(dwsPath))
            {
                try
                {
                    using (var dwsDb = new Database(false, true))
                    {
                        dwsDb.ReadDwgFile(dwsPath, FileOpenMode.OpenForReadAndAllShare, false, null);
                        report.Rules.Add(CheckLayersVsDws(db, dwsDb));
                        report.Rules.Add(CheckTextStylesVsDws(db, dwsDb));
                        report.Rules.Add(CheckDimStylesVsDws(db, dwsDb));
                        report.Rules.Add(CheckLinetypesVsDws(db, dwsDb));
                    }
                }
                catch (System.Exception ex)
                {
                    report.Rules.Add(new RuleResult
                    {
                        RuleId   = "DWS_LOAD_ERROR",
                        Category = "StandardsFile",
                        Result   = "fail",
                        Message  = $"Could not open standards file: {ex.Message}",
                    });
                }
            }

            // ── Custom JSON rules ────────────────────────────────────────────────
            var rules = runParams?.Rules;
            if (rules != null)
            {
                if (rules.LayerNamePatterns.Count > 0)
                    report.Rules.Add(CheckLayerNamePatterns(db, rules.LayerNamePatterns));

                if (rules.RequiredLayers.Count > 0)
                    report.Rules.Add(CheckRequiredLayers(db, rules.RequiredLayers));

                if (rules.ForbiddenLayers.Count > 0)
                    report.Rules.Add(CheckForbiddenLayers(db, rules.ForbiddenLayers));

                if (rules.TextStyleWhitelist.Count > 0)
                    report.Rules.Add(CheckTextStyleWhitelist(db, rules.TextStyleWhitelist));

                if (rules.DimStyleWhitelist.Count > 0)
                    report.Rules.Add(CheckDimStyleWhitelist(db, rules.DimStyleWhitelist));

                if (rules.RequiredBlocks.Count > 0)
                    report.Rules.Add(CheckRequiredBlocks(db, rules.RequiredBlocks));

                if (rules.LinetypeWhitelist.Count > 0)
                    report.Rules.Add(CheckLinetypeWhitelist(db, rules.LinetypeWhitelist));

                if (rules.RequirePurge)
                    report.Rules.Add(CheckPurgeStatus(db));
            }

            if (report.Rules.Count == 0)
            {
                report.Rules.Add(new RuleResult
                {
                    RuleId      = "NO_RULES_CONFIGURED",
                    Category    = "Configuration",
                    Result      = "warning",
                    Message     = "No standards file or rules payload provided. Nothing to check.",
                    Remediation = "Provide a .dws standards file via the standardsFile argument, or a JSON rules payload via the paramsFile argument.",
                });
            }

            // ── Roll up overall status ────────────────────────────────────────────
            report.Summary.Pass    = report.Rules.Count(r => r.Result == "pass");
            report.Summary.Warning = report.Rules.Count(r => r.Result == "warning");
            report.Summary.Fail    = report.Rules.Count(r => r.Result == "fail");

            report.OverallStatus = report.Summary.Fail > 0    ? "fail"
                                 : report.Summary.Warning > 0 ? "pass-with-warnings"
                                 :                              "pass";
            return report;
        }

        // ═══════════════════════════════════════════════════════════════════════
        // DWS-based checks — open the .dws as a secondary Database and compare
        // ═══════════════════════════════════════════════════════════════════════

        private static RuleResult CheckLayersVsDws(Database dwg, Database dws)
        {
            var result = new RuleResult
            {
                RuleId   = "DWS_LAYER_COMPLIANCE",
                Category = "LayerNaming",
                Message  = "Layers comply with the standards file.",
            };

            var stdNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            // key = name (lower), value = (color, linetype)
            var stdProps = new Dictionary<string, (short Color, string Linetype)>(StringComparer.OrdinalIgnoreCase);

            using (var tr = dws.TransactionManager.StartTransaction())
            {
                var lt = (LayerTable)tr.GetObject(dws.LayerTableId, OpenMode.ForRead);
                foreach (ObjectId id in lt)
                {
                    var ltr = (LayerTableRecord)tr.GetObject(id, OpenMode.ForRead);
                    stdNames.Add(ltr.Name);
                    string ltype = GetLinetypeName(tr, ltr.LinetypeObjectId);
                    stdProps[ltr.Name] = (ltr.Color.ColorIndex, ltype);
                }
                tr.Commit();
            }

            using (var tr = dwg.TransactionManager.StartTransaction())
            {
                var lt = (LayerTable)tr.GetObject(dwg.LayerTableId, OpenMode.ForRead);
                foreach (ObjectId id in lt)
                {
                    var ltr = (LayerTableRecord)tr.GetObject(id, OpenMode.ForRead);
                    string name = ltr.Name;
                    if (name == "0" || name == "Defpoints") continue; // reserved layers always allowed

                    if (!stdNames.Contains(name))
                    {
                        result.Offenders.Add(new Offender
                        {
                            ObjectType = "Layer",
                            Name       = name,
                            Value      = name,
                            Expected   = "Layer must be defined in the standards file",
                        });
                    }
                    else if (stdProps.TryGetValue(name, out var std))
                    {
                        // Check color
                        if (std.Color != 0 && ltr.Color.ColorIndex != std.Color)
                        {
                            result.Offenders.Add(new Offender
                            {
                                ObjectType = "Layer",
                                Name       = name,
                                Value      = $"color={ltr.Color.ColorIndex}",
                                Expected   = $"color={std.Color} (per standards file)",
                            });
                        }
                        // Check linetype
                        string dwgLtype = GetLinetypeName(tr, ltr.LinetypeObjectId);
                        if (!string.IsNullOrEmpty(std.Linetype) && !std.Linetype.Equals(dwgLtype, StringComparison.OrdinalIgnoreCase))
                        {
                            result.Offenders.Add(new Offender
                            {
                                ObjectType = "Layer",
                                Name       = name,
                                Value      = $"linetype={dwgLtype}",
                                Expected   = $"linetype={std.Linetype} (per standards file)",
                            });
                        }
                    }
                }
                tr.Commit();
            }

            result.Result = result.Offenders.Count == 0 ? "pass" : "fail";
            if (result.Result == "fail")
            {
                result.Message     = $"{result.Offenders.Count} layer(s) are non-compliant with the standards file.";
                result.Remediation = "Rename non-standard layers to match the standards file names, and correct colors and linetypes where flagged.";
            }
            return result;
        }

        private static RuleResult CheckTextStylesVsDws(Database dwg, Database dws)
        {
            var result = new RuleResult
            {
                RuleId   = "DWS_TEXT_STYLE_COMPLIANCE",
                Category = "TextStyles",
                Message  = "Text styles comply with the standards file.",
            };

            var stdStyles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            using (var tr = dws.TransactionManager.StartTransaction())
            {
                var tt = (TextStyleTable)tr.GetObject(dws.TextStyleTableId, OpenMode.ForRead);
                foreach (ObjectId id in tt)
                {
                    var ttr = (TextStyleTableRecord)tr.GetObject(id, OpenMode.ForRead);
                    stdStyles.Add(ttr.Name);
                }
                tr.Commit();
            }

            using (var tr = dwg.TransactionManager.StartTransaction())
            {
                var tt = (TextStyleTable)tr.GetObject(dwg.TextStyleTableId, OpenMode.ForRead);
                foreach (ObjectId id in tt)
                {
                    var ttr = (TextStyleTableRecord)tr.GetObject(id, OpenMode.ForRead);
                    if (ttr.Name == "Standard") continue;
                    if (!stdStyles.Contains(ttr.Name))
                    {
                        result.Offenders.Add(new Offender
                        {
                            ObjectType = "TextStyle",
                            Name       = ttr.Name,
                            Expected   = "Text style must be defined in the standards file",
                        });
                    }
                }
                tr.Commit();
            }

            result.Result = result.Offenders.Count == 0 ? "pass" : "fail";
            if (result.Result == "fail")
            {
                result.Message     = $"{result.Offenders.Count} text style(s) are non-compliant with the standards file.";
                result.Remediation = "Remove or rename non-standard text styles to match the approved styles defined in the standards file.";
            }
            return result;
        }

        private static RuleResult CheckDimStylesVsDws(Database dwg, Database dws)
        {
            var result = new RuleResult
            {
                RuleId   = "DWS_DIM_STYLE_COMPLIANCE",
                Category = "DimensionStyles",
                Message  = "Dimension styles comply with the standards file.",
            };

            var stdStyles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            using (var tr = dws.TransactionManager.StartTransaction())
            {
                var dt = (DimStyleTable)tr.GetObject(dws.DimStyleTableId, OpenMode.ForRead);
                foreach (ObjectId id in dt)
                {
                    var dsr = (DimStyleTableRecord)tr.GetObject(id, OpenMode.ForRead);
                    stdStyles.Add(dsr.Name);
                }
                tr.Commit();
            }

            using (var tr = dwg.TransactionManager.StartTransaction())
            {
                var dt = (DimStyleTable)tr.GetObject(dwg.DimStyleTableId, OpenMode.ForRead);
                foreach (ObjectId id in dt)
                {
                    var dsr = (DimStyleTableRecord)tr.GetObject(id, OpenMode.ForRead);
                    if (dsr.Name == "Standard") continue;
                    if (!stdStyles.Contains(dsr.Name))
                    {
                        result.Offenders.Add(new Offender
                        {
                            ObjectType = "DimStyle",
                            Name       = dsr.Name,
                            Expected   = "Dimension style must be defined in the standards file",
                        });
                    }
                }
                tr.Commit();
            }

            result.Result = result.Offenders.Count == 0 ? "pass" : "fail";
            if (result.Result == "fail")
            {
                result.Message     = $"{result.Offenders.Count} dimension style(s) are non-compliant with the standards file.";
                result.Remediation = "Remove or rename non-standard dimension styles to match the approved styles defined in the standards file.";
            }
            return result;
        }

        private static RuleResult CheckLinetypesVsDws(Database dwg, Database dws)
        {
            var result = new RuleResult
            {
                RuleId   = "DWS_LINETYPE_COMPLIANCE",
                Category = "Linetypes",
                Message  = "Linetypes comply with the standards file.",
            };

            var stdLtypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            using (var tr = dws.TransactionManager.StartTransaction())
            {
                var ltt = (LinetypeTable)tr.GetObject(dws.LinetypeTableId, OpenMode.ForRead);
                foreach (ObjectId id in ltt)
                {
                    var lttr = (LinetypeTableRecord)tr.GetObject(id, OpenMode.ForRead);
                    stdLtypes.Add(lttr.Name);
                }
                tr.Commit();
            }

            // Always-present linetypes that are exempt from the standard check
            var exempt = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                { "Continuous", "ByLayer", "ByBlock" };

            using (var tr = dwg.TransactionManager.StartTransaction())
            {
                var ltt = (LinetypeTable)tr.GetObject(dwg.LinetypeTableId, OpenMode.ForRead);
                foreach (ObjectId id in ltt)
                {
                    var lttr = (LinetypeTableRecord)tr.GetObject(id, OpenMode.ForRead);
                    if (exempt.Contains(lttr.Name)) continue;
                    if (!stdLtypes.Contains(lttr.Name))
                    {
                        result.Offenders.Add(new Offender
                        {
                            ObjectType = "Linetype",
                            Name       = lttr.Name,
                            Expected   = "Linetype must be defined in the standards file",
                        });
                    }
                }
                tr.Commit();
            }

            result.Result = result.Offenders.Count == 0 ? "pass" : "fail";
            if (result.Result == "fail")
            {
                result.Message     = $"{result.Offenders.Count} linetype(s) are non-compliant with the standards file.";
                result.Remediation = "Remove non-standard linetypes from the drawing or load the correct linetype definitions from the standards file.";
            }
            return result;
        }

        // ═══════════════════════════════════════════════════════════════════════
        // Custom JSON rule checks
        // ═══════════════════════════════════════════════════════════════════════

        private static RuleResult CheckLayerNamePatterns(Database db, List<string> patterns)
        {
            var result = new RuleResult
            {
                RuleId   = "LAYER_NAME_PATTERN",
                Category = "LayerNaming",
                Message  = "All layer names match the required naming conventions.",
            };

            var compiledPatterns = patterns
                .Select(p => { try { return new Regex(p, RegexOptions.IgnoreCase); } catch { return null; } })
                .Where(r => r != null)
                .ToList();

            if (compiledPatterns.Count == 0)
            {
                result.Result  = "warning";
                result.Message = "No valid regex patterns in layerNamePatterns — nothing checked.";
                return result;
            }

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                foreach (ObjectId id in lt)
                {
                    var ltr  = (LayerTableRecord)tr.GetObject(id, OpenMode.ForRead);
                    string name = ltr.Name;
                    if (name == "0" || name == "Defpoints") continue;

                    bool matched = compiledPatterns.Any(p => p!.IsMatch(name));
                    if (!matched)
                    {
                        result.Offenders.Add(new Offender
                        {
                            ObjectType = "Layer",
                            Name       = name,
                            Value      = name,
                            Expected   = $"Matches one of: {string.Join(" | ", patterns)}",
                        });
                    }
                }
                tr.Commit();
            }

            result.Result = result.Offenders.Count == 0 ? "pass" : "fail";
            if (result.Result == "fail")
            {
                result.Message     = $"{result.Offenders.Count} layer(s) have non-conforming names.";
                result.Remediation = $"Rename the flagged layers to match one of the required patterns: {string.Join(", ", patterns)}";
            }
            return result;
        }

        private static RuleResult CheckRequiredLayers(Database db, List<RequiredLayer> required)
        {
            var result = new RuleResult
            {
                RuleId   = "REQUIRED_LAYERS",
                Category = "LayerNaming",
                Message  = "All required layers are present with correct properties.",
            };

            // Store ObjectIds (immutable, safe to use after transaction commits) not DBObject refs
            var existing = new Dictionary<string, ObjectId>(StringComparer.OrdinalIgnoreCase);
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                foreach (ObjectId id in lt)
                {
                    var ltr = (LayerTableRecord)tr.GetObject(id, OpenMode.ForRead);
                    existing[ltr.Name] = id;
                }
                tr.Commit();
            }

            foreach (var req in required)
            {
                if (!existing.TryGetValue(req.Name, out ObjectId layerId))
                {
                    result.Offenders.Add(new Offender
                    {
                        ObjectType = "Layer",
                        Name       = req.Name,
                        Value      = "(missing)",
                        Expected   = "Layer must exist",
                    });
                    continue;
                }
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var record = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForRead);
                    if (req.Color.HasValue && record.Color.ColorIndex != req.Color.Value)
                    {
                        result.Offenders.Add(new Offender
                        {
                            ObjectType = "Layer",
                            Name       = req.Name,
                            Value      = $"color={record.Color.ColorIndex}",
                            Expected   = $"color={req.Color.Value}",
                        });
                    }
                    if (!string.IsNullOrEmpty(req.Linetype))
                    {
                        string actual = GetLinetypeName(tr, record.LinetypeObjectId);
                        if (!actual.Equals(req.Linetype, StringComparison.OrdinalIgnoreCase))
                        {
                            result.Offenders.Add(new Offender
                            {
                                ObjectType = "Layer",
                                Name       = req.Name,
                                Value      = $"linetype={actual}",
                                Expected   = $"linetype={req.Linetype}",
                            });
                        }
                    }
                    tr.Commit();
                }
            }

            result.Result = result.Offenders.Count == 0 ? "pass" : "fail";
            if (result.Result == "fail")
            {
                result.Message     = $"{result.Offenders.Count} required layer(s) are missing or have incorrect properties.";
                result.Remediation = "Add the missing layers and/or correct their color and linetype to match the project standards.";
            }
            return result;
        }

        private static RuleResult CheckForbiddenLayers(Database db, List<string> forbidden)
        {
            var result = new RuleResult
            {
                RuleId   = "FORBIDDEN_LAYERS",
                Category = "LayerNaming",
                Message  = "No forbidden layers found in the drawing.",
            };

            var forbidSet = new HashSet<string>(forbidden, StringComparer.OrdinalIgnoreCase);
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                foreach (ObjectId id in lt)
                {
                    var ltr = (LayerTableRecord)tr.GetObject(id, OpenMode.ForRead);
                    if (forbidSet.Contains(ltr.Name))
                    {
                        result.Offenders.Add(new Offender
                        {
                            ObjectType = "Layer",
                            Name       = ltr.Name,
                            Value      = ltr.Name,
                            Expected   = "Layer name is on the forbidden list",
                        });
                    }
                }
                tr.Commit();
            }

            result.Result = result.Offenders.Count == 0 ? "pass" : "fail";
            if (result.Result == "fail")
            {
                result.Message     = $"{result.Offenders.Count} forbidden layer(s) found.";
                result.Remediation = "Delete or rename the flagged layers, moving any objects on them to the appropriate standard layer first.";
            }
            return result;
        }

        private static RuleResult CheckTextStyleWhitelist(Database db, List<string> whitelist)
        {
            var result = new RuleResult
            {
                RuleId   = "TEXT_STYLE_WHITELIST",
                Category = "TextStyles",
                Message  = "All text styles are on the approved list.",
            };

            var allowed = new HashSet<string>(whitelist, StringComparer.OrdinalIgnoreCase);
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var tt = (TextStyleTable)tr.GetObject(db.TextStyleTableId, OpenMode.ForRead);
                foreach (ObjectId id in tt)
                {
                    var ttr = (TextStyleTableRecord)tr.GetObject(id, OpenMode.ForRead);
                    if (!allowed.Contains(ttr.Name))
                    {
                        result.Offenders.Add(new Offender
                        {
                            ObjectType = "TextStyle",
                            Name       = ttr.Name,
                            Expected   = $"One of: {string.Join(", ", whitelist)}",
                        });
                    }
                }
                tr.Commit();
            }

            result.Result = result.Offenders.Count == 0 ? "pass" : "fail";
            if (result.Result == "fail")
            {
                result.Message     = $"{result.Offenders.Count} unapproved text style(s) found.";
                result.Remediation = "Replace references to non-standard text styles with approved styles, then purge the unapproved style definitions.";
            }
            return result;
        }

        private static RuleResult CheckDimStyleWhitelist(Database db, List<string> whitelist)
        {
            var result = new RuleResult
            {
                RuleId   = "DIM_STYLE_WHITELIST",
                Category = "DimensionStyles",
                Message  = "All dimension styles are on the approved list.",
            };

            var allowed = new HashSet<string>(whitelist, StringComparer.OrdinalIgnoreCase);
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var dt = (DimStyleTable)tr.GetObject(db.DimStyleTableId, OpenMode.ForRead);
                foreach (ObjectId id in dt)
                {
                    var dsr = (DimStyleTableRecord)tr.GetObject(id, OpenMode.ForRead);
                    if (!allowed.Contains(dsr.Name))
                    {
                        result.Offenders.Add(new Offender
                        {
                            ObjectType = "DimStyle",
                            Name       = dsr.Name,
                            Expected   = $"One of: {string.Join(", ", whitelist)}",
                        });
                    }
                }
                tr.Commit();
            }

            result.Result = result.Offenders.Count == 0 ? "pass" : "fail";
            if (result.Result == "fail")
            {
                result.Message     = $"{result.Offenders.Count} unapproved dimension style(s) found.";
                result.Remediation = "Replace references to non-standard dimension styles with approved styles, then purge the unapproved style definitions.";
            }
            return result;
        }

        private static RuleResult CheckRequiredBlocks(Database db, List<string> required)
        {
            var result = new RuleResult
            {
                RuleId   = "REQUIRED_BLOCKS",
                Category = "BlockDefinitions",
                Message  = "All required block definitions are present.",
            };

            var existing = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                foreach (ObjectId id in bt)
                {
                    var btr = (BlockTableRecord)tr.GetObject(id, OpenMode.ForRead);
                    existing.Add(btr.Name);
                }
                tr.Commit();
            }

            foreach (string req in required)
            {
                if (!existing.Contains(req))
                {
                    result.Offenders.Add(new Offender
                    {
                        ObjectType = "Block",
                        Name       = req,
                        Value      = "(missing)",
                        Expected   = "Block definition must exist",
                    });
                }
            }

            result.Result = result.Offenders.Count == 0 ? "pass" : "fail";
            if (result.Result == "fail")
            {
                result.Message     = $"{result.Offenders.Count} required block(s) are missing.";
                result.Remediation = "Insert or define the required blocks. For title blocks, insert from the company block library.";
            }
            return result;
        }

        private static RuleResult CheckLinetypeWhitelist(Database db, List<string> whitelist)
        {
            var result = new RuleResult
            {
                RuleId   = "LINETYPE_WHITELIST",
                Category = "Linetypes",
                Message  = "All linetypes are on the approved list.",
            };

            var allowed = new HashSet<string>(whitelist, StringComparer.OrdinalIgnoreCase);
            var exempt  = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Continuous", "ByLayer", "ByBlock" };

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var ltt = (LinetypeTable)tr.GetObject(db.LinetypeTableId, OpenMode.ForRead);
                foreach (ObjectId id in ltt)
                {
                    var lttr = (LinetypeTableRecord)tr.GetObject(id, OpenMode.ForRead);
                    if (exempt.Contains(lttr.Name)) continue;
                    if (!allowed.Contains(lttr.Name))
                    {
                        result.Offenders.Add(new Offender
                        {
                            ObjectType = "Linetype",
                            Name       = lttr.Name,
                            Expected   = $"One of: {string.Join(", ", whitelist)}",
                        });
                    }
                }
                tr.Commit();
            }

            result.Result = result.Offenders.Count == 0 ? "pass" : "fail";
            if (result.Result == "fail")
            {
                result.Message     = $"{result.Offenders.Count} unapproved linetype(s) found.";
                result.Remediation = "Remove unapproved linetype definitions (PURGE) and load only the approved linetypes from the company linetype file.";
            }
            return result;
        }

        private static RuleResult CheckPurgeStatus(Database db)
        {
            var result = new RuleResult
            {
                RuleId   = "PURGE_STATUS",
                Category = "DrawingHygiene",
            };

            // Count unreferenced symbols using PurgeCheck (available headlessly).
            // An ObjectIdCollection is populated with all purgeable objects.
            var purgeable = new ObjectIdCollection();
            db.Purge(purgeable);

            if (purgeable.Count == 0)
            {
                result.Result  = "pass";
                result.Message = "Drawing has no unreferenced (purgeable) symbol definitions.";
                return result;
            }

            // Categorize by object type name
            var countByType = new Dictionary<string, int>();
            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (ObjectId id in purgeable)
                {
                    try
                    {
                        var obj = tr.GetObject(id, OpenMode.ForRead);
                        string typeName = obj.GetType().Name;
                        if (!countByType.TryGetValue(typeName, out int c))
                            countByType[typeName] = 0;
                        countByType[typeName]++;
                    }
                    catch { /* skip inaccessible objects */ }
                }
                tr.Commit();
            }

            foreach (var kvp in countByType)
            {
                result.Offenders.Add(new Offender
                {
                    ObjectType = kvp.Key,
                    Name       = $"{kvp.Value} unreferenced {kvp.Key}(s)",
                    Expected   = "0 purgeable objects of this type",
                });
            }

            result.Result      = "fail";
            result.Message     = $"Drawing contains {purgeable.Count} purgeable (unreferenced) object(s) — run PURGE before submission.";
            result.Remediation = "Open the drawing locally and run PURGE > PURGEALL, or use the Design Automation AutoCADDrawingUpdater to run PURGE headlessly before checking.";
            return result;
        }

        // ─── Helpers ──────────────────────────────────────────────────────────

        private static string GetLinetypeName(Transaction tr, ObjectId linetypeId)
        {
            try
            {
                var lttr = (LinetypeTableRecord)tr.GetObject(linetypeId, OpenMode.ForRead);
                return lttr.Name;
            }
            catch { return ""; }
        }
    }
}
