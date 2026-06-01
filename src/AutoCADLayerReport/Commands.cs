// Scaffolded from templates/autocad/CommandClass.cs.
// Capability: report every layer (name, color index, linetype, on/off/frozen/locked)
// + a count to result.json. Uses only stable LayerTable APIs to avoid cross-version traps.
//
// Proven command pattern: [assembly: CommandClass] + [assembly: ExtensionApplication(null)],
// static [CommandMethod]. The verb LAYERREPORT must match PackageContents <Command> and the
// activity's COMMAND/.scr value.
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Autodesk.AutoCAD.ApplicationServices.Core;   // Application (core console)
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using Newtonsoft.Json;

[assembly: CommandClass(typeof(AutoCADLayerReport.Commands))]
[assembly: ExtensionApplication(null)]

namespace AutoCADLayerReport
{
    public class Commands
    {
        [CommandMethod("LAYERREPORT", CommandFlags.Modal)]
        public static void LayerReport()
        {
            try
            {
                // In accoreconsole the INPUT drawing is the active document. Read its database —
                // HostApplicationServices.WorkingDatabase alone returns an EMPTY in-memory db
                // (layer "0" only), so prefer MdiActiveDocument.Database and fall back.
                var doc = Application.DocumentManager.MdiActiveDocument;
                var db = doc?.Database ?? HostApplicationServices.WorkingDatabase;
                var layers = new List<LayerRow>();

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    foreach (ObjectId id in lt)
                    {
                        var ltr = (LayerTableRecord)tr.GetObject(id, OpenMode.ForRead);

                        string linetype = "";
                        try
                        {
                            var ltype = (LinetypeTableRecord)tr.GetObject(ltr.LinetypeObjectId, OpenMode.ForRead);
                            linetype = ltype.Name;
                        }
                        catch { /* linetype lookup is best-effort */ }

                        layers.Add(new LayerRow
                        {
                            Name       = ltr.Name,
                            ColorIndex = ltr.Color.ColorIndex,
                            Linetype   = linetype,
                            IsOff      = ltr.IsOff,
                            IsFrozen   = ltr.IsFrozen,
                            IsLocked   = ltr.IsLocked,
                            IsPlottable = ltr.IsPlottable,
                        });
                    }
                    tr.Commit();
                }

                var report = new LayerReportData
                {
                    ExtractedAt = DateTime.UtcNow.ToString("o"),
                    LayerCount  = layers.Count,
                    Layers      = layers,
                };

                string json = JsonConvert.SerializeObject(report, Formatting.Indented);
                // UTF-8 WITHOUT BOM — Encoding.UTF8 emits a BOM that breaks strict JSON parsers.
                File.WriteAllText("result.json", json, new UTF8Encoding(false));
            }
            catch (System.Exception ex)
            {
                var err = new { ok = false, error = ex.Message, stack = ex.StackTrace };
                File.WriteAllText("result.json",
                    JsonConvert.SerializeObject(err, Formatting.Indented),
                    new UTF8Encoding(false));
            }
        }
    }

    internal class LayerReportData
    {
        public string ExtractedAt { get; set; } = "";
        public int LayerCount { get; set; }
        public List<LayerRow> Layers { get; set; } = new List<LayerRow>();
    }

    internal class LayerRow
    {
        public string Name { get; set; } = "";
        public short ColorIndex { get; set; }
        public string Linetype { get; set; } = "";
        public bool IsOff { get; set; }
        public bool IsFrozen { get; set; }
        public bool IsLocked { get; set; }
        public bool IsPlottable { get; set; }
    }
}
