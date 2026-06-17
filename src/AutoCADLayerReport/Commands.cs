// Scaffolded from templates/autocad/CommandClass.cs.
// Capability: report every layer (name, color index, linetype, on/off/frozen/locked)
// as RFC 4180 CSV to result.csv. Uses only stable LayerTable APIs to avoid cross-version traps.
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

                // Tabular output: header row + one row per layer.
                var sb = new StringBuilder();
                sb.Append("Name,ColorIndex,Linetype,IsOff,IsFrozen,IsLocked,IsPlottable\r\n");
                foreach (var l in layers)
                    sb.Append(string.Join(",",
                        Csv(l.Name),
                        l.ColorIndex.ToString(),
                        Csv(l.Linetype),
                        Bool(l.IsOff), Bool(l.IsFrozen), Bool(l.IsLocked), Bool(l.IsPlottable)) + "\r\n");

                // UTF-8 WITHOUT BOM — Encoding.UTF8 emits a BOM that trips strict parsers.
                File.WriteAllText("result.csv", sb.ToString(), new UTF8Encoding(false));
            }
            catch (System.Exception ex)
            {
                // Still emit result.csv so the caller always gets a parseable artifact.
                File.WriteAllText("result.csv",
                    "error\r\n" + Csv(ex.Message) + "\r\n",
                    new UTF8Encoding(false));
            }
        }

        private static string Bool(bool b) => b ? "true" : "false";

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
