// ════════════════════════════════════════════════════════════════════════════
//  APS DA — AutoCAD plugin entry point (template)
//
//  PROVEN COMMAND PATTERN (loads on the +24_3 engine):
//    [assembly: CommandClass(typeof(...))]   ← registers the command class
//    [assembly: ExtensionApplication(null)]  ← NO IExtensionApplication needed
//    static [CommandMethod] handlers
//
//  net48 TRAPS (each cost a debugging cycle):
//    • Dictionary.GetValueOrDefault is .NET Std 2.1+ — NOT on net48. Use TryGetValue.
//    • Write result.json with `new UTF8Encoding(false)` — Encoding.UTF8 emits a BOM
//      that breaks strict JSON parsers (JS JSON.parse fails on the leading ﻿).
//
//  REPLACE: AUDemoAcadSmoke, APSAUDEMOACAD, and the command body.
// ════════════════════════════════════════════════════════════════════════════
using System;
using System.IO;
using System.Linq;
using System.Text;
using Autodesk.AutoCAD.ApplicationServices.Core;   // Application (core console)
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using Newtonsoft.Json;

[assembly: CommandClass(typeof(AUDemoAcadSmoke.Commands))]
[assembly: ExtensionApplication(null)]

namespace AUDemoAcadSmoke
{
    // ── Optional: params.json contract ──────────────────────────────────────
    // If your activity has a "params" get-arg, the MCP uploads the user's config
    // as params.json. Deserialise it here. Remove if not needed.
    // internal class RunParams
    // {
    //     [JsonProperty("inputMode")] public string? InputMode { get; set; }
    //     // add your fields here
    // }
    // ────────────────────────────────────────────────────────────────────────

    public class Commands
    {
        // The Global/Local verb here MUST match a <Command .../> in PackageContents.xml.
        [CommandMethod("APSAUDEMOACAD", CommandFlags.Modal)]
        public static void RunPrimary()
        {
            try
            {
                // Diagnostic: log all files visible in the working directory.
                // Useful when debugging missing inputs or unexpected DA file names.
                var wdFiles = Directory.GetFiles(Directory.GetCurrentDirectory());
                Console.WriteLine($"[APSAUDEMOACAD] Working dir files: {string.Join(", ", wdFiles.Select(f => Path.GetFileName(f)))}");

                // In accoreconsole the INPUT drawing (opened via /i) is the active document.
                // Resolve its database the way the shipped AutoCADDrawingMetadataExtractor does —
                // active-document-first, falling back to WorkingDatabase.
                var doc = Application.DocumentManager.MdiActiveDocument;
                var db = doc?.Database ?? HostApplicationServices.WorkingDatabase;

                // ── Optional: read params.json ───────────────────────────────
                // RunParams? runParams = null;
                // if (File.Exists("params.json"))
                //     runParams = JsonConvert.DeserializeObject<RunParams>(File.ReadAllText("params.json"));
                // ────────────────────────────────────────────────────────────

                // ── Optional: multi-format input detection ───────────────────
                // When a capability accepts CSV, XLSX, or JSON from a single input arg,
                // DA does not preserve file extensions. Detect format by magic bytes:
                //   PK\x03\x04 (first 4 bytes) → XLSX
                //   [ or {      (first byte)    → JSON
                //   anything else               → CSV
                //
                // if (File.Exists("input.dat")) {
                //     var magic = new byte[4];
                //     using (var fs = File.OpenRead("input.dat")) fs.Read(magic, 0, 4);
                //     bool isXlsx = magic[0] == 0x50 && magic[1] == 0x4B && magic[2] == 0x03 && magic[3] == 0x04;
                //     bool isJson = magic[0] == (byte)'[' || magic[0] == (byte)'{';
                //     // Use ExcelDataReader 3.6.0 + ExcelDataReader.DataSet for XLSX (see csproj).
                // }
                // ────────────────────────────────────────────────────────────

                // ── your extraction / processing logic ──────────────────────
                var result = new { ok = true, command = "APSAUDEMOACAD", framework = "net48" };
                // ────────────────────────────────────────────────────────────

                // UTF-8 WITHOUT BOM — required for downstream JSON.parse to succeed.
                File.WriteAllText("result.json",
                    JsonConvert.SerializeObject(result, Formatting.Indented),
                    new UTF8Encoding(false));
            }
            catch (System.Exception ex)
            {
                // On failure, still emit a result.json so the workitem produces a
                // diagnosable artifact rather than an empty output.
                var err = new { ok = false, error = ex.Message, stack = ex.StackTrace };
                File.WriteAllText("result.json",
                    JsonConvert.SerializeObject(err, Formatting.Indented),
                    new UTF8Encoding(false));
            }
        }
    }
}
