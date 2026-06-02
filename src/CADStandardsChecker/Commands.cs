// Scaffolded from templates/autocad/CommandClass.cs.
// Capability: validate a DWG against CAD standards (layer naming, text styles, dim styles,
// linetypes, required blocks, drawing hygiene) and write report.json + summary.csv + remediation.md.
//
// Activity inputs (localName → file):
//   inputFile    → active DWG (opened via accoreconsole /i)
//   paramsFile   → "params.json"    (optional: custom JSON rules)
//   standardsFile→ "standards.dws"  (optional: AutoCAD DWS standards file)
//
// Activity outputs:
//   reportJson   → "report.json"
//   summaryCsv   → "summary.csv"
//   remediationMd→ "remediation.md"
using System;
using System.IO;
using System.Linq;
using System.Text;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using Newtonsoft.Json;

[assembly: CommandClass(typeof(CADStandardsChecker.Commands))]
[assembly: ExtensionApplication(null)]

namespace CADStandardsChecker
{
    public class Commands
    {
        // Verb must match <Command Global="CHECKSTANDARDS"> in PackageContents.xml
        // and COMMAND env var in the CI/publish script.
        [CommandMethod("CHECKSTANDARDS", CommandFlags.Modal)]
        public static void CheckStandards()
        {
            try
            {
                Console.WriteLine("[CHECKSTANDARDS] Starting CAD standards compliance check...");

                var wdFiles = Directory.GetFiles(Directory.GetCurrentDirectory());
                Console.WriteLine($"[CHECKSTANDARDS] Working dir files: {string.Join(", ", wdFiles.Select(f => Path.GetFileName(f)))}");

                // Resolve input DWG database (active document or working database)
                var doc = Application.DocumentManager.MdiActiveDocument;
                var db  = doc?.Database ?? HostApplicationServices.WorkingDatabase;

                string fileName = doc?.Name ?? "input.dwg";

                // Optional: custom rules JSON
                RunParams? runParams = null;
                if (File.Exists("params.json"))
                {
                    Console.WriteLine("[CHECKSTANDARDS] Loading params.json...");
                    runParams = JsonConvert.DeserializeObject<RunParams>(File.ReadAllText("params.json"));
                }

                // Optional: .dws standards file
                string? dwsPath = File.Exists("standards.dws") ? "standards.dws" : null;
                if (dwsPath != null)
                    Console.WriteLine("[CHECKSTANDARDS] Found standards.dws — will compare against standards file.");

                // Run compliance check
                var report = ComplianceChecker.Run(db, runParams, dwsPath, fileName);

                Console.WriteLine($"[CHECKSTANDARDS] Overall status: {report.OverallStatus}  " +
                    $"(pass={report.Summary.Pass} warn={report.Summary.Warning} fail={report.Summary.Fail})");

                // ── Write report.json ────────────────────────────────────────
                string reportJson = JsonConvert.SerializeObject(report, Formatting.Indented);
                File.WriteAllText("report.json", reportJson, new UTF8Encoding(false));
                Console.WriteLine("[CHECKSTANDARDS] Wrote report.json");

                // ── Write summary.csv ────────────────────────────────────────
                WriteSummaryCsv(report);
                Console.WriteLine("[CHECKSTANDARDS] Wrote summary.csv");

                // ── Write remediation.md ─────────────────────────────────────
                WriteRemediationMd(report);
                Console.WriteLine("[CHECKSTANDARDS] Wrote remediation.md");

                // Also write result.json (smoke gate checks for this file)
                var gate = new
                {
                    ok            = true,
                    overallStatus = report.OverallStatus,
                    pass          = report.Summary.Pass,
                    warning       = report.Summary.Warning,
                    fail          = report.Summary.Fail,
                };
                File.WriteAllText("result.json", JsonConvert.SerializeObject(gate, Formatting.Indented), new UTF8Encoding(false));

                Console.WriteLine("[CHECKSTANDARDS] Done.");
            }
            catch (System.Exception ex)
            {
                Console.WriteLine($"[CHECKSTANDARDS] ERROR: {ex.Message}");
                var err = new { ok = false, error = ex.Message, stack = ex.StackTrace };
                string errJson = JsonConvert.SerializeObject(err, Formatting.Indented);
                // Always emit both outputs so the workitem has diagnosable artifacts
                File.WriteAllText("result.json",  errJson, new UTF8Encoding(false));
                File.WriteAllText("report.json",  errJson, new UTF8Encoding(false));
                File.WriteAllText("summary.csv",  $"file,overallStatus,pass,warning,fail\ninput.dwg,error,0,0,0\n");
                File.WriteAllText("remediation.md", $"# Compliance Check Error\n\n{ex.Message}\n");
            }
        }

        private static void WriteSummaryCsv(ComplianceReport report)
        {
            var sb = new StringBuilder();
            sb.AppendLine("file,overallStatus,pass,warning,fail");
            sb.AppendLine(
                $"{EscapeCsv(report.File)}," +
                $"{report.OverallStatus}," +
                $"{report.Summary.Pass}," +
                $"{report.Summary.Warning}," +
                $"{report.Summary.Fail}");
            File.WriteAllText("summary.csv", sb.ToString(), new UTF8Encoding(false));
        }

        private static void WriteRemediationMd(ComplianceReport report)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"# CAD Standards Compliance Report");
            sb.AppendLine($"");
            sb.AppendLine($"**File:** {report.File}  ");
            sb.AppendLine($"**Checked:** {report.CheckedAt}  ");
            sb.AppendLine($"**Overall Status:** `{report.OverallStatus}`");
            sb.AppendLine($"");
            sb.AppendLine($"| Result | Count |");
            sb.AppendLine($"|--------|-------|");
            sb.AppendLine($"| Pass   | {report.Summary.Pass}   |");
            sb.AppendLine($"| Warning| {report.Summary.Warning}|");
            sb.AppendLine($"| Fail   | {report.Summary.Fail}   |");
            sb.AppendLine($"");

            var failing = report.Rules.Where(r => r.Result != "pass").ToList();
            if (failing.Count == 0)
            {
                sb.AppendLine("All checks passed. No remediation required.");
                File.WriteAllText("remediation.md", sb.ToString(), new UTF8Encoding(false));
                return;
            }

            sb.AppendLine($"## Issues Requiring Attention");
            sb.AppendLine($"");

            foreach (var rule in failing)
            {
                sb.AppendLine($"### [{rule.Result.ToUpperInvariant()}] {rule.RuleId} — {rule.Category}");
                sb.AppendLine($"");
                sb.AppendLine($"**Message:** {rule.Message}");
                sb.AppendLine($"");
                if (!string.IsNullOrEmpty(rule.Remediation))
                {
                    sb.AppendLine($"**How to fix:** {rule.Remediation}");
                    sb.AppendLine($"");
                }
                if (rule.Offenders.Count > 0)
                {
                    sb.AppendLine($"**Offending objects ({rule.Offenders.Count}):**");
                    sb.AppendLine($"");
                    sb.AppendLine($"| Type | Name | Actual Value | Expected |");
                    sb.AppendLine($"|------|------|-------------|----------|");
                    foreach (var o in rule.Offenders)
                    {
                        sb.AppendLine($"| {o.ObjectType} | `{o.Name}` | {o.Value ?? "-"} | {o.Expected ?? "-"} |");
                    }
                    sb.AppendLine($"");
                }
            }

            File.WriteAllText("remediation.md", sb.ToString(), new UTF8Encoding(false));
        }

        private static string EscapeCsv(string s)
        {
            if (s.Contains(',') || s.Contains('"') || s.Contains('\n'))
                return "\"" + s.Replace("\"", "\"\"") + "\"";
            return s;
        }
    }
}
