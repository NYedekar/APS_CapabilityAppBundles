# APS Capability AppBundles

> **Design Automation plugins that power the WorkflowSkills MCP server.**

This repository contains the C# plugin source code, `.bundle` manifests, and CI/CD pipelines for all custom AppBundles used by [WorkflowSkills MCP](https://git.autodesk.com/yedekan/WorkflowSkills-MCP). Each bundle runs as a cloud compute job on Autodesk Platform Services (APS) Design Automation and is invoked automatically when you ask Claude to process an Autodesk file.

---

## Table of Contents

- [How it fits together](#how-it-fits-together)
- [Bundles](#bundles)
- [Prerequisites](#prerequisites)
- [Deployment — CI (recommended)](#deployment--ci-recommended)
- [Deployment — manual](#deployment--manual)
- [Development](#development)
- [Repository structure](#repository-structure)
- [Support](#support)

---

## How it fits together

```
Claude (Desktop or CLI)
    │
    ▼
WorkflowSkills MCP server          ← git.autodesk.com/yedekan/WorkflowSkills-MCP
    │  submits WorkItems via APS DA API
    ▼
APS Design Automation
    │  runs plugin on cloud engine
    ▼
AppBundle (this repo)              ← your deployed copy, registered under YOUR nickname
    │  writes outputs to OSS
    ▼
MCP server retrieves results → Claude returns them to you
```

The MCP server resolves your APS DA nickname automatically from your credentials at runtime. **You do not need to edit any config files or the capability registry** — just deploy the bundles to your account and they are discovered automatically.

---

## Bundles

| Bundle | Engine | What it does | Outputs |
|--------|--------|-------------|---------|
| `RevitExtractor` | Autodesk.Revit+2026 | Extracts all element parameters and properties from a Revit model | `result.csv` + `result.json` |
| `RevitPDFExport` | Autodesk.Revit+2026 | Exports all views and sheets to PDF | `result.zip` (one PDF per sheet) |
| `RevitIFCExport` | Autodesk.Revit+2026 | Exports a Revit model to IFC | `result.ifc` |
| `RevitSheetListTD` | Autodesk.Revit+2026 | Generates a structured sheet list from a Revit model | `result.json` |
| `AutoCADDrawingMetadataExtractor` | Autodesk.AutoCAD+24_3 | Extracts layers, blocks, layouts, and drawing statistics from a DWG | `result.json` |
| `AutoCADLayerReportTD` | Autodesk.AutoCAD+24_3 | Generates a detailed layer report from a DWG | `result.json` |

> **Note on AutoCAD engine version:** AutoCAD bundles target `+24_3` (AutoCAD 2024, .NET Framework 4.8). The newer `.NET 8` engines (`+25_0`, `+26_0`) silently fail to load `.NET Framework` assemblies — `+24_3` is the proven, officially-documented stack.

---

## Prerequisites

- An **APS application** with Design Automation API enabled — create one at [aps.autodesk.com/myapps](https://aps.autodesk.com/myapps)
- Your **APS DA nickname** — visible at [aps.autodesk.com/myapps](https://aps.autodesk.com/myapps) under your app's DA section
- A repository on GitHub or git.autodesk.com with **Actions** enabled (for CI deployment)

You do **not** need a local .NET / Visual Studio environment — the CI pipeline builds everything on `windows-latest` runners.

---

## Deployment — CI (recommended)

Each bundle has its own GitHub Actions workflow that builds the plugin, packages the `.bundle` folder, and deploys it to APS — all automatically.

### Step 1 — Fork or clone this repo

```bash
# Autodesk internal
git clone https://git.autodesk.com/yedekan/APS_CapabilityAppBundles.git

# or GitHub
git clone https://github.com/NYedekar/APS_CapabilityAppBundles.git
```

### Step 2 — Add repository secrets

In the repo's **Settings → Secrets and variables → Actions**, add:

| Secret | Value |
|--------|-------|
| `APS_CLIENT_ID` | Your APS app Client ID |
| `APS_CLIENT_SECRET` | Your APS app Client Secret |
| `APS_NICKNAME` | Your APS DA nickname |

### Step 3 — Run the workflows

Go to **Actions** and trigger each workflow manually (using **Run workflow**), or push a change to `main` to trigger them automatically.

| Workflow | Deploys |
|----------|---------|
| `Build & Publish AppBundle to APS` | RevitExtractor |
| `Build & Publish RevitPDFExport to APS` | RevitPDFExport |
| `Build & Publish RevitIFCExport to APS` | RevitIFCExport |
| `Build & Publish RevitSheetList to APS` | RevitSheetList |
| `Build & Publish AutoCAD Metadata Extractor` | AutoCADDrawingMetadataExtractor |
| `Build & Publish AutoCADLayerReport to APS` | AutoCADLayerReport |

Each workflow:
1. Builds the C# plugin DLL using MSBuild on a Windows runner
2. Packages it into a `.bundle` zip with the correct structure
3. Uploads the zip as a build artifact (useful for debugging)
4. Registers the AppBundle in your APS account via `scripts/publish-appbundle.js`
5. Creates or updates the Activity definition via `scripts/publish-activity.js`

Once all workflows complete, the bundles are live in your APS account under `<YOUR_NICKNAME>.<BundleName>+prod`.

---

## Deployment — manual

If you prefer to deploy without CI (e.g. from a local machine):

### Prerequisites
- [Node.js](https://nodejs.org) v18+
- [.NET Framework 4.8](https://dotnet.microsoft.com/download/dotnet-framework) + [MSBuild](https://visualstudio.microsoft.com/downloads/#build-tools-for-visual-studio-2022) (Windows only)
- [NuGet CLI](https://www.nuget.org/downloads)

### Build a bundle

```powershell
# Example: RevitExtractor
nuget restore src/RevitExtractor/RevitExtractor.csproj
msbuild src/RevitExtractor/RevitExtractor.csproj `
  /p:Configuration=Release `
  /p:Platform=x64 `
  /p:OutputPath="$PWD\bundle\RevitExtractor.bundle\Contents\"

# Zip it
Compress-Archive -Path bundle\RevitExtractor.bundle -DestinationPath RevitExtractor.zip -Force
```

### Publish to APS

```bash
export APS_CLIENT_ID=your_client_id
export APS_CLIENT_SECRET=your_client_secret
export APS_NICKNAME=your_nickname

node scripts/publish-appbundle.js
node scripts/publish-activity.js
```

### Smoke test

```bash
node scripts/smoke-test-activity.js
```

---

## Development

### Adding a new bundle

1. Create a new C# project under `src/<BundleName>/`
2. Create a `.bundle` folder under `bundle/<BundleName>.bundle/` with:
   - `PackageContents.xml` — DA manifest (see existing bundles for reference)
   - `Contents/` — plugin DLL + `.addin` file (for Revit) or no `.addin` (for AutoCAD)
3. Add a `publish-<name>-appbundle.js` and `publish-<name>-activity.js` under `scripts/`
4. Copy an existing workflow from `.github/workflows/` and update `BUNDLE_NAME` and `ENGINE_VERSION`
5. Register the new capability in the [WorkflowSkills MCP capability registry](https://git.autodesk.com/yedekan/WorkflowSkills-MCP/blob/main/data/capability-registry.json)

### Engine version reference

| Product | Engine ID | Notes |
|---------|-----------|-------|
| Revit 2026 | `Autodesk.Revit+2026` | Current recommended for Revit bundles |
| AutoCAD 2024 | `Autodesk.AutoCAD+24_3` | Use for all AutoCAD bundles (.NET Framework 4.8) |
| Inventor 2024 | `Autodesk.Inventor+24` | |
| 3ds Max 2024 | `Autodesk.3dsMax+2024` | |

---

## Repository structure

```
src/
  RevitExtractor/                   # Extracts all Revit element parameters
  RevitPDFExport/                   # Exports Revit views/sheets to PDF
  RevitIFCExport/                   # Exports Revit model to IFC
  RevitSheetList/                   # Generates Revit sheet list
  AutoCADDrawingMetadataExtractor/  # Extracts DWG metadata (layers, blocks, layouts)
  AutoCADLayerReport/               # Generates DWG layer report

bundle/
  *.bundle/                         # Packaged bundle manifests (DLL contents added by CI)
    PackageContents.xml
    Contents/
      *.addin                       # Revit only

scripts/
  publish-appbundle.js              # Upload + register AppBundle to APS DA
  publish-activity.js               # Create or update Activity definition
  smoke-test-activity.js            # Submit a test WorkItem and verify output

.github/workflows/
  build-and-publish.yml                       # RevitExtractor
  build-publish-revit-pdf-export.yml          # RevitPDFExport
  build-publish-revit-ifc-export.yml          # RevitIFCExport
  build-publish-revit-sheetlist.yml           # RevitSheetList
  build-publish-autocad-metadata-extractor.yml
  build-publish-autocad-layerreport.yml

test/
  sample.rvt                        # Test Revit model
  sample.dwg                        # Test AutoCAD drawing
  condo-skylight.dwg
```

---

## Support

**Neeraj Yedekar**
Product Manager, Autodesk
[neeraj.yedekar@autodesk.com](mailto:neeraj.yedekar@autodesk.com)

Open an issue:
- Autodesk internal: [git.autodesk.com/yedekan/APS_CapabilityAppBundles/issues](https://git.autodesk.com/yedekan/APS_CapabilityAppBundles/issues)
- GitHub: [github.com/NYedekar/APS_CapabilityAppBundles/issues](https://github.com/NYedekar/APS_CapabilityAppBundles/issues)

---

## License

MIT
