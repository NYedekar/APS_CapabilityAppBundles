<#
.SYNOPSIS
    Pre-flight validator for APS Design Automation AppBundle zip files.

.DESCRIPTION
    Checks zip structure, PackageContents.xml schema, and required DLLs
    before uploading to APS. Catches the most common silent-failure causes.

.PARAMETER BundlePath
    Path to the .zip file to validate.

.PARAMETER Product
    Target product: 'autocad', 'revit', or 'inventor'. Affects which checks run.

.EXAMPLE
    .\Validate-Bundle.ps1 -BundlePath .\MyPlugin.zip -Product autocad
    .\Validate-Bundle.ps1 -BundlePath .\MyRevitPlugin.zip -Product revit
    .\Validate-Bundle.ps1 -BundlePath .\MyInventorPlugin.zip -Product inventor
#>

param(
    [Parameter(Mandatory)]
    [string]$BundlePath,

    [Parameter(Mandatory)]
    [ValidateSet('autocad', 'revit', 'inventor')]
    [string]$Product
)

$pass = 0
$fail = 0

function Check($label, $result, $detail = "") {
    if ($result) {
        Write-Host "  [PASS] $label" -ForegroundColor Green
        $script:pass++
    } else {
        Write-Host "  [FAIL] $label$(if ($detail) { ' — ' + $detail })" -ForegroundColor Red
        $script:fail++
    }
}

# ── Load zip entries ──────────────────────────────────────────────────────────
if (-not (Test-Path $BundlePath)) {
    Write-Error "File not found: $BundlePath"
    exit 1
}

Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::OpenRead((Resolve-Path $BundlePath).Path)
$entries = $zip.Entries | ForEach-Object { $_.FullName }

Write-Host "`nValidating: $BundlePath  (product=$Product)`n"

# ── Section 1: zip root layout ────────────────────────────────────────────────
Write-Host "1. Zip root layout"

# Derive the .bundle root from ANY entry under it — do not require an explicit
# directory entry. Compress-Archive emits only file entries (e.g.
# "Foo.bundle/PackageContents.xml"), not a bare "Foo.bundle/" directory entry,
# so matching on a trailing-slash directory entry finds nothing even on a valid zip.
$bundleRoots = @($entries |
    ForEach-Object { if ($_ -match '^([^/]+\.bundle)(/|$)') { $matches[1] } } |
    Select-Object -Unique)
Check "Exactly one .bundle folder at zip root" ($bundleRoots.Count -eq 1) "Found: $($bundleRoots -join ', ')"

# @(...) above forces array context. Without it, a single unique match unwraps to a
# scalar string, and $bundleRoots[0] would return its first CHARACTER ("A"), not the
# whole name — silently breaking every downstream "$bundleRoot/..." path check.
$bundleRoot = if ($bundleRoots.Count -ge 1) { $bundleRoots[0] } else { "UNKNOWN.bundle" }

$manifestEntry = "$bundleRoot/PackageContents.xml"
Check "PackageContents.xml present at bundle root" ($entries -contains $manifestEntry)

$contentsFolder = "$bundleRoot/Contents/"
$inContents = $entries | Where-Object { $_ -like "$contentsFolder*" }
Check "Contents/ folder present" ($inContents.Count -gt 0)

# ── Section 2: DLL presence ───────────────────────────────────────────────────
Write-Host "`n2. Required DLLs in Contents/"

$dlls = $inContents | Where-Object { $_ -match '\.dll$' }
Check "At least one plugin DLL in Contents/" ($dlls.Count -gt 0)

$hasNewtonsoft = $inContents | Where-Object { $_ -match 'Newtonsoft\.Json\.dll$' }
Check "Newtonsoft.Json.dll in Contents/ (required on net48)" ($hasNewtonsoft.Count -gt 0)

if ($Product -eq 'revit') {
    $hasBridge = $inContents | Where-Object { $_ -match 'DesignAutomationBridge\.dll$' }
    Check "DesignAutomationBridge.dll in Contents/ (Revit DA bridge — must ship)" ($hasBridge.Count -gt 0)

    $hasAddin = $inContents | Where-Object { $_ -match '\.addin$' }
    Check ".addin file in Contents/ (Revit auto-load descriptor)" ($hasAddin.Count -gt 0)
}

if ($Product -eq 'inventor') {
    $hasInvAddin = $inContents | Where-Object { $_ -match '\.Inventor\.addin$' }
    Check ".Inventor.addin file in Contents/ (Inventor auto-load descriptor)" ($hasInvAddin.Count -gt 0)
}

# ── Section 3: PackageContents.xml content checks ─────────────────────────────
Write-Host "`n3. PackageContents.xml content"

$manifestXml = $null
try {
    $entry = $zip.GetEntry($manifestEntry)
    if ($entry) {
        $reader = New-Object System.IO.StreamReader($entry.Open())
        $manifestXml = [xml]$reader.ReadToEnd()
        $reader.Close()
    }
} catch {
    Write-Host "  [WARN] Could not parse PackageContents.xml: $_" -ForegroundColor Yellow
}

if ($manifestXml) {
    # ProductCode (GUID) present — AutoCAD and Inventor carry one in PackageContents.xml.
    # Revit does NOT: it registers through the .addin AddInId, and working Revit bundles
    # have no ProductCode attribute, so requiring it for Revit is a false positive.
    if ($Product -ne 'revit') {
        $productCode = $manifestXml.ApplicationPackage.ProductCode
        Check "ProductCode (GUID) present" ($productCode -and $productCode -match '^\{?[0-9A-Fa-f-]{32,36}\}?$')
    }

    if ($Product -eq 'autocad') {
        # No SupportedLocales — this causes silent skip when worker locale doesn't match
        $supportedLocales = $manifestXml.ApplicationPackage.Components.RuntimeRequirements.SupportedLocales
        Check "No SupportedLocales (omitting = correct for DA)" ([string]::IsNullOrEmpty($supportedLocales)) "SupportedLocales='$supportedLocales' causes silent bundle skip"

        # LoadOnAutoCADStartup + LoadOnCommandInvocation
        $loadFlags = $manifestXml.ApplicationPackage.Components.ComponentEntry
        $startup = $loadFlags | Where-Object { $_.LoadOnAutoCADStartup -eq 'True' }
        $cmd = $loadFlags | Where-Object { $_.LoadOnCommandInvocation -eq 'True' }
        Check "LoadOnAutoCADStartup=True on at least one ComponentEntry" ($startup.Count -gt 0)
        Check "LoadOnCommandInvocation=True on at least one ComponentEntry" ($cmd.Count -gt 0)

        # Command verbs defined
        $commands = $manifestXml.SelectNodes("//Command")
        Check "At least one <Command> element defined" ($commands.Count -gt 0)

        # Warn on known built-in verb collisions
        $builtins = @('CHECKSTANDARDS','AUDIT','PURGE','LAYER','STANDARDS','MATCHPROP','RECTANG','CIRCLE','LINE')
        foreach ($cmd in $commands) {
            $verb = $cmd.Global
            if ($verb -in $builtins) {
                Write-Host "  [WARN] Command verb '$verb' is a known AutoCAD built-in — rename to avoid silent clash" -ForegroundColor Yellow
                $script:fail++
            }
        }

    } elseif ($Product -eq 'revit') {
        # ModuleName points to .addin (not .dll)
        $moduleName = $manifestXml.ApplicationPackage.Components.ComponentEntry.ModuleName
        Check "ModuleName points to .addin file (not .dll)" ($moduleName -match '\.addin$') "ModuleName='$moduleName'"

        # RuntimeRequirements inside Components
        $rtInComponents = $manifestXml.SelectSingleNode("//Components/RuntimeRequirements")
        $rtTopLevel = $manifestXml.ApplicationPackage.RuntimeRequirements
        Check "RuntimeRequirements is inside <Components> (not top-level)" ($rtInComponents -ne $null -and $rtTopLevel -eq $null)

        # SeriesMin/SeriesMax present
        $seriesMin = $manifestXml.SelectSingleNode("//RuntimeRequirements/@SeriesMin")
        $seriesMax = $manifestXml.SelectSingleNode("//RuntimeRequirements/@SeriesMax")
        Check "SeriesMin present" ($seriesMin -ne $null)
        Check "SeriesMax present" ($seriesMax -ne $null)

    } elseif ($Product -eq 'inventor') {
        # ModuleName points to the .Inventor.addin descriptor (not the .dll)
        $moduleName = $manifestXml.ApplicationPackage.Components.ComponentEntry.ModuleName
        Check "ModuleName points to .addin file (not .dll)" ($moduleName -match '\.addin$') "ModuleName='$moduleName'"

        # Platform should be Inventor. RuntimeRequirements sits directly under <Components>
        # (sibling of ComponentEntry), not nested inside ComponentEntry — same as the
        # AutoCAD SupportedLocales check above.
        $platform = $manifestXml.ApplicationPackage.Components.RuntimeRequirements.Platform
        Check "RuntimeRequirements Platform='Inventor'" ($platform -eq 'Inventor') "Platform='$platform'"
    }
}

$zip.Dispose()

# ── Summary ───────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "─────────────────────────────────────────"
if ($fail -eq 0) {
    Write-Host "  RESULT: ALL $pass checks passed" -ForegroundColor Green
} else {
    Write-Host "  RESULT: $pass passed, $fail failed — fix issues before publishing" -ForegroundColor Red
}
Write-Host "─────────────────────────────────────────`n"

exit $fail
