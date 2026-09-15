# SharePoint Classic-to-Modern Migration Runbook v2

## Complete Script Mapping, Guardrails, Validation, and Master Orchestrator

This runbook provides a reproducible Windows-based PowerShell process for rebuilding a Classic SharePoint experience as a Modern SharePoint page/site and migrating the folder hierarchy and file content into the Modern destination.

The runbook is intentionally divided into complete `.ps1` files. Each script below is a full script, not a partial snippet. A new specialist should be able to create the files exactly as shown, run them one stage at a time, review the output, and troubleshoot before proceeding.

> **Safety principle:** Discover -> classify -> preserve -> build -> pilot -> validate -> migrate -> validate -> cut over.
>
> **Source preservation rule:** The migration scripts do not delete or overwrite the Classic source.

> **Environment note:** These examples are written for SharePoint Online using PnP PowerShell. Authentication and some cmdlet behavior can differ in Commercial Microsoft 365, GCC High, DoD, or SharePoint Server/on-premises environments. Validate your tenant's approved authentication method before production use.

---

## 1. Recommended Windows Folder Structure

Create the following structure on the migration workstation:

```text
C:\SPModernization
|
|-- config.psd1
|-- Invoke-SharePointModernization.ps1
|
|-- Scripts
|   |-- 00-Preflight.ps1
|   |-- 01-Inventory.ps1
|   |-- 02-Classify-And-Backup-Page.ps1
|   |-- 03-Build-Modern-Target.ps1
|   |-- 04-Migrate-Pilot.ps1
|   |-- 05-Validate-Pilot.ps1
|   |-- 06-Migrate-All.ps1
|   |-- 07-Validate-All.ps1
|   `-- 08-Cutover.ps1
|
|-- Reports
|-- Backup
|-- Logs
`-- Temp
```

For the first migration, run scripts `00` through `08` individually. Use the master script only after the individual stages have been tested and understood.

---

# 2. `config.psd1`

## Purpose

The configuration file centralizes site URLs, page names, library names, pilot folder, authentication settings, and environment safeguards. Do not hard-code these values throughout every script.

Create:

```powershell
notepad C:\SPModernization\config.psd1
```

Paste:

```powershell
@{

    # ============================================================
    # SOURCE / CLASSIC SHAREPOINT
    # ============================================================

    SourceSiteUrl = "https://YOUR-SOURCE-SHAREPOINT-SITE"

    # Site-relative page path.
    # Example:
    # SitePages/Home.aspx
    # Pages/default.aspx
    SourcePage = "SitePages/Home.aspx"

    # Display name of the source document library.
    SourceLibrary = "Shared Documents"


    # ============================================================
    # TARGET / MODERN SHAREPOINT
    # ============================================================

    TargetSiteUrl = "https://YOUR-TARGET-SHAREPOINT-SITE"

    # Do NOT include .aspx here.
    TargetPageName = "Military-Personnel-Flight-Modern"

    TargetLibrary = "MPF Documents"


    # ============================================================
    # PILOT MIGRATION
    # ============================================================

    # Start with one relatively small folder.
    PilotFolder = "CSS Training"


    # ============================================================
    # AUTHENTICATION
    # ============================================================

    # If your SharePoint administrator gave you an approved
    # Entra Application / Client ID, enter it here.
    #
    # Otherwise leave this empty and the scripts will attempt
    # ordinary interactive authentication.
    #
    ClientId = ""


    # ============================================================
    # SAFETY
    # ============================================================

    # TEST or PROD
    Environment = "TEST"

}
```

Do not place passwords, CAC PINs, client secrets, certificate private keys, or other credentials in this file.

---

# 3. `00-Preflight.ps1`

## Purpose

This script does not change SharePoint. It verifies the workstation, PowerShell, PnP.PowerShell, configuration, working directories, and production safety settings.

Create:

```powershell
notepad C:\SPModernization\Scripts\00-Preflight.ps1
```

Paste:

```powershell
# ================================================================
# 00-Preflight.ps1
#
# PURPOSE:
# Validate the local migration workstation before connecting
# to or modifying SharePoint.
# ================================================================

$ErrorActionPreference = "Stop"

$RootPath   = "C:\SPModernization"
$ConfigPath = Join-Path $RootPath "config.psd1"

Write-Host ""
Write-Host "============================================================"
Write-Host " SHAREPOINT MODERNIZATION - PREFLIGHT"
Write-Host "============================================================"
Write-Host ""

# ----------------------------------------------------------------
# 1. Check configuration file
# ----------------------------------------------------------------

if (-not (Test-Path $ConfigPath)) {
    throw "Configuration file was not found: $ConfigPath"
}

$Config = Import-PowerShellDataFile $ConfigPath

Write-Host "[PASS] Configuration file found."

# ----------------------------------------------------------------
# 2. Check PowerShell version
# ----------------------------------------------------------------

$PowerShellVersion = $PSVersionTable.PSVersion

Write-Host "PowerShell Version: $PowerShellVersion"

if ($PowerShellVersion.Major -lt 7) {

    Write-Warning @"
PowerShell 7 or later is recommended for current PnP.PowerShell use.

Current version:
$PowerShellVersion
"@

}
else {

    Write-Host "[PASS] PowerShell version is suitable." `
        -ForegroundColor Green

}

# ----------------------------------------------------------------
# 3. Check PnP.PowerShell
# ----------------------------------------------------------------

$PnPModule = Get-Module PnP.PowerShell `
    -ListAvailable |
    Sort-Object Version -Descending |
    Select-Object -First 1

if (-not $PnPModule) {

    throw @"
PnP.PowerShell is not installed.

Install it with:

Install-Module PnP.PowerShell -Scope CurrentUser
"@

}

Write-Host "[PASS] PnP.PowerShell installed."
Write-Host "PnP Version: $($PnPModule.Version)"

# ----------------------------------------------------------------
# 4. Verify required configuration values
# ----------------------------------------------------------------

$RequiredSettings = @(
    "SourceSiteUrl",
    "SourcePage",
    "SourceLibrary",
    "TargetSiteUrl",
    "TargetPageName",
    "TargetLibrary",
    "PilotFolder",
    "Environment"
)

foreach ($Setting in $RequiredSettings) {

    if (
        -not $Config.ContainsKey($Setting) -or
        [string]::IsNullOrWhiteSpace(
            [string]$Config[$Setting]
        )
    ) {

        throw "Required configuration setting is missing: $Setting"

    }

}

Write-Host "[PASS] Required configuration values are present."

# ----------------------------------------------------------------
# 5. Ensure working directories exist
# ----------------------------------------------------------------

$Directories = @(
    "$RootPath\Scripts",
    "$RootPath\Reports",
    "$RootPath\Backup",
    "$RootPath\Logs",
    "$RootPath\Temp"
)

foreach ($Directory in $Directories) {

    if (-not (Test-Path $Directory)) {

        New-Item `
            -ItemType Directory `
            -Path $Directory `
            -Force |
            Out-Null

        Write-Host "Created directory: $Directory"

    }

}

Write-Host "[PASS] Working directories verified."

# ----------------------------------------------------------------
# 6. Display migration configuration
# ----------------------------------------------------------------

Write-Host ""
Write-Host "---------------- CONFIGURATION ----------------"
Write-Host "Source Site:    $($Config.SourceSiteUrl)"
Write-Host "Source Page:    $($Config.SourcePage)"
Write-Host "Source Library: $($Config.SourceLibrary)"
Write-Host ""
Write-Host "Target Site:    $($Config.TargetSiteUrl)"
Write-Host "Target Page:    $($Config.TargetPageName)"
Write-Host "Target Library: $($Config.TargetLibrary)"
Write-Host ""
Write-Host "Pilot Folder:   $($Config.PilotFolder)"
Write-Host "Environment:    $($Config.Environment)"
Write-Host "------------------------------------------------"
Write-Host ""

# ----------------------------------------------------------------
# 7. Production safety guardrail
# ----------------------------------------------------------------

if ($Config.Environment.ToUpper() -eq "PROD") {

    Write-Warning "THIS CONFIGURATION IS MARKED AS PRODUCTION."

    $Confirmation = Read-Host `
        "Type PRODUCTION exactly to acknowledge"

    if ($Confirmation -ne "PRODUCTION") {

        throw "Production operation cancelled."

    }

}

Write-Host ""
Write-Host "============================================================"
Write-Host " PREFLIGHT PASSED"
Write-Host "============================================================"
Write-Host ""
Write-Host "Next step:"
Write-Host ".\Scripts\01-Inventory.ps1"
Write-Host ""
```

Run:

```powershell
cd C:\SPModernization
.\Scripts\00-Preflight.ps1
```

Stop and read the output before continuing.

---

# 4. `01-Inventory.ps1`

## Purpose

This script is read-only. It inventories the source site, lists/libraries, navigation, folders, and files.

Create:

```powershell
notepad C:\SPModernization\Scripts\01-Inventory.ps1
```

Paste:

```powershell
# ================================================================
# 01-Inventory.ps1
#
# PURPOSE:
# Perform a read-only inventory of the Classic/source SharePoint
# environment.
#
# THIS SCRIPT DOES NOT MODIFY SHAREPOINT.
# ================================================================

$ErrorActionPreference = "Stop"

$RootPath   = "C:\SPModernization"
$ConfigPath = Join-Path $RootPath "config.psd1"
$ReportPath = Join-Path $RootPath "Reports"

$Config = Import-PowerShellDataFile $ConfigPath

function Connect-ToPnPSite {

    param (
        [Parameter(Mandatory)]
        [string]$Url
    )

    if (
        $Config.ContainsKey("ClientId") -and
        -not [string]::IsNullOrWhiteSpace($Config.ClientId)
    ) {

        return Connect-PnPOnline `
            -Url $Url `
            -ClientId $Config.ClientId `
            -Interactive `
            -ReturnConnection

    }
    else {

        return Connect-PnPOnline `
            -Url $Url `
            -Interactive `
            -ReturnConnection

    }

}

Write-Host ""
Write-Host "============================================================"
Write-Host " SOURCE SHAREPOINT INVENTORY"
Write-Host "============================================================"
Write-Host ""

Write-Host "Connecting to source SharePoint..."

$SourceConnection = Connect-ToPnPSite `
    -Url $Config.SourceSiteUrl

Write-Host "[PASS] Connected."

$Web = Get-PnPWeb `
    -Includes Title,Url,ServerRelativeUrl `
    -Connection $SourceConnection

$Web |
    Select-Object `
        Title,
        Url,
        ServerRelativeUrl |
    Export-Csv `
        "$ReportPath\Source-Web.csv" `
        -NoTypeInformation

Write-Host "[PASS] Site information exported."

$Lists = Get-PnPList `
    -Connection $SourceConnection

$Lists |
    Where-Object {
        $_.Hidden -eq $false
    } |
    Select-Object `
        Title,
        BaseTemplate,
        ItemCount,
        DefaultViewUrl |
    Export-Csv `
        "$ReportPath\Source-Lists.csv" `
        -NoTypeInformation

Write-Host "[PASS] Lists and libraries exported."

try {

    $Navigation = Get-PnPNavigationNode `
        -Location QuickLaunch `
        -Connection $SourceConnection

    $Navigation |
        Select-Object `
            Title,
            Url |
        Export-Csv `
            "$ReportPath\Source-Navigation.csv" `
            -NoTypeInformation

    Write-Host "[PASS] Quick Launch navigation exported."

}
catch {

    Write-Warning `
        "Quick Launch navigation could not be inventoried: $($_.Exception.Message)"

}

$SourceList = Get-PnPList `
    -Identity $Config.SourceLibrary `
    -Includes RootFolder `
    -Connection $SourceConnection

if (-not $SourceList) {

    throw "Source library not found: $($Config.SourceLibrary)"

}

$LibraryRootServerRelative =
    $SourceList.RootFolder.ServerRelativeUrl.TrimEnd('/')

$WebServerRelative =
    $Web.ServerRelativeUrl.TrimEnd('/')

if ([string]::IsNullOrWhiteSpace($WebServerRelative)) {
    $WebServerRelative = "/"
}

if ($WebServerRelative -eq "/") {

    $LibrarySiteRelative =
        $LibraryRootServerRelative.TrimStart('/')

}
else {

    $LibrarySiteRelative =
        $LibraryRootServerRelative.Substring(
            $WebServerRelative.Length
        ).TrimStart('/')

}

Write-Host ""
Write-Host "Source Library Root:"
Write-Host $LibraryRootServerRelative

Write-Host ""
Write-Host "Library Site Relative Path:"
Write-Host $LibrarySiteRelative

Write-Host ""
Write-Host "Scanning folders..."

$Folders = @(
    Get-PnPFolderItem `
        -FolderSiteRelativeUrl $LibrarySiteRelative `
        -ItemType Folder `
        -Recursive `
        -Connection $SourceConnection
)

$Folders |
    Select-Object `
        Name,
        ServerRelativeUrl |
    Export-Csv `
        "$ReportPath\Source-Folders.csv" `
        -NoTypeInformation

Write-Host "[PASS] Folders inventoried: $($Folders.Count)"

Write-Host "Scanning files..."

$Files = @(
    Get-PnPFolderItem `
        -FolderSiteRelativeUrl $LibrarySiteRelative `
        -ItemType File `
        -Recursive `
        -Connection $SourceConnection
)

$Files |
    Select-Object `
        Name,
        ServerRelativeUrl,
        Length |
    Export-Csv `
        "$ReportPath\Source-Files.csv" `
        -NoTypeInformation

Write-Host "[PASS] Files inventoried: $($Files.Count)"

Write-Host ""
Write-Host "============================================================"
Write-Host " INVENTORY COMPLETE"
Write-Host "============================================================"
Write-Host ""
Write-Host "Folders: $($Folders.Count)"
Write-Host "Files:   $($Files.Count)"
Write-Host ""
Write-Host "Review these files before continuing:"
Write-Host ""
Write-Host "$ReportPath\Source-Web.csv"
Write-Host "$ReportPath\Source-Lists.csv"
Write-Host "$ReportPath\Source-Navigation.csv"
Write-Host "$ReportPath\Source-Folders.csv"
Write-Host "$ReportPath\Source-Files.csv"
Write-Host ""
Write-Host "DO NOT continue until the inventory makes sense."
Write-Host ""
```

Run:

```powershell
.\Scripts\01-Inventory.ps1
```

Then inspect the CSV files under `C:\SPModernization\Reports`.

---

# 5. `02-Classify-And-Backup-Page.ps1`

## Purpose

This script classifies the source page as Modern, Classic, or Unknown.

Decision logic:

```text
Modern  -> Do not convert
Classic -> Preserve source and prepare a new Modern replacement
Unknown -> Stop and investigate manually
```

Create:

```powershell
notepad C:\SPModernization\Scripts\02-Classify-And-Backup-Page.ps1
```

Paste:

```powershell
# ================================================================
# 02-Classify-And-Backup-Page.ps1
#
# PURPOSE:
# Determine whether the source SharePoint page is Classic,
# Modern, or Unknown.
#
# If Classic, create a reference backup.
#
# THE SOURCE PAGE IS NOT DELETED OR OVERWRITTEN.
# ================================================================

$ErrorActionPreference = "Stop"

$RootPath   = "C:\SPModernization"
$ConfigPath = Join-Path $RootPath "config.psd1"
$BackupRoot = Join-Path $RootPath "Backup"
$ReportPath = Join-Path $RootPath "Reports"

$Config = Import-PowerShellDataFile $ConfigPath

function Connect-ToPnPSite {

    param (
        [Parameter(Mandatory)]
        [string]$Url
    )

    if (
        $Config.ContainsKey("ClientId") -and
        -not [string]::IsNullOrWhiteSpace($Config.ClientId)
    ) {

        return Connect-PnPOnline `
            -Url $Url `
            -ClientId $Config.ClientId `
            -Interactive `
            -ReturnConnection

    }
    else {

        return Connect-PnPOnline `
            -Url $Url `
            -Interactive `
            -ReturnConnection

    }

}

Write-Host ""
Write-Host "============================================================"
Write-Host " PAGE CLASSIFICATION"
Write-Host "============================================================"
Write-Host ""

$SourceConnection = Connect-ToPnPSite `
    -Url $Config.SourceSiteUrl

$PageName = Split-Path `
    $Config.SourcePage `
    -Leaf

$PageType = "Unknown"
$ModernPage = $null
$PageItem = $null

try {

    $ModernPage = Get-PnPPage `
        -Identity $PageName `
        -Connection $SourceConnection `
        -ErrorAction Stop

    if ($ModernPage) {

        $PageType = "Modern"

    }

}
catch {

    Write-Host `
        "Page was not recognized as a Modern client-side page."

}

try {

    $PageItem = Get-PnPFile `
        -Url $Config.SourcePage `
        -AsListItem `
        -Connection $SourceConnection `
        -ErrorAction Stop

    if ($PageType -ne "Modern") {

        $Fields = $PageItem.FieldValues

        $HasWikiContent =
            $Fields.ContainsKey("WikiField") -and
            -not [string]::IsNullOrWhiteSpace(
                [string]$Fields["WikiField"]
            )

        $HasPublishingContent =
            $Fields.ContainsKey("PublishingPageContent") -and
            -not [string]::IsNullOrWhiteSpace(
                [string]$Fields["PublishingPageContent"]
            )

        if ($HasWikiContent -or $HasPublishingContent) {

            $PageType = "Classic"

        }

    }

}
catch {

    throw @"
Unable to inspect source page:

$($Config.SourcePage)

Error:
$($_.Exception.Message)
"@

}

$ClassificationResult = [PSCustomObject]@{

    SourceSite = $Config.SourceSiteUrl
    SourcePage = $Config.SourcePage
    PageType   = $PageType
    CheckedAt  = Get-Date

}

$ClassificationResult |
    Export-Csv `
        "$ReportPath\Page-Classification.csv" `
        -NoTypeInformation

switch ($PageType) {

    "Modern" {

        Write-Host ""
        Write-Host "RESULT: MODERN PAGE" `
            -ForegroundColor Green

        Write-Host ""
        Write-Host @"
The source page is already recognized as a Modern SharePoint page.

No page conversion will be attempted.

The content migration can still continue if the target site/library
is different from the source.
"@

    }

    "Classic" {

        Write-Host ""
        Write-Host "RESULT: CLASSIC PAGE" `
            -ForegroundColor Yellow

        Write-Host ""
        Write-Host "Creating a reference backup..."

        $Timestamp = Get-Date `
            -Format "yyyyMMdd-HHmmss"

        $BackupDirectory =
            Join-Path `
                $BackupRoot `
                "ClassicPage-$Timestamp"

        New-Item `
            -ItemType Directory `
            -Path $BackupDirectory `
            -Force |
            Out-Null

        Get-PnPFile `
            -Url $Config.SourcePage `
            -Path $BackupDirectory `
            -FileName $PageName `
            -AsFile `
            -Force `
            -Connection $SourceConnection

        if ($PageItem) {

            $PageItem.FieldValues |
                ConvertTo-Json -Depth 10 |
                Set-Content `
                    "$BackupDirectory\Page-Metadata.json"

        }

        Write-Host ""
        Write-Host "Reference backup created:"
        Write-Host $BackupDirectory

        Write-Host ""
        Write-Host @"
IMPORTANT:

This backup is only a reference copy of the page and metadata.

It is NOT a complete SharePoint disaster-recovery backup.

The primary rollback protection is that the original Classic site
and Classic page remain untouched.
"@

    }

    default {

        Write-Host ""
        Write-Host "RESULT: UNKNOWN PAGE TYPE" `
            -ForegroundColor Red

        throw @"
The script cannot safely classify the page.

MIGRATION STOPPED.

Do not assume the page is Classic.

A SharePoint administrator should inspect the page manually before
continuing.
"@

    }

}

Write-Host ""
Write-Host "Classification report:"
Write-Host "$ReportPath\Page-Classification.csv"
Write-Host ""
```

For the first real-world migration, stop after this script and review the inventory and classification results before designing the target.

---

# 6. `03-Build-Modern-Target.ps1`

## Purpose

This script creates the target Modern document library and target Modern page only when they do not already exist.

Create:

```powershell
notepad C:\SPModernization\Scripts\03-Build-Modern-Target.ps1
```

Paste:

```powershell
# ================================================================
# 03-Build-Modern-Target.ps1
#
# PURPOSE:
# Create the Modern destination page and document library.
#
# Guardrails prevent duplicate creation.
# ================================================================

$ErrorActionPreference = "Stop"

$RootPath   = "C:\SPModernization"
$ConfigPath = Join-Path $RootPath "config.psd1"

$Config = Import-PowerShellDataFile $ConfigPath

function Connect-ToPnPSite {

    param (
        [Parameter(Mandatory)]
        [string]$Url
    )

    if (
        $Config.ContainsKey("ClientId") -and
        -not [string]::IsNullOrWhiteSpace($Config.ClientId)
    ) {

        return Connect-PnPOnline `
            -Url $Url `
            -ClientId $Config.ClientId `
            -Interactive `
            -ReturnConnection

    }
    else {

        return Connect-PnPOnline `
            -Url $Url `
            -Interactive `
            -ReturnConnection

    }

}

Write-Host ""
Write-Host "============================================================"
Write-Host " BUILD MODERN TARGET"
Write-Host "============================================================"
Write-Host ""

$TargetConnection = Connect-ToPnPSite `
    -Url $Config.TargetSiteUrl

$TargetLibrary = $null

try {

    $TargetLibrary = Get-PnPList `
        -Identity $Config.TargetLibrary `
        -Connection $TargetConnection `
        -ErrorAction Stop

}
catch {

    $TargetLibrary = $null

}

if ($TargetLibrary) {

    Write-Host `
        "[SKIP] Target library already exists: $($Config.TargetLibrary)" `
        -ForegroundColor Yellow

}
else {

    Write-Host `
        "Creating target document library: $($Config.TargetLibrary)"

    New-PnPList `
        -Title $Config.TargetLibrary `
        -Template DocumentLibrary `
        -Connection $TargetConnection

    Write-Host `
        "[PASS] Target library created." `
        -ForegroundColor Green

}

$TargetPageFile =
    "$($Config.TargetPageName).aspx"

$ExistingPage = $null

try {

    $ExistingPage = Get-PnPPage `
        -Identity $TargetPageFile `
        -Connection $TargetConnection `
        -ErrorAction Stop

}
catch {

    $ExistingPage = $null

}

if ($ExistingPage) {

    Write-Host `
        "[SKIP] Modern page already exists: $TargetPageFile" `
        -ForegroundColor Yellow

}
else {

    Write-Host "Creating Modern page: $TargetPageFile"

    Add-PnPPage `
        -Name $Config.TargetPageName `
        -LayoutType Article `
        -Connection $TargetConnection

    Write-Host `
        "[PASS] Modern page created." `
        -ForegroundColor Green

}

if (-not $ExistingPage) {

    try {

        Add-PnPPageSection `
            -Page $TargetPageFile `
            -SectionTemplate OneColumn `
            -Order 1 `
            -Connection $TargetConnection

        Add-PnPPageTextPart `
            -Page $TargetPageFile `
            -Text @"
<h1>Military Personnel Flight</h1>
<p>Personnel services, resources, documents, and information.</p>
"@ `
            -Section 1 `
            -Column 1 `
            -Connection $TargetConnection

        Write-Host `
            "[PASS] Initial page content added."

    }
    catch {

        Write-Warning @"
The Modern page was created, but the initial text/section could not
be added automatically.

You can configure the visual layout manually in the Modern
SharePoint page editor.

Error:
$($_.Exception.Message)
"@

    }

}

Write-Host ""
Write-Host "============================================================"
Write-Host " MODERN TARGET BUILD COMPLETE"
Write-Host "============================================================"
Write-Host ""
Write-Host "STOP HERE and visually inspect the target."
Write-Host ""
Write-Host "Target Site:"
Write-Host $Config.TargetSiteUrl
Write-Host ""
Write-Host "Target Library:"
Write-Host $Config.TargetLibrary
Write-Host ""
Write-Host "Target Page:"
Write-Host $TargetPageFile
Write-Host ""
```

---

# 7. `04-Migrate-Pilot.ps1`

## Purpose

This migrates only one configured pilot folder. The pilot is used to validate folder recreation, file transfer, access, and overall migration behavior before attempting a full migration.

Create:

```powershell
notepad C:\SPModernization\Scripts\04-Migrate-Pilot.ps1
```

Paste:

```powershell
# ================================================================
# 04-Migrate-Pilot.ps1
#
# PURPOSE:
# Migrate ONE pilot folder from the Classic/source library
# to the Modern/target library.
#
# Folder hierarchy is recreated automatically.
#
# This basic migration preserves files and folder hierarchy.
# Do not assume version history, author information, timestamps,
# approvals, retention labels, custom metadata, or permissions
# are preserved by this process.
# ================================================================

$ErrorActionPreference = "Stop"

$RootPath   = "C:\SPModernization"
$ConfigPath = Join-Path $RootPath "config.psd1"
$LogPath    = Join-Path $RootPath "Logs"
$TempPath   = Join-Path $RootPath "Temp"

$Config = Import-PowerShellDataFile $ConfigPath

function Connect-ToPnPSite {

    param (
        [Parameter(Mandatory)]
        [string]$Url
    )

    if (
        $Config.ContainsKey("ClientId") -and
        -not [string]::IsNullOrWhiteSpace($Config.ClientId)
    ) {

        return Connect-PnPOnline `
            -Url $Url `
            -ClientId $Config.ClientId `
            -Interactive `
            -ReturnConnection

    }
    else {

        return Connect-PnPOnline `
            -Url $Url `
            -Interactive `
            -ReturnConnection

    }

}

function Invoke-WithRetry {

    param (

        [Parameter(Mandatory)]
        [scriptblock]$Operation,

        [int]$MaximumAttempts = 4

    )

    $Attempt = 1

    while ($Attempt -le $MaximumAttempts) {

        try {

            return & $Operation

        }
        catch {

            if ($Attempt -ge $MaximumAttempts) {
                throw
            }

            $DelaySeconds =
                [Math]::Pow(2, $Attempt)

            Write-Warning `
                "Attempt $Attempt failed. Retrying in $DelaySeconds seconds."

            Start-Sleep `
                -Seconds $DelaySeconds

            $Attempt++

        }

    }

}

Write-Host ""
Write-Host "============================================================"
Write-Host " PILOT MIGRATION"
Write-Host "============================================================"
Write-Host ""

Write-Host "Pilot Folder: $($Config.PilotFolder)"
Write-Host ""

$SourceConnection = Connect-ToPnPSite `
    -Url $Config.SourceSiteUrl

$TargetConnection = Connect-ToPnPSite `
    -Url $Config.TargetSiteUrl

$SourceWeb = Get-PnPWeb `
    -Includes ServerRelativeUrl `
    -Connection $SourceConnection

$SourceList = Get-PnPList `
    -Identity $Config.SourceLibrary `
    -Includes RootFolder `
    -Connection $SourceConnection

$SourceRoot =
    $SourceList.RootFolder.ServerRelativeUrl.TrimEnd('/')

$SourceWebRoot =
    $SourceWeb.ServerRelativeUrl.TrimEnd('/')

if ([string]::IsNullOrWhiteSpace($SourceWebRoot)) {
    $SourceWebRoot = "/"
}

if ($SourceWebRoot -eq "/") {

    $SourceLibrarySiteRelative =
        $SourceRoot.TrimStart('/')

}
else {

    $SourceLibrarySiteRelative =
        $SourceRoot.Substring(
            $SourceWebRoot.Length
        ).TrimStart('/')

}

$AllFiles = @(
    Get-PnPFolderItem `
        -FolderSiteRelativeUrl $SourceLibrarySiteRelative `
        -ItemType File `
        -Recursive `
        -Connection $SourceConnection
)

$PilotPrefix =
    "$($Config.PilotFolder.Trim('/'))/"

$PilotFiles = @(
    $AllFiles |
        Where-Object {

            $RelativePath =
                $_.ServerRelativeUrl.Substring(
                    $SourceRoot.Length
                ).TrimStart('/')

            $RelativePath.StartsWith(
                $PilotPrefix,
                [System.StringComparison]::OrdinalIgnoreCase
            )

        }
)

if ($PilotFiles.Count -eq 0) {

    throw @"
No files were found for pilot folder:

$($Config.PilotFolder)

Verify the actual folder structure before continuing.
"@

}

Write-Host "Pilot files discovered: $($PilotFiles.Count)"
Write-Host ""

$Confirmation =
    Read-Host "Type PILOT to begin pilot migration"

if ($Confirmation -ne "PILOT") {

    throw "Pilot migration cancelled."

}

$MigrationLog = @()
$Counter = 0

foreach ($File in $PilotFiles) {

    $Counter++
    $LocalTempFile = $null

    try {

        $RelativePath =
            $File.ServerRelativeUrl.Substring(
                $SourceRoot.Length
            ).TrimStart('/')

        Write-Host ""
        Write-Host "[$Counter/$($PilotFiles.Count)] $RelativePath"

        $LastSlash =
            $RelativePath.LastIndexOf('/')

        if ($LastSlash -gt 0) {

            $RelativeFolder =
                $RelativePath.Substring(
                    0,
                    $LastSlash
                )

        }
        else {

            $RelativeFolder = ""

        }

        $TargetFolderSiteRelative =
            $Config.TargetLibrary

        if (-not [string]::IsNullOrWhiteSpace($RelativeFolder)) {

            $TargetFolderSiteRelative =
                "$($Config.TargetLibrary)/$RelativeFolder"

        }

        $ResolvedTargetFolder =
            Resolve-PnPFolder `
                -SiteRelativePath $TargetFolderSiteRelative `
                -Connection $TargetConnection

        $Extension =
            [System.IO.Path]::GetExtension($File.Name)

        $TempName =
            "$([guid]::NewGuid().ToString())$Extension"

        $LocalTempFile =
            Join-Path $TempPath $TempName

        Invoke-WithRetry {

            Get-PnPFile `
                -Url $File.ServerRelativeUrl `
                -Path $TempPath `
                -FileName $TempName `
                -AsFile `
                -Force `
                -Connection $SourceConnection

        } | Out-Null

        Invoke-WithRetry {

            Add-PnPFile `
                -Path $LocalTempFile `
                -Folder $ResolvedTargetFolder.ServerRelativeUrl `
                -NewFileName $File.Name `
                -Connection $TargetConnection

        } | Out-Null

        $MigrationLog += [PSCustomObject]@{

            RelativePath = $RelativePath
            Source       = $File.ServerRelativeUrl
            Status       = "SUCCESS"
            Error        = ""

        }

        Write-Host "SUCCESS" `
            -ForegroundColor Green

    }
    catch {

        $MigrationLog += [PSCustomObject]@{

            RelativePath = $RelativePath
            Source       = $File.ServerRelativeUrl
            Status       = "FAILED"
            Error        = $_.Exception.Message

        }

        Write-Warning `
            "FAILED: $($_.Exception.Message)"

    }
    finally {

        if (
            $LocalTempFile -and
            (Test-Path $LocalTempFile)
        ) {

            Remove-Item `
                $LocalTempFile `
                -Force `
                -ErrorAction SilentlyContinue

        }

    }

}

$Timestamp =
    Get-Date -Format "yyyyMMdd-HHmmss"

$LogFile =
    "$LogPath\Pilot-Migration-$Timestamp.csv"

$MigrationLog |
    Export-Csv `
        $LogFile `
        -NoTypeInformation

$SuccessCount =
    @($MigrationLog |
        Where-Object Status -eq "SUCCESS").Count

$FailedCount =
    @($MigrationLog |
        Where-Object Status -eq "FAILED").Count

Write-Host ""
Write-Host "============================================================"
Write-Host " PILOT MIGRATION COMPLETE"
Write-Host "============================================================"
Write-Host ""
Write-Host "Total:   $($MigrationLog.Count)"
Write-Host "Success: $SuccessCount"
Write-Host "Failed:  $FailedCount"
Write-Host ""
Write-Host "Log:"
Write-Host $LogFile
Write-Host ""
Write-Host "Next step:"
Write-Host ".\Scripts\05-Validate-Pilot.ps1"
Write-Host ""
```

---

# 8. `05-Validate-Pilot.ps1`

## Purpose

This script compares the source and target pilot content by relative file path and file size.

Create:

```powershell
notepad C:\SPModernization\Scripts\05-Validate-Pilot.ps1
```

Paste:

```powershell
# ================================================================
# 05-Validate-Pilot.ps1
#
# PURPOSE:
# Validate the pilot migration by comparing:
# - relative file paths
# - file counts
# - file sizes
# ================================================================

$ErrorActionPreference = "Stop"

$RootPath   = "C:\SPModernization"
$ConfigPath = Join-Path $RootPath "config.psd1"
$ReportPath = Join-Path $RootPath "Reports"

$Config = Import-PowerShellDataFile $ConfigPath

function Connect-ToPnPSite {

    param (
        [Parameter(Mandatory)]
        [string]$Url
    )

    if (
        $Config.ContainsKey("ClientId") -and
        -not [string]::IsNullOrWhiteSpace($Config.ClientId)
    ) {

        return Connect-PnPOnline `
            -Url $Url `
            -ClientId $Config.ClientId `
            -Interactive `
            -ReturnConnection

    }
    else {

        return Connect-PnPOnline `
            -Url $Url `
            -Interactive `
            -ReturnConnection

    }

}

function Get-LibraryData {

    param (

        [Parameter(Mandatory)]
        $Connection,

        [Parameter(Mandatory)]
        [string]$LibraryName

    )

    $Web = Get-PnPWeb `
        -Includes ServerRelativeUrl `
        -Connection $Connection

    $List = Get-PnPList `
        -Identity $LibraryName `
        -Includes RootFolder `
        -Connection $Connection

    $RootServerRelative =
        $List.RootFolder.ServerRelativeUrl.TrimEnd('/')

    $WebRoot =
        $Web.ServerRelativeUrl.TrimEnd('/')

    if ([string]::IsNullOrWhiteSpace($WebRoot)) {
        $WebRoot = "/"
    }

    if ($WebRoot -eq "/") {

        $SiteRelative =
            $RootServerRelative.TrimStart('/')

    }
    else {

        $SiteRelative =
            $RootServerRelative.Substring(
                $WebRoot.Length
            ).TrimStart('/')

    }

    return [PSCustomObject]@{

        RootServerRelative = $RootServerRelative
        SiteRelative       = $SiteRelative

    }

}

Write-Host ""
Write-Host "============================================================"
Write-Host " VALIDATE PILOT"
Write-Host "============================================================"
Write-Host ""

$SourceConnection =
    Connect-ToPnPSite `
        -Url $Config.SourceSiteUrl

$TargetConnection =
    Connect-ToPnPSite `
        -Url $Config.TargetSiteUrl

$SourceData =
    Get-LibraryData `
        -Connection $SourceConnection `
        -LibraryName $Config.SourceLibrary

$TargetData =
    Get-LibraryData `
        -Connection $TargetConnection `
        -LibraryName $Config.TargetLibrary

$SourceAllFiles = @(
    Get-PnPFolderItem `
        -FolderSiteRelativeUrl $SourceData.SiteRelative `
        -ItemType File `
        -Recursive `
        -Connection $SourceConnection
)

$TargetAllFiles = @(
    Get-PnPFolderItem `
        -FolderSiteRelativeUrl $TargetData.SiteRelative `
        -ItemType File `
        -Recursive `
        -Connection $TargetConnection
)

$PilotPrefix =
    "$($Config.PilotFolder.Trim('/'))/"

$SourceMap = @{}

foreach ($File in $SourceAllFiles) {

    $Relative =
        $File.ServerRelativeUrl.Substring(
            $SourceData.RootServerRelative.Length
        ).TrimStart('/')

    if (
        $Relative.StartsWith(
            $PilotPrefix,
            [System.StringComparison]::OrdinalIgnoreCase
        )
    ) {

        $SourceMap[$Relative.ToLowerInvariant()] =
            [long]$File.Length

    }

}

$TargetMap = @{}

foreach ($File in $TargetAllFiles) {

    $Relative =
        $File.ServerRelativeUrl.Substring(
            $TargetData.RootServerRelative.Length
        ).TrimStart('/')

    if (
        $Relative.StartsWith(
            $PilotPrefix,
            [System.StringComparison]::OrdinalIgnoreCase
        )
    ) {

        $TargetMap[$Relative.ToLowerInvariant()] =
            [long]$File.Length

    }

}

$Missing = @()
$Unexpected = @()
$SizeMismatch = @()

foreach ($Path in $SourceMap.Keys) {

    if (-not $TargetMap.ContainsKey($Path)) {

        $Missing += [PSCustomObject]@{
            RelativePath = $Path
        }

    }
    elseif ($SourceMap[$Path] -ne $TargetMap[$Path]) {

        $SizeMismatch += [PSCustomObject]@{

            RelativePath = $Path
            SourceSize   = $SourceMap[$Path]
            TargetSize   = $TargetMap[$Path]

        }

    }

}

foreach ($Path in $TargetMap.Keys) {

    if (-not $SourceMap.ContainsKey($Path)) {

        $Unexpected += [PSCustomObject]@{
            RelativePath = $Path
        }

    }

}

$Timestamp =
    Get-Date -Format "yyyyMMdd-HHmmss"

$Missing |
    Export-Csv `
        "$ReportPath\Pilot-Missing-$Timestamp.csv" `
        -NoTypeInformation

$Unexpected |
    Export-Csv `
        "$ReportPath\Pilot-Unexpected-$Timestamp.csv" `
        -NoTypeInformation

$SizeMismatch |
    Export-Csv `
        "$ReportPath\Pilot-SizeMismatch-$Timestamp.csv" `
        -NoTypeInformation

Write-Host "Source pilot files: $($SourceMap.Count)"
Write-Host "Target pilot files: $($TargetMap.Count)"
Write-Host ""
Write-Host "Missing:       $($Missing.Count)"
Write-Host "Unexpected:    $($Unexpected.Count)"
Write-Host "Size mismatch: $($SizeMismatch.Count)"
Write-Host ""

if (
    $Missing.Count -eq 0 -and
    $SizeMismatch.Count -eq 0
) {

    Write-Host `
        "AUTOMATED PILOT VALIDATION PASSED." `
        -ForegroundColor Green

}
else {

    Write-Warning `
        "Pilot validation found problems. DO NOT run the full migration."

}

Write-Host ""
Write-Host @"
AUTOMATED VALIDATION IS NOT USER ACCEPTANCE TESTING.

Before proceeding, manually verify:

- Files open correctly
- Folder hierarchy looks correct
- Modern page links point to the right locations
- Expected users can access the content
- Metadata requirements have been addressed
- Version-history requirements have been addressed

Have the content owner approve the pilot.
"@
```

---

# 9. `06-Migrate-All.ps1`

## Purpose

This script migrates all files and recreates the folder hierarchy. It is designed to be rerunnable by using an idempotency decision before copying each file.

```text
Target missing                 -> COPY
Target exists + same size      -> SKIP
Target exists + different size -> CONFLICT
```

Conflicting files are not overwritten automatically.

Create:

```powershell
notepad C:\SPModernization\Scripts\06-Migrate-All.ps1
```

Paste:

```powershell
# ================================================================
# 06-Migrate-All.ps1
#
# PURPOSE:
# Migrate all files and recreate folder hierarchy.
#
# IDEMPOTENCY RULES:
# Target missing                 -> COPY
# Target exists + same size      -> SKIP
# Target exists + different size -> CONFLICT
#
# Conflicting files are NOT overwritten automatically.
# ================================================================

$ErrorActionPreference = "Stop"

$RootPath   = "C:\SPModernization"
$ConfigPath = Join-Path $RootPath "config.psd1"
$LogPath    = Join-Path $RootPath "Logs"
$TempPath   = Join-Path $RootPath "Temp"

$Config = Import-PowerShellDataFile $ConfigPath

function Connect-ToPnPSite {

    param (
        [Parameter(Mandatory)]
        [string]$Url
    )

    if (
        $Config.ContainsKey("ClientId") -and
        -not [string]::IsNullOrWhiteSpace($Config.ClientId)
    ) {

        return Connect-PnPOnline `
            -Url $Url `
            -ClientId $Config.ClientId `
            -Interactive `
            -ReturnConnection

    }
    else {

        return Connect-PnPOnline `
            -Url $Url `
            -Interactive `
            -ReturnConnection

    }

}

function Invoke-WithRetry {

    param (

        [Parameter(Mandatory)]
        [scriptblock]$Operation,

        [int]$MaximumAttempts = 4

    )

    $Attempt = 1

    while ($Attempt -le $MaximumAttempts) {

        try {

            return & $Operation

        }
        catch {

            if ($Attempt -ge $MaximumAttempts) {
                throw
            }

            $Delay =
                [Math]::Pow(2, $Attempt)

            Write-Warning `
                "Attempt $Attempt failed. Retrying in $Delay seconds."

            Start-Sleep `
                -Seconds $Delay

            $Attempt++

        }

    }

}

Write-Host ""
Write-Host "============================================================"
Write-Host " FULL SHAREPOINT MIGRATION"
Write-Host "============================================================"
Write-Host ""

Write-Warning @"
You are about to start the FULL content migration.

Source:
$($Config.SourceSiteUrl)

Target:
$($Config.TargetSiteUrl)

Source Library:
$($Config.SourceLibrary)

Target Library:
$($Config.TargetLibrary)

Existing conflicting target files will NOT be overwritten.
"@

$Confirmation =
    Read-Host "Type MIGRATE-ALL to continue"

if ($Confirmation -ne "MIGRATE-ALL") {

    throw "Full migration cancelled."

}

$SourceConnection =
    Connect-ToPnPSite `
        -Url $Config.SourceSiteUrl

$TargetConnection =
    Connect-ToPnPSite `
        -Url $Config.TargetSiteUrl

$SourceWeb =
    Get-PnPWeb `
        -Includes ServerRelativeUrl `
        -Connection $SourceConnection

$SourceList =
    Get-PnPList `
        -Identity $Config.SourceLibrary `
        -Includes RootFolder `
        -Connection $SourceConnection

$SourceRoot =
    $SourceList.RootFolder.ServerRelativeUrl.TrimEnd('/')

$SourceWebRoot =
    $SourceWeb.ServerRelativeUrl.TrimEnd('/')

if ([string]::IsNullOrWhiteSpace($SourceWebRoot)) {
    $SourceWebRoot = "/"
}

if ($SourceWebRoot -eq "/") {

    $SourceLibrarySiteRelative =
        $SourceRoot.TrimStart('/')

}
else {

    $SourceLibrarySiteRelative =
        $SourceRoot.Substring(
            $SourceWebRoot.Length
        ).TrimStart('/')

}

$TargetList =
    Get-PnPList `
        -Identity $Config.TargetLibrary `
        -Includes RootFolder `
        -Connection $TargetConnection

$TargetRoot =
    $TargetList.RootFolder.ServerRelativeUrl.TrimEnd('/')

$AllFiles = @(
    Get-PnPFolderItem `
        -FolderSiteRelativeUrl $SourceLibrarySiteRelative `
        -ItemType File `
        -Recursive `
        -Connection $SourceConnection
)

if ($AllFiles.Count -eq 0) {

    throw "No source files were discovered."

}

Write-Host ""
Write-Host "Source files discovered: $($AllFiles.Count)"
Write-Host ""

$MigrationLog = @()
$Counter = 0

foreach ($File in $AllFiles) {

    $Counter++
    $LocalTempFile = $null

    try {

        $RelativePath =
            $File.ServerRelativeUrl.Substring(
                $SourceRoot.Length
            ).TrimStart('/')

        Write-Host ""
        Write-Host "[$Counter/$($AllFiles.Count)] $RelativePath"

        $LastSlash =
            $RelativePath.LastIndexOf('/')

        if ($LastSlash -gt 0) {

            $RelativeFolder =
                $RelativePath.Substring(
                    0,
                    $LastSlash
                )

        }
        else {

            $RelativeFolder = ""

        }

        $TargetFolderSiteRelative =
            $Config.TargetLibrary

        if (-not [string]::IsNullOrWhiteSpace($RelativeFolder)) {

            $TargetFolderSiteRelative =
                "$($Config.TargetLibrary)/$RelativeFolder"

        }

        $ResolvedTargetFolder =
            Resolve-PnPFolder `
                -SiteRelativePath $TargetFolderSiteRelative `
                -Connection $TargetConnection

        $TargetFileServerRelative =
            "$($ResolvedTargetFolder.ServerRelativeUrl.TrimEnd('/'))/$($File.Name)"

        $ExistingTargetFile = $null

        try {

            $ExistingTargetFile =
                Get-PnPFile `
                    -Url $TargetFileServerRelative `
                    -AsFileObject `
                    -Connection $TargetConnection `
                    -ErrorAction Stop

        }
        catch {

            $ExistingTargetFile = $null

        }

        if ($ExistingTargetFile) {

            $SourceSize =
                [long]$File.Length

            $TargetSize =
                [long]$ExistingTargetFile.Length

            if ($SourceSize -eq $TargetSize) {

                Write-Host `
                    "SKIP - target already exists with matching size." `
                    -ForegroundColor Yellow

                $MigrationLog += [PSCustomObject]@{

                    RelativePath = $RelativePath
                    Status       = "SKIPPED"
                    SourceSize   = $SourceSize
                    TargetSize   = $TargetSize
                    Error        = ""

                }

                continue

            }
            else {

                Write-Warning `
                    "CONFLICT - target exists but file size differs."

                $MigrationLog += [PSCustomObject]@{

                    RelativePath = $RelativePath
                    Status       = "CONFLICT"
                    SourceSize   = $SourceSize
                    TargetSize   = $TargetSize
                    Error        = "Target file exists with different size."

                }

                continue

            }

        }

        $Extension =
            [System.IO.Path]::GetExtension($File.Name)

        $TempName =
            "$([guid]::NewGuid().ToString())$Extension"

        $LocalTempFile =
            Join-Path $TempPath $TempName

        Invoke-WithRetry {

            Get-PnPFile `
                -Url $File.ServerRelativeUrl `
                -Path $TempPath `
                -FileName $TempName `
                -AsFile `
                -Force `
                -Connection $SourceConnection

        } | Out-Null

        Invoke-WithRetry {

            Add-PnPFile `
                -Path $LocalTempFile `
                -Folder $ResolvedTargetFolder.ServerRelativeUrl `
                -NewFileName $File.Name `
                -Connection $TargetConnection

        } | Out-Null

        $MigrationLog += [PSCustomObject]@{

            RelativePath = $RelativePath
            Status       = "SUCCESS"
            SourceSize   = [long]$File.Length
            TargetSize   = [long]$File.Length
            Error        = ""

        }

        Write-Host "SUCCESS" `
            -ForegroundColor Green

    }
    catch {

        $MigrationLog += [PSCustomObject]@{

            RelativePath = $RelativePath
            Status       = "FAILED"
            SourceSize   = [long]$File.Length
            TargetSize   = ""
            Error        = $_.Exception.Message

        }

        Write-Warning `
            "FAILED: $($_.Exception.Message)"

    }
    finally {

        if (
            $LocalTempFile -and
            (Test-Path $LocalTempFile)
        ) {

            Remove-Item `
                $LocalTempFile `
                -Force `
                -ErrorAction SilentlyContinue

        }

    }

}

$Timestamp =
    Get-Date -Format "yyyyMMdd-HHmmss"

$LogFile =
    "$LogPath\Full-Migration-$Timestamp.csv"

$MigrationLog |
    Export-Csv `
        $LogFile `
        -NoTypeInformation

$Successful =
    @($MigrationLog |
        Where-Object Status -eq "SUCCESS").Count

$Skipped =
    @($MigrationLog |
        Where-Object Status -eq "SKIPPED").Count

$Conflicts =
    @($MigrationLog |
        Where-Object Status -eq "CONFLICT").Count

$Failed =
    @($MigrationLog |
        Where-Object Status -eq "FAILED").Count

Write-Host ""
Write-Host "============================================================"
Write-Host " FULL MIGRATION COMPLETE"
Write-Host "============================================================"
Write-Host ""
Write-Host "Successful: $Successful"
Write-Host "Skipped:    $Skipped"
Write-Host "Conflicts:  $Conflicts"
Write-Host "Failed:     $Failed"
Write-Host ""
Write-Host "Migration log:"
Write-Host $LogFile
Write-Host ""
Write-Host "Next:"
Write-Host ".\Scripts\07-Validate-All.ps1"
Write-Host ""
```

---

# 10. `07-Validate-All.ps1`

## Purpose

This script compares the entire source library against the entire target library and produces reports for missing files, unexpected files, and size mismatches.

Create:

```powershell
notepad C:\SPModernization\Scripts\07-Validate-All.ps1
```

Paste:

```powershell
# ================================================================
# 07-Validate-All.ps1
#
# PURPOSE:
# Compare the complete source and target libraries.
#
# Produces:
# Missing
# Unexpected
# SizeMismatch
# ================================================================

$ErrorActionPreference = "Stop"

$RootPath   = "C:\SPModernization"
$ConfigPath = Join-Path $RootPath "config.psd1"
$ReportPath = Join-Path $RootPath "Reports"

$Config = Import-PowerShellDataFile $ConfigPath

function Connect-ToPnPSite {

    param (
        [Parameter(Mandatory)]
        [string]$Url
    )

    if (
        $Config.ContainsKey("ClientId") -and
        -not [string]::IsNullOrWhiteSpace($Config.ClientId)
    ) {

        return Connect-PnPOnline `
            -Url $Url `
            -ClientId $Config.ClientId `
            -Interactive `
            -ReturnConnection

    }
    else {

        return Connect-PnPOnline `
            -Url $Url `
            -Interactive `
            -ReturnConnection

    }

}

function Get-LibraryInformation {

    param (

        $Connection,

        [string]$LibraryName

    )

    $Web =
        Get-PnPWeb `
            -Includes ServerRelativeUrl `
            -Connection $Connection

    $List =
        Get-PnPList `
            -Identity $LibraryName `
            -Includes RootFolder `
            -Connection $Connection

    $Root =
        $List.RootFolder.ServerRelativeUrl.TrimEnd('/')

    $WebRoot =
        $Web.ServerRelativeUrl.TrimEnd('/')

    if ([string]::IsNullOrWhiteSpace($WebRoot)) {
        $WebRoot = "/"
    }

    if ($WebRoot -eq "/") {

        $SiteRelative =
            $Root.TrimStart('/')

    }
    else {

        $SiteRelative =
            $Root.Substring(
                $WebRoot.Length
            ).TrimStart('/')

    }

    return [PSCustomObject]@{

        Root         = $Root
        SiteRelative = $SiteRelative

    }

}

Write-Host ""
Write-Host "============================================================"
Write-Host " FULL MIGRATION VALIDATION"
Write-Host "============================================================"
Write-Host ""

$SourceConnection =
    Connect-ToPnPSite `
        -Url $Config.SourceSiteUrl

$TargetConnection =
    Connect-ToPnPSite `
        -Url $Config.TargetSiteUrl

$SourceInfo =
    Get-LibraryInformation `
        -Connection $SourceConnection `
        -LibraryName $Config.SourceLibrary

$TargetInfo =
    Get-LibraryInformation `
        -Connection $TargetConnection `
        -LibraryName $Config.TargetLibrary

$SourceFiles = @(
    Get-PnPFolderItem `
        -FolderSiteRelativeUrl $SourceInfo.SiteRelative `
        -ItemType File `
        -Recursive `
        -Connection $SourceConnection
)

$TargetFiles = @(
    Get-PnPFolderItem `
        -FolderSiteRelativeUrl $TargetInfo.SiteRelative `
        -ItemType File `
        -Recursive `
        -Connection $TargetConnection
)

$SourceMap = @{}

foreach ($File in $SourceFiles) {

    $RelativePath =
        $File.ServerRelativeUrl.Substring(
            $SourceInfo.Root.Length
        ).TrimStart('/')

    $SourceMap[$RelativePath.ToLowerInvariant()] =
        [long]$File.Length

}

$TargetMap = @{}

foreach ($File in $TargetFiles) {

    $RelativePath =
        $File.ServerRelativeUrl.Substring(
            $TargetInfo.Root.Length
        ).TrimStart('/')

    $TargetMap[$RelativePath.ToLowerInvariant()] =
        [long]$File.Length

}

$Missing = @()
$Unexpected = @()
$SizeMismatch = @()

foreach ($Path in $SourceMap.Keys) {

    if (-not $TargetMap.ContainsKey($Path)) {

        $Missing += [PSCustomObject]@{

            RelativePath = $Path
            SourceSize   = $SourceMap[$Path]

        }

    }
    elseif ($SourceMap[$Path] -ne $TargetMap[$Path]) {

        $SizeMismatch += [PSCustomObject]@{

            RelativePath = $Path
            SourceSize   = $SourceMap[$Path]
            TargetSize   = $TargetMap[$Path]

        }

    }

}

foreach ($Path in $TargetMap.Keys) {

    if (-not $SourceMap.ContainsKey($Path)) {

        $Unexpected += [PSCustomObject]@{

            RelativePath = $Path
            TargetSize   = $TargetMap[$Path]

        }

    }

}

$Timestamp =
    Get-Date -Format "yyyyMMdd-HHmmss"

$Missing |
    Export-Csv `
        "$ReportPath\Missing-$Timestamp.csv" `
        -NoTypeInformation

$Unexpected |
    Export-Csv `
        "$ReportPath\Unexpected-$Timestamp.csv" `
        -NoTypeInformation

$SizeMismatch |
    Export-Csv `
        "$ReportPath\SizeMismatch-$Timestamp.csv" `
        -NoTypeInformation

Write-Host "Source files:     $($SourceMap.Count)"
Write-Host "Target files:     $($TargetMap.Count)"
Write-Host ""
Write-Host "Missing:          $($Missing.Count)"
Write-Host "Unexpected:       $($Unexpected.Count)"
Write-Host "Size mismatches:  $($SizeMismatch.Count)"
Write-Host ""

if (
    $Missing.Count -eq 0 -and
    $SizeMismatch.Count -eq 0
) {

    Write-Host `
        "AUTOMATED CONTENT VALIDATION PASSED." `
        -ForegroundColor Green

}
else {

    Write-Warning @"
VALIDATION FAILED.

DO NOT PERFORM CUTOVER.

Review the generated reports.
"@

}

Write-Host ""
Write-Host @"
Before cutover you must still perform USER ACCEPTANCE TESTING:

- Open representative documents
- Test all Modern page links
- Test access as appropriate user groups
- Verify navigation
- Verify expected metadata behavior
- Verify business processes/workflows
- Obtain content owner/supervisor approval
"@
```

---

# 11. `08-Cutover.ps1`

## Purpose

Cutover should be small and deliberate. This script publishes the Modern page and optionally sets it as the site homepage. It does not delete the Classic source.

Create:

```powershell
notepad C:\SPModernization\Scripts\08-Cutover.ps1
```

Paste:

```powershell
# ================================================================
# 08-Cutover.ps1
#
# PURPOSE:
# Publish the Modern page and optionally configure it as the
# SharePoint homepage.
#
# THE CLASSIC SOURCE IS NOT DELETED.
# ================================================================

$ErrorActionPreference = "Stop"

$RootPath   = "C:\SPModernization"
$ConfigPath = Join-Path $RootPath "config.psd1"

$Config = Import-PowerShellDataFile $ConfigPath

function Connect-ToPnPSite {

    param (
        [Parameter(Mandatory)]
        [string]$Url
    )

    if (
        $Config.ContainsKey("ClientId") -and
        -not [string]::IsNullOrWhiteSpace($Config.ClientId)
    ) {

        return Connect-PnPOnline `
            -Url $Url `
            -ClientId $Config.ClientId `
            -Interactive `
            -ReturnConnection

    }
    else {

        return Connect-PnPOnline `
            -Url $Url `
            -Interactive `
            -ReturnConnection

    }

}

Write-Host ""
Write-Host "============================================================"
Write-Host " PRODUCTION CUTOVER"
Write-Host "============================================================"
Write-Host ""

Write-Warning @"
CUTOVER SHOULD ONLY OCCUR AFTER:

1. Pilot validation passed.
2. Full validation passed.
3. User Acceptance Testing passed.
4. Content owner approved the Modern page.
5. Supervisor approved cutover.

The Classic page/site will NOT be deleted.
"@

$Confirmation =
    Read-Host "Type CUTOVER exactly to continue"

if ($Confirmation -ne "CUTOVER") {

    throw "Cutover cancelled."

}

$TargetConnection =
    Connect-ToPnPSite `
        -Url $Config.TargetSiteUrl

$TargetPage =
    "$($Config.TargetPageName).aspx"

Write-Host "Publishing Modern page..."

Set-PnPPage `
    -Identity $TargetPage `
    -Publish `
    -Connection $TargetConnection

Write-Host `
    "[PASS] Modern page published." `
    -ForegroundColor Green

Write-Host ""
$SetHomePage =
    Read-Host "Should this page become the site homepage? Type YES or NO"

if ($SetHomePage -eq "YES") {

    $SecondConfirmation =
        Read-Host `
            "Type SET-HOMEPAGE to confirm"

    if ($SecondConfirmation -eq "SET-HOMEPAGE") {

        Set-PnPHomePage `
            -RootFolderRelativeUrl `
            "SitePages/$TargetPage" `
            -Connection $TargetConnection

        Write-Host `
            "[PASS] Modern page set as homepage." `
            -ForegroundColor Green

    }
    else {

        Write-Host "Homepage change skipped."

    }

}
else {

    Write-Host `
        "Page published but homepage was not changed."

}

Write-Host ""
Write-Host "============================================================"
Write-Host " CUTOVER COMPLETE"
Write-Host "============================================================"
Write-Host ""
Write-Host @"
The Classic site/page remains preserved.

Do NOT delete it at this stage.

Recommended next phase:

Modern production use
        |
Observation period
        |
User feedback
        |
Leadership approval
        |
Archive/retire Classic separately
"@
```

---

# 12. Master Script: `Invoke-SharePointModernization.ps1`

## Purpose

The master script orchestrates all numbered scripts while preserving manual checkpoints. It should only be used after the operator understands and has tested the individual scripts.

Create:

```powershell
notepad C:\SPModernization\Invoke-SharePointModernization.ps1
```

Paste:

```powershell
# ================================================================
# Invoke-SharePointModernization.ps1
#
# MASTER ORCHESTRATION SCRIPT
#
# PURPOSE:
# Guide an operator through the complete SharePoint modernization
# workflow while preserving human validation checkpoints.
#
# Do not use this master script until the individual scripts
# have been tested and understood.
# ================================================================

$ErrorActionPreference = "Stop"

$RootPath   = "C:\SPModernization"
$ScriptPath = Join-Path $RootPath "Scripts"

function Invoke-MigrationStage {

    param (

        [Parameter(Mandatory)]
        [string]$Script,

        [Parameter(Mandatory)]
        [string]$Description

    )

    $FullPath =
        Join-Path $ScriptPath $Script

    if (-not (Test-Path $FullPath)) {

        throw "Required script not found: $FullPath"

    }

    Write-Host ""
    Write-Host "============================================================"
    Write-Host " $Description"
    Write-Host "============================================================"
    Write-Host ""

    & $FullPath

    if ($LASTEXITCODE -and $LASTEXITCODE -ne 0) {

        throw "$Script returned an error."

    }

}

function Wait-ForApproval {

    param (

        [Parameter(Mandatory)]
        [string]$Message,

        [string]$RequiredText = "CONTINUE"

    )

    Write-Host ""
    Write-Host "------------------------------------------------------------"
    Write-Host " HUMAN CHECKPOINT"
    Write-Host "------------------------------------------------------------"
    Write-Host ""
    Write-Host $Message
    Write-Host ""

    $Response =
        Read-Host "Type $RequiredText to continue"

    if ($Response -ne $RequiredText) {

        throw "Workflow stopped by operator."

    }

}

Write-Host ""
Write-Host "############################################################"
Write-Host "#                                                          #"
Write-Host "#      SHAREPOINT CLASSIC -> MODERN MIGRATION              #"
Write-Host "#                                                          #"
Write-Host "#              MASTER RUNBOOK                              #"
Write-Host "#                                                          #"
Write-Host "############################################################"
Write-Host ""

Write-Warning @"
This script orchestrates the migration.

It intentionally contains manual approval checkpoints.

Automation does NOT remove the need for human review.
"@

$Start =
    Read-Host "Type START to begin"

if ($Start -ne "START") {

    throw "Migration workflow cancelled."

}

Invoke-MigrationStage `
    -Script "00-Preflight.ps1" `
    -Description "STAGE 00 - PREFLIGHT"

Wait-ForApproval `
    -Message @"
Review the preflight information.

Verify:

- Correct source URL
- Correct target URL
- Correct source library
- Correct target library
- Correct source page
- Correct environment
"@

Invoke-MigrationStage `
    -Script "01-Inventory.ps1" `
    -Description "STAGE 01 - INVENTORY"

Wait-ForApproval `
    -Message @"
STOP and inspect:

C:\SPModernization\Reports\Source-Web.csv
C:\SPModernization\Reports\Source-Lists.csv
C:\SPModernization\Reports\Source-Navigation.csv
C:\SPModernization\Reports\Source-Folders.csv
C:\SPModernization\Reports\Source-Files.csv

Confirm that the discovered structure represents the actual
Classic SharePoint environment.
"@

Invoke-MigrationStage `
    -Script "02-Classify-And-Backup-Page.ps1" `
    -Description "STAGE 02 - PAGE CLASSIFICATION"

Wait-ForApproval `
    -Message @"
Review the page classification.

If:

MODERN
    No conversion is required.

CLASSIC
    Confirm that the Classic source remains untouched and
    the reference backup was created.

UNKNOWN
    The classification script should already have stopped
    the workflow.

Do not proceed unless the classification is understood.
"@

Invoke-MigrationStage `
    -Script "03-Build-Modern-Target.ps1" `
    -Description "STAGE 03 - BUILD MODERN TARGET"

Wait-ForApproval `
    -Message @"
Open the Modern SharePoint site in the browser.

Verify:

- Modern page exists
- Target document library exists
- No duplicate page was created
- No duplicate library was created
- Layout is acceptable for pilot testing
"@

Invoke-MigrationStage `
    -Script "04-Migrate-Pilot.ps1" `
    -Description "STAGE 04 - PILOT MIGRATION"

Invoke-MigrationStage `
    -Script "05-Validate-Pilot.ps1" `
    -Description "STAGE 05 - VALIDATE PILOT"

Wait-ForApproval `
    -Message @"
DO NOT approve based only on file counts.

Manually verify the pilot:

- Folder hierarchy
- Documents
- Files open
- Modern page links
- Permissions/access
- Business functionality
- Required metadata behavior

Obtain pilot approval from the appropriate content owner.
"@ `
    -RequiredText "PILOT-APPROVED"

Invoke-MigrationStage `
    -Script "06-Migrate-All.ps1" `
    -Description "STAGE 06 - FULL MIGRATION"

Invoke-MigrationStage `
    -Script "07-Validate-All.ps1" `
    -Description "STAGE 07 - FULL VALIDATION"

Wait-ForApproval `
    -Message @"
Review all migration logs and validation reports.

Desired state:

Missing files        = 0
Size mismatches      = 0
Failed migrations    = 0

Any conflicts or unexpected target files must be reviewed.

Perform complete User Acceptance Testing before cutover.
"@ `
    -RequiredText "UAT-APPROVED"

Invoke-MigrationStage `
    -Script "08-Cutover.ps1" `
    -Description "STAGE 08 - CUTOVER"

Write-Host ""
Write-Host "############################################################"
Write-Host "#                                                          #"
Write-Host "#           MIGRATION WORKFLOW COMPLETE                    #"
Write-Host "#                                                          #"
Write-Host "############################################################"
Write-Host ""

Write-Host @"
The Modern environment has reached the cutover stage.

The Classic environment should remain preserved until a separate
retirement/archive decision is approved.

Continue monitoring:

- Broken links
- Access problems
- Missing content
- User feedback
- Workflow/functionality issues
- Search behavior
- Metadata behavior
"@
```

---

# 13. Script-to-Stage Mapping

```text
                    config.psd1
                        |
                        v
               00-Preflight.ps1
                        |
                   SAFE TO RUN?
                        |
                        v
                01-Inventory.ps1
                        |
                 WHAT EXISTS?
                        |
                        v
       02-Classify-And-Backup-Page.ps1
                        |
                WHAT PAGE TYPE?
                 /      |      \
            Classic   Modern   Unknown
               |         |        |
          Preserve      Keep      STOP
               \         /
                \       /
                    v
       03-Build-Modern-Target.ps1
                    |
                    v
          04-Migrate-Pilot.ps1
                    |
               ONE FOLDER
                    |
                    v
          05-Validate-Pilot.ps1
                    |
              CONTENT OWNER
                 APPROVAL
                    |
                    v
           06-Migrate-All.ps1
                    |
              ALL CONTENT
                    |
                    v
           07-Validate-All.ps1
                    |
                  UAT
                    |
                    v
              08-Cutover.ps1
                    |
                    v
              MODERN LIVE

Classic remains preserved
```

---

# 14. Recommended First-Run Sequence

Do not start with the master script during your first migration.

Open PowerShell 7:

```powershell
pwsh
```

Go to the project directory:

```powershell
cd C:\SPModernization
```

Run:

```powershell
.\Scripts\00-Preflight.ps1
```

Stop and review the output.

Then run:

```powershell
.\Scripts\01-Inventory.ps1
```

Stop and inspect the CSV reports.

Then run:

```powershell
.\Scripts\02-Classify-And-Backup-Page.ps1
```

For the first real-world migration, stop here and review:

```text
Source-Lists.csv
Source-Navigation.csv
Source-Folders.csv
Source-Files.csv
Page-Classification.csv
```

The discovery stage determines whether items such as AGR, CSS Training, Career Development, Force Management, and other links are actually folders, libraries, pages, subsites, or other resources.

After the environment is understood, continue with:

```powershell
.\Scripts\03-Build-Modern-Target.ps1
.\Scripts\04-Migrate-Pilot.ps1
.\Scripts\05-Validate-Pilot.ps1
```

After pilot approval:

```powershell
.\Scripts\06-Migrate-All.ps1
.\Scripts\07-Validate-All.ps1
```

Only after validation and UAT:

```powershell
.\Scripts\08-Cutover.ps1
```

After the procedure is mature and understood, the master orchestrator can be used:

```powershell
.\Invoke-SharePointModernization.ps1
```

---

# 15. Troubleshooting and Operational Guardrails

| Situation | Meaning | Action |
|---|---|---|
| `401 Unauthorized` | Authentication failure | Reauthenticate and verify tenant/app configuration |
| `403 Access Denied` | Insufficient permissions | Verify SharePoint/site/app permissions |
| `404 File Not Found` | Incorrect path or library name | Inspect actual server-relative URL |
| `429 Too Many Requests` | SharePoint throttling | Retry with exponential delay |
| `503 Service Unavailable` | Temporary service condition | Retry after delay |
| Page type `Unknown` | Classification is not safe | Stop and manually investigate |
| Target file exists, same size | Likely already migrated | Skip |
| Target file exists, different size | Potential conflict | Log and manually review |
| Script Editor or custom JavaScript | May have no direct Modern equivalent | Redesign instead of blind conversion |
| Missing metadata | Basic file transfer only preserved content | Add a metadata migration strategy |
| Missing version history | Basic transfer did not preserve versions | Use an appropriate migration method/tool |
| Unique permissions missing | Permissions were not migrated | Design permission migration separately |
| Retention/records restriction | Compliance policy may block or alter behavior | Coordinate with compliance/SharePoint admin |

---

# 16. What “Move Everything” Actually Means

Before full migration, confirm which layers are in scope:

1. Folders
2. Files
3. Metadata
4. Version history
5. Permissions
6. Page functionality
7. Workflows/automation
8. Retention/compliance settings

Moving a file is much simpler than preserving its full history and governance state. If leadership requires preservation of properties such as original Created/Modified dates, authors, approval status, custom content types, version history, unique permissions, or retention labels, stop before `06-Migrate-All.ps1` and design the correct migration method.

---

# 17. Key Learning Concepts

This runbook is not only about SharePoint. It demonstrates several reusable automation principles:

- **Preflight checks:** validate prerequisites before making changes.
- **Discovery before migration:** do not automate assumptions.
- **Guardrails:** stop when the script cannot safely decide.
- **Source preservation:** never make rollback depend on hope.
- **Pilot migrations:** validate on a small scope before full execution.
- **Idempotency:** rerunning a script should not duplicate or corrupt data.
- **Retry logic:** transient service failures should not automatically ruin a long migration.
- **Logging:** every success, failure, skip, and conflict should be auditable.
- **Automated validation:** compare source and target, not just console messages.
- **Human checkpoints:** automation does not remove operational judgment.
- **UAT:** technical completion is not the same as business acceptance.
- **Separated cutover:** migration and go-live should be separate decisions.
- **Delayed retirement:** keep Classic available until Modern has proven stable and retirement is approved.

These principles apply equally to data migrations, ETL pipelines, cloud migrations, application deployments, and Knowledge Management automation.

---

## Final Operator Rule

For a new specialist, the safest mental model is:

```text
Read -> Run -> Stop -> Inspect -> Understand -> Approve -> Continue
```

Do not replace that process with:

```text
Run everything -> Hope -> Troubleshoot production
```

