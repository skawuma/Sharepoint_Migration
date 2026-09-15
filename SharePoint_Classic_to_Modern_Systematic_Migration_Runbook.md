---
title: "SharePoint Classic-to-Modern Systematic Migration Runbook"
subtitle: "A reproducible, step-by-step PowerShell process for Windows"
date: "September 2026"
geometry: margin=0.75in
fontsize: 10pt
---

# Purpose

For this Project, do **not** begin with one giant PowerShell script. Build the migration as a small toolkit with numbered stages. This makes the work safer, easier to troubleshoot, easier to repeat, and teachable to the next Knowledge Management specialist.

The operating principle is:

> **Discover -> classify -> preserve -> build -> pilot -> validate -> migrate -> validate -> cut over.**

One rule should exist throughout the project:

> **The migration scripts never delete or overwrite the Classic source.**

This runbook assumes **SharePoint Online**. If the environment is SharePoint Server/on-premises, stop after the preflight stage because authentication and several commands need to be adapted.

# 1. How to organize the Windows machine

Use PowerShell 7 and create a dedicated working directory.

Open PowerShell:

```powershell
pwsh
```

Check the version:

```powershell
$PSVersionTable
```

Create the migration workspace:

```powershell
New-Item -ItemType Directory -Path C:\SPModernization -Force

Set-Location C:\SPModernization

New-Item -ItemType Directory -Path `
    .\Scripts,
    .\Logs,
    .\Reports,
    .\Backup,
    .\Temp `
    -Force
```

The machine should eventually look like:

```text
C:\SPModernization
|
|-- config.psd1
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

Run these scripts **one at a time initially**.

Only after the migration has been successfully performed several times in testing should you create one orchestration script that runs them sequentially.

# 2. Install the tools

Run:

```powershell
Get-ExecutionPolicy -List
```

If the organization's policy permits it:

```powershell
Set-ExecutionPolicy `
    -ExecutionPolicy RemoteSigned `
    -Scope CurrentUser
```

Do **not** try to bypass an Air Force/enterprise Group Policy if it blocks script execution. Have the administrator approve the appropriate execution method.

Install PnP PowerShell:

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
```

Then:

```powershell
Import-Module PnP.PowerShell
```

Verify:

```powershell
Get-Module PnP.PowerShell -ListAvailable
```

Then test:

```powershell
Get-Command Connect-PnPOnline
```

You should get the command definition rather than:

```text
The term 'Connect-PnPOnline' is not recognized
```

# 3. Create one configuration file

Do **not** hard-code URLs throughout eight different scripts.

Create:

```powershell
notepad C:\SPModernization\config.psd1
```

Put something like this inside:

```powershell
@{

    # SOURCE
    SourceSiteUrl = "https://SOURCE-SITE-URL"

    SourcePage = "SitePages/Home.aspx"

    SourceLibrary = "Shared Documents"


    # TARGET
    TargetSiteUrl = "https://TARGET-SITE-URL"

    TargetPageName = ""

    TargetLibrary = "MPF Documents"


    # PILOT
    PilotFolder = "CSS Training"


    # Authentication
    ClientId = "YOUR-APP-ID-IF-REQUIRED"


    # Safety
    Environment = "TEST"

}
```

Do not put:

```text
password
client secret
certificate private key
CAC PIN
```

in this configuration file.

The organization may also use a different Microsoft cloud/authentication configuration than commercial Microsoft 365. Use the tenant/app configuration provided by the SharePoint/Entra administrator.

# 4. Script 00 - Preflight

Create:

```powershell
notepad C:\SPModernization\Scripts\00-Preflight.ps1
```

Use:

```powershell
$ErrorActionPreference = "Stop"

$Config = Import-PowerShellDataFile `
    "C:\SPModernization\config.psd1"

Write-Host "=== SHAREPOINT MIGRATION PREFLIGHT ===" -ForegroundColor Cyan

# PowerShell check

Write-Host "PowerShell Version:"
$PSVersionTable.PSVersion


# PnP check

$PnP = Get-Module PnP.PowerShell -ListAvailable |
    Sort-Object Version -Descending |
    Select-Object -First 1

if (-not $PnP) {
    throw "PnP.PowerShell is not installed."
}

Write-Host "PnP.PowerShell version: $($PnP.Version)"


# Required configuration

$Required = @(
    "SourceSiteUrl",
    "SourcePage",
    "SourceLibrary",
    "TargetSiteUrl",
    "TargetPageName",
    "TargetLibrary"
)

foreach ($Setting in $Required) {

    if ([string]::IsNullOrWhiteSpace($Config[$Setting])) {
        throw "Missing configuration value: $Setting"
    }

}


Write-Host ""
Write-Host "Source Site:  $($Config.SourceSiteUrl)"
Write-Host "Target Site:  $($Config.TargetSiteUrl)"
Write-Host "Source Page:  $($Config.SourcePage)"
Write-Host "Environment:  $($Config.Environment)"


if ($Config.Environment -eq "PROD") {

    Write-Warning "YOU ARE CONFIGURED FOR PRODUCTION."

    $Confirmation = Read-Host `
        "Type PRODUCTION to continue"

    if ($Confirmation -ne "PRODUCTION") {
        throw "Production execution cancelled."
    }

}


Write-Host ""
Write-Host "PRE-FLIGHT PASSED." -ForegroundColor Green
```

Run it:

```powershell
cd C:\SPModernization

.\Scripts\00-Preflight.ps1
```

## Stop here if it fails

Do not proceed to migration because PowerShell says:

```text
mostly looks okay
```

Every prerequisite should pass.

# 5. Script 01 - Inventory the Classic environment

This script should be **100% read-only**.

Nothing should be changed.

Create:

```powershell
notepad .\Scripts\01-Inventory.ps1
```

Start with:

```powershell
$ErrorActionPreference = "Stop"

$Config = Import-PowerShellDataFile `
    "C:\SPModernization\config.psd1"


Write-Host "Connecting to source..."


$SourceConnection = Connect-PnPOnline `
    -Url $Config.SourceSiteUrl `
    -Interactive `
    -ReturnConnection
```

Depending upon the organization's authentication configuration, an administrator-provided Client ID may be required.

For example:

```powershell
$SourceConnection = Connect-PnPOnline `
    -Url $Config.SourceSiteUrl `
    -ClientId $Config.ClientId `
    -Interactive `
    -ReturnConnection
```

Do not randomly alternate authentication methods. Use the one approved for the tenant.

## Inventory the site

Add:

```powershell
$Web = Get-PnPWeb `
    -Connection $SourceConnection

$Web |
    Select-Object Title,
                  Url,
                  ServerRelativeUrl |
    Export-Csv `
        "C:\SPModernization\Reports\Source-Web.csv" `
        -NoTypeInformation
```

## Inventory document libraries and lists

```powershell
Get-PnPList `
    -Connection $SourceConnection |
    Where-Object {
        $_.Hidden -eq $false
    } |
    Select-Object Title,
                  BaseTemplate,
                  ItemCount,
                  DefaultViewUrl |
    Export-Csv `
        "C:\SPModernization\Reports\Source-Lists.csv" `
        -NoTypeInformation
```

## Inventory navigation

```powershell
Get-PnPNavigationNode `
    -Location QuickLaunch `
    -Connection $SourceConnection |
    Select-Object Title,
                  Url |
    Export-Csv `
        "C:\SPModernization\Reports\Source-Navigation.csv" `
        -NoTypeInformation
```

This step is particularly important for the MPF page.

Something displayed as:

```text
AGR
```

may actually point to:

```text
/Shared Documents/AGR
```

while:

```text
Career Development
```

may point to:

```text
/SitePages/CareerDevelopment.aspx
```

Those are completely different migration cases.

# 6. Inventory folders

Use:

```powershell
$Folders = Get-PnPFolderItem `
    -FolderSiteRelativeUrl $Config.SourceLibrary `
    -ItemType Folder `
    -Recursive `
    -Connection $SourceConnection


$Folders |
    Select-Object Name,
                  ServerRelativeUrl |
    Export-Csv `
        "C:\SPModernization\Reports\Source-Folders.csv" `
        -NoTypeInformation
```

Now you could see:

```text
AGR
AGR/Forms
AGR/Policy
AGR/Applications

CSS Training
CSS Training/Slides
CSS Training/References

Force Management
Force Management/Policy
```

# 7. Inventory files

Still in `01-Inventory.ps1`:

```powershell
$Files = Get-PnPFolderItem `
    -FolderSiteRelativeUrl $Config.SourceLibrary `
    -ItemType File `
    -Recursive `
    -Connection $SourceConnection


$Files |
    Select-Object Name,
                  ServerRelativeUrl,
                  Length |
    Export-Csv `
        "C:\SPModernization\Reports\Source-Files.csv" `
        -NoTypeInformation
```

Then output your counts:

```powershell
Write-Host ""
Write-Host "Folder Count: $($Folders.Count)"
Write-Host "File Count:   $($Files.Count)"
```

Run:

```powershell
.\Scripts\01-Inventory.ps1
```

Then physically inspect:

```text
C:\SPModernization\Reports
```

before continuing.

# 8. Script 02 - Determine whether the page is Classic or Modern

This answers one of the biggest questions:

> Should the script blindly convert the page?

**No.**

It should first classify it.

The logic should be:

```text
Find page
    |
    +-- Modern?
    |      |
    |      +-- YES --> Don't convert it
    |
    +-- Classic?
    |      |
    |      +-- YES --> Preserve it
    |                  Create separate Modern replacement
    |
    +-- Unknown?
           |
           +-- STOP
               Manual investigation required
```

Unknown should **never** mean:

> "Let's assume it is Classic."

# 9. A practical page classification function

Create:

```powershell
notepad .\Scripts\02-Classify-And-Backup-Page.ps1
```

Start:

```powershell
$ErrorActionPreference = "Stop"

$Config = Import-PowerShellDataFile `
    "C:\SPModernization\config.psd1"


$SourceConnection = Connect-PnPOnline `
    -Url $Config.SourceSiteUrl `
    -Interactive `
    -ReturnConnection


$PageName = Split-Path `
    $Config.SourcePage `
    -Leaf
```

Now test whether PnP sees it as a modern client-side page.

```powershell
$PageType = "Unknown"


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

    Write-Host "Page was not identified as a Modern client-side page."

}
```

Then inspect the underlying SharePoint file/list item:

```powershell
try {

    $PageItem = Get-PnPFile `
        -Url $Config.SourcePage `
        -AsListItem `
        -Connection $SourceConnection `
        -ErrorAction Stop


    if ($PageType -ne "Modern") {

        $Fields = $PageItem.FieldValues

        if (
            $Fields.ContainsKey("WikiField") -or
            $Fields.ContainsKey("PublishingPageContent")
        ) {

            $PageType = "Classic"

        }

    }

}
catch {

    throw "Could not inspect source page: $($_.Exception.Message)"

}
```

Now your guardrail:

```powershell
Write-Host ""
Write-Host "PAGE TYPE: $PageType"


switch ($PageType) {

    "Modern" {

        Write-Host `
            "Page is already Modern. No conversion will be attempted." `
            -ForegroundColor Green

    }


    "Classic" {

        Write-Host `
            "Classic page detected." `
            -ForegroundColor Yellow

        Write-Host `
            "The original will remain untouched."

    }


    default {

        throw @"
Page type could not be safely identified.

Migration stopped.

Perform manual inspection before continuing.
"@

    }

}
```

That is the behavior you want.

# 10. Preserve the Classic page

If the page is Classic, download a reference copy.

```powershell
if ($PageType -eq "Classic") {

    $Timestamp = Get-Date -Format "yyyyMMdd-HHmmss"

    $BackupDirectory =
        "C:\SPModernization\Backup\$Timestamp"

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


    $PageItem.FieldValues |
        ConvertTo-Json -Depth 10 |
        Set-Content `
            "$BackupDirectory\Page-Metadata.json"


    Write-Host `
        "Classic page reference backup created at:"
    Write-Host `
        $BackupDirectory

}
```

## Important distinction

This is **not a complete SharePoint backup system**.

The real rollback protection is:

> **Leave the Classic SharePoint source untouched.**

If the organization requires a complete recoverable site backup, that must use the organization's SharePoint backup/retention/migration mechanisms.

# 11. Do not immediately run `ConvertTo-PnPPage`

For the first migration, prefer:

```text
Classic.aspx
     remains untouched

         |
         v

Military-Personnel-Flight-Modern.aspx
     newly created
```

instead of:

```text
Classic.aspx
       |
       v
overwrite/transform
       |
       v
hope everything works
```

Page transformation tools can be extremely useful, but Classic pages may contain:

```text
custom JavaScript
Content Editor web parts
Script Editor web parts
old List View web parts
publishing features
custom master pages
unsupported web parts
```

Those are exactly the kinds of things that make automatic conversion unreliable.

The learning objective should therefore be:

> **Analyze first, transform second.**

# 12. Script 03 - Build the Modern destination

Now create the Modern page.

Create:

```powershell
notepad .\Scripts\03-Build-Modern-Target.ps1
```

Connect:

```powershell
$ErrorActionPreference = "Stop"

$Config = Import-PowerShellDataFile `
    "C:\SPModernization\config.psd1"


$TargetConnection = Connect-PnPOnline `
    -Url $Config.TargetSiteUrl `
    -Interactive `
    -ReturnConnection
```

# 13. First check whether the Modern page already exists

Guardrail:

```powershell
$ExistingPage = $null


try {

    $ExistingPage = Get-PnPPage `
        -Identity "$($Config.TargetPageName).aspx" `
        -Connection $TargetConnection `
        -ErrorAction Stop

}
catch {

    # Page doesn't exist.
}
```

Then:

```powershell
if ($ExistingPage) {

    Write-Host `
        "Modern page already exists. Creation skipped." `
        -ForegroundColor Yellow

}
else {

    Add-PnPPage `
        -Name $Config.TargetPageName `
        -LayoutType Article `
        -Connection $TargetConnection

    Write-Host `
        "Modern page created." `
        -ForegroundColor Green

}
```

This makes the operation more **idempotent**.

That means running the script twice should not create:

```text
Military-Personnel-Flight-Modern.aspx
Military-Personnel-Flight-Modern1.aspx
Military-Personnel-Flight-Modern2.aspx
Military-Personnel-Flight-Modern3.aspx
```

# 14. Check whether the target library exists

```powershell
$TargetLibrary = Get-PnPList `
    -Identity $Config.TargetLibrary `
    -Connection $TargetConnection `
    -ErrorAction SilentlyContinue
```

Then:

```powershell
if (-not $TargetLibrary) {

    New-PnPList `
        -Title $Config.TargetLibrary `
        -Template DocumentLibrary `
        -Connection $TargetConnection

    Write-Host "Target document library created."

}
else {

    Write-Host "Target document library already exists."

}
```

Again:

```text
Check first
|
v
Create only if necessary
```

# 15. Script 04 - Pilot migration

Never migrate everything first.

For the MPF structure, choose something manageable such as:

```text
CSS Training
```

or whichever folder has the least content.

The goal is:

```text
Classic

Shared Documents
`-- CSS Training
    |-- Slides
    |-- References
    `-- Documents


             |
             v


Modern

MPF Documents
`-- CSS Training
    |-- Slides
    |-- References
    `-- Documents
```

Exactly the same hierarchy.

# 16. Why use a temporary local copy during the learning phase

There are several SharePoint server-side copy/migration methods, but for teaching purposes a very transparent workflow is:

```text
SharePoint Source
       |
       v
temporary local file
       |
       v
SharePoint Target
```

You can see every stage and log every failure.

However:

> This basic technique is designed primarily to preserve **folder structure and file content**.

It should **not automatically be assumed** to preserve:

```text
version history
Created By
Modified By
original Created date
original Modified date
approval state
retention labels
unique permissions
content types
custom metadata
```

If leadership requires those, stop before bulk migration and establish the preservation requirements first. You may need tenant-native migration tooling or more sophisticated metadata handling.

# 17. Retrieve only the pilot files

Inside `04-Migrate-Pilot.ps1`:

```powershell
$Config = Import-PowerShellDataFile `
    "C:\SPModernization\config.psd1"


$SourceConnection = Connect-PnPOnline `
    -Url $Config.SourceSiteUrl `
    -Interactive `
    -ReturnConnection


$TargetConnection = Connect-PnPOnline `
    -Url $Config.TargetSiteUrl `
    -Interactive `
    -ReturnConnection


$AllFiles = Get-PnPFolderItem `
    -FolderSiteRelativeUrl $Config.SourceLibrary `
    -ItemType File `
    -Recursive `
    -Connection $SourceConnection
```

Filter:

```powershell
$PilotFiles = $AllFiles |
    Where-Object {
        $_.ServerRelativeUrl -like `
        "*$($Config.PilotFolder)/*"
    }


Write-Host `
    "Pilot file count: $($PilotFiles.Count)"
```

Stop if:

```powershell
if ($PilotFiles.Count -eq 0) {
    throw "No files found in pilot folder."
}
```

# 18. Calculate relative paths

Imagine SharePoint gives:

```text
/sites/MPF/Shared Documents/CSS Training/Slides/Training1.pptx
```

We want:

```text
CSS Training/Slides/Training1.pptx
```

Obtain the library root:

```powershell
$SourceList = Get-PnPList `
    -Identity $Config.SourceLibrary `
    -Includes RootFolder `
    -Connection $SourceConnection


$SourceRoot =
    $SourceList.RootFolder.ServerRelativeUrl.TrimEnd('/')
```

Then:

```powershell
foreach ($File in $PilotFiles) {

    $RelativePath =
        $File.ServerRelativeUrl.Substring(
            $SourceRoot.Length
        ).TrimStart('/')


    Write-Host $RelativePath

}
```

You should see:

```text
CSS Training/Slides/Training1.pptx
CSS Training/Slides/Training2.pptx
CSS Training/References/Reference.pdf
```

This is a major checkpoint.

Before copying anything, confirm that these paths look correct.

# 19. Recreate the target folder automatically

Inside your loop:

```powershell
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
```

Then:

```powershell
$TargetFolder =
    $Config.TargetLibrary


if ($RelativeFolder) {

    $TargetFolder =
        "$($Config.TargetLibrary)/$RelativeFolder"

}


Resolve-PnPFolder `
    -SiteRelativePath $TargetFolder `
    -Connection $TargetConnection |
    Out-Null
```

So PowerShell dynamically creates:

```text
MPF Documents
`-- CSS Training
    `-- Slides
```

before uploading the file.

# 20. Download the source file

```powershell
$LocalFile = Join-Path `
    "C:\SPModernization\Temp" `
    $File.Name


Get-PnPFile `
    -Url $File.ServerRelativeUrl `
    -Path "C:\SPModernization\Temp" `
    -FileName $File.Name `
    -AsFile `
    -Force `
    -Connection $SourceConnection
```

# 21. Upload it into the Modern library

```powershell
Add-PnPFile `
    -Path $LocalFile `
    -Folder $TargetFolder `
    -Connection $TargetConnection
```

Then:

```powershell
Remove-Item `
    $LocalFile `
    -Force
```

# 22. Do not copy files without logging

Wrap the operation:

```powershell
$MigrationLog = @()


foreach ($File in $PilotFiles) {

    try {

        # Migration logic here


        $MigrationLog += [PSCustomObject]@{

            Source = $File.ServerRelativeUrl
            Status = "SUCCESS"
            Error  = ""

        }

    }
    catch {

        $MigrationLog += [PSCustomObject]@{

            Source = $File.ServerRelativeUrl
            Status = "FAILED"
            Error  = $_.Exception.Message

        }

    }

}
```

At the end:

```powershell
$Timestamp = Get-Date -Format "yyyyMMdd-HHmmss"


$MigrationLog |
    Export-Csv `
        "C:\SPModernization\Logs\Pilot-$Timestamp.csv" `
        -NoTypeInformation
```

# 23. Add retry logic

SharePoint Online may temporarily throttle requests.

Create a reusable function:

```powershell
function Invoke-WithRetry {

    param (

        [Parameter(Mandatory)]
        [scriptblock]$Operation,

        [int]$MaximumAttempts = 4

    )


    $Attempt = 1


    while ($Attempt -le $MaximumAttempts) {

        try {

            & $Operation

            return

        }
        catch {

            if ($Attempt -eq $MaximumAttempts) {
                throw
            }


            $Delay = [Math]::Pow(2, $Attempt)

            Write-Warning `
                "Attempt $Attempt failed. Retrying in $Delay seconds."


            Start-Sleep `
                -Seconds $Delay


            $Attempt++

        }

    }

}
```

Then instead of:

```powershell
Add-PnPFile ...
```

use:

```powershell
Invoke-WithRetry {

    Add-PnPFile `
        -Path $LocalFile `
        -Folder $TargetFolder `
        -Connection $TargetConnection

}
```

Now a transient SharePoint error does not immediately terminate the entire migration.

# 24. Script 05 - Validate the pilot

Do not proceed simply because PowerShell says:

```text
Completed.
```

Validation should answer:

```text
Did every source file appear at the target?

Do paths match?

Do file sizes match?

Are there missing files?

Are there unexpected extra files?

Can the documents actually open?
```

Collect source:

```powershell
$SourceFiles = Get-PnPFolderItem `
    -FolderSiteRelativeUrl $Config.SourceLibrary `
    -ItemType File `
    -Recursive `
    -Connection $SourceConnection |
    Where-Object {
        $_.ServerRelativeUrl -like `
        "*$($Config.PilotFolder)/*"
    }
```

Collect target:

```powershell
$TargetFiles = Get-PnPFolderItem `
    -FolderSiteRelativeUrl $Config.TargetLibrary `
    -ItemType File `
    -Recursive `
    -Connection $TargetConnection |
    Where-Object {
        $_.ServerRelativeUrl -like `
        "*$($Config.PilotFolder)/*"
    }
```

Then:

```powershell
Write-Host `
    "Source file count: $($SourceFiles.Count)"

Write-Host `
    "Target file count: $($TargetFiles.Count)"
```

If:

```text
Source = 42
Target = 42
```

that is encouraging.

It is not sufficient by itself.

Compare paths and sizes too.

# 25. Pilot approval checkpoint

Before Script 06, manually verify:

```text
CSS Training
|
|-- correct subfolders
|-- correct documents
|-- documents open
|-- filenames match
|-- file counts match
|-- expected permissions work
`-- modern page links work
```

Then have whoever owns the content verify it.

Only after that:

```text
PILOT APPROVED
```

should you proceed.

# 26. Script 06 - Full migration

This is essentially the tested pilot script except:

```powershell
$PilotFiles
```

becomes:

```powershell
$AllFiles
```

But add another major guardrail.

Before copying a target file:

```text
Does target already exist?
```

If not:

```text
COPY
```

If target exists and same size:

```text
SKIP
```

If target exists but differs:

```text
CONFLICT
```

Do **not** automatically overwrite it.

The behavior should be:

```text
SOURCE FILE
    |
    v
Does target exist?
    |
 +--+--+
 |     |
 NO   YES
 |     |
Copy   Same size?
       |
    +--+--+
    |     |
   YES   NO
    |     |
   Skip  CONFLICT
           |
           v
      Human review
```

This makes the migration rerunnable.

# 27. This concept is called idempotency

It is worth learning because it applies well beyond SharePoint.

An idempotent migration script should be safe to rerun.

Bad migration:

```text
Run #1
100 documents

Run #2
200 documents

Run #3
300 documents
```

Good migration:

```text
Run #1
100 migrated

Run #2
100 detected
100 already correct
0 duplicated
```

That is how you want the automation to behave.

# 28. Script 07 - Full validation

After bulk migration produce three reports:

```text
Missing.csv
Unexpected.csv
SizeMismatch.csv
```

The final report might show:

```text
SOURCE FILES       1,482

TARGET FILES       1,482

MATCHED            1,482

MISSING                0

SIZE MISMATCH          0

FAILED                 0
```

That is the sort of result you want to hand your supervisor.

Not:

> "I went through the folders and they seemed to be there."

# 29. Modern page validation comes next

Suppose the source Classic page displayed:

```text
Military Personnel Flight

Personnel Systems Management
AGR
Career Development
CSS Toolkit
CSS Training
Customer Support
Force Management
Wing Talent Manager Consultant
Virtual In/Out Processing
Status
```

The Modern page should contain the corresponding navigation/Quick Links:

```text
AGR
   |
   v
MPF Documents/AGR


CSS Training
   |
   v
MPF Documents/CSS Training


Force Management
   |
   v
MPF Documents/Force Management
```

Do not just validate files.

Validate the **user journey**.

A user should be able to start at the new page, click:

```text
CSS Training
```

and reach the correct Modern location.

# 30. Script 08 - Cutover

This script should be intentionally difficult to run accidentally.

Something like:

```powershell
Write-Warning @"

YOU ARE ABOUT TO PERFORM THE CUTOVER.

Classic content will NOT be deleted.

Modern page:
$($Config.TargetPageName)

"@


$Confirmation =
    Read-Host "Type CUTOVER to continue"


if ($Confirmation -ne "CUTOVER") {

    throw "Cutover cancelled."

}
```

Only then would you update navigation or the homepage.

If this Modern page is intended to become the site homepage, an operation may resemble:

```powershell
Set-PnPHomePage `
    -RootFolderRelativeUrl `
    "SitePages/$($Config.TargetPageName).aspx" `
    -Connection $TargetConnection
```

But only run that if it truly is supposed to become the homepage.

The MPF page could simply be a section page and not the site homepage.

# 31. Do not delete the Classic page during cutover

The architecture should temporarily look like:

```text
                  USERS
                    |
                    v
          +-----------------+
          |   Modern Page   |
          +--------+--------+
                   |
                   v
           Modern Libraries


              meanwhile


         +------------------+
         | Classic Page     |
         | preserved        |
         +------------------+
```

If something serious happens:

```text
Modern
   |
   v
problem detected
   |
   v
restore navigation
   |
   v
Classic remains available
```

That is the rollback plan.

# 32. When should Classic eventually be deleted?

Not during the migration.

Treat retirement as a separate project stage:

```text
Migration
   |
   v
Validation
   |
   v
User Acceptance Testing
   |
   v
Production cutover
   |
   v
Observation period
   |
   v
Supervisor/content-owner approval
   |
   v
Archive/retire Classic
```

That prevents a migration problem from turning into data loss.

# 33. The complete execution order

For a new specialist, this is the runbook to put on the desk:

```text
00-Preflight.ps1

STOP
Review results.


        |
        v


01-Inventory.ps1

STOP
Inspect:
- lists
- libraries
- navigation
- folders
- files


        |
        v


02-Classify-And-Backup-Page.ps1

IF MODERN:
    Do not convert.

IF CLASSIC:
    Preserve source.
    Plan new Modern page.

IF UNKNOWN:
    STOP.


        |
        v


03-Build-Modern-Target.ps1

STOP
Visually inspect target.


        |
        v


04-Migrate-Pilot.ps1

Migrate ONE folder.


        |
        v


05-Validate-Pilot.ps1

STOP.

Have content owner verify.


        |
        v


06-Migrate-All.ps1

Bulk migration.


        |
        v


07-Validate-All.ps1

Missing = 0
Failures = 0
Unexpected = reviewed


        |
        v


USER ACCEPTANCE TEST


        |
        v


08-Cutover.ps1


        |
        v


MONITOR


        |
        v


Archive Classic later
```

That is how to teach someone to perform this migration.

# 34. Do not run all scripts at once initially

For the first migration:

**No.**

Run them separately.

You specifically want opportunities where the operator must stop and inspect the result.

For example:

```text
01 Inventory
```

might reveal:

> `Customer Support` is not a folder at all. It is another SharePoint page.

That discovery might completely change what Script 03 needs to build.

If you had one giant script, it could simply continue doing the wrong thing.

# 35. Later you can create a master script

After the procedure is mature, you could create:

```text
Invoke-SharePointModernization.ps1
```

that does:

```powershell
& ".\Scripts\00-Preflight.ps1"

Read-Host "Review preflight and press ENTER"


& ".\Scripts\01-Inventory.ps1"

Read-Host "Review inventory and press ENTER"


& ".\Scripts\02-Classify-And-Backup-Page.ps1"

Read-Host "Review page classification and press ENTER"


& ".\Scripts\03-Build-Modern-Target.ps1"

Read-Host "Review target and press ENTER"
```

Notice something important.

Even the master script contains **checkpoints**.

Automation does not mean removing judgment.

# 36. Common failure scenarios the next specialist should understand

| Failure | Meaning | Response |
|---|---|---|
| `401 Unauthorized` | Authentication failure | Reauthenticate/check tenant configuration |
| `403 Access Denied` | Insufficient SharePoint/app permissions | Verify site admin/app permissions |
| `404 File Not Found` | Incorrect library/path/URL | Inspect actual server-relative URL |
| `429 Too Many Requests` | SharePoint throttling | Retry with exponential delay |
| `503 Service Unavailable` | Temporary service condition | Retry |
| Page type `Unknown` | Script cannot classify safely | Stop and investigate manually |
| Target file already exists | Previous run/user-created file | Compare before overwrite |
| Same filename, different size | Potential conflict | Log and manually review |
| Script Editor web part | Modern equivalent may not exist | Redesign |
| Missing metadata | Copy only preserved file content | Add metadata migration strategy |
| Version history missing | Basic upload did not preserve versions | Use appropriate migration method |
| Unique permissions missing | Permissions were not migrated | Design permission migration separately |
| Checked-out document | Source state may interfere | Resolve checkout/content-owner issue |
| Retention/record restriction | Compliance policy applies | Coordinate with SharePoint/compliance admin |

# 37. An important lesson about "all content"

Before a supervisor says:

> "Move everything."

Ask exactly what **everything** means.

There are at least six different layers:

```text
1. Folders
2. Files
3. Metadata
4. Version history
5. Permissions
6. Page functionality
```

Those are not the same migration problem.

Moving:

```text
AGR/application.pdf
```

is relatively straightforward.

Preserving:

```text
Created:       March 2022
Created By:    TSgt Smith
Modified:      June 2025
Modified By:   SSgt Jones
Version:       8.3
Approval:      Approved
Content Type:  MPF Policy
Permissions:   MPF only
```

is a much more sophisticated migration.

Get that requirement decided **before Script 06**.

# 38. The first real session should stop after Script 02

When next at the Knowledge Management shop, do only:

```powershell
.\Scripts\00-Preflight.ps1
```

then:

```powershell
.\Scripts\01-Inventory.ps1
```

then:

```powershell
.\Scripts\02-Classify-And-Backup-Page.ps1
```

And stop.

Do not migrate yet.

Those three steps will tell you:

```text
What type of SharePoint environment is this?

What exactly is the Classic page?

What libraries exist?

What does each menu item actually point to?

How many folders?

How many documents?

What is the actual structure?

Are there custom web parts?

Is the destination a different site or the same site?
```

Once those results are known, that is the point where Scripts 03-08 should be tailored to the actual SharePoint environment rather than guessed.

The larger lesson is that this is not merely a collection of PowerShell commands. It is a **controlled, restartable, auditable migration pipeline** built around:

```text
preconditions
discovery
decision logic
dry runs
idempotency
retries
validation
logging
cutover
rollback
```

Those same principles transfer directly to database migrations, cloud migrations, ETL pipelines, and Knowledge Management automation.
