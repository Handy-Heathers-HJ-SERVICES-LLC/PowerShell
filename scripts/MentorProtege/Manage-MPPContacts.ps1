<#
.SYNOPSIS
    Manages contacts for the Mentor-Protégé Program (MPP) using an Excel spreadsheet.

.DESCRIPTION
    This script provides a complete contact management solution for the Mentor-Protégé Program.
    It supports creating, reading, updating, and deleting contact records stored in an Excel (.xlsx)
    file. Contacts can be searched and filtered by name, role, organization, or other criteria.
    The Excel file can also be committed to the GitHub repository for version-controlled, centralized
    storage.

    Required module: ImportExcel (Install-Module -Name ImportExcel -Scope CurrentUser)

.PARAMETER Action
    The action to perform. Valid values:
      Initialize   - Create a new contacts Excel file with the standard template headers.
      Add          - Add a new contact to the spreadsheet.
      Update       - Update an existing contact matched by email address.
      Remove       - Remove a contact matched by email address.
      Search       - Search/filter contacts by Name, Role, or Organization.
      Export       - Export all contacts to a formatted Excel report.
      Import       - Import contacts from an external Excel file, merging with existing data.
      List         - Display all contacts in the console.
      Sync         - Commit and push the Excel file to the configured Git repository.

.PARAMETER ContactsFile
    Path to the Excel (.xlsx) contacts file. Defaults to 'MPP-Contacts.xlsx' in the same
    directory as this script.

.PARAMETER Name
    Contact's full name. Required for Add and used as display for Update/Remove.

.PARAMETER Role
    Contact's role. Valid values: Mentor, Protege. Required for Add.

.PARAMETER Organization
    Contact's organization or company name.

.PARAMETER Email
    Contact's email address. Used as the unique identifier for Update and Remove actions.

.PARAMETER Phone
    Contact's phone number.

.PARAMETER Industry
    Contact's industry or sector.

.PARAMETER Location
    Contact's city, state, or region.

.PARAMETER Notes
    Additional notes about the contact.

.PARAMETER SearchTerm
    The term to search for when Action is 'Search'. Searches across Name, Role, and Organization.

.PARAMETER FilterRole
    Filter results to a specific role when Action is 'Search' or 'List'. Valid values: Mentor, Protege.

.PARAMETER FilterIndustry
    Filter results to a specific industry when Action is 'Search' or 'List'.

.PARAMETER FilterLocation
    Filter results to a specific location when Action is 'Search' or 'List'.

.PARAMETER ImportFile
    Path to an external Excel file to import contacts from when Action is 'Import'.

.PARAMETER ExportFile
    Path for the exported Excel report when Action is 'Export'. Defaults to
    'MPP-Contacts-Export-<timestamp>.xlsx'.

.PARAMETER CommitMessage
    Git commit message used when Action is 'Sync'. Defaults to an auto-generated message with
    timestamp.

.EXAMPLE
    # Create a new contacts file
    .\Manage-MPPContacts.ps1 -Action Initialize

.EXAMPLE
    # Add a new mentor
    .\Manage-MPPContacts.ps1 -Action Add -Name "Jane Smith" -Role Mentor `
        -Organization "Acme Corp" -Email "jane.smith@acme.com" -Phone "555-1234" `
        -Industry "Technology" -Location "Seattle, WA" -Notes "Expert in cloud computing"

.EXAMPLE
    # Add a new protégé
    .\Manage-MPPContacts.ps1 -Action Add -Name "John Doe" -Role Protege `
        -Organization "State University" -Email "john.doe@stateuniversity.edu" `
        -Industry "Engineering" -Location "Austin, TX"

.EXAMPLE
    # Update an existing contact's phone number
    .\Manage-MPPContacts.ps1 -Action Update -Email "jane.smith@acme.com" -Phone "555-9999"

.EXAMPLE
    # Remove a contact by email
    .\Manage-MPPContacts.ps1 -Action Remove -Email "john.doe@stateuniversity.edu"

.EXAMPLE
    # Search contacts by keyword
    .\Manage-MPPContacts.ps1 -Action Search -SearchTerm "Acme"

.EXAMPLE
    # List all mentors in the Technology industry
    .\Manage-MPPContacts.ps1 -Action List -FilterRole Mentor -FilterIndustry "Technology"

.EXAMPLE
    # Export all contacts to a timestamped Excel file
    .\Manage-MPPContacts.ps1 -Action Export

.EXAMPLE
    # Import contacts from another Excel file
    .\Manage-MPPContacts.ps1 -Action Import -ImportFile "C:\Downloads\NewContacts.xlsx"

.EXAMPLE
    # Sync the contacts file to Git
    .\Manage-MPPContacts.ps1 -Action Sync -CommitMessage "Add Q1 2025 contacts"

.NOTES
    Author:  HJ Services LLC
    Version: 1.0.0
    Prerequisites:
      - PowerShell 7.0 or later
      - ImportExcel module: Install-Module -Name ImportExcel -Scope CurrentUser
      - Git (for Sync action)
#>


[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [ValidateSet('Initialize', 'Add', 'Update', 'Remove', 'Search', 'Export', 'Import', 'List', 'Sync')]
    [string] $Action,

    [Parameter()]
    [string] $ContactsFile = (Join-Path $PSScriptRoot 'MPP-Contacts.xlsx'),

    # Contact fields
    [Parameter()]
    [string] $Name,

    [Parameter()]
    [ValidateSet('Mentor', 'Protege')]
    [string] $Role,

    [Parameter()]
    [string] $Organization,

    [Parameter()]
    [string] $Email,

    [Parameter()]
    [string] $Phone,

    [Parameter()]
    [string] $Industry,

    [Parameter()]
    [string] $Location,

    [Parameter()]
    [string] $Notes,

    # Search / filter parameters
    [Parameter()]
    [string] $SearchTerm,

    [Parameter()]
    [ValidateSet('Mentor', 'Protege')]
    [string] $FilterRole,

    [Parameter()]
    [string] $FilterIndustry,

    [Parameter()]
    [string] $FilterLocation,

    # Import / export parameters
    [Parameter()]
    [string] $ImportFile,

    [Parameter()]
    [string] $ExportFile,

    # Sync parameter
    [Parameter()]
    [string] $CommitMessage
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region --- Helper Functions ---

function Assert-ImportExcelModule {
    <#
    .SYNOPSIS
        Ensures the ImportExcel module is available, prompting the user to install it if not.
    #>
    if (-not (Get-Module -Name ImportExcel -ListAvailable)) {
        Write-Warning "The 'ImportExcel' module is required but not installed."
        Write-Host "Run the following command to install it:" -ForegroundColor Yellow
        Write-Host "  Install-Module -Name ImportExcel -Scope CurrentUser" -ForegroundColor Cyan
        throw "Missing required module: ImportExcel"
    }
    Import-Module ImportExcel -ErrorAction Stop
}

function Get-ContactsFromFile {
    <#
    .SYNOPSIS
        Reads all contacts from the Excel file and returns them as an array of PSCustomObjects.
    #>
    param([string] $FilePath)

    if (-not (Test-Path -Path $FilePath)) {
        throw "Contacts file not found: '$FilePath'. Run with -Action Initialize to create it."
    }

    $contacts = Import-Excel -Path $FilePath -WorksheetName 'Contacts' -ErrorAction Stop
    return $contacts
}

function Save-ContactsToFile {
    <#
    .SYNOPSIS
        Writes the supplied array of contact objects back to the Excel file.
    #>
    param(
        [string] $FilePath,
        [object[]] $Contacts
    )

    # Remove and recreate the worksheet so stale rows are cleared
    $excelPackage = Open-ExcelPackage -Path $FilePath -Create
    $worksheetName = 'Contacts'

    # Delete existing worksheet if present
    if ($excelPackage.Workbook.Worksheets[$worksheetName]) {
        $excelPackage.Workbook.Worksheets.Delete($worksheetName)
    }
    Close-ExcelPackage -ExcelPackage $excelPackage -NoSave

    if ($Contacts.Count -gt 0) {
        $Contacts | Export-Excel -Path $FilePath -WorksheetName $worksheetName `
            -AutoSize -BoldTopRow -FreezeTopRow -TableStyle Medium9 -ClearSheet
    } else {
        # Write headers only when there are no contacts
        $emptyRow = [PSCustomObject]@{
            Name         = ''
            Role         = ''
            Organization = ''
            Email        = ''
            Phone        = ''
            Industry     = ''
            Location     = ''
            Notes        = ''
            DateAdded    = ''
            LastUpdated  = ''
        }
        $emptyRow | Export-Excel -Path $FilePath -WorksheetName $worksheetName `
            -AutoSize -BoldTopRow -FreezeTopRow -TableStyle Medium9 -ClearSheet

        # Remove the empty placeholder row
        $pkg = Open-ExcelPackage -Path $FilePath
        $ws  = $pkg.Workbook.Worksheets[$worksheetName]
        if ($ws.Dimension.Rows -gt 1) {
            $ws.DeleteRow(2)
        }
        Close-ExcelPackage -ExcelPackage $pkg -Save
    }
}

function ConvertTo-Contact {
    <#
    .SYNOPSIS
        Creates a new contact PSCustomObject with all standard fields.
    #>
    
    param(
        [string] $Name,
        [string] $Role,
        [string] $Organization,
        [string] $Email,
        [string] $Phone,
        [string] $Industry,
        [string] $Location,
        [string] $Notes,
        [string] $DateAdded    = (Get-Date -Format 'yyyy-MM-dd'),
        [string] $LastUpdated  = (Get-Date -Format 'yyyy-MM-dd')
    )

    return [PSCustomObject]@{
        Name         = $Name
        Role         = $Role
        Organization = $Organization
        Email        = $Email
        Phone        = $Phone
        Industry     = $Industry
        Location     = $Location
        Notes        = $Notes
        DateAdded    = $DateAdded
        LastUpdated  = $LastUpdated
    }
}

function Write-ContactTable {
    <#
    .SYNOPSIS
        Outputs a formatted contact table to the host console.
    #>
    param([object[]] $Contacts)

    if ($null -eq $Contacts -or $Contacts.Count -eq 0) {
        Write-Host "No contacts found." -ForegroundColor Yellow
        return
    }

    $Contacts | Format-Table -Property Name, Role, Organization, Email, Phone, Industry, Location `
        -AutoSize -Wrap
    Write-Host "Total: $($Contacts.Count) contact(s)" -ForegroundColor Cyan
}

#endregion

#region --- Action Handlers ---

function Invoke-Initialize {
    <#
    .SYNOPSIS
        Creates a new contacts Excel file with the standard template.
    #>
    param([string] $FilePath)

    if (Test-Path -Path $FilePath) {
        Write-Warning "Contacts file already exists: '$FilePath'"
        $overwrite = Read-Host "Overwrite? (y/N)"
        if ($overwrite -ne 'y' -and $overwrite -ne 'Y') {
            Write-Host "Initialization cancelled." -ForegroundColor Yellow
            return
        }
        Remove-Item -Path $FilePath -Force
    }

    $templateRows = @(
        ConvertTo-Contact -Name 'Example Mentor' -Role 'Mentor' `
            -Organization 'Sample Corp' -Email 'mentor@example.com' `
            -Phone '555-0100' -Industry 'Technology' -Location 'Seattle, WA' `
            -Notes 'Sample mentor entry — delete before use.'
        ConvertTo-Contact -Name 'Example Protege' -Role 'Protege' `
            -Organization 'State University' -Email 'protege@example.edu' `
            -Phone '555-0200' -Industry 'Engineering' -Location 'Austin, TX' `
            -Notes 'Sample protégé entry — delete before use.'
    )

    $templateRows | Export-Excel -Path $FilePath -WorksheetName 'Contacts' `
        -AutoSize -BoldTopRow -FreezeTopRow -TableStyle Medium9

    Write-Host "Contacts file created: '$FilePath'" -ForegroundColor Green
    Write-Host "Delete the sample rows before adding real contacts." -ForegroundColor Yellow
}

function Invoke-Add {
    <#
    .SYNOPSIS
        Adds a new contact to the spreadsheet.
    #>
    param(
        [string] $FilePath,
        [string] $Name,
        [string] $Role,
        [string] $Organization,
        [string] $Email,
        [string] $Phone,
        [string] $Industry,
        [string] $Location,
        [string] $Notes
    )

    if ([string]::IsNullOrWhiteSpace($Name))  { throw "Parameter -Name is required for action 'Add'." }
    if ([string]::IsNullOrWhiteSpace($Role))  { throw "Parameter -Role is required for action 'Add'." }
    if ([string]::IsNullOrWhiteSpace($Email)) { throw "Parameter -Email is required for action 'Add'." }

    $contacts = Get-ContactsFromFile -FilePath $FilePath

    # Check for duplicate email
    $existingContact = $contacts | Where-Object { $_.Email -eq $Email }
    if ($existingContact) {
        throw "A contact with email '$Email' already exists. Use -Action Update to modify it."
    }

    $newContact = ConvertTo-Contact -Name $Name -Role $Role -Organization $Organization `
        -Email $Email -Phone $Phone -Industry $Industry -Location $Location -Notes $Notes

    $updatedContacts = @($contacts) + @($newContact)
    Save-ContactsToFile -FilePath $FilePath -Contacts $updatedContacts

    Write-Host "Contact added: $Name ($Role)" -ForegroundColor Green
}

function Invoke-Update {
    <#
    .SYNOPSIS
        Updates an existing contact matched by email address.
    #>
    param(
        [string] $FilePath,
        [string] $Email,
        [string] $Name,
        [string] $Role,
        [string] $Organization,
        [string] $Phone,
        [string] $Industry,
        [string] $Location,
        [string] $Notes
    )

    if ([string]::IsNullOrWhiteSpace($Email)) { throw "Parameter -Email is required for action 'Update'." }

    $contacts = Get-ContactsFromFile -FilePath $FilePath

    if (-not ($contacts | Where-Object { $_.Email -eq $Email })) {
        throw "No contact found with email '$Email'."
    }

    $updatedContacts = $contacts | ForEach-Object {
        if ($_.Email -eq $Email) {
            if ($Name)         { $_.Name         = $Name         }
            if ($Role)         { $_.Role         = $Role         }
            if ($Organization) { $_.Organization = $Organization }
            if ($Phone)        { $_.Phone        = $Phone        }
            if ($Industry)     { $_.Industry     = $Industry     }
            if ($Location)     { $_.Location     = $Location     }
            if ($Notes)        { $_.Notes        = $Notes        }
            $_.LastUpdated = (Get-Date -Format 'yyyy-MM-dd')
        }
        $_
    }

    Save-ContactsToFile -FilePath $FilePath -Contacts $updatedContacts
    Write-Host "Contact updated: $Email" -ForegroundColor Green
}

function Invoke-Remove {
    <#
    .SYNOPSIS
        Removes a contact matched by email address.
    #>
    param(
        [string] $FilePath,
        [string] $Email
    )

    if ([string]::IsNullOrWhiteSpace($Email)) { throw "Parameter -Email is required for action 'Remove'." }

    $contacts = Get-ContactsFromFile -FilePath $FilePath
    $before = $contacts.Count
    $updatedContacts = @($contacts | Where-Object { $_.Email -ne $Email })

    if ($updatedContacts.Count -eq $before) {
        throw "No contact found with email '$Email'."
    }

    Save-ContactsToFile -FilePath $FilePath -Contacts $updatedContacts
    Write-Host "Contact removed: $Email" -ForegroundColor Green
}

function Invoke-Search {
    <#
    .SYNOPSIS
        Searches contacts by a keyword or filters by role, industry, or location.
    #>
    param(
        [string] $FilePath,
        [string] $SearchTerm,
        [string] $FilterRole,
        [string] $FilterIndustry,
        [string] $FilterLocation
    )

    $contacts = Get-ContactsFromFile -FilePath $FilePath

    if (-not [string]::IsNullOrWhiteSpace($SearchTerm)) {
        $contacts = $contacts | Where-Object {
            $_.Name         -like "*$SearchTerm*" -or
            $_.Role         -like "*$SearchTerm*" -or
            $_.Organization -like "*$SearchTerm*" -or
            $_.Email        -like "*$SearchTerm*" -or
            $_.Industry     -like "*$SearchTerm*" -or
            $_.Location     -like "*$SearchTerm*" -or
            $_.Notes        -like "*$SearchTerm*"
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($FilterRole)) {
        $contacts = $contacts | Where-Object { $_.Role -eq $FilterRole }
    }

    if (-not [string]::IsNullOrWhiteSpace($FilterIndustry)) {
        $contacts = $contacts | Where-Object { $_.Industry -like "*$FilterIndustry*" }
    }

    if (-not [string]::IsNullOrWhiteSpace($FilterLocation)) {
        $contacts = $contacts | Where-Object { $_.Location -like "*$FilterLocation*" }
    }

    Write-ContactTable -Contacts $contacts
}

function Invoke-List {
    <#
    .SYNOPSIS
        Lists all contacts, with optional role/industry/location filtering.
    #>
    param(
        [string] $FilePath,
        [string] $FilterRole,
        [string] $FilterIndustry,
        [string] $FilterLocation
    )

    Invoke-Search -FilePath $FilePath -FilterRole $FilterRole `
        -FilterIndustry $FilterIndustry -FilterLocation $FilterLocation
}

function Invoke-Export {
    <#
    .SYNOPSIS
        Exports all contacts to a new timestamped Excel report file.
    #>
    param(
        [string] $FilePath,
        [string] $ExportFile
    )

    if ([string]::IsNullOrWhiteSpace($ExportFile)) {
        $timestamp  = Get-Date -Format 'yyyyMMdd-HHmmss'
        $exportDir  = Split-Path -Parent $FilePath
        $ExportFile = Join-Path $exportDir "MPP-Contacts-Export-$timestamp.xlsx"
    }

    $contacts = Get-ContactsFromFile -FilePath $FilePath

    $contacts | Export-Excel -Path $ExportFile -WorksheetName 'Contacts' `
        -AutoSize -BoldTopRow -FreezeTopRow -TableStyle Medium9 `
        -Title 'Mentor-Protégé Program — Contact List' `
        -TitleBold -TitleSize 14

    Write-Host "Contacts exported to: '$ExportFile'" -ForegroundColor Green
    Write-Host "Total records exported: $($contacts.Count)" -ForegroundColor Cyan
}

function Invoke-Import {
    <#
    .SYNOPSIS
        Imports contacts from an external Excel file, merging with existing data.
        Existing contacts matched by email are updated; new contacts are appended.
    #>
    param(
        [string] $FilePath,
        [string] $ImportFile
    )

    if ([string]::IsNullOrWhiteSpace($ImportFile)) { throw "Parameter -ImportFile is required for action 'Import'." }
    if (-not (Test-Path -Path $ImportFile))         { throw "Import file not found: '$ImportFile'." }

    $existingContacts = Get-ContactsFromFile -FilePath $FilePath
    $incomingContacts = Import-Excel -Path $ImportFile -WorksheetName 'Contacts' -ErrorAction Stop

    $addedCount   = 0
    $updatedCount = 0

    foreach ($incoming in $incomingContacts) {
        $existing = $existingContacts | Where-Object { $_.Email -eq $incoming.Email }
        if ($existing) {
            # Update fields from incoming row
            $existing.Name         = $incoming.Name
            $existing.Role         = $incoming.Role
            $existing.Organization = $incoming.Organization
            $existing.Phone        = $incoming.Phone
            $existing.Industry     = $incoming.Industry
            $existing.Location     = $incoming.Location
            $existing.Notes        = $incoming.Notes
            $existing.LastUpdated  = (Get-Date -Format 'yyyy-MM-dd')
            $updatedCount++
        } else {
            $newContact = ConvertTo-Contact `
                -Name         $incoming.Name `
                -Role         $incoming.Role `
                -Organization $incoming.Organization `
                -Email        $incoming.Email `
                -Phone        $incoming.Phone `
                -Industry     $incoming.Industry `
                -Location     $incoming.Location `
                -Notes        $incoming.Notes `
                -DateAdded    (Get-Date -Format 'yyyy-MM-dd')
            $existingContacts = @($existingContacts) + @($newContact)
            $addedCount++
        }
    }

    Save-ContactsToFile -FilePath $FilePath -Contacts $existingContacts

    Write-Host "Import complete:" -ForegroundColor Green
    Write-Host "  Added:   $addedCount contact(s)" -ForegroundColor Cyan
    Write-Host "  Updated: $updatedCount contact(s)" -ForegroundColor Cyan
}

function Invoke-Sync {
    <#
    .SYNOPSIS
        Commits and pushes the contacts Excel file to the Git repository.
    #>
    param(
        [string] $FilePath,
        [string] $CommitMessage
    )

    if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
        throw "Git is not installed or not in PATH. Cannot sync."
    }

    if (-not (Test-Path -Path $FilePath)) {
        throw "Contacts file not found: '$FilePath'. Run -Action Initialize first."
    }

    # Determine the repo root from the contacts file location
    $repoRoot = git -C (Split-Path -Parent $FilePath) rev-parse --show-toplevel 2>&1
    $gitExitCode = $LASTEXITCODE
    if ($gitExitCode -ne 0) {
        throw "The directory containing the contacts file is not inside a Git repository."
    }

    if ([string]::IsNullOrWhiteSpace($CommitMessage)) {
        $CommitMessage = "chore: update MPP contacts - $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
    }

    $relativePath = [System.IO.Path]::GetRelativePath($repoRoot, $FilePath)

    Push-Location $repoRoot
    try {
        git add $relativePath
        if ($LASTEXITCODE -ne 0) {
            throw "'git add' failed with exit code $LASTEXITCODE."
        }

        $statusOutput = git status --porcelain $relativePath
        if ([string]::IsNullOrWhiteSpace($statusOutput)) {
            Write-Host "No changes to commit for '$relativePath'." -ForegroundColor Yellow
        } else {
            git commit -m $CommitMessage
            if ($LASTEXITCODE -ne 0) {
                throw "'git commit' failed with exit code $LASTEXITCODE."
            }

            git push
            if ($LASTEXITCODE -ne 0) {
                throw "'git push' failed with exit code $LASTEXITCODE. Check remote configuration and credentials."
            }

            Write-Host "Contacts file synced to repository." -ForegroundColor Green
        }
    } finally {
        Pop-Location
    }
}

#endregion

#region --- Main Entry Point ---

Assert-ImportExcelModule

switch ($Action) {
    'Initialize' {
        if ($PSCmdlet.ShouldProcess($ContactsFile, 'Initialize contacts file')) {
            Invoke-Initialize -FilePath $ContactsFile
        }
    }
    'Add' {
        if ($PSCmdlet.ShouldProcess($ContactsFile, "Add contact '$Name'")) {
            Invoke-Add -FilePath $ContactsFile -Name $Name -Role $Role `
                -Organization $Organization -Email $Email -Phone $Phone `
                -Industry $Industry -Location $Location -Notes $Notes
        }
    }
    'Update' {
        if ($PSCmdlet.ShouldProcess($ContactsFile, "Update contact '$Email'")) {
            Invoke-Update -FilePath $ContactsFile -Email $Email -Name $Name -Role $Role `
                -Organization $Organization -Phone $Phone `
                -Industry $Industry -Location $Location -Notes $Notes
        }
    }
    'Remove' {
        if ($PSCmdlet.ShouldProcess($ContactsFile, "Remove contact '$Email'")) {
            Invoke-Remove -FilePath $ContactsFile -Email $Email
        }
    }
    'Search' {
        Invoke-Search -FilePath $ContactsFile -SearchTerm $SearchTerm `
            -FilterRole $FilterRole -FilterIndustry $FilterIndustry `
            -FilterLocation $FilterLocation
    }
    'List' {
        Invoke-List -FilePath $ContactsFile -FilterRole $FilterRole `
            -FilterIndustry $FilterIndustry -FilterLocation $FilterLocation
    }
    'Export' {
        if ($PSCmdlet.ShouldProcess($ContactsFile, 'Export contacts')) {
            Invoke-Export -FilePath $ContactsFile -ExportFile $ExportFile
        }
    }
    'Import' {
        if ($PSCmdlet.ShouldProcess($ContactsFile, "Import contacts from '$ImportFile'")) {
            Invoke-Import -FilePath $ContactsFile -ImportFile $ImportFile
        }
    }
    'Sync' {
        if ($PSCmdlet.ShouldProcess($ContactsFile, 'Sync contacts to Git')) {
            Invoke-Sync -FilePath $ContactsFile -CommitMessage $CommitMessage
        }
    }
}

#endregion
