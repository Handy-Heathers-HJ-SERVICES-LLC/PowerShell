<#
.SYNOPSIS
    Synchronizes GitHub issues and milestones with external project management tools.

.DESCRIPTION
    Exports GitHub issues to a CSV file with task management fields (task name, priority,
    status, and deadlines) suitable for import into tools like Microsoft Project, Excel,
    or any CSV-compatible project management application.

    Optionally imports a CSV produced by an external tool and updates GitHub issue labels
    and milestones based on the task status columns in that CSV.

.PARAMETER Owner
    The GitHub repository owner (user or organization name).

.PARAMETER Repo
    The GitHub repository name.

.PARAMETER Token
    A GitHub Personal Access Token (PAT) with repo read/write access.

.PARAMETER ExportPath
    Path for the exported CSV file. Default: './github-tasks.csv'.

.PARAMETER ImportPath
    Path to an external-tool CSV to import back into GitHub.
    When supplied, the script reads the CSV and updates matching GitHub issues.

.PARAMETER State
    Filter issues by state when exporting. Accepted values: 'open', 'closed', 'all'.
    Default: 'open'.

.PARAMETER MilestoneNumber
    When specified, exports only issues belonging to this milestone number.

.PARAMETER LabelFilter
    One or more label names to filter exported issues. If omitted, all issues matching
    -State are exported.

.EXAMPLE
    # Export open issues to CSV
    .\Sync-GitHubTasks.ps1 -Owner 'Handy-Heathers-HJ-SERVICES-LLC' -Repo 'PowerShell' `
        -Token $env:GITHUB_TOKEN -ExportPath './tasks.csv'

.EXAMPLE
    # Export issues for a specific milestone
    .\Sync-GitHubTasks.ps1 -Owner 'Handy-Heathers-HJ-SERVICES-LLC' -Repo 'PowerShell' `
        -Token $env:GITHUB_TOKEN -MilestoneNumber 3 -ExportPath './milestone3-tasks.csv'

.EXAMPLE
    # Import updated tasks from CSV and sync back to GitHub
    .\Sync-GitHubTasks.ps1 -Owner 'Handy-Heathers-HJ-SERVICES-LLC' -Repo 'PowerShell' `
        -Token $env:GITHUB_TOKEN -ImportPath './tasks-updated.csv'

.NOTES
    CSV columns used for import:
      IssueNumber, ExternalStatus, ExternalPriority, DueDate
    Any unrecognised issue numbers are skipped with a warning.
#>

[CmdletBinding(DefaultParameterSetName = 'Export')]
param(
    [Parameter(Mandatory)]
    [string] $Owner,

    [Parameter(Mandatory)]
    [string] $Repo,

    [Parameter()]
    [string] $Token = $env:GITHUB_TOKEN,

    [Parameter(ParameterSetName = 'Export')]
    [string] $ExportPath = './github-tasks.csv',

    [Parameter(ParameterSetName = 'Import', Mandatory)]
    [string] $ImportPath,

    [Parameter(ParameterSetName = 'Export')]
    [ValidateSet('open', 'closed', 'all')]
    [string] $State = 'open',

    [Parameter(ParameterSetName = 'Export')]
    [int] $MilestoneNumber,

    [Parameter(ParameterSetName = 'Export')]
    [string[]] $LabelFilter
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region Helpers

function Get-GitHubHeaders {
    param([string] $AuthToken)
    $headers = @{
        'Accept'               = 'application/vnd.github+json'
        'X-GitHub-Api-Version' = '2022-11-28'
    }
    if ($AuthToken) {
        $headers['Authorization'] = "Bearer $AuthToken"
    }
    return $headers
}

function Invoke-GitHubApi {
    param(
        [string]    $Uri,
        [hashtable] $Headers,
        [string]    $Method = 'GET',
        [string]    $Body   = $null
    )
    $results = @()
    $page    = 1
    do {
        $pagedUri = "$Uri$(if ($Uri -match '\?') { '&' } else { '?' })per_page=100&page=$page"
        Write-Verbose "GET $pagedUri"
        $invokeParams = @{
            Uri     = $pagedUri
            Headers = $Headers
            Method  = $Method
        }
        if ($Body -and $Method -ne 'GET') {
            $invokeParams['Body']        = $Body
            $invokeParams['ContentType'] = 'application/json'
        }
        $response = Invoke-RestMethod @invokeParams
        if ($response -is [array]) {
            $results += $response
        } else {
            $results += @($response)
        }
        $page++
    } while ($Method -eq 'GET' -and $response -is [array] -and $response.Count -eq 100)
    return $results
}

function Get-PriorityFromLabels {
    param([array] $Labels)
    $names = $Labels | ForEach-Object { $_.name.ToLower() }
    if ($names -contains 'critical' -or $names -contains 'priority: critical') { return 'Critical' }
    if ($names -contains 'high'     -or $names -contains 'priority: high')     { return 'High' }
    if ($names -contains 'medium'   -or $names -contains 'priority: medium')   { return 'Medium' }
    if ($names -contains 'low'      -or $names -contains 'priority: low')      { return 'Low' }
    return 'Normal'
}

function Get-StatusFromIssue {
    param($Issue)
    if ($Issue.state -eq 'closed') { return 'Closed' }
    $names = $Issue.labels | ForEach-Object { $_.name.ToLower() }
    if ($names -contains 'in progress' -or $names -contains 'status: in progress') { return 'In Progress' }
    if ($names -contains 'blocked'     -or $names -contains 'status: blocked')     { return 'Blocked' }
    if ($names -contains 'review'      -or $names -contains 'status: review')      { return 'In Review' }
    return 'Open'
}

#endregion

#region Export

function Export-IssuesToCsv {
    param(
        [string]    $Owner,
        [string]    $Repo,
        [hashtable] $Headers,
        [string]    $State,
        [int]       $MilestoneNumber,
        [string[]]  $LabelFilter,
        [string]    $ExportPath
    )

    $uri = "https://api.github.com/repos/$Owner/$Repo/issues?state=$State&sort=created&direction=asc"
    if ($MilestoneNumber -gt 0) {
        $uri += "&milestone=$MilestoneNumber"
    }
    if ($LabelFilter -and $LabelFilter.Count -gt 0) {
        $uri += "&labels=$($LabelFilter -join ',')"
    }

    Write-Host "Fetching issues from $Owner/$Repo (state=$State)..."
    $allIssues = Invoke-GitHubApi -Uri $uri -Headers $Headers

    # Exclude pull requests (GitHub API returns PRs under /issues)
    $issues = @($allIssues | Where-Object { -not $_.pull_request })
    Write-Host "  Found $($issues.Count) issue(s)."

    $rows = $issues | ForEach-Object {
        $issue = $_
        [PSCustomObject]@{
            IssueNumber   = $issue.number
            TaskName      = $issue.title
            State         = $issue.state
            Status        = Get-StatusFromIssue -Issue $issue
            Priority      = Get-PriorityFromLabels -Labels $issue.labels
            Labels        = ($issue.labels | ForEach-Object { $_.name }) -join '; '
            Assignees     = ($issue.assignees | ForEach-Object { $_.login }) -join '; '
            Milestone     = if ($issue.milestone) { $issue.milestone.title } else { '' }
            MilestoneDue  = if ($issue.milestone -and $issue.milestone.due_on) {
                                ([datetime]$issue.milestone.due_on).ToString('yyyy-MM-dd')
                            } else { '' }
            CreatedDate   = ([datetime]$issue.created_at).ToString('yyyy-MM-dd')
            UpdatedDate   = ([datetime]$issue.updated_at).ToString('yyyy-MM-dd')
            ClosedDate    = if ($issue.closed_at) { ([datetime]$issue.closed_at).ToString('yyyy-MM-dd') } else { '' }
            URL           = $issue.html_url
            Body          = ($issue.body -replace '[\r\n]+', ' ').Trim()
        }
    }

    $rows | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host "Exported $($rows.Count) task(s) to: $ExportPath"
}

#endregion

#region Import / Sync back to GitHub

function Import-TasksFromCsv {
    param(
        [string]    $Owner,
        [string]    $Repo,
        [hashtable] $Headers,
        [string]    $ImportPath
    )

    if (-not (Test-Path $ImportPath)) {
        throw "Import file not found: $ImportPath"
    }

    $rows = Import-Csv -Path $ImportPath -Encoding UTF8
    Write-Host "Imported $($rows.Count) row(s) from: $ImportPath"

    $updatedCount = 0
    $skippedCount = 0

    foreach ($row in $rows) {
        if (-not $row.IssueNumber) {
            Write-Warning "Row missing IssueNumber, skipping."
            $skippedCount++
            continue
        }

        $issueNumber = [int]$row.IssueNumber
        Write-Verbose "Processing issue #$issueNumber"

        # Fetch current issue state
        try {
            $currentIssue = Invoke-RestMethod `
                -Uri "https://api.github.com/repos/$Owner/$Repo/issues/$issueNumber" `
                -Headers $Headers -Method GET
        } catch {
            Write-Warning "Issue #$issueNumber not found in $Owner/$Repo — skipping."
            $skippedCount++
            continue
        }

        $patchBody    = @{}
        $labelsToAdd  = @($currentIssue.labels | ForEach-Object { $_.name })

        # Map ExternalStatus to GitHub state and labels
        if ($row.PSObject.Properties['ExternalStatus'] -and $row.ExternalStatus) {
            switch ($row.ExternalStatus.Trim().ToLower()) {
                'closed'      { $patchBody['state'] = 'closed' }
                'done'        { $patchBody['state'] = 'closed' }
                'complete'    { $patchBody['state'] = 'closed' }
                'completed'   { $patchBody['state'] = 'closed' }
                'in progress' { $labelsToAdd = @($labelsToAdd | Where-Object { $_ -ne 'in progress' }) + 'in progress' }
                'blocked'     { $labelsToAdd = @($labelsToAdd | Where-Object { $_ -ne 'blocked' })     + 'blocked' }
                'open'        { $patchBody['state'] = 'open' }
                default       { Write-Verbose "  Unrecognized ExternalStatus '$($row.ExternalStatus)' for #$issueNumber — status unchanged." }
            }
        }

        # Map ExternalPriority to labels
        if ($row.PSObject.Properties['ExternalPriority'] -and $row.ExternalPriority) {
            $priorityLabels = @('critical', 'high', 'medium', 'low', 'normal',
                                'priority: critical', 'priority: high', 'priority: medium', 'priority: low')
            $labelsToAdd = @($labelsToAdd | Where-Object { $priorityLabels -notcontains $_.ToLower() })
            $newPriorityLabel = "priority: $($row.ExternalPriority.Trim().ToLower())"
            $labelsToAdd += $newPriorityLabel
        }

        $patchBody['labels'] = $labelsToAdd

        # Map DueDate to milestone (create milestone if it doesn't exist)
        if ($row.PSObject.Properties['DueDate'] -and $row.DueDate) {
            $dueDateStr = $row.DueDate.Trim()
            if ($dueDateStr) {
                try {
                    $dueDate = [datetime]::Parse($dueDateStr)
                    # Try to find an existing milestone with this due date
                    $milestones = Invoke-GitHubApi `
                        -Uri "https://api.github.com/repos/$Owner/$Repo/milestones?state=open" `
                        -Headers $Headers
                    $matchingMilestone = $milestones | Where-Object {
                        $_.due_on -and ([datetime]$_.due_on).Date -eq $dueDate.Date
                    } | Select-Object -First 1
                    if ($matchingMilestone) {
                        $patchBody['milestone'] = $matchingMilestone.number
                    } else {
                        # Create a new milestone
                        $msBody = @{
                            title  = "Due $($dueDate.ToString('yyyy-MM-dd'))"
                            due_on = $dueDate.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'")
                        } | ConvertTo-Json
                        $newMs = Invoke-RestMethod `
                            -Uri "https://api.github.com/repos/$Owner/$Repo/milestones" `
                            -Headers $Headers -Method POST -Body $msBody -ContentType 'application/json'
                        $patchBody['milestone'] = $newMs.number
                        Write-Verbose "  Created milestone '$($newMs.title)' (#$($newMs.number))"
                    }
                } catch {
                    Write-Warning "  Could not parse DueDate '$dueDateStr' for issue #$issueNumber — skipping due date."
                }
            }
        }

        # Only send PATCH if there are changes
        if ($patchBody.Count -gt 0) {
            $bodyJson = $patchBody | ConvertTo-Json -Depth 5
            Invoke-RestMethod `
                -Uri "https://api.github.com/repos/$Owner/$Repo/issues/$issueNumber" `
                -Headers $Headers -Method PATCH -Body $bodyJson -ContentType 'application/json' | Out-Null
            Write-Host "  Updated issue #$issueNumber"
            $updatedCount++
        } else {
            Write-Verbose "  No changes for issue #$issueNumber."
            $skippedCount++
        }
    }

    Write-Host "Sync complete — updated: $updatedCount, skipped/unchanged: $skippedCount"
}

#endregion

# ── Main ──────────────────────────────────────────────────────────────────────

if (-not $Token) {
    throw 'A GitHub token is required. Supply -Token or set $env:GITHUB_TOKEN.'
}

$headers = Get-GitHubHeaders -AuthToken $Token

if ($PSCmdlet.ParameterSetName -eq 'Import') {
    Write-Host "Mode: Import — syncing CSV tasks back to GitHub..."
    Import-TasksFromCsv -Owner $Owner -Repo $Repo -Headers $headers -ImportPath $ImportPath
} else {
    Write-Host "Mode: Export — exporting GitHub issues to CSV..."
    $milestoneParam = if ($PSBoundParameters.ContainsKey('MilestoneNumber')) { $MilestoneNumber } else { 0 }
    Export-IssuesToCsv `
        -Owner $Owner -Repo $Repo -Headers $headers `
        -State $State -MilestoneNumber $milestoneParam `
        -LabelFilter $LabelFilter -ExportPath $ExportPath
}

Write-Host "Done."
