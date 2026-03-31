<#
.SYNOPSIS
    Generates a daily or weekly status update report from a GitHub repository.

.DESCRIPTION
    Fetches open/closed issues, pull requests, and commits from a GitHub repository
    for a given timeframe (daily or weekly) and generates a Markdown report.
    Optionally uploads the report as a file to the repository or sends it via email.

.PARAMETER Owner
    The GitHub repository owner (user or organization name).

.PARAMETER Repo
    The GitHub repository name.

.PARAMETER Token
    A GitHub Personal Access Token (PAT) with repo read access.

.PARAMETER Timeframe
    The reporting timeframe. Accepted values: 'Daily', 'Weekly'. Default: 'Weekly'.

.PARAMETER OutputPath
    Path where the Markdown report file will be saved. Default: './status-report.md'.

.PARAMETER SendEmail
    If specified, sends the report via email using Send-MailMessage.

.PARAMETER EmailTo
    The recipient email address. Required when -SendEmail is used.

.PARAMETER EmailFrom
    The sender email address. Required when -SendEmail is used.

.PARAMETER SmtpServer
    The SMTP server hostname. Required when -SendEmail is used.

.PARAMETER UploadToRepo
    If specified, uploads the generated report as a file to the GitHub repository.

.PARAMETER UploadBranch
    The branch to upload the report to. Default: 'main'. Only used with -UploadToRepo.

.EXAMPLE
    .\Get-StatusUpdate.ps1 -Owner 'Handy-Heathers-HJ-SERVICES-LLC' -Repo 'PowerShell' `
        -Token $env:GITHUB_TOKEN -Timeframe 'Weekly' -OutputPath './weekly-report.md'

.EXAMPLE
    .\Get-StatusUpdate.ps1 -Owner 'Handy-Heathers-HJ-SERVICES-LLC' -Repo 'PowerShell' `
        -Token $env:GITHUB_TOKEN -Timeframe 'Daily' -SendEmail `
        -EmailTo 'team@example.com' -EmailFrom 'reports@example.com' -SmtpServer 'smtp.example.com'

.NOTES
    Requires a GitHub Personal Access Token stored in the -Token parameter or $env:GITHUB_TOKEN.
#>

[CmdletBinding(DefaultParameterSetName = 'File')]
param(
    [Parameter(Mandatory)]
    [string] $Owner,

    [Parameter(Mandatory)]
    [string] $Repo,

    [Parameter()]
    [string] $Token = $env:GITHUB_TOKEN,

    [Parameter()]
    [ValidateSet('Daily', 'Weekly')]
    [string] $Timeframe = 'Weekly',

    [Parameter()]
    [string] $OutputPath = './status-report.md',

    [Parameter(ParameterSetName = 'Email')]
    [switch] $SendEmail,

    [Parameter(ParameterSetName = 'Email', Mandatory)]
    [string] $EmailTo,

    [Parameter(ParameterSetName = 'Email', Mandatory)]
    [string] $EmailFrom,

    [Parameter(ParameterSetName = 'Email', Mandatory)]
    [string] $SmtpServer,

    [Parameter()]
    [switch] $UploadToRepo,

    [Parameter()]
    [string] $UploadBranch = 'main'
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
        [string] $Uri,
        [hashtable] $Headers,
        [string] $Method = 'GET'
    )
    $results = @()
    $page = 1
    do {
        $pagedUri = "$Uri$(if ($Uri -match '\?') { '&' } else { '?' })per_page=100&page=$page"
        Write-Verbose "GET $pagedUri"
        $response = Invoke-RestMethod -Uri $pagedUri -Headers $Headers -Method $Method
        if ($response -is [array]) {
            $results += $response
        } else {
            $results += @($response)
        }
        $page++
    } while ($response -is [array] -and $response.Count -eq 100)
    return $results
}

#endregion

#region Data Fetching

function Get-Issues {
    param(
        [string] $Owner,
        [string] $Repo,
        [hashtable] $Headers,
        [datetime] $Since
    )
    $sinceStr = $Since.ToString('yyyy-MM-ddTHH:mm:ssZ')
    $openIssues  = Invoke-GitHubApi -Uri "https://api.github.com/repos/$Owner/$Repo/issues?state=open&since=$sinceStr&sort=updated&direction=desc" -Headers $Headers
    $closedIssues = Invoke-GitHubApi -Uri "https://api.github.com/repos/$Owner/$Repo/issues?state=closed&since=$sinceStr&sort=updated&direction=desc" -Headers $Headers
    # Filter out pull requests (GitHub returns PRs as issues)
    $openIssues   = @($openIssues  | Where-Object { -not $_.pull_request })
    $closedIssues = @($closedIssues | Where-Object { -not $_.pull_request })
    return [PSCustomObject]@{ Open = $openIssues; Closed = $closedIssues }
}

function Get-PullRequests {
    param(
        [string] $Owner,
        [string] $Repo,
        [hashtable] $Headers,
        [datetime] $Since
    )
    $openPRs   = Invoke-GitHubApi -Uri "https://api.github.com/repos/$Owner/$Repo/pulls?state=open&sort=updated&direction=desc" -Headers $Headers
    $closedPRs = Invoke-GitHubApi -Uri "https://api.github.com/repos/$Owner/$Repo/pulls?state=closed&sort=updated&direction=desc" -Headers $Headers

    $openPRs   = @($openPRs   | Where-Object { [datetime]$_.updated_at -ge $Since })
    $closedPRs = @($closedPRs | Where-Object { [datetime]$_.updated_at -ge $Since })
    return [PSCustomObject]@{ Open = $openPRs; Closed = $closedPRs }
}

function Get-RecentCommits {
    param(
        [string] $Owner,
        [string] $Repo,
        [hashtable] $Headers,
        [datetime] $Since
    )
    $sinceStr = $Since.ToString('yyyy-MM-ddTHH:mm:ssZ')
    $commits = Invoke-GitHubApi -Uri "https://api.github.com/repos/$Owner/$Repo/commits?since=$sinceStr" -Headers $Headers
    return $commits
}

#endregion

#region Report Generation

function New-MarkdownReport {
    param(
        [string]   $Owner,
        [string]   $Repo,
        [string]   $Timeframe,
        [datetime] $Since,
        [datetime] $Now,
        $Issues,
        $PullRequests,
        $Commits
    )

    $dateRange = "$($Since.ToString('yyyy-MM-dd')) to $($Now.ToString('yyyy-MM-dd'))"
    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add("# $Timeframe Status Report: $Owner/$Repo")
    $lines.Add("")
    $lines.Add("**Report Period:** $dateRange  ")
    $lines.Add("**Generated:** $($Now.ToString('yyyy-MM-dd HH:mm:ss')) UTC")
    $lines.Add("")

    # Summary table
    $lines.Add("## Summary")
    $lines.Add("")
    $lines.Add("| Category | Count |")
    $lines.Add("|----------|-------|")
    $lines.Add("| Open Issues | $($Issues.Open.Count) |")
    $lines.Add("| Closed Issues | $($Issues.Closed.Count) |")
    $lines.Add("| Open Pull Requests | $($PullRequests.Open.Count) |")
    $lines.Add("| Closed/Merged Pull Requests | $($PullRequests.Closed.Count) |")
    $lines.Add("| Commits | $($Commits.Count) |")
    $lines.Add("")

    # Issues
    $lines.Add("## Issues")
    $lines.Add("")
    if ($Issues.Open.Count -gt 0) {
        $lines.Add("### Open Issues ($($Issues.Open.Count))")
        $lines.Add("")
        foreach ($issue in $Issues.Open) {
            $labels = ($issue.labels | ForEach-Object { "``$($_.name)``" }) -join ' '
            $labelStr = if ($labels) { " — $labels" } else { '' }
            $lines.Add("- [#$($issue.number)]($($issue.html_url)) **$($issue.title)**$labelStr")
        }
        $lines.Add("")
    } else {
        $lines.Add("_No open issues in this period._")
        $lines.Add("")
    }

    if ($Issues.Closed.Count -gt 0) {
        $lines.Add("### Closed Issues ($($Issues.Closed.Count))")
        $lines.Add("")
        foreach ($issue in $Issues.Closed) {
            $lines.Add("- [#$($issue.number)]($($issue.html_url)) ~~$($issue.title)~~")
        }
        $lines.Add("")
    } else {
        $lines.Add("_No issues closed in this period._")
        $lines.Add("")
    }

    # Pull Requests
    $lines.Add("## Pull Requests")
    $lines.Add("")
    if ($PullRequests.Open.Count -gt 0) {
        $lines.Add("### Open Pull Requests ($($PullRequests.Open.Count))")
        $lines.Add("")
        foreach ($pr in $PullRequests.Open) {
            $lines.Add("- [#$($pr.number)]($($pr.html_url)) **$($pr.title)** (by @$($pr.user.login))")
        }
        $lines.Add("")
    } else {
        $lines.Add("_No open pull requests in this period._")
        $lines.Add("")
    }

    if ($PullRequests.Closed.Count -gt 0) {
        $lines.Add("### Closed/Merged Pull Requests ($($PullRequests.Closed.Count))")
        $lines.Add("")
        foreach ($pr in $PullRequests.Closed) {
            $mergeIcon = if ($pr.merged_at) { '✅' } else { '❌' }
            $lines.Add("- $mergeIcon [#$($pr.number)]($($pr.html_url)) $($pr.title) (by @$($pr.user.login))")
        }
        $lines.Add("")
    } else {
        $lines.Add("_No pull requests closed/merged in this period._")
        $lines.Add("")
    }

    # Commits
    $lines.Add("## Recent Commits ($($Commits.Count))")
    $lines.Add("")
    if ($Commits.Count -gt 0) {
        foreach ($commit in $Commits) {
            $shortSha = $commit.sha.Substring(0, 7)
            $message  = ($commit.commit.message -split "`n")[0]
            $author   = $commit.commit.author.name
            $date     = ([datetime]$commit.commit.author.date).ToString('yyyy-MM-dd')
            $lines.Add("- [``$shortSha``]($($commit.html_url)) $message — *$author* ($date)")
        }
        $lines.Add("")
    } else {
        $lines.Add("_No commits in this period._")
        $lines.Add("")
    }

    return $lines -join "`n"
}

#endregion

#region Upload to Repo

function Push-ReportToRepo {
    param(
        [string] $Owner,
        [string] $Repo,
        [string] $AuthToken,
        [string] $Branch,
        [string] $Content,
        [string] $Timeframe,
        [datetime] $Now
    )
    $headers = Get-GitHubHeaders -AuthToken $AuthToken
    $fileName = "reports/status-report-$($Now.ToString('yyyy-MM-dd'))-$($Timeframe.ToLower()).md"
    $encodedContent = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($Content))

    # Check if file already exists (to get its SHA for update)
    $existingSha = $null
    try {
        $existing = Invoke-RestMethod -Uri "https://api.github.com/repos/$Owner/$Repo/contents/$fileName" -Headers $headers -Method GET
        $existingSha = $existing.sha
    } catch {
        # File does not exist yet — will create it
    }

    $body = @{
        message = "chore: add $Timeframe status report for $($Now.ToString('yyyy-MM-dd'))"
        content = $encodedContent
        branch  = $Branch
    }
    if ($existingSha) {
        $body['sha'] = $existingSha
    }

    $bodyJson = $body | ConvertTo-Json -Depth 5
    Invoke-RestMethod -Uri "https://api.github.com/repos/$Owner/$Repo/contents/$fileName" `
        -Headers $headers -Method PUT -Body $bodyJson -ContentType 'application/json' | Out-Null
    Write-Host "Report uploaded to repository: $fileName (branch: $Branch)"
}

#endregion

# ── Main ──────────────────────────────────────────────────────────────────────

if (-not $Token) {
    throw 'A GitHub token is required. Supply -Token or set $env:GITHUB_TOKEN.'
}

$now   = [datetime]::UtcNow
$since = if ($Timeframe -eq 'Daily') { $now.AddDays(-1) } else { $now.AddDays(-7) }

$headers = Get-GitHubHeaders -AuthToken $Token

Write-Host "Fetching $Timeframe data for $Owner/$Repo (since $($since.ToString('yyyy-MM-dd HH:mm:ss')) UTC)..."

$issues       = Get-Issues       -Owner $Owner -Repo $Repo -Headers $headers -Since $since
$pullRequests = Get-PullRequests -Owner $Owner -Repo $Repo -Headers $headers -Since $since
$commits      = Get-RecentCommits -Owner $Owner -Repo $Repo -Headers $headers -Since $since

$reportContent = New-MarkdownReport `
    -Owner $Owner -Repo $Repo -Timeframe $Timeframe `
    -Since $since -Now $now `
    -Issues $issues -PullRequests $pullRequests -Commits $commits

# Save report locally
$reportContent | Set-Content -Path $OutputPath -Encoding UTF8
Write-Host "Report saved to: $OutputPath"

# Optional: upload to repository
if ($UploadToRepo) {
    Push-ReportToRepo -Owner $Owner -Repo $Repo -AuthToken $Token `
        -Branch $UploadBranch -Content $reportContent -Timeframe $Timeframe -Now $now
}

# Optional: send via email
if ($SendEmail) {
    $subject = "$Timeframe Status Report — $Owner/$Repo ($($now.ToString('yyyy-MM-dd')))"
    Send-MailMessage -To $EmailTo -From $EmailFrom -Subject $subject `
        -Body $reportContent -SmtpServer $SmtpServer -Encoding UTF8
    Write-Host "Report emailed to: $EmailTo"
}

Write-Host "Done."
