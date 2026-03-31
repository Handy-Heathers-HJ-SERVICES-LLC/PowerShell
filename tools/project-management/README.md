# Project Management Automation Scripts

This directory contains PowerShell scripts that automate project management tasks for the **Handy-Heathers-HJ-SERVICES-LLC/PowerShell** repository using the GitHub REST API.

---

## Prerequisites

- **PowerShell 7+** (cross-platform — runs on Windows, macOS, and Linux)
- A **GitHub Personal Access Token (PAT)** with at minimum:
  - `repo` scope for reading issues, pull requests, and commits
  - `repo` scope write access when using the import/sync-back feature or `-UploadToRepo`
- Internet access to reach `api.github.com`

Store the token in an environment variable to avoid embedding it in commands:

```powershell
$env:GITHUB_TOKEN = '<your-token-here>'
```

---

## Scripts

### 1. `Get-StatusUpdate.ps1` — Daily / Weekly Status Reports

Generates a Markdown status report that summarises repository activity over the past day or week, including open/closed issues, pull requests, and recent commits.

#### Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `Owner` | `string` | ✅ | — | GitHub owner (user or org) |
| `Repo` | `string` | ✅ | — | Repository name |
| `Token` | `string` | ✅* | `$env:GITHUB_TOKEN` | GitHub PAT |
| `Timeframe` | `string` | ❌ | `Weekly` | `Daily` or `Weekly` |
| `OutputPath` | `string` | ❌ | `./status-report.md` | Path for the saved report |
| `UploadToRepo` | `switch` | ❌ | — | Upload report to the repository |
| `UploadBranch` | `string` | ❌ | `main` | Branch to upload to |
| `SendEmail` | `switch` | ❌ | — | Email the report |
| `EmailTo` | `string` | ❌** | — | Recipient address |
| `EmailFrom` | `string` | ❌** | — | Sender address |
| `SmtpServer` | `string` | ❌** | — | SMTP server hostname |

\* Required if `$env:GITHUB_TOKEN` is not set.  
\*\* Required when `-SendEmail` is used.

#### Examples

```powershell
# Generate a weekly report and save it locally
.\Get-StatusUpdate.ps1 `
    -Owner 'Handy-Heathers-HJ-SERVICES-LLC' `
    -Repo  'PowerShell' `
    -Token $env:GITHUB_TOKEN `
    -Timeframe 'Weekly' `
    -OutputPath './weekly-report.md'
```

```powershell
# Generate a daily report and upload it to the repository
.\Get-StatusUpdate.ps1 `
    -Owner 'Handy-Heathers-HJ-SERVICES-LLC' `
    -Repo  'PowerShell' `
    -Token $env:GITHUB_TOKEN `
    -Timeframe 'Daily' `
    -UploadToRepo `
    -UploadBranch 'main'
```

```powershell
# Generate a weekly report and email it
.\Get-StatusUpdate.ps1 `
    -Owner      'Handy-Heathers-HJ-SERVICES-LLC' `
    -Repo       'PowerShell' `
    -Token      $env:GITHUB_TOKEN `
    -Timeframe  'Weekly' `
    -SendEmail `
    -EmailTo    'team@example.com' `
    -EmailFrom  'reports@example.com' `
    -SmtpServer 'smtp.example.com'
```

#### Report Format

The generated Markdown report includes:

- **Summary table** — counts for open/closed issues, open/merged PRs, and commits.
- **Issues** — open issues with labels and links, closed issues with links.
- **Pull Requests** — open PRs with author, closed/merged PRs flagged with ✅ (merged) or ❌ (closed without merge).
- **Recent Commits** — short SHA link, commit message (first line), author, and date.

---

### 2. `Sync-GitHubTasks.ps1` — Task Synchronization

Exports GitHub issues to a CSV file suitable for import into Microsoft Project, Excel, or any CSV-compatible project management tool. Can also re-import an updated CSV to synchronize changes back to GitHub.

#### Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `Owner` | `string` | ✅ | — | GitHub owner (user or org) |
| `Repo` | `string` | ✅ | — | Repository name |
| `Token` | `string` | ✅* | `$env:GITHUB_TOKEN` | GitHub PAT |
| `ExportPath` | `string` | ❌ | `./github-tasks.csv` | Path for the exported CSV |
| `ImportPath` | `string` | ❌** | — | CSV file to import back into GitHub |
| `State` | `string` | ❌ | `open` | `open`, `closed`, or `all` |
| `MilestoneNumber` | `int` | ❌ | — | Limit export to a specific milestone |
| `LabelFilter` | `string[]` | ❌ | — | Filter by one or more label names |

\* Required if `$env:GITHUB_TOKEN` is not set.  
\*\* When `-ImportPath` is supplied, the script runs in *import* mode; all export parameters are ignored.

#### Examples

```powershell
# Export all open issues to CSV
.\Sync-GitHubTasks.ps1 `
    -Owner 'Handy-Heathers-HJ-SERVICES-LLC' `
    -Repo  'PowerShell' `
    -Token $env:GITHUB_TOKEN `
    -ExportPath './tasks.csv'
```

```powershell
# Export issues for milestone #2 only
.\Sync-GitHubTasks.ps1 `
    -Owner           'Handy-Heathers-HJ-SERVICES-LLC' `
    -Repo            'PowerShell' `
    -Token           $env:GITHUB_TOKEN `
    -MilestoneNumber 2 `
    -ExportPath      './milestone2-tasks.csv'
```

```powershell
# Export open high-priority issues
.\Sync-GitHubTasks.ps1 `
    -Owner       'Handy-Heathers-HJ-SERVICES-LLC' `
    -Repo        'PowerShell' `
    -Token       $env:GITHUB_TOKEN `
    -LabelFilter 'high', 'critical' `
    -ExportPath  './high-priority-tasks.csv'
```

```powershell
# Import an updated CSV and sync changes back to GitHub
.\Sync-GitHubTasks.ps1 `
    -Owner      'Handy-Heathers-HJ-SERVICES-LLC' `
    -Repo       'PowerShell' `
    -Token      $env:GITHUB_TOKEN `
    -ImportPath './tasks-updated.csv'
```

#### Exported CSV Columns

| Column | Description |
|--------|-------------|
| `IssueNumber` | GitHub issue number |
| `TaskName` | Issue title |
| `State` | `open` or `closed` |
| `Status` | Derived from labels: `Open`, `In Progress`, `Blocked`, `In Review`, `Closed` |
| `Priority` | Derived from labels: `Critical`, `High`, `Medium`, `Low`, `Normal` |
| `Labels` | Semi-colon separated list of all labels |
| `Assignees` | Semi-colon separated list of assignee logins |
| `Milestone` | Milestone title (if any) |
| `MilestoneDue` | Milestone due date `yyyy-MM-dd` (if set) |
| `CreatedDate` | Issue creation date `yyyy-MM-dd` |
| `UpdatedDate` | Last update date `yyyy-MM-dd` |
| `ClosedDate` | Closed date `yyyy-MM-dd` (blank if open) |
| `URL` | Link to the issue on GitHub |
| `Body` | Issue description (newlines replaced by spaces) |

#### Import CSV Columns

When updating issues from an external tool, the import CSV must include `IssueNumber` and at least one of the following update columns:

| Column | Description |
|--------|-------------|
| `IssueNumber` | **Required** — GitHub issue number to update |
| `ExternalStatus` | `open`, `closed`, `done`, `in progress`, `blocked`, `review` |
| `ExternalPriority` | `critical`, `high`, `medium`, `low`, `normal` |
| `DueDate` | Date string (e.g. `2025-06-30`) — creates/assigns a milestone |

---

## Automating with GitHub Actions

You can schedule these scripts to run automatically using a GitHub Actions workflow.

### Weekly Report Workflow

Create `.github/workflows/weekly-status-report.yml`:

```yaml
name: Weekly Status Report
on:
  schedule:
    - cron: '0 9 * * 1'   # Every Monday at 09:00 UTC
  workflow_dispatch:       # Allow manual runs

jobs:
  report:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4

      - name: Generate Weekly Report
        shell: pwsh
        env:
          GITHUB_TOKEN: ${{ secrets.GITHUB_TOKEN }}
        run: |
          ./tools/project-management/Get-StatusUpdate.ps1 `
            -Owner '${{ github.repository_owner }}' `
            -Repo  'PowerShell' `
            -Timeframe 'Weekly' `
            -UploadToRepo `
            -UploadBranch 'main'
```

### Daily Report Workflow

Create `.github/workflows/daily-status-report.yml`:

```yaml
name: Daily Status Report
on:
  schedule:
    - cron: '0 8 * * *'   # Every day at 08:00 UTC
  workflow_dispatch:

jobs:
  report:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4

      - name: Generate Daily Report
        shell: pwsh
        env:
          GITHUB_TOKEN: ${{ secrets.GITHUB_TOKEN }}
        run: |
          ./tools/project-management/Get-StatusUpdate.ps1 `
            -Owner '${{ github.repository_owner }}' `
            -Repo  'PowerShell' `
            -Timeframe 'Daily' `
            -OutputPath './daily-report.md'

      - name: Upload Report Artifact
        uses: actions/upload-artifact@v4
        with:
          name: daily-status-report
          path: ./daily-report.md
```

---

## Priority and Status Label Conventions

The scripts use the following label naming conventions to determine task priority and status.

### Priority Labels

| Priority | Recognised Labels |
|----------|-------------------|
| Critical | `critical`, `priority: critical` |
| High | `high`, `priority: high` |
| Medium | `medium`, `priority: medium` |
| Low | `low`, `priority: low` |
| Normal | *(default when no priority label is present)* |

### Status Labels

| Status | Recognised Labels |
|--------|-------------------|
| In Progress | `in progress`, `status: in progress` |
| Blocked | `blocked`, `status: blocked` |
| In Review | `review`, `status: review` |
| Closed | *(issue state = closed)* |
| Open | *(default)* |

---

## Security Notes

- **Never hard-code** your GitHub PAT in scripts or workflow files. Always use environment variables or GitHub Actions secrets.
- The PAT only needs `repo` scope. Use a fine-grained token with minimal permissions where possible.
- The `Send-MailMessage` cmdlet used for email delivery is available on Windows PowerShell. On PowerShell 7 it may require the [PSMTP module](https://www.powershellgallery.com/packages/Mailozaurr) or an alternative email solution.
