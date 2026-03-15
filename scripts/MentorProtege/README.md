# Mentor-Protégé Program — Contact Manager

A PowerShell script that manages contacts for the **Mentor-Protégé Program (MPP)** using an
Excel spreadsheet (`.xlsx`). Contacts are stored locally and can be synced to this GitHub
repository for centralized, version-controlled tracking.

---

## Prerequisites

| Requirement | Version | Notes |
|---|---|---|
| PowerShell | 7.0 + | [Download](https://aka.ms/powershell) |
| ImportExcel module | latest | `Install-Module -Name ImportExcel -Scope CurrentUser` |
| Git | any | Required only for the `Sync` action |

Install the required module once:

```powershell
Install-Module -Name ImportExcel -Scope CurrentUser -Force
```

---

## Quick Start

```powershell
# 1. Clone / navigate to the repository
cd path\to\PowerShell

# 2. Create a new contacts file
.\scripts\MentorProtege\Manage-MPPContacts.ps1 -Action Initialize

# 3. Add your first contact
.\scripts\MentorProtege\Manage-MPPContacts.ps1 -Action Add `
    -Name "Jane Smith" -Role Mentor `
    -Organization "Acme Corp" -Email "jane.smith@acme.com" `
    -Phone "555-1234" -Industry "Technology" -Location "Seattle, WA"

# 4. List all contacts
.\scripts\MentorProtege\Manage-MPPContacts.ps1 -Action List
```

---

## Excel Spreadsheet Structure

The script uses a single worksheet named **Contacts** with the following columns:

| Column | Description | Required |
|---|---|---|
| **Name** | Contact's full name | ✅ |
| **Role** | `Mentor` or `Protege` | ✅ |
| **Organization** | Company, university, or agency | |
| **Email** | Unique identifier — used for updates and deletes | ✅ |
| **Phone** | Phone number | |
| **Industry** | Industry or sector (e.g., Technology, Engineering) | |
| **Location** | City, state, or region | |
| **Notes** | Free-text notes | |
| **DateAdded** | Auto-set when contact is first added (`yyyy-MM-dd`) | |
| **LastUpdated** | Auto-updated on every modification (`yyyy-MM-dd`) | |

> **Tip:** Open `MPP-Contacts.xlsx` directly in Excel to view or edit contacts manually,
> then use `-Action Import` to bring your changes back into the managed file.

---

## Actions Reference

### `Initialize` — Create a new contacts file

Creates `MPP-Contacts.xlsx` (or the file specified by `-ContactsFile`) with the standard
column headers and two sample rows that you should delete before use.

```powershell
.\Manage-MPPContacts.ps1 -Action Initialize

# Use a custom file path
.\Manage-MPPContacts.ps1 -Action Initialize -ContactsFile "C:\MPP\contacts.xlsx"
```

---

### `Add` — Add a new contact

Appends a new contact row. **Name**, **Role**, and **Email** are required.
Email must be unique — the script will reject duplicates.

```powershell
# Add a mentor
.\Manage-MPPContacts.ps1 -Action Add `
    -Name "Jane Smith" -Role Mentor `
    -Organization "Acme Corp" -Email "jane.smith@acme.com" `
    -Phone "555-1234" -Industry "Technology" -Location "Seattle, WA" `
    -Notes "Cloud computing expert, available Tuesdays"

# Add a protege
.\Manage-MPPContacts.ps1 -Action Add `
    -Name "John Doe" -Role Protege `
    -Organization "State University" -Email "john.doe@stateuniversity.edu" `
    -Industry "Engineering" -Location "Austin, TX"
```

---

### `Update` — Update an existing contact

Locates the contact by **Email** and updates only the fields you supply.
Fields you omit are left unchanged.

```powershell
# Update phone number
.\Manage-MPPContacts.ps1 -Action Update -Email "jane.smith@acme.com" -Phone "555-9999"

# Update multiple fields
.\Manage-MPPContacts.ps1 -Action Update -Email "jane.smith@acme.com" `
    -Organization "New Corp" -Location "Portland, OR" -Notes "Relocated to Portland"
```

---

### `Remove` — Remove a contact

Deletes the contact row matched by **Email**.

```powershell
.\Manage-MPPContacts.ps1 -Action Remove -Email "john.doe@stateuniversity.edu"
```

---

### `Search` — Search and filter contacts

Searches across Name, Role, Organization, Email, Industry, Location, and Notes.
Combine `-SearchTerm` with role/industry/location filters for precise results.

```powershell
# Keyword search
.\Manage-MPPContacts.ps1 -Action Search -SearchTerm "Acme"

# Filter by role
.\Manage-MPPContacts.ps1 -Action Search -FilterRole Mentor

# Filter by industry
.\Manage-MPPContacts.ps1 -Action Search -FilterIndustry "Technology"

# Filter by location
.\Manage-MPPContacts.ps1 -Action Search -FilterLocation "Seattle"

# Combined: mentors in Seattle
.\Manage-MPPContacts.ps1 -Action Search -FilterRole Mentor -FilterLocation "Seattle"
```

---

### `List` — List all contacts

Displays every contact in a formatted table.  
Supports the same `-FilterRole`, `-FilterIndustry`, and `-FilterLocation` options as `Search`.

```powershell
# List all contacts
.\Manage-MPPContacts.ps1 -Action List

# List all proteges in Engineering
.\Manage-MPPContacts.ps1 -Action List -FilterRole Protege -FilterIndustry "Engineering"
```

---

### `Export` — Export to a dated Excel file

Creates a formatted, timestamped Excel export (suitable for sharing or archiving).
Does **not** modify `MPP-Contacts.xlsx`.

```powershell
# Export with auto-generated filename
.\Manage-MPPContacts.ps1 -Action Export

# Export to a specific path
.\Manage-MPPContacts.ps1 -Action Export -ExportFile "C:\Shared\Q1-2025-Contacts.xlsx"
```

---

### `Import` — Import from an external Excel file

Merges contacts from another `.xlsx` file (same column structure) into the main contacts file.

- Rows whose **Email** already exists are **updated**.
- Rows with a new **Email** are **appended**.

```powershell
.\Manage-MPPContacts.ps1 -Action Import -ImportFile "C:\Downloads\NewContacts.xlsx"
```

---

### `Sync` — Commit and push to Git

Stages the contacts file and pushes it to the configured remote, creating a version-controlled
audit trail of every change.

```powershell
# Sync with an auto-generated commit message
.\Manage-MPPContacts.ps1 -Action Sync

# Sync with a custom commit message
.\Manage-MPPContacts.ps1 -Action Sync -CommitMessage "Add Q1 2025 mentor contacts"
```

> **Note:** Your Git remote must be configured and authenticated before running `Sync`.

---

## Using a Custom Contacts File Path

Every action accepts `-ContactsFile` to override the default path
(`MPP-Contacts.xlsx` next to the script):

```powershell
.\Manage-MPPContacts.ps1 -Action List -ContactsFile "D:\MentorProgram\contacts.xlsx"
```

---

## Typical Workflow

```text
1. Initialize   →  Create MPP-Contacts.xlsx
2. Add          →  Add mentor and protégé contacts
3. List/Search  →  Browse and verify entries
4. Update       →  Correct or enrich contact details
5. Export       →  Share a snapshot with your team
6. Sync         →  Push the updated file to GitHub
```

---

## Editing the Excel File Manually

You can open `MPP-Contacts.xlsx` directly in Microsoft Excel or LibreOffice Calc, make
changes, save the file, and the next time the script runs it will read your changes.

If you add contacts directly in Excel (rather than via `-Action Add`), the **DateAdded** and
**LastUpdated** columns will be blank for those rows — fill them in manually or use
`-Action Update` to trigger an automatic timestamp.

---

## Troubleshooting

| Problem | Solution |
|---|---|
| `Missing required module: ImportExcel` | Run `Install-Module -Name ImportExcel -Scope CurrentUser` |
| `Contacts file not found` | Run `-Action Initialize` to create the file first |
| `A contact with email '...' already exists` | Use `-Action Update` instead of `-Action Add` |
| `Git is not installed or not in PATH` | Install Git and ensure it is on your `PATH` |
| `The directory is not inside a Git repository` | Move the contacts file inside a Git-tracked folder |

---

## File Layout

```
scripts/
└── MentorProtege/
    ├── Manage-MPPContacts.ps1   ← this script
    ├── MPP-Contacts.xlsx        ← generated on Initialize (gitignored by default for privacy)
    └── README.md                ← this file
```

> **Privacy notice:** `MPP-Contacts.xlsx` may contain personal information (names, email
> addresses, phone numbers). Review your organization's data-handling policies before
> committing this file to a public repository. Consider using a private repository or
> encrypting the file.

---

## License

This script is provided under the same license as the parent repository.
