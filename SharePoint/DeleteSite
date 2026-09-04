# Remove-VBOSiteData-Interactive

Interactive PowerShell script to remove a SharePoint or OneDrive site's backed-up
data from a **Veeam Backup for Microsoft 365** repository.

Wraps `Remove-VBOEntityData` in a guided workflow: no variables to edit, no
copy-pasting of site URLs, a mandatory dry-run, and an explicit confirmation
step before anything is deleted.

---

## Why

The documented one-liner works, but is unforgiving in day-to-day operations:

```powershell
$repository = Get-VBORepository -Name "ABC Daily Backup"
$site = Get-VBOEntityData -Type Site -Repository $repository -Name "Support"
Remove-VBOEntityData -Repository $repository -Site $site
```

Three problems show up quickly:

1. **`-Name` does not do partial matching.** It accepts no wildcards and compares
   exactly. Worse, for personal sites the server-side name is the URL, not the
   display title, so searching for `Adele` never finds
   `Adele Vance / adelev_contoso_onmicrosoft_com`.
2. **`Get-VBOEntityData` can return more than one object** for the same display
   name (renamed or re-created sites), which throws a parameter-binding error
   when piped straight into `Remove-VBOEntityData`.
3. **Removal is irreversible** and there is nothing between a typo and data loss.

This script pulls the site list once, filters it client-side against title *and*
URL, forces you to look at a `-WhatIf` preview, and requires you to type
`DELETE` before it runs.

---

## Requirements

| Requirement | Detail |
|---|---|
| Veeam Backup for Microsoft 365 | v8.x (tested against build 8.5) |
| PowerShell module | `Veeam.Archiver.PowerShell` |
| PowerShell | 5.1 (Windows) |
| Permissions | An account allowed to connect to the VB365 server and manage backup data |

The module ships with VB365 and normally lives at:

```
C:\Program Files\Veeam\Backup365\Veeam.Archiver.PowerShell\Veeam.Archiver.PowerShell.psd1
```

---

## Usage

```powershell
.\Remove-VBOSiteData-Interactive.ps1
```

Optional parameters:

```powershell
# Custom transcript location
.\Remove-VBOSiteData-Interactive.ps1 -LogPath "D:\Logs\vbo-site-removal.log"

# Show more entries per page when browsing large tenants
.\Remove-VBOSiteData-Interactive.ps1 -PageSize 50
```

### Walkthrough

```
========================================================================
  Remove site data - Veeam Backup for Microsoft 365
========================================================================

VB365 server (press Enter for localhost):
Use alternate credentials? (y/N): n
Connected to localhost.

========================================================================
  Repository selection
========================================================================
  [1] Backup / object storage repository
  [2] Archive repository (long-term)
Repository type (Enter for 1): 1

  [  1] ABC Daily Backup   (type: Local)
  [  2] Azure Blob Prod    (type: ObjectStorage)
  [  0] Cancel

  Enter the NUMBER shown in brackets to select a repository.
  Selection: 2

Selected repository: Azure Blob Prod

========================================================================
  Site lookup
========================================================================
Retrieving all sites from the repository (single query)...
412 site(s) in this repository.
Filter by name or URL (press Enter to list all): Adele
1 site(s) match 'Adele'.

  [  1] Adele Vance   [https://contoso-my.sharepoint.com/personal/adelev_...]
  [  0] Cancel
  Selection: 1

========================================================================
  Operation summary
========================================================================
Repository : Azure Blob Prod
Site       : Adele Vance
URL        : https://contoso-my.sharepoint.com/personal/adelev_...
Scope      : All restore points for this site in this repository
Reversible : NO

Running dry-run (-WhatIf):
What if: Performing the operation "Remove" on target "Adele Vance".

========================================================================
  Confirmation
========================================================================
This removal is permanent and cannot be undone.
Type DELETE (uppercase) to proceed:
```

### Selection menu

Every menu accepts three kinds of input:

| Input | Behaviour |
|---|---|
| A number | Selects that entry |
| Free text | Filters the list; auto-selects and asks for confirmation on a single match |
| `all` | Clears the filter and shows the full list |
| `0` | Cancels |

Site filtering matches against title, display name, URL and email, so
`Adele`, `adelev`, `contoso-my` and `onmicrosoft` all resolve to the same entry.

---

## Safety features

- **Mandatory dry-run** — `-WhatIf` output is printed before the confirmation prompt
- **Case-sensitive confirmation** — the literal string `DELETE` is required; Enter or `delete` aborts
- **Immutability pre-check** — inspects the repository object and warns before you invest time in the flow
- **Transcript** — full session written to `%TEMP%` by default, useful when attaching evidence to a support case
- **Clean disconnect** — `Disconnect-VBOServer` runs in a `finally` block

---

## Limitations

- **Immutability blocks removal.** Entity data cannot be removed from object
  storage repositories with data immutability enabled. The operation will fail
  until the immutability period expires.
- **Sites are all-or-nothing.** Unlike the `-User` and `-Group` parameter sets of
  `Remove-VBOEntityData`, the `-Site` set has no sub-switches. It removes that
  site's data across every restore point in the repository.
- **One site per run.** For batch removals, replace `Select-FromList` with a
  comma-separated selection and wrap the removal in a `foreach`, keeping a single
  confirmation at the end.
- **Initial query can be slow.** The full site list is fetched once so filtering
  can be done locally; on very large tenants the first query takes a while.

---

## This is not the same as unprotecting a site

If the goal is to **stop backing up** a site rather than purge existing data, do
not use this script. Use:

```powershell
Remove-VBOBackupItem -Job $job -BackupItem $item
```

That removes the object from the backup job and leaves existing restore points
in place until retention ages them out.

---

## References

- [Remove-VBOEntityData](https://helpcenter.veeam.com/docs/vbo365/powershell/remove-vboentitydata.html)
- [Get-VBOEntityData](https://helpcenter.veeam.com/docs/vbo365/powershell/get-vboentitydata.html)
- [Get-VBORepository](https://helpcenter.veeam.com/docs/vbo365/powershell/get-vborepository.html)
- [Remove-VBOBackupItem](https://helpcenter.veeam.com/docs/vbo365/powershell/remove-vbobackupitem.html)

---

## Disclaimer

This script permanently deletes backup data. It is provided as-is, is not a Veeam
product, and is not supported by Veeam. Test it in a lab against a non-production
repository before using it on customer data.

## License

MIT
