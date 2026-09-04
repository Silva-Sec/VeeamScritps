


<#
.SYNOPSIS
    Interactive removal of SharePoint site data from a Veeam Backup for
    Microsoft 365 repository.
 
.DESCRIPTION
    Prompts the operator for every input instead of requiring manual edits.
    Flow:
      1. Connect to the VB365 server
      2. Pick a repository (standard or archive / long-term)
      3. Search sites in that repository
      4. Review the target and run a mandatory dry-run (-WhatIf)
      5. Type an explicit confirmation word before anything is deleted
 
    Writes a transcript to disk for audit purposes.
 
.NOTES
    Removal is IRREVERSIBLE. It will fail against object storage
    repositories that have data immutability enabled.
#>
 
[CmdletBinding()]
param(
    [string]$LogPath = "$env:TEMP\Remove-VBOSiteData_$(Get-Date -Format 'yyyyMMdd_HHmmss').log",
    [int]$PageSize = 25
)
 
$ErrorActionPreference = 'Stop'
 
#region Helpers -----------------------------------------------------------
 
function Write-Header {
    param([string]$Text)
    Write-Host ''
    Write-Host ('=' * 72) -ForegroundColor DarkCyan
    Write-Host "  $Text" -ForegroundColor Cyan
    Write-Host ('=' * 72) -ForegroundColor DarkCyan
}
 
function Get-FirstProperty {
    # Returns the value of the first property that exists and is populated.
    param($Object, [string[]]$Names)
    foreach ($n in $Names) {
        if ($Object.PSObject.Properties.Name -contains $n -and $Object.$n) {
            return $Object.$n
        }
    }
    return '(n/a)'
}
 
function Get-SiteSearchText {
    # Everything a site can reasonably be searched by, in one string.
    param($Site)
    $parts = foreach ($n in 'Title', 'Name', 'DisplayName', 'URL', 'Url', 'SiteUrl', 'Email') {
        if ($Site.PSObject.Properties.Name -contains $n -and $Site.$n) { [string]$Site.$n }
    }
    ($parts | Select-Object -Unique) -join ' '
}
 
function Select-FromList {
<#
    Numbered menu that also accepts free text.
 
      - a number      -> selects that entry
      - any text      -> filters the list (auto-selects on a single match)
      - "all"         -> restores the full list
      - "0"           -> cancels
#>
    param(
        [Parameter(Mandatory)] [object[]]$Items,
        [Parameter(Mandatory)] [scriptblock]$Label,
        [scriptblock]$SearchText,
        [string]$ItemName = 'item',
        [int]$PageSize = 25
    )
 
    if (-not $SearchText) { $SearchText = $Label }
    $working = @($Items)
 
    while ($true) {
 
        if ($working.Count -eq 0) {
            Write-Host "  No $ItemName matches the filter. Showing the full list again." -ForegroundColor Yellow
            $working = @($Items)
        }
 
        $shown = [Math]::Min($working.Count, $PageSize)
        Write-Host ''
        for ($i = 0; $i -lt $shown; $i++) {
            Write-Host ('  [{0,3}] {1}' -f ($i + 1), (& $Label $working[$i]))
        }
        if ($working.Count -gt $shown) {
            Write-Host ("  ... and {0} more. Type part of the name to narrow the list." -f ($working.Count - $shown)) -ForegroundColor DarkGray
        }
        Write-Host '  [  0] Cancel' -ForegroundColor DarkGray
 
        Write-Host ''
        Write-Host "  Enter the NUMBER shown in brackets to select a $ItemName." -ForegroundColor Yellow
        Write-Host '  Or type part of the name to filter, "all" to reset, 0 to cancel.' -ForegroundColor DarkGray
        $answer = Read-Host "  Selection"
        $answer = $answer.Trim()
 
        if ([string]::IsNullOrWhiteSpace($answer)) {
            Write-Host '  Nothing entered. Please enter a number from the list.' -ForegroundColor Yellow
            continue
        }
 
        if ($answer -eq '0') { return $null }
 
        if ($answer -match '^\d+$') {
            $index = [int]$answer
            if ($index -ge 1 -and $index -le $shown) {
                return $working[$index - 1]
            }
            Write-Host ("  Number out of range. Valid numbers are 1 to {0}." -f $shown) -ForegroundColor Yellow
            continue
        }
 
        if ($answer -eq 'all') {
            $working = @($Items)
            Write-Host ("  Filter cleared. Showing all {0} entries." -f $Items.Count) -ForegroundColor DarkGray
            continue
        }
 
        # Treat anything else as a filter over the ORIGINAL list.
        $matches = @($Items | Where-Object { (& $SearchText $_) -like "*$answer*" })
 
        if ($matches.Count -eq 1) {
            Write-Host ("  One match: {0}" -f (& $Label $matches[0])) -ForegroundColor Green
            $ok = Read-Host '  Use this one? (Y/n)'
            if ($ok -notmatch '^[nN]') { return $matches[0] }
            $working = @($Items)
            continue
        }
 
        if ($matches.Count -eq 0) {
            Write-Host "  No $ItemName matches '$answer'." -ForegroundColor Yellow
            continue
        }
 
        Write-Host ("  {0} matches for '{1}'. Pick a number below." -f $matches.Count, $answer) -ForegroundColor DarkGray
        $working = $matches
    }
}
 
#endregion ----------------------------------------------------------------
 
Start-Transcript -Path $LogPath | Out-Null
 
try {
    Write-Header 'Remove site data - Veeam Backup for Microsoft 365'
    Write-Host "Session log: $LogPath" -ForegroundColor DarkGray
 
    # --- 1. Module --------------------------------------------------------
    if (-not (Get-Module -Name Veeam.Archiver.PowerShell)) {
        Write-Host "`nLoading Veeam.Archiver.PowerShell module..." -ForegroundColor DarkGray
        Import-Module Veeam.Archiver.PowerShell
    }
 
    # --- 2. Connection ----------------------------------------------------
    Write-Header 'Server connection'
    $server = Read-Host 'VB365 server (press Enter for localhost)'
    if ([string]::IsNullOrWhiteSpace($server)) { $server = 'localhost' }
 
    $useCred = Read-Host 'Use alternate credentials? (y/N)'
    if ($useCred -match '^[yY]') {
        $cred = Get-Credential -Message "Credentials for $server"
        Connect-VBOServer -Server $server -Credential $cred | Out-Null
    }
    else {
        Connect-VBOServer -Server $server | Out-Null
    }
    Write-Host "Connected to $server." -ForegroundColor Green
 
    # --- 3. Repository ----------------------------------------------------
    Write-Header 'Repository selection'
    Write-Host '  [1] Backup / object storage repository'
    Write-Host '  [2] Archive repository (long-term)'
    $repoType = Read-Host 'Repository type (Enter for 1)'
 
    $repos = if ($repoType -eq '2') { @(Get-VBORepository -LongTerm) } else { @(Get-VBORepository) }
 
    if ($repos.Count -eq 0) { throw 'No repositories found for the selected type.' }
 
    $repository = Select-FromList -Items $repos -ItemName 'repository' -PageSize $PageSize -Label {
        param($r) '{0}   (type: {1})' -f $r.Name, (Get-FirstProperty $r @('Type', 'ObjectStorageRepository'))
    } -SearchText {
        param($r) $r.Name
    }
    if (-not $repository) { Write-Host 'Cancelled by operator.' -ForegroundColor Yellow; return }
 
    Write-Host "`nSelected repository: $($repository.Name)" -ForegroundColor Green
 
    # Early immutability warning
    $immutableProp = $repository.PSObject.Properties |
                     Where-Object { $_.Name -match 'Immutab' -and $_.Value -eq $true }
    if ($immutableProp) {
        Write-Host 'WARNING: this repository appears to have immutability enabled.' -ForegroundColor Red
        Write-Host 'VB365 will most likely reject the removal.' -ForegroundColor Red
    }
 
    # --- 4. Site lookup ---------------------------------------------------
    Write-Header 'Site lookup'
    # NOTE: Get-VBOEntityData -Name does an exact match and accepts no
    # wildcards, and for personal sites the server-side name is the URL
    # rather than the display title. Partial names therefore return nothing.
    # We pull the full list once and filter client-side instead.
    Write-Host 'Retrieving all sites from the repository (single query)...' -ForegroundColor DarkGray
    Write-Host 'This may take a while on large repositories.' -ForegroundColor DarkGray
 
    $allSites = @(Get-VBOEntityData -Type Site -Repository $repository)
 
    if ($allSites.Count -eq 0) {
        Write-Host 'This repository contains no backed-up site data.' -ForegroundColor Yellow
        return
    }
 
    Write-Host ("{0} site(s) in this repository." -f $allSites.Count) -ForegroundColor Green
 
    $filter = Read-Host 'Filter by name or URL (press Enter to list all)'
    if ([string]::IsNullOrWhiteSpace($filter)) {
        $sites = $allSites
    }
    else {
        $sites = @($allSites | Where-Object { (Get-SiteSearchText $_) -like "*$filter*" })
        if ($sites.Count -eq 0) {
            Write-Host "No site matches '$filter'. Showing the full list instead." -ForegroundColor Yellow
            $sites = $allSites
        }
        else {
            Write-Host ("{0} site(s) match '{1}'." -f $sites.Count, $filter) -ForegroundColor Green
        }
    }
 
    $site = Select-FromList -Items $sites -ItemName 'site' -PageSize $PageSize -Label {
        param($s) '{0}   [{1}]' -f (Get-FirstProperty $s @('Title', 'Name', 'DisplayName')),
                                    (Get-FirstProperty $s @('URL', 'Url', 'SiteUrl'))
    } -SearchText {
        param($s) Get-SiteSearchText $s
    }
    if (-not $site) { Write-Host 'Cancelled by operator.' -ForegroundColor Yellow; return }
 
    # --- 5. Summary and dry-run -------------------------------------------
    Write-Header 'Operation summary'
    [PSCustomObject]@{
        Repository = $repository.Name
        Site       = Get-FirstProperty $site @('Title', 'Name', 'DisplayName')
        URL        = Get-FirstProperty $site @('URL', 'Url', 'SiteUrl')
        Scope      = 'All restore points for this site in this repository'
        Reversible = 'NO'
    } | Format-List
 
    Write-Host 'Running dry-run (-WhatIf):' -ForegroundColor Yellow
    Remove-VBOEntityData -Repository $repository -Site $site -WhatIf
 
    # --- 6. Explicit confirmation -----------------------------------------
    Write-Header 'Confirmation'
    Write-Host 'This removal is permanent and cannot be undone.' -ForegroundColor Red
    $confirm = Read-Host 'Type DELETE (uppercase) to proceed'
 
    if ($confirm -cne 'DELETE') {
        Write-Host "`nAborted. Nothing was removed." -ForegroundColor Yellow
        return
    }
 
    # --- 7. Execution -------------------------------------------------------
    Write-Host "`nRemoving data..." -ForegroundColor Yellow
    $start = Get-Date
 
    Remove-VBOEntityData -Repository $repository -Site $site -Confirm:$false
 
    Write-Host ("Completed in {0:hh\:mm\:ss}." -f ((Get-Date) - $start)) -ForegroundColor Green
 
    $session = Get-VBODataManagementSession | Sort-Object CreationTime -Descending | Select-Object -First 1
    if ($session) {
        Write-Host "`nLatest data management session:" -ForegroundColor DarkGray
        $session | Format-List
    }
}
catch {
    Write-Host "`nERROR: $($_.Exception.Message)" -ForegroundColor Red
    if ($_.Exception.Message -match 'immutab') {
        Write-Host 'Likely cause: immutability enabled on the object storage repository.' -ForegroundColor Red
        Write-Host 'Removal will only be possible after the immutability period expires.' -ForegroundColor Red
    }
}
finally {
    if (Get-Command Disconnect-VBOServer -ErrorAction SilentlyContinue) {
        Disconnect-VBOServer -ErrorAction SilentlyContinue
    }
    Stop-Transcript | Out-Null
    Write-Host "`nLog saved to: $LogPath" -ForegroundColor DarkGray
}
 
