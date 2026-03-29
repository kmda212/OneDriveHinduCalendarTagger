<#
.SYNOPSIS
    Hindu Festival Photo Album Creator for OneDrive Personal.

.DESCRIPTION
    Scans your OneDrive photos and matches them to Hindu festival dates using
    the Calendarific API. Creates named albums (e.g. DiwaliLifetime) and adds
    matched photos by reference — photos stay in their original OneDrive location.

    Progress is saved to OneDrive after every checkpoint so the script can
    safely resume after any interruption.

.PARAMETER Rescan
    Force a full re-scan of all photos from the beginning.
    By default the script resumes from its last saved checkpoint.

.EXAMPLE
    .\FestivalAlbums.ps1
    .\FestivalAlbums.ps1 -Rescan
#>

[CmdletBinding()]
param(
    [switch]$Rescan,
    [switch]$DryRun,

    # Optional path to a local festivals_config.json produced by Get-HinduFestivals.ps1.
    # When provided, the file is uploaded to OneDrive and used for this run.
    # On subsequent runs without this param, the copy stored in OneDrive is used.
    [string]$FestivalsConfig = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── Ensure Microsoft Graph SDK modules are available ──────────────────────────
foreach ($mod in @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Files')) {
    if (-not (Get-Module -ListAvailable -Name $mod)) {
        Write-Host "[Setup] Installing $mod..." -ForegroundColor Yellow
        Install-Module $mod -Scope CurrentUser -Force -AllowClobber
    }
}
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

# ═══════════════════════════════════════════════════════════════════════════════
#  CONFIGURATION  — Credentials are collected via popup on first run and saved
#                   locally. All other settings can be changed here freely.
# ═══════════════════════════════════════════════════════════════════════════════

$Config = @{

    # Authentication is handled by Connect-MgGraph (Microsoft Graph PowerShell SDK).
    # No ClientId registration required — the SDK's built-in app is used.
    # The Calendarific API key is stored securely in OneDrive after first entry.

    # How many past years to scan for festival photos.
    YearsToScan        = 30

    # Re-download calendar if the cached copy is older than this (days).
    CacheRefreshDays   = 90

    # Write progress to OneDrive after every N photo additions.
    # Lower = safer resume, but slightly more API calls.
    CheckpointEvery    = 20

    # OneDrive folder where state, cache, and settings files are stored (auto-created).
    StateFolder        = 'Apps/FestivalTimeline'

    # ── Festivals to create albums for ───────────────────────────────────────
    # This list is used only if festivals_config.json does not exist.
    # Run Get-HinduFestivals.ps1 to generate festivals_config.json with
    # your own selection — that file takes priority over this list.
    FestivalsToTrack   = @(
        'Diwali'
        'Holi'
        'Navratri'
        'Raksha Bandhan'
        'Dussehra'
        'Janmashtami'
        'Ganesh Chaturthi'
        'Makar Sankranti'
        'Maha Shivaratri'
    )
}

# OneDrive path for the festival selection (loaded after auth)
$OneDriveFestivalsConfigPath = "$($Config.StateFolder)/festivals_config.json"

# OneDrive path — stores the Calendarific API key (never in git, never on disk)
$OneDriveSettingsPath = "$($Config.StateFolder)/settings.json"

# Album name = festival name + this suffix  (e.g. "Diwali Over the years")
$AlbumSuffix = 'Over the years'

# Image file extensions to scan
$ImageExtensions = @('.jpg', '.jpeg', '.heic', '.png', '.gif', '.bmp', '.tiff')

# Microsoft Graph base URL
$GraphBase = 'https://graph.microsoft.com/v1.0'

# ═══════════════════════════════════════════════════════════════════════════════
#  AUTHENTICATION  — Microsoft Graph PowerShell SDK
#  Connect-MgGraph opens a browser sign-in page automatically.
#  Token refresh is handled by the SDK — no manual token management needed.
# ═══════════════════════════════════════════════════════════════════════════════

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Connect-ToOneDrive {
    Write-Host '[Auth] Connecting to OneDrive via Microsoft Graph...' -ForegroundColor Cyan
    try {
        Connect-MgGraph -Scopes 'Files.ReadWrite' -NoWelcome -ErrorAction Stop
        Write-Host '[Auth] Signed in to OneDrive successfully.' -ForegroundColor Green
    } catch {
        Write-Error "[Auth] Sign-in failed: $_"
        exit 1
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
#  GRAPH API HELPER
# ═══════════════════════════════════════════════════════════════════════════════

function Invoke-Graph {
    param(
        [string]   $Method = 'GET',
        [string]   $Uri,
        [hashtable]$Body
    )

    $params = @{ Method = $Method; Uri = $Uri }

    if ($Body) {
        $params.ContentType = 'application/json'
        $params.Body        = ($Body | ConvertTo-Json -Depth 10)
    }

    # -OutputType PSObject converts the top-level object but nested collections
    # (e.g. the 'value' array items) can still come back as Hashtables in some
    # SDK versions. A JSON round-trip forces deep conversion to PSCustomObject
    # so dot-notation works on every nested property ($item.photo, $item.file, etc.)
    $raw = Invoke-MgGraphRequest @params -OutputType Json
    return $raw | ConvertFrom-Json
}

function Load-OneDriveSettings {
    Write-Host '[Settings] Loading settings from OneDrive...' -ForegroundColor Cyan
    $raw = Read-OneDriveJson -RelativePath $OneDriveSettingsPath
    if ($raw -and $raw.CalendarificApiKey) {
        Write-Host '[Settings] Calendarific API key loaded from OneDrive.' -ForegroundColor Green
        return $raw.CalendarificApiKey
    }
    return $null
}

function Save-OneDriveSettings {
    param([string]$ApiKey)
    Write-OneDriveJson -RelativePath $OneDriveSettingsPath -Data @{
        CalendarificApiKey = $ApiKey
        saved_on           = (Get-Date -Format 'o')
    }
    Write-Host '[Settings] API key saved to OneDrive.' -ForegroundColor Green
}

function Load-FestivalsConfig {
    <#
      Priority order:
        1. -FestivalsConfig <path> was passed → validate, upload to OneDrive, use it
        2. No param → try to load from OneDrive
        3. Not in OneDrive either → fall back to built-in list
    #>

    if (-not [string]::IsNullOrWhiteSpace($FestivalsConfig)) {
        # ── Local file explicitly provided ──
        if (-not (Test-Path $FestivalsConfig)) {
            Write-Warning "[Config] -FestivalsConfig file not found: $FestivalsConfig — using built-in list."
            return
        }
        try {
            $fc = Get-Content $FestivalsConfig -Raw | ConvertFrom-Json
            if ($fc.FestivalsToTrack -and $fc.FestivalsToTrack.Count -gt 0) {
                $Config.FestivalsToTrack = @($fc.FestivalsToTrack)
                Write-Host "[Config] Loaded $($Config.FestivalsToTrack.Count) festivals from local file: $FestivalsConfig" -ForegroundColor Green

                # Upload to OneDrive so future runs without -FestivalsConfig use it
                Write-Host '[Config] Uploading festivals_config.json to OneDrive...' -ForegroundColor Cyan
                Write-OneDriveJson -RelativePath $OneDriveFestivalsConfigPath -Data $fc
                Write-Host '[Config] festivals_config.json saved to OneDrive.' -ForegroundColor Green
            } else {
                Write-Warning '[Config] Provided file has no festivals — using built-in list.'
            }
        } catch {
            Write-Warning "[Config] Could not read $FestivalsConfig`: $_ — using built-in list."
        }
        return
    }

    # ── No local file provided — try OneDrive ──
    Write-Host '[Config] Loading festivals_config.json from OneDrive...' -ForegroundColor Cyan
    $raw = Read-OneDriveJson -RelativePath $OneDriveFestivalsConfigPath
    if ($raw -and $raw.FestivalsToTrack -and $raw.FestivalsToTrack.Count -gt 0) {
        $Config.FestivalsToTrack = @($raw.FestivalsToTrack)
        Write-Host "[Config] Loaded $($Config.FestivalsToTrack.Count) festivals from OneDrive." -ForegroundColor Green
        return
    }

    # ── Not found anywhere — fall back ──
    Write-Host '[Config] No festivals_config.json in OneDrive — using built-in FestivalsToTrack list.' -ForegroundColor Gray
    Write-Host '         Tip: run Get-HinduFestivals.ps1 then pass -FestivalsConfig to upload your selection.' -ForegroundColor Gray
}

function Show-ApiKeyDialog {
    param([string]$Current = '')

    $form = [System.Windows.Forms.Form]@{
        Text            = 'Festival Albums — Calendarific API Key'
        Size            = [System.Drawing.Size]::new(520, 280)
        StartPosition   = 'CenterScreen'
        FormBorderStyle = 'FixedDialog'
        MaximizeBox     = $false
        MinimizeBox     = $false
        BackColor       = [System.Drawing.Color]::WhiteSmoke
    }

    $header = [System.Windows.Forms.Label]@{
        Text      = '🗓️  Calendarific API Key Required'
        Location  = [System.Drawing.Point]::new(20, 15)
        Size      = [System.Drawing.Size]::new(470, 28)
        Font      = [System.Drawing.Font]::new('Segoe UI', 13, [System.Drawing.FontStyle]::Bold)
        ForeColor = [System.Drawing.Color]::DarkSlateBlue
    }

    $sub = [System.Windows.Forms.Label]@{
        Text      = 'This key is used to fetch Hindu festival dates. It will be saved securely' +
                    ' in your OneDrive — not on this machine or in any code.'
        Location  = [System.Drawing.Point]::new(20, 47)
        Size      = [System.Drawing.Size]::new(470, 36)
        Font      = [System.Drawing.Font]::new('Segoe UI', 9)
        ForeColor = [System.Drawing.Color]::DimGray
    }

    $lbl = [System.Windows.Forms.Label]@{
        Text     = 'Calendarific API Key'
        Location = [System.Drawing.Point]::new(20, 97)
        Size     = [System.Drawing.Size]::new(200, 20)
        Font     = [System.Drawing.Font]::new('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    }

    $link = [System.Windows.Forms.LinkLabel]@{
        Text     = 'Get a free key →'
        Location = [System.Drawing.Point]::new(230, 97)
        Size     = [System.Drawing.Size]::new(260, 20)
        Font     = [System.Drawing.Font]::new('Segoe UI', 9)
    }
    $link.add_LinkClicked({ Start-Process 'https://calendarific.com/sign-up' })

    $txt = [System.Windows.Forms.TextBox]@{
        Location        = [System.Drawing.Point]::new(20, 121)
        Size            = [System.Drawing.Size]::new(460, 26)
        Font            = [System.Drawing.Font]::new('Consolas', 10)
        Text            = $Current
        PlaceholderText = 'Paste your Calendarific API key here'
    }

    $note = [System.Windows.Forms.Label]@{
        Text      = 'Free tier: 1,000 calls/month. This script uses ~30 calls total per run.'
        Location  = [System.Drawing.Point]::new(20, 153)
        Size      = [System.Drawing.Size]::new(460, 20)
        Font      = [System.Drawing.Font]::new('Segoe UI', 8)
        ForeColor = [System.Drawing.Color]::Gray
    }

    $btnOk = [System.Windows.Forms.Button]@{
        Text         = 'Save to OneDrive'
        Location     = [System.Drawing.Point]::new(310, 200)
        Size         = [System.Drawing.Size]::new(160, 36)
        BackColor    = [System.Drawing.Color]::DarkSlateBlue
        ForeColor    = [System.Drawing.Color]::White
        FlatStyle    = 'Flat'
        Font         = [System.Drawing.Font]::new('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
        DialogResult = 'OK'
    }

    $btnCancel = [System.Windows.Forms.Button]@{
        Text         = 'Cancel'
        Location     = [System.Drawing.Point]::new(220, 200)
        Size         = [System.Drawing.Size]::new(80, 36)
        FlatStyle    = 'Flat'
        Font         = [System.Drawing.Font]::new('Segoe UI', 10)
        DialogResult = 'Cancel'
    }

    $savNote = [System.Windows.Forms.Label]@{
        Text      = '🔒 Saved to OneDrive: Apps/FestivalTimeline/settings.json'
        Location  = [System.Drawing.Point]::new(20, 210)
        Size      = [System.Drawing.Size]::new(360, 20)
        Font      = [System.Drawing.Font]::new('Segoe UI', 8)
        ForeColor = [System.Drawing.Color]::DimGray
    }

    $form.AcceptButton = $btnOk
    $form.CancelButton = $btnCancel
    $form.Controls.AddRange(@($header, $sub, $lbl, $link, $txt, $note, $btnOk, $btnCancel, $savNote))

    if ($form.ShowDialog() -ne 'OK') { Write-Host '[Settings] Cancelled.'; exit 0 }

    $key = $txt.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($key)) {
        [System.Windows.Forms.MessageBox]::Show('API key is required.', 'Missing Value',
            'OK', 'Warning') | Out-Null
        return Show-ApiKeyDialog -Current $key
    }

    return $key
}

# ═══════════════════════════════════════════════════════════════════════════════
#  ONEDRIVE FILE HELPERS  — all calls use Invoke-MgGraphRequest
# ═══════════════════════════════════════════════════════════════════════════════

function Read-OneDriveJson {
    param([string]$RelativePath)
    try {
        $raw = Invoke-MgGraphRequest -Method GET `
            -Uri "$GraphBase/me/drive/root:/$RelativePath`:/content" `
            -OutputType Json -ErrorAction Stop
        return $raw | ConvertFrom-Json
    }
    catch {
        # Invoke-MgGraphRequest surfaces HTTP errors in ErrorDetails.Message (the raw
        # response body), not in Exception.Message — check both to catch 404s silently.
        $detail = "$($_.Exception.Message) $($_.ErrorDetails.Message) $($_.Exception.InnerException)"
        if ($detail -match '404|itemNotFound') { return $null }
        throw
    }
}

function Write-OneDriveJson {
    param([string]$RelativePath, [object]$Data)
    $bytes = [System.Text.Encoding]::UTF8.GetBytes(($Data | ConvertTo-Json -Depth 20 -Compress))
    Invoke-MgGraphRequest -Method PUT `
        -Uri "$GraphBase/me/drive/root:/$RelativePath`:/content" `
        -ContentType 'application/json; charset=utf-8' `
        -Body $bytes | Out-Null
}

function Ensure-OneDriveFolder {
    param([string]$FolderPath)
    $current = ''
    foreach ($part in ($FolderPath -split '/')) {
        $parentUri = if ($current -eq '') {
            "$GraphBase/me/drive/root/children"
        } else {
            "$GraphBase/me/drive/root:/$current`:/children"
        }
        $current = if ($current -eq '') { $part } else { "$current/$part" }
        try {
            Invoke-MgGraphRequest -Method POST -Uri $parentUri `
                -ContentType 'application/json' `
                -Body (@{
                    name = $part
                    folder = @{}
                    '@microsoft.graph.conflictBehavior' = 'replace'
                } | ConvertTo-Json) | Out-Null
        } catch { <# Folder already exists — safe to ignore #> }
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
#  STATE MANAGEMENT
#  Stored as a hashtable in memory, persisted as JSON to OneDrive.
# ═══════════════════════════════════════════════════════════════════════════════

$StatePath = "$($Config.StateFolder)/progress_state.json"

# Recursively converts PSCustomObject (from ConvertFrom-Json) to hashtable
function ConvertTo-Hashtable {
    param($Obj)
    if ($Obj -is [System.Management.Automation.PSCustomObject]) {
        $ht = @{}
        foreach ($p in $Obj.PSObject.Properties) {
            $ht[$p.Name] = ConvertTo-Hashtable $p.Value
        }
        return $ht
    }
    if ($Obj -is [System.Object[]]) {
        return @($Obj | ForEach-Object { ConvertTo-Hashtable $_ })
    }
    return $Obj
}

function New-EmptyState {
    return @{
        schema_version   = 1
        last_run         = $null
        completed_phases = @()
        current_phase    = 'init'
        photo_scan       = @{
            extension_index    = 0      # index into $ImageExtensions currently being scanned
            next_link          = $null  # Graph API pagination token for the current extension
            total_scanned      = 0
            festival_photo_map = @{}    # festival_name -> [file_id, ...]
        }
        albums           = @{}          # festival_name -> album_id
        photos_added     = @{}          # festival_name -> [file_id, ...]
    }
}

function Load-State {
    Write-Host '[State] Loading state from OneDrive...' -ForegroundColor Cyan
    $raw = Read-OneDriveJson -RelativePath $StatePath
    if ($null -eq $raw) {
        Write-Host '[State] No previous state found — starting fresh.' -ForegroundColor Yellow
        return New-EmptyState
    }

    $state = ConvertTo-Hashtable $raw

    # Ensure all keys exist (forward-compatibility with older state files)
    if (-not $state.ContainsKey('photo_scan'))       { $state.photo_scan = @{} }
    if (-not $state.photo_scan.ContainsKey('extension_index'))    { $state.photo_scan.extension_index = 0 }
    if (-not $state.photo_scan.ContainsKey('next_link'))          { $state.photo_scan.next_link = $null }
    if (-not $state.photo_scan.ContainsKey('total_scanned'))      { $state.photo_scan.total_scanned = 0 }
    if (-not $state.photo_scan.ContainsKey('festival_photo_map')) { $state.photo_scan.festival_photo_map = @{} }
    if (-not $state.ContainsKey('albums'))           { $state.albums = @{} }
    if (-not $state.ContainsKey('photos_added'))     { $state.photos_added = @{} }
    if (-not $state.ContainsKey('completed_phases')) { $state.completed_phases = @() }

    if ($Rescan) {
        Write-Host '[State] -Rescan: clearing photo scan progress (albums and added photos preserved).' -ForegroundColor Yellow
        $state.photo_scan       = @{
            extension_index    = 0
            next_link          = $null
            total_scanned      = 0
            festival_photo_map = @{}
        }
        $state.completed_phases = @($state.completed_phases | Where-Object { $_ -ne 'photo_scan' })
    }

    $phases = if ($state.completed_phases.Count -gt 0) { $state.completed_phases -join ', ' } else { 'none' }
    Write-Host "[State] Loaded. Completed phases: $phases" -ForegroundColor Green
    return $state
}

function Save-State {
    param([hashtable]$State)
    if ($DryRun) {
        Write-Host '[State] DryRun — state not saved.' -ForegroundColor DarkGray
        return
    }
    $State.last_run = (Get-Date -Format 'o')
    Write-OneDriveJson -RelativePath $StatePath -Data $State
    Write-Host '[State] Progress saved to OneDrive.' -ForegroundColor Gray
}

# ═══════════════════════════════════════════════════════════════════════════════
#  CALENDAR CACHE
#  Festival dates fetched from Calendarific, cached in OneDrive.
#  Returns flat hashtable: { "YYYY-MM-DD" = "FestivalName" }
# ═══════════════════════════════════════════════════════════════════════════════

$CalendarPath = "$($Config.StateFolder)/calendar_cache.json"

function Get-FestivalDates {
    Write-Host "`n[Calendar] Checking festival date cache..." -ForegroundColor Cyan

    $cache       = Read-OneDriveJson -RelativePath $CalendarPath
    $needRefresh = $true

    if ($null -ne $cache) {
        try {
            $age = ((Get-Date) - [DateTime]::Parse($cache.cached_on)).TotalDays
            if ($age -lt $Config.CacheRefreshDays) {
                Write-Host "[Calendar] Cache is $([int]$age) days old — using cached data." -ForegroundColor Green
                $needRefresh = $false
            } else {
                Write-Host "[Calendar] Cache is $([int]$age) days old — refreshing." -ForegroundColor Yellow
            }
        } catch {
            Write-Host '[Calendar] Cache timestamp unreadable — refreshing.' -ForegroundColor Yellow
        }
    } else {
        Write-Host '[Calendar] No cache found — downloading festival dates.' -ForegroundColor Yellow
    }

    if ($needRefresh) {
        $festivalDates = @{}
        $currentYear   = (Get-Date).Year
        $startYear     = $currentYear - $Config.YearsToScan
        $totalYears    = $Config.YearsToScan + 1

        Write-Host "[Calendar] Fetching $totalYears years ($startYear–$currentYear) from Calendarific..." -ForegroundColor Cyan

        for ($yr = $startYear; $yr -le $currentYear; $yr++) {
            $pct = [int](($yr - $startYear) / $totalYears * 100)
            Write-Progress -Activity 'Downloading Hindu calendar' -Status "Year $yr" -PercentComplete $pct
            try {
                $url      = "https://calendarific.com/api/v2/holidays?api_key=$($Config.CalendarificApiKey)&country=IN&year=$yr&type=religious"
                $response = Invoke-RestMethod -Uri $url -Method GET -ErrorAction Stop
                foreach ($h in $response.response.holidays) {
                    if ($h.name -in $Config.FestivalsToTrack) {
                        $dateKey = $h.date.iso.Substring(0, 10)  # YYYY-MM-DD only
                        $festivalDates[$dateKey] = $h.name
                    }
                }
            } catch {
                Write-Warning "[Calendar] Could not fetch year $yr`: $_"
            }
        }
        Write-Progress -Activity 'Downloading Hindu calendar' -Completed

        $cache = @{
            cached_on      = (Get-Date -Format 'o')
            festival_dates = $festivalDates
        }
        Write-OneDriveJson -RelativePath $CalendarPath -Data $cache
        Write-Host "[Calendar] Cached $($festivalDates.Count) festival date entries to OneDrive." -ForegroundColor Green
    }

    # Build flat hashtable for O(1) date lookup
    $lookup  = @{}
    $dateMap = $cache.festival_dates
    if ($dateMap -is [System.Management.Automation.PSCustomObject]) {
        foreach ($p in $dateMap.PSObject.Properties) { $lookup[$p.Name] = $p.Value }
    } elseif ($dateMap -is [hashtable]) {
        foreach ($k in $dateMap.Keys)                { $lookup[$k]      = $dateMap[$k] }
    }

    Write-Host "[Calendar] $($lookup.Count) festival dates ready." -ForegroundColor Green
    return $lookup
}

# ═══════════════════════════════════════════════════════════════════════════════
#  PHOTO SCAN
#  Paginates through all OneDrive photos by image extension.
#  Saves extension index + page token after each page so the scan
#  can resume exactly where it left off after any interruption.
# ═══════════════════════════════════════════════════════════════════════════════

function Invoke-PhotoScan {
    param([hashtable]$State, [hashtable]$FestivalDates)

    Write-Host "`n[Scan] Starting photo scan..." -ForegroundColor Cyan

    # Restore partially-built photo map from state
    $photoMap = @{}
    foreach ($k in $State.photo_scan.festival_photo_map.Keys) {
        $photoMap[$k] = [System.Collections.Generic.List[string]]($State.photo_scan.festival_photo_map[$k])
    }

    $scanned          = [int]$State.photo_scan.total_scanned
    $script:ExifCount = 0   # tracks how many photos had real EXIF takenDateTime
    $resumeExtIdx = [int]$State.photo_scan.extension_index
    $resumeLink   = $State.photo_scan.next_link

    for ($extIdx = $resumeExtIdx; $extIdx -lt $ImageExtensions.Count; $extIdx++) {
        $ext = $ImageExtensions[$extIdx]

        # On the resumed extension use the saved nextLink; all others start fresh
        $uri = if ($extIdx -eq $resumeExtIdx -and $resumeLink) {
            $resumeLink
        } else {
            "$GraphBase/me/drive/root/search(q='$ext')?`$select=id,name,photo,createdDateTime,file&`$top=200"
        }
        $resumeLink = $null  # only apply saved link once

        Write-Host "[Scan] Extension $($extIdx+1)/$($ImageExtensions.Count): $ext" -ForegroundColor Gray

        while ($uri) {
            $page = Invoke-Graph -Uri $uri

            foreach ($item in $page.value) {
                if (-not $item.file) { continue }  # skip folders

                # Prefer EXIF capture date (photo facet), fall back to file creation date.
                # Safe null-chain: .photo is absent on non-JPEG files, chaining onto
                # $null throws in strict mode, so we guard with PSObject.Properties.
                $takenDateTime = if ($item.photo -ne $null) {
                    $item.photo.PSObject.Properties['takenDateTime']?.Value
                } else { $null }
                $rawDate = if ($takenDateTime) { $takenDateTime } else { $item.createdDateTime }
                if ($takenDateTime) { $script:ExifCount++ }

                $dateKey = try { ([DateTime]::Parse($rawDate)).ToString('yyyy-MM-dd') } catch { $null }

                if ($dateKey -and $FestivalDates.ContainsKey($dateKey)) {
                    $festival = $FestivalDates[$dateKey]
                    if (-not $photoMap.ContainsKey($festival)) {
                        $photoMap[$festival] = [System.Collections.Generic.List[string]]::new()
                    }
                    if ($item.id -notin $photoMap[$festival]) {
                        $photoMap[$festival].Add($item.id)
                    }
                }
                $scanned++
            }

            $totalFound = ($photoMap.Values | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum
            Write-Progress -Activity 'Scanning OneDrive photos' `
                -Status "Extension: $ext  |  Scanned: $scanned  |  Festival photos: $totalFound"

            $nextUri = $page.'@odata.nextLink'

            # Checkpoint: save current extension, page token, and photo map
            $mapSnapshot = @{}
            foreach ($k in $photoMap.Keys) { $mapSnapshot[$k] = @($photoMap[$k]) }

            $State.photo_scan.extension_index    = $extIdx
            $State.photo_scan.next_link          = $nextUri
            $State.photo_scan.total_scanned      = $scanned
            $State.photo_scan.festival_photo_map = $mapSnapshot
            Save-State -State $State

            $uri = $nextUri
        }

        # Extension complete — advance index and clear page token
        $State.photo_scan.extension_index = $extIdx + 1
        $State.photo_scan.next_link       = $null
        Save-State -State $State
    }

    Write-Progress -Activity 'Scanning OneDrive photos' -Completed
    $fallbackCount = $scanned - $script:ExifCount
    Write-Host "[Scan] Complete. Total files scanned: $scanned  |  EXIF date used: $($script:ExifCount)  |  Fallback to created date: $fallbackCount" -ForegroundColor Green
    foreach ($festival in ($photoMap.Keys | Sort-Object)) {
        Write-Host "  $festival`: $($photoMap[$festival].Count) photo(s)" -ForegroundColor Gray
    }

    # Finalise state
    $finalMap = @{}
    foreach ($k in $photoMap.Keys) { $finalMap[$k] = @($photoMap[$k]) }

    $State.photo_scan.festival_photo_map = $finalMap
    $State.photo_scan.next_link          = $null
    $State.photo_scan.total_scanned      = $scanned
    $State.completed_phases              = @($State.completed_phases) + 'photo_scan'
    Save-State -State $State

    return $finalMap
}

# ═══════════════════════════════════════════════════════════════════════════════
#  ALBUM MANAGEMENT
#  Uses OneDrive bundles (photo albums). Photos are added by reference —
#  they stay in their original location. No copies, no moves.
# ═══════════════════════════════════════════════════════════════════════════════

function Get-OrCreate-Album {
    param([string]$FestivalName, [hashtable]$State)

    $albumName = "$FestivalName $AlbumSuffix"

    # Fastest path: already tracked in state
    if ($State.albums.ContainsKey($FestivalName)) {
        $id = $State.albums[$FestivalName]
        Write-Host "[Album] '$albumName' found in state: $id" -ForegroundColor Gray
        return $id
    }

    # Search OneDrive bundles for an existing album with this name
    Write-Host "[Album] Searching OneDrive for existing album '$albumName'..." -ForegroundColor Cyan
    try {
        $bundles  = Invoke-Graph -Uri "$GraphBase/me/drive/bundles?`$select=id,name"
        $existing = $bundles.value | Where-Object { $_.name -eq $albumName } | Select-Object -First 1
        if ($existing) {
            Write-Host "[Album] Found '$albumName': $($existing.id)" -ForegroundColor Green
            $State.albums[$FestivalName] = $existing.id
            return $existing.id
        }
    } catch {
        Write-Warning "[Album] Could not list bundles: $_"
    }

    # Create a new album
    if ($DryRun) {
        Write-Host "[DryRun] Would create album '$albumName'" -ForegroundColor DarkYellow
        return 'dry-run-album-id'
    }
    Write-Host "[Album] Creating new album '$albumName'..." -ForegroundColor Cyan
    $newAlbum = Invoke-Graph -Method POST -Uri "$GraphBase/me/drive/bundles" -Body @{
        name                                = $albumName
        bundle                              = @{ album = @{} }
        '@microsoft.graph.conflictBehavior' = 'fail'
    }
    Write-Host "[Album] Created '$albumName': $($newAlbum.id)" -ForegroundColor Green
    $State.albums[$FestivalName] = $newAlbum.id
    return $newAlbum.id
}

function Add-PhotosToAlbum {
    param(
        [string]   $FestivalName,
        [string]   $AlbumId,
        [string[]] $FileIds,
        [hashtable]$State
    )

    $albumName = "$FestivalName $AlbumSuffix"
    Write-Host "`n[Add] Adding photos to '$albumName'..." -ForegroundColor Cyan

    # Build a hash set of already-added IDs for O(1) skip lookup
    if (-not $State.photos_added.ContainsKey($FestivalName)) {
        $State.photos_added[$FestivalName] = @()
    }
    $alreadyAdded = @{}
    foreach ($id in $State.photos_added[$FestivalName]) { $alreadyAdded[$id] = $true }

    $added   = 0
    $skipped = 0
    $failed  = 0
    $batch   = 0
    $total   = $FileIds.Count

    foreach ($fileId in $FileIds) {

        # Skip photos already in this album
        if ($alreadyAdded.ContainsKey($fileId)) {
            $skipped++
            continue
        }

        try {
            # DryRun: skip the actual Graph API write, just count
            if ($DryRun) {
                $added++
                continue
            }

            # Add photo to album by reference — photo stays in its original location
            Invoke-Graph -Method POST `
                -Uri "$GraphBase/me/drive/bundles/$AlbumId/children" `
                -Body @{ id = $fileId } | Out-Null

            $State.photos_added[$FestivalName] = @($State.photos_added[$FestivalName]) + $fileId
            $alreadyAdded[$fileId] = $true
            $added++
            $batch++

            if ($batch -ge $Config.CheckpointEvery) {
                Save-State -State $State
                $batch = 0
            }
        } catch {
            Write-Warning "[Add] Failed to add photo $fileId`: $_"
            $failed++
        }

        $done = $added + $skipped + $failed
        Write-Progress -Activity "Adding to $albumName" `
            -Status "Added: $added  |  Skipped: $skipped  |  Failed: $failed" `
            -PercentComplete ([int]($done / [Math]::Max($total, 1) * 100))
    }

    Write-Progress -Activity "Adding to $albumName" -Completed

    # Final checkpoint for this festival
    Save-State -State $State
    Write-Host "[Add] $albumName — Added: $added  |  Already in album (skipped): $skipped  |  Failed: $failed" -ForegroundColor Green
}

# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════════════════════════

$bar = '═' * 62
Write-Host "`n$bar" -ForegroundColor Cyan
Write-Host '  Hindu Festival Photo Album Creator' -ForegroundColor Cyan
Write-Host "  $(Get-Date -Format 'yyyy-MM-dd HH:mm')  |  OneDrive Personal" -ForegroundColor Cyan
Write-Host "$bar" -ForegroundColor Cyan
Write-Host ''
if ($Rescan) {
    Write-Host '  Mode: FULL RESCAN (ignoring previous scan progress)' -ForegroundColor Yellow
} elseif ($DryRun) {
    Write-Host '  Mode: DRY RUN — no albums created, no photos added, no state saved' -ForegroundColor Magenta
    Write-Host '         API key WILL be saved to OneDrive if not already present.' -ForegroundColor DarkGray
} else {
    Write-Host '  Mode: Resume from last checkpoint' -ForegroundColor Green
}
Write-Host "  Tracking festivals: $($Config.FestivalsToTrack -join ', ')" -ForegroundColor Gray
Write-Host ''

# ── Phase 1: Authenticate ──────────────────────────────────────────────────
Connect-ToOneDrive

# Ensure the state folder exists in OneDrive before any reads/writes
Ensure-OneDriveFolder -FolderPath $Config.StateFolder

# ── Load API key from OneDrive (after auth so we can read OneDrive) ─────────
Write-Host "`n[Settings] Checking for Calendarific API key in OneDrive..." -ForegroundColor Cyan
$apiKey = Load-OneDriveSettings
if (-not $apiKey) {
    Write-Host '[Settings] API key not found in OneDrive — opening dialog...' -ForegroundColor Yellow
    $apiKey = Show-ApiKeyDialog
    # Always save the key to OneDrive — even in DryRun (this is the one write DryRun allows)
    Save-OneDriveSettings -ApiKey $apiKey
} else {
    Write-Host '[Settings] API key found in OneDrive.' -ForegroundColor Green
}
$Config.CalendarificApiKey = $apiKey

# ── Load festival selection (local file → OneDrive → built-in fallback) ─────
Load-FestivalsConfig

# ── Load State ─────────────────────────────────────────────────────────────
$state = Load-State

# ── Phase 2: Calendar Cache ────────────────────────────────────────────────
$festivalDates = Get-FestivalDates

# ── Phase 3: Photo Scan ────────────────────────────────────────────────────
$photoMap = $null

if ('photo_scan' -in $state.completed_phases) {
    Write-Host "`n[Scan] Photo scan already complete — loading results from state." -ForegroundColor Green
    $photoMap = $state.photo_scan.festival_photo_map
} else {
    $photoMap = Invoke-PhotoScan -State $state -FestivalDates $festivalDates
}

# ── Phase 4 + 5: Create Albums and Add Photos ──────────────────────────────
$festivalsWithPhotos = @($photoMap.Keys | Where-Object { @($photoMap[$_]).Count -gt 0 } | Sort-Object)

if ($festivalsWithPhotos.Count -eq 0) {
    Write-Host "`n[Albums] No festival photos found in your OneDrive." -ForegroundColor Yellow
    Write-Host '         Tip: verify that FestivalsToTrack names match Calendarific exactly.' -ForegroundColor Gray
} else {
    Write-Host "`n[Albums] Processing $($festivalsWithPhotos.Count) festival(s)..." -ForegroundColor Cyan

    foreach ($festival in $festivalsWithPhotos) {
        $fileIds = @($photoMap[$festival])
        Write-Host "`n  ── $festival ($($fileIds.Count) photo(s)) ──" -ForegroundColor Yellow

        $albumId = Get-OrCreate-Album -FestivalName $festival -State $state
        Save-State -State $state

        Add-PhotosToAlbum -FestivalName $festival -AlbumId $albumId -FileIds $fileIds -State $state
    }
}

# ── Mark run complete ──────────────────────────────────────────────────────
if (-not $DryRun) {
    $state.completed_phases = @($state.completed_phases) + 'done'
    $state.current_phase    = 'complete'
    Save-State -State $state
}

if ($DryRun) {
    $totalPhotos = ($festivalsWithPhotos | ForEach-Object { @($photoMap[$_]).Count } | Measure-Object -Sum).Sum
    Write-Host "`n$bar" -ForegroundColor Magenta
    Write-Host '  DRY RUN COMPLETE — nothing was changed in OneDrive.' -ForegroundColor Magenta
    Write-Host "$bar" -ForegroundColor Magenta
    Write-Host ''
    Write-Host '  What WOULD have happened:' -ForegroundColor White
    Write-Host "    Albums to create/reuse : $($festivalsWithPhotos.Count)" -ForegroundColor White
    Write-Host "    Photos to add (total)  : $totalPhotos" -ForegroundColor White
    Write-Host ''
    foreach ($festival in $festivalsWithPhotos) {
        $count = @($photoMap[$festival]).Count
        Write-Host ("    {0,-30} {1,4} photo(s) → {2} {3}" -f $festival, $count, $festival, $AlbumSuffix) -ForegroundColor Gray
    }
    Write-Host ''
    Write-Host '  To apply these changes, run:  .\FestivalAlbums.ps1' -ForegroundColor Yellow
    Write-Host "$bar`n" -ForegroundColor Magenta
} else {
    Write-Host "`n$bar" -ForegroundColor Green
    Write-Host '  All done! Your festival albums have been updated.' -ForegroundColor Green
    Write-Host '  Open OneDrive → Photos → Albums to see them.' -ForegroundColor Green
    Write-Host "$bar`n" -ForegroundColor Green
}
