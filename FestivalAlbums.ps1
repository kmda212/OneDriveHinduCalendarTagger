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
    [switch]$Rescan
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ═══════════════════════════════════════════════════════════════════════════════
#  CONFIGURATION  — Edit this section before running for the first time
# ═══════════════════════════════════════════════════════════════════════════════

$Config = @{

    # Your Azure App Registration Client ID.
    # Leave blank to see setup instructions.
    ClientId           = ''

    # Your Calendarific API key.
    # Leave blank to see setup instructions.
    CalendarificApiKey = ''

    # How many past years to scan for festival photos.
    YearsToScan        = 30

    # Re-download calendar if the cached copy is older than this (days).
    CacheRefreshDays   = 90

    # Write progress to OneDrive after every N photo additions.
    # Lower = safer resume, but slightly more API calls.
    CheckpointEvery    = 20

    # OneDrive folder where state and cache files are stored (auto-created).
    StateFolder        = 'Apps/FestivalTimeline'

    # ── Festivals to create albums for ───────────────────────────────────────
    # Names must match Calendarific API holiday names for country=IN exactly.
    # Add or remove lines freely before running.
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

# Album name = festival name + this suffix  (e.g. "DiwaliLifetime")
$AlbumSuffix = 'Lifetime'

# Image file extensions to scan
$ImageExtensions = @('.jpg', '.jpeg', '.heic', '.png', '.gif', '.bmp', '.tiff')

# Microsoft Graph base URL
$GraphBase = 'https://graph.microsoft.com/v1.0'

# ═══════════════════════════════════════════════════════════════════════════════
#  FIRST-TIME SETUP GUIDE
# ═══════════════════════════════════════════════════════════════════════════════

function Show-SetupGuide {
    param([string[]]$Missing)

    $bar = '═' * 62
    Write-Host "`n$bar" -ForegroundColor Cyan
    Write-Host '  FIRST-TIME SETUP REQUIRED' -ForegroundColor Yellow
    Write-Host "$bar`n" -ForegroundColor Cyan

    $step = 1

    if ('ClientId' -in $Missing) {
        Write-Host "  STEP $step — Register a Microsoft Azure App" -ForegroundColor Green
        Write-Host ('  ' + '─' * 55)
        Write-Host '  This gives the script permission to access YOUR OneDrive.'
        Write-Host '  Takes ~3 minutes. No credit card required.'
        Write-Host ''
        Write-Host '  1. Open in your browser:'
        Write-Host '       https://aka.ms/AppRegistrations' -ForegroundColor Cyan
        Write-Host '     Sign in with the same account as your OneDrive.'
        Write-Host ''
        Write-Host '  2. Click [ + New registration ]'
        Write-Host '       Name:                    Festival Photo Albums'
        Write-Host '       Supported account types: Personal Microsoft accounts only'
        Write-Host '       Redirect URI:            (leave blank)'
        Write-Host '       Click [ Register ]'
        Write-Host ''
        Write-Host '  3. On the app Overview page:'
        Write-Host '       Copy "Application (client) ID"'
        Write-Host '       Paste it into $Config.ClientId in this script.'
        Write-Host ''
        Write-Host '  4. Left menu → Authentication'
        Write-Host '       → Add a platform → Mobile and desktop applications'
        Write-Host '       → Check: https://login.microsoftonline.com/common/oauth2/nativeclient'
        Write-Host '       → Advanced settings → Allow public client flows = Yes'
        Write-Host '       → [ Save ]'
        Write-Host ''
        Write-Host '  5. Left menu → API permissions'
        Write-Host '       → Add a permission → Microsoft Graph → Delegated'
        Write-Host '       → Search "Files.ReadWrite" → check it → Add permissions'
        Write-Host ''
        $step++
    }

    if ('CalendarificApiKey' -in $Missing) {
        Write-Host "  STEP $step — Get a Free Calendarific API Key" -ForegroundColor Green
        Write-Host ('  ' + '─' * 55)
        Write-Host '  Provides Hindu festival dates for all years.'
        Write-Host '  Free tier: 1,000 API calls/month. This script uses ~30 calls total.'
        Write-Host ''
        Write-Host '  1. Open: https://calendarific.com/sign-up' -ForegroundColor Cyan
        Write-Host '  2. Create a free account and verify your email.'
        Write-Host '  3. After login, go to your Dashboard.'
        Write-Host '  4. Copy the API Key shown on the dashboard.'
        Write-Host '       Paste it into $Config.CalendarificApiKey in this script.'
        Write-Host ''
    }

    Write-Host '  After updating the config, re-run:  .\FestivalAlbums.ps1' -ForegroundColor Yellow
    Write-Host "$bar`n" -ForegroundColor Cyan
    exit 0
}

# Validate config before doing anything else
$missingConfig = @()
if ([string]::IsNullOrWhiteSpace($Config.ClientId))           { $missingConfig += 'ClientId' }
if ([string]::IsNullOrWhiteSpace($Config.CalendarificApiKey)) { $missingConfig += 'CalendarificApiKey' }
if ($missingConfig.Count -gt 0) { Show-SetupGuide -Missing $missingConfig }

# ═══════════════════════════════════════════════════════════════════════════════
#  AUTHENTICATION  — Microsoft Graph Device Code Flow
# ═══════════════════════════════════════════════════════════════════════════════

$script:AccessToken = $null
$script:TokenExpiry = [DateTime]::MinValue

function Get-AccessToken {
    Write-Host '[Auth] Requesting device code...' -ForegroundColor Cyan

    $deviceCode = Invoke-RestMethod -Method POST `
        -Uri 'https://login.microsoftonline.com/consumers/oauth2/v2.0/devicecode' `
        -ContentType 'application/x-www-form-urlencoded' `
        -Body @{
            client_id = $Config.ClientId
            scope     = 'https://graph.microsoft.com/Files.ReadWrite offline_access'
        }

    $bar = '═' * 62
    Write-Host "`n$bar" -ForegroundColor Yellow
    Write-Host '  ACTION REQUIRED — Sign in to Microsoft' -ForegroundColor Yellow
    Write-Host "$bar" -ForegroundColor Yellow
    Write-Host ''
    Write-Host '  1. Open this URL in your browser:'
    Write-Host "       $($deviceCode.verification_uri)" -ForegroundColor Cyan
    Write-Host ''
    Write-Host '  2. Enter this code when prompted:'
    Write-Host "       $($deviceCode.user_code)" -ForegroundColor Green
    Write-Host ''
    Write-Host '  Waiting for sign-in...' -ForegroundColor Gray

    $interval  = [int]$deviceCode.interval
    $expiresIn = [int]$deviceCode.expires_in
    $elapsed   = 0

    while ($elapsed -lt $expiresIn) {
        Start-Sleep -Seconds $interval
        $elapsed += $interval
        try {
            $token = Invoke-RestMethod -Method POST `
                -Uri 'https://login.microsoftonline.com/consumers/oauth2/v2.0/token' `
                -ContentType 'application/x-www-form-urlencoded' `
                -Body @{
                    client_id   = $Config.ClientId
                    grant_type  = 'urn:ietf:params:oauth:grant-type:device_code'
                    device_code = $deviceCode.device_code
                }
            $script:AccessToken = $token.access_token
            $script:TokenExpiry = (Get-Date).AddSeconds([int]$token.expires_in - 60)
            Write-Host '[Auth] Signed in successfully.' -ForegroundColor Green
            return
        }
        catch {
            $err = $null
            try { $err = ($_.ErrorDetails.Message | ConvertFrom-Json).error } catch {}
            if ($err -eq 'authorization_pending') { continue }
            if ($err -eq 'expired_token') { Write-Error '[Auth] Code expired. Re-run the script.'; return }
            throw
        }
    }
    Write-Error '[Auth] Timed out waiting for sign-in.'
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

    # Re-authenticate if token is close to expiry
    if ((Get-Date) -ge $script:TokenExpiry) {
        Write-Host '[Auth] Token expired, re-authenticating...' -ForegroundColor Yellow
        Get-AccessToken
    }

    $params = @{
        Method  = $Method
        Uri     = $Uri
        Headers = @{ Authorization = "Bearer $script:AccessToken" }
    }

    if ($Body) {
        $params.ContentType = 'application/json'
        $params.Body        = ($Body | ConvertTo-Json -Depth 10)
    }

    return Invoke-RestMethod @params
}

# ═══════════════════════════════════════════════════════════════════════════════
#  ONEDRIVE FILE HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

function Read-OneDriveJson {
    param([string]$RelativePath)
    try {
        return Invoke-RestMethod -Method GET `
            -Uri "$GraphBase/me/drive/root:/$RelativePath`:/content" `
            -Headers @{ Authorization = "Bearer $script:AccessToken" }
    }
    catch {
        $code = $_.Exception.Response.StatusCode.value__
        if ($code -eq 404) { return $null }
        throw
    }
}

function Write-OneDriveJson {
    param([string]$RelativePath, [object]$Data)
    $bytes = [System.Text.Encoding]::UTF8.GetBytes(($Data | ConvertTo-Json -Depth 20 -Compress))
    Invoke-RestMethod -Method PUT `
        -Uri "$GraphBase/me/drive/root:/$RelativePath`:/content" `
        -Headers @{ Authorization = "Bearer $script:AccessToken" } `
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
            Invoke-RestMethod -Method POST -Uri $parentUri `
                -Headers @{ Authorization = "Bearer $script:AccessToken" } `
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

    $scanned      = [int]$State.photo_scan.total_scanned
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

                # Prefer EXIF capture date, fall back to file creation date
                $rawDate = if ($item.photo -and $item.photo.takenDateTime) {
                    $item.photo.takenDateTime
                } else {
                    $item.createdDateTime
                }

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
    Write-Host "[Scan] Complete. Total files scanned: $scanned" -ForegroundColor Green
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

    $albumName = "$FestivalName$AlbumSuffix"

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

    $albumName = "$FestivalName$AlbumSuffix"
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
} else {
    Write-Host '  Mode: Resume from last checkpoint' -ForegroundColor Green
}
Write-Host "  Tracking festivals: $($Config.FestivalsToTrack -join ', ')" -ForegroundColor Gray
Write-Host ''

# ── Phase 1: Authenticate ──────────────────────────────────────────────────
Get-AccessToken

# Ensure the state folder exists in OneDrive before any reads/writes
Ensure-OneDriveFolder -FolderPath $Config.StateFolder

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
$state.completed_phases = @($state.completed_phases) + 'done'
$state.current_phase    = 'complete'
Save-State -State $state

Write-Host "`n$bar" -ForegroundColor Green
Write-Host '  All done! Your festival albums have been updated.' -ForegroundColor Green
Write-Host '  Open OneDrive → Photos → Albums to see them.' -ForegroundColor Green
Write-Host "$bar`n" -ForegroundColor Green
