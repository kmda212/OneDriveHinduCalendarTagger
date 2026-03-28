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

# Required for HttpUtility.ParseQueryString (OAuth redirect parsing)
Add-Type -AssemblyName System.Web

# ═══════════════════════════════════════════════════════════════════════════════
#  CONFIGURATION  — Credentials are collected via popup on first run and saved
#                   locally. All other settings can be changed here freely.
# ═══════════════════════════════════════════════════════════════════════════════

$Config = @{

    # ClientId is the only value stored locally (%APPDATA%\FestivalAlbums\config.json).
    # The Calendarific API key is stored securely in OneDrive (Apps/FestivalTimeline/settings.json)
    # and is loaded after sign-in. Neither value is stored in this script.

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

# Load festival selection from festivals_config.json if present (overrides built-in list)
$FestivalsConfigPath = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) 'festivals_config.json'
if (Test-Path $FestivalsConfigPath) {
    try {
        $fc = Get-Content $FestivalsConfigPath -Raw | ConvertFrom-Json
        if ($fc.FestivalsToTrack -and $fc.FestivalsToTrack.Count -gt 0) {
            $Config.FestivalsToTrack = @($fc.FestivalsToTrack)
            Write-Host "[Config] Loaded $($Config.FestivalsToTrack.Count) festivals from festivals_config.json" -ForegroundColor Green
        }
    } catch {
        Write-Warning "[Config] Could not read festivals_config.json — using built-in list."
    }
} else {
    Write-Host "[Config] No festivals_config.json found — using built-in FestivalsToTrack list." -ForegroundColor Gray
    Write-Host "         Tip: run Get-HinduFestivals.ps1 to build your own selection." -ForegroundColor Gray
}

# Local path — stores only the Azure App ClientId (not the API key)
$LocalConfigPath = "$env:APPDATA\FestivalAlbums\config.json"

# OneDrive path — stores the Calendarific API key (never in git, never on disk)
$OneDriveSettingsPath = "$($Config.StateFolder)/settings.json"

# Album name = festival name + this suffix  (e.g. "DiwaliLifetime")
$AlbumSuffix = 'Lifetime'

# Image file extensions to scan
$ImageExtensions = @('.jpg', '.jpeg', '.heic', '.png', '.gif', '.bmp', '.tiff')

# Microsoft Graph base URL
$GraphBase = 'https://graph.microsoft.com/v1.0'

# ═══════════════════════════════════════════════════════════════════════════════
#  LOCAL CONFIG  — Stores only the Azure App ClientId on this machine.
#  The Calendarific API key is stored in OneDrive instead (see below).
# ═══════════════════════════════════════════════════════════════════════════════

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Load-LocalConfig {
    if (Test-Path $LocalConfigPath) {
        try {
            $raw = Get-Content $LocalConfigPath -Raw | ConvertFrom-Json
            if ($raw.ClientId) { return @{ ClientId = $raw.ClientId } }
        } catch {}
    }
    return $null
}

function Save-LocalConfig {
    param([string]$ClientId)
    $dir = Split-Path $LocalConfigPath
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
    @{ ClientId = $ClientId } | ConvertTo-Json | Set-Content -Path $LocalConfigPath -Encoding UTF8
}

function Show-ClientIdDialog {
    param([string]$Current = '')

    $form = [System.Windows.Forms.Form]@{
        Text            = 'Festival Albums — Azure App Setup'
        Size            = [System.Drawing.Size]::new(520, 310)
        StartPosition   = 'CenterScreen'
        FormBorderStyle = 'FixedDialog'
        MaximizeBox     = $false
        MinimizeBox     = $false
        BackColor       = [System.Drawing.Color]::WhiteSmoke
    }

    $header = [System.Windows.Forms.Label]@{
        Text      = '🪔  Festival Albums for OneDrive'
        Location  = [System.Drawing.Point]::new(20, 15)
        Size      = [System.Drawing.Size]::new(470, 28)
        Font      = [System.Drawing.Font]::new('Segoe UI', 13, [System.Drawing.FontStyle]::Bold)
        ForeColor = [System.Drawing.Color]::DarkSlateBlue
    }

    $sub = [System.Windows.Forms.Label]@{
        Text      = 'Enter your Azure App Client ID to allow this script to access your OneDrive.'
        Location  = [System.Drawing.Point]::new(20, 47)
        Size      = [System.Drawing.Size]::new(470, 20)
        Font      = [System.Drawing.Font]::new('Segoe UI', 9)
        ForeColor = [System.Drawing.Color]::DimGray
    }

    $lbl = [System.Windows.Forms.Label]@{
        Text     = 'Azure App Client ID'
        Location = [System.Drawing.Point]::new(20, 90)
        Size     = [System.Drawing.Size]::new(200, 20)
        Font     = [System.Drawing.Font]::new('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    }

    $link = [System.Windows.Forms.LinkLabel]@{
        Text     = 'How to register a free app →'
        Location = [System.Drawing.Point]::new(230, 90)
        Size     = [System.Drawing.Size]::new(260, 20)
        Font     = [System.Drawing.Font]::new('Segoe UI', 9)
    }
    $link.add_LinkClicked({ Start-Process 'https://aka.ms/AppRegistrations' })

    $txt = [System.Windows.Forms.TextBox]@{
        Location        = [System.Drawing.Point]::new(20, 114)
        Size            = [System.Drawing.Size]::new(460, 26)
        Font            = [System.Drawing.Font]::new('Consolas', 10)
        Text            = $Current
        PlaceholderText = 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'
    }

    $note = [System.Windows.Forms.Label]@{
        Text      = 'Register at portal.azure.com → App Registrations. Personal accounts only.' +
                    ' Add redirect URI http://localhost:8765/ and Files.ReadWrite permission.'
        Location  = [System.Drawing.Point]::new(20, 146)
        Size      = [System.Drawing.Size]::new(460, 36)
        Font      = [System.Drawing.Font]::new('Segoe UI', 8)
        ForeColor = [System.Drawing.Color]::Gray
    }

    $btnOk = [System.Windows.Forms.Button]@{
        Text         = 'Save && Continue'
        Location     = [System.Drawing.Point]::new(320, 224)
        Size         = [System.Drawing.Size]::new(155, 36)
        BackColor    = [System.Drawing.Color]::DarkSlateBlue
        ForeColor    = [System.Drawing.Color]::White
        FlatStyle    = 'Flat'
        Font         = [System.Drawing.Font]::new('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
        DialogResult = 'OK'
    }

    $btnCancel = [System.Windows.Forms.Button]@{
        Text         = 'Cancel'
        Location     = [System.Drawing.Point]::new(230, 224)
        Size         = [System.Drawing.Size]::new(80, 36)
        FlatStyle    = 'Flat'
        Font         = [System.Drawing.Font]::new('Segoe UI', 10)
        DialogResult = 'Cancel'
    }

    $savNote = [System.Windows.Forms.Label]@{
        Text      = '🔒 Saved locally to %APPDATA%\FestivalAlbums\config.json'
        Location  = [System.Drawing.Point]::new(20, 234)
        Size      = [System.Drawing.Size]::new(360, 20)
        Font      = [System.Drawing.Font]::new('Segoe UI', 8)
        ForeColor = [System.Drawing.Color]::DimGray
    }

    $form.AcceptButton = $btnOk
    $form.CancelButton = $btnCancel
    $form.Controls.AddRange(@($header, $sub, $lbl, $link, $txt, $note, $btnOk, $btnCancel, $savNote))

    if ($form.ShowDialog() -ne 'OK') { Write-Host '[Setup] Cancelled.'; exit 0 }

    $clientId = $txt.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($clientId)) {
        [System.Windows.Forms.MessageBox]::Show('Client ID is required.', 'Missing Value',
            'OK', 'Warning') | Out-Null
        return Show-ClientIdDialog -Current $clientId
    }

    Save-LocalConfig -ClientId $clientId
    Write-Host '[Setup] Client ID saved locally.' -ForegroundColor Green
    return $clientId
}

# Load or prompt for ClientId
$localCreds = Load-LocalConfig
if (-not $localCreds) {
    Write-Host '[Setup] No Client ID found — opening setup dialog...' -ForegroundColor Yellow
    $Config.ClientId = Show-ClientIdDialog
} else {
    $Config.ClientId = $localCreds.ClientId
}

# ═══════════════════════════════════════════════════════════════════════════════
#  ONEDRIVE SETTINGS  — Calendarific API key stored in OneDrive, not locally.
#  Loaded after authentication. Popup shown if key is missing or blank.
#  File: Apps/FestivalTimeline/settings.json  (inside your OneDrive)
# ═══════════════════════════════════════════════════════════════════════════════

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
#  AUTHENTICATION  — Microsoft Graph Authorization Code Flow (browser popup)
#  Opens the Microsoft login page in the default browser. A local HTTP listener
#  on localhost catches the OAuth redirect and extracts the auth code.
#  No credentials are typed into the terminal.
# ═══════════════════════════════════════════════════════════════════════════════

$script:AccessToken  = $null
$script:RefreshToken = $null
$script:TokenExpiry  = [DateTime]::MinValue

# Fixed redirect URI — must also be registered in your Azure app:
#   Authentication → Add platform → Web → http://localhost:8765/
$OAuthRedirectUri = 'http://localhost:8765/'
$OAuthScope       = 'https://graph.microsoft.com/Files.ReadWrite offline_access'

function Get-AccessToken {
    Write-Host '[Auth] Opening Microsoft sign-in in your browser...' -ForegroundColor Cyan

    # Build the authorization URL
    $state   = [System.Guid]::NewGuid().ToString('N')
    $authUrl = "https://login.microsoftonline.com/consumers/oauth2/v2.0/authorize" +
               "?client_id=$($Config.ClientId)" +
               "&response_type=code" +
               "&redirect_uri=$([Uri]::EscapeDataString($OAuthRedirectUri))" +
               "&scope=$([Uri]::EscapeDataString($OAuthScope))" +
               "&state=$state" +
               "&prompt=select_account"

    # Start local HTTP listener BEFORE opening browser
    $listener = [System.Net.HttpListener]::new()
    $listener.Prefixes.Add($OAuthRedirectUri)
    try { $listener.Start() }
    catch {
        Write-Error "[Auth] Cannot start local listener on $OAuthRedirectUri. Port 8765 may be in use."
        return
    }

    # Open browser
    Start-Process $authUrl

    Write-Host '[Auth] Waiting for sign-in to complete in browser...' -ForegroundColor Gray

    # Wait for the redirect (60-second timeout)
    $contextTask = $listener.GetContextAsync()
    $waited      = 0
    while (-not $contextTask.IsCompleted -and $waited -lt 120) {
        Start-Sleep -Milliseconds 500
        $waited++
    }

    $listener.Stop()

    if (-not $contextTask.IsCompleted) {
        Write-Error '[Auth] Timed out waiting for browser sign-in. Please re-run the script.'
        return
    }

    $context  = $contextTask.Result
    $rawUrl   = $context.Request.Url.ToString()

    # Send a friendly close page to the browser
    $html     = '<html><body style="font-family:Segoe UI;text-align:center;margin-top:80px">' +
                '<h2>&#10003; Signed in successfully!</h2>' +
                '<p>You can close this tab and return to the terminal.</p></body></html>'
    $bytes    = [System.Text.Encoding]::UTF8.GetBytes($html)
    $context.Response.ContentType   = 'text/html; charset=utf-8'
    $context.Response.ContentLength64 = $bytes.Length
    $context.Response.OutputStream.Write($bytes, 0, $bytes.Length)
    $context.Response.Close()

    # Parse query string from redirect URL
    $query = [System.Web.HttpUtility]::ParseQueryString(([Uri]$rawUrl).Query)

    if ($query['error']) {
        Write-Error "[Auth] Sign-in error: $($query['error_description'])"
        return
    }

    $returnedState = $query['state']
    if ($returnedState -ne $state) {
        Write-Error '[Auth] State mismatch — possible CSRF. Aborting.'
        return
    }

    $code = $query['code']
    if (-not $code) {
        Write-Error '[Auth] No authorization code returned. Please try again.'
        return
    }

    # Exchange auth code for tokens
    $token = Invoke-RestMethod -Method POST `
        -Uri 'https://login.microsoftonline.com/consumers/oauth2/v2.0/token' `
        -ContentType 'application/x-www-form-urlencoded' `
        -Body @{
            client_id    = $Config.ClientId
            grant_type   = 'authorization_code'
            code         = $code
            redirect_uri = $OAuthRedirectUri
            scope        = $OAuthScope
        }

    $script:AccessToken  = $token.access_token
    $script:RefreshToken = $token.refresh_token
    $script:TokenExpiry  = (Get-Date).AddSeconds([int]$token.expires_in - 60)

    Write-Host '[Auth] Signed in to OneDrive successfully.' -ForegroundColor Green
}

function Refresh-AccessToken {
    Write-Host '[Auth] Refreshing access token...' -ForegroundColor Gray
    $token = Invoke-RestMethod -Method POST `
        -Uri 'https://login.microsoftonline.com/consumers/oauth2/v2.0/token' `
        -ContentType 'application/x-www-form-urlencoded' `
        -Body @{
            client_id     = $Config.ClientId
            grant_type    = 'refresh_token'
            refresh_token = $script:RefreshToken
            scope         = $OAuthScope
        }
    $script:AccessToken  = $token.access_token
    $script:RefreshToken = $token.refresh_token
    $script:TokenExpiry  = (Get-Date).AddSeconds([int]$token.expires_in - 60)
    Write-Host '[Auth] Token refreshed.' -ForegroundColor Gray
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

    # Silently refresh token if close to expiry (no browser popup needed)
    if ((Get-Date) -ge $script:TokenExpiry) {
        if ($script:RefreshToken) { Refresh-AccessToken }
        else                      { Get-AccessToken }
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

# ── Load API key from OneDrive (after auth so we can read OneDrive) ─────────
Write-Host "`n[Settings] Checking for Calendarific API key in OneDrive..." -ForegroundColor Cyan
$apiKey = Load-OneDriveSettings
if (-not $apiKey) {
    Write-Host '[Settings] API key not found in OneDrive — opening dialog...' -ForegroundColor Yellow
    $apiKey = Show-ApiKeyDialog
    Save-OneDriveSettings -ApiKey $apiKey
}
$Config.CalendarificApiKey = $apiKey

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
