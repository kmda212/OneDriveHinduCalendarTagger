<#
.SYNOPSIS
    Lists all Hindu festivals from Calendarific and lets you pick which ones to track.

.DESCRIPTION
    Fetches religious holidays for India (country=IN) for the current year,
    displays them in a checkbox picker, and saves your selection to
    festivals_config.json in the same folder. FestivalAlbums.ps1 reads this
    file automatically, overriding its built-in FestivalsToTrack list.

.PARAMETER ApiKey
    Your Calendarific API key.

.EXAMPLE
    .\Get-HinduFestivals.ps1 -ApiKey "your-api-key-here"
#>

param(
    [Parameter(Mandatory)]
    [string]$ApiKey
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ScriptDir      = Split-Path -Parent $MyInvocation.MyCommand.Path
$ConfigFile     = Join-Path $ScriptDir 'festivals_config.json'
$year           = (Get-Date).Year

# ── Fetch from Calendarific ───────────────────────────────────────────────────

Write-Host "`nFetching Hindu/religious festivals for India ($year)..." -ForegroundColor Cyan

try {
    $response = Invoke-RestMethod `
        -Uri "https://calendarific.com/api/v2/holidays?api_key=$ApiKey&country=IN&year=$year&type=religious" `
        -Method GET -ErrorAction Stop
} catch {
    Write-Host "[Error] API call failed: $_" -ForegroundColor Red
    Write-Host "Check your API key and free tier quota." -ForegroundColor Yellow
    exit 1
}

$holidays = $response.response.holidays
if (-not $holidays -or $holidays.Count -eq 0) {
    Write-Host "[Error] No holidays returned. Check your API key." -ForegroundColor Red
    exit 1
}

# Deduplicate by name, sort alphabetically
$allFestivals = $holidays |
    Group-Object name |
    ForEach-Object { $_.Group[0] } |
    Sort-Object name |
    ForEach-Object { $_.name }

Write-Host "Found $($allFestivals.Count) distinct festivals." -ForegroundColor Green

# ── Load any previously saved selection ──────────────────────────────────────

$previousSelection = @()
if (Test-Path $ConfigFile) {
    try {
        $prev = Get-Content $ConfigFile -Raw | ConvertFrom-Json
        $previousSelection = @($prev.FestivalsToTrack)
        Write-Host "Loaded previous selection ($($previousSelection.Count) festivals)." -ForegroundColor Gray
    } catch {}
}

# ── Checkbox Picker (WinForms) ────────────────────────────────────────────────

$form = [System.Windows.Forms.Form]@{
    Text            = 'Select Festivals to Track'
    Size            = [System.Drawing.Size]::new(480, 560)
    StartPosition   = 'CenterScreen'
    FormBorderStyle = 'FixedDialog'
    MaximizeBox     = $false
    MinimizeBox     = $false
    BackColor       = [System.Drawing.Color]::WhiteSmoke
}

$header = [System.Windows.Forms.Label]@{
    Text      = '🪔  Choose Festivals to Track'
    Location  = [System.Drawing.Point]::new(16, 12)
    Size      = [System.Drawing.Size]::new(440, 28)
    Font      = [System.Drawing.Font]::new('Segoe UI', 13, [System.Drawing.FontStyle]::Bold)
    ForeColor = [System.Drawing.Color]::DarkSlateBlue
}

$sub = [System.Windows.Forms.Label]@{
    Text      = "Check all festivals you want to create albums for. Names are saved exactly as shown."
    Location  = [System.Drawing.Point]::new(16, 44)
    Size      = [System.Drawing.Size]::new(440, 36)
    Font      = [System.Drawing.Font]::new('Segoe UI', 9)
    ForeColor = [System.Drawing.Color]::DimGray
}

$clb = [System.Windows.Forms.CheckedListBox]@{
    Location      = [System.Drawing.Point]::new(16, 86)
    Size          = [System.Drawing.Size]::new(434, 370)
    Font          = [System.Drawing.Font]::new('Segoe UI', 10)
    CheckOnClick  = $true
    BorderStyle   = 'FixedSingle'
    BackColor     = [System.Drawing.Color]::White
}

foreach ($festival in $allFestivals) {
    $checked = $festival -in $previousSelection
    $clb.Items.Add($festival, $checked) | Out-Null
}

$btnAll = [System.Windows.Forms.Button]@{
    Text      = 'Select All'
    Location  = [System.Drawing.Point]::new(16, 466)
    Size      = [System.Drawing.Size]::new(90, 30)
    FlatStyle = 'Flat'
    Font      = [System.Drawing.Font]::new('Segoe UI', 9)
}
$btnAll.add_Click({
    for ($i = 0; $i -lt $clb.Items.Count; $i++) { $clb.SetItemChecked($i, $true) }
})

$btnNone = [System.Windows.Forms.Button]@{
    Text      = 'Clear All'
    Location  = [System.Drawing.Point]::new(114, 466)
    Size      = [System.Drawing.Size]::new(80, 30)
    FlatStyle = 'Flat'
    Font      = [System.Drawing.Font]::new('Segoe UI', 9)
}
$btnNone.add_Click({
    for ($i = 0; $i -lt $clb.Items.Count; $i++) { $clb.SetItemChecked($i, $false) }
})

$btnSave = [System.Windows.Forms.Button]@{
    Text         = 'Save Selection'
    Location     = [System.Drawing.Point]::new(320, 462)
    Size         = [System.Drawing.Size]::new(130, 36)
    BackColor    = [System.Drawing.Color]::DarkSlateBlue
    ForeColor    = [System.Drawing.Color]::White
    FlatStyle    = 'Flat'
    Font         = [System.Drawing.Font]::new('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    DialogResult = 'OK'
}

$btnCancel = [System.Windows.Forms.Button]@{
    Text         = 'Cancel'
    Location     = [System.Drawing.Point]::new(230, 462)
    Size         = [System.Drawing.Size]::new(80, 36)
    FlatStyle    = 'Flat'
    Font         = [System.Drawing.Font]::new('Segoe UI', 10)
    DialogResult = 'Cancel'
}

$form.AcceptButton = $btnSave
$form.CancelButton = $btnCancel
$form.Controls.AddRange(@($header, $sub, $clb, $btnAll, $btnNone, $btnSave, $btnCancel))

$result = $form.ShowDialog()

if ($result -ne 'OK') {
    Write-Host 'Cancelled — no changes saved.' -ForegroundColor Yellow
    exit 0
}

$selected = @($clb.CheckedItems)

if ($selected.Count -eq 0) {
    Write-Host 'No festivals selected — nothing saved.' -ForegroundColor Yellow
    exit 0
}

# ── Save to festivals_config.json ─────────────────────────────────────────────

$configData = @{
    FestivalsToTrack = $selected
    generated_on     = (Get-Date -Format 'o')
    source_year      = $year
}

$configData | ConvertTo-Json -Depth 5 | Set-Content -Path $ConfigFile -Encoding UTF8

# ── Summary ───────────────────────────────────────────────────────────────────

$bar = '═' * 50
Write-Host "`n$bar" -ForegroundColor Green
Write-Host "  Saved $($selected.Count) festivals to:" -ForegroundColor Green
Write-Host "  $ConfigFile" -ForegroundColor Cyan
Write-Host "$bar" -ForegroundColor Green
Write-Host ''
$selected | ForEach-Object { Write-Host "  ✔ $_" -ForegroundColor White }
Write-Host ''
Write-Host 'FestivalAlbums.ps1 will use this file automatically on next run.' -ForegroundColor Yellow

