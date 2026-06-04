<#
.SYNOPSIS
    Batch export Neware .ndax files to Excel (.xlsx)
    Uses BTSDAExReport.exe "export custom" - identical to manual export (8 sheets)

.PARAMETER InputDir
    Folder containing .ndax files (default: script directory)

.PARAMETER OutputDir
    Destination for exported .xlsx files (default: InputDir\exported)

.PARAMETER Recurse
    Search subdirectories recursively

.PARAMETER Force
    Overwrite existing .xlsx files

.EXAMPLE
    .\neware_batch_export.ps1 -InputDir "E:\data\ndax" -OutputDir "E:\data\xlsx"
    .\neware_batch_export.ps1 -InputDir "E:\data" -Recurse -Force
#>

param(
    [string]$InputDir  = $PSScriptRoot,
    [string]$OutputDir = "",
    [switch]$Recurse,
    [switch]$Force
)

$BTSDA_EXE   = "E:\software\BTSClient80\BTSDAExReport.exe"
$EXPORT_TYPE = "custom"   # custom = 8-sheet full export, same as manual

# ---- Validate ----------------------------------------------------------
if (-not (Test-Path $BTSDA_EXE)) {
    Write-Error "BTSDAExReport.exe not found: $BTSDA_EXE"; exit 1
}
if (-not (Test-Path $InputDir)) {
    Write-Error "Input directory not found: $InputDir"; exit 1
}
if (-not $OutputDir) {
    $OutputDir = Join-Path $InputDir "exported"
}
if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
    Write-Host "Created output directory: $OutputDir"
}

# ---- Collect files -----------------------------------------------------
$searchOption = if ($Recurse) { "AllDirectories" } else { "TopDirectoryOnly" }
$ndaxFiles = [System.IO.Directory]::GetFiles($InputDir, "*.ndax", $searchOption)

if ($ndaxFiles.Count -eq 0) {
    Write-Host "No .ndax files found in: $InputDir"; exit 0
}
Write-Host "Found $($ndaxFiles.Count) .ndax file(s)"

# ---- Batch export ------------------------------------------------------
$success = 0
$skipped = 0
$failed  = 0

foreach ($ndax in $ndaxFiles) {
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($ndax)
    $xlsxPath = Join-Path $OutputDir "$baseName.xlsx"

    if ((Test-Path $xlsxPath) -and -not $Force) {
        Write-Host "  [SKIP]  $baseName.xlsx (use -Force to overwrite)"
        $skipped++
        continue
    }

    Write-Host "  [EXPORT] $([System.IO.Path]::GetFileName($ndax)) ..."
    $proc = Start-Process -FilePath $BTSDA_EXE `
        -ArgumentList "export", $EXPORT_TYPE, "`"$ndax`"", "`"$xlsxPath`"" `
        -PassThru -Wait -WindowStyle Hidden

    if ($proc.ExitCode -eq 0 -and (Test-Path $xlsxPath)) {
        $sizeMB = [math]::Round((Get-Item $xlsxPath).Length / 1MB, 2)
        Write-Host "  [OK]    $baseName.xlsx ($sizeMB MB)"
        $success++
    } else {
        Write-Host "  [FAIL]  $baseName  exit=$($proc.ExitCode)" -ForegroundColor Red
        $failed++
    }
}

# ---- Summary -----------------------------------------------------------
Write-Host ""
Write-Host "Done: $success exported, $skipped skipped, $failed failed"
