#requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSCommandPath
$inputScript = Join-Path $root "QuickPDFA.ps1"
$outputDir = Join-Path $root "dist"
$outputExe = Join-Path $outputDir "QuickPDFA.exe"

$logoPreferred = Join-Path $root "assets\LOGO.ico"
$logoFallback = Join-Path $root "assets\ICON.ico"
$logoPath = if (Test-Path -LiteralPath $logoPreferred) {
    $logoPreferred
}
elseif (Test-Path -LiteralPath $logoFallback) {
    $logoFallback
}
else {
    $null
}

if (-not (Test-Path -LiteralPath $inputScript)) {
    throw "Input script not found: $inputScript"
}

if (-not $logoPath) {
    throw "No icon file found. Expected 'assets\\LOGO.ico' (preferred) or 'assets\\ICON.ico' (fallback)."
}

$invokeCommand = Get-Command -Name "Invoke-PS2EXE" -ErrorAction SilentlyContinue
$legacyCommand = Get-Command -Name "ps2exe" -ErrorAction SilentlyContinue

if (-not $invokeCommand -and -not $legacyCommand) {
    Write-Host "ps2exe was not found on this system." -ForegroundColor Yellow
    Write-Host "Install it with:" -ForegroundColor Yellow
    Write-Host "  Install-Module -Name ps2exe -Scope CurrentUser" -ForegroundColor Yellow
    exit 1
}

if (-not (Test-Path -LiteralPath $outputDir)) {
    [void](New-Item -Path $outputDir -ItemType Directory -Force)
}

$baseArgs = @(
    "-inputFile", $inputScript,
    "-outputFile", $outputExe,
    "-iconFile", $logoPath,
    "-title", "QuickPDFA",
    "-product", "QuickPDFA",
    "-company", "Matteo Tognolo",
    "-description", "QuickPDFA",
    "-version", "1.0",
    "-noConsole"
)

$baseParams = @{
    inputFile = $inputScript
    outputFile = $outputExe
    iconFile = $logoPath
    title = "QuickPDFA"
    product = "QuickPDFA"
    company = "Matteo Tognolo"
    description = "QuickPDFA"
    version = "1.0"
    noConsole = $true
}

Write-Host "Building QuickPDFA executable..." -ForegroundColor Cyan
Write-Host "Input : $inputScript"
Write-Host "Output: $outputExe"
Write-Host "Icon  : $logoPath"

if ($invokeCommand) {
    if ($invokeCommand.CommandType -eq "Application") {
        & $invokeCommand.Source @baseArgs
    }
    else {
        & $invokeCommand.Name @baseParams
    }
}
else {
    if ($legacyCommand.CommandType -eq "Application") {
        & $legacyCommand.Source @baseArgs
    }
    else {
        & $legacyCommand.Name @baseParams
    }
}

if (-not (Test-Path -LiteralPath $outputExe)) {
    throw "Build finished without creating output: $outputExe"
}

Write-Host "Build completed successfully." -ForegroundColor Green
Write-Host "Executable: $outputExe" -ForegroundColor Green
