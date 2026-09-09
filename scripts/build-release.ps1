param(
    [string]$Version = "1.1.0",
    [string]$Runtime = "win-x64"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$artifacts = Join-Path $root "artifacts"
$publishDir = Join-Path $artifacts "OfficeLegacyConverter-v$Version-$Runtime"
$zipPath = "$publishDir.zip"

Remove-Item $publishDir -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item $zipPath -Force -ErrorAction SilentlyContinue

dotnet publish (Join-Path $root "OfficeLegacyConverter.csproj") `
    -c Release `
    -r $Runtime `
    --self-contained true `
    -p:Version=$Version `
    -o $publishDir

if ($LASTEXITCODE -ne 0) {
    throw "dotnet publish failed with exit code $LASTEXITCODE"
}

$tesseractSource = Join-Path $env:ProgramFiles "Tesseract-OCR"
if (Test-Path (Join-Path $tesseractSource "tesseract.exe")) {
    $tesseractTarget = Join-Path $publishDir "tools\tesseract"
    New-Item -ItemType Directory -Force -Path $tesseractTarget | Out-Null
    Copy-Item (Join-Path $tesseractSource "tesseract.exe") $tesseractTarget -Force
    Copy-Item (Join-Path $tesseractSource "*.dll") $tesseractTarget -Force
    $license = Join-Path $tesseractSource "doc\LICENSE"
    if (Test-Path $license) {
        Copy-Item $license (Join-Path $tesseractTarget "TESSERACT-LICENSE.txt") -Force
    }
}
else {
    Write-Warning "Tesseract is not installed; OCR remains an optional runtime dependency."
}

$exe = Join-Path $publishDir "OfficeLegacyConverter.exe"
$fileVersion = [Diagnostics.FileVersionInfo]::GetVersionInfo($exe).FileVersion
if (-not $fileVersion.StartsWith($Version)) {
    throw "Release version mismatch: expected $Version, actual $fileVersion"
}

& tar.exe -a -c -f $zipPath -C $publishDir .
if ($LASTEXITCODE -ne 0) {
    throw "Archive creation failed with exit code $LASTEXITCODE"
}
Write-Output "Release: $publishDir"
Write-Output "Archive: $zipPath"
Write-Output "FileVersion: $fileVersion"
