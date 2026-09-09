#Requires -Version 5.1
<#
.SYNOPSIS
    產生 OfficeLegacyConverter 操作標準程序（SOP）PowerPoint 簡報
.DESCRIPTION
    呼叫 docs\SopGenerator C# 專案，透過 PowerPoint COM 自動化建立 SOP 簡報。
.EXAMPLE
    .\Generate-SOP.ps1
    .\Generate-SOP.ps1 -OutputPath "C:\temp\SOP.pptx"
#>
param(
    [string]$OutputPath = (Join-Path $PSScriptRoot "OfficeLegacyConverter_SOP.pptx")
)

$ErrorActionPreference = "Stop"
$generatorDir = Join-Path $PSScriptRoot "SopGenerator"

if (-not (Test-Path $generatorDir)) {
    Write-Error "找不到 SopGenerator 專案目錄：$generatorDir"
    exit 1
}

Push-Location $generatorDir
try {
    dotnet build -c Release
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

    dotnet run -c Release --no-build -- $OutputPath
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

    if (Test-Path $OutputPath) {
        $info = Get-Item $OutputPath
        Write-Host "完成：$($info.FullName) ($([math]::Round($info.Length / 1KB, 1)) KB)"
    }
}
finally {
    Pop-Location
}
