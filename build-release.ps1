param(
    [switch]$NoPause,
    [string]$Version = $env:PLENARO_VERSION
)

$ErrorActionPreference = "Stop"

function Write-Step {
    param([string]$Message)
    Write-Host ""
    Write-Host "==> $Message" -ForegroundColor Cyan
}

function Write-Failure {
    param([string]$Message)
    Write-Host ""
    Write-Host "FEHLER: $Message" -ForegroundColor Red
}

function Invoke-DotNetStep {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [string[]]$Arguments
    )

    Write-Step $Name
    & dotnet @Arguments
    $exitCode = $LASTEXITCODE
    if ($exitCode -ne 0) {
        throw "dotnet $($Arguments -join ' ') ist mit Exit-Code $exitCode fehlgeschlagen."
    }
}

try {
    Set-Location $PSScriptRoot

    $projectFile = Join-Path $PSScriptRoot "TaskTool.Wpf.csproj"
    if (-not (Test-Path -LiteralPath $projectFile -PathType Leaf)) {
        throw "Projektdatei wurde nicht gefunden: $projectFile"
    }

    $dotnetCommand = Get-Command dotnet -ErrorAction SilentlyContinue
    if ($null -eq $dotnetCommand) {
        throw "dotnet wurde nicht im PATH gefunden. Bitte installiere das .NET 8 SDK und starte PowerShell danach neu."
    }

    $publishDirectory = Join-Path $PSScriptRoot "artifacts\publish\win-x64"
    if ([string]::IsNullOrWhiteSpace($Version)) {
        $tag = (& git describe --tags --exact-match 2>$null)
        if ($tag -match '^[vV](\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?)$') { $Version = $Matches[1] }
        else { $Version = "2.1.0-dev" }
    }

    Write-Step "Ausgabeordner vorbereiten"
    if (Test-Path -LiteralPath $publishDirectory) {
        Remove-Item -LiteralPath $publishDirectory -Recurse -Force
    }
    New-Item -ItemType Directory -Path $publishDirectory -Force | Out-Null

    Invoke-DotNetStep "Clean" @(
        "clean", $projectFile,
        "-c", "Release"
    )

    Invoke-DotNetStep "Restore" @(
        "restore", $projectFile,
        "-r", "win-x64"
    )

    Invoke-DotNetStep "Publish" @(
        "publish", $projectFile,
        "-c", "Release",
        "-r", "win-x64",
        "--self-contained", "true",
        "-p:PublishSingleFile=true",
        "-p:PublishTrimmed=false",
        "-p:IncludeNativeLibrariesForSelfExtract=true",
        "-p:EnableCompileTimeAppIcon=true",
        "-p:Version=$Version",
        "-p:InformationalVersion=$Version",
        "-o", $publishDirectory
    )

    $exeFiles = Get-ChildItem -LiteralPath $publishDirectory -Filter "*.exe" -File | Sort-Object Length -Descending
    if (-not $exeFiles) {
        throw "Publish abgeschlossen, aber im Ausgabeordner wurde keine EXE gefunden: $publishDirectory"
    }

    $exeFile = $exeFiles[0]
    $sizeMb = [Math]::Round($exeFile.Length / 1MB, 2)

    Write-Host ""
    Write-Host "Release-Build erfolgreich erstellt." -ForegroundColor Green
    Write-Host "EXE: $($exeFile.FullName)" -ForegroundColor Green
    Write-Host "Groesse: $sizeMb MB" -ForegroundColor Green

    if (Get-Command explorer.exe -ErrorAction SilentlyContinue) {
        Start-Process explorer.exe -ArgumentList $publishDirectory
    }

    exit 0
}
catch {
    Write-Failure $_.Exception.Message
    exit 1
}
finally {
    if (-not $NoPause -and -not [Console]::IsInputRedirected) {
        Write-Host ""
        Read-Host "Zum Schliessen Enter druecken"
    }
}
