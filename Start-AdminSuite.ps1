#
# ==============================================================================
# ACTIVE DIRECTORY UND ENTRA ID ADMIN SUITE - HYBRID BOOTSTRAPPER
# Startet die Suite lokal aus dem Ordner ODER live per One-Liner aus GitHub:
# irm https://raw.githubusercontent.com/Mfi1979/AdminSuite/main/Start-AdminSuite.ps1 | iex
# ==============================================================================

# GitHub Basis-URL fuer den Web-Abruf
$RepoOwner  = "Mfi1979"
$RepoName   = "AdminSuite"
$RepoBranch = "main"
$BaseRawUrl = "https://raw.githubusercontent.com/$RepoOwner/$RepoName/$RepoBranch"

# Assemblies laden
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.DirectoryServices
[System.Windows.Forms.Application]::EnableVisualStyles()

# Sprachdateien (JSON)
$LanguageFiles = @(
    "Languages/de-DE.json",
    "Languages/en-US.json"
)

# Moduldateien in exakter Abhaengigkeitsreihenfolge
$ModuleFiles = @(
    "Config/UITheme.ps1",
    "Config/I18N.ps1",
    "Core/Helpers.ps1",
    "Core/SystemInfo.ps1",
    "Tools/Tool01_LastLogon.ps1",
    "Tools/Tool02_ADAudit.ps1",
    "Tools/Tool03_EntraStatus.ps1",
    "Tools/Tool04_GroupsGPO.ps1",
    "Tools/Tool05_Win11Check.ps1",
    "Tools/Tool06_DomainOverview.ps1",
    "Tools/Tool07_OSSupportAudit.ps1",
    "Tools/Tool08_ClientSoftware.ps1",
    "Tools/Tool09_ACLCompare.ps1",
    "Tools/Tool10_OUGroupFinder.ps1",
    "Tools/Tool11_PasswordPolicy.ps1",
    "Tools/Tool12_UserPasswordAge.ps1",
    "GUI/MainWindow.ps1"
)

# Pruefen, ob lokal ausgefuehrt oder ueber das Web gestreamt
$IsLocal = ($PSScriptRoot -and (Test-Path "$PSScriptRoot\Config\I18N.ps1"))

if ($IsLocal) {
    Write-Host "Lade Admin Suite lokal aus: $PSScriptRoot" -ForegroundColor Cyan

    # 1. PowerShell-Module laden
    foreach ($file in $ModuleFiles) {
        $localPath = Join-Path $PSScriptRoot ($file -replace '/', '\')
        if (Test-Path $localPath) {
            . $localPath
        } else {
            Write-Warning "Datei nicht gefunden: $localPath"
        }
    }

    # 2. Lokale Sprachdateien einlesen
    foreach ($langFile in $LanguageFiles) {
        $localLangPath = Join-Path $PSScriptRoot ($langFile -replace '/', '\')
        if (Test-Path $localLangPath) {
            try {
                $rawJson = Get-Content -Path $localLangPath -Raw -Encoding UTF8
                if (Get-Command Register-LanguageJson -ErrorAction SilentlyContinue) {
                    Register-LanguageJson -FileName (Split-Path $langFile -Leaf) -JsonContent $rawJson
                }
            } catch {
                Write-Warning "Fehler beim Einlesen von ${localLangPath}: $($_.Exception.Message)"
            }
        }
    }
} else {
    Write-Host "Lade Admin Suite Module von GitHub ($BaseRawUrl)..." -ForegroundColor Cyan
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $webClient = New-Object System.Net.WebClient
    $webClient.Headers.Add("User-Agent", "PowerShell-AdminSuite-Loader")
    $webClient.Encoding = [System.Text.Encoding]::UTF8

    # Cache-Buster fuer Untermodule
    $cacheBuster = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()

    # 1. PowerShell-Module aus dem Web laden und global einbinden
    foreach ($file in $ModuleFiles) {
        $fileUrl = "$BaseRawUrl/$file`?t=$cacheBuster"
        try {
            $code = $webClient.DownloadString($fileUrl)
            . ([ScriptBlock]::Create($code))
        } catch {
            Write-Error "Fehler beim Laden von $fileUrl : $($_.Exception.Message)"
        }
    }

    # 2. Sprachdateien (JSON) aus dem Web laden und registrieren
    foreach ($langFile in $LanguageFiles) {
        $langUrl = "$BaseRawUrl/$langFile`?t=$cacheBuster"
        try {
            $jsonContent = $webClient.DownloadString($langUrl)
            if (Get-Command Register-LanguageJson -ErrorAction SilentlyContinue) {
                Register-LanguageJson -FileName ($langFile.Split('/')[-1]) -JsonContent $jsonContent
            }
        } catch {
            Write-Warning "Fehler beim Laden der Sprachdatei ${langUrl}: $($_.Exception.Message)"
        }
    }
}

# Startet das Haupt-Dashboard
Start-AdminSuiteMainWindow