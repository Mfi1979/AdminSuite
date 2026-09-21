<#
.SYNOPSIS
    28_Export-04-Teams.ps1
    Exportiert alle Teams-Instanzen, Sichtbarkeitseinstellungen und Gruppen-IDs.
#>
[CmdletBinding()]
param([string]$ConfigFile = "28_config.json")

if ($host.Name -eq "Windows PowerShell ISE Host") {
    Write-Warning "Bitte in nativer powershell.exe starten."
    return
}

Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process -Force
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls13

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$baseTenant = Join-Path -Path $scriptDir -ChildPath "TenantData"
$latest = Get-ChildItem -Path $baseTenant -Directory | Sort-Object Name -Descending | Select-Object -First 1
$targetDir = if ($latest) { 
    $latest.FullName 
} else {
    $ts = Get-Date -Format "yyyyMMdd_HHmm"
    $nd = Join-Path -Path $baseTenant -ChildPath $ts
    New-Item -ItemType Directory -Path $nd -Force | Out-Null
    $nd
}

Write-Host "=== [4/4] EXPORT: MICROSOFT TEAMS ===" -ForegroundColor Cyan
Write-Host "Zielverzeichnis: $targetDir" -ForegroundColor Cyan

Import-Module Microsoft.Graph.Authentication -RequiredVersion 2.39.0 -ErrorAction Stop
Connect-MgGraph -Scopes "Group.Read.All" -NoWelcome

Write-Host "Lese Teams-Instanzen aus..." -ForegroundColor Yellow
$teamsGroups = Get-MgGroup -Filter "resourceProvisioningOptions/Any(x:x eq 'Team')" -Property Id, DisplayName, Description, Visibility -All
$teamsReport = @()

if ($teamsGroups) {
    foreach ($tm in $teamsGroups) {
        $teamsReport += [PSCustomObject]@{
            "TeamName"     = [string]$tm.DisplayName
            "Sichtbarkeit" = [string]$tm.Visibility
            "GroupId"      = [string]$tm.Id
            "Beschreibung" = [string]$tm.Description
        }
    }
} else {
    $teamsReport += [PSCustomObject]@{ "Hinweis" = "Keine Microsoft Teams im Tenant vorhanden." }
}

$teamsReport | Export-Csv -Path "$targetDir\10_TeamsInventory.csv" -NoTypeInformation -Encoding UTF8

Write-Host "Teams-Export erfolgreich abgeschlossen!" -ForegroundColor Green