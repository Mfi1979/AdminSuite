<#
================================================================================
 BUILD-SUITE: BUNDLER-SKRIPT
 Fügt alle Einzeldateien zu einer einzigen eigenständigen 'AdminSuite_Full.ps1' zusammen.
================================================================================
#>

$files = @(
    "Config\UITheme.ps1",
    "Config\I18N.ps1",
    "Core\Helpers.ps1",
    "Core\SystemInfo.ps1",
    "Tools\Tool01_LastLogon.ps1",
    "Tools\Tool02_ADAudit.ps1",
    "Tools\Tool03_EntraStatus.ps1",
    "Tools\Tool04_GroupsGPO.ps1",
    "Tools\Tool05_Win11Check.ps1",
    "Tools\Tool06_DomainOverview.ps1",
    "Tools\Tool07_OSSupportAudit.ps1",
    "Tools\Tool08_ClientSoftware.ps1",
    "Tools\Tool09_ACLCompare.ps1",
    "Tools\Tool10_OUGroupFinder.ps1",
    "Tools\Tool11_PasswordPolicy.ps1",
    "GUI\MainWindow.ps1"
)

$bundle = New-Object System.Text.StringBuilder
[void]$bundle.AppendLine("<#")
[void]$bundle.AppendLine("================================================================================")
[void]$bundle.AppendLine(" ACTIVE DIRECTORY & ENTRA ID ADMIN SUITE (VOLLSTÄNDIGE GESAMTEDITION)")
[void]$bundle.AppendLine(" Automatisch generiert durch Build-Suite.ps1 | Alle Module 1 bis 11 integriert")
[void]$bundle.AppendLine("================================================================================")
[void]$bundle.AppendLine("#>")
[void]$bundle.AppendLine("")
[void]$bundle.AppendLine("Add-Type -AssemblyName System.Windows.Forms")
[void]$bundle.AppendLine("Add-Type -AssemblyName System.Drawing")
[void]$bundle.AppendLine("Add-Type -AssemblyName System.DirectoryServices")
[void]$bundle.AppendLine("[System.Windows.Forms.Application]::EnableVisualStyles()")
[void]$bundle.AppendLine("")

foreach ($f in $files) {
    $fullPath = Join-Path $PSScriptRoot $f
    if (Test-Path $fullPath) {
        $content = Get-Content -Path $fullPath -Raw -Encoding UTF8
        [void]$bundle.AppendLine($content)
        [void]$bundle.AppendLine("")
    } else {
        Write-Warning "Konnte Datei nicht finden: $fullPath"
    }
}

[void]$bundle.AppendLine("Start-AdminSuiteMainWindow")

$outputPath = Join-Path $PSScriptRoot "AdminSuite_Full.ps1"
Set-Content -Path $outputPath -Value $bundle.ToString() -Encoding UTF8
Write-Host "✅ Erfolgreich generiert: $outputPath" -ForegroundColor Green
