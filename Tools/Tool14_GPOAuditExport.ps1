# ============================================
# GPO Audit - Alle GPOs auf Verlinkung prüfen
# ============================================

Import-Module GroupPolicy

# Ausgabe-Dateien
$outputPath = "C:\_Admin\GPO_Audit"
$csvPath    = Join-Path $outputPath "GPO_Link_Status.csv"
$txtPath    = Join-Path $outputPath "GPO_Link_Status.txt"

# Ausgabeordner erstellen
if (-not (Test-Path $outputPath)) {
    New-Item -ItemType Directory -Path $outputPath -Force | Out-Null
}

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "        GPO LINK AUDIT" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

try {
    # Alle GPOs auslesen
    $gpos = Get-GPO -All -ErrorAction Stop

    Write-Host "Anzahl gefundener GPOs: $($gpos.Count)" -ForegroundColor Green
    Write-Host ""

    $results = foreach ($gpo in $gpos) {

        Write-Host "Prüfe: $($gpo.DisplayName)" -ForegroundColor Gray

        try {
            # GPO-Report als XML abrufen
            $report = [xml](Get-GPOReport `
                -Guid $gpo.Id `
                -ReportType Xml `
                -ErrorAction Stop)

            # Links auslesen
            $links = @($report.GPO.LinksTo)

            if ($links.Count -eq 0) {

                [PSCustomObject]@{
                    GPOName        = $gpo.DisplayName
                    GUID           = $gpo.Id
                    Status         = "NICHT VERLINKT"
                    AnzahlLinks    = 0
                    Verknuepfungen = ""
                }

            }
            else {

                # Alle Link-Ziele zusammenfassen
                $linkInfo = foreach ($link in $links) {

                    $enabled = if ($link.Enabled -eq "true") {
                        "Aktiviert"
                    }
                    else {
                        "Deaktiviert"
                    }

                    "$($link.SOMPath) [$enabled]"
                }

                [PSCustomObject]@{
                    GPOName        = $gpo.DisplayName
                    GUID           = $gpo.Id
                    Status         = "VERLINKT"
                    AnzahlLinks    = $links.Count
                    Verknuepfungen = ($linkInfo -join " | ")
                }
            }
        }
        catch {

            [PSCustomObject]@{
                GPOName        = $gpo.DisplayName
                GUID           = $gpo.Id
                Status         = "FEHLER"
                AnzahlLinks    = 0
                Verknuepfungen = $_.Exception.Message
            }
        }
    }

    # ============================================
    # Ausgabe Konsole
    # ============================================

    Write-Host ""
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host "ERGEBNIS" -ForegroundColor Cyan
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host ""

    $results |
        Sort-Object Status, GPOName |
        Format-Table GPOName, GUID, Status, AnzahlLinks -AutoSize

    # ============================================
    # CSV speichern
    # ============================================

    $results |
        Sort-Object GPOName |
        Export-Csv `
            -Path $csvPath `
            -NoTypeInformation `
            -Encoding UTF8

    # ============================================
    # TXT Report
    # ============================================

    $timestamp = Get-Date -Format "dd.MM.yyyy HH:mm:ss"

    $reportLines = @()

    $reportLines += "============================================================"
    $reportLines += "                 GPO LINK AUDIT"
    $reportLines += "============================================================"
    $reportLines += "Prüfungszeitpunkt : $timestamp"
    $reportLines += "Anzahl GPOs       : $($results.Count)"
    $reportLines += ""

    $linkedCount = @(
        $results | Where-Object { $_.Status -eq "VERLINKT" }
    ).Count

    $unlinkedCount = @(
        $results | Where-Object { $_.Status -eq "NICHT VERLINKT" }
    ).Count

    $errorCount = @(
        $results | Where-Object { $_.Status -eq "FEHLER" }
    ).Count

    $reportLines += "Verlinkte GPOs    : $linkedCount"
    $reportLines += "Nicht verlinkt    : $unlinkedCount"
    $reportLines += "Fehler            : $errorCount"
    $reportLines += ""
    $reportLines += "============================================================"
    $reportLines += ""

    foreach ($item in ($results | Sort-Object Status, GPOName)) {

        $reportLines += "GPO: $($item.GPOName)"
        $reportLines += "GUID: $($item.GUID)"
        $reportLines += "Status: $($item.Status)"
        $reportLines += "Anzahl Links: $($item.AnzahlLinks)"

        if ($item.Verknuepfungen) {
            $reportLines += "Verknüpfungen:"
            $reportLines += "  $($item.Verknuepfungen)"
        }
        else {
            $reportLines += "Verknüpfungen: Keine"
        }

        $reportLines += "------------------------------------------------------------"
    }

    $reportLines |
        Out-File `
            -FilePath $txtPath `
            -Encoding UTF8

    # ============================================
    # Zusammenfassung
    # ============================================

    Write-Host ""
    Write-Host "============================================" -ForegroundColor Green
    Write-Host "AUDIT ABGESCHLOSSEN" -ForegroundColor Green
    Write-Host "============================================" -ForegroundColor Green
    Write-Host ""

    Write-Host "GPOs gesamt       : $($results.Count)"
    Write-Host "Verlinkt          : $linkedCount" -ForegroundColor Green
    Write-Host "Nicht verlinkt    : $unlinkedCount" -ForegroundColor Yellow
    Write-Host "Fehler            : $errorCount" -ForegroundColor Red
    Write-Host ""

    Write-Host "CSV : $csvPath" -ForegroundColor Cyan
    Write-Host "TXT : $txtPath" -ForegroundColor Cyan
    Write-Host ""

    # Nicht verlinkte GPOs zusätzlich anzeigen
    if ($unlinkedCount -gt 0) {

        Write-Host "NICHT VERLINKTE GPOs:" -ForegroundColor Yellow
        Write-Host ""

        $results |
            Where-Object { $_.Status -eq "NICHT VERLINKT" } |
            Sort-Object GPOName |
            Format-Table GPOName, GUID -AutoSize
    }

}
catch {

    Write-Host ""
    Write-Host "FEHLER beim Auslesen der GPOs:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
}
