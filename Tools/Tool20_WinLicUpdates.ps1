
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ==============================================================================
# ZENTRALE LAYOUT- & DESIGN-KONFIGURATION
# ==============================================================================
$script:UITheme = @{
    HeaderHeight       = 34
    RowHeight          = 28
    HeaderPaddingLeft  = 8
    HeaderPaddingRight = 8
    CellPaddingLeft    = 8
    CellPaddingRight   = 8
    HeaderFont         = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    CellFont           = New-Object System.Drawing.Font("Segoe UI", 9.0, [System.Drawing.FontStyle]::Regular)
    HeaderBackColor    = [System.Drawing.Color]::FromArgb(238, 242, 248)
    HeaderForeColor    = [System.Drawing.Color]::FromArgb(30, 41, 59)
    GridLineColor      = [System.Drawing.Color]::FromArgb(226, 232, 240)
    RowBackColor       = [System.Drawing.Color]::White
    RowAltBackColor    = [System.Drawing.Color]::FromArgb(248, 250, 252)
    SelectionBackColor = [System.Drawing.Color]::FromArgb(203, 228, 249)
    SelectionForeColor = [System.Drawing.Color]::Black
}

function Apply-StandardGridTheme {
    param([System.Windows.Forms.DataGridView]$grid)
    if (-not $grid) { return }

    $grid.EnableHeadersVisualStyles = $false
    $grid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
    $grid.ColumnHeadersHeight = $script:UITheme.HeaderHeight
    $grid.RowTemplate.Height = $script:UITheme.RowHeight

    $hdrStyle = New-Object System.Windows.Forms.DataGridViewCellStyle
    $hdrStyle.Font = $script:UITheme.HeaderFont
    $hdrStyle.BackColor = $script:UITheme.HeaderBackColor
    $hdrStyle.ForeColor = $script:UITheme.HeaderForeColor
    $hdrStyle.Alignment = [System.Drawing.ContentAlignment]::MiddleLeft
    $hdrStyle.Padding = New-Object System.Windows.Forms.Padding($script:UITheme.HeaderPaddingLeft, 0, $script:UITheme.HeaderPaddingRight, 0)
    $grid.ColumnHeadersDefaultCellStyle = $hdrStyle

    $cellStyle = New-Object System.Windows.Forms.DataGridViewCellStyle
    $cellStyle.Font = $script:UITheme.CellFont
    $cellStyle.Alignment = [System.Drawing.ContentAlignment]::MiddleLeft
    $cellStyle.Padding = New-Object System.Windows.Forms.Padding($script:UITheme.CellPaddingLeft, 2, $script:UITheme.CellPaddingRight, 2)
    $cellStyle.SelectionBackColor = $script:UITheme.SelectionBackColor
    $cellStyle.SelectionForeColor = $script:UITheme.SelectionForeColor
    $grid.DefaultCellStyle = $cellStyle

    $grid.AlternatingRowsDefaultCellStyle.BackColor = $script:UITheme.RowAltBackColor
    $grid.AlternatingRowsDefaultCellStyle.SelectionBackColor = $script:UITheme.SelectionBackColor
    $grid.AlternatingRowsDefaultCellStyle.SelectionForeColor = $script:UITheme.SelectionForeColor

    $grid.GridColor = $script:UITheme.GridLineColor
    $grid.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $grid.CellBorderStyle = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
    $grid.BackgroundColor = [System.Drawing.Color]::White
    $grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grid.MultiSelect = $false
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.ReadOnly = $true
    $grid.RowHeadersVisible = $false
}

# ==============================================================================
# HAUPTFENSTER: TOOL 20 (UPDATE-QUELLEN & LIZENZ-DIAGNOSE)
# ==============================================================================
function Show-Tool20UpdateAndActivation {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Tool 20 - Windows Update Quelle & Lizenz-Status ($env:COMPUTERNAME)"
    $form.Size = New-Object System.Drawing.Size(1080, 750)
    $form.MinimumSize = New-Object System.Drawing.Size(920, 600)
    $form.StartPosition = "CenterScreen"
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.BackColor = [System.Drawing.Color]::FromArgb(243, 245, 249)

    # 1. Dashboard-Kopfbereich (Status-Kacheln)
    $pnlHeader = New-Object System.Windows.Forms.Panel
    $pnlHeader.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlHeader.Height = 145
    $pnlHeader.BackColor = [System.Drawing.Color]::FromArgb(255, 255, 255)
    $pnlHeader.Padding = New-Object System.Windows.Forms.Padding(15, 12, 15, 10)
    $form.Controls.Add($pnlHeader)

    # Kachel 1: Update-Quelle
    $cardSource = New-Object System.Windows.Forms.Panel
    $cardSource.Location = New-Object System.Drawing.Point(15, 12)
    $cardSource.Size = New-Object System.Drawing.Size(510, 80)
    $cardSource.BackColor = [System.Drawing.Color]::FromArgb(246, 248, 252)
    $cardSource.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $pnlHeader.Controls.Add($cardSource)

    $lblSourceCaption = New-Object System.Windows.Forms.Label
    $lblSourceCaption.Location = New-Object System.Drawing.Point(10, 8)
    $lblSourceCaption.Size = New-Object System.Drawing.Size(480, 18)
    $lblSourceCaption.Text = "ERMITTELTE UPDATE-VERWALTUNG"
    $lblSourceCaption.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
    $lblSourceCaption.ForeColor = [System.Drawing.Color]::FromArgb(100, 116, 139)
    $cardSource.Controls.Add($lblSourceCaption)

    $lblSourceValue = New-Object System.Windows.Forms.Label
    $lblSourceValue.Location = New-Object System.Drawing.Point(10, 27)
    $lblSourceValue.Size = New-Object System.Drawing.Size(490, 24)
    $lblSourceValue.Text = "Ermittle..."
    $lblSourceValue.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $lblSourceValue.ForeColor = [System.Drawing.Color]::FromArgb(15, 23, 42)
    $cardSource.Controls.Add($lblSourceValue)

    $lblSourceSub = New-Object System.Windows.Forms.Label
    $lblSourceSub.Location = New-Object System.Drawing.Point(10, 53)
    $lblSourceSub.Size = New-Object System.Drawing.Size(490, 20)
    $lblSourceSub.Text = "Prüfe SCCM, WSUS und PolicyManager..."
    $lblSourceSub.ForeColor = [System.Drawing.Color]::FromArgb(71, 85, 105)
    $cardSource.Controls.Add($lblSourceSub)

    # Kachel 2: Windows-Aktivierung & Key
    $cardLic = New-Object System.Windows.Forms.Panel
    $cardLic.Location = New-Object System.Drawing.Point(540, 12)
    $cardLic.Size = New-Object System.Drawing.Size(500, 80)
    $cardLic.BackColor = [System.Drawing.Color]::FromArgb(246, 248, 252)
    $cardLic.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $pnlHeader.Controls.Add($cardLic)

    $lblLicCaption = New-Object System.Windows.Forms.Label
    $lblLicCaption.Location = New-Object System.Drawing.Point(10, 8)
    $lblLicCaption.Size = New-Object System.Drawing.Size(470, 18)
    $lblLicCaption.Text = "WINDOWS AKTIVIERUNG & MAINBOARD OEM-KEY"
    $lblLicCaption.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
    $lblLicCaption.ForeColor = [System.Drawing.Color]::FromArgb(100, 116, 139)
    $cardLic.Controls.Add($lblLicCaption)

    $lblLicStatus = New-Object System.Windows.Forms.Label
    $lblLicStatus.Location = New-Object System.Drawing.Point(10, 27)
    $lblLicStatus.Size = New-Object System.Drawing.Size(470, 24)
    $lblLicStatus.Text = "Lese Lizenzstatus..."
    $lblLicStatus.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $lblLicStatus.ForeColor = [System.Drawing.Color]::FromArgb(22, 101, 52)
    $cardLic.Controls.Add($lblLicStatus)

    $lblLicSub = New-Object System.Windows.Forms.Label
    $lblLicSub.Location = New-Object System.Drawing.Point(10, 53)
    $lblLicSub.Size = New-Object System.Drawing.Size(470, 20)
    $lblLicSub.Text = "Mainboard Key: Abfrage läuft..."
    $lblLicSub.Font = New-Object System.Drawing.Font("Consolas", 8.5)
    $lblLicSub.ForeColor = [System.Drawing.Color]::FromArgb(51, 65, 85)
    $cardLic.Controls.Add($lblLicSub)

    # Aktionsleiste unter den Kacheln
    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Location = New-Object System.Drawing.Point(15, 102)
    $btnRefresh.Size = New-Object System.Drawing.Size(130, 32)
    $btnRefresh.Text = "🔄 Neu einlesen"
    $btnRefresh.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnRefresh.BackColor = [System.Drawing.Color]::FromArgb(240, 244, 248)
    $btnRefresh.FlatStyle = [System.Windows.Forms.FlatStyle]::System
    $pnlHeader.Controls.Add($btnRefresh)

    $btnCopyKey = New-Object System.Windows.Forms.Button
    $btnCopyKey.Location = New-Object System.Drawing.Point(155, 102)
    $btnCopyKey.Size = New-Object System.Drawing.Size(180, 32)
    $btnCopyKey.Text = "📋 Mainboard-Key kopieren"
    $btnCopyKey.FlatStyle = [System.Windows.Forms.FlatStyle]::System
    $pnlHeader.Controls.Add($btnCopyKey)

    $lblFooterHint = New-Object System.Windows.Forms.Label
    $lblFooterHint.Location = New-Object System.Drawing.Point(345, 109)
    $lblFooterHint.Size = New-Object System.Drawing.Size(600, 20)
    $lblFooterHint.ForeColor = [System.Drawing.Color]::FromArgb(100, 116, 139)
    $lblFooterHint.Text = "Status: Bereit. CIM-Abfrage auf UEFI/BIOS und Richtlinien ausgeführt."
    $pnlHeader.Controls.Add($lblFooterHint)

    # 2. Registerkarten (Übersichtstabelle vs. Detail-Log)
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Dock = [System.Windows.Forms.DockStyle]::Fill
    $form.Controls.Add($tabControl)
    $tabControl.BringToFront()

    # Tab 1: Tabelle
    $tabGrid = New-Object System.Windows.Forms.TabPage
    $tabGrid.Text = "Konfigurationsübersicht"
    $tabGrid.Padding = New-Object System.Windows.Forms.Padding(6)
    $tabControl.TabPages.Add($tabGrid)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = [System.Windows.Forms.DockStyle]::Fill
    $tabGrid.Controls.Add($grid)
    Apply-StandardGridTheme -grid $grid

    # Tab 2: Detail-Log
    $tabLog = New-Object System.Windows.Forms.TabPage
    $tabLog.Text = "Technischer Log & Registry-Pfade"
    $tabLog.Padding = New-Object System.Windows.Forms.Padding(6)
    $tabControl.TabPages.Add($tabLog)

    $txtLog = New-Object System.Windows.Forms.TextBox
    $txtLog.Dock = [System.Windows.Forms.DockStyle]::Fill
    $txtLog.Multiline = $true
    $txtLog.ReadOnly = $true
    $txtLog.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
    $txtLog.Font = New-Object System.Drawing.Font("Consolas", 9.0)
    $txtLog.BackColor = [System.Drawing.Color]::FromArgb(253, 254, 255)
    $tabLog.Controls.Add($txtLog)

    # Status-Variable für Key
    $script:ExtractedOemKey = ""

    # 3. Auslese- und Diagnosefunktion
    $analyzeSystem = {
        $lblFooterHint.Text = "Lese Systemwerte, UEFI-Tabellen und Richtlinien ein..."
        $form.Refresh()

        $rawList = [System.Collections.Generic.List[PSCustomObject]]::new()
        $log = New-Object System.Text.StringBuilder

        [void]$log.AppendLine("================================================================================")
        [void]$log.AppendLine(" SYSTEM-DIAGNOSE: WINDOWS UPDATE & LIZENZSTATUS")
        [void]$log.AppendLine(" Computer: $env:COMPUTERNAME | Zeit: $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')")
        [void]$log.AppendLine("================================================================================")
        [void]$log.AppendLine("")

        # ----------------------------------------------------------------------
        # A. MAINBOARD-KEY & AKTIVIERUNG
        # ----------------------------------------------------------------------
        $licService = Get-CimInstance -ClassName SoftwareLicensingService -ErrorAction SilentlyContinue
        $biosKey = if ($licService -and $licService.OA3xOriginalProductKey) { 
            $licService.OA3xOriginalProductKey 
        } else { 
            $null 
        }
        $script:ExtractedOemKey = if ($biosKey) { $biosKey } else { "" }

        $activeProd = Get-CimInstance -ClassName SoftwareLicensingProduct `
            -Filter "ApplicationId = '55c92734-d682-4d71-983e-d6ec3f16059f' and PartialProductKey is not null" `
            -ErrorAction SilentlyContinue | Select-Object -First 1

        $licStatusStr = switch ($activeProd.LicenseStatus) {
            1 { "Aktiviert (Lizenziert)" }
            2 { "OOB-Gnadenfrist (Initial Grace)" }
            3 { "OOT-Gnadenfrist (Additional Grace)" }
            4 { "Nicht-Original (Non-Genuine)" }
            5 { "Nicht aktiviert (Notification Mode)" }
            6 { "Erweiterte Gnadenfrist (Extended Grace)" }
            default { "Unbekannt ($($activeProd.LicenseStatus))" }
        }

        $licChannel = if ($activeProd.Description) {
            if ($activeProd.Description -match "VOLUME_KMSCLIENT") { "KMS Client" }
            elseif ($activeProd.Description -match "VOLUME_MAK")   { "Volume (MAK)" }
            elseif ($activeProd.Description -match "OEM")          { "OEM / Systemhersteller" }
            elseif ($activeProd.Description -match "RETAIL")       { "Retail / Digitale Bindung" }
            else { $activeProd.Description }
        } else { "N/A" }

        # Kachel-Aktualisierung Aktivierung
        $lblLicStatus.Text = $licStatusStr
        if ($activeProd.LicenseStatus -eq 1) {
            $lblLicStatus.ForeColor = [System.Drawing.Color]::FromArgb(22, 101, 52)
        } else {
            $lblLicStatus.ForeColor = [System.Drawing.Color]::FromArgb(185, 28, 28)
        }

        $lblLicSub.Text = if ($biosKey) { "Mainboard-Key: $biosKey" } else { "Mainboard-Key: Nicht im UEFI vorhanden (Retail/KMS)" }

        # Einträge für Grid
        $rawList.Add([PSCustomObject]@{ "Kategorie" = "Aktivierung"; "Komponente / Setting" = "Windows Lizenzstatus"; "Aktiver Wert" = $licStatusStr; "Quelle / Registry-Pfad" = "CIM: SoftwareLicensingProduct" })
        $rawList.Add([PSCustomObject]@{ "Kategorie" = "Aktivierung"; "Komponente / Setting" = "Lizenzkanal"; "Aktiver Wert" = $licChannel; "Quelle / Registry-Pfad" = "CIM: Description" })
        $rawList.Add([PSCustomObject]@{ "Kategorie" = "Aktivierung"; "Komponente / Setting" = "Aktiver Key (Teilstring)"; "Aktiver Wert" = if ($activeProd.PartialProductKey) { "XXXXX-XXXXX-XXXXX-XXXXX-$($activeProd.PartialProductKey)" } else { "N/A" }; "Quelle / Registry-Pfad" = "CIM: PartialProductKey" })
        $rawList.Add([PSCustomObject]@{ "Kategorie" = "Aktivierung"; "Komponente / Setting" = "Mainboard-Key (MSDM-Tabelle)"; "Aktiver Wert" = if ($biosKey) { $biosKey } else { "Nicht im BIOS/UEFI hinterlegt" }; "Quelle / Registry-Pfad" = "ACPI: MSDM Table (OA3x)" })

        # ----------------------------------------------------------------------
        # B. WINDOWS UPDATE QUELLEN-ANALYSE
        # ----------------------------------------------------------------------
        $wuRegPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
        $auRegPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"
        $wuProps   = Get-ItemProperty -Path $wuRegPath -ErrorAction SilentlyContinue
        $auProps   = Get-ItemProperty -Path $auRegPath -ErrorAction SilentlyContinue

        $wsusServer = if ($wuProps.WUServer) { $wuProps.WUServer } else { $null }
        $wsusStatus = if ($wuProps.WUStatusServer) { $wuProps.WUStatusServer } else { $null }
        $useWUServer = if ($auProps.UseWUServer -ne $null) { [int]$auProps.UseWUServer } else { $null }
        $targetGroup = if ($wuProps.TargetGroup) { $wuProps.TargetGroup } else { "-" }

        $ccmService = Get-Service -Name "CcmExec" -ErrorAction SilentlyContinue
        $hasCcm = ($ccmService -ne $null)

        $intunePath  = "HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\Update"
        $intuneProps = Get-ItemProperty -Path $intunePath -ErrorAction SilentlyContinue
        $hasIntune   = $false
        if ($intuneProps) {
            if ($intuneProps.QualityUpdateDeferralInDays -ne $null -or 
                $intuneProps.FeatureUpdateDeferralInDays -ne $null -or 
                $intuneProps.ConfigureDeadlineForQualityUpdates -ne $null) {
                $hasIntune = $true
            }
        }

        # Primäre Quelle klassifizieren
        $sourceTitle = ""
        $sourceSub   = ""
        $sourceColor = [System.Drawing.Color]::Navy

        if ($hasCcm -and $wsusServer -and $useWUServer -eq 1) {
            $sourceTitle = "Microsoft Endpoint Configuration Manager (SCCM)"
            $sourceSub   = "Updates über SCCM-Client ($($ccmService.Status)) & lokalen SUP-Server"
            $sourceColor = [System.Drawing.Color]::FromArgb(30, 58, 138)
        } elseif ($wsusServer -and $useWUServer -eq 1) {
            $sourceTitle = "Lokaler WSUS Server (Windows Server Update Services)"
            $sourceSub   = "WSUS-Server: $wsusServer"
            $sourceColor = [System.Drawing.Color]::FromArgb(67, 56, 202)
        } elseif ($hasIntune) {
            $sourceTitle = "Microsoft Intune Update Rings (WUfB)"
            $sourceSub   = "Verwaltet über Cloud MDM-Richtlinien (Windows Update for Business)"
            $sourceColor = [System.Drawing.Color]::FromArgb(21, 128, 61)
        } elseif ($useWUServer -eq 0 -and $wsusServer) {
            $sourceTitle = "Dual-Scan / Microsoft CDN (WSUS inaktiv)"
            $sourceSub   = "UseWUServer ist 0: Richtlinien greifen direkt auf Microsoft CDN zurück"
            $sourceColor = [System.Drawing.Color]::FromArgb(180, 83, 9)
        } else {
            $sourceTitle = "Windows Update (Standard / Unverwaltet)"
            $sourceSub   = "Keine Gruppenrichtlinien oder Intune-Ringe konfiguriert (Direct CDN)"
            $sourceColor = [System.Drawing.Color]::FromArgb(71, 85, 105)
        }

        $lblSourceValue.Text = $sourceTitle
        $lblSourceValue.ForeColor = $sourceColor
        $lblSourceSub.Text = $sourceSub

        # Update-Werte ins Grid füllen
        $rawList.Add([PSCustomObject]@{ "Kategorie" = "WSUS / GPO"; "Komponente / Setting" = "WUServer (Update-Quelle)"; "Aktiver Wert" = if ($wsusServer) { $wsusServer } else { "Nicht konfiguriert" }; "Quelle / Registry-Pfad" = $wuRegPath })
        $rawList.Add([PSCustomObject]@{ "Kategorie" = "WSUS / GPO"; "Komponente / Setting" = "WUStatusServer (Reporting)"; "Aktiver Wert" = if ($wsusStatus) { $wsusStatus } else { "Nicht konfiguriert" }; "Quelle / Registry-Pfad" = $wuRegPath })
        $rawList.Add([PSCustomObject]@{ "Kategorie" = "WSUS / GPO"; "Komponente / Setting" = "UseWUServer"; "Aktiver Wert" = if ($useWUServer -ne $null) { "$useWUServer (1 = WSUS aktiv, 0 = Ignorieren)" } else { "Nicht definiert" }; "Quelle / Registry-Pfad" = $auRegPath })
        $rawList.Add([PSCustomObject]@{ "Kategorie" = "WSUS / GPO"; "Komponente / Setting" = "TargetGroup (Zielgruppe)"; "Aktiver Wert" = $targetGroup; "Quelle / Registry-Pfad" = $wuRegPath })

        $rawList.Add([PSCustomObject]@{ "Kategorie" = "ConfigMgr / SCCM"; "Komponente / Setting" = "CcmExec Client Dienst"; "Aktiver Wert" = if ($hasCcm) { "Installiert ($($ccmService.Status))" } else { "Nicht installiert" }; "Quelle / Registry-Pfad" = "Service: CcmExec" })

        if ($intuneProps) {
            $rawList.Add([PSCustomObject]@{ "Kategorie" = "Intune Ring"; "Komponente / Setting" = "QualityUpdateDeferralInDays"; "Aktiver Wert" = if ($intuneProps.QualityUpdateDeferralInDays -ne $null) { "$($intuneProps.QualityUpdateDeferralInDays) Tage Verzögerung" } else { "Nicht konfiguriert" }; "Quelle / Registry-Pfad" = $intunePath })
            $rawList.Add([PSCustomObject]@{ "Kategorie" = "Intune Ring"; "Komponente / Setting" = "FeatureUpdateDeferralInDays"; "Aktiver Wert" = if ($intuneProps.FeatureUpdateDeferralInDays -ne $null) { "$($intuneProps.FeatureUpdateDeferralInDays) Tage Verzögerung" } else { "Nicht konfiguriert" }; "Quelle / Registry-Pfad" = $intunePath })
            $rawList.Add([PSCustomObject]@{ "Kategorie" = "Intune Ring"; "Komponente / Setting" = "ConfigureDeadlineForQualityUpdates"; "Aktiver Wert" = if ($intuneProps.ConfigureDeadlineForQualityUpdates -ne $null) { "$($intuneProps.ConfigureDeadlineForQualityUpdates) Tage Frist" } else { "Nicht konfiguriert" }; "Quelle / Registry-Pfad" = $intunePath })
            $rawList.Add([PSCustomObject]@{ "Kategorie" = "Intune Ring"; "Komponente / Setting" = "UpdateServiceUrlAlternate"; "Aktiver Wert" = if ($intuneProps.UpdateServiceUrlAlternate) { $intuneProps.UpdateServiceUrlAlternate } else { "Nicht definiert" }; "Quelle / Registry-Pfad" = $intunePath })
        } else {
            $rawList.Add([PSCustomObject]@{ "Kategorie" = "Intune Ring"; "Komponente / Setting" = "MDM Policy Status"; "Aktiver Wert" = "Keine PolicyManager Update-Ringe vorhanden"; "Quelle / Registry-Pfad" = $intunePath })
        }

        # Grid füllen & Spaltenbreiten anpassen
        $arrG = [System.Collections.ArrayList]::new()
        foreach ($item in $rawList) { [void]$arrG.Add($item) }
        $grid.DataSource = $arrG
        $grid.Columns["Kategorie"].Width = 140
        $grid.Columns["Komponente / Setting"].Width = 260
        $grid.Columns["Aktiver Wert"].Width = 320
        $grid.Columns["Quelle / Registry-Pfad"].AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill

        # Log aufbauen
        [void]$log.AppendLine("[1] AKTIVIERUNG & LIZENZDETAILS:")
        [void]$log.AppendLine("  Aktivierungsstatus:     $licStatusStr")
        [void]$log.AppendLine("  Lizenzkanal:            $licChannel")
        [void]$log.AppendLine("  Teilschlüssel (Aktiv):  $($activeProd.PartialProductKey)")
        [void]$log.AppendLine("  UEFI/BIOS OEM-Key:      $(if ($biosKey) { $biosKey } else { 'Kein Key in MSDM-Tabelle' })")
        [void]$log.AppendLine("")
        [void]$log.AppendLine("[2] UPDATE-QUELLEN DETAILPROTOKOLL:")
        [void]$log.AppendLine("  Primäre Zuordnung:      $sourceTitle")
        [void]$log.AppendLine("  WSUS WUServer:          $wsusServer")
        [void]$log.AppendLine("  WSUS WUStatusServer:    $wsusStatus")
        [void]$log.AppendLine("  UseWUServer Flag:       $useWUServer")
        [void]$log.AppendLine("  SCCM CcmExec:           $(if ($hasCcm) { 'Installiert (' + $ccmService.Status + ')' } else { 'Nicht vorhanden' })")
        [void]$log.AppendLine("")
        [void]$log.AppendLine("[3] REGISTRY DUMP ($wuRegPath):")
        if ($wuProps) {
            $wuProps.PSObject.Properties | Where-Object { $_.Name -notmatch '^__' } | ForEach-Object {
                [void]$log.AppendLine("  $($_.Name) = $($_.Value)")
            }
        } else {
            [void]$log.AppendLine("  (Schlüssel nicht angelegt)")
        }
        [void]$log.AppendLine("")
        [void]$log.AppendLine("[4] INTUNE POLICYMANAGER DUMP ($intunePath):")
        if ($intuneProps) {
            $intuneProps.PSObject.Properties | Where-Object { $_.Name -notmatch '^__' } | ForEach-Object {
                [void]$log.AppendLine("  $($_.Name) = $($_.Value)")
            }
        } else {
            [void]$log.AppendLine("  (Keine Intune-Richtlinien im PolicyManager gefunden)")
        }

        $txtLog.Text = $log.ToString()
        $lblFooterHint.Text = "Status: Abfrage erfolgreich abgeschlossen ($(Get-Date -Format 'HH:mm:ss'))."
    }

    # Event-Verdrahtung
    $btnRefresh.Add_Click({ & $analyzeSystem })
    $btnCopyKey.Add_Click({
        if ($script:ExtractedOemKey) {
            [System.Windows.Forms.Clipboard]::SetText($script:ExtractedOemKey)
            [System.Windows.Forms.MessageBox]::Show("Mainboard-Key in die Zwischenablage kopiert:`n`n$($script:ExtractedOemKey)", "Key kopiert", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } else {
            [System.Windows.Forms.MessageBox]::Show("Auf diesem Mainboard ist kein OEM-Key in der ACPI-MSDM-Tabelle hinterlegt.", "Kein Key vorhanden", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    })

    $form.Add_Shown({ & $analyzeSystem })
    [void]$form.ShowDialog()
}

# Direktes Starten:
Show-Tool20UpdateAndActivation
