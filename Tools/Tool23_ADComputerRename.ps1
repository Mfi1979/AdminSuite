<#
.SYNOPSIS
    Tool 23: AD Remote Computer Rename & AD/Entra Inspector Suite
.DESCRIPTION
    - Tab 1: Remote-Umbenennung von Windows-Clients ohne WinRM via RPC/SMB inkl. CSV-Steuerung.
    - Tab 2: AD-Konto Inspector zur Überprüfung & Korrektur von DisplayName und Description.
    - Nahtloser Übergang: Rechner aus der Umbenennungs-CSV können im Inspector direkt ausgewählt werden.
    - Buttons zum direkten Öffnen von Quell-CSV, Logdateien und Fehlerberichten.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Import-Module ActiveDirectory -ErrorAction SilentlyContinue

if (-not [System.Windows.Forms.Application]::RenderWithVisualStyles) {
    try { [System.Windows.Forms.Application]::EnableVisualStyles() } catch {}
}

# -----------------------------------------------------------------------------
# POPUP: BESTÄTIGUNG MIT TABELLE VOR DEM BATCH-RENAME
# -----------------------------------------------------------------------------
function Show-RenameConfirmDialog {
    param([array]$ComputersList)

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Umbenennung bestätigen - Vorab-Prüfung"
    $dlg.Size = New-Object System.Drawing.Size(650, 480)
    $dlg.MinimumSize = New-Object System.Drawing.Size(550, 380)
    $dlg.StartPosition = "CenterParent"
    $dlg.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
    $dlg.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

    $pnlTop = New-Object System.Windows.Forms.Panel
    $pnlTop.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlTop.Height = 75
    $pnlTop.Padding = New-Object System.Windows.Forms.Padding(15, 10, 15, 5)
    $dlg.Controls.Add($pnlTop)

    $lblCount = New-Object System.Windows.Forms.Label
    $lblCount.Text = "Es wurden $($ComputersList.Count) Computer in der CSV gefunden."
    $lblCount.Dock = [System.Windows.Forms.DockStyle]::Top
    $lblCount.Height = 24
    $lblCount.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $lblCount.ForeColor = [System.Drawing.Color]::FromArgb(30, 41, 59)
    $pnlTop.Controls.Add($lblCount)

    $lblWarn = New-Object System.Windows.Forms.Label
    $lblWarn.Text = "Achtung: Erreichbare Clients werden nach erfolgreichem Rename-Befehl sofort neu gestartet (-Force -Restart)!"
    $lblWarn.Dock = [System.Windows.Forms.DockStyle]::Fill
    $lblWarn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
    $lblWarn.ForeColor = [System.Drawing.Color]::FromArgb(180, 40, 40)
    $pnlTop.Controls.Add($lblWarn)

    $pnlBottom = New-Object System.Windows.Forms.Panel
    $pnlBottom.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $pnlBottom.Height = 55
    $pnlBottom.BackColor = [System.Drawing.Color]::FromArgb(238, 242, 246)
    $dlg.Controls.Add($pnlBottom)

    $btnConfirm = New-Object System.Windows.Forms.Button
    $btnConfirm.Text = "Ja, Umbenennung starten"
    $btnConfirm.Location = New-Object System.Drawing.Point(260, 10)
    $btnConfirm.Size = New-Object System.Drawing.Size(210, 34)
    $btnConfirm.BackColor = [System.Drawing.Color]::FromArgb(190, 40, 40)
    $btnConfirm.ForeColor = [System.Drawing.Color]::White
    $btnConfirm.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnConfirm.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    $btnConfirm.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $pnlBottom.Controls.Add($btnConfirm)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Abbrechen"
    $btnCancel.Location = New-Object System.Drawing.Point(485, 10)
    $btnCancel.Size = New-Object System.Drawing.Size(130, 34)
    $btnCancel.BackColor = [System.Drawing.Color]::FromArgb(220, 225, 230)
    $btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::No
    $pnlBottom.Controls.Add($btnCancel)

    $dlg.AcceptButton = $btnConfirm
    $dlg.CancelButton = $btnCancel

    $grpGrid = New-Object System.Windows.Forms.GroupBox
    $grpGrid.Text = "Geplante Änderungen (Voransicht)"
    $grpGrid.Dock = [System.Windows.Forms.DockStyle]::Fill
    $grpGrid.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $dlg.Controls.Add($grpGrid)
    $grpGrid.BringToFront()

    $gridPreview = New-Object System.Windows.Forms.DataGridView
    $gridPreview.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridPreview.BackgroundColor = [System.Drawing.Color]::White
    $gridPreview.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $gridPreview.CellBorderStyle = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
    $gridPreview.ReadOnly = $true
    $gridPreview.AllowUserToAddRows = $false
    $gridPreview.AllowUserToDeleteRows = $false
    $gridPreview.RowHeadersVisible = $false
    $gridPreview.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $gridPreview.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill

    $gridPreview.EnableHeadersVisualStyles = $false
    $gridPreview.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
    $gridPreview.ColumnHeadersHeight = 34
    $gridPreview.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    $gridPreview.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(235, 239, 245)
    $gridPreview.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(35, 45, 55)
    $gridPreview.ColumnHeadersDefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleLeft
    $gridPreview.ColumnHeadersDefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(8, 2, 8, 2)

    $gridPreview.RowTemplate.Height = 28
    $gridPreview.DefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9.0, [System.Drawing.FontStyle]::Regular)
    $gridPreview.DefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(8, 2, 8, 2)
    $grpGrid.Controls.Add($gridPreview)

    [void]$gridPreview.Columns.Add("OldName", "Alter Name")
    [void]$gridPreview.Columns.Add("NewName", "Neuer Name")

    foreach ($entry in $ComputersList) {
        [void]$gridPreview.Rows.Add($entry.AlterName.Trim(), $entry.NeuerName.Trim())
    }

    $diagResult = $dlg.ShowDialog()
    $dlg.Dispose()
    return ($diagResult -eq [System.Windows.Forms.DialogResult]::Yes)
}

# -----------------------------------------------------------------------------
# POPUP: ABSCHLUSS-DIALOG MIT LOG-ZUGRIFF
# -----------------------------------------------------------------------------
function Show-ExecutionFinishedDialog {
    param(
        [int]$FailedCount,
        [string]$FailCsvPath,
        [string]$LogTxtPath
    )

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = if ($FailedCount -gt 0) { "Abgeschlossen mit Fehlern" } else { "Erfolgreich abgeschlossen" }
    $dlg.Size = New-Object System.Drawing.Size(720, 270)
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
    $dlg.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 252)

    $pnlMsg = New-Object System.Windows.Forms.Panel
    $pnlMsg.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlMsg.Height = 140
    $pnlMsg.Padding = New-Object System.Windows.Forms.Padding(20, 15, 20, 10)
    $dlg.Controls.Add($pnlMsg)

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Dock = [System.Windows.Forms.DockStyle]::Top
    $lblTitle.Height = 30
    $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 11.0, [System.Drawing.FontStyle]::Bold)
    if ($FailedCount -gt 0) {
        $lblTitle.Text = "Der Vorgang wurde mit $FailedCount Warnung(en) / Fehler(n) beendet."
        $lblTitle.ForeColor = [System.Drawing.Color]::FromArgb(180, 40, 40)
    } else {
        $lblTitle.Text = "Alle Computer wurden erfolgreich umbenannt und im AD bereinigt!"
        $lblTitle.ForeColor = [System.Drawing.Color]::FromArgb(20, 120, 40)
    }
    $pnlMsg.Controls.Add($lblTitle)

    $txtInfo = New-Object System.Windows.Forms.TextBox
    $txtInfo.Dock = [System.Windows.Forms.DockStyle]::Fill
    $txtInfo.Multiline = $true
    $txtInfo.ReadOnly = $true
    $txtInfo.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 252)
    $txtInfo.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $txtInfo.Font = New-Object System.Drawing.Font("Segoe UI", 9.0)

    $infoLines = @()
    if ($FailedCount -gt 0 -and (Test-Path $FailCsvPath)) {
        $infoLines += "Fehlerbericht CSV:`r`n$FailCsvPath"
    }
    if (Test-Path $LogTxtPath) {
        $infoLines += "Detail-Log TXT:`r`n$LogTxtPath"
    }
    $txtInfo.Text = $infoLines -join "`r`n`r`n"
    $pnlMsg.Controls.Add($txtInfo)
    $lblTitle.BringToFront()

    $pnlBtns = New-Object System.Windows.Forms.FlowLayoutPanel
    $pnlBtns.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $pnlBtns.Height = 58
    $pnlBtns.FlowDirection = [System.Windows.Forms.FlowDirection]::RightToLeft
    $pnlBtns.Padding = New-Object System.Windows.Forms.Padding(10, 10, 15, 10)
    $pnlBtns.BackColor = [System.Drawing.Color]::FromArgb(235, 240, 246)
    $dlg.Controls.Add($pnlBtns)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "Schließen"
    $btnClose.Size = New-Object System.Drawing.Size(110, 34)
    $btnClose.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $pnlBtns.Controls.Add($btnClose)

    $btnOpenDir = New-Object System.Windows.Forms.Button
    $btnOpenDir.Text = "Ordner öffnen"
    $btnOpenDir.Size = New-Object System.Drawing.Size(130, 34)
    $btnOpenDir.Add_Click({
        $dir = Split-Path -Parent $LogTxtPath
        if (Test-Path $dir) { Start-Process "explorer.exe" -ArgumentList "`"$dir`"" }
    })
    $pnlBtns.Controls.Add($btnOpenDir)

    if (Test-Path $LogTxtPath) {
        $btnOpenTxt = New-Object System.Windows.Forms.Button
        $btnOpenTxt.Text = "Detail-Log öffnen"
        $btnOpenTxt.Size = New-Object System.Drawing.Size(150, 34)
        $btnOpenTxt.Add_Click({ Start-Process -FilePath $LogTxtPath })
        $pnlBtns.Controls.Add($btnOpenTxt)
    }

    if ($FailedCount -gt 0 -and (Test-Path $FailCsvPath)) {
        $btnOpenFail = New-Object System.Windows.Forms.Button
        $btnOpenFail.Text = "Fehler-CSV öffnen"
        $btnOpenFail.Size = New-Object System.Drawing.Size(155, 34)
        $btnOpenFail.BackColor = [System.Drawing.Color]::FromArgb(254, 235, 235)
        $btnOpenFail.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnOpenFail.Font = New-Object System.Drawing.Font("Segoe UI", 9.0, [System.Drawing.FontStyle]::Bold)
        $btnOpenFail.Add_Click({ Start-Process -FilePath $FailCsvPath })
        $pnlBtns.Controls.Add($btnOpenFail)
    }

    $dlg.AcceptButton = $btnClose
    [void]$dlg.ShowDialog()
    $dlg.Dispose()
}

# -----------------------------------------------------------------------------
# HAUPTMODUL: TOOL 23 MIT ZWEI TABS
# -----------------------------------------------------------------------------
function Show-ADRenameComputersTool {
    $defaultDir = "C:\_Admin\Scripts\JHA"
    $defaultCsv = Join-Path $defaultDir "RenameClientsList.csv"
    $script:LastLogPath = ""
    $script:LastFailCsv = ""
    $script:InspectorADObj = $null

    $script:DomainDns = ""
    try {
        $script:DomainDns = (Get-ADDomain).DNSRoot.ToLower()
    } catch {
        $script:DomainDns = [System.Net.Dns]::GetHostByName("").HostName -replace '^[^.]*\.', ''
    }

    # Hauptfenster
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Tool 23: AD Remote Computer Rename & Entra-ID Inspector"
    $form.Size = New-Object System.Drawing.Size(1180, 820)
    $form.MinimumSize = New-Object System.Drawing.Size(1020, 680)
    $form.StartPosition = "CenterScreen"
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
    $form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

    # Tab-Steuerung
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Dock = [System.Windows.Forms.DockStyle]::Fill
    $tabControl.Font = New-Object System.Drawing.Font("Segoe UI", 10.0, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($tabControl)

    $tabRename    = New-Object System.Windows.Forms.TabPage
    $tabRename.Text = "1. Remote Rename Batch (CSV)"
    $tabRename.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
    $tabControl.TabPages.Add($tabRename)

    $tabInspector = New-Object System.Windows.Forms.TabPage
    $tabInspector.Text = "2. AD Konto-Check & Entra Sync"
    $tabInspector.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
    $tabControl.TabPages.Add($tabInspector)

    # =========================================================================
    # REITER 1: REMOTE RENAME BATCH
    # =========================================================================
    $pnlTop = New-Object System.Windows.Forms.GroupBox
    $pnlTop.Text = "Konfiguration & CSV-Steuerung"
    $pnlTop.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlTop.Height = 135
    $pnlTop.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    $tabRename.Controls.Add($pnlTop)

    $lblCsv = New-Object System.Windows.Forms.Label
    $lblCsv.Text = "Rename-CSV Pfad:"
    $lblCsv.Location = New-Object System.Drawing.Point(15, 28)
    $lblCsv.Size = New-Object System.Drawing.Size(150, 24)
    $lblCsv.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Regular)
    $pnlTop.Controls.Add($lblCsv)

    $txtCsvPath = New-Object System.Windows.Forms.TextBox
    $txtCsvPath.Text = $defaultCsv
    $txtCsvPath.Location = New-Object System.Drawing.Point(170, 25)
    $txtCsvPath.Size = New-Object System.Drawing.Size(460, 26)
    $txtCsvPath.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Regular)
    $pnlTop.Controls.Add($txtCsvPath)

    $btnBrowse = New-Object System.Windows.Forms.Button
    $btnBrowse.Text = "Durchsuchen..."
    $btnBrowse.Location = New-Object System.Drawing.Point(640, 23)
    $btnBrowse.Size = New-Object System.Drawing.Size(125, 30)
    $btnBrowse.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnBrowse.Font = New-Object System.Drawing.Font("Segoe UI", 9.0)
    $pnlTop.Controls.Add($btnBrowse)

    $btnCreateTpl = New-Object System.Windows.Forms.Button
    $btnCreateTpl.Text = "Vorlage erstellen"
    $btnCreateTpl.Location = New-Object System.Drawing.Point(775, 23)
    $btnCreateTpl.Size = New-Object System.Drawing.Size(150, 30)
    $btnCreateTpl.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnCreateTpl.BackColor = [System.Drawing.Color]::FromArgb(235, 240, 248)
    $btnCreateTpl.Font = New-Object System.Drawing.Font("Segoe UI", 9.0)
    $pnlTop.Controls.Add($btnCreateTpl)

    $btnOpenCsv = New-Object System.Windows.Forms.Button
    $btnOpenCsv.Text = "CSV öffnen / bearbeiten"
    $btnOpenCsv.Location = New-Object System.Drawing.Point(935, 23)
    $btnOpenCsv.Size = New-Object System.Drawing.Size(185, 30)
    $btnOpenCsv.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnOpenCsv.BackColor = [System.Drawing.Color]::FromArgb(235, 240, 248)
    $btnOpenCsv.Font = New-Object System.Drawing.Font("Segoe UI", 9.0)
    $pnlTop.Controls.Add($btnOpenCsv)

    $btnStart = New-Object System.Windows.Forms.Button
    $btnStart.Text = "Umbenennung starten"
    $btnStart.Location = New-Object System.Drawing.Point(170, 72)
    $btnStart.Size = New-Object System.Drawing.Size(220, 38)
    $btnStart.BackColor = [System.Drawing.Color]::FromArgb(190, 40, 40)
    $btnStart.ForeColor = [System.Drawing.Color]::White
    $btnStart.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnStart.Font = New-Object System.Drawing.Font("Segoe UI", 10.0, [System.Drawing.FontStyle]::Bold)
    $pnlTop.Controls.Add($btnStart)

    $btnOpenLogDir = New-Object System.Windows.Forms.Button
    $btnOpenLogDir.Text = "Log-Ordner öffnen"
    $btnOpenLogDir.Location = New-Object System.Drawing.Point(400, 72)
    $btnOpenLogDir.Size = New-Object System.Drawing.Size(160, 38)
    $btnOpenLogDir.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnOpenLogDir.BackColor = [System.Drawing.Color]::FromArgb(235, 240, 248)
    $btnOpenLogDir.Font = New-Object System.Drawing.Font("Segoe UI", 9.0, [System.Drawing.FontStyle]::Bold)
    $pnlTop.Controls.Add($btnOpenLogDir)

    $lblWarn = New-Object System.Windows.Forms.Label
    $lblWarn.Text = "Achtung: Führt bei Erreichbarkeit sofortigen Neustart (-Force -Restart) aus!"
    $lblWarn.Location = New-Object System.Drawing.Point(575, 82)
    $lblWarn.AutoSize = $true
    $lblWarn.ForeColor = [System.Drawing.Color]::FromArgb(170, 30, 30)
    $lblWarn.Font = New-Object System.Drawing.Font("Segoe UI", 9.0, [System.Drawing.FontStyle]::Italic)
    $pnlTop.Controls.Add($lblWarn)

    $splitRename = New-Object System.Windows.Forms.SplitContainer
    $splitRename.Dock = [System.Windows.Forms.DockStyle]::Fill
    $splitRename.Orientation = [System.Windows.Forms.Orientation]::Horizontal
    $splitRename.SplitterDistance = 290
    $tabRename.Controls.Add($splitRename)
    $splitRename.BringToFront()

    $grpGrid = New-Object System.Windows.Forms.GroupBox
    $grpGrid.Text = "Verarbeitungsliste (Ergebnisse)"
    $grpGrid.Dock = [System.Windows.Forms.DockStyle]::Fill
    $grpGrid.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    $splitRename.Panel1.Controls.Add($grpGrid)

    $gridResults = New-Object System.Windows.Forms.DataGridView
    $gridResults.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridResults.BackgroundColor = [System.Drawing.Color]::White
    $gridResults.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $gridResults.CellBorderStyle = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
    $gridResults.ReadOnly = $true
    $gridResults.AllowUserToAddRows = $false
    $gridResults.AllowUserToDeleteRows = $false
    $gridResults.RowHeadersVisible = $false
    $gridResults.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $gridResults.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $gridResults.EnableHeadersVisualStyles = $false
    $gridResults.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
    $gridResults.ColumnHeadersHeight = 36
    $gridResults.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 10.0, [System.Drawing.FontStyle]::Bold)
    $gridResults.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(235, 239, 245)
    $gridResults.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(35, 45, 55)
    $gridResults.ColumnHeadersDefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleLeft
    $gridResults.ColumnHeadersDefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(8, 4, 8, 4)
    $gridResults.RowTemplate.Height = 30
    $gridResults.DefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Regular)
    $gridResults.DefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(8, 2, 8, 2)
    $grpGrid.Controls.Add($gridResults)

    [void]$gridResults.Columns.Add("OldName", "Alter Name")
    [void]$gridResults.Columns.Add("NewName", "Neuer Name")
    [void]$gridResults.Columns.Add("Status", "Status")
    [void]$gridResults.Columns.Add("Details", "Details / Fehler")
    $gridResults.Columns["OldName"].FillWeight = 20
    $gridResults.Columns["NewName"].FillWeight = 20
    $gridResults.Columns["Status"].FillWeight = 25
    $gridResults.Columns["Details"].FillWeight = 35

    $grpLog = New-Object System.Windows.Forms.GroupBox
    $grpLog.Text = "Ausführungs-Log"
    $grpLog.Dock = [System.Windows.Forms.DockStyle]::Fill
    $grpLog.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    $splitRename.Panel2.Controls.Add($grpLog)

    $txtLog = New-Object System.Windows.Forms.RichTextBox
    $txtLog.Dock = [System.Windows.Forms.DockStyle]::Fill
    $txtLog.BackColor = [System.Drawing.Color]::FromArgb(20, 24, 30)
    $txtLog.ForeColor = [System.Drawing.Color]::White
    $txtLog.Font = New-Object System.Drawing.Font("Consolas", 9.5)
    $txtLog.ReadOnly = $true
    $grpLog.Controls.Add($txtLog)

    # =========================================================================
    # REITER 2: AD KONTO-INSPECTOR & ENTRA ID SYNC
    # =========================================================================
    $grpInspSearch = New-Object System.Windows.Forms.GroupBox
    $grpInspSearch.Text = "Computer-Auswahl (manuell oder aus Rename-CSV)"
    $grpInspSearch.Dock = [System.Windows.Forms.DockStyle]::Top
    $grpInspSearch.Height = 85
    $grpInspSearch.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    $tabInspector.Controls.Add($grpInspSearch)

    $lblInspSelect = New-Object System.Windows.Forms.Label
    $lblInspSelect.Text = "PC-Name wählen:"
    $lblInspSelect.Location = New-Object System.Drawing.Point(15, 32)
    $lblInspSelect.Size = New-Object System.Drawing.Size(125, 24)
    $lblInspSelect.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Regular)
    $grpInspSearch.Controls.Add($lblInspSelect)

    $cmbInspComputers = New-Object System.Windows.Forms.ComboBox
    $cmbInspComputers.Location = New-Object System.Drawing.Point(145, 29)
    $cmbInspComputers.Size = New-Object System.Drawing.Size(320, 28)
    $cmbInspComputers.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Regular)
    $grpInspSearch.Controls.Add($cmbInspComputers)

    $btnInspCheck = New-Object System.Windows.Forms.Button
    $btnInspCheck.Text = "AD Prüfen"
    $btnInspCheck.Location = New-Object System.Drawing.Point(475, 27)
    $btnInspCheck.Size = New-Object System.Drawing.Size(125, 32)
    $btnInspCheck.BackColor = [System.Drawing.Color]::FromArgb(37, 99, 235)
    $btnInspCheck.ForeColor = [System.Drawing.Color]::White
    $btnInspCheck.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnInspCheck.Font = New-Object System.Drawing.Font("Segoe UI", 9.0, [System.Drawing.FontStyle]::Bold)
    $grpInspSearch.Controls.Add($btnInspCheck)

    $btnLoadFromCsv = New-Object System.Windows.Forms.Button
    $btnLoadFromCsv.Text = "Aus CSV laden"
    $btnLoadFromCsv.Location = New-Object System.Drawing.Point(610, 27)
    $btnLoadFromCsv.Size = New-Object System.Drawing.Size(150, 32)
    $btnLoadFromCsv.BackColor = [System.Drawing.Color]::FromArgb(235, 240, 248)
    $btnLoadFromCsv.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnLoadFromCsv.Font = New-Object System.Drawing.Font("Segoe UI", 9.0, [System.Drawing.FontStyle]::Bold)
    $grpInspSearch.Controls.Add($btnLoadFromCsv)

    $grpInspStatus = New-Object System.Windows.Forms.GroupBox
    $grpInspStatus.Text = "AD-Attributstatus & Entra-ID Synchronisations-Relevanz"
    $grpInspStatus.Dock = [System.Windows.Forms.DockStyle]::Top
    $grpInspStatus.Height = 160
    $grpInspStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    $tabInspector.Controls.Add($grpInspStatus)
    $grpInspStatus.BringToFront()

    $createInspRow = {
        param($Parent, $Top, $Title, $HasStatus = $false)
        $lblT = New-Object System.Windows.Forms.Label
        $lblT.Text = $Title
        $lblT.Location = New-Object System.Drawing.Point(15, $Top)
        $lblT.Size = New-Object System.Drawing.Size(140, 22)
        $lblT.Font = New-Object System.Drawing.Font("Segoe UI", 9.0, [System.Drawing.FontStyle]::Bold)
        $lblT.ForeColor = [System.Drawing.Color]::FromArgb(50, 60, 75)
        $Parent.Controls.Add($lblT)

        $lblV = New-Object System.Windows.Forms.Label
        $lblV.Text = "-"
        $lblV.Location = New-Object System.Drawing.Point(160, $Top)
        $lblV.Size = New-Object System.Drawing.Size(520, 22)
        $lblV.Font = New-Object System.Drawing.Font("Segoe UI", 9.0, [System.Drawing.FontStyle]::Regular)
        $Parent.Controls.Add($lblV)

        $lblS = $null
        if ($HasStatus) {
            $lblS = New-Object System.Windows.Forms.Label
            $lblS.Text = ""
            $lblS.Location = New-Object System.Drawing.Point(690, $Top)
            $lblS.Size = New-Object System.Drawing.Size(180, 22)
            $lblS.Font = New-Object System.Drawing.Font("Segoe UI", 9.0, [System.Drawing.FontStyle]::Bold)
            $Parent.Controls.Add($lblS)
        }
        return @{ Val = $lblV; Status = $lblS }
    }

    $rSam  = & $createInspRow $grpInspStatus 26  "sAMAccountName:" $false
    $rDisp = & $createInspRow $grpInspStatus 54  "displayName:"    $true
    $rDns  = & $createInspRow $grpInspStatus 82  "dNSHostName:"    $true
    $rDesc = & $createInspRow $grpInspStatus 110 "Description:"    $false

    $pnlInspBottom = New-Object System.Windows.Forms.Panel
    $pnlInspBottom.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $pnlInspBottom.Height = 55
    $pnlInspBottom.BackColor = [System.Drawing.Color]::FromArgb(238, 242, 246)
    $tabInspector.Controls.Add($pnlInspBottom)

    $btnFixAD = New-Object System.Windows.Forms.Button
    $btnFixAD.Text = "AD-Attribute korrigieren"
    $btnFixAD.Location = New-Object System.Drawing.Point(15, 10)
    $btnFixAD.Size = New-Object System.Drawing.Size(200, 34)
    $btnFixAD.BackColor = [System.Drawing.Color]::FromArgb(5, 150, 105)
    $btnFixAD.ForeColor = [System.Drawing.Color]::White
    $btnFixAD.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnFixAD.Font = New-Object System.Drawing.Font("Segoe UI", 9.0, [System.Drawing.FontStyle]::Bold)
    $btnFixAD.Enabled = $false
    $pnlInspBottom.Controls.Add($btnFixAD)

    $btnTriggerEntra = New-Object System.Windows.Forms.Button
    $btnTriggerEntra.Text = "Entra Connect Sync triggern"
    $btnTriggerEntra.Location = New-Object System.Drawing.Point(225, 10)
    $btnTriggerEntra.Size = New-Object System.Drawing.Size(210, 34)
    $btnTriggerEntra.BackColor = [System.Drawing.Color]::FromArgb(75, 85, 99)
    $btnTriggerEntra.ForeColor = [System.Drawing.Color]::White
    $btnTriggerEntra.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnTriggerEntra.Font = New-Object System.Drawing.Font("Segoe UI", 9.0, [System.Drawing.FontStyle]::Bold)
    $pnlInspBottom.Controls.Add($btnTriggerEntra)

    $btnClientInfo = New-Object System.Windows.Forms.Button
    $btnClientInfo.Text = "Client dsregcmd Befehle"
    $btnClientInfo.Location = New-Object System.Drawing.Point(445, 10)
    $btnClientInfo.Size = New-Object System.Drawing.Size(190, 34)
    $btnClientInfo.BackColor = [System.Drawing.Color]::FromArgb(59, 130, 246)
    $btnClientInfo.ForeColor = [System.Drawing.Color]::White
    $btnClientInfo.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnClientInfo.Font = New-Object System.Drawing.Font("Segoe UI", 9.0, [System.Drawing.FontStyle]::Bold)
    $pnlInspBottom.Controls.Add($btnClientInfo)

    $grpInspLog = New-Object System.Windows.Forms.GroupBox
    $grpInspLog.Text = "Aktivitäten / Protokoll"
    $grpInspLog.Dock = [System.Windows.Forms.DockStyle]::Fill
    $grpInspLog.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    $tabInspector.Controls.Add($grpInspLog)
    $grpInspLog.BringToFront()

    $txtInspLog = New-Object System.Windows.Forms.RichTextBox
    $txtInspLog.Dock = [System.Windows.Forms.DockStyle]::Fill
    $txtInspLog.BackColor = [System.Drawing.Color]::FromArgb(20, 24, 30)
    $txtInspLog.ForeColor = [System.Drawing.Color]::White
    $txtInspLog.Font = New-Object System.Drawing.Font("Consolas", 9.5)
    $txtInspLog.ReadOnly = $true
    $grpInspLog.Controls.Add($txtInspLog)

    # -------------------------------------------------------------------------
    # HILFSLOGIKEN & ACTIONS
    # -------------------------------------------------------------------------
    $logBatch = {
        param([string]$Msg, [System.Drawing.Color]$Color = [System.Drawing.Color]::White, [string]$FileLogPath = "")
        $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        $entry = "[$ts] $Msg`r`n"
        $txtLog.SelectionStart = $txtLog.TextLength
        $txtLog.SelectionLength = 0
        $txtLog.SelectionColor = $Color
        $txtLog.AppendText($entry)
        $txtLog.ScrollToCaret()

        if ($FileLogPath) {
            try { $entry.TrimEnd("`r`n") | Out-File -FilePath $FileLogPath -Append -Encoding utf8 } catch {}
        }
    }

    $logInspector = {
        param([string]$Msg, [System.Drawing.Color]$Color = [System.Drawing.Color]::White)
        $ts = (Get-Date).ToString("HH:mm:ss")
        $entry = "[$ts] $Msg`r`n"
        $txtInspLog.SelectionStart = $txtInspLog.TextLength
        $txtInspLog.SelectionLength = 0
        $txtInspLog.SelectionColor = $Color
        $txtInspLog.AppendText($entry)
        $txtInspLog.ScrollToCaret()
    }

    # CSV in ComboBox des Inspectors laden
    $populateInspectorDropdown = {
        $csvFile = $txtCsvPath.Text.Trim()
        if (Test-Path $csvFile) {
            try {
                $imported = @(Import-Csv -Path $csvFile -Delimiter ",")
                $cmbInspComputers.Items.Clear()
                foreach ($row in $imported) {
                    if ($row.NeuerName) {
                        [void]$cmbInspComputers.Items.Add($row.NeuerName.Trim())
                    }
                }
                if ($cmbInspComputers.Items.Count -gt 0) {
                    $cmbInspComputers.SelectedIndex = 0
                }
            } catch {}
        }
    }

    # -------------------------------------------------------------------------
    # EVENTS: REITER 1 (RENAME)
    # -------------------------------------------------------------------------
    $btnBrowse.Add_Click({
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Filter = "CSV-Dateien (*.csv)|*.csv|Alle Dateien (*.*)|*.*"
        $ofd.InitialDirectory = if (Test-Path $txtCsvPath.Text) { Split-Path $txtCsvPath.Text } else { "C:\" }
        if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtCsvPath.Text = $ofd.FileName
            & $populateInspectorDropdown
        }
    })

    $btnCreateTpl.Add_Click({
        $csvFile = $txtCsvPath.Text.Trim()
        $dir = Split-Path -Path $csvFile -Parent
        try {
            if (-not [string]::IsNullOrWhiteSpace($dir) -and (-not (Test-Path $dir))) {
                New-Item -ItemType Directory -Path $dir -Force | Out-Null
            }
            if (-not (Test-Path $csvFile)) {
                $templateContent = @"
AlterName,NeuerName
CLIENT001,NB-GF-01
CLIENT002,WS-BU-05
"@
                $templateContent.Trim() | Out-File -FilePath $csvFile -Encoding utf8
                [System.Windows.Forms.MessageBox]::Show("Vorlage erstellt unter:`r`n$csvFile", "Erstellt", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                & $populateInspectorDropdown
            } else {
                [System.Windows.Forms.MessageBox]::Show("Die Datei existiert bereits unter:`r`n$csvFile", "Hinweis", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Erstellen der Vorlage:`r`n$($_.Exception.Message)", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })

    $btnOpenCsv.Add_Click({
        $csvFile = $txtCsvPath.Text.Trim()
        if (-not (Test-Path $csvFile)) {
            $diag = [System.Windows.Forms.MessageBox]::Show("Die Datei existiert noch nicht. Soll sie jetzt neu erstellt werden?", "Datei nicht gefunden", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($diag -eq [System.Windows.Forms.DialogResult]::Yes) {
                $btnCreateTpl.PerformClick()
            } else { return }
        }
        if (Test-Path $csvFile) { Start-Process -FilePath $csvFile }
    })

    $btnOpenLogDir.Add_Click({
        $dir = if (-not [string]::IsNullOrWhiteSpace($script:LastLogPath) -and (Test-Path (Split-Path $script:LastLogPath))) {
            Split-Path $script:LastLogPath
        } elseif (Test-Path $txtCsvPath.Text) {
            Split-Path $txtCsvPath.Text
        } else {
            $defaultDir
        }
        if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        Start-Process "explorer.exe" -ArgumentList "`"$dir`""
    })

    $btnStart.Add_Click({
        $csvFile = $txtCsvPath.Text.Trim()
        if (-not (Test-Path $csvFile)) {
            [System.Windows.Forms.MessageBox]::Show("Die CSV-Datei existiert nicht!`r`n$csvFile", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        try {
            $computers = @(Import-Csv -Path $csvFile -Delimiter ",")
            if ($computers.Count -eq 0 -or -not $computers[0].PSObject.Properties['AlterName'] -or -not $computers[0].PSObject.Properties['NeuerName']) {
                [System.Windows.Forms.MessageBox]::Show("Die CSV muss die Spalten 'AlterName' und 'NeuerName' enthalten.", "Formatfehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Einlesen:`r`n$($_.Exception.Message)", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }

        if (-not (Show-RenameConfirmDialog -ComputersList $computers)) { return }

        $domainCred = $null
        try {
            $domainCred = Get-Credential -UserName "$env:USERDOMAIN\" -Message "Bitte Administrator-Zugangsdaten eingeben"
        } catch { return }
        if ($null -eq $domainCred) { return }

        $btnStart.Enabled = $false
        $btnBrowse.Enabled = $false
        $btnCreateTpl.Enabled = $false
        $btnOpenCsv.Enabled = $false
        $gridResults.Rows.Clear()
        $txtLog.Clear()

        $adminUser   = $domainCred.UserName
        $targetDir   = Split-Path -Path $csvFile -Parent
        $logPath     = Join-Path $targetDir "RenameLog_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').txt"
        $failCsvPath = Join-Path $targetDir "Fehlgeschlagene_Computer_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"

        $script:LastLogPath = $logPath
        $script:LastFailCsv = $failCsvPath
        $failedList = [System.Collections.Generic.List[PSCustomObject]]::new()

        & $logBatch "=== START: Remote-Rename gestartet ===" ([System.Drawing.Color]::Cyan) $logPath
        & $logBatch "Einträge in CSV: $($computers.Count)" ([System.Drawing.Color]::White) $logPath

        foreach ($entry in $computers) {
            $oldName = $entry.AlterName.Trim()
            $newName = $entry.NeuerName.Trim()
            $status = "Unbekannt"
            $fehlerDetails = ""

            & $logBatch "--------------------------------------------------" ([System.Drawing.Color]::Gray) $logPath
            & $logBatch "Verarbeite Client: $oldName -> $newName" ([System.Drawing.Color]::FromArgb(100, 180, 255)) $logPath
            [System.Windows.Forms.Application]::DoEvents()

            if (-not (Test-Connection -ComputerName $oldName -Count 1 -Quiet)) {
                $status = "Offline"
                $fehlerDetails = "Client antwortet nicht auf ICMP (Ping)."
                & $logBatch "[-] $fehlerDetails" ([System.Drawing.Color]::Orange) $logPath
            } else {
                & $logBatch "[+] Ping erfolgreich." ([System.Drawing.Color]::LightGray) $logPath

                $tcpSocket = New-Object System.Net.Sockets.TcpClient
                $portOk = $false
                try {
                    $asyncResult = $tcpSocket.BeginConnect($oldName, 445, $null, $null)
                    $waitHandle = $asyncResult.AsyncWaitHandle.WaitOne(1500, $false)
                    if ($waitHandle -and $tcpSocket.Connected) {
                        $tcpSocket.EndConnect($asyncResult)
                        $portOk = $true
                    }
                } catch { $portOk = $false }
                finally { $tcpSocket.Close(); $tcpSocket.Dispose() }

                if (-not $portOk) {
                    $status = "RPC/SMB blockiert"
                    $fehlerDetails = "Port 445 nicht erreichbar (Datei- und Druckerfreigabe / Client-Firewall)."
                    & $logBatch "[-] $fehlerDetails" ([System.Drawing.Color]::Orange) $logPath
                } else {
                    & $logBatch "[+] Port 445 erreichbar. Sende Rename-Befehl..." ([System.Drawing.Color]::LightGray) $logPath

                    try {
                        Rename-Computer -ComputerName $oldName -NewName $newName -DomainCredential $domainCred -Force -Restart -ErrorAction Stop
                        & $logBatch "[*] Befehl gesendet. Warte auf AD-Synchronisation..." ([System.Drawing.Color]::Yellow) $logPath
                        [System.Windows.Forms.Application]::DoEvents()

                        $adCheck = $null
                        $timeoutSeconds = 25
                        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

                        while ($stopwatch.Elapsed.TotalSeconds -lt $timeoutSeconds) {
                            Start-Sleep -Seconds 3
                            $adCheck = Get-ADComputer -Filter "Name -eq '$newName' -or sAMAccountName -eq '$newName$'" -Credential $domainCred -Properties Description, DisplayName -ErrorAction SilentlyContinue
                            if ($adCheck) { break }
                            [System.Windows.Forms.Application]::DoEvents()
                        }
                        $stopwatch.Stop()

                        if ($adCheck) {
                            & $logBatch "[SUCCESS] $oldName erfolgreich in $newName umbenannt (AD bestätigt)!" ([System.Drawing.Color]::LightGreen) $logPath

                            try {
                                $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                                $descEntry = "Umbenannt von $oldName auf $newName am $ts durch $adminUser"
                                $newDesc = if ([string]::IsNullOrWhiteSpace($adCheck.Description)) { $descEntry } else { "$($adCheck.Description) | $descEntry" }

                                Set-ADComputer -Identity $adCheck.DistinguishedName -Credential $domainCred -DisplayName $newName -Description $newDesc -ErrorAction Stop
                                & $logBatch "[+] AD-Attribute angepasst: DisplayName='$newName', Description gesetzt." ([System.Drawing.Color]::LightGreen) $logPath
                                $status = "Erfolgreich"
                            } catch {
                                $status = "Erfolgreich mit Attributwarnung"
                                $fehlerDetails = "PC umbenannt, Attribute unvollständig: $($_.Exception.Message)"
                                & $logBatch "[!] $fehlerDetails" ([System.Drawing.Color]::Yellow) $logPath
                            }
                        } else {
                            $status = "AD-Verifikation Timeout"
                            $fehlerDetails = "Rename abgesetzt, aber Objekt nach 25s im AD noch nicht als '$newName' auffindbar."
                            & $logBatch "[-] $fehlerDetails" ([System.Drawing.Color]::Salmon) $logPath
                        }
                    } catch {
                        $status = "Rename fehlgeschlagen"
                        $fehlerDetails = $_.Exception.Message
                        & $logBatch "[-] Rename fehlgeschlagen: $fehlerDetails" ([System.Drawing.Color]::Salmon) $logPath
                    }
                }
            }

            $rowIndex = $gridResults.Rows.Add($oldName, $newName, $status, $fehlerDetails)
            $row = $gridResults.Rows[$rowIndex]
            if ($status -eq "Erfolgreich") {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(235, 250, 235)
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkGreen
            } else {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 238, 238)
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkRed
                $failedList.Add([PSCustomObject]@{
                    AlterName        = $oldName
                    GewuenschterName = $newName
                    Status           = $status
                    FehlerDetails    = $fehlerDetails
                })
            }
            [System.Windows.Forms.Application]::DoEvents()
        }

        & $logBatch "--------------------------------------------------" ([System.Drawing.Color]::Gray) $logPath
        if ($failedList.Count -gt 0) {
            try {
                $failedList | Export-Csv -Path $failCsvPath -NoTypeInformation -Delimiter ";" -Encoding utf8
                & $logBatch "=== FINISH: Beendet mit $($failedList.Count) Fehlern. ===" ([System.Drawing.Color]::Orange) $logPath
            } catch {}
        } else {
            & $logBatch "=== FINISH: Alle Clients erfolgreich verarbeitet! ===" ([System.Drawing.Color]::LightGreen) $logPath
        }

        & $populateInspectorDropdown
        Show-ExecutionFinishedDialog -FailedCount $failedList.Count -FailCsvPath $failCsvPath -LogTxtPath $logPath

        $btnStart.Enabled = $true
        $btnBrowse.Enabled = $true
        $btnCreateTpl.Enabled = $true
        $btnOpenCsv.Enabled = $true
    })

    # -------------------------------------------------------------------------
    # EVENTS: REITER 2 (INSPECTOR)
    # -------------------------------------------------------------------------
    $btnLoadFromCsv.Add_Click({
        & $populateInspectorDropdown
        & $logInspector "Computernamen aus CSV in die Auswahlliste übernommen." ([System.Drawing.Color]::Cyan)
    })

    $btnInspCheck.Add_Click({
        $searchName = $cmbInspComputers.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($searchName)) {
            [System.Windows.Forms.MessageBox]::Show("Bitte einen Computernamen auswählen oder eingeben!", "Eingabe fehlt", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        & $logInspector "Suche nach Computer-Objekt '$searchName' im AD..." ([System.Drawing.Color]::Cyan)

        try {
            $comp = Get-ADComputer -Filter "Name -eq '$searchName' -or sAMAccountName -eq '$searchName$' -or sAMAccountName -eq '$searchName'" `
                                   -Properties DisplayName, dNSHostName, Description, sAMAccountName -ErrorAction Stop

            if (-not $comp) {
                & $logInspector "[-] Kein Computerobjekt für '$searchName' gefunden." ([System.Drawing.Color]::Salmon)
                $btnFixAD.Enabled = $false
                $rSam.Val.Text = "-"
                $rDisp.Val.Text = "-"
                $rDisp.Status.Text = ""
                $rDns.Val.Text = "-"
                $rDns.Status.Text = ""
                $rDesc.Val.Text = "-"
                return
            }

            $script:InspectorADObj = $comp
            $cleanName = $comp.Name

            $rSam.Val.Text  = $comp.sAMAccountName
            $rDisp.Val.Text = if ($comp.DisplayName) { $comp.DisplayName } else { "<NICHT GESETZT>" }
            $rDns.Val.Text  = if ($comp.dNSHostName) { $comp.dNSHostName } else { "<NICHT GESETZT>" }
            $rDesc.Val.Text = if ($comp.Description) { $comp.Description } else { "<LEER>" }

            # DisplayName Prüfung
            if ([string]::IsNullOrWhiteSpace($comp.DisplayName)) {
                $rDisp.Status.Text = "[!] Fehlt"
                $rDisp.Status.ForeColor = [System.Drawing.Color]::Red
                & $logInspector "[WARNUNG] 'displayName' ist leer! Entra Connect nutzt Fallback '$($comp.sAMAccountName)'." ([System.Drawing.Color]::Orange)
            } elseif ($comp.DisplayName.EndsWith('$')) {
                $rDisp.Status.Text = "[!] Hat $"
                $rDisp.Status.ForeColor = [System.Drawing.Color]::Red
                & $logInspector "[FEHLER] 'displayName' endet mit '$' -> Entra zeigt Dollarzeichen." ([System.Drawing.Color]::Salmon)
            } elseif ($comp.DisplayName -ne $cleanName) {
                $rDisp.Status.Text = "[!] Abweichend"
                $rDisp.Status.ForeColor = [System.Drawing.Color]::Orange
                & $logInspector "[WARNUNG] 'displayName' ($($comp.DisplayName)) weicht von Name ab ($cleanName)." ([System.Drawing.Color]::Orange)
            } else {
                $rDisp.Status.Text = "[OK]"
                $rDisp.Status.ForeColor = [System.Drawing.Color]::DarkGreen
            }

            # dNSHostName Prüfung
            $expectedDns = "$cleanName.$($script:DomainDns)".ToLower()
            if ([string]::IsNullOrWhiteSpace($comp.dNSHostName) -or $comp.dNSHostName.ToLower() -ne $expectedDns) {
                $rDns.Status.Text = "[!] Abweichend"
                $rDns.Status.ForeColor = [System.Drawing.Color]::Orange
                & $logInspector "[INFO] dNSHostName weicht ab. Client registriert diesen beim nächsten Booten selbst." ([System.Drawing.Color]::Yellow)
            } else {
                $rDns.Status.Text = "[OK]"
                $rDns.Status.ForeColor = [System.Drawing.Color]::DarkGreen
            }

            $btnFixAD.Enabled = $true
            & $logInspector "[+] Objekt '$cleanName' erfolgreich geladen. Bereit für Aktionen." ([System.Drawing.Color]::LightGreen)
        } catch {
            & $logInspector "[-] Fehler beim Abfragen des AD-Objekts: $($_.Exception.Message)" ([System.Drawing.Color]::Salmon)
            $btnFixAD.Enabled = $false
        }
    })

    $cmbInspComputers.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $_.SuppressKeyPress = $true
            $btnInspCheck.PerformClick()
        }
    })

    $btnFixAD.Add_Click({
        if (-not $script:InspectorADObj) { return }

        $comp = $script:InspectorADObj
        $cleanName = $comp.Name
        $user = "$env:USERDOMAIN\$env:USERNAME"
        $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        $logEntry = "Korrektur am $ts durch $user"
        $newDesc = if ([string]::IsNullOrWhiteSpace($comp.Description)) { $logEntry } else { "$($comp.Description) | $logEntry" }

        try {
            & $logInspector "Setze DisplayName und Description für '$cleanName'..." ([System.Drawing.Color]::Yellow)
            Set-ADComputer -Identity $comp.DistinguishedName -DisplayName $cleanName -Description $newDesc -ErrorAction Stop

            & $logInspector "[ERFOLG] DisplayName='$cleanName' und Description angepasst." ([System.Drawing.Color]::LightGreen)
            $rDisp.Val.Text = $cleanName
            $rDisp.Status.Text = "[OK]"
            $rDisp.Status.ForeColor = [System.Drawing.Color]::DarkGreen
            $rDesc.Val.Text = $newDesc

            [System.Windows.Forms.MessageBox]::Show("AD-Attribute wurden erfolgreich bereinigt!`r`n`r`nDisplayName: $cleanName", "Erfolg", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } catch {
            & $logInspector "[-] Fehler: $($_.Exception.Message)" ([System.Drawing.Color]::Salmon)
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Aktualisieren: $($_.Exception.Message)", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })

    $btnTriggerEntra.Add_Click({
        $inputForm = New-Object System.Windows.Forms.Form
        $inputForm.Text = "Entra Sync Server angeben"
        $inputForm.Size = New-Object System.Drawing.Size(420, 170)
        $inputForm.StartPosition = "CenterParent"
        $inputForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $inputForm.MaximizeBox = $false
        $inputForm.MinimizeBox = $false

        $lblIn = New-Object System.Windows.Forms.Label
        $lblIn.Text = "Name des Entra Connect Servers ('localhost' falls lokal):"
        $lblIn.Location = New-Object System.Drawing.Point(15, 15)
        $lblIn.Size = New-Object System.Drawing.Size(370, 25)
        $inputForm.Controls.Add($lblIn)

        $txtIn = New-Object System.Windows.Forms.TextBox
        $txtIn.Text = "localhost"
        $txtIn.Location = New-Object System.Drawing.Point(18, 45)
        $txtIn.Size = New-Object System.Drawing.Size(365, 25)
        $inputForm.Controls.Add($txtIn)

        $btnOk = New-Object System.Windows.Forms.Button
        $btnOk.Text = "Starten"
        $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $btnOk.Location = New-Object System.Drawing.Point(200, 85)
        $btnOk.Size = New-Object System.Drawing.Size(85, 30)
        $inputForm.Controls.Add($btnOk)

        $btnCancel = New-Object System.Windows.Forms.Button
        $btnCancel.Text = "Abbrechen"
        $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $btnCancel.Location = New-Object System.Drawing.Point(295, 85)
        $btnCancel.Size = New-Object System.Drawing.Size(88, 30)
        $inputForm.Controls.Add($btnCancel)
        $inputForm.AcceptButton = $btnOk

        if ($inputForm.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK -or [string]::IsNullOrWhiteSpace($txtIn.Text)) { return }
        $syncServer = $txtIn.Text.Trim()

        & $logInspector "Starte Delta-Sync auf Server '$syncServer'..." ([System.Drawing.Color]::Yellow)

        try {
            if ($syncServer -eq "localhost" -or $syncServer -eq $env:COMPUTERNAME) {
                Start-ADSyncSyncCycle -PolicyType Delta -ErrorAction Stop
            } else {
                Invoke-Command -ComputerName $syncServer -ScriptBlock { Start-ADSyncSyncCycle -PolicyType Delta } -ErrorAction Stop
            }
            & $logInspector "[ERFOLG] Entra Connect Delta-Sync erfolgreich gestartet!" ([System.Drawing.Color]::LightGreen)
            [System.Windows.Forms.MessageBox]::Show("Delta-Sync gestartet! Daten sollten in ca. 2-5 Minuten im Entra Portal synchronisiert sein.", "Sync gestartet", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } catch {
            & $logInspector "[-] Fehler beim Sync-Aufruf: $($_.Exception.Message)" ([System.Drawing.Color]::Salmon)
            & $logInspector "[TIPP] Führe direkt auf dem Sync-Server aus: Start-ADSyncSyncCycle -PolicyType Delta" ([System.Drawing.Color]::Yellow)
        }
    })

    $btnClientInfo.Add_Click({
        $pc = if ($script:InspectorADObj) { $script:InspectorADObj.Name } else { "<CLIENT_NAME>" }

        & $logInspector "==========================================================" ([System.Drawing.Color]::Cyan)
        & $logInspector "MANUELLE DIAGNOSE-BEFEHLE FÜR DEN CLIENT ($pc):" ([System.Drawing.Color]::White)
        & $logInspector "----------------------------------------------------------" ([System.Drawing.Color]::Gray)
        & $logInspector "1. Prüfen, ob der neue Hostname an Entra gemeldet wird:" ([System.Drawing.Color]::LightGray)
        & $logInspector "   dsregcmd /status" ([System.Drawing.Color]::Yellow)
        & $logInspector "   (Unter 'Diagnostic Data' sollte 'HostNameUpdated : YES' stehen)" ([System.Drawing.Color]::LightGray)
        & $logInspector "" ([System.Drawing.Color]::White)
        & $logInspector "2. Entra-Join-Task am Client sofort anstoßen (Admin CMD):" ([System.Drawing.Color]::LightGray)
        & $logInspector "   schtasks /run /tn `"\Microsoft\Windows\Workplace Join\Automatic-Device-Join`"" ([System.Drawing.Color]::Yellow)
        & $logInspector "" ([System.Drawing.Color]::White)
        & $logInspector "3. Intune-Synchronisation am Client forcieren:" ([System.Drawing.Color]::LightGray)
        & $logInspector "   Restart-Service IntuneManagementExtension" ([System.Drawing.Color]::Yellow)
        & $logInspector "==========================================================" ([System.Drawing.Color]::Cyan)
    })

    # Initiales Füllen der ComboBox beim Start
    & $populateInspectorDropdown

    [void]$form.ShowDialog()
}

# Standalone Aufruf
Show-ADRenameComputersTool
