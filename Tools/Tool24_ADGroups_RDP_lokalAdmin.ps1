<#
.SYNOPSIS
    Tool 24: AD Computer RDP & lokalAdmin Gruppen-Manager
    Version: 2.9.0 (Fix: btnLoadMembersTab3 Event & Auto-Load bei Tab-Wechsel)
.DESCRIPTION
    - Tab 1: Computer & Gruppenprüfung (Spalte 2: Beschreibung, Befehlsbox, Tabellen- & Mitglieder-Export).
    - Tab 2: Gegenprüfung verwaister Gruppen ohne Computerobjekt (Einzel-Backups vor Löschen).
    - Tab 3: Gruppenmitglieder-Übersicht mit Spalten Computer-Status & Typ,
             Filtern nach Server/Client, Status & Text sowie Bereinigung mit Vorab-Export.
    - Automatisches Einlesen der Mitglieder bei Wechsel auf Tab 3.
    - Fester Pfad für CSV-Backups & Audit-Logfile (Activity_Log_JJJJMMDD.log).
    - Echte AD-Verifikation nach Erstellung.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.DirectoryServices
Import-Module ActiveDirectory -ErrorAction SilentlyContinue

[System.Windows.Forms.Application]::EnableVisualStyles()

# ------------------------------------------------------------------------------
# DOMÄNEN-VERBINDUNG
# ------------------------------------------------------------------------------
try {
    $domainObj  = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
    $rootDse    = [ADSI]"LDAP://RootDSE"
    $domainDN   = $rootDse.defaultNamingContext.Value
    $domainName = $domainObj.Name
    $rootDse.Dispose()
} catch {
    [System.Windows.Forms.MessageBox]::Show(
        "Keine Verbindung zu einer Active Directory Domäne möglich.",
        "AD Verbindungsfehler",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    return
}

# ------------------------------------------------------------------------------
# DATENSPEICHER
# ------------------------------------------------------------------------------
$script:RawCheckResults           = [System.Collections.Generic.List[PSCustomObject]]::new()
$script:CurrentFilteredRows       = [System.Collections.Generic.List[PSCustomObject]]::new()
$script:PendingCreateCmds         = [System.Collections.Generic.List[string]]::new()
$script:PendingDeleteGroups       = [System.Collections.Generic.List[PSCustomObject]]::new()
$script:OrphanedGroupsList        = [System.Collections.Generic.List[PSCustomObject]]::new()

$script:AllMembersList            = [System.Collections.Generic.List[PSCustomObject]]::new()
$script:CurrentMemberRows         = [System.Collections.Generic.List[PSCustomObject]]::new()
$script:PendingRemoveMemberCmds   = [System.Collections.Generic.List[string]]::new()

# ------------------------------------------------------------------------------
# THEME & LAYOUT HELPER
# ------------------------------------------------------------------------------
$UITheme = @{
    HeaderHeight    = 28
    RowHeight       = 24
    HeaderFont      = New-Object System.Drawing.Font("Segoe UI", 9.0, [System.Drawing.FontStyle]::Bold)
    CellFont        = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Regular)
    HeaderBackColor = [System.Drawing.Color]::FromArgb(238, 242, 246)
    HeaderForeColor = [System.Drawing.Color]::FromArgb(30, 41, 59)
    GridLineColor   = [System.Drawing.Color]::FromArgb(226, 232, 240)
}

function Apply-StandardGridTheme {
    param([System.Windows.Forms.DataGridView]$Grid)
    $Grid.EnableHeadersVisualStyles = $false
    $Grid.BorderStyle               = [System.Windows.Forms.BorderStyle]::FixedSingle
    $Grid.CellBorderStyle           = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
    $Grid.GridColor                 = $UITheme.GridLineColor
    $Grid.BackgroundColor           = [System.Drawing.Color]::White
    $Grid.RowHeadersVisible         = $false
    $Grid.AllowUserToAddRows        = $false
    $Grid.AllowUserToDeleteRows     = $false
    $Grid.SelectionMode             = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $Grid.MultiSelect               = $true
    $Grid.ReadOnly                  = $true
    $Grid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
    $Grid.ColumnHeadersHeight       = $UITheme.HeaderHeight
    $Grid.ColumnHeadersDefaultCellStyle.Font      = $UITheme.HeaderFont
    $Grid.ColumnHeadersDefaultCellStyle.BackColor = $UITheme.HeaderBackColor
    $Grid.ColumnHeadersDefaultCellStyle.ForeColor = $UITheme.HeaderForeColor
    $Grid.DefaultCellStyle.Font                   = $UITheme.CellFont
    $Grid.RowTemplate.Height                      = $UITheme.RowHeight
}

function Enable-UniversalGridSorting {
    param([System.Windows.Forms.DataGridView]$Grid)

    $Grid.Add_ColumnHeaderMouseClick({
        param($sender, $e)
        $targetGrid = $sender
        $col = $targetGrid.Columns[$e.ColumnIndex]
        if (-not $col) { return }

        $propName = if ($col.DataPropertyName) { $col.DataPropertyName } else { $col.HeaderText }
        if (-not $propName) { return }

        if (-not $targetGrid.Tag -or -not ($targetGrid.Tag -is [hashtable])) {
            $targetGrid.Tag = @{ LastCol = ""; Asc = $true }
        }

        $state = $targetGrid.Tag
        if ($state.LastCol -eq $propName) {
            $state.Asc = -not $state.Asc
        } else {
            $state.LastCol = $propName
            $state.Asc = $true
        }

        $items = @($targetGrid.DataSource)
        if ($null -eq $items -or $items.Count -le 1) { return }

        $sorted = $items | Sort-Object -Property @{
            Expression = {
                $val = $_.$propName
                if ($null -eq $val -or $val -eq "" -or $val -eq "-") { return "" }
                if ($val -as [int]) { return [int]$val }
                return $val
            }
            Descending = (-not $state.Asc)
        }

        $targetGrid.SuspendLayout()
        $arr = [System.Collections.ArrayList]::new()
        foreach ($item in $sorted) { [void]$arr.Add($item) }
        $targetGrid.DataSource = $null
        $targetGrid.DataSource = $arr
        $targetGrid.ClearSelection()

        foreach ($c in $targetGrid.Columns) {
            $c.HeaderCell.SortGlyphDirection = [System.Windows.Forms.SortOrder]::None
        }
        $targetGrid.Columns[$e.ColumnIndex].HeaderCell.SortGlyphDirection = `
            $(if ($state.Asc) { [System.Windows.Forms.SortOrder]::Ascending } else { [System.Windows.Forms.SortOrder]::Descending })

        $targetGrid.ResumeLayout()
    })
}

# ------------------------------------------------------------------------------
# HAUPTFORMULAR
# ------------------------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Tool 24: AD Computer RDP & lokalAdmin Gruppen-Manager - Domäne: $domainName"
$form.Size = New-Object System.Drawing.Size(1660, 940)
$form.MinimumSize = New-Object System.Drawing.Size(1280, 750)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = [System.Windows.Forms.DockStyle]::Fill
$form.Controls.Add($tabControl)

# ==============================================================================
# REGISTER 1: COMPUTER & GRUPPENPRÜFUNG
# ==============================================================================
$tabComp = New-Object System.Windows.Forms.TabPage
$tabComp.Text = "Computer & Gruppenprüfung"
$tabComp.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
$tabControl.TabPages.Add($tabComp)

# CONFIG PANEL
$pnlTop = New-Object System.Windows.Forms.GroupBox
$pnlTop.Text = "Konfiguration, OU-Zielauswahl, Filter & Backup-Pfad"
$pnlTop.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlTop.Height = 145
$pnlTop.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$tabComp.Controls.Add($pnlTop)

# Zeile 1: Ziel-OUs
$lblOuRdp = New-Object System.Windows.Forms.Label
$lblOuRdp.Text = "Ziel-OU RDP:"
$lblOuRdp.Location = New-Object System.Drawing.Point(15, 22)
$lblOuRdp.AutoSize = $true
$lblOuRdp.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$pnlTop.Controls.Add($lblOuRdp)

$cmbOuRdp = New-Object System.Windows.Forms.ComboBox
$cmbOuRdp.Location = New-Object System.Drawing.Point(120, 19)
$cmbOuRdp.Size = New-Object System.Drawing.Size(560, 23)
$cmbOuRdp.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbOuRdp.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$pnlTop.Controls.Add($cmbOuRdp)

$lblOuAdmin = New-Object System.Windows.Forms.Label
$lblOuAdmin.Text = "Ziel-OU lokalAdmin:"
$lblOuAdmin.Location = New-Object System.Drawing.Point(700, 22)
$lblOuAdmin.AutoSize = $true
$lblOuAdmin.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$pnlTop.Controls.Add($lblOuAdmin)

$cmbOuAdmin = New-Object System.Windows.Forms.ComboBox
$cmbOuAdmin.Location = New-Object System.Drawing.Point(835, 19)
$cmbOuAdmin.Size = New-Object System.Drawing.Size(560, 23)
$cmbOuAdmin.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbOuAdmin.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$pnlTop.Controls.Add($cmbOuAdmin)

# Zeile 2: Filterleiste
$lblFilter = New-Object System.Windows.Forms.Label
$lblFilter.Text = "Name (*):"
$lblFilter.Location = New-Object System.Drawing.Point(15, 58)
$lblFilter.AutoSize = $true
$lblFilter.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$pnlTop.Controls.Add($lblFilter)

$txtCompFilter = New-Object System.Windows.Forms.TextBox
$txtCompFilter.Location = New-Object System.Drawing.Point(75, 55)
$txtCompFilter.Size = New-Object System.Drawing.Size(120, 23)
$txtCompFilter.Text = "*"
$txtCompFilter.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$pnlTop.Controls.Add($txtCompFilter)

$lblTypeFilter = New-Object System.Windows.Forms.Label
$lblTypeFilter.Text = "Typ:"
$lblTypeFilter.Location = New-Object System.Drawing.Point(210, 58)
$lblTypeFilter.AutoSize = $true
$lblTypeFilter.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$pnlTop.Controls.Add($lblTypeFilter)

$cmbTypeFilter = New-Object System.Windows.Forms.ComboBox
$cmbTypeFilter.Location = New-Object System.Drawing.Point(245, 55)
$cmbTypeFilter.Size = New-Object System.Drawing.Size(155, 23)
$cmbTypeFilter.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbTypeFilter.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
[void]$cmbTypeFilter.Items.AddRange(@("Alle (PC & Server)", "Nur Computer (Clients)", "Nur Server"))
$cmbTypeFilter.SelectedIndex = 0
$pnlTop.Controls.Add($cmbTypeFilter)

$lblActiveFilter = New-Object System.Windows.Forms.Label
$lblActiveFilter.Text = "PC-Status:"
$lblActiveFilter.Location = New-Object System.Drawing.Point(415, 58)
$lblActiveFilter.AutoSize = $true
$lblActiveFilter.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$pnlTop.Controls.Add($lblActiveFilter)

$cmbActiveFilter = New-Object System.Windows.Forms.ComboBox
$cmbActiveFilter.Location = New-Object System.Drawing.Point(485, 55)
$cmbActiveFilter.Size = New-Object System.Drawing.Size(165, 23)
$cmbActiveFilter.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbActiveFilter.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
[void]$cmbActiveFilter.Items.AddRange(@("Nur Aktive", "Nur Inaktiv (Deaktiviert)", "Alle Konten"))
$cmbActiveFilter.SelectedIndex = 0
$pnlTop.Controls.Add($cmbActiveFilter)

$lblStatusFilter = New-Object System.Windows.Forms.Label
$lblStatusFilter.Text = "Bedarf:"
$lblStatusFilter.Location = New-Object System.Drawing.Point(665, 58)
$lblStatusFilter.AutoSize = $true
$lblStatusFilter.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$pnlTop.Controls.Add($lblStatusFilter)

$cmbStatusFilter = New-Object System.Windows.Forms.ComboBox
$cmbStatusFilter.Location = New-Object System.Drawing.Point(720, 55)
$cmbStatusFilter.Size = New-Object System.Drawing.Size(190, 23)
$cmbStatusFilter.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbStatusFilter.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
[void]$cmbStatusFilter.Items.AddRange(@("Alle anzeigen", "Nur mit Handlungsbedarf", "Nur Abgeschlossene (OK)"))
$cmbStatusFilter.SelectedIndex = 0
$pnlTop.Controls.Add($cmbStatusFilter)

$btnScan = New-Object System.Windows.Forms.Button
$btnScan.Text = "AD prüfen"
$btnScan.Location = New-Object System.Drawing.Point(920, 52)
$btnScan.Size = New-Object System.Drawing.Size(105, 28)
$btnScan.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$btnScan.ForeColor = [System.Drawing.Color]::White
$btnScan.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$pnlTop.Controls.Add($btnScan)

$btnExportTab1 = New-Object System.Windows.Forms.Button
$btnExportTab1.Text = "Tabelle Export"
$btnExportTab1.Location = New-Object System.Drawing.Point(1030, 52)
$btnExportTab1.Size = New-Object System.Drawing.Size(110, 28)
$btnExportTab1.BackColor = [System.Drawing.Color]::FromArgb(16, 124, 65)
$btnExportTab1.ForeColor = [System.Drawing.Color]::White
$btnExportTab1.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$pnlTop.Controls.Add($btnExportTab1)

$btnExportMembersTab1 = New-Object System.Windows.Forms.Button
$btnExportMembersTab1.Text = "👥 Mitglieder exportieren"
$btnExportMembersTab1.Location = New-Object System.Drawing.Point(1145, 52)
$btnExportMembersTab1.Size = New-Object System.Drawing.Size(170, 28)
$btnExportMembersTab1.BackColor = [System.Drawing.Color]::FromArgb(40, 100, 170)
$btnExportMembersTab1.ForeColor = [System.Drawing.Color]::White
$btnExportMembersTab1.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnExportMembersTab1.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$pnlTop.Controls.Add($btnExportMembersTab1)

# Zeile 3: Backup-Pfad
$lblBackupPath = New-Object System.Windows.Forms.Label
$lblBackupPath.Text = "Backup- & Log-Pfad:"
$lblBackupPath.Location = New-Object System.Drawing.Point(15, 98)
$lblBackupPath.AutoSize = $true
$lblBackupPath.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$lblBackupPath.ForeColor = [System.Drawing.Color]::FromArgb(120, 40, 20)
$pnlTop.Controls.Add($lblBackupPath)

$txtBackupDir = New-Object System.Windows.Forms.TextBox
$txtBackupDir.Location = New-Object System.Drawing.Point(150, 95)
$txtBackupDir.Size = New-Object System.Drawing.Size(530, 23)
$txtBackupDir.Text = "C:\Install\Backup\AD_Groups"
$txtBackupDir.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$pnlTop.Controls.Add($txtBackupDir)

$btnBrowseBackup = New-Object System.Windows.Forms.Button
$btnBrowseBackup.Text = "Pfad wählen..."
$btnBrowseBackup.Location = New-Object System.Drawing.Point(690, 93)
$btnBrowseBackup.Size = New-Object System.Drawing.Size(120, 26)
$btnBrowseBackup.BackColor = [System.Drawing.Color]::FromArgb(235, 240, 245)
$btnBrowseBackup.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnBrowseBackup.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$pnlTop.Controls.Add($btnBrowseBackup)

# Status-Banner
$pnlStatus = New-Object System.Windows.Forms.Panel
$pnlStatus.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlStatus.Height = 28
$pnlStatus.BackColor = [System.Drawing.Color]::FromArgb(235, 240, 248)
$tabComp.Controls.Add($pnlStatus)
$pnlTop.SendToBack()

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Bereit. Ziel-OUs prüfen und AD-Prüfung starten."
$lblStatus.Dock = [System.Windows.Forms.DockStyle]::Fill
$lblStatus.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$lblStatus.Padding = New-Object System.Windows.Forms.Padding(15, 0, 0, 0)
$lblStatus.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(20, 60, 120)
$pnlStatus.Controls.Add($lblStatus)

# Split-Container
$splitContainer = New-Object System.Windows.Forms.SplitContainer
$splitContainer.Dock = [System.Windows.Forms.DockStyle]::Fill
$splitContainer.Orientation = [System.Windows.Forms.Orientation]::Horizontal
$splitContainer.SplitterDistance = 450
$tabComp.Controls.Add($splitContainer)
$splitContainer.BringToFront()

$grpGrid = New-Object System.Windows.Forms.GroupBox
$grpGrid.Text = "Gefundene AD-Objekte (Klick auf Spaltenkopf sortiert | Zeilen markieren = Befehle/Export für Auswahl)"
$grpGrid.Dock = [System.Windows.Forms.DockStyle]::Fill
$grpGrid.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$splitContainer.Panel1.Controls.Add($grpGrid)

$gridData = New-Object System.Windows.Forms.DataGridView
$gridData.Dock = [System.Windows.Forms.DockStyle]::Fill
$gridData.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::DisplayedCells
Apply-StandardGridTheme $gridData
Enable-UniversalGridSorting -Grid $gridData
$grpGrid.Controls.Add($gridData)

$grpCmd = New-Object System.Windows.Forms.GroupBox
$grpCmd.Text = "Generierte PowerShell-Befehle"
$grpCmd.Dock = [System.Windows.Forms.DockStyle]::Fill
$grpCmd.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$splitContainer.Panel2.Controls.Add($grpCmd)

$txtCommands = New-Object System.Windows.Forms.TextBox
$txtCommands.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtCommands.Multiline = $true
$txtCommands.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
$txtCommands.Font = New-Object System.Drawing.Font("Consolas", 9.0)
$txtCommands.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
$txtCommands.ForeColor = [System.Drawing.Color]::FromArgb(220, 220, 220)
$txtCommands.ReadOnly = $true
$grpCmd.Controls.Add($txtCommands)

$pnlCmdButtons = New-Object System.Windows.Forms.Panel
$pnlCmdButtons.Dock = [System.Windows.Forms.DockStyle]::Right
$pnlCmdButtons.Width = 270
$pnlCmdButtons.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
$grpCmd.Controls.Add($pnlCmdButtons)

$btnCopyCmd = New-Object System.Windows.Forms.Button
$btnCopyCmd.Text = "Befehle kopieren"
$btnCopyCmd.Location = New-Object System.Drawing.Point(15, 10)
$btnCopyCmd.Size = New-Object System.Drawing.Size(240, 30)
$btnCopyCmd.BackColor = [System.Drawing.Color]::FromArgb(40, 100, 170)
$btnCopyCmd.ForeColor = [System.Drawing.Color]::White
$btnCopyCmd.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$pnlCmdButtons.Controls.Add($btnCopyCmd)

$btnCreateMissing = New-Object System.Windows.Forms.Button
$btnCreateMissing.Text = "Gruppen anlegen (0)"
$btnCreateMissing.Location = New-Object System.Drawing.Point(15, 46)
$btnCreateMissing.Size = New-Object System.Drawing.Size(240, 32)
$btnCreateMissing.BackColor = [System.Drawing.Color]::FromArgb(16, 124, 65)
$btnCreateMissing.ForeColor = [System.Drawing.Color]::White
$btnCreateMissing.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnCreateMissing.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$pnlCmdButtons.Controls.Add($btnCreateMissing)

$btnDeleteOrphaned = New-Object System.Windows.Forms.Button
$btnDeleteOrphaned.Text = "Backup & Inaktive bereinigen (0)"
$btnDeleteOrphaned.Location = New-Object System.Drawing.Point(15, 84)
$btnDeleteOrphaned.Size = New-Object System.Drawing.Size(240, 36)
$btnDeleteOrphaned.BackColor = [System.Drawing.Color]::FromArgb(180, 40, 40)
$btnDeleteOrphaned.ForeColor = [System.Drawing.Color]::White
$btnDeleteOrphaned.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnDeleteOrphaned.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$pnlCmdButtons.Controls.Add($btnDeleteOrphaned)

# ==============================================================================
# REGISTER 2: GEGENPRÜFUNG VERWAISTER GRUPPEN
# ==============================================================================
$tabOrphan = New-Object System.Windows.Forms.TabPage
$tabOrphan.Text = "Gegenprüfung: Verwaiste AD-Gruppen"
$tabOrphan.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
$tabControl.TabPages.Add($tabOrphan)

$pnlOrphanTop = New-Object System.Windows.Forms.Panel
$pnlOrphanTop.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlOrphanTop.Height = 55
$pnlOrphanTop.BackColor = [System.Drawing.Color]::FromArgb(240, 243, 246)
$tabOrphan.Controls.Add($pnlOrphanTop)

$btnScanOrphans = New-Object System.Windows.Forms.Button
$btnScanOrphans.Text = "Verwaiste Gruppen suchen"
$btnScanOrphans.Location = New-Object System.Drawing.Point(15, 12)
$btnScanOrphans.Size = New-Object System.Drawing.Size(200, 30)
$btnScanOrphans.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$btnScanOrphans.ForeColor = [System.Drawing.Color]::White
$btnScanOrphans.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnScanOrphans.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$pnlOrphanTop.Controls.Add($btnScanOrphans)

$btnExportOrphans = New-Object System.Windows.Forms.Button
$btnExportOrphans.Text = "CSV-Backup in Zielordner"
$btnExportOrphans.Location = New-Object System.Drawing.Point(225, 12)
$btnExportOrphans.Size = New-Object System.Drawing.Size(210, 30)
$btnExportOrphans.BackColor = [System.Drawing.Color]::FromArgb(16, 124, 65)
$btnExportOrphans.ForeColor = [System.Drawing.Color]::White
$btnExportOrphans.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnExportOrphans.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$pnlOrphanTop.Controls.Add($btnExportOrphans)

$btnDeleteOrphanList = New-Object System.Windows.Forms.Button
$btnDeleteOrphanList.Text = "Backup & Verwaiste löschen"
$btnDeleteOrphanList.Location = New-Object System.Drawing.Point(445, 12)
$btnDeleteOrphanList.Size = New-Object System.Drawing.Size(240, 30)
$btnDeleteOrphanList.BackColor = [System.Drawing.Color]::FromArgb(180, 40, 40)
$btnDeleteOrphanList.ForeColor = [System.Drawing.Color]::White
$btnDeleteOrphanList.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnDeleteOrphanList.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$pnlOrphanTop.Controls.Add($btnDeleteOrphanList)

$lblOrphanCount = New-Object System.Windows.Forms.Label
$lblOrphanCount.Location = New-Object System.Drawing.Point(700, 18)
$lblOrphanCount.AutoSize = $true
$lblOrphanCount.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$lblOrphanCount.ForeColor = [System.Drawing.Color]::DarkRed
$pnlOrphanTop.Controls.Add($lblOrphanCount)

$gridOrphans = New-Object System.Windows.Forms.DataGridView
$gridOrphans.Dock = [System.Windows.Forms.DockStyle]::Fill
$gridOrphans.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::DisplayedCells
Apply-StandardGridTheme $gridOrphans
Enable-UniversalGridSorting -Grid $gridOrphans
$tabOrphan.Controls.Add($gridOrphans)
$gridOrphans.BringToFront()

# ==============================================================================
# REGISTER 3: GRUPPENMITGLIEDER & BEREINIGUNG
# ==============================================================================
$tabMembers = New-Object System.Windows.Forms.TabPage
$tabMembers.Text = "👥 Gruppenmitglieder & Bereinigung"
$tabMembers.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
$tabControl.TabPages.Add($tabMembers)

# Top-Leiste Register 3 mit Typ- & Status-Filtern
$pnlMembersTop = New-Object System.Windows.Forms.Panel
$pnlMembersTop.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlMembersTop.Height = 55
$pnlMembersTop.BackColor = [System.Drawing.Color]::FromArgb(240, 243, 246)
$tabMembers.Controls.Add($pnlMembersTop)

$btnLoadMembersTab3 = New-Object System.Windows.Forms.Button
$btnLoadMembersTab3.Text = "🔄 Mitglieder einlesen"
$btnLoadMembersTab3.Location = New-Object System.Drawing.Point(12, 12)
$btnLoadMembersTab3.Size = New-Object System.Drawing.Size(160, 30)
$btnLoadMembersTab3.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$btnLoadMembersTab3.ForeColor = [System.Drawing.Color]::White
$btnLoadMembersTab3.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnLoadMembersTab3.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$pnlMembersTop.Controls.Add($btnLoadMembersTab3)

# Filter Typ
$lblMbrType = New-Object System.Windows.Forms.Label
$lblMbrType.Text = "Typ:"
$lblMbrType.Location = New-Object System.Drawing.Point(180, 18)
$lblMbrType.AutoSize = $true
$lblMbrType.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$pnlMembersTop.Controls.Add($lblMbrType)

$cmbMbrTypeFilter = New-Object System.Windows.Forms.ComboBox
$cmbMbrTypeFilter.Location = New-Object System.Drawing.Point(215, 15)
$cmbMbrTypeFilter.Size = New-Object System.Drawing.Size(145, 23)
$cmbMbrTypeFilter.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbMbrTypeFilter.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
[void]$cmbMbrTypeFilter.Items.AddRange(@("Alle (PC & Server)", "Nur Computer (Clients)", "Nur Server"))
$cmbMbrTypeFilter.SelectedIndex = 0
$pnlMembersTop.Controls.Add($cmbMbrTypeFilter)

# Filter PC-Status
$lblMbrStatus = New-Object System.Windows.Forms.Label
$lblMbrStatus.Text = "Status:"
$lblMbrStatus.Location = New-Object System.Drawing.Point(370, 18)
$lblMbrStatus.AutoSize = $true
$lblMbrStatus.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$pnlMembersTop.Controls.Add($lblMbrStatus)

$cmbMbrStatusFilter = New-Object System.Windows.Forms.ComboBox
$cmbMbrStatusFilter.Location = New-Object System.Drawing.Point(420, 15)
$cmbMbrStatusFilter.Size = New-Object System.Drawing.Size(155, 23)
$cmbMbrStatusFilter.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbMbrStatusFilter.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
[void]$cmbMbrStatusFilter.Items.AddRange(@("Alle Konten", "Nur Aktive", "Nur Inaktiv (Deaktiviert)"))
$cmbMbrStatusFilter.SelectedIndex = 0
$pnlMembersTop.Controls.Add($cmbMbrStatusFilter)

# Textfilter
$lblMemberFilter = New-Object System.Windows.Forms.Label
$lblMemberFilter.Text = "Suche:"
$lblMemberFilter.Location = New-Object System.Drawing.Point(585, 18)
$lblMemberFilter.AutoSize = $true
$lblMemberFilter.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$pnlMembersTop.Controls.Add($lblMemberFilter)

$txtMemberFilter = New-Object System.Windows.Forms.TextBox
$txtMemberFilter.Location = New-Object System.Drawing.Point(635, 15)
$txtMemberFilter.Size = New-Object System.Drawing.Size(160, 23)
$txtMemberFilter.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$pnlMembersTop.Controls.Add($txtMemberFilter)

$btnExportTab3Csv = New-Object System.Windows.Forms.Button
$btnExportTab3Csv.Text = "📥 CSV Export"
$btnExportTab3Csv.Location = New-Object System.Drawing.Point(805, 12)
$btnExportTab3Csv.Size = New-Object System.Drawing.Size(110, 30)
$btnExportTab3Csv.BackColor = [System.Drawing.Color]::FromArgb(16, 124, 65)
$btnExportTab3Csv.ForeColor = [System.Drawing.Color]::White
$btnExportTab3Csv.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnExportTab3Csv.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$pnlMembersTop.Controls.Add($btnExportTab3Csv)

$lblMembersCount = New-Object System.Windows.Forms.Label
$lblMembersCount.Location = New-Object System.Drawing.Point(925, 18)
$lblMembersCount.AutoSize = $true
$lblMembersCount.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$lblMembersCount.ForeColor = [System.Drawing.Color]::FromArgb(20, 60, 120)
$pnlMembersTop.Controls.Add($lblMembersCount)

# Split-Container Register 3
$splitMembers = New-Object System.Windows.Forms.SplitContainer
$splitMembers.Dock = [System.Windows.Forms.DockStyle]::Fill
$splitMembers.Orientation = [System.Windows.Forms.Orientation]::Horizontal
$splitMembers.SplitterDistance = 470
$tabMembers.Controls.Add($splitMembers)
$splitMembers.BringToFront()

$grpMembersGrid = New-Object System.Windows.Forms.GroupBox
$grpMembersGrid.Text = "👥 Gruppenmitglieder (Klick auf Spaltenkopf sortiert | Zeilen markieren = Befehle für Auswahl)"
$grpMembersGrid.Dock = [System.Windows.Forms.DockStyle]::Fill
$grpMembersGrid.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$splitMembers.Panel1.Controls.Add($grpMembersGrid)

$gridMembers = New-Object System.Windows.Forms.DataGridView
$gridMembers.Dock = [System.Windows.Forms.DockStyle]::Fill
$gridMembers.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::DisplayedCells
Apply-StandardGridTheme $gridMembers
Enable-UniversalGridSorting -Grid $gridMembers
$grpMembersGrid.Controls.Add($gridMembers)

$grpMembersCmd = New-Object System.Windows.Forms.GroupBox
$grpMembersCmd.Text = "⚡ PowerShell-Befehle zur Mitglieder-Bereinigung (Remove-ADGroupMember)"
$grpMembersCmd.Dock = [System.Windows.Forms.DockStyle]::Fill
$grpMembersCmd.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$splitMembers.Panel2.Controls.Add($grpMembersCmd)

$txtMembersCommands = New-Object System.Windows.Forms.TextBox
$txtMembersCommands.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtMembersCommands.Multiline = $true
$txtMembersCommands.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
$txtMembersCommands.Font = New-Object System.Drawing.Font("Consolas", 9.0)
$txtMembersCommands.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
$txtMembersCommands.ForeColor = [System.Drawing.Color]::FromArgb(220, 220, 220)
$txtMembersCommands.ReadOnly = $true
$grpMembersCmd.Controls.Add($txtMembersCommands)

$pnlMembersCmdButtons = New-Object System.Windows.Forms.Panel
$pnlMembersCmdButtons.Dock = [System.Windows.Forms.DockStyle]::Right
$pnlMembersCmdButtons.Width = 270
$pnlMembersCmdButtons.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
$grpMembersCmd.Controls.Add($pnlMembersCmdButtons)

$btnCopyMemberCmd = New-Object System.Windows.Forms.Button
$btnCopyMemberCmd.Text = "Befehle kopieren"
$btnCopyMemberCmd.Location = New-Object System.Drawing.Point(15, 10)
$btnCopyMemberCmd.Size = New-Object System.Drawing.Size(240, 30)
$btnCopyMemberCmd.BackColor = [System.Drawing.Color]::FromArgb(40, 100, 170)
$btnCopyMemberCmd.ForeColor = [System.Drawing.Color]::White
$btnCopyMemberCmd.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$pnlMembersCmdButtons.Controls.Add($btnCopyMemberCmd)

$btnRemoveSelectedMembers = New-Object System.Windows.Forms.Button
$btnRemoveSelectedMembers.Text = "Auswahl entfernen (0)"
$btnRemoveSelectedMembers.Location = New-Object System.Drawing.Point(15, 46)
$btnRemoveSelectedMembers.Size = New-Object System.Drawing.Size(240, 34)
$btnRemoveSelectedMembers.BackColor = [System.Drawing.Color]::FromArgb(180, 40, 40)
$btnRemoveSelectedMembers.ForeColor = [System.Drawing.Color]::White
$btnRemoveSelectedMembers.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnRemoveSelectedMembers.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$pnlMembersCmdButtons.Controls.Add($btnRemoveSelectedMembers)

$btnRemoveAllListedMembers = New-Object System.Windows.Forms.Button
$btnRemoveAllListedMembers.Text = "Alle Gelisteten entfernen (0)"
$btnRemoveAllListedMembers.Location = New-Object System.Drawing.Point(15, 85)
$btnRemoveAllListedMembers.Size = New-Object System.Drawing.Size(240, 34)
$btnRemoveAllListedMembers.BackColor = [System.Drawing.Color]::FromArgb(150, 20, 20)
$btnRemoveAllListedMembers.ForeColor = [System.Drawing.Color]::White
$btnRemoveAllListedMembers.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnRemoveAllListedMembers.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$pnlMembersCmdButtons.Controls.Add($btnRemoveAllListedMembers)

$btnBrowseBackup.Add_Click({
    $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
    $fbd.Description = "Wählen Sie das Zielverzeichnis für Backups und Logs:"
    $fbd.SelectedPath = $txtBackupDir.Text.Trim()
    if ($fbd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtBackupDir.Text = $fbd.SelectedPath
    }
})

# ------------------------------------------------------------------------------
# ZENTRALE AUDIT-LOG-FUNKTION
# ------------------------------------------------------------------------------
function Write-ToolLog {
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS")][string]$Severity = "INFO"
    )

    $targetPath = $txtBackupDir.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($targetPath)) { $targetPath = "C:\Install\Backup\AD_Groups" }
    if (-not (Test-Path -LiteralPath $targetPath)) {
        try { New-Item -ItemType Directory -Path $targetPath -Force | Out-Null } catch {}
    }

    $logFile = Join-Path -Path $targetPath -ChildPath "Activity_Log_$(Get-Date -Format 'yyyyMMdd').log"
    $timeStr = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $userStr = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    $logLine = "[$timeStr] [$Severity] [$userStr] $Message"

    try {
        Add-Content -Path $logFile -Value $logLine -Encoding UTF8
    } catch {}
}

# ------------------------------------------------------------------------------
# OU-LADEN & VORAUSWAHL
# ------------------------------------------------------------------------------
function Load-DomainOUsAndPreselect {
    $rootAdsi = $null
    $ouSearcher = $null
    $ous = $null
    try {
        $rootAdsi = [ADSI]"LDAP://$domainDN"
        $ouSearcher = [System.DirectoryServices.DirectorySearcher]::new($rootAdsi)
        $ouSearcher.Filter = "(objectCategory=organizationalUnit)"
        $ouSearcher.PropertiesToLoad.AddRange(@("distinguishedName", "name"))
        $ouSearcher.PageSize = 1000
        $ous = $ouSearcher.FindAll()

        $ouList = [System.Collections.Generic.List[string]]::new()
        foreach ($ou in $ous) {
            $dn = [string]$ou.Properties["distinguishedname"][0]
            $ouList.Add($dn)
        }
        $ouList.Sort()
        $ouList.Insert(0, "CN=Users,$domainDN")

        foreach ($dn in $ouList) {
            [void]$cmbOuRdp.Items.Add($dn)
            [void]$cmbOuAdmin.Items.Add($dn)
        }

        $rdpIndex = -1
        for ($i = 0; $i -lt $ouList.Count; $i++) {
            $ouRdn = ($ouList[$i] -split ',')[0]
            if ($ouRdn -match "RDP") { $rdpIndex = $i; break }
        }
        if ($rdpIndex -ge 0) { $cmbOuRdp.SelectedIndex = $rdpIndex }
        elseif ($cmbOuRdp.Items.Count -gt 0) { $cmbOuRdp.SelectedIndex = 0 }

        $adminIndex = -1
        for ($i = 0; $i -lt $ouList.Count; $i++) {
            $ouRdn = ($ouList[$i] -split ',')[0]
            if ($ouRdn -match "lokalAdmin|localAdmin") { $adminIndex = $i; break }
        }
        if ($adminIndex -ge 0) { $cmbOuAdmin.SelectedIndex = $adminIndex }
        elseif ($cmbOuAdmin.Items.Count -gt 0) { $cmbOuAdmin.SelectedIndex = 0 }
    } catch {
        [void]$cmbOuRdp.Items.Add("CN=Users,$domainDN")
        [void]$cmbOuAdmin.Items.Add("CN=Users,$domainDN")
        $cmbOuRdp.SelectedIndex = 0
        $cmbOuAdmin.SelectedIndex = 0
    } finally {
        if ($null -ne $ous) { $ous.Dispose() }
        if ($null -ne $ouSearcher) { $ouSearcher.Dispose() }
        if ($null -ne $rootAdsi) { $rootAdsi.Dispose() }
    }
}

# ------------------------------------------------------------------------------
# ZENTRALE FUNKTION: MITGLIEDER EINER GRUPPE AUSLESEN
# ------------------------------------------------------------------------------
function Get-GroupMembersDetail {
    param(
        [string]$GroupName,
        [string]$GroupDN,
        [string]$ComputerName = "-",
        [string]$ComputerDescription = "-",
        [string]$ComputerStatus = "Aktiviert",
        [string]$ComputerType = "Computer"
    )

    $curDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $rows = [System.Collections.Generic.List[PSCustomObject]]::new()

    try {
        $gEntry = [ADSI]"LDAP://$GroupDN"
        $members = @($gEntry.Properties["member"])
        $gEntry.Dispose()

        if ($members -and $members.Count -gt 0) {
            foreach ($mDN in $members) {
                $mName = ""
                $mSam  = ""
                $mType = "Unbekannt"
                $mEntry = $null
                try {
                    $mEntry = [ADSI]"LDAP://$mDN"
                    $mName = [string]$mEntry.Properties["name"].Value
                    $mSam  = [string]$mEntry.Properties["samaccountname"].Value
                    $classes = @($mEntry.Properties["objectClass"])
                    if ($classes -contains "user") { $mType = "Benutzer" }
                    elseif ($classes -contains "group") { $mType = "Gruppe" }
                    elseif ($classes -contains "computer") { $mType = "Computer" }
                } catch { 
                    $mName = "Nicht auflösbar" 
                } finally {
                    if ($null -ne $mEntry) { $mEntry.Dispose() }
                }

                $rows.Add([PSCustomObject]@{
                    "Computer"              = $ComputerName
                    "Computer-Status"       = $ComputerStatus
                    "Typ"                   = $ComputerType
                    "Computer-Beschreibung" = $ComputerDescription
                    "Gruppenname"           = $GroupName
                    "Mitglied Typ"          = $mType
                    "Mitgliedsname"         = $mName
                    "SamAccountName"        = $mSam
                    "Mitglieds-DN"          = $mDN
                    "Gruppen-DN"            = $GroupDN
                    "Domain Name"           = $domainName
                    "Export-Zeitpunkt"      = $curDate
                })
            }
        } else {
            $rows.Add([PSCustomObject]@{
                "Computer"              = $ComputerName
                "Computer-Status"       = $ComputerStatus
                "Typ"                   = $ComputerType
                "Computer-Beschreibung" = $ComputerDescription
                "Gruppenname"           = $GroupName
                "Mitglied Typ"          = "-"
                "Mitgliedsname"         = "(Keine Mitglieder)"
                "SamAccountName"        = "-"
                "Mitglieds-DN"          = "-"
                "Gruppen-DN"            = $GroupDN
                "Domain Name"           = $domainName
                "Export-Zeitpunkt"      = $curDate
            })
        }
    } catch {
        Write-ToolLog -Message "Fehler beim Lesen der Mitglieder fuer Gruppe '$GroupName': $($_.Exception.Message)" -Severity "WARN"
    }

    return $rows
}

# ------------------------------------------------------------------------------
# DEFINITION DER SCRIPTBLÖCKE FÜR REGISTER 1 & 3
# ------------------------------------------------------------------------------

# 1. Befehle für Register 1 generieren
$updateCommandsAction = {
    $targetOuRdp   = if ($cmbOuRdp.SelectedItem) { $cmbOuRdp.SelectedItem.ToString() } else { "CN=Users,$domainDN" }
    $targetOuAdmin = if ($cmbOuAdmin.SelectedItem) { $cmbOuAdmin.SelectedItem.ToString() } else { "CN=Users,$domainDN" }

    $isUserSelection = $false
    $targetRows = @()

    if ($gridData.SelectedRows.Count -gt 0) {
        $isUserSelection = $true
        foreach ($r in $gridData.SelectedRows) {
            if ($r.DataBoundItem) {
                $targetRows += $r.DataBoundItem
            }
        }
    } else {
        $targetRows = @($script:CurrentFilteredRows)
    }

    $script:PendingCreateCmds.Clear()
    $script:PendingDeleteGroups.Clear()
    $sbCommands = New-Object System.Text.StringBuilder

    [void]$sbCommands.AppendLine("# ===========================================================================")
    if ($isUserSelection) {
        [void]$sbCommands.AppendLine("# BEFEHLE FÜR MARKIERTEN BEREICH ($($targetRows.Count) markierte Zeile(n))")
    } else {
        [void]$sbCommands.AppendLine("# BEFEHLE FÜR AKTUELLE FILTERANSICHT ($($targetRows.Count) angezeigte Objekte)")
    }
    [void]$sbCommands.AppendLine("# ===========================================================================")

    foreach ($row in $targetRows) {
        $cName      = $row.Computername
        $cDesc      = $row._RawDesc
        $isDisabled = $row._IsDisabled
        $hasRdp     = $row._HasRdp
        $hasAdmin   = $row._HasAdmin
        $expectedRdpGrp   = "$($cName)-RDP"
        $expectedAdminGrp = "$($cName)-lokalAdmin"
        $descBase = if (-not [string]::IsNullOrWhiteSpace($cDesc)) { $cDesc } else { $cName }

        if (-not $isDisabled) {
            if (-not $hasRdp) {
                $descRdp = "Remote Desktop Berechtigung fuer $descBase"
                $cmd = "New-ADGroup -Name '$expectedRdpGrp' -SamAccountName '$expectedRdpGrp' -GroupScope DomainLocal -GroupCategory Security -Path '$targetOuRdp' -Description '$descRdp'"
                $script:PendingCreateCmds.Add($cmd)
                [void]$sbCommands.AppendLine($cmd)
            }
            if (-not $hasAdmin) {
                $descAdmin = "Lokale Administratorrechte fuer $descBase"
                $cmd = "New-ADGroup -Name '$expectedAdminGrp' -SamAccountName '$expectedAdminGrp' -GroupScope DomainLocal -GroupCategory Security -Path '$targetOuAdmin' -Description '$descAdmin'"
                $script:PendingCreateCmds.Add($cmd)
                [void]$sbCommands.AppendLine($cmd)
            }
        } else {
            if ($hasRdp) {
                $script:PendingDeleteGroups.Add([PSCustomObject]@{
                    GroupName           = $expectedRdpGrp
                    ComputerName        = $cName
                    ComputerDescription = $cDesc
                    DistinguishedName   = $row._RdpInfo.DistinguishedName
                    Members             = $row._RdpInfo.Members
                })
                $delCmd = "Remove-ADGroup -Identity '$expectedRdpGrp' -Confirm:`$false"
                [void]$sbCommands.AppendLine("# Inaktiv: $cName ($cDesc) -> Automatisches Backup je Gruppe vor Loeschen aktiv!")
                [void]$sbCommands.AppendLine($delCmd)
            }
            if ($hasAdmin) {
                $script:PendingDeleteGroups.Add([PSCustomObject]@{
                    GroupName           = $expectedAdminGrp
                    ComputerName        = $cName
                    ComputerDescription = $cDesc
                    DistinguishedName   = $row._AdminInfo.DistinguishedName
                    Members             = $row._AdminInfo.Members
                })
                $delCmd = "Remove-ADGroup -Identity '$expectedAdminGrp' -Confirm:`$false"
                [void]$sbCommands.AppendLine("# Inaktiv: $cName ($cDesc) -> Automatisches Backup je Gruppe vor Loeschen aktiv!")
                [void]$sbCommands.AppendLine($delCmd)
            }
        }
    }

    $txtCommands.Text = $sbCommands.ToString()

    if ($isUserSelection) {
        $btnCreateMissing.Text  = "Gruppen für Auswahl anlegen ($($script:PendingCreateCmds.Count))"
        $btnDeleteOrphaned.Text = "Auswahl bereinigen ($($script:PendingDeleteGroups.Count))"
        $grpCmd.Text = "Generierte PowerShell-Befehle (Nur für $($targetRows.Count) markierte Zeilen)"
    } else {
        $btnCreateMissing.Text  = "Gruppen anlegen ($($script:PendingCreateCmds.Count))"
        $btnDeleteOrphaned.Text = "Backup & Inaktive bereinigen ($($script:PendingDeleteGroups.Count))"
        $grpCmd.Text = "Generierte PowerShell-Befehle (Gefilterte Ansicht: $($targetRows.Count) Objekte)"
    }
}

# 2. Filter für Register 1
$filterResultsAction = {
    $typeSel   = $cmbTypeFilter.SelectedIndex
    $activeSel = $cmbActiveFilter.SelectedIndex
    $statusSel = $cmbStatusFilter.SelectedIndex

    $filtered = $script:RawCheckResults | Where-Object {
        $matchType = if ($typeSel -eq 1) { $_.Typ -eq "Computer" }
                     elseif ($typeSel -eq 2) { $_.Typ -eq "Server" }
                     else { $true }

        $matchActive = if ($activeSel -eq 0) { $_."Computer-Status" -eq "Aktiviert" }
                       elseif ($activeSel -eq 1) { $_."Computer-Status" -like "*Deaktiviert*" }
                       else { $true }

        $matchStatus = if ($statusSel -eq 1) { $_."Handlungsbedarf" -eq "Ja" }
                       elseif ($statusSel -eq 2) { $_."Handlungsbedarf" -eq "Nein" }
                       else { $true }

        $matchType -and $matchActive -and $matchStatus
    }

    $script:CurrentFilteredRows.Clear()
    foreach ($f in $filtered) { [void]$script:CurrentFilteredRows.Add($f) }

    $gridData.SuspendLayout()
    $gridData.DataSource = $null
    $arrList = [System.Collections.ArrayList]::new()
    foreach ($item in $filtered) { [void]$arrList.Add($item) }
    $gridData.DataSource = $arrList
    $gridData.ClearSelection()
    $gridData.ResumeLayout()

    & $updateCommandsAction
    $lblStatus.Text = "Gefiltert: $(@($filtered).Count) von $($script:RawCheckResults.Count) Objekten geladen."
}

# 3. AD-Prüfung ausführen
$runCheckAction = {
    $cFilter = $txtCompFilter.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($cFilter)) { $cFilter = "*" }

    $btnScan.Enabled = $false
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $lblStatus.Text = "Frage AD-Computer und vorhandene Gruppen ab..."
    $form.Refresh()

    $script:RawCheckResults.Clear()

    $rootEntry = $null
    $gSearcher = $null
    $allGroups = $null
    $cSearcher = $null
    $computers = $null

    try {
        $rootEntry = [System.DirectoryServices.DirectoryEntry]::new("LDAP://$domainDN")

        $gSearcher = [System.DirectoryServices.DirectorySearcher]::new($rootEntry)
        $gSearcher.Filter = "(objectCategory=group)"
        $gSearcher.PropertiesToLoad.AddRange(@("sAMAccountName", "distinguishedName", "member"))
        $gSearcher.PageSize = 1000
        $allGroups = $gSearcher.FindAll()

        $groupDict = [System.Collections.Generic.Dictionary[string, Object]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($g in $allGroups) {
            $sam = if ($g.Properties["samaccountname"].Count -gt 0) { [string]$g.Properties["samaccountname"][0] } else { "" }
            if ($sam) {
                $groupDict[$sam] = @{
                    DistinguishedName = [string]$g.Properties["distinguishedname"][0]
                    Members           = @($g.Properties["member"])
                }
            }
        }

        $cSearcher = [System.DirectoryServices.DirectorySearcher]::new($rootEntry)
        $cSearcher.Filter = "(&(objectCategory=computer)(name=$cFilter))"
        $cSearcher.PropertiesToLoad.AddRange(@("name", "description", "distinguishedName", "userAccountControl", "operatingSystem"))
        $cSearcher.PageSize = 1000
        $computers = $cSearcher.FindAll()

        $rawList = [System.Collections.Generic.List[PSCustomObject]]::new()

        foreach ($comp in $computers) {
            $cName = [string]$comp.Properties["name"][0]
            $cDesc = if ($comp.Properties["description"].Count -gt 0) { [string]$comp.Properties["description"][0] } else { "" }
            $cDN   = [string]$comp.Properties["distinguishedname"][0]
            $cOS   = if ($comp.Properties["operatingsystem"].Count -gt 0) { [string]$comp.Properties["operatingsystem"][0] } else { "" }
            $uac   = if ($comp.Properties["useraccountcontrol"].Count -gt 0) { [int]$comp.Properties["useraccountcontrol"][0] } else { 0 }

            $isDisabled = ($uac -band 2) -eq 2
            $activeStatus = if ($isDisabled) { "Deaktiviert (Inaktiv)" } else { "Aktiviert" }
            $objType = if ($cOS -match "Server") { "Server" } else { "Computer" }

            $expectedRdpGrp   = "$($cName)-RDP"
            $expectedAdminGrp = "$($cName)-lokalAdmin"

            $hasRdp   = $groupDict.ContainsKey($expectedRdpGrp)
            $hasAdmin = $groupDict.ContainsKey($expectedAdminGrp)

            $actionReq = "Nein"
            $statusText = ""

            if (-not $isDisabled) {
                if ($hasRdp -and $hasAdmin) {
                    $statusText = "[OK] Vollständig"
                } elseif (-not $hasRdp -and -not $hasAdmin) {
                    $statusText = "[x] Beide Gruppen fehlen (Erstellen)"
                    $actionReq = "Ja"
                } elseif (-not $hasRdp) {
                    $statusText = "[!] RDP-Gruppe fehlt (Erstellen)"
                    $actionReq = "Ja"
                } else {
                    $statusText = "[!] lokalAdmin-Gruppe fehlt (Erstellen)"
                    $actionReq = "Ja"
                }
            } else {
                if (-not $hasRdp -and -not $hasAdmin) {
                    $statusText = "Bereinigt (Keine Gruppen vorhanden)"
                } else {
                    $statusText = "[!] Verwaiste Gruppen vorhanden (Löschen)"
                    $actionReq = "Ja"
                }
            }

            $rawList.Add([PSCustomObject]@{
                "Computername"          = $cName
                "Computer-Beschreibung" = if ($cDesc) { $cDesc } else { "-" }
                "Computer-Status"       = $activeStatus
                "Typ"                   = $objType
                "Gruppen-Status"        = $statusText
                "Handlungsbedarf"       = $actionReq
                "RDP-Gruppe"            = if ($hasRdp) { "Vorhanden" } else { "Fehlt" }
                "lokalAdmin-Gruppe"     = if ($hasAdmin) { "Vorhanden" } else { "Fehlt" }
                "Betriebssystem"        = if ($cOS) { $cOS } else { "-" }
                "DistinguishedName"     = $cDN
                "_HasRdp"               = $hasRdp
                "_HasAdmin"             = $hasAdmin
                "_IsDisabled"           = $isDisabled
                "_RawDesc"              = $cDesc
                "_RdpInfo"              = if ($hasRdp) { $groupDict[$expectedRdpGrp] } else { $null }
                "_AdminInfo"            = if ($hasAdmin) { $groupDict[$expectedAdminGrp] } else { $null }
            })
        }

        $sortedComps = $rawList | Sort-Object "Computername"
        foreach ($sc in $sortedComps) { [void]$script:RawCheckResults.Add($sc) }

        Write-ToolLog -Message "AD-Prüfung abgeschlossen. $($computers.Count) Computerobjekte eingelesen." -Severity "INFO"
        & $filterResultsAction
    } catch {
        $lblStatus.Text = "Fehler bei der Abfrage: $($_.Exception.Message)"
        Write-ToolLog -Message "Fehler bei AD-Abfrage: $($_.Exception.Message)" -Severity "ERROR"
        [System.Windows.Forms.MessageBox]::Show("Fehler bei der Abfrage: $($_.Exception.Message)", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } finally {
        if ($null -ne $allGroups) { $allGroups.Dispose() }
        if ($null -ne $gSearcher) { $gSearcher.Dispose() }
        if ($null -ne $computers) { $computers.Dispose() }
        if ($null -ne $cSearcher) { $cSearcher.Dispose() }
        if ($null -ne $rootEntry) { $rootEntry.Dispose() }

        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $btnScan.Enabled = $true
    }
}

# 4. Befehle für Register 3 generieren
$updateMemberCommandsAction = {
    $isUserSelection = $false
    $targetRows = @()

    if ($gridMembers.SelectedRows.Count -gt 0) {
        $isUserSelection = $true
        foreach ($r in $gridMembers.SelectedRows) {
            if ($r.DataBoundItem) { $targetRows += $r.DataBoundItem }
        }
    } else {
        $targetRows = @($script:CurrentMemberRows)
    }

    $script:PendingRemoveMemberCmds.Clear()
    $sbCmds = New-Object System.Text.StringBuilder
    [void]$sbCmds.AppendLine("# ===========================================================================")
    if ($isUserSelection) {
        [void]$sbCmds.AppendLine("# BEFEHLE FÜR MARKIERTEN BEREICH ($($targetRows.Count) markierte Mitglieder)")
    } else {
        [void]$sbCmds.AppendLine("# BEFEHLE FÜR AKTUELLE FILTERANSICHT ($($targetRows.Count) angezeigte Mitglieder)")
    }
    [void]$sbCmds.AppendLine("# ===========================================================================")

    foreach ($row in $targetRows) {
        if ($row.Mitgliedsname -ne "(Keine Mitglieder)" -and -not [string]::IsNullOrWhiteSpace($row.'Mitglieds-DN')) {
            $grpIdentity = $row.Gruppenname
            $memberIdentity = if (-not [string]::IsNullOrWhiteSpace($row.SamAccountName) -and $row.SamAccountName -ne "-") { $row.SamAccountName } else { $row.'Mitglieds-DN' }
            $cmd = "Remove-ADGroupMember -Identity '$grpIdentity' -Members '$memberIdentity' -Confirm:`$false"
            $script:PendingRemoveMemberCmds.Add($cmd)
            [void]$sbCmds.AppendLine($cmd)
        }
    }

    $txtMembersCommands.Text = $sbCmds.ToString()

    if ($isUserSelection) {
        $btnRemoveSelectedMembers.Text = "Auswahl entfernen ($($script:PendingRemoveMemberCmds.Count))"
    } else {
        $btnRemoveSelectedMembers.Text = "Auswahl entfernen (0)"
    }
    
    $validAllCount = ($script:CurrentMemberRows | Where-Object { $_.Mitgliedsname -ne "(Keine Mitglieder)" -and -not [string]::IsNullOrWhiteSpace($_.'Mitglieds-DN') }).Count
    $btnRemoveAllListedMembers.Text = "Alle Gelisteten entfernen ($validAllCount)"
}

# 5. Filter für Register 3
$filterMembersAction = {
    $filterVal = $txtMemberFilter.Text.Trim()
    $typeSel   = $cmbMbrTypeFilter.SelectedIndex
    $statusSel = $cmbMbrStatusFilter.SelectedIndex

    $filtered = $script:AllMembersList | Where-Object {
        $matchType = if ($typeSel -eq 1) { $_.Typ -eq "Computer" }
                     elseif ($typeSel -eq 2) { $_.Typ -eq "Server" }
                     else { $true }

        $matchStatus = if ($statusSel -eq 1) { $_."Computer-Status" -eq "Aktiviert" }
                       elseif ($statusSel -eq 2) { $_."Computer-Status" -like "*Deaktiviert*" }
                       else { $true }

        $matchText = if ([string]::IsNullOrWhiteSpace($filterVal)) {
            $true
        } else {
            $_.Computer -like "*$filterVal*" -or 
            $_.Gruppenname -like "*$filterVal*" -or 
            $_.Mitgliedsname -like "*$filterVal*" -or 
            $_.SamAccountName -like "*$filterVal*"
        }

        $matchType -and $matchStatus -and $matchText
    }

    $script:CurrentMemberRows.Clear()
    foreach ($f in $filtered) { [void]$script:CurrentMemberRows.Add($f) }

    $gridMembers.SuspendLayout()
    $gridMembers.DataSource = $null
    $arrMembers = [System.Collections.ArrayList]::new()
    foreach ($item in $filtered) { [void]$arrMembers.Add($item) }
    $gridMembers.DataSource = $arrMembers
    $gridMembers.ClearSelection()
    $gridMembers.ResumeLayout()

    $lblMembersCount.Text = "Gelistet: $($script:CurrentMemberRows.Count) von $($script:AllMembersList.Count)"
    & $updateMemberCommandsAction
}

# 6. Klick-Event für Register 3 (Mitglieder einlesen)
$loadMembersActionTab3 = {
    $btnLoadMembersTab3.Enabled = $false
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $lblMembersCount.Text = "Lese Mitglieder aus Active Directory ein..."
    $form.Refresh()

    $script:AllMembersList.Clear()

    try {
        if ($script:RawCheckResults.Count -eq 0) {
            & $runCheckAction
        }

        foreach ($row in $script:RawCheckResults) {
            $cName   = [string]$row.Computername
            $cDesc   = [string]$row.'Computer-Beschreibung'
            $cStatus = [string]$row.'Computer-Status'
            $cType   = [string]$row.Typ

            if ($row._HasRdp -and $row._RdpInfo) {
                $mListRdp = Get-GroupMembersDetail -GroupName "$cName-RDP" `
                                                   -GroupDN $row._RdpInfo.DistinguishedName `
                                                   -ComputerName $cName `
                                                   -ComputerDescription $cDesc `
                                                   -ComputerStatus $cStatus `
                                                   -ComputerType $cType
                foreach ($m in $mListRdp) { [void]$script:AllMembersList.Add($m) }
            }
            if ($row._HasAdmin -and $row._AdminInfo) {
                $mListAdmin = Get-GroupMembersDetail -GroupName "$cName-lokalAdmin" `
                                                     -GroupDN $row._AdminInfo.DistinguishedName `
                                                     -ComputerName $cName `
                                                     -ComputerDescription $cDesc `
                                                     -ComputerStatus $cStatus `
                                                     -ComputerType $cType
                foreach ($m in $mListAdmin) { [void]$script:AllMembersList.Add($m) }
            }
        }

        Write-ToolLog -Message "Mitglieder in Register 3 geladen ($($script:AllMembersList.Count) Einträge)." -Severity "INFO"

        if ($script:AllMembersList.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "Es wurden für die gefundenen Computer keine vorhandenen RDP- oder lokalAdmin-Gruppen im Active Directory ermittelt.`r`n`r`nBitte prüfen Sie zuerst in Register 1, ob die Gruppen angelegt sind.",
                "Hinweis",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
    } catch {
        Write-ToolLog -Message "Fehler beim Einlesen der Mitglieder in Tab 3: $($_.Exception.Message)" -Severity "ERROR"
        [System.Windows.Forms.MessageBox]::Show("Fehler beim Einlesen der Mitglieder:`r`n$($_.Exception.Message)", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $btnLoadMembersTab3.Enabled = $true
        & $filterMembersAction
    }
}

# ------------------------------------------------------------------------------
# FARB- & KONTRASTKORREKTUR (TAB 1)
# ------------------------------------------------------------------------------
$gridData.Add_SelectionChanged({
    & $updateCommandsAction
})

$gridData.Add_RowPrePaint({
    param($s, $e)
    if ($e.RowIndex -ge 0 -and $e.RowIndex -lt $gridData.Rows.Count) {
        $row = $gridData.Rows[$e.RowIndex]
        $pcStatus   = [string]$row.Cells["Computer-Status"].Value
        $needAction = [string]$row.Cells["Handlungsbedarf"].Value

        if ($pcStatus -like "*Deaktiviert*") {
            if ($needAction -eq "Ja") {
                $row.DefaultCellStyle.BackColor          = [System.Drawing.Color]::FromArgb(255, 243, 205)
                $row.DefaultCellStyle.ForeColor          = [System.Drawing.Color]::FromArgb(133, 100, 4)
                $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(255, 220, 140)
                $row.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::FromArgb(90, 60, 0)
            } else {
                $row.DefaultCellStyle.BackColor          = [System.Drawing.Color]::FromArgb(245, 245, 245)
                $row.DefaultCellStyle.ForeColor          = [System.Drawing.Color]::FromArgb(120, 120, 120)
                $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(220, 220, 220)
                $row.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::FromArgb(50, 50, 50)
            }
        } else {
            if ($needAction -eq "Ja") {
                $row.DefaultCellStyle.BackColor          = [System.Drawing.Color]::FromArgb(255, 235, 235)
                $row.DefaultCellStyle.ForeColor          = [System.Drawing.Color]::FromArgb(180, 20, 20)
                $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(250, 190, 190)
                $row.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::FromArgb(130, 0, 0)
            } else {
                $row.DefaultCellStyle.BackColor          = [System.Drawing.Color]::FromArgb(240, 253, 244)
                $row.DefaultCellStyle.ForeColor          = [System.Drawing.Color]::FromArgb(20, 100, 40)
                $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(195, 240, 210)
                $row.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::FromArgb(10, 60, 20)
            }
        }
    }
})

$btnExportTab1.Add_Click({
    $exportSource = if ($gridData.SelectedRows.Count -gt 0) {
        $list = [System.Collections.Generic.List[PSCustomObject]]::new()
        foreach ($r in $gridData.SelectedRows) {
            if ($r.DataBoundItem) { [void]$list.Add($r.DataBoundItem) }
        }
        $list
    } else {
        $script:CurrentFilteredRows
    }

    if ($exportSource.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Keine Zeilen zum Exportieren vorhanden.", "Hinweis", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    $targetPath = $txtBackupDir.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($targetPath)) { $targetPath = "C:\Install\Backup\AD_Groups" }
    if (-not (Test-Path -LiteralPath $targetPath)) {
        try { New-Item -ItemType Directory -Path $targetPath -Force | Out-Null } catch {}
    }

    $fileName = "AD_Computer_Gruppen_Status_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $fullPath = Join-Path -Path $targetPath -ChildPath $fileName

    $cleanRows = $exportSource | Select-Object Computername, 'Computer-Beschreibung', Computer-Status, Typ, Gruppen-Status, Handlungsbedarf, RDP-Gruppe, lokalAdmin-Gruppe, Betriebssystem, DistinguishedName

    try {
        $cleanRows | Export-Csv -Path $fullPath -Delimiter ';' -NoTypeInformation -Encoding UTF8
        Write-ToolLog -Message "Tabelle Register 1 exportiert: $fullPath ($($cleanRows.Count) Zeilen)" -Severity "INFO"
        [System.Windows.Forms.MessageBox]::Show("Tabelle erfolgreich exportiert nach:`r`n$fullPath", "Export abgeschlossen", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } catch {
        Write-ToolLog -Message "Fehler beim CSV-Export: $($_.Exception.Message)" -Severity "ERROR"
        [System.Windows.Forms.MessageBox]::Show("Fehler beim CSV-Export: $($_.Exception.Message)", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

$btnExportMembersTab1.Add_Click({
    $exportSource = if ($gridData.SelectedRows.Count -gt 0) {
        $list = [System.Collections.Generic.List[PSCustomObject]]::new()
        foreach ($r in $gridData.SelectedRows) {
            if ($r.DataBoundItem) { [void]$list.Add($r.DataBoundItem) }
        }
        $list
    } else {
        $script:CurrentFilteredRows
    }

    if ($exportSource.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Keine Computer zur Ermittlung von Gruppenmitgliedern vorhanden.", "Hinweis", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    $targetPath = $txtBackupDir.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($targetPath)) { $targetPath = "C:\Install\Backup\AD_Groups" }
    if (-not (Test-Path -LiteralPath $targetPath)) {
        try { New-Item -ItemType Directory -Path $targetPath -Force | Out-Null } catch {}
    }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $allMemberRows = [System.Collections.Generic.List[PSCustomObject]]::new()
    $exportedGroupsCount = 0

    foreach ($row in $exportSource) {
        $cName = $row.Computername
        $cDesc = $row.'Computer-Beschreibung'

        if ($row._HasRdp -and $row._RdpInfo) {
            $mList = Get-GroupMembersDetail -GroupName "$cName-RDP" -GroupDN $row._RdpInfo.DistinguishedName -ComputerName $cName -ComputerDescription $cDesc
            foreach ($m in $mList) { [void]$allMemberRows.Add($m) }
            $exportedGroupsCount++
        }
        if ($row._HasAdmin -and $row._AdminInfo) {
            $mList = Get-GroupMembersDetail -GroupName "$cName-lokalAdmin" -GroupDN $row._AdminInfo.DistinguishedName -ComputerName $cName -ComputerDescription $cDesc
            foreach ($m in $mList) { [void]$allMemberRows.Add($m) }
            $exportedGroupsCount++
        }
    }

    $form.Cursor = [System.Windows.Forms.Cursors]::Default

    if ($allMemberRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Für die ausgewählten Computer existieren noch keine RDP- oder lokalAdmin-Gruppen im Active Directory.", "Hinweis", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    $fileName = "AD_Mitglieder_Export_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $fullPath = Join-Path -Path $targetPath -ChildPath $fileName

    try {
        $allMemberRows | Export-Csv -Path $fullPath -Delimiter ';' -NoTypeInformation -Encoding UTF8
        Write-ToolLog -Message "Mitglieder manuell exportiert nach $fullPath ($exportedGroupsCount Gruppen analysiert, $($allMemberRows.Count) Zeilen)" -Severity "SUCCESS"
        [System.Windows.Forms.MessageBox]::Show(
            "Mitglieder erfolgreich exportiert!`r`n`r`n" +
            "Datei: $fullPath`r`n" +
            "Erfasste Gruppen: $exportedGroupsCount`r`n" +
            "Zeilen insgesamt: $($allMemberRows.Count)",
            "Mitglieder-Export abgeschlossen",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    } catch {
        Write-ToolLog -Message "Fehler beim Schreiben des Mitglieder-Exports: $($_.Exception.Message)" -Severity "ERROR"
        [System.Windows.Forms.MessageBox]::Show("Fehler beim Exportieren: $($_.Exception.Message)", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# ------------------------------------------------------------------------------
# GEGENPRÜFUNG REGISTER 2
# ------------------------------------------------------------------------------
$scanOrphansAction = {
    $btnScanOrphans.Enabled = $false
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $script:OrphanedGroupsList.Clear()
    $gridOrphans.DataSource = $null
    $lblOrphanCount.Text = "Prüfe Gruppen auf fehlende Computerobjekte..."
    $form.Refresh()

    $rootEntry = $null
    $cSearcher = $null
    $foundComps = $null
    $gSearcher = $null
    $patternGroups = $null

    try {
        $rootEntry = [System.DirectoryServices.DirectoryEntry]::new("LDAP://$domainDN")

        $cSearcher = [System.DirectoryServices.DirectorySearcher]::new($rootEntry)
        $cSearcher.Filter = "(objectCategory=computer)"
        $cSearcher.PropertiesToLoad.AddRange(@("name"))
        $cSearcher.PageSize = 1000
        $foundComps = $cSearcher.FindAll()

        $allCompNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($c in $foundComps) {
            if ($c.Properties["name"].Count -gt 0) {
                [void]$allCompNames.Add([string]$c.Properties["name"][0])
            }
        }

        $gSearcher = [System.DirectoryServices.DirectorySearcher]::new($rootEntry)
        $gSearcher.Filter = "(&(objectCategory=group)(|(name=*-RDP)(name=RDP-*)(name=*-lokalAdmin)(name=*-localAdmin)(name=lokalAdmin-*)))"
        $gSearcher.PropertiesToLoad.AddRange(@("name", "sAMAccountName", "description", "distinguishedName", "member"))
        $gSearcher.PageSize = 1000
        $patternGroups = $gSearcher.FindAll()

        $orphanRaw = [System.Collections.Generic.List[PSCustomObject]]::new()

        foreach ($g in $patternGroups) {
            $gName = [string]$g.Properties["name"][0]
            $gDN   = [string]$g.Properties["distinguishedname"][0]
            $gDesc = if ($g.Properties["description"].Count -gt 0) { [string]$g.Properties["description"][0] } else { "-" }
            $mCount = if ($g.Properties["member"]) { $g.Properties["member"].Count } else { 0 }

            $extractedPC = ""
            $gType = ""

            if ($gName -match "^(.+)-(?:lokalAdmin|localAdmin)$") {
                $extractedPC = $Matches[1].Trim()
                $gType = "lokalAdmin"
            } elseif ($gName -match "^(?:lokalAdmin|localAdmin)-(.+)$") {
                $extractedPC = $Matches[1].Trim()
                $gType = "lokalAdmin"
            } elseif ($gName -match "^(.+)-RDP$") {
                $extractedPC = $Matches[1].Trim()
                $gType = "RDP"
            } elseif ($gName -match "^RDP-(.+)$") {
                $extractedPC = $Matches[1].Trim()
                $gType = "RDP"
            }

            if ($extractedPC -and (-not $allCompNames.Contains($extractedPC))) {
                $ouPath = if ($gDN -match "OU=.*") { $gDN.Substring($gDN.IndexOf("OU=")) } else { "CN=Users/Builtin" }

                $orphanRaw.Add([PSCustomObject]@{
                    "Gruppenname"           = $gName
                    "Typ"                   = $gType
                    "Vermuteter Computer"   = $extractedPC
                    "Computer im AD"        = "Nicht vorhanden (Gelöscht/Verwaist)"
                    "Mitglieder-Anzahl"     = [int]$mCount
                    "Beschreibung"          = $gDesc
                    "OU / Pfad"             = $ouPath
                    "DistinguishedName"     = $gDN
                    "_RawMembers"           = @($g.Properties["member"])
                })
            }
        }

        $sortedOrphans = $orphanRaw | Sort-Object "Vermuteter Computer"
        foreach ($so in $sortedOrphans) { [void]$script:OrphanedGroupsList.Add($so) }

        $gridOrphans.SuspendLayout()
        $gridOrphans.DataSource = $null
        $arrOrphans = [System.Collections.ArrayList]::new()
        foreach ($o in $script:OrphanedGroupsList) { [void]$arrOrphans.Add($o) }
        $gridOrphans.DataSource = $arrOrphans
        $gridOrphans.ClearSelection()
        $gridOrphans.ResumeLayout()

        $lblOrphanCount.Text = "Gefundene verwaiste Gruppen ohne Computerobjekt: $($script:OrphanedGroupsList.Count)"
        Write-ToolLog -Message "Gegenprüfung abgeschlossen. $($script:OrphanedGroupsList.Count) verwaiste Gruppen identifiziert." -Severity "INFO"
    } catch {
        Write-ToolLog -Message "Fehler bei Gegenprüfung verwaister Gruppen: $($_.Exception.Message)" -Severity "ERROR"
        [System.Windows.Forms.MessageBox]::Show("Fehler bei der Gegenprüfung: $($_.Exception.Message)", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } finally {
        if ($null -ne $foundComps) { $foundComps.Dispose() }
        if ($null -ne $cSearcher) { $cSearcher.Dispose() }
        if ($null -ne $patternGroups) { $patternGroups.Dispose() }
        if ($null -ne $gSearcher) { $gSearcher.Dispose() }
        if ($null -ne $rootEntry) { $rootEntry.Dispose() }

        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $btnScanOrphans.Enabled = $true
    }
}

$gridOrphans.Add_RowPrePaint({
    param($s, $e)
    if ($e.RowIndex -ge 0 -and $e.RowIndex -lt $gridOrphans.Rows.Count) {
        $row = $gridOrphans.Rows[$e.RowIndex]
        $row.DefaultCellStyle.BackColor          = [System.Drawing.Color]::FromArgb(255, 235, 235)
        $row.DefaultCellStyle.ForeColor          = [System.Drawing.Color]::FromArgb(180, 20, 20)
        $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(250, 190, 190)
        $row.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::FromArgb(130, 0, 0)
    }
})

# --------------------------------------------------------------------------
# BUTTON ACTIONS TAB 1 (ERSTELLUNG & INAKTIVE BEREINIGEN MIT EINZEL-BACKUPS)
# --------------------------------------------------------------------------
$btnCopyCmd.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtCommands.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Keine Befehle vorhanden.", "Hinweis", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    [System.Windows.Forms.Clipboard]::SetText($txtCommands.Text)
    [System.Windows.Forms.MessageBox]::Show("Befehle wurden in die Zwischenablage kopiert.", "Erfolg", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})

$btnCreateMissing.Add_Click({
    if ($script:PendingCreateCmds.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Für die aktuelle Auswahl gibt es keine fehlenden Gruppen zu erstellen.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    $diag = [System.Windows.Forms.MessageBox]::Show(
        "Möchten Sie $($script:PendingCreateCmds.Count) fehlende Gruppen für die aktuelle Auswahl im Active Directory anlegen?",
        "Gruppenerstellung bestätigen",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    if ($diag -eq [System.Windows.Forms.DialogResult]::Yes) {
        $created = 0
        $failed = 0
        $errorDetails = [System.Collections.Generic.List[string]]::new()
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

        Write-ToolLog -Message "Starte Erstellung von $($script:PendingCreateCmds.Count) Gruppen..." -Severity "INFO"

        foreach ($cmd in $script:PendingCreateCmds) {
            $grpName = if ($cmd -match "-Name '([^']+)'") { $Matches[1] } else { "Unbekannt" }
            try {
                $cmdToRun = "$cmd -ErrorAction Stop"
                Invoke-Expression $cmdToRun

                $verifiedGroup = Get-ADGroup -Filter "sAMAccountName -eq '$grpName'" -ErrorAction SilentlyContinue
                if ($null -ne $verifiedGroup) {
                    $created++
                    Write-ToolLog -Message "Gruppe '$grpName' erfolgreich erstellt und verifiziert." -Severity "SUCCESS"
                } else {
                    $failed++
                    $msg = "Gruppe '$grpName': Befehl ausgeführt, Gruppe im AD nicht auffindbar."
                    $errorDetails.Add(" $msg")
                    Write-ToolLog -Message $msg -Severity "ERROR"
                }
            } catch {
                $failed++
                $errMsg = $_.Exception.Message
                $errorDetails.Add(" Gruppe '$grpName': $errMsg")
                Write-ToolLog -Message "Fehler bei Erstellung von '$grpName': $errMsg" -Severity "ERROR"
            }
        }

        $form.Cursor = [System.Windows.Forms.Cursors]::Default

        if ($failed -gt 0) {
            $errorSummary = ($errorDetails | Select-Object -First 6) -join "`r`n"
            if ($errorDetails.Count -gt 6) {
                $errorSummary += "`r`n... und $($errorDetails.Count - 6) weitere Fehler."
            }

            [System.Windows.Forms.MessageBox]::Show(
                "Vorgang mit Fehlern beendet!`r`n`r`n" +
                "Tatsächlich erstellt: $created`r`n" +
                "Fehlgeschlagen: $failed (Keine Berechtigung auf Ziel-OU oder Syntaxfehler)`r`n`r`n" +
                "Details:`r`n$errorSummary`r`n`r`n" +
                "Hinweis: Details wurden im Activity-Log protokolliert.",
                "Gruppenerstellung Fehlgeschlagen",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
        } else {
            [System.Windows.Forms.MessageBox]::Show(
                "Alle $created Gruppen wurden erfolgreich im Active Directory erstellt und verifiziert.",
                "Erfolg",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }

        & $runCheckAction
    }
})

$btnDeleteOrphaned.Add_Click({
    if ($script:PendingDeleteGroups.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Für die aktuelle Auswahl wurden keine verwaisten Gruppen inaktiver Objekte gefunden.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    $diag = [System.Windows.Forms.MessageBox]::Show(
        "Achtung: Es wurden $($script:PendingDeleteGroups.Count) Gruppen inaktiver Rechner in der aktuellen Auswahl gefunden.`r`n`r`n" +
        "Vor dem Löschen wird für jede Gruppe automatisch ein separates Mitglieder-Backup in folgenden Ordner geschrieben:`r`n$($txtBackupDir.Text.Trim())`r`n`r`n" +
        "Möchten Sie das Backup erstellen und die Gruppen danach löschen?",
        "Löschen & Backup bestätigen",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($diag -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

    $targetPath = $txtBackupDir.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($targetPath)) { $targetPath = "C:\Install\Backup\AD_Groups" }
    if (-not (Test-Path -LiteralPath $targetPath)) {
        try { New-Item -ItemType Directory -Path $targetPath -Force | Out-Null } catch {}
    }

    $timeStamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $backupCount = 0
    Write-ToolLog -Message "Starte Vorab-Backup fuer $($script:PendingDeleteGroups.Count) Gruppen inaktiver PCs..." -Severity "INFO"

    foreach ($grp in $script:PendingDeleteGroups) {
        $grpName = $grp.GroupName
        $membersDetail = Get-GroupMembersDetail -GroupName $grpName -GroupDN $grp.DistinguishedName -ComputerName $grp.ComputerName -ComputerDescription $grp.ComputerDescription
        $cleanGrpName = $grpName -replace '[\\/:*?"<>|]', '_'
        $singleBackupFile = Join-Path -Path $targetPath -ChildPath "Backup_${cleanGrpName}_${timeStamp}.csv"

        try {
            $membersDetail | Export-Csv -Path $singleBackupFile -Delimiter ';' -NoTypeInformation -Encoding UTF8
            $backupCount++
            Write-ToolLog -Message "Sicherungsdatei erstellt: $singleBackupFile ($($membersDetail.Count) Zeilen)" -Severity "INFO"
        } catch {
            Write-ToolLog -Message "Fehler beim Erstellen der Sicherung fuer '$grpName': $($_.Exception.Message)" -Severity "ERROR"
        }
    }

    $deleted = 0
    $delFailed = 0
    foreach ($grp in $script:PendingDeleteGroups) {
        $gEntry = $null
        $parent = $null
        try {
            $gEntry = [ADSI]"LDAP://$($grp.DistinguishedName)"
            $parent = [ADSI]$gEntry.Parent
            $parent.Children.Remove($gEntry)
            $parent.CommitChanges()
            $deleted++
            Write-ToolLog -Message "Gruppe '$($grp.GroupName)' aus AD gelöscht." -Severity "SUCCESS"
        } catch { 
            $delFailed++ 
            Write-ToolLog -Message "Fehler beim Löschen von '$($grp.GroupName)': $($_.Exception.Message)" -Severity "ERROR"
        } finally {
            if ($null -ne $gEntry) { $gEntry.Dispose() }
            if ($null -ne $parent) { $parent.Dispose() }
        }
    }

    $form.Cursor = [System.Windows.Forms.Cursors]::Default
    [System.Windows.Forms.MessageBox]::Show(
        "Vorgang beendet!`r`n`r`n" +
        "Einzel-Backups der Mitglieder erstellt: $backupCount`r`n" +
        "Zielverzeichnis: $targetPath`r`n`r`n" +
        "Gelöschte Gruppen im AD: $deleted`r`nFehlgeschlagen: $delFailed",
        "Abgeschlossen",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )

    & $runCheckAction
})

$btnExportOrphans.Add_Click({
    if ($script:OrphanedGroupsList.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Keine verwaisten Gruppen zum Exportieren vorhanden.", "Hinweis", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $targetPath = $txtBackupDir.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($targetPath)) { $targetPath = "C:\Install\Backup\AD_Groups" }
    if (-not (Test-Path -LiteralPath $targetPath)) {
        try { New-Item -ItemType Directory -Path $targetPath -Force | Out-Null } catch {}
    }

    $timeStamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $count = 0

    foreach ($grp in $script:OrphanedGroupsList) {
        $membersDetail = Get-GroupMembersDetail -GroupName $grp.Gruppenname -GroupDN $grp.DistinguishedName -ComputerName $grp."Vermuteter Computer" -ComputerDescription $grp.Beschreibung
        $cleanName = $grp.Gruppenname -replace '[\\/:*?"<>|]', '_'
        $singleFile = Join-Path -Path $targetPath -ChildPath "Backup_Verwaist_${cleanName}_${timeStamp}.csv"

        try {
            $membersDetail | Export-Csv -Path $singleFile -Delimiter ';' -NoTypeInformation -Encoding UTF8
            $count++
        } catch {}
    }

    $form.Cursor = [System.Windows.Forms.Cursors]::Default
    Write-ToolLog -Message "Manueller Einzel-Export fuer $count verwaiste Gruppen nach $targetPath durchgefuehrt." -Severity "SUCCESS"
    [System.Windows.Forms.MessageBox]::Show("Mitglieder für $count verwaiste Gruppen erfolgreich in Einzelfiles exportiert nach:`r`n$targetPath", "Export abgeschlossen", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})

$btnDeleteOrphanList.Add_Click({
    if ($script:OrphanedGroupsList.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Keine verwaisten Gruppen vorhanden.", "Hinweis", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    $diag = [System.Windows.Forms.MessageBox]::Show(
        "Möchten Sie alle $($script:OrphanedGroupsList.Count) verwaisten Gruppen unwiderruflich aus dem Active Directory löschen?`r`n`r`n" +
        "Vor dem Löschen wird für jede Gruppe automatisch eine eigene Mitglieder-Sicherung in folgenden Pfad geschrieben:`r`n$($txtBackupDir.Text.Trim())",
        "Löschen bestätigen",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($diag -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

    $targetPath = $txtBackupDir.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($targetPath)) { $targetPath = "C:\Install\Backup\AD_Groups" }
    if (-not (Test-Path -LiteralPath $targetPath)) {
        try { New-Item -ItemType Directory -Path $targetPath -Force | Out-Null } catch {}
    }

    $timeStamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $backupCount = 0
    Write-ToolLog -Message "Starte Vorab-Backup fuer $($script:OrphanedGroupsList.Count) verwaiste Gruppen..." -Severity "INFO"

    foreach ($grp in $localOrphanedGroupsList) {
        $membersDetail = Get-GroupMembersDetail -GroupName $grp.Gruppenname -GroupDN $grp.DistinguishedName -ComputerName $grp."Vermuteter Computer" -ComputerDescription $grp.Beschreibung
        $cleanName = $grp.Gruppenname -replace '[\\/:*?"<>|]', '_'
        $singleFile = Join-Path -Path $targetPath -ChildPath "Backup_Verwaist_${cleanName}_${timeStamp}.csv"

        try {
            $membersDetail | Export-Csv -Path $singleFile -Delimiter ';' -NoTypeInformation -Encoding UTF8
            $backupCount++
        } catch {
            Write-ToolLog -Message "Fehler beim Sichern von verwaister Gruppe '$($grp.Gruppenname)': $($_.Exception.Message)" -Severity "ERROR"
        }
    }

    $deleted = 0
    $delFailed = 0

    foreach ($grp in $script:OrphanedGroupsList) {
        $gEntry = $null
        $parent = $null
        try {
            $gEntry = [ADSI]"LDAP://$($grp.DistinguishedName)"
            $parent = [ADSI]$gEntry.Parent
            $parent.Children.Remove($gEntry)
            $parent.CommitChanges()
            $deleted++
            Write-ToolLog -Message "Verwaiste Gruppe '$($grp.Gruppenname)' gelöscht." -Severity "SUCCESS"
        } catch {
            $delFailed++
            Write-ToolLog -Message "Fehler beim Löschen der verwaisten Gruppe '$($grp.Gruppenname)': $($_.Exception.Message)" -Severity "ERROR"
        } finally {
            if ($null -ne $gEntry) { $gEntry.Dispose() }
            if ($null -ne $parent) { $parent.Dispose() }
        }
    }

    $form.Cursor = [System.Windows.Forms.Cursors]::Default
    [System.Windows.Forms.MessageBox]::Show(
        "Löschvorgang beendet!`r`n`r`n" +
        "Einzel-Sicherungen erstellt: $backupCount`r`n" +
        "Zielverzeichnis: $targetPath`r`n`r`n" +
        "Gelöscht: $deleted`r`nFehlgeschlagen: $delFailed",
        "Ergebnis",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )

    & $scanOrphansAction
})

# ------------------------------------------------------------------------------
# EVENTS REGISTER 3
# ------------------------------------------------------------------------------
# HIER IST DAS VERMISSTE BUTTON-EVENT FEST VERBUNDEN:
$btnLoadMembersTab3.Add_Click($loadMembersActionTab3)

$cmbMbrTypeFilter.Add_SelectedIndexChanged({ & $filterMembersAction })
$cmbMbrStatusFilter.Add_SelectedIndexChanged({ & $filterMembersAction })
$txtMemberFilter.Add_TextChanged({ & $filterMembersAction })

$gridMembers.Add_SelectionChanged({
    & $updateMemberCommandsAction
})

$gridMembers.Add_RowPrePaint({
    param($s, $e)
    if ($e.RowIndex -ge 0 -and $e.RowIndex -lt $gridMembers.Rows.Count) {
        $row = $gridMembers.Rows[$e.RowIndex]
        $pcStatus = [string]$row.Cells["Computer-Status"].Value
        $mbrName  = [string]$row.Cells["Mitgliedsname"].Value

        if ($pcStatus -like "*Deaktiviert*") {
            $row.DefaultCellStyle.BackColor          = [System.Drawing.Color]::FromArgb(245, 245, 245)
            $row.DefaultCellStyle.ForeColor          = [System.Drawing.Color]::FromArgb(120, 120, 120)
            $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(220, 220, 220)
            $row.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::FromArgb(50, 50, 50)
        } elseif ($mbrName -eq "(Keine Mitglieder)") {
            $row.DefaultCellStyle.BackColor          = [System.Drawing.Color]::White
            $row.DefaultCellStyle.ForeColor          = [System.Drawing.Color]::Gray
            $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(235, 240, 248)
            $row.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::FromArgb(30, 41, 59)
        } else {
            $row.DefaultCellStyle.BackColor          = [System.Drawing.Color]::White
            $row.DefaultCellStyle.ForeColor          = [System.Drawing.Color]::FromArgb(30, 41, 59)
            $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(205, 230, 255)
            $row.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::FromArgb(0, 40, 100)
        }
    }
})

$btnCopyMemberCmd.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtMembersCommands.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Keine Befehle vorhanden.", "Hinweis", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    [System.Windows.Forms.Clipboard]::SetText($txtMembersCommands.Text)
    [System.Windows.Forms.MessageBox]::Show("Befehle wurden in die Zwischenablage kopiert.", "Erfolg", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})

$btnExportTab3Csv.Add_Click({
    if ($script:CurrentMemberRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Keine Mitgliederzeilen zum Exportieren vorhanden.", "Hinweis", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    $targetPath = $txtBackupDir.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($targetPath)) { $targetPath = "C:\Install\Backup\AD_Groups" }
    if (-not (Test-Path -LiteralPath $targetPath)) {
        try { New-Item -ItemType Directory -Path $targetPath -Force | Out-Null } catch {}
    }

    $fileName = "AD_Gruppenmitglieder_Tab3_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $fullPath = Join-Path -Path $targetPath -ChildPath $fileName

    try {
        $script:CurrentMemberRows | Export-Csv -Path $fullPath -Delimiter ';' -NoTypeInformation -Encoding UTF8
        Write-ToolLog -Message "Register 3 CSV-Export erstellt: $fullPath ($($script:CurrentMemberRows.Count) Zeilen)" -Severity "SUCCESS"
        [System.Windows.Forms.MessageBox]::Show("Mitgliederliste erfolgreich exportiert nach:`r`n$fullPath", "Export abgeschlossen", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } catch {
        Write-ToolLog -Message "Fehler beim Exportieren in Register 3: $($_.Exception.Message)" -Severity "ERROR"
        [System.Windows.Forms.MessageBox]::Show("Fehler beim Exportieren: $($_.Exception.Message)", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

function Execute-MemberRemoval {
    param(
        [System.Collections.IEnumerable]$MembersToRemove,
        [string]$RemovalTypeLabel
    )

    $validMembers = @($MembersToRemove | Where-Object { $_.Mitgliedsname -ne "(Keine Mitglieder)" -and -not [string]::IsNullOrWhiteSpace($_.'Mitglieds-DN') })
    if ($validMembers.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("In der getroffenen Auswahl befinden sich keine gültigen Gruppenmitglieder zum Entfernen.", "Hinweis", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    $diag = [System.Windows.Forms.MessageBox]::Show(
        "Achtung: Es werden $($validMembers.Count) Gruppenmitgliedschaft(en) aus dem Active Directory entfernt.`r`n`r`n" +
        "Vor dem Entfernen wird automatisch ein Backup der betroffenen Mitglieder in folgenden Pfad geschrieben:`r`n$($txtBackupDir.Text.Trim())`r`n`r`n" +
        "Möchten Sie das Backup erstellen und die Mitglieder danach entfernen?",
        "Mitglieder-Entfernung bestätigen",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($diag -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

    $targetPath = $txtBackupDir.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($targetPath)) { $targetPath = "C:\Install\Backup\AD_Groups" }
    if (-not (Test-Path -LiteralPath $targetPath)) {
        try { New-Item -ItemType Directory -Path $targetPath -Force | Out-Null } catch {}
    }

    # 1. Automatisches Backup vor der Bereinigung
    $timeStamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $backupFile = Join-Path -Path $targetPath -ChildPath "Backup_MitgliederBereinigung_${RemovalTypeLabel}_${timeStamp}.csv"

    try {
        $validMembers | Export-Csv -Path $backupFile -Delimiter ';' -NoTypeInformation -Encoding UTF8
        Write-ToolLog -Message "Vorab-Sicherung vor Mitglied-Bereinigung erstellt: $backupFile ($($validMembers.Count) Einträge)" -Severity "INFO"
    } catch {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        Write-ToolLog -Message "Abbruch: Backup-Datei konnte nicht geschrieben werden: $($_.Exception.Message)" -Severity "ERROR"
        [System.Windows.Forms.MessageBox]::Show("Abbruch: Die Backup-Datei konnte nicht geschrieben werden:`r`n$($_.Exception.Message)", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # 2. Entfernen aus dem Active Directory
    $removed = 0
    $failed = 0
    $errorList = [System.Collections.Generic.List[string]]::new()

    foreach ($m in $validMembers) {
        $grpIdentity = $m.Gruppenname
        $memberIdentity = if (-not [string]::IsNullOrWhiteSpace($m.SamAccountName) -and $m.SamAccountName -ne "-") { $m.SamAccountName } else { $m.'Mitglieds-DN' }
        try {
            Remove-ADGroupMember -Identity $grpIdentity -Members $memberIdentity -Confirm:$false -ErrorAction Stop
            $removed++
            Write-ToolLog -Message "Mitglied '$memberIdentity' aus Gruppe '$grpIdentity' entfernt." -Severity "SUCCESS"
        } catch {
            $failed++
            $errMsg = $_.Exception.Message
            $errorList.Add("• '$memberIdentity' aus '$grpIdentity': $errMsg")
            Write-ToolLog -Message "Fehler beim Entfernen von '$memberIdentity' aus '$grpIdentity': $errMsg" -Severity "ERROR"
        }
    }

    $form.Cursor = [System.Windows.Forms.Cursors]::Default

    $msgText = "Vorgang beendet!`r`n`r`n" +
               "Sicherungsdatei erstellt unter:`r`n$backupFile`r`n`r`n" +
               "Erfolgreich entfernt: $removed`r`nFehlgeschlagen: $failed"

    if ($failed -gt 0) {
        $errSummary = ($errorList | Select-Object -First 5) -join "`r`n"
        $msgText += "`r`n`r`nErste Fehler:`r`n$errSummary"
        [System.Windows.Forms.MessageBox]::Show($msgText, "Bereinigung mit Fehlern", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    } else {
        [System.Windows.Forms.MessageBox]::Show($msgText, "Bereinigung abgeschlossen", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }

    & $loadMembersActionTab3
}

$btnRemoveSelectedMembers.Add_Click({
    $selectedItems = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($r in $gridMembers.SelectedRows) {
        if ($r.DataBoundItem) { [void]$selectedItems.Add($r.DataBoundItem) }
    }
    Execute-MemberRemoval -MembersToRemove $selectedItems -RemovalTypeLabel "Auswahl"
})

$btnRemoveAllListedMembers.Add_Click({
    Execute-MemberRemoval -MembersToRemove $script:CurrentMemberRows -RemovalTypeLabel "AlleGelisteten"
})

# Automatisches Laden beim Anklicken von Register 3
$tabControl.Add_SelectedIndexChanged({
    if ($tabControl.SelectedTab -eq $tabMembers -and $script:AllMembersList.Count -eq 0) {
        & $loadMembersActionTab3
    }
})

# Events Register 1 & 2 verknüpfen
$btnScan.Add_Click($runCheckAction)
$btnScanOrphans.Add_Click($scanOrphansAction)

$cmbTypeFilter.Add_SelectedIndexChanged($filterResultsAction)
$cmbActiveFilter.Add_SelectedIndexChanged($filterResultsAction)
$cmbStatusFilter.Add_SelectedIndexChanged($filterResultsAction)

$txtCompFilter.Add_KeyDown({
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $_.SuppressKeyPress = $true
        & $runCheckAction
    }
})

$form.Add_FormClosing({
    param($sender, $e)
    Write-ToolLog -Message "Tool 24 beendet." -Severity "INFO"
})

Load-DomainOUsAndPreselect
Write-ToolLog -Message "Tool 24 gestartet." -Severity "INFO"

try {
    [void]$form.ShowDialog()
} finally {
    if ($null -ne $form) {
        $form.Dispose()
    }

    $script:RawCheckResults.Clear()
    $script:CurrentFilteredRows.Clear()
    $script:PendingCreateCmds.Clear()
    $script:PendingDeleteGroups.Clear()
    $script:OrphanedGroupsList.Clear()
    $script:AllMembersList.Clear()
    $script:CurrentMemberRows.Clear()
    $script:PendingRemoveMemberCmds.Clear()

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    [System.GC]::Collect()
}