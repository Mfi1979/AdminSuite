<#
==================================================================================
 Tool 15: Active Directory GPO Enterprise Suite
 Version: 1.5.2 (DPI- & Layout-optimiert, ASCII-only, Clean-Exit)
 
 Register 1: GPO Uebersicht & Verlinkungs-Analyse
             - Schnelle ADSI/LDAP-Abfrage aller GPOs, WMI-Filter & OU-Verlinkungen
             - Spaltenbreiten automatisch an Inhalt angepasst (AllCells)
             - Spaltenbezeichnungen "Benutzer" und "Computer"
             - Farbliche Kennzeichnung:
               * [Blau/Lila] Default Domain Policy & Default Domain Controllers Policy
               * [Gruen] OK (Nur Computer aktiv ODER nur Benutzer aktiv)
               * [Rot] Nicht OK (Beide aktiviert oder vollstaendig deaktiviert)
             - Klickbare Spaltensortierung auf allen Spalten
             - Master-Detail: Zeigt rechts alle Verlinkungsziele (OUs/Domaene)
             - PowerShell- & Tool-Info-Dialog via Button
 
 Register 2: GPO Richtlinien-Einstellungen & Inspektor (Praeziser XML-Parser)
             - Bereinigte Spalte "Konfigurierter Wert" (reine Werte ohne Explain-Texte)
             - Vollstaendige Richtlinienerklaerung rechts im Detailbereich
 
 Register 3: GPO Backup & Verknuepfungs-Audit
             - DPI-sichere 2-Zeilen-Kopfleiste (kein Abschneiden von Buttons/Texten)
             - Konfigurierbarer Ziel-Pfad mit Ordnerauswahl-Dialog
             - Einzelsicherung oder Gesamtsicherung aller GPOs
             - Datums- & Zeitstruktur: [Zielpfad]\[GPO-Name]\[JJJJMMTT_HHMM]
             - Erstellung von 'GPO_Link_Info.txt'
 
 Register 4: GPO-Vergleich (Diff mit getrennten Status- & Werte-Spalten)
             - 2 beliebige GPOs gegeneinander vergleichen
             - Spalten: Bereich, Kategorie, Einstellung, Status GPO 1, Wert GPO 1, Status GPO 2, Wert GPO 2, Diff-Status
             - Erkennt Parameter- & Zahlenwert-Abweichungen (Gelb/Orange)
             - Identische Einstellungen vollstaendig in GRUEN
==================================================================================
#>

[System.Windows.Forms.Application]::EnableVisualStyles()
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.DirectoryServices
Import-Module GroupPolicy -ErrorAction Stop

$script:ToolVersion = "v1.5.2"

function Show-Tool15 {
    [CmdletBinding()]
    param()

    # --- Domaenenpruefung ---
    try {
        $domainInfo = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
        $domainDN = ([ADSI]"LDAP://RootDSE").defaultNamingContext.Value
        $domainName = $domainInfo.Name
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Active Directory Domaene nicht erreichbar oder Computer nicht domaenengebunden.",
            "Fehler",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }

    # --- Hauptfenster ---
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Tool 15 - Active Directory GPO Enterprise Suite ($domainName) - $script:ToolVersion"
    $form.Size = New-Object System.Drawing.Size(1680, 930)
    $form.StartPosition = "CenterScreen"
    $form.MinimumSize = New-Object System.Drawing.Size(1250, 750)
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Dock = [System.Windows.Forms.DockStyle]::Fill

    # =========================================================================
    # REGISTER 1: GPO Uebersicht & Verlinkungs-Analyse
    # =========================================================================
    $tabOverview = New-Object System.Windows.Forms.TabPage
    $tabOverview.Text = "1. GPO Uebersicht & Verlinkungs-Analyse"

    $panelOverviewTop = New-Object System.Windows.Forms.Panel
    $panelOverviewTop.Dock = [System.Windows.Forms.DockStyle]::Top
    $panelOverviewTop.Height = 78
    $panelOverviewTop.BackColor = [System.Drawing.Color]::FromArgb(245, 248, 252)
    $panelOverviewTop.Padding = New-Object System.Windows.Forms.Padding(10)

    # Zeile 1
    $lblViewFilter = New-Object System.Windows.Forms.Label
    $lblViewFilter.Text = "Ansicht:"
    $lblViewFilter.Location = New-Object System.Drawing.Point(12, 16)
    $lblViewFilter.AutoSize = $true
    $lblViewFilter.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $comboViewMode = New-Object System.Windows.Forms.ComboBox
    $comboViewMode.Location = New-Object System.Drawing.Point(72, 13)
    $comboViewMode.Size = New-Object System.Drawing.Size(220, 25)
    $comboViewMode.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$comboViewMode.Items.Add("Alle GPOs")
    [void]$comboViewMode.Items.Add("GPOs mit WMI-Filter")
    [void]$comboViewMode.Items.Add("Nicht verlinkte GPOs (Unlinked)")
    $comboViewMode.SelectedIndex = 0

    $lblOverviewSearch = New-Object System.Windows.Forms.Label
    $lblOverviewSearch.Text = "Suche:"
    $lblOverviewSearch.Location = New-Object System.Drawing.Point(305, 16)
    $lblOverviewSearch.AutoSize = $true
    $lblOverviewSearch.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $txtOverviewSearch = New-Object System.Windows.Forms.TextBox
    $txtOverviewSearch.Location = New-Object System.Drawing.Point(355, 13)
    $txtOverviewSearch.Size = New-Object System.Drawing.Size(180, 25)

    $btnRefreshOverview = New-Object System.Windows.Forms.Button
    $btnRefreshOverview.Text = "Neu laden"
    $btnRefreshOverview.Location = New-Object System.Drawing.Point(545, 10)
    $btnRefreshOverview.Size = New-Object System.Drawing.Size(100, 30)
    $btnRefreshOverview.BackColor = [System.Drawing.Color]::FromArgb(225, 238, 255)

    $btnExportOverviewCsv = New-Object System.Windows.Forms.Button
    $btnExportOverviewCsv.Text = "CSV Export"
    $btnExportOverviewCsv.Location = New-Object System.Drawing.Point(652, 10)
    $btnExportOverviewCsv.Size = New-Object System.Drawing.Size(110, 30)
    $btnExportOverviewCsv.BackColor = [System.Drawing.Color]::FromArgb(230, 245, 230)

    $btnPsInfo = New-Object System.Windows.Forms.Button
    $btnPsInfo.Text = "PS- & Tool-Info"
    $btnPsInfo.Location = New-Object System.Drawing.Point(768, 10)
    $btnPsInfo.Size = New-Object System.Drawing.Size(120, 30)
    $btnPsInfo.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 245)
    $btnPsInfo.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)

    # Zeile 2: Legende
    $lblLegendOverview = New-Object System.Windows.Forms.Label
    $lblLegendOverview.Text = "Legende:  [Blau/Lila] Default GPO  |  [Gruen] OK (Nur Computer oder Benutzer)  |  [Rot] Nicht OK (Beide aktiv / inaktiv)"
    $lblLegendOverview.Location = New-Object System.Drawing.Point(12, 48)
    $lblLegendOverview.AutoSize = $true
    $lblLegendOverview.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Italic)

    $panelOverviewTop.Controls.Add($lblViewFilter)
    $panelOverviewTop.Controls.Add($comboViewMode)
    $panelOverviewTop.Controls.Add($lblOverviewSearch)
    $panelOverviewTop.Controls.Add($txtOverviewSearch)
    $panelOverviewTop.Controls.Add($btnRefreshOverview)
    $panelOverviewTop.Controls.Add($btnExportOverviewCsv)
    $panelOverviewTop.Controls.Add($btnPsInfo)
    $panelOverviewTop.Controls.Add($lblLegendOverview)

    $splitOverview = New-Object System.Windows.Forms.SplitContainer
    $splitOverview.Dock = [System.Windows.Forms.DockStyle]::Fill
    $splitOverview.SplitterDistance = 980
    $splitOverview.SplitterWidth = 6

    # Linke Seite Tab 1
    $panelOvLeft = New-Object System.Windows.Forms.Panel
    $panelOvLeft.Dock = [System.Windows.Forms.DockStyle]::Fill
    $panelOvLeft.Padding = New-Object System.Windows.Forms.Padding(10, 8, 4, 10)

    $lblOvMasterTitle = New-Object System.Windows.Forms.Label
    $lblOvMasterTitle.Text = "Gruppenrichtlinien der Domaene (Klick auf Spaltenkopf zum Sortieren):"
    $lblOvMasterTitle.Dock = [System.Windows.Forms.DockStyle]::Top
    $lblOvMasterTitle.Height = 25
    $lblOvMasterTitle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $gridOvMaster = New-Object System.Windows.Forms.DataGridView
    $gridOvMaster.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridOvMaster.ReadOnly = $true
    $gridOvMaster.AllowUserToAddRows = $false
    $gridOvMaster.AllowUserToDeleteRows = $false
    $gridOvMaster.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $gridOvMaster.MultiSelect = $false
    $gridOvMaster.RowHeadersVisible = $false
    $gridOvMaster.ColumnHeadersVisible = $true
    $gridOvMaster.EnableHeadersVisualStyles = $false
    $gridOvMaster.BackgroundColor = [System.Drawing.Color]::White
    $gridOvMaster.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $gridOvMaster.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
    $gridOvMaster.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(230, 236, 245)
    $gridOvMaster.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $gridOvMaster.ColumnHeadersHeight = 32

    $panelOvLeft.Controls.Add($gridOvMaster)
    $panelOvLeft.Controls.Add($lblOvMasterTitle)
    $splitOverview.Panel1.Controls.Add($panelOvLeft)

    # Rechte Seite Tab 1
    $panelOvRight = New-Object System.Windows.Forms.Panel
    $panelOvRight.Dock = [System.Windows.Forms.DockStyle]::Fill
    $panelOvRight.Padding = New-Object System.Windows.Forms.Padding(4, 8, 10, 10)

    $lblOvDetailsTitle = New-Object System.Windows.Forms.Label
    $lblOvDetailsTitle.Text = "Verlinkungsziele der GPO (OUs / Domaene):"
    $lblOvDetailsTitle.Dock = [System.Windows.Forms.DockStyle]::Top
    $lblOvDetailsTitle.Height = 25
    $lblOvDetailsTitle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $gridOvDetails = New-Object System.Windows.Forms.DataGridView
    $gridOvDetails.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridOvDetails.ReadOnly = $true
    $gridOvDetails.AllowUserToAddRows = $false
    $gridOvDetails.AllowUserToDeleteRows = $false
    $gridOvDetails.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $gridOvDetails.RowHeadersVisible = $false
    $gridOvDetails.ColumnHeadersVisible = $true
    $gridOvDetails.EnableHeadersVisualStyles = $false
    $gridOvDetails.BackgroundColor = [System.Drawing.Color]::White
    $gridOvDetails.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $gridOvDetails.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(230, 236, 245)
    $gridOvDetails.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $gridOvDetails.ColumnHeadersHeight = 32
    $gridOvDetails.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill

    $panelOvRight.Controls.Add($gridOvDetails)
    $panelOvRight.Controls.Add($lblOvDetailsTitle)
    $splitOverview.Panel2.Controls.Add($panelOvRight)

    $tabOverview.Controls.Add($splitOverview)
    $tabOverview.Controls.Add($panelOverviewTop)

    # =========================================================================
    # REGISTER 2: GPO Richtlinien-Einstellungen & Inspektor
    # =========================================================================
    $tabSettings = New-Object System.Windows.Forms.TabPage
    $tabSettings.Text = "2. GPO Richtlinien-Einstellungen & Inspektor"

    $panelSettingsTop = New-Object System.Windows.Forms.Panel
    $panelSettingsTop.Dock = [System.Windows.Forms.DockStyle]::Top
    $panelSettingsTop.Height = 65
    $panelSettingsTop.BackColor = [System.Drawing.Color]::FromArgb(242, 245, 250)
    $panelSettingsTop.Padding = New-Object System.Windows.Forms.Padding(10)

    $lblGpo = New-Object System.Windows.Forms.Label
    $lblGpo.Text = "GPO:"
    $lblGpo.Location = New-Object System.Drawing.Point(12, 20)
    $lblGpo.AutoSize = $true
    $lblGpo.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $comboGpo = New-Object System.Windows.Forms.ComboBox
    $comboGpo.Location = New-Object System.Drawing.Point(55, 17)
    $comboGpo.Size = New-Object System.Drawing.Size(360, 25)
    $comboGpo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

    $lblFilter = New-Object System.Windows.Forms.Label
    $lblFilter.Text = "Suche:"
    $lblFilter.Location = New-Object System.Drawing.Point(425, 20)
    $lblFilter.AutoSize = $true
    $lblFilter.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $txtFilter = New-Object System.Windows.Forms.TextBox
    $txtFilter.Location = New-Object System.Drawing.Point(475, 17)
    $txtFilter.Size = New-Object System.Drawing.Size(200, 25)

    $btnLoadSettings = New-Object System.Windows.Forms.Button
    $btnLoadSettings.Text = "Laden"
    $btnLoadSettings.Location = New-Object System.Drawing.Point(685, 15)
    $btnLoadSettings.Size = New-Object System.Drawing.Size(95, 30)
    $btnLoadSettings.BackColor = [System.Drawing.Color]::FromArgb(225, 238, 255)

    $btnExportCsv = New-Object System.Windows.Forms.Button
    $btnExportCsv.Text = "CSV Export"
    $btnExportCsv.Location = New-Object System.Drawing.Point(788, 15)
    $btnExportCsv.Size = New-Object System.Drawing.Size(105, 30)
    $btnExportCsv.BackColor = [System.Drawing.Color]::FromArgb(230, 245, 230)

    $lblSettingsStatus = New-Object System.Windows.Forms.Label
    $lblSettingsStatus.Text = "Bereit."
    $lblSettingsStatus.Location = New-Object System.Drawing.Point(905, 21)
    $lblSettingsStatus.AutoSize = $true
    $lblSettingsStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)

    $panelSettingsTop.Controls.Add($lblGpo)
    $panelSettingsTop.Controls.Add($comboGpo)
    $panelSettingsTop.Controls.Add($lblFilter)
    $panelSettingsTop.Controls.Add($txtFilter)
    $panelSettingsTop.Controls.Add($btnLoadSettings)
    $panelSettingsTop.Controls.Add($btnExportCsv)
    $panelSettingsTop.Controls.Add($lblSettingsStatus)

    # SplitContainer Tab 2
    $splitSettings = New-Object System.Windows.Forms.SplitContainer
    $splitSettings.Dock = [System.Windows.Forms.DockStyle]::Fill
    $splitSettings.SplitterDistance = 900
    $splitSettings.SplitterWidth = 6

    $panelSettingsLeft = New-Object System.Windows.Forms.Panel
    $panelSettingsLeft.Dock = [System.Windows.Forms.DockStyle]::Fill
    $panelSettingsLeft.Padding = New-Object System.Windows.Forms.Padding(10, 8, 4, 10)

    $lblTableTitle = New-Object System.Windows.Forms.Label
    $lblTableTitle.Text = "Konfigurierte Einstellungen:"
    $lblTableTitle.Dock = [System.Windows.Forms.DockStyle]::Top
    $lblTableTitle.Height = 25
    $lblTableTitle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $gridSettings = New-Object System.Windows.Forms.DataGridView
    $gridSettings.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridSettings.ReadOnly = $true
    $gridSettings.AllowUserToAddRows = $false
    $gridSettings.AllowUserToDeleteRows = $false
    $gridSettings.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $gridSettings.MultiSelect = $false
    $gridSettings.RowHeadersVisible = $false
    $gridSettings.ColumnHeadersVisible = $true
    $gridSettings.EnableHeadersVisualStyles = $false
    $gridSettings.BackgroundColor = [System.Drawing.Color]::White
    $gridSettings.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $gridSettings.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(230, 236, 245)
    $gridSettings.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $gridSettings.ColumnHeadersHeight = 32
    $gridSettings.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 253)

    [void]$gridSettings.Columns.Add("colScope", "Bereich")
    [void]$gridSettings.Columns.Add("colCategory", "Kategorie / Pfad")
    [void]$gridSettings.Columns.Add("colName", "Einstellung (Name)")
    [void]$gridSettings.Columns.Add("colValue", "Konfigurierter Wert")
    [void]$gridSettings.Columns.Add("colState", "Status")
    [void]$gridSettings.Columns.Add("colSupported", "Unterstuetzt ab")
    [void]$gridSettings.Columns.Add("colExplain", "Erklaerung")
    [void]$gridSettings.Columns.Add("colGpo", "GPO")

    $gridSettings.Columns["colScope"].Width = 90
    $gridSettings.Columns["colCategory"].Width = 240
    $gridSettings.Columns["colName"].Width = 300
    $gridSettings.Columns["colValue"].AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
    $gridSettings.Columns["colState"].Visible = $false
    $gridSettings.Columns["colSupported"].Visible = $false
    $gridSettings.Columns["colExplain"].Visible = $false
    $gridSettings.Columns["colGpo"].Visible = $false

    $panelSettingsLeft.Controls.Add($gridSettings)
    $panelSettingsLeft.Controls.Add($lblTableTitle)
    $splitSettings.Panel1.Controls.Add($panelSettingsLeft)

    # Rechte Seite Tab 2 (Detailbereich)
    $panelSettingsRight = New-Object System.Windows.Forms.Panel
    $panelSettingsRight.Dock = [System.Windows.Forms.DockStyle]::Fill
    $panelSettingsRight.Padding = New-Object System.Windows.Forms.Padding(4, 8, 10, 10)

    $lblDescHeader = New-Object System.Windows.Forms.Label
    $lblDescHeader.Text = "Erlaeuterung & Richtlinien-Details:"
    $lblDescHeader.Dock = [System.Windows.Forms.DockStyle]::Top
    $lblDescHeader.Height = 25
    $lblDescHeader.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $txtDescription = New-Object System.Windows.Forms.TextBox
    $txtDescription.Dock = [System.Windows.Forms.DockStyle]::Fill
    $txtDescription.Multiline = $true
    $txtDescription.ReadOnly = $true
    $txtDescription.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $txtDescription.BackColor = [System.Drawing.Color]::FromArgb(252, 252, 254)
    $txtDescription.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)

    $panelSettingsRight.Controls.Add($txtDescription)
    $panelSettingsRight.Controls.Add($lblDescHeader)
    $splitSettings.Panel2.Controls.Add($panelSettingsRight)

    $tabSettings.Controls.Add($splitSettings)
    $tabSettings.Controls.Add($panelSettingsTop)

    # =========================================================================
    # REGISTER 3: GPO Backup & Verknuepfungs-Audit (DPI-sichere 2-Zeilen-Leiste)
    # =========================================================================
    $tabBackup = New-Object System.Windows.Forms.TabPage
    $tabBackup.Text = "3. GPO Backup & Verknuepfungs-Status"

    $panelBackupTop = New-Object System.Windows.Forms.Panel
    $panelBackupTop.Dock = [System.Windows.Forms.DockStyle]::Top
    $panelBackupTop.Height = 110
    $panelBackupTop.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
    $panelBackupTop.Padding = New-Object System.Windows.Forms.Padding(10)

    # Zeile 1: Pfad & Durchsuchen
    $lblTargetPath = New-Object System.Windows.Forms.Label
    $lblTargetPath.Text = "Backup Ziel-Pfad:"
    $lblTargetPath.Location = New-Object System.Drawing.Point(12, 18)
    $lblTargetPath.AutoSize = $true
    $lblTargetPath.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $txtBackupTargetDir = New-Object System.Windows.Forms.TextBox
    $txtBackupTargetDir.Location = New-Object System.Drawing.Point(135, 15)
    $txtBackupTargetDir.Size = New-Object System.Drawing.Size(480, 25)
    $txtBackupTargetDir.Text = "C:\Install\Backup\GPO"

    $btnBrowseFolder = New-Object System.Windows.Forms.Button
    $btnBrowseFolder.Text = "Durchsuchen..."
    $btnBrowseFolder.Location = New-Object System.Drawing.Point(625, 12)
    $btnBrowseFolder.Size = New-Object System.Drawing.Size(125, 30)

    # Zeile 2: Aktionen & Buttons (Großzügige Breiten gegen Abschneiden)
    $btnLoadGpos = New-Object System.Windows.Forms.Button
    $btnLoadGpos.Text = "GPO-Liste laden"
    $btnLoadGpos.Location = New-Object System.Drawing.Point(12, 56)
    $btnLoadGpos.Size = New-Object System.Drawing.Size(140, 34)

    $btnBackupSelected = New-Object System.Windows.Forms.Button
    $btnBackupSelected.Text = "Ausgewaehlte GPO sichern"
    $btnBackupSelected.Location = New-Object System.Drawing.Point(160, 56)
    $btnBackupSelected.Size = New-Object System.Drawing.Size(220, 34)
    $btnBackupSelected.BackColor = [System.Drawing.Color]::FromArgb(230, 245, 230)

    $btnBackupAll = New-Object System.Windows.Forms.Button
    $btnBackupAll.Text = "ALLE GPOs sichern"
    $btnBackupAll.Location = New-Object System.Drawing.Point(390, 56)
    $btnBackupAll.Size = New-Object System.Drawing.Size(180, 34)
    $btnBackupAll.BackColor = [System.Drawing.Color]::FromArgb(255, 245, 230)

    $lblBackupNote = New-Object System.Windows.Forms.Label
    $lblBackupNote.Text = "Format: [Zielpfad]\[Name der GPO]\[JJJJMMTT_HHMM]"
    $lblBackupNote.Location = New-Object System.Drawing.Point(585, 65)
    $lblBackupNote.AutoSize = $true
    $lblBackupNote.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Italic)

    $panelBackupTop.Controls.Add($lblTargetPath)
    $panelBackupTop.Controls.Add($txtBackupTargetDir)
    $panelBackupTop.Controls.Add($btnBrowseFolder)
    $panelBackupTop.Controls.Add($btnLoadGpos)
    $panelBackupTop.Controls.Add($btnBackupSelected)
    $panelBackupTop.Controls.Add($btnBackupAll)
    $panelBackupTop.Controls.Add($lblBackupNote)

    $panelBackupMain = New-Object System.Windows.Forms.Panel
    $panelBackupMain.Dock = [System.Windows.Forms.DockStyle]::Fill
    $panelBackupMain.Padding = New-Object System.Windows.Forms.Padding(10)

    $lblGpoGrid = New-Object System.Windows.Forms.Label
    $lblGpoGrid.Text = "1. Gruppenrichtlinien der Domaene:"
    $lblGpoGrid.Location = New-Object System.Drawing.Point(10, 5)
    $lblGpoGrid.AutoSize = $true
    $lblGpoGrid.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $gridGpos = New-Object System.Windows.Forms.DataGridView
    $gridGpos.Location = New-Object System.Drawing.Point(10, 28)
    $gridGpos.Size = New-Object System.Drawing.Size(780, 670)
    $gridGpos.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
    $gridGpos.ReadOnly = $true
    $gridGpos.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
    $gridGpos.AllowUserToAddRows = $false
    $gridGpos.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $gridGpos.MultiSelect = $false
    $gridGpos.RowHeadersVisible = $false
    $gridGpos.BackgroundColor = [System.Drawing.Color]::White

    $lblLinks = New-Object System.Windows.Forms.Label
    $lblLinks.Text = "2. Verknuepfungs-Ziele (OUs / Sites):"
    $lblLinks.Location = New-Object System.Drawing.Point(805, 5)
    $lblLinks.AutoSize = $true
    $lblLinks.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $gridLinks = New-Object System.Windows.Forms.DataGridView
    $gridLinks.Location = New-Object System.Drawing.Point(805, 28)
    $gridLinks.Size = New-Object System.Drawing.Size(705, 270)
    $gridLinks.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $gridLinks.ReadOnly = $true
    $gridLinks.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $gridLinks.AllowUserToAddRows = $false
    $gridLinks.RowHeadersVisible = $false
    $gridLinks.BackgroundColor = [System.Drawing.Color]::White

    $lblLog = New-Object System.Windows.Forms.Label
    $lblLog.Text = "3. Backup- & Aktivitaets-Protokoll:"
    $lblLog.Location = New-Object System.Drawing.Point(805, 308)
    $lblLog.AutoSize = $true
    $lblLog.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $txtLog = New-Object System.Windows.Forms.TextBox
    $txtLog.Location = New-Object System.Drawing.Point(805, 330)
    $txtLog.Size = New-Object System.Drawing.Size(705, 368)
    $txtLog.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $txtLog.Multiline = $true
    $txtLog.ReadOnly = $true
    $txtLog.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $txtLog.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 252)
    $txtLog.Font = New-Object System.Drawing.Font("Consolas", 9)

    $panelBackupMain.Controls.Add($lblGpoGrid)
    $panelBackupMain.Controls.Add($gridGpos)
    $panelBackupMain.Controls.Add($lblLinks)
    $panelBackupMain.Controls.Add($gridLinks)
    $panelBackupMain.Controls.Add($lblLog)
    $panelBackupMain.Controls.Add($txtLog)

    $tabBackup.Controls.Add($panelBackupMain)
    $tabBackup.Controls.Add($panelBackupTop)

    # =========================================================================
    # REGISTER 4: GPO-Vergleich (Diff)
    # =========================================================================
    $tabCompare = New-Object System.Windows.Forms.TabPage
    $tabCompare.Text = "4. GPO-Vergleich (Diff)"

    $panelCompareTop = New-Object System.Windows.Forms.Panel
    $panelCompareTop.Dock = [System.Windows.Forms.DockStyle]::Top
    $panelCompareTop.Height = 70
    $panelCompareTop.BackColor = [System.Drawing.Color]::FromArgb(242, 245, 250)
    $panelCompareTop.Padding = New-Object System.Windows.Forms.Padding(10)

    $lblGpo1 = New-Object System.Windows.Forms.Label
    $lblGpo1.Text = "GPO 1 (Basis):"
    $lblGpo1.Location = New-Object System.Drawing.Point(12, 22)
    $lblGpo1.AutoSize = $true
    $lblGpo1.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $comboGpo1 = New-Object System.Windows.Forms.ComboBox
    $comboGpo1.Location = New-Object System.Drawing.Point(105, 18)
    $comboGpo1.Size = New-Object System.Drawing.Size(260, 25)
    $comboGpo1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

    $lblGpo2 = New-Object System.Windows.Forms.Label
    $lblGpo2.Text = "GPO 2 (Vergleich):"
    $lblGpo2.Location = New-Object System.Drawing.Point(380, 22)
    $lblGpo2.AutoSize = $true
    $lblGpo2.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $comboGpo2 = New-Object System.Windows.Forms.ComboBox
    $comboGpo2.Location = New-Object System.Drawing.Point(500, 18)
    $comboGpo2.Size = New-Object System.Drawing.Size(260, 25)
    $comboGpo2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

    $btnCompare = New-Object System.Windows.Forms.Button
    $btnCompare.Text = "Vergleichen"
    $btnCompare.Location = New-Object System.Drawing.Point(775, 16)
    $btnCompare.Size = New-Object System.Drawing.Size(110, 30)
    $btnCompare.BackColor = [System.Drawing.Color]::FromArgb(225, 238, 255)

    $chkOnlyDiffs = New-Object System.Windows.Forms.CheckBox
    $chkOnlyDiffs.Text = "Nur Unterschiede anzeigen"
    $chkOnlyDiffs.Location = New-Object System.Drawing.Point(900, 21)
    $chkOnlyDiffs.AutoSize = $true
    $chkOnlyDiffs.Checked = $true
    $chkOnlyDiffs.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $btnExportCompareCsv = New-Object System.Windows.Forms.Button
    $btnExportCompareCsv.Text = "Diff CSV Export"
    $btnExportCompareCsv.Location = New-Object System.Drawing.Point(1090, 16)
    $btnExportCompareCsv.Size = New-Object System.Drawing.Size(125, 30)
    $btnExportCompareCsv.BackColor = [System.Drawing.Color]::FromArgb(230, 245, 230)

    $lblCompareStatus = New-Object System.Windows.Forms.Label
    $lblCompareStatus.Text = "Waehlen Sie zwei GPOs fuer den Vergleich."
    $lblCompareStatus.Location = New-Object System.Drawing.Point(1230, 22)
    $lblCompareStatus.AutoSize = $true
    $lblCompareStatus.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Italic)

    $panelCompareTop.Controls.Add($lblGpo1)
    $panelCompareTop.Controls.Add($comboGpo1)
    $panelCompareTop.Controls.Add($lblGpo2)
    $panelCompareTop.Controls.Add($comboGpo2)
    $panelCompareTop.Controls.Add($btnCompare)
    $panelCompareTop.Controls.Add($chkOnlyDiffs)
    $panelCompareTop.Controls.Add($btnExportCompareCsv)
    $panelCompareTop.Controls.Add($lblCompareStatus)

    $panelCompareMain = New-Object System.Windows.Forms.Panel
    $panelCompareMain.Dock = [System.Windows.Forms.DockStyle]::Fill
    $panelCompareMain.Padding = New-Object System.Windows.Forms.Padding(10, 8, 10, 10)

    $lblCompareLegend = New-Object System.Windows.Forms.Label
    $lblCompareLegend.Text = "Legende:  [Gruen] Identische Einstellung  |  [Gelb/Orange] Abweichender Wert / Status  |  [Rot] Nur in GPO 1  |  [Blau] Nur in GPO 2"
    $lblCompareLegend.Dock = [System.Windows.Forms.DockStyle]::Top
    $lblCompareLegend.Height = 22
    $lblCompareLegend.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Italic)

    $gridCompare = New-Object System.Windows.Forms.DataGridView
    $gridCompare.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridCompare.ReadOnly = $true
    $gridCompare.AllowUserToAddRows = $false
    $gridCompare.AllowUserToDeleteRows = $false
    $gridCompare.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $gridCompare.MultiSelect = $false
    $gridCompare.RowHeadersVisible = $false
    $gridCompare.ColumnHeadersVisible = $true
    $gridCompare.EnableHeadersVisualStyles = $false
    $gridCompare.BackgroundColor = [System.Drawing.Color]::White
    $gridCompare.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $gridCompare.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(230, 236, 245)
    $gridCompare.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $gridCompare.ColumnHeadersHeight = 32

    [void]$gridCompare.Columns.Add("colCmpScope", "Bereich")
    [void]$gridCompare.Columns.Add("colCmpCategory", "Kategorie / Pfad")
    [void]$gridCompare.Columns.Add("colCmpName", "Einstellung (Name)")
    [void]$gridCompare.Columns.Add("colCmpState1", "Status (GPO 1)")
    [void]$gridCompare.Columns.Add("colCmpVal1", "Wert in GPO 1")
    [void]$gridCompare.Columns.Add("colCmpState2", "Status (GPO 2)")
    [void]$gridCompare.Columns.Add("colCmpVal2", "Wert in GPO 2")
    [void]$gridCompare.Columns.Add("colCmpStatus", "Vergleichs-Status")

    $gridCompare.Columns["colCmpScope"].Width = 85
    $gridCompare.Columns["colCmpCategory"].Width = 200
    $gridCompare.Columns["colCmpName"].Width = 250
    $gridCompare.Columns["colCmpState1"].Width = 100
    $gridCompare.Columns["colCmpVal1"].Width = 220
    $gridCompare.Columns["colCmpState2"].Width = 100
    $gridCompare.Columns["colCmpVal2"].Width = 220
    $gridCompare.Columns["colCmpStatus"].AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill

    $panelCompareMain.Controls.Add($gridCompare)
    $panelCompareMain.Controls.Add($lblCompareLegend)

    $tabCompare.Controls.Add($panelCompareMain)
    $tabCompare.Controls.Add($panelCompareTop)

    $tabControl.TabPages.Add($tabOverview)
    $tabControl.TabPages.Add($tabSettings)
    $tabControl.TabPages.Add($tabBackup)
    $tabControl.TabPages.Add($tabCompare)
    $form.Controls.Add($tabControl)

    # Lokale Datencontainer
    $rawOverviewList = [System.Collections.Generic.List[PSCustomObject]]::new()
    $gpoLinksCache   = @{}
    $rawSettingsList = [System.Collections.Generic.List[PSCustomObject]]::new()
    $rawCompareList  = [System.Collections.Generic.List[PSCustomObject]]::new()

    # =========================================================================
    # SPALTENSORTIERUNG FUER DATAGRIDVIEWS
    # =========================================================================
    function Enable-GridSorting {
        param([System.Windows.Forms.DataGridView]$Grid)

        $Grid.Add_ColumnHeaderMouseClick({
            param($sender, $e)
            if ($form.IsDisposed -or $Grid.IsDisposed) { return }

            $targetGrid = $sender
            $colProp = $targetGrid.Columns[$e.ColumnIndex].DataPropertyName
            if (-not $colProp) { $colProp = $targetGrid.Columns[$e.ColumnIndex].HeaderText }

            if ($targetGrid.Tag -and $targetGrid.Tag.Column -eq $colProp) {
                $asc = -not $targetGrid.Tag.Ascending
            } else {
                $asc = $true
            }
            $targetGrid.Tag = @{ Column = $colProp; Ascending = $asc }

            $data = @($targetGrid.DataSource)
            if ($null -eq $data -or $data.Count -le 1) { return }

            $sorted = $data | Sort-Object -Property @{
                Expression = {
                    $val = $_.$colProp
                    if ($null -eq $val) { return "" }
                    if ($colProp -eq "Link-Anzahl" -and ($val -as [int])) { return [int]$val }
                    return $val
                }
                Descending = (-not $asc)
            }

            $arr = [System.Collections.ArrayList]::new()
            foreach ($item in $sorted) { [void]$arr.Add($item) }
            $targetGrid.DataSource = $arr

            foreach ($col in $targetGrid.Columns) {
                $col.HeaderCell.SortGlyphDirection = [System.Windows.Forms.SortOrder]::None
            }
            $targetGrid.Columns[$e.ColumnIndex].HeaderCell.SortGlyphDirection = if ($asc) { 
                [System.Windows.Forms.SortOrder]::Ascending 
            } else { 
                [System.Windows.Forms.SortOrder]::Descending 
            }
        })
    }

    Enable-GridSorting -Grid $gridOvMaster
    Enable-GridSorting -Grid $gridOvDetails

    # =========================================================================
    # LOGIK TAB 1: ADSI GPO-Uebersicht & OU-Links
    # =========================================================================
    function Update-OverviewDisplay {
        if ($form.IsDisposed -or $gridOvMaster.IsDisposed) { return }
        $mode = $comboViewMode.SelectedItem
        $filterText = $txtOverviewSearch.Text.Trim()
        
        $filtered = $rawOverviewList | Where-Object {
            $item = $_
            $matchMode = switch ($mode) {
                "GPOs mit WMI-Filter"              { $item."WMI-Filter" -ne "-" }
                "Nicht verlinkte GPOs (Unlinked)" { $item."Verlinkt" -eq "Nein" }
                default                           { $true }
            }
            $matchSearch = if ([string]::IsNullOrWhiteSpace($filterText)) { $true } else {
                $item."GPO Name" -like "*$filterText*" -or $item."WMI-Filter" -like "*$filterText*" -or $item."Gesamt-Status" -like "*$filterText*"
            }
            $matchMode -and $matchSearch
        }

        $arr = [System.Collections.ArrayList]::new()
        foreach ($it in $filtered) { [void]$arr.Add($it) }
        $gridOvMaster.DataSource = $arr

        if ($gridOvMaster.Columns["GUID"]) { $gridOvMaster.Columns["GUID"].Visible = $false }
        if ($gridOvMaster.Columns["WMI Query"]) { $gridOvMaster.Columns["WMI Query"].Visible = $false }

        $lblLegendOverview.Text = "Status: $($arr.Count) von $($rawOverviewList.Count) GPOs  |  [Blau/Lila] Default GPO  |  [Gruen] OK  |  [Rot] Nicht OK"
    }

    $loadOverviewAction = {
        $lblLegendOverview.Text = "Lade AD-Struktur..."
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $form.Refresh()

        $wmiSearcher = $null
        $wmiResults = $null
        $linkSearcher = $null
        $ouResults = $null
        $gpoSearcher = $null
        $gpoResults = $null

        try {
            $gpoLinksCache.Clear()
            $rawOverviewList.Clear()

            # 1. WMI Filter
            $wmiMap = @{}
            $wmiSearcher = [System.DirectoryServices.DirectorySearcher]::new(
                [System.DirectoryServices.DirectoryEntry]::new("LDAP://CN=SOM,CN=WMIPolicy,CN=System,$domainDN")
            )
            $wmiSearcher.Filter = "(objectClass=msWMI-Som)"
            $wmiSearcher.PropertiesToLoad.AddRange(@("msWMI-Name", "msWMI-ID", "msWMI-Parm2"))
            
            try {
                $wmiResults = $wmiSearcher.FindAll()
                foreach ($w in $wmiResults) {
                    $id = $w.Properties["mswmi-id"][0]
                    $wName = $w.Properties["mswmi-name"][0]
                    $wQuery = if ($w.Properties["mswmi-parm2"]) { $w.Properties["mswmi-parm2"][0] } else { "" }
                    $wmiMap[$id] = [PSCustomObject]@{ Name = $wName; Query = $wQuery }
                }
            } catch {}

            # 2. OU & Domain Verlinkungen
            $linkSearcher = [System.DirectoryServices.DirectorySearcher]::new(
                [System.DirectoryServices.DirectoryEntry]::new("LDAP://$domainDN")
            )
            $linkSearcher.Filter = "(|(objectClass=organizationalUnit)(objectClass=domainDNS))"
            $linkSearcher.PropertiesToLoad.AddRange(@("distinguishedName", "gPLink", "name"))
            $linkSearcher.SearchScope = [System.DirectoryServices.SearchScope]::Subtree
            
            $ouResults = $linkSearcher.FindAll()
            foreach ($ou in $ouResults) {
                if ($ou.Properties["gplink"]) {
                    $rawGpLink = $ou.Properties["gplink"][0]
                    $targetDN = $ou.Properties["distinguishedname"][0]
                    $matches = [regex]::Matches($rawGpLink, "cn=({[a-fA-F0-9-]+})", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                    foreach ($m in $matches) {
                        $gpoGuid = $m.Groups[1].Value.ToUpper()
                        if (-not $gpoLinksCache.ContainsKey($gpoGuid)) {
                            $gpoLinksCache[$gpoGuid] = [System.Collections.Generic.List[string]]::new()
                        }
                        $gpoLinksCache[$gpoGuid].Add($targetDN)
                    }
                }
            }

            # 3. GPO Container
            $gpoSearcher = [System.DirectoryServices.DirectorySearcher]::new(
                [System.DirectoryServices.DirectoryEntry]::new("LDAP://CN=Policies,CN=System,$domainDN")
            )
            $gpoSearcher.Filter = "(objectClass=groupPolicyContainer)"
            $gpoSearcher.PropertiesToLoad.AddRange(@("displayName", "name", "flags", "gPCWQLFilter", "whenCreated", "whenChanged"))

            $gpoResults = $gpoSearcher.FindAll()
            foreach ($g in $gpoResults) {
                $guid = $g.Properties["name"][0].ToUpper()
                $displayName = if ($g.Properties["displayname"]) { $g.Properties["displayname"][0] } else { $guid }
                $flags = if ($g.Properties["flags"]) { [int]$g.Properties["flags"][0] } else { 0 }

                $userStatus = if (($flags -band 1) -eq 1) { "Deaktiviert" } else { "Aktiviert" }
                $compStatus = if (($flags -band 2) -eq 2) { "Deaktiviert" } else { "Aktiviert" }

                $isDefaultGPO = ($guid -in @("{31B2F340-016D-11D2-945F-00C04FB984F9}", "{6AC1786C-016F-11D2-945F-00C04FB984F9}")) -or 
                                ($displayName -match "^Default Domain Policy$" -or $displayName -match "^Default Domain Controllers Policy$")

                $overallStatus = if ($isDefaultGPO) {
                    "Sonderstellung (Default GPO)"
                } else {
                    switch ($flags) {
                        1 { "OK (Nur Computer)" }
                        2 { "OK (Nur Benutzer)" }
                        0 { "Nicht OK (Beide aktiviert)" }
                        3 { "Nicht OK (Vollstaendig deaktiviert)" }
                        default { "Nicht OK (Unbekannt: $flags)" }
                    }
                }

                $isLinked = $false
                $linkedCount = 0
                if ($gpoLinksCache.ContainsKey($guid) -and $gpoLinksCache[$guid].Count -gt 0) {
                    $isLinked = $true
                    $linkedCount = $gpoLinksCache[$guid].Count
                }

                $wmiFilterName = "-"
                $wmiFilterQuery = "-"
                if ($g.Properties["gpcwqlfilter"]) {
                    $rawWmi = $g.Properties["gpcwqlfilter"][0]
                    if ($rawWmi -match "({[a-fA-F0-9-]+})") {
                        $wmiGuid = $matches[1]
                        if ($wmiMap.ContainsKey($wmiGuid)) {
                            $wmiFilterName = $wmiMap[$wmiGuid].Name
                            $wmiFilterQuery = $wmiMap[$wmiGuid].Query
                        } else { $wmiFilterName = $wmiGuid }
                    } else { $wmiFilterName = $rawWmi }
                }

                $created = if ($g.Properties["whencreated"]) { (Get-Date $g.Properties["whencreated"][0]).ToString("dd.MM.yyyy HH:mm") } else { "-" }
                $changed = if ($g.Properties["whenchanged"]) { (Get-Date $g.Properties["whenchanged"][0]).ToString("dd.MM.yyyy HH:mm") } else { "-" }

                $rawOverviewList.Add([PSCustomObject]@{
                    "GPO Name"      = $displayName
                    "Gesamt-Status" = $overallStatus
                    "Verlinkt"      = if ($isLinked) { "Ja" } else { "Nein" }
                    "Link-Anzahl"   = $linkedCount
                    "Benutzer"      = $userStatus
                    "Computer"      = $compStatus
                    "WMI-Filter"    = $wmiFilterName
                    "WMI Query"     = $wmiFilterQuery
                    "GUID"          = $guid
                    "Erstellt am"   = $created
                    "Geaendert am"  = $changed
                })
            }

            Update-OverviewDisplay
        } catch {
            $lblLegendOverview.Text = "Fehler: $($_.Exception.Message)"
        } finally {
            if ($wmiResults)   { $wmiResults.Dispose() }
            if ($wmiSearcher)  { $wmiSearcher.Dispose() }
            if ($ouResults)    { $ouResults.Dispose() }
            if ($linkSearcher) { $linkSearcher.Dispose() }
            if ($gpoResults)   { $gpoResults.Dispose() }
            if ($gpoSearcher)  { $gpoSearcher.Dispose() }
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }

    $gridOvMaster.Add_DataBindingComplete({
        if ($form.IsDisposed -or $gridOvMaster.IsDisposed) { return }
        foreach ($row in $gridOvMaster.Rows) {
            $status = [string]$row.Cells["Gesamt-Status"].Value
            $gpoName = [string]$row.Cells["GPO Name"].Value

            if ($status -match "Sonderstellung" -or $gpoName -match "^Default Domain Policy$" -or $gpoName -match "^Default Domain Controllers Policy$") {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(232, 238, 255)
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(0, 45, 135)
                $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(200, 215, 255)
                $row.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
                $row.DefaultCellStyle.Font = New-Object System.Drawing.Font($gridOvMaster.Font, [System.Drawing.FontStyle]::Bold)
            }
            elseif ($status -match "^Nicht OK") {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::MistyRose
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkRed
                $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::Salmon
                $row.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
                $row.DefaultCellStyle.Font = New-Object System.Drawing.Font($gridOvMaster.Font, [System.Drawing.FontStyle]::Bold)
            }
            else {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(235, 247, 235)
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkGreen
                $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(200, 235, 200)
                $row.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
            }
        }
    })

    $btnRefreshOverview.Add_Click($loadOverviewAction)
    $comboViewMode.Add_SelectedIndexChanged({ Update-OverviewDisplay })
    $txtOverviewSearch.Add_TextChanged({ Update-OverviewDisplay })

    $gridOvMaster.Add_SelectionChanged({
        if ($form.IsDisposed -or $gridOvMaster.IsDisposed -or $gridOvDetails.IsDisposed) { return }
        if ($gridOvMaster.SelectedRows.Count -gt 0) {
            $selectedRow = $gridOvMaster.SelectedRows[0]
            $guid = [string]$selectedRow.Cells["GUID"].Value
            $gName = [string]$selectedRow.Cells["GPO Name"].Value

            $lblOvDetailsTitle.Text = "Verlinkungsziele fuer: [$gName]"

            $arrDetails = [System.Collections.ArrayList]::new()
            if ($guid -and $gpoLinksCache.ContainsKey($guid) -and $gpoLinksCache[$guid].Count -gt 0) {
                foreach ($dn in $gpoLinksCache[$guid]) {
                    $type = "Organizational Unit (OU)"
                    $simpleName = $dn
                    if ($dn -match "^OU=([^,]+)") {
                        $simpleName = $matches[1]
                        $type = "OU"
                    } elseif ($dn -match "^DC=") {
                        $type = "Domaenen-Root"
                        $simpleName = $domainName
                    }
                    [void]$arrDetails.Add([PSCustomObject]@{
                        "Typ"                = $type
                        "Name / Ziel"        = $simpleName
                        "DistinguishedName"  = $dn
                    })
                }
            } else {
                [void]$arrDetails.Add([PSCustomObject]@{
                    "Typ"                = "Info"
                    "Name / Ziel"        = "-- Keine Verknuepfung --"
                    "DistinguishedName"  = "[Hinweis] Diese GPO ist aktuell nirgendwo verlinkt (Unlinked)."
                })
            }
            $gridOvDetails.DataSource = $arrDetails
        }
    })

    $btnExportOverviewCsv.Add_Click({
        if ($gridOvMaster.Rows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Keine Daten zum Exportieren vorhanden.", "Hinweis", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        $targetBase = $txtBackupTargetDir.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($targetBase)) { $targetBase = "C:\Install\Backup\GPO" }
        if (-not (Test-Path $targetBase)) { New-Item -ItemType Directory -Path $targetBase -Force | Out-Null }

        $dateStr = Get-Date -Format "yyyyMMdd"
        $timeStr = Get-Date -Format "HHmm"
        $csvFile = Join-Path $targetBase "GPO_Overview_Export_${dateStr}_${timeStr}.csv"

        $enriched = foreach ($row in @($gridOvMaster.DataSource)) {
            $gGuid = $row.GUID
            $linksStr = if ($gpoLinksCache.ContainsKey($gGuid)) { ($gpoLinksCache[$gGuid] -join " | ") } else { "Keine" }
            [PSCustomObject]@{
                "GPO Name"         = $row."GPO Name"
                "Gesamt-Status"    = $row."Gesamt-Status"
                "Verlinkt"         = $row."Verlinkt"
                "Link-Anzahl"      = $row."Link-Anzahl"
                "Benutzer"         = $row."Benutzer"
                "Computer"         = $row."Computer"
                "Verlinkungsziele" = $linksStr
                "WMI-Filter"       = $row."WMI-Filter"
                "WMI Query"        = $row."WMI Query"
                "GUID"             = $row."GUID"
                "Erstellt am"      = $row."Erstellt am"
                "Geaendert am"     = $row."Geaendert am"
            }
        }

        $enriched | Export-Csv -Path $csvFile -Delimiter ";" -NoTypeInformation -Encoding UTF8
        [System.Windows.Forms.MessageBox]::Show("Uebersichtsbericht erfolgreich exportiert!`n`nPfad: $csvFile", "Export abgeschlossen", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    })

    # =========================================================================
    # PRAEZISER XML-PARSER (Trennt Parameter-Werte strikt von Explain-Texten)
    # =========================================================================
    function Get-ParsedGpoSettings {
        param([string]$GpoId, [string]$GpoDisplayName)

        $dateStr = Get-Date -Format "yyyyMMdd"
        $timeStr = Get-Date -Format "HHmm"
        $list = [System.Collections.Generic.List[PSCustomObject]]::new()

        [xml]$xml = Get-GPOReport -Guid $GpoId -ReportType Xml -ErrorAction Stop

        $parseSection = {
            param($sectionNode, $scope)
            if ($null -eq $sectionNode -or -not $sectionNode.ExtensionData) { return }

            foreach ($ext in $sectionNode.ExtensionData.Extension) {
                $extType = if ($ext.type) { $ext.type } else { $ext.LocalName }
                $extCategory = if ($ext.Name) { $ext.Name } else { "Erweiterung" }

                # 1. Administrative Vorlagen & Richtlinien (<Policy>)
                $policies = $ext.SelectNodes(".//*[local-name()='Policy']")
                if ($policies -and $policies.Count -gt 0) {
                    foreach ($p in $policies) {
                        $pName = if ($p.SelectSingleNode("./*[local-name()='Name']")) { $p.SelectSingleNode("./*[local-name()='Name']").InnerText.Trim() } elseif ($p.Name) { $p.Name.Trim() } else { "Unbenannte Richtlinie" }
                        $rawState = if ($p.SelectSingleNode("./*[local-name()='State']")) { $p.SelectSingleNode("./*[local-name()='State']").InnerText.Trim() } elseif ($p.State) { $p.State.Trim() } else { "Enabled" }
                        $pState = switch ($rawState) { "Enabled" { "Aktiviert" } "Disabled" { "Deaktiviert" } default { $rawState } }
                        $pCat = if ($p.SelectSingleNode("./*[local-name()='Category']")) { $p.SelectSingleNode("./*[local-name()='Category']").InnerText.Trim() } else { $extCategory }
                        
                        $pSupported = if ($p.SelectSingleNode("./*[local-name()='Supported' or local-name()='SupportedOn']")) {
                            $p.SelectSingleNode("./*[local-name()='Supported' or local-name()='SupportedOn']").InnerText.Trim()
                        } else { "Keine Angabe" }

                        $pExplain = if ($p.SelectSingleNode("./*[local-name()='Explain' or local-name()='ExplainText']")) {
                            $p.SelectSingleNode("./*[local-name()='Explain' or local-name()='ExplainText']").InnerText.Trim()
                        } else { "Keine Erklaerung in der Richtlinienvorlage hinterlegt." }

                        $ignoredTags = @("Name", "State", "Explain", "ExplainText", "Supported", "SupportedOn", "Category", "Text")
                        $paramValues = @()

                        foreach ($child in $p.ChildNodes) {
                            if ($child.LocalName -in $ignoredTags) { continue }

                            $valNodes = $child.SelectNodes(".//*[local-name()='Value' or local-name()='Data' or local-name()='Setting' or local-name()='Decimal' or local-name()='String']")
                            $nameNode = $child.SelectSingleNode("./*[local-name()='Name' or local-name()='Label']")
                            $optLabel = if ($nameNode) { $nameNode.InnerText.Trim() } else { "" }

                            $extractedList = @()
                            if ($valNodes -and $valNodes.Count -gt 0) {
                                foreach ($vn in $valNodes) {
                                    if (-not [string]::IsNullOrWhiteSpace($vn.InnerText)) {
                                        $extractedList += $vn.InnerText.Trim()
                                    }
                                }
                            }
                            elseif ($child.Attributes["value"]) {
                                $extractedList += $child.Attributes["value"].Value.Trim()
                            }
                            elseif (-not [string]::IsNullOrWhiteSpace($child.InnerText) -and $child.ChildNodes.Count -le 1) {
                                $extractedList += $child.InnerText.Trim()
                            }

                            if ($extractedList.Count -gt 0) {
                                $joinedVals = $extractedList -join ", "
                                if (-not [string]::IsNullOrWhiteSpace($optLabel) -and $optLabel -ne $joinedVals) {
                                    $paramValues += "$($optLabel): $joinedVals"
                                } else {
                                    $paramValues += "$joinedVals"
                                }
                            }
                        }

                        $cleanValue = if ($paramValues.Count -gt 0) { $paramValues -join " | " } else { $pState }

                        $list.Add([PSCustomObject]@{
                            Scope     = $scope
                            Category  = $pCat
                            Name      = $pName
                            Value     = $cleanValue
                            State     = $pState
                            Supported = $pSupported
                            Explain   = $pExplain
                            GpoName   = $GpoDisplayName
                            Datum     = $dateStr
                            Uhrzeit   = $timeStr
                        })
                    }
                }

                # 2. Systemdienste (Print Spooler etc.)
                $services = $ext.SelectNodes(".//*[local-name()='SystemServices'] | .//*[local-name()='Service']")
                if ($services -and $services.Count -gt 0) {
                    foreach ($svc in $services) {
                        $svcName = if ($svc.SelectSingleNode("./*[local-name()='Display']/*[local-name()='Name']")) {
                            $svc.SelectSingleNode("./*[local-name()='Display']/*[local-name()='Name']").InnerText.Trim()
                        } elseif ($svc.SelectSingleNode("./*[local-name()='Name']")) {
                            $svc.SelectSingleNode("./*[local-name()='Name']").InnerText.Trim()
                        } else { "Systemdienst" }

                        $mode = if ($svc.SelectSingleNode("./*[local-name()='StartupMode']")) {
                            $svc.SelectSingleNode("./*[local-name()='StartupMode']").InnerText.Trim()
                        } elseif ($svc.SelectSingleNode("./*[local-name()='Mode']")) {
                            $svc.SelectSingleNode("./*[local-name()='Mode']").InnerText.Trim()
                        } else { "Konfiguriert" }

                        $modeDE = switch ($mode) {
                            "Disabled"  { "Deaktiviert (Disabled)" }
                            "Automatic" { "Automatisch (Automatic)" }
                            "Manual"    { "Manuell (Manual)" }
                            default     { $mode }
                        }

                        $list.Add([PSCustomObject]@{
                            Scope     = $scope
                            Category  = "Sicherheitseinstellungen / Systemdienste"
                            Name      = $svcName
                            Value     = "Starttyp: $modeDE"
                            State     = $modeDE
                            Supported = "Windows Systemdienste"
                            Explain   = "Startmodus fuer den Windows-Dienst '$svcName': $modeDE"
                            GpoName   = $GpoDisplayName
                            Datum     = $dateStr
                            Uhrzeit   = $timeStr
                        })
                    }
                }

                # 3. Group Policy Preferences (GPP)
                $gppNodes = $ext.SelectNodes(".//*[local-name()='Properties']")
                if ($gppNodes -and $gppNodes.Count -gt 0) {
                    foreach ($prop in $gppNodes) {
                        $parent = $prop.ParentNode
                        $itemType = $parent.LocalName
                        
                        $gppCategory = switch ($itemType) {
                            "Drive"               { "Praeferenzen (GPP) / Laufwerkszuordnungen" }
                            "Shortcut"            { "Praeferenzen (GPP) / Verknuepfungen" }
                            "File"                { "Praeferenzen (GPP) / Dateien" }
                            "Folder"              { "Praeferenzen (GPP) / Ordner" }
                            "Registry"            { "Praeferenzen (GPP) / Registrierung" }
                            "TaskV2"              { "Praeferenzen (GPP) / Geplante Aufgaben" }
                            "ImmediateTaskV2"     { "Praeferenzen (GPP) / Sofortige Aufgaben" }
                            "Group"               { "Praeferenzen (GPP) / Lokale Gruppen" }
                            "User"                { "Praeferenzen (GPP) / Lokale Benutzer" }
                            "SharedPrinter"       { "Praeferenzen (GPP) / Netzwerkdrucker" }
                            "LocalPrinter"        { "Praeferenzen (GPP) / Lokale Drucker" }
                            "EnvironmentVariable" { "Praeferenzen (GPP) / Umgebungsvariablen" }
                            default               { "Praeferenzen (GPP) / $itemType" }
                        }

                        $itemName = if ($parent.Attributes["name"]) {
                            $parent.Attributes["name"].Value
                        } elseif ($prop.Attributes["name"]) {
                            $prop.Attributes["name"].Value
                        } elseif ($prop.Attributes["path"]) {
                            $prop.Attributes["path"].Value
                        } elseif ($prop.Attributes["letter"]) {
                            "Laufwerk $($prop.Attributes['letter'].Value):"
                        } else { "GPP $itemType" }

                        $actionCode = if ($prop.Attributes["action"]) { $prop.Attributes["action"].Value } else { "U" }
                        $actionText = switch ($actionCode) {
                            "C" { "Erstellen (Create)" }
                            "U" { "Aktualisieren (Update)" }
                            "R" { "Ersetzen (Replace)" }
                            "D" { "Loeschen (Delete)" }
                            default { $actionCode }
                        }

                        $valParts = @("Aktion: $actionText")
                        if ($prop.Attributes["path"])       { $valParts += "Pfad: $($prop.Attributes['path'].Value)" }
                        if ($prop.Attributes["targetPath"]) { $valParts += "Ziel: $($prop.Attributes['targetPath'].Value)" }
                        if ($prop.Attributes["fromPath"])   { $valParts += "Quelle: $($prop.Attributes['fromPath'].Value)" }
                        if ($prop.Attributes["location"])   { $valParts += "Ort: $($prop.Attributes['location'].Value)" }
                        if ($prop.Attributes["hive"])       { $valParts += "$($prop.Attributes['hive'].Value)\$($prop.Attributes['key'].Value)\$($prop.Attributes['name'].Value)" }
                        if ($prop.Attributes["value"])      { $valParts += "Wert: $($prop.Attributes['value'].Value) ($($prop.Attributes['type'].Value))" }
                        if ($prop.Attributes["groupName"])  { $valParts += "Gruppe: $($prop.Attributes['groupName'].Value)" }
                        if ($prop.Attributes["userName"])   { $valParts += "Benutzer: $($prop.Attributes['userName'].Value)" }

                        $cleanGppVal = $valParts -join " | "

                        $descLines = @("GPP OBJEKT: $itemName", "TYP:        $itemType", "AKTION:     $actionText", "--------------------------------------------------")
                        foreach ($att in $prop.Attributes) {
                            $descLines += " - $($att.Name): $($att.Value)"
                        }

                        $list.Add([PSCustomObject]@{
                            Scope     = $scope
                            Category  = $gppCategory
                            Name      = $itemName
                            Value     = $cleanGppVal
                            State     = $actionText
                            Supported = "Gruppenrichtlinien-Praeferenzen (GPP)"
                            Explain   = $descLines -join "`r`n"
                            GpoName   = $GpoDisplayName
                            Datum     = $dateStr
                            Uhrzeit   = $timeStr
                        })
                    }
                }

                # 4. Sicherheitsoptionen & Audit
                $secOptions = $ext.SelectNodes(".//*[local-name()='SecurityOptions']/* | .//*[local-name()='Account']/* | .//*[local-name()='KerberosPolicy']/* | .//*[local-name()='Audit']/* | .//*[local-name()='AuditPolicy']/*")
                if ($secOptions -and $secOptions.Count -gt 0) {
                    foreach ($sec in $secOptions) {
                        $secName = if ($sec.SelectSingleNode("./*[local-name()='Display']/*[local-name()='Name']")) {
                            $sec.SelectSingleNode("./*[local-name()='Display']/*[local-name()='Name']").InnerText.Trim()
                        } elseif ($sec.SelectSingleNode("./*[local-name()='Name']")) {
                            $sec.SelectSingleNode("./*[local-name()='Name']").InnerText.Trim()
                        } else { $sec.LocalName }

                        $secVal = if ($sec.SelectSingleNode("./*[local-name()='Display']/*[local-name()='DisplayString']")) {
                            $sec.SelectSingleNode("./*[local-name()='Display']/*[local-name()='DisplayString']").InnerText.Trim()
                        } elseif ($sec.SelectSingleNode("./*[local-name()='SettingNumber']")) {
                            $sec.SelectSingleNode("./*[local-name()='SettingNumber']").InnerText.Trim()
                        } elseif ($sec.SelectSingleNode("./*[local-name()='SettingBoolean']")) {
                            $sec.SelectSingleNode("./*[local-name()='SettingBoolean']").InnerText.Trim()
                        } else { $sec.InnerText.Trim() }

                        if (-not [string]::IsNullOrWhiteSpace($secVal) -and $secVal -ne $secName) {
                            $list.Add([PSCustomObject]@{
                                Scope     = $scope
                                Category  = "Sicherheitseinstellungen / Sicherheitsoptionen"
                                Name      = $secName
                                Value     = $secVal
                                State     = "Konfiguriert"
                                Supported = "Windows Sicherheitsrichtlinie"
                                Explain   = "Sicherheitsoption im Bereich $scope.`r`nRichtlinie: $secName"
                                GpoName   = $GpoDisplayName
                                Datum     = $dateStr
                                Uhrzeit   = $timeStr
                            })
                        }
                    }
                }

                # 5. Benutzerrechte (User Rights)
                $userRights = $ext.SelectNodes(".//*[local-name()='UserRightsAssignment']")
                if ($userRights -and $userRights.Count -gt 0) {
                    foreach ($ur in $userRights) {
                        $rightName = if ($ur.SelectSingleNode("./*[local-name()='Name']")) { $ur.SelectSingleNode("./*[local-name()='Name']").InnerText.Trim() } else { $ur.LocalName }
                        $members = ($ur.SelectNodes(".//*[local-name()='Member']/*[local-name()='Name']").InnerText -join ", ")
                        if (-not [string]::IsNullOrWhiteSpace($members)) {
                            $list.Add([PSCustomObject]@{
                                Scope     = $scope
                                Category  = "Sicherheitseinstellungen / Zuweisen von Benutzerrechten"
                                Name      = $rightName
                                Value     = $members
                                State     = "Zugewiesen"
                                Supported = "Benutzerrechte-Richtlinie"
                                Explain   = "Zugewiesene Konten / Gruppen fuer '$rightName':`r`n$members"
                                GpoName   = $GpoDisplayName
                                Datum     = $dateStr
                                Uhrzeit   = $timeStr
                            })
                        }
                    }
                }

                # 6. Registry-Einstellungen
                $regNodes = $ext.SelectNodes(".//*[local-name()='RegistrySetting']")
                if ($regNodes -and $regNodes.Count -gt 0) {
                    foreach ($reg in $regNodes) {
                        $keyPath = if ($reg.SelectSingleNode("./*[local-name()='KeyPath']")) { $reg.SelectSingleNode("./*[local-name()='KeyPath']").InnerText.Trim() } else { "" }
                        $valName = if ($reg.SelectSingleNode("./*[local-name()='ValueName']")) { $reg.SelectSingleNode("./*[local-name()='ValueName']").InnerText.Trim() } else { "(Standard)" }
                        $valData = if ($reg.SelectSingleNode("./*[local-name()='Value']")) { $reg.SelectSingleNode("./*[local-name()='Value']").InnerText.Trim() } else { "" }
                        $regType = if ($reg.SelectSingleNode("./*[local-name()='Type']")) { $reg.SelectSingleNode("./*[local-name()='Type']").InnerText.Trim() } else { "REG_SZ" }

                        $list.Add([PSCustomObject]@{
                            Scope     = $scope
                            Category  = "Registry-Richtlinie"
                            Name      = "$keyPath\$valName"
                            Value     = "$valData ($regType)"
                            State     = "Aktiviert"
                            Supported = "Registry-Eintrag"
                            Explain   = "Direkter Registry-Eintrag:`r`nPfad: $keyPath`r`nName: $valName`r`nTyp:  $regType`r`nWert: $valData"
                            GpoName   = $GpoDisplayName
                            Datum     = $dateStr
                            Uhrzeit   = $timeStr
                        })
                    }
                }
            }
        }

        & $parseSection $xml.GPO.Computer "Computer"
        & $parseSection $xml.GPO.User "User"

        return $list
    }

    # =========================================================================
    # LOGIK TAB 2: GPO Settings Inspector
    # =========================================================================
    function Update-SettingsGridDisplay {
        if ($form.IsDisposed -or $gridSettings.IsDisposed) { return }
        $filterText = $txtFilter.Text.Trim()
        $gridSettings.Rows.Clear()

        $filteredItems = if ([string]::IsNullOrWhiteSpace($filterText)) {
            $rawSettingsList
        } else {
            $rawSettingsList | Where-Object {
                $_.Name -like "*$filterText*" -or 
                $_.Category -like "*$filterText*" -or 
                $_.Value -like "*$filterText*" -or 
                $_.Scope -like "*$filterText*"
            }
        }

        $filteredArray = @($filteredItems)
        foreach ($item in $filteredArray) {
            [void]$gridSettings.Rows.Add(
                $item.Scope,
                $item.Category,
                $item.Name,
                $item.Value,
                $item.State,
                $item.Supported,
                $item.Explain,
                $item.GpoName
            )
        }

        if ($filteredArray.Count -gt 0) {
            $lblSettingsStatus.Text = "$($filteredArray.Count) Einstellung(en) geladen."
            $gridSettings.Rows[0].Selected = $true
        } else {
            $lblSettingsStatus.Text = "Keine Einstellungen gefunden."
            $txtDescription.Text = "Keine Eintraege fuer die Auswahl vorhanden."
        }
    }

    $loadSettingsAction = {
        $selectedGpoName = $comboGpo.SelectedItem
        if ([string]::IsNullOrWhiteSpace($selectedGpoName)) { return }

        $lblSettingsStatus.Text = "Lese Einstellungen ein..."
        $form.Refresh()

        $rawSettingsList.Clear()

        if ($selectedGpoName -eq "-- ALLE GPOs laden --") {
            $gpoList = Get-GPO -All
            foreach ($g in $gpoList) {
                try {
                    $items = Get-ParsedGpoSettings -GpoId $g.Id -GpoDisplayName $g.DisplayName
                    foreach ($item in $items) { $rawSettingsList.Add($item) }
                } catch {}
            }
        } else {
            try {
                $gpo = Get-GPO -Name $selectedGpoName -ErrorAction Stop
                $items = Get-ParsedGpoSettings -GpoId $gpo.Id -GpoDisplayName $gpo.DisplayName
                foreach ($item in $items) { $rawSettingsList.Add($item) }
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Fehler beim Abruf: $_", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }

        Update-SettingsGridDisplay
    }

    $btnLoadSettings.Add_Click($loadSettingsAction)
    $txtFilter.Add_TextChanged({ Update-SettingsGridDisplay })

    $gridSettings.Add_SelectionChanged({
        if ($form.IsDisposed -or $gridSettings.IsDisposed -or $txtDescription.IsDisposed) { return }
        if ($gridSettings.SelectedRows.Count -gt 0) {
            $row = $gridSettings.SelectedRows[0]
            $detailText  = "RICHTLINIE : $($row.Cells['colName'].Value)`r`n"
            $detailText += "GPO        : $($row.Cells['colGpo'].Value)`r`n"
            $detailText += "BEREICH    : $($row.Cells['colScope'].Value)-Konfiguration`r`n"
            $detailText += "PFAD       : $($row.Cells['colCategory'].Value)`r`n"
            $detailText += "STATUS     : $($row.Cells['colState'].Value)`r`n"
            $detailText += "WERT       : $($row.Cells['colValue'].Value)`r`n"
            $detailText += "KOMPATIBEL : $($row.Cells['colSupported'].Value)`r`n"
            $detailText += "======================================================================`r`n`r`n"
            $detailText += "[ERLAEUTERUNG / BESCHREIBUNG]`r`n"
            $detailText += [string]$row.Cells['colExplain'].Value

            $txtDescription.Text = $detailText
        }
    })

    $btnExportCsv.Add_Click({
        if ($rawSettingsList.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Keine Daten zum Exportieren vorhanden.", "Hinweis", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $targetBase = $txtBackupTargetDir.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($targetBase)) { $targetBase = "C:\Install\Backup\GPO" }
        if (-not (Test-Path $targetBase)) { New-Item -ItemType Directory -Path $targetBase -Force | Out-Null }

        $dateStr = Get-Date -Format "yyyyMMdd"
        $timeStr = Get-Date -Format "HHmm"
        $csvFile = Join-Path $targetBase "GPO_Settings_Export_${dateStr}_${timeStr}.csv"

        $rawSettingsList | Export-Csv -Path $csvFile -Delimiter ";" -NoTypeInformation -Encoding UTF8
        [System.Windows.Forms.MessageBox]::Show("Erfolgreich $($rawSettingsList.Count) Eintraege exportiert!`n`nPfad: $csvFile", "Export abgeschlossen", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    })

    # =========================================================================
    # LOGIK TAB 3: GPO Backup & Audit
    # =========================================================================
    $btnBrowseFolder.Add_Click({
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description = "Waehlen Sie das Basis-Backup-Verzeichnis fuer die GPOs aus:"
        $dialog.SelectedPath = $txtBackupTargetDir.Text.Trim()
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtBackupTargetDir.Text = $dialog.SelectedPath
        }
    })

    $loadGposAction = {
        if ($form.IsDisposed -or $gridGpos.IsDisposed) { return }
        $gridGpos.DataSource = $null
        $tableGpos = New-Object System.Data.DataTable
        [void]$tableGpos.Columns.Add("GPO Name")
        [void]$tableGpos.Columns.Add("Status")
        [void]$tableGpos.Columns.Add("Verknuepft")
        [void]$tableGpos.Columns.Add("GPO ID (GUID)")

        $allGpos = Get-GPO -All | Sort-Object DisplayName
        foreach ($g in $allGpos) {
            $report = [xml](Get-GPOReport -Guid $g.Id -ReportType Xml)
            $linkCount = @($report.GPO.LinksTo).Count
            $isLinked = if ($linkCount -gt 0) { "JA ($linkCount)" } else { "NEIN (Unlinked)" }

            [void]$tableGpos.Rows.Add($g.DisplayName, $g.GpoStatus, $isLinked, $g.Id.ToString())
        }
        $gridGpos.DataSource = $tableGpos
    }

    $btnLoadGpos.Add_Click($loadGposAction)

    $gridGpos.Add_SelectionChanged({
        if ($form.IsDisposed -or $gridGpos.IsDisposed -or $gridLinks.IsDisposed) { return }
        if ($gridGpos.SelectedRows.Count -gt 0) {
            $guid = $gridGpos.SelectedRows[0].Cells["GPO ID (GUID)"].Value
            $name = $gridGpos.SelectedRows[0].Cells["GPO Name"].Value
            $lblLinks.Text = "2. Verknuepfungs-Ziele fuer: [$name]"

            $tableLinks = New-Object System.Data.DataTable
            [void]$tableLinks.Columns.Add("Verknuepfungs-Ziel (SOMPath)")
            [void]$tableLinks.Columns.Add("Aktiviert (LinkEnabled)")
            [void]$tableLinks.Columns.Add("Keine Vererbung (Enforced)")

            try {
                $report = [xml](Get-GPOReport -Guid $guid -ReportType Xml)
                $links = $report.GPO.LinksTo
                if ($links) {
                    foreach ($l in $links) {
                        [void]$tableLinks.Rows.Add($l.SOMPath, $l.Enabled, $l.NoOverride)
                    }
                } else {
                    [void]$tableLinks.Rows.Add("-- Keine Verknuepfung vorhanden --", "-", "-")
                }
            } catch {
                [void]$tableLinks.Rows.Add("Fehler beim Abruf", "-", "-")
            }
            $gridLinks.DataSource = $tableLinks
        }
    })

    function Backup-SingleGPOWithLog {
        param([string]$gpoGuid, [string]$gpoName, [string]$basePath)
        $safeName = $gpoName -replace '[\\/:*?"<>|]', '_'
        $timestamp = Get-Date -Format "yyyyMMdd_HHmm"
        $targetPath = Join-Path (Join-Path $basePath $safeName) $timestamp

        if (-not (Test-Path $targetPath)) { New-Item -ItemType Directory -Path $targetPath -Force | Out-Null }
        Backup-GPO -Guid $gpoGuid -Path $targetPath | Out-Null

        $txtPath = Join-Path $targetPath "GPO_Link_Info.txt"
        $report = [xml](Get-GPOReport -Guid $gpoGuid -ReportType Xml)
        $links = $report.GPO.LinksTo
        $timeFormatted = Get-Date -Format "dd.MM.yyyy HH:mm:ss"

        "==================================================" | Out-File -FilePath $txtPath -Encoding UTF8
        " GPO AUDIT & BACKUP REPORT"                         | Out-File -FilePath $txtPath -Append -Encoding UTF8
        "==================================================" | Out-File -FilePath $txtPath -Append -Encoding UTF8
        "GPO Name        : $gpoName"                          | Out-File -FilePath $txtPath -Append -Encoding UTF8
        "GPO ID (GUID)   : $gpoGuid"                          | Out-File -FilePath $txtPath -Append -Encoding UTF8
        "Zeitpunkt       : $timeFormatted"                    | Out-File -FilePath $txtPath -Append -Encoding UTF8
        "Ziel-Ordner     : $targetPath"                       | Out-File -FilePath $txtPath -Append -Encoding UTF8
        "--------------------------------------------------" | Out-File -FilePath $txtPath -Append -Encoding UTF8
        if ($links) {
            "Verknuepfungen:" | Out-File -FilePath $txtPath -Append -Encoding UTF8
            foreach ($l in $links) {
                " - Ziel (SOM): $($l.SOMPath) | Aktiv: $($l.Enabled) | Enforced: $($l.NoOverride)" | Out-File -FilePath $txtPath -Append -Encoding UTF8
            }
        } else {
            "Status: Unverknuepft (Unlinked GPO)" | Out-File -FilePath $txtPath -Append -Encoding UTF8
        }

        $txtLog.AppendText("[$((Get-Date).ToString('HH:mm:ss'))] [OK] Gesichert: '$gpoName' -> $targetPath`r`n")
    }

    $btnBackupSelected.Add_Click({
        $basePath = $txtBackupTargetDir.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($basePath)) {
            [System.Windows.Forms.MessageBox]::Show("Bitte geben Sie einen gueltigen Zielpfad ein.", "Pfad fehlt", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        if ($gridGpos.SelectedRows.Count -gt 0) {
            $guid = $gridGpos.SelectedRows[0].Cells["GPO ID (GUID)"].Value
            $name = $gridGpos.SelectedRows[0].Cells["GPO Name"].Value
            
            $txtLog.AppendText("[$((Get-Date).ToString('HH:mm:ss'))] Starte Einzelsicherung fuer '$name'...`r`n")
            try {
                Backup-SingleGPOWithLog -gpoGuid $guid -gpoName $name -basePath $basePath
                [System.Windows.Forms.MessageBox]::Show("Backup fuer GPO '$name' erfolgreich abgeschlossen!`n`nOrdner: $basePath\$name", "Sicherung OK", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } catch {
                $txtLog.AppendText("[$((Get-Date).ToString('HH:mm:ss'))] [FEHLER] bei '$name': $_`r`n")
                [System.Windows.Forms.MessageBox]::Show("Fehler beim Backup: $_", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("Bitte waehlen Sie zuerst eine GPO in der Tabelle aus.", "Hinweis", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    })

    $btnBackupAll.Add_Click({
        $basePath = $txtBackupTargetDir.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($basePath)) {
            [System.Windows.Forms.MessageBox]::Show("Bitte geben Sie einen gueltigen Zielpfad ein.", "Pfad fehlt", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        $allGpos = Get-GPO -All
        $confirm = [System.Windows.Forms.MessageBox]::Show("Moechten Sie wirklich alle $($allGpos.Count) Gruppenrichtlinien nach '$basePath' sichern?", "Bestaetigung", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        $txtLog.AppendText("[$((Get-Date).ToString('HH:mm:ss'))] === Starte Gesamtsicherung aller $($allGpos.Count) GPOs ===`r`n")
        $form.Refresh()

        foreach ($g in $allGpos) {
            try {
                Backup-SingleGPOWithLog -gpoGuid $g.Id -gpoName $g.DisplayName -basePath $basePath
            } catch {
                $txtLog.AppendText("[$((Get-Date).ToString('HH:mm:ss'))] [FEHLER] bei '$($g.DisplayName)': $_`r`n")
            }
        }
        $txtLog.AppendText("[$((Get-Date).ToString('HH:mm:ss'))] === Gesamtsicherung abgeschlossen! ===`r`n")
        [System.Windows.Forms.MessageBox]::Show("Gesamtsicherung aller $($allGpos.Count) GPOs abgeschlossen!", "Fertig", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    })

    # =========================================================================
    # LOGIK TAB 4: GPO-Vergleich (Diff Engine mit Wert-Vergleich)
    # =========================================================================
    function Update-CompareGridDisplay {
        if ($form.IsDisposed -or $gridCompare.IsDisposed) { return }
        $gridCompare.Rows.Clear()
        $onlyDiffs = $chkOnlyDiffs.Checked
        $diffCount = 0
        $totalCount = $rawCompareList.Count

        foreach ($item in $rawCompareList) {
            if ($item.DiffStatus -ne "Identisch") { $diffCount++ }
        }

        $suppressFilter = ($diffCount -eq 0 -and $totalCount -gt 0)

        foreach ($item in $rawCompareList) {
            $isDiff = $item.DiffStatus -ne "Identisch"
            if ($onlyDiffs -and -not $isDiff -and -not $suppressFilter) { continue }

            $rowIndex = $gridCompare.Rows.Add(
                $item.Scope,
                $item.Category,
                $item.Name,
                $item.StateGPO1,
                $item.ValueGPO1,
                $item.StateGPO2,
                $item.ValueGPO2,
                $item.DiffStatus
            )

            $row = $gridCompare.Rows[$rowIndex]
            switch -Regex ($item.DiffStatus) {
                "Identisch" {
                    $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(235, 247, 235)
                    $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkGreen
                }
                "Abweichend" {
                    $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::LemonChiffon
                    $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkOrange
                    $row.DefaultCellStyle.Font = New-Object System.Drawing.Font($gridCompare.Font, [System.Drawing.FontStyle]::Bold)
                }
                "Nur in GPO 1" {
                    $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::MistyRose
                    $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkRed
                }
                "Nur in GPO 2" {
                    $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::AliceBlue
                    $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkBlue
                }
            }
        }

        if ($diffCount -eq 0 -and $totalCount -gt 0) {
            $lblCompareStatus.Text = "[OK] 0 Unterschiede! Alle $totalCount Richtlinien sind identisch (Gruen)."
        } else {
            $lblCompareStatus.Text = "$diffCount Unterschied(e) bei $totalCount untersuchten Richtlinien."
        }
    }

    $btnCompare.Add_Click({
        $gpoName1 = $comboGpo1.SelectedItem
        $gpoName2 = $comboGpo2.SelectedItem

        if ([string]::IsNullOrWhiteSpace($gpoName1) -or [string]::IsNullOrWhiteSpace($gpoName2)) {
            [System.Windows.Forms.MessageBox]::Show("Bitte waehlen Sie zwei GPOs fuer den Vergleich aus.", "Auswahl unvollstaendig", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        $lblCompareStatus.Text = "Analysiere & Vergleiche..."
        $form.Refresh()
        $rawCompareList.Clear()

        try {
            $gpoObj1 = Get-GPO -Name $gpoName1 -ErrorAction Stop
            $gpoObj2 = Get-GPO -Name $gpoName2 -ErrorAction Stop

            $list1 = Get-ParsedGpoSettings -GpoId $gpoObj1.Id -GpoDisplayName $gpoName1
            $list2 = Get-ParsedGpoSettings -GpoId $gpoObj2.Id -GpoDisplayName $gpoName2

            $dict1 = @{}
            foreach ($item in $list1) { $dict1["$($item.Scope)|$($item.Category)|$($item.Name)"] = $item }

            $dict2 = @{}
            foreach ($item in $list2) { $dict2["$($item.Scope)|$($item.Category)|$($item.Name)"] = $item }

            $allKeys = [System.Collections.Generic.HashSet[string]]::new()
            foreach ($k in $dict1.Keys) { [void]$allKeys.Add($k) }
            foreach ($k in $dict2.Keys) { [void]$allKeys.Add($k) }

            foreach ($k in ($allKeys | Sort-Object)) {
                $item1 = $dict1[$k]
                $item2 = $dict2[$k]

                $scope    = if ($item1) { $item1.Scope } else { $item2.Scope }
                $category = if ($item1) { $item1.Category } else { $item2.Category }
                $name     = if ($item1) { $item1.Name } else { $item2.Name }

                $state1 = if ($item1) { $item1.State } else { "-" }
                $val1   = if ($item1) { $item1.Value } else { "-- [Nicht konfiguriert] --" }

                $state2 = if ($item2) { $item2.State } else { "-" }
                $val2   = if ($item2) { $item2.Value } else { "-- [Nicht konfiguriert] --" }

                $diffStatus = ""
                if ($null -ne $item1 -and $null -ne $item2) {
                    $val1Clean = "$($item1.Value)".Trim()
                    $val2Clean = "$($item2.Value)".Trim()
                    $state1Clean = "$($item1.State)".Trim()
                    $state2Clean = "$($item2.State)".Trim()

                    if ($val1Clean -eq $val2Clean -and $state1Clean -eq $state2Clean) {
                        $diffStatus = "Identisch"
                    }
                    elseif ($state1Clean -ne $state2Clean) {
                        $diffStatus = "Abweichender Status ($state1Clean vs $state2Clean)"
                    }
                    else {
                        $diffStatus = "Abweichender Wert"
                    }
                }
                elseif ($null -ne $item1 -and $null -eq $item2) {
                    $diffStatus = "Nur in GPO 1 ($gpoName1)"
                }
                else {
                    $diffStatus = "Nur in GPO 2 ($gpoName2)"
                }

                $rawCompareList.Add([PSCustomObject]@{
                    Scope      = $scope
                    Category   = $category
                    Name       = $name
                    StateGPO1  = $state1
                    ValueGPO1  = $val1
                    StateGPO2  = $state2
                    ValueGPO2  = $val2
                    DiffStatus = $diffStatus
                    GPO1       = $gpoName1
                    GPO2       = $gpoName2
                })
            }

            Update-CompareGridDisplay
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Vergleich: $_", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $lblCompareStatus.Text = "Fehler beim Einlesen der GPOs."
        }
    })

    $chkOnlyDiffs.Add_CheckedChanged({ Update-CompareGridDisplay })

    $btnExportCompareCsv.Add_Click({
        if ($rawCompareList.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Fuehren Sie zuerst einen Vergleich durch.", "Hinweis", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        $targetBase = $txtBackupTargetDir.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($targetBase)) { $targetBase = "C:\Install\Backup\GPO" }
        if (-not (Test-Path $targetBase)) { New-Item -ItemType Directory -Path $targetBase -Force | Out-Null }

        $dateStr = Get-Date -Format "yyyyMMdd"
        $timeStr = Get-Date -Format "HHmm"
        $gpoClean1 = $comboGpo1.SelectedItem -replace '[\\/:*?"<>|]', '_'
        $gpoClean2 = $comboGpo2.SelectedItem -replace '[\\/:*?"<>|]', '_'
        $csvFile = Join-Path $targetBase "GPO_Compare_${gpoClean1}_vs_${gpoClean2}_${dateStr}_${timeStr}.csv"

        $rawCompareList | Export-Csv -Path $csvFile -Delimiter ";" -NoTypeInformation -Encoding UTF8
        [System.Windows.Forms.MessageBox]::Show("Vergleichsbericht erfolgreich exportiert!`n`nPfad: $csvFile", "Export abgeschlossen", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    })

    # =========================================================================
    # EVENT: POWERSHELL- & TOOL-INFO DIALOG
    # =========================================================================
    $btnPsInfo.Add_Click({
        $psVer    = $PSVersionTable.PSVersion.ToString()
        $psEdit   = if ($PSVersionTable.PSEdition) { $PSVersionTable.PSEdition } else { "Desktop (Windows PowerShell)" }
        $clrVer   = if ($PSVersionTable.CLRVersion) { $PSVersionTable.CLRVersion.ToString() } else { "Core CLR (.NET)" }
        $osVer    = [System.Environment]::OSVersion.VersionString
        $bitProc  = if ([System.Environment]::Is64BitProcess) { "64-Bit" } else { "32-Bit" }
        $execPol  = (Get-ExecutionPolicy).ToString()
        $userCtx  = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        $adMod    = try { (Get-Module -Name ActiveDirectory -ListAvailable | Select-Object -First 1).Version.ToString() } catch { "Nicht geladen" }
        $gpoMod   = try { (Get-Module -Name GroupPolicy -ListAvailable | Select-Object -First 1).Version.ToString() } catch { "Standard" }

        $infoMsg  = "========================================================`r`n"
        $infoMsg += " TOOL 15: GPO ENTERPRISE SUITE`r`n"
        $infoMsg += " Version : $script:ToolVersion`r`n"
        $infoMsg += " Domaene : $domainName`r`n"
        $infoMsg += " Benutzer: $userCtx`r`n"
        $infoMsg += "========================================================`r`n`r`n"
        $infoMsg += "[POWERSHELL UMGEBUNG]`r`n"
        $infoMsg += " - PowerShell Version : $psVer`r`n"
        $infoMsg += " - PowerShell Edition : $psEdit`r`n"
        $infoMsg += " - .NET CLR Version   : $clrVer`r`n"
        $infoMsg += " - Prozess-Architektur: $bitProc`r`n"
        $infoMsg += " - Execution-Policy   : $execPol`r`n`r`n"
        $infoMsg += "[SYSTEM & MODULE]`r`n"
        $infoMsg += " - Betriebssystem     : $osVer`r`n"
        $infoMsg += " - GroupPolicy Modul  : $gpoMod`r`n"
        $infoMsg += " - ActiveDirectory Mod: $adMod`r`n"

        [System.Windows.Forms.MessageBox]::Show($infoMsg, "PowerShell & Tool Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    })

    # =========================================================================
    # INITIALISIERUNG BEIM START & SAUBERES SCHLIESSEN
    # =========================================================================
    $form.Add_Shown({
        $comboGpo.Items.Clear()
        $comboGpo1.Items.Clear()
        $comboGpo2.Items.Clear()

        [void]$comboGpo.Items.Add("-- ALLE GPOs laden --")
        $allGpos = Get-GPO -All | Sort-Object DisplayName
        foreach ($g in $allGpos) {
            [void]$comboGpo.Items.Add($g.DisplayName)
            [void]$comboGpo1.Items.Add($g.DisplayName)
            [void]$comboGpo2.Items.Add($g.DisplayName)
        }

        if ($comboGpo.Items.Count -gt 1) { $comboGpo.SelectedIndex = 1 }
        if ($comboGpo1.Items.Count -gt 0) { $comboGpo1.SelectedIndex = 0 }
        if ($comboGpo2.Items.Count -gt 1) { $comboGpo2.SelectedIndex = 1 }

        & $loadOverviewAction
        & $loadSettingsAction
        & $loadGposAction
    })

    $form.Add_FormClosing({
        $gridOvMaster.DataSource = $null
        $gridOvDetails.DataSource = $null
        $gridSettings.DataSource = $null
        $gridGpos.DataSource = $null
        $gridLinks.DataSource = $null
        $gridCompare.DataSource = $null
    })

    try {
        [void]$form.ShowDialog()
    } finally {
        $form.Dispose()
    }
}

# --- Standalone- & Suite-Aufruf ---
if ($MyInvocation.InvocationName -ne '.' -and $MyInvocation.Line -notmatch 'Open-Tool') {
    Show-Tool15
}
