<#
==================================================================================
 Tool 15: Active Directory GPO Enterprise Suite
 Version: 1.7.4 (Lazy Loading, DDP-Integritaetswarnung, Pfad in Tab 2, ISE-Safe)
 
 Register 1: GPO Uebersicht & Verlinkungs-Analyse
             - Button "GPOs einlesen" mit Live-Fortschrittsbalken & Statustext
             - Automatische Erkennung & Warnung bei DDP-Namensduplikaten / Umbenennung
             - Buttons "Snapshot speichern" & "Snapshot laden" (.gposnap)
             - Schnelle ADSI/LDAP-Abfrage aller GPOs, WMI-Filter & OU-Verlinkungen
             - Farbliche Kennzeichnung (Default GPO / Warnung / OK / Nicht OK)
             - Master-Detail: Zeigt rechts alle Verlinkungsziele (OUs/Domaene/Sites)
 
 Register 2: GPO Richtlinien-Einstellungen & Inspektor
             - Sichtbarer Export-Zielpfad mit Durchsuchen-Dialog
             - Vollstaendiger 6-Komponenten-XML-Parser (ADMX, Services, GPP, Security, Rights, Reg)
             - Synchronisierter Ansichts-Filter (Alle / Nur verlinkte / Unlinked)
             - Lazy Loading (startet sofort ohne Hänger)
 
 Register 3: GPO Backup & Verknuepfungs-Audit
             - Schnelle Bestandsliste ueber LDAP-Cache (kein Netzwerk-Freeze mehr)
             - HTML (Ausgewaehlt): HTML-Bericht der markierten GPO erstellen & oeffnen
             - HTML (Gefilterte): HTML-Massenexport aller gefilterten GPOs in Unterordner
             - Ansichts-Filter: Alle / Nur verlinkte / Nicht verlinkte (Unlinked)
             - Selektives Backup (z.B. nur alle 13 ungelinkten GPOs sichern)
 
 Register 4: GPO-Vergleich (Diff) & DDP-Baseline-Check
             - Button "DDP vs. MS-Standard" fuer 1-Klick Soll/Ist-Vergleich
             - Integrierte Microsoft Werkszustand-Referenz (14 Kontorichtlinien)
             - Intelligente Normalisierung deutscher & englischer Richtliniennamen
             - 2 beliebige GPOs gegeneinander vergleichen
==================================================================================
#>

# 1. Assemblies laden
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.DirectoryServices
Import-Module GroupPolicy -ErrorAction Stop

# 2. VisualStyles initialisieren
try {
    [System.Windows.Forms.Application]::EnableVisualStyles()
} catch {}

$script:ToolVersion = "v1.7.4"
$script:StandardDdpGuid = "31B2F340-016D-11D2-945F-00C04FB984F9"
$script:StandardDdcpGuid = "6AC1786C-016F-11D2-945F-00C04FB984F9"
$script:DdpBaselineName = "[Referenz] Microsoft Default Domain Policy (Standard-Werte)"

function Show-Tool15 {
    [CmdletBinding()]
    param()

    # --- Domaenenpruefung ---
    try {
        $domainInfo = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
        $rootDse = [ADSI]"LDAP://RootDSE"
        $domainDN = $rootDse.defaultNamingContext.Value
        $domainName = $domainInfo.Name
        $rootDse.Dispose()
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Active Directory Domaene nicht erreichbar oder Computer nicht domaenengebunden.",
            "Fehler",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }

    $isClosing = $false

    # --- Hauptfenster ---
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Tool 15 - Active Directory GPO Enterprise Suite ($domainName) - $script:ToolVersion"
    $form.Size = New-Object System.Drawing.Size(1680, 950)
    $form.StartPosition = "CenterScreen"
    $form.MinimumSize = New-Object System.Drawing.Size(1250, 750)
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    # --- Untere Statusleiste fuer Fortschrittsanzeige ---
    $panelBottomStatus = New-Object System.Windows.Forms.Panel
    $panelBottomStatus.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $panelBottomStatus.Height = 32
    $panelBottomStatus.BackColor = [System.Drawing.Color]::FromArgb(240, 242, 246)
    $panelBottomStatus.Padding = New-Object System.Windows.Forms.Padding(10, 4, 10, 4)

    $lblProgressInfo = New-Object System.Windows.Forms.Label
    $lblProgressInfo.Dock = [System.Windows.Forms.DockStyle]::Fill
    $lblProgressInfo.Text = "Bereit."
    $lblProgressInfo.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $lblProgressInfo.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $pbarGlobal = New-Object System.Windows.Forms.ProgressBar
    $pbarGlobal.Dock = [System.Windows.Forms.DockStyle]::Right
    $pbarGlobal.Width = 360
    $pbarGlobal.Visible = $false

    $panelBottomStatus.Controls.Add($lblProgressInfo)
    $panelBottomStatus.Controls.Add($pbarGlobal)

    # --- TabControl mit optimierter Schrift & Polsterung ---
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Dock = [System.Windows.Forms.DockStyle]::Fill
    $tabControl.Font = New-Object System.Drawing.Font("Segoe UI", 10.5, [System.Drawing.FontStyle]::Bold)
    $tabControl.Padding = New-Object System.Drawing.Point(14, 6)

    # =========================================================================
    # REGISTER 1: GPO Uebersicht & Verlinkungs-Analyse
    # =========================================================================
    $tabOverview = New-Object System.Windows.Forms.TabPage
    $tabOverview.Text = "1. GPO Uebersicht & Verlinkungs-Analyse"
    $tabOverview.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)

    $panelOverviewTop = New-Object System.Windows.Forms.Panel
    $panelOverviewTop.Dock = [System.Windows.Forms.DockStyle]::Top
    $panelOverviewTop.Height = 85
    $panelOverviewTop.BackColor = [System.Drawing.Color]::FromArgb(245, 248, 252)
    $panelOverviewTop.Padding = New-Object System.Windows.Forms.Padding(10)

    $lblViewFilter = New-Object System.Windows.Forms.Label
    $lblViewFilter.Text = "Ansicht:"
    $lblViewFilter.Location = New-Object System.Drawing.Point(12, 16)
    $lblViewFilter.AutoSize = $true
    $lblViewFilter.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $comboViewMode = New-Object System.Windows.Forms.ComboBox
    $comboViewMode.Location = New-Object System.Drawing.Point(75, 13)
    $comboViewMode.Size = New-Object System.Drawing.Size(200, 25)
    $comboViewMode.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$comboViewMode.Items.Add("Alle GPOs")
    [void]$comboViewMode.Items.Add("GPOs mit WMI-Filter")
    [void]$comboViewMode.Items.Add("Nicht verlinkte GPOs (Unlinked)")
    $comboViewMode.SelectedIndex = 0

    $lblOverviewSearch = New-Object System.Windows.Forms.Label
    $lblOverviewSearch.Text = "Suche:"
    $lblOverviewSearch.Location = New-Object System.Drawing.Point(285, 16)
    $lblOverviewSearch.AutoSize = $true
    $lblOverviewSearch.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $txtOverviewSearch = New-Object System.Windows.Forms.TextBox
    $txtOverviewSearch.Location = New-Object System.Drawing.Point(335, 13)
    $txtOverviewSearch.Size = New-Object System.Drawing.Size(130, 25)

    $btnLoadAdGpos = New-Object System.Windows.Forms.Button
    $btnLoadAdGpos.Text = "GPOs einlesen"
    $btnLoadAdGpos.Location = New-Object System.Drawing.Point(475, 10)
    $btnLoadAdGpos.Size = New-Object System.Drawing.Size(120, 30)
    $btnLoadAdGpos.BackColor = [System.Drawing.Color]::FromArgb(225, 238, 255)
    $btnLoadAdGpos.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $btnSaveSnapshot = New-Object System.Windows.Forms.Button
    $btnSaveSnapshot.Text = "Snapshot speichern"
    $btnSaveSnapshot.Location = New-Object System.Drawing.Point(602, 10)
    $btnSaveSnapshot.Size = New-Object System.Drawing.Size(140, 30)
    $btnSaveSnapshot.BackColor = [System.Drawing.Color]::FromArgb(255, 245, 230)

    $btnLoadSnapshot = New-Object System.Windows.Forms.Button
    $btnLoadSnapshot.Text = "Snapshot laden"
    $btnLoadSnapshot.Location = New-Object System.Drawing.Point(748, 10)
    $btnLoadSnapshot.Size = New-Object System.Drawing.Size(125, 30)
    $btnLoadSnapshot.BackColor = [System.Drawing.Color]::FromArgb(235, 245, 255)

    $btnExportOverviewCsv = New-Object System.Windows.Forms.Button
    $btnExportOverviewCsv.Text = "CSV Export"
    $btnExportOverviewCsv.Location = New-Object System.Drawing.Point(880, 10)
    $btnExportOverviewCsv.Size = New-Object System.Drawing.Size(100, 30)
    $btnExportOverviewCsv.BackColor = [System.Drawing.Color]::FromArgb(230, 245, 230)

    $btnPsInfo = New-Object System.Windows.Forms.Button
    $btnPsInfo.Text = "PS- & Tool-Info"
    $btnPsInfo.Location = New-Object System.Drawing.Point(987, 10)
    $btnPsInfo.Size = New-Object System.Drawing.Size(115, 30)
    $btnPsInfo.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 245)
    $btnPsInfo.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)

    $lblLegendOverview = New-Object System.Windows.Forms.Label
    $lblLegendOverview.Text = "Legende:  [Blau] Echte Default GPO  |  [Orange] DDP-Integritaetswarnung  |  [Gruen] OK (1 Seite aktiv)  |  [Rot] Nicht OK"
    $lblLegendOverview.Location = New-Object System.Drawing.Point(12, 52)
    $lblLegendOverview.AutoSize = $true
    $lblLegendOverview.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Italic)

    $panelOverviewTop.Controls.Add($lblViewFilter)
    $panelOverviewTop.Controls.Add($comboViewMode)
    $panelOverviewTop.Controls.Add($lblOverviewSearch)
    $panelOverviewTop.Controls.Add($txtOverviewSearch)
    $panelOverviewTop.Controls.Add($btnLoadAdGpos)
    $panelOverviewTop.Controls.Add($btnSaveSnapshot)
    $panelOverviewTop.Controls.Add($btnLoadSnapshot)
    $panelOverviewTop.Controls.Add($btnExportOverviewCsv)
    $panelOverviewTop.Controls.Add($btnPsInfo)
    $panelOverviewTop.Controls.Add($lblLegendOverview)

    $splitOverview = New-Object System.Windows.Forms.SplitContainer
    $splitOverview.Dock = [System.Windows.Forms.DockStyle]::Fill
    $splitOverview.SplitterDistance = 980
    $splitOverview.SplitterWidth = 6

    $panelOvLeft = New-Object System.Windows.Forms.Panel
    $panelOvLeft.Dock = [System.Windows.Forms.DockStyle]::Fill
    $panelOvLeft.Padding = New-Object System.Windows.Forms.Padding(10, 8, 4, 10)

    $lblOvMasterTitle = New-Object System.Windows.Forms.Label
    $lblOvMasterTitle.Text = "Gruppenrichtlinien der Domaene (Klick auf Spaltenkopf zum Sortieren):"
    $lblOvMasterTitle.Dock = [System.Windows.Forms.DockStyle]::Top
    $lblOvMasterTitle.Height = 28
    $lblOvMasterTitle.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)

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
    $gridOvMaster.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    $gridOvMaster.ColumnHeadersHeight = 34
    $gridOvMaster.RowTemplate.Height = 28
    $gridOvMaster.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(249, 251, 254)

    $panelOvLeft.Controls.Add($gridOvMaster)
    $panelOvLeft.Controls.Add($lblOvMasterTitle)
    $splitOverview.Panel1.Controls.Add($panelOvLeft)

    $panelOvRight = New-Object System.Windows.Forms.Panel
    $panelOvRight.Dock = [System.Windows.Forms.DockStyle]::Fill
    $panelOvRight.Padding = New-Object System.Windows.Forms.Padding(4, 8, 10, 10)

    $lblOvDetailsTitle = New-Object System.Windows.Forms.Label
    $lblOvDetailsTitle.Text = "Verlinkungsziele der GPO (OUs / Domaene):"
    $lblOvDetailsTitle.Dock = [System.Windows.Forms.DockStyle]::Top
    $lblOvDetailsTitle.Height = 46
    $lblOvDetailsTitle.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)

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
    $gridOvDetails.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    $gridOvDetails.ColumnHeadersHeight = 34
    $gridOvDetails.RowTemplate.Height = 28
    $gridOvDetails.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(249, 251, 254)
    $gridOvDetails.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill

    $panelOvRight.Controls.Add($gridOvDetails)
    $panelOvRight.Controls.Add($lblOvDetailsTitle)
    $splitOverview.Panel2.Controls.Add($panelOvRight)

    $tabOverview.Controls.Add($splitOverview)
    $tabOverview.Controls.Add($panelOverviewTop)

    # =========================================================================
    # REGISTER 2: GPO Richtlinien-Einstellungen & Inspektor (Mit Pfadanzeige)
    # =========================================================================
    $tabSettings = New-Object System.Windows.Forms.TabPage
    $tabSettings.Text = "2. GPO Richtlinien-Einstellungen & Inspektor"
    $tabSettings.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)

    $panelSettingsTop = New-Object System.Windows.Forms.Panel
    $panelSettingsTop.Dock = [System.Windows.Forms.DockStyle]::Top
    $panelSettingsTop.Height = 110
    $panelSettingsTop.BackColor = [System.Drawing.Color]::FromArgb(242, 245, 250)
    $panelSettingsTop.Padding = New-Object System.Windows.Forms.Padding(10)

    # Zeile 1: Pfad & Export
    $lblSettingsTargetPath = New-Object System.Windows.Forms.Label
    $lblSettingsTargetPath.Text = "Export Ziel-Pfad:"
    $lblSettingsTargetPath.Location = New-Object System.Drawing.Point(12, 17)
    $lblSettingsTargetPath.AutoSize = $true
    $lblSettingsTargetPath.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $txtSettingsExportDir = New-Object System.Windows.Forms.TextBox
    $txtSettingsExportDir.Location = New-Object System.Drawing.Point(135, 14)
    $txtSettingsExportDir.Size = New-Object System.Drawing.Size(440, 25)
    $txtSettingsExportDir.Text = "C:\Install\Backup\GPO"

    $btnBrowseSettingsDir = New-Object System.Windows.Forms.Button
    $btnBrowseSettingsDir.Text = "Durchsuchen..."
    $btnBrowseSettingsDir.Location = New-Object System.Drawing.Point(585, 11)
    $btnBrowseSettingsDir.Size = New-Object System.Drawing.Size(115, 30)

    $btnExportCsv = New-Object System.Windows.Forms.Button
    $btnExportCsv.Text = "CSV Export"
    $btnExportCsv.Location = New-Object System.Drawing.Point(710, 11)
    $btnExportCsv.Size = New-Object System.Drawing.Size(110, 30)
    $btnExportCsv.BackColor = [System.Drawing.Color]::FromArgb(230, 245, 230)

    $lblSettingsStatus = New-Object System.Windows.Forms.Label
    $lblSettingsStatus.Text = "Bereit."
    $lblSettingsStatus.Location = New-Object System.Drawing.Point(835, 17)
    $lblSettingsStatus.AutoSize = $true
    $lblSettingsStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)

    # Zeile 2: Filter, GPO-Auswahl, Laden & Suche
    $lblSettingsViewFilter = New-Object System.Windows.Forms.Label
    $lblSettingsViewFilter.Text = "Ansicht:"
    $lblSettingsViewFilter.Location = New-Object System.Drawing.Point(12, 58)
    $lblSettingsViewFilter.AutoSize = $true
    $lblSettingsViewFilter.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $comboSettingsViewMode = New-Object System.Windows.Forms.ComboBox
    $comboSettingsViewMode.Location = New-Object System.Drawing.Point(75, 55)
    $comboSettingsViewMode.Size = New-Object System.Drawing.Size(210, 25)
    $comboSettingsViewMode.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$comboSettingsViewMode.Items.Add("Alle GPOs")
    [void]$comboSettingsViewMode.Items.Add("Nur verlinkte GPOs")
    [void]$comboSettingsViewMode.Items.Add("Nicht verlinkte GPOs (Unlinked)")
    $comboSettingsViewMode.SelectedIndex = 0

    $lblGpo = New-Object System.Windows.Forms.Label
    $lblGpo.Text = "GPO:"
    $lblGpo.Location = New-Object System.Drawing.Point(295, 58)
    $lblGpo.AutoSize = $true
    $lblGpo.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $comboGpo = New-Object System.Windows.Forms.ComboBox
    $comboGpo.Location = New-Object System.Drawing.Point(340, 55)
    $comboGpo.Size = New-Object System.Drawing.Size(360, 25)
    $comboGpo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

    $btnLoadSettings = New-Object System.Windows.Forms.Button
    $btnLoadSettings.Text = "Laden"
    $btnLoadSettings.Location = New-Object System.Drawing.Point(710, 52)
    $btnLoadSettings.Size = New-Object System.Drawing.Size(110, 32)
    $btnLoadSettings.BackColor = [System.Drawing.Color]::FromArgb(225, 238, 255)
    $btnLoadSettings.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $lblFilter = New-Object System.Windows.Forms.Label
    $lblFilter.Text = "Suche:"
    $lblFilter.Location = New-Object System.Drawing.Point(830, 58)
    $lblFilter.AutoSize = $true
    $lblFilter.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $txtFilter = New-Object System.Windows.Forms.TextBox
    $txtFilter.Location = New-Object System.Drawing.Point(885, 55)
    $txtFilter.Size = New-Object System.Drawing.Size(200, 25)

    $panelSettingsTop.Controls.Add($lblSettingsTargetPath)
    $panelSettingsTop.Controls.Add($txtSettingsExportDir)
    $panelSettingsTop.Controls.Add($btnBrowseSettingsDir)
    $panelSettingsTop.Controls.Add($btnExportCsv)
    $panelSettingsTop.Controls.Add($lblSettingsStatus)
    $panelSettingsTop.Controls.Add($lblSettingsViewFilter)
    $panelSettingsTop.Controls.Add($comboSettingsViewMode)
    $panelSettingsTop.Controls.Add($lblGpo)
    $panelSettingsTop.Controls.Add($comboGpo)
    $panelSettingsTop.Controls.Add($btnLoadSettings)
    $panelSettingsTop.Controls.Add($lblFilter)
    $panelSettingsTop.Controls.Add($txtFilter)

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
    $lblTableTitle.Height = 28
    $lblTableTitle.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)

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
    $gridSettings.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    $gridSettings.ColumnHeadersHeight = 34
    $gridSettings.RowTemplate.Height = 28
    $gridSettings.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(249, 251, 254)

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

    $panelSettingsRight = New-Object System.Windows.Forms.Panel
    $panelSettingsRight.Dock = [System.Windows.Forms.DockStyle]::Fill
    $panelSettingsRight.Padding = New-Object System.Windows.Forms.Padding(4, 8, 10, 10)

    $lblDescHeader = New-Object System.Windows.Forms.Label
    $lblDescHeader.Text = "Erlaeuterung & Richtlinien-Details:"
    $lblDescHeader.Dock = [System.Windows.Forms.DockStyle]::Top
    $lblDescHeader.Height = 28
    $lblDescHeader.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)

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
    # REGISTER 3: GPO Backup & Verknuepfungs-Audit
    # =========================================================================
    $tabBackup = New-Object System.Windows.Forms.TabPage
    $tabBackup.Text = "3. GPO Backup & Verknuepfungs-Status"
    $tabBackup.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)

    $panelBackupTop = New-Object System.Windows.Forms.Panel
    $panelBackupTop.Dock = [System.Windows.Forms.DockStyle]::Top
    $panelBackupTop.Height = 108
    $panelBackupTop.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
    $panelBackupTop.Padding = New-Object System.Windows.Forms.Padding(10)

    $lblTargetPath = New-Object System.Windows.Forms.Label
    $lblTargetPath.Text = "Backup Ziel-Pfad:"
    $lblTargetPath.Location = New-Object System.Drawing.Point(12, 17)
    $lblTargetPath.AutoSize = $true
    $lblTargetPath.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $txtBackupTargetDir = New-Object System.Windows.Forms.TextBox
    $txtBackupTargetDir.Location = New-Object System.Drawing.Point(140, 14)
    $txtBackupTargetDir.Size = New-Object System.Drawing.Size(360, 25)
    $txtBackupTargetDir.Text = "C:\Install\Backup\GPO"

    $btnBrowseFolder = New-Object System.Windows.Forms.Button
    $btnBrowseFolder.Text = "Durchsuchen..."
    $btnBrowseFolder.Location = New-Object System.Drawing.Point(510, 11)
    $btnBrowseFolder.Size = New-Object System.Drawing.Size(105, 30)

    $btnExportBackupCsv = New-Object System.Windows.Forms.Button
    $btnExportBackupCsv.Text = "CSV Export"
    $btnExportBackupCsv.Location = New-Object System.Drawing.Point(625, 11)
    $btnExportBackupCsv.Size = New-Object System.Drawing.Size(95, 30)
    $btnExportBackupCsv.BackColor = [System.Drawing.Color]::FromArgb(230, 245, 230)

    $btnExportSelectedHtml = New-Object System.Windows.Forms.Button
    $btnExportSelectedHtml.Text = "HTML (Ausgewaehlt)"
    $btnExportSelectedHtml.Location = New-Object System.Drawing.Point(730, 11)
    $btnExportSelectedHtml.Size = New-Object System.Drawing.Size(155, 30)
    $btnExportSelectedHtml.BackColor = [System.Drawing.Color]::FromArgb(255, 245, 230)

    $btnExportAllHtml = New-Object System.Windows.Forms.Button
    $btnExportAllHtml.Text = "HTML (Gefilterte)"
    $btnExportAllHtml.Location = New-Object System.Drawing.Point(895, 11)
    $btnExportAllHtml.Size = New-Object System.Drawing.Size(150, 30)
    $btnExportAllHtml.BackColor = [System.Drawing.Color]::FromArgb(255, 238, 220)

    $lblBackupNote = New-Object System.Windows.Forms.Label
    $lblBackupNote.Text = "Format: [Zielpfad]\[Name der GPO]\[JJJJMMTT_HHMM]"
    $lblBackupNote.Location = New-Object System.Drawing.Point(1060, 18)
    $lblBackupNote.AutoSize = $true
    $lblBackupNote.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Italic)

    $lblBackupViewFilter = New-Object System.Windows.Forms.Label
    $lblBackupViewFilter.Text = "Ansicht:"
    $lblBackupViewFilter.Location = New-Object System.Drawing.Point(12, 58)
    $lblBackupViewFilter.AutoSize = $true
    $lblBackupViewFilter.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $comboBackupFilter = New-Object System.Windows.Forms.ComboBox
    $comboBackupFilter.Location = New-Object System.Drawing.Point(72, 55)
    $comboBackupFilter.Size = New-Object System.Drawing.Size(200, 25)
    $comboBackupFilter.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$comboBackupFilter.Items.Add("Alle GPOs")
    [void]$comboBackupFilter.Items.Add("Nur verlinkte GPOs")
    [void]$comboBackupFilter.Items.Add("Nicht verlinkte GPOs (Unlinked)")
    $comboBackupFilter.SelectedIndex = 0

    $lblBackupSearch = New-Object System.Windows.Forms.Label
    $lblBackupSearch.Text = "Suche:"
    $lblBackupSearch.Location = New-Object System.Drawing.Point(282, 58)
    $lblBackupSearch.AutoSize = $true
    $lblBackupSearch.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $txtBackupSearch = New-Object System.Windows.Forms.TextBox
    $txtBackupSearch.Location = New-Object System.Drawing.Point(335, 55)
    $txtBackupSearch.Size = New-Object System.Drawing.Size(140, 25)

    $btnLoadGpos = New-Object System.Windows.Forms.Button
    $btnLoadGpos.Text = "GPO-Liste laden"
    $btnLoadGpos.Location = New-Object System.Drawing.Point(485, 52)
    $btnLoadGpos.Size = New-Object System.Drawing.Size(120, 32)

    $btnBackupSelected = New-Object System.Windows.Forms.Button
    $btnBackupSelected.Text = "Ausgewaehlte GPO sichern"
    $btnBackupSelected.Location = New-Object System.Drawing.Point(615, 52)
    $btnBackupSelected.Size = New-Object System.Drawing.Size(190, 32)
    $btnBackupSelected.BackColor = [System.Drawing.Color]::FromArgb(230, 245, 230)

    $btnBackupAll = New-Object System.Windows.Forms.Button
    $btnBackupAll.Text = "ALLE GPOs sichern"
    $btnBackupAll.Location = New-Object System.Drawing.Point(815, 52)
    $btnBackupAll.Size = New-Object System.Drawing.Size(210, 32)
    $btnBackupAll.BackColor = [System.Drawing.Color]::FromArgb(255, 245, 230)

    $panelBackupTop.Controls.Add($lblTargetPath)
    $panelBackupTop.Controls.Add($txtBackupTargetDir)
    $panelBackupTop.Controls.Add($btnBrowseFolder)
    $panelBackupTop.Controls.Add($btnExportBackupCsv)
    $panelBackupTop.Controls.Add($btnExportSelectedHtml)
    $panelBackupTop.Controls.Add($btnExportAllHtml)
    $panelBackupTop.Controls.Add($lblBackupNote)
    $panelBackupTop.Controls.Add($lblBackupViewFilter)
    $panelBackupTop.Controls.Add($comboBackupFilter)
    $panelBackupTop.Controls.Add($lblBackupSearch)
    $panelBackupTop.Controls.Add($txtBackupSearch)
    $panelBackupTop.Controls.Add($btnLoadGpos)
    $panelBackupTop.Controls.Add($btnBackupSelected)
    $panelBackupTop.Controls.Add($btnBackupAll)

    $splitBackupMain = New-Object System.Windows.Forms.SplitContainer
    $splitBackupMain.Dock = [System.Windows.Forms.DockStyle]::Fill
    $splitBackupMain.SplitterDistance = 750
    $splitBackupMain.SplitterWidth = 6

    $panelGpoLeft = New-Object System.Windows.Forms.Panel
    $panelGpoLeft.Dock = [System.Windows.Forms.DockStyle]::Fill
    $panelGpoLeft.Padding = New-Object System.Windows.Forms.Padding(10, 8, 4, 10)

    $lblGpoGrid = New-Object System.Windows.Forms.Label
    $lblGpoGrid.Text = "1. Gruppenrichtlinien der Domaene:"
    $lblGpoGrid.Dock = [System.Windows.Forms.DockStyle]::Top
    $lblGpoGrid.Height = 28
    $lblGpoGrid.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)

    $gridGpos = New-Object System.Windows.Forms.DataGridView
    $gridGpos.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridGpos.ReadOnly = $true
    $gridGpos.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
    $gridGpos.AllowUserToAddRows = $false
    $gridGpos.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $gridGpos.MultiSelect = $false
    $gridGpos.RowHeadersVisible = $false
    $gridGpos.BackgroundColor = [System.Drawing.Color]::White
    $gridGpos.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $gridGpos.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(230, 236, 245)
    $gridGpos.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    $gridGpos.ColumnHeadersHeight = 34
    $gridGpos.RowTemplate.Height = 28
    $gridGpos.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(249, 251, 254)

    $panelGpoLeft.Controls.Add($gridGpos)
    $panelGpoLeft.Controls.Add($lblGpoGrid)
    $splitBackupMain.Panel1.Controls.Add($panelGpoLeft)

    $splitBackupRight = New-Object System.Windows.Forms.SplitContainer
    $splitBackupRight.Dock = [System.Windows.Forms.DockStyle]::Fill
    $splitBackupRight.Orientation = [System.Windows.Forms.Orientation]::Horizontal
    $splitBackupRight.SplitterDistance = 280
    $splitBackupRight.SplitterWidth = 6

    $panelLinks = New-Object System.Windows.Forms.Panel
    $panelLinks.Dock = [System.Windows.Forms.DockStyle]::Fill
    $panelLinks.Padding = New-Object System.Windows.Forms.Padding(4, 8, 10, 4)

    $lblLinks = New-Object System.Windows.Forms.Label
    $lblLinks.Text = "2. Verknuepfungs-Ziele (OUs / Sites):"
    $lblLinks.Dock = [System.Windows.Forms.DockStyle]::Top
    $lblLinks.Height = 28
    $lblLinks.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)

    $gridLinks = New-Object System.Windows.Forms.DataGridView
    $gridLinks.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridLinks.ReadOnly = $true
    $gridLinks.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $gridLinks.AllowUserToAddRows = $false
    $gridLinks.RowHeadersVisible = $false
    $gridLinks.BackgroundColor = [System.Drawing.Color]::White
    $gridLinks.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $gridLinks.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(230, 236, 245)
    $gridLinks.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    $gridLinks.ColumnHeadersHeight = 34
    $gridLinks.RowTemplate.Height = 28
    $gridLinks.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(249, 251, 254)

    $panelLinks.Controls.Add($gridLinks)
    $panelLinks.Controls.Add($lblLinks)
    $splitBackupRight.Panel1.Controls.Add($panelLinks)

    $panelLog = New-Object System.Windows.Forms.Panel
    $panelLog.Dock = [System.Windows.Forms.DockStyle]::Fill
    $panelLog.Padding = New-Object System.Windows.Forms.Padding(4, 4, 10, 10)

    $lblLog = New-Object System.Windows.Forms.Label
    $lblLog.Text = "3. Backup- & Aktivitaets-Protokoll:"
    $lblLog.Dock = [System.Windows.Forms.DockStyle]::Top
    $lblLog.Height = 28
    $lblLog.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)

    $txtLog = New-Object System.Windows.Forms.TextBox
    $txtLog.Dock = [System.Windows.Forms.DockStyle]::Fill
    $txtLog.Multiline = $true
    $txtLog.ReadOnly = $true
    $txtLog.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $txtLog.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 252)
    $txtLog.Font = New-Object System.Drawing.Font("Consolas", 9)

    $panelLog.Controls.Add($txtLog)
    $panelLog.Controls.Add($lblLog)
    $splitBackupRight.Panel2.Controls.Add($panelLog)

    $splitBackupMain.Panel2.Controls.Add($splitBackupRight)

    $tabBackup.Controls.Add($splitBackupMain)
    $tabBackup.Controls.Add($panelBackupTop)

    # =========================================================================
    # REGISTER 4: GPO-Vergleich (Diff) & DDP-Baseline-Check
    # =========================================================================
    $tabCompare = New-Object System.Windows.Forms.TabPage
    $tabCompare.Text = "4. GPO-Vergleich (Diff)"
    $tabCompare.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)

    $panelCompareTop = New-Object System.Windows.Forms.Panel
    $panelCompareTop.Dock = [System.Windows.Forms.DockStyle]::Top
    $panelCompareTop.Height = 90
    $panelCompareTop.BackColor = [System.Drawing.Color]::FromArgb(242, 245, 250)
    $panelCompareTop.Padding = New-Object System.Windows.Forms.Padding(10)

    $lblGpo1 = New-Object System.Windows.Forms.Label
    $lblGpo1.Text = "GPO 1 (Basis):"
    $lblGpo1.Location = New-Object System.Drawing.Point(12, 16)
    $lblGpo1.AutoSize = $true
    $lblGpo1.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $comboGpo1 = New-Object System.Windows.Forms.ComboBox
    $comboGpo1.Location = New-Object System.Drawing.Point(115, 13)
    $comboGpo1.Size = New-Object System.Drawing.Size(280, 25)
    $comboGpo1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

    $lblGpo2 = New-Object System.Windows.Forms.Label
    $lblGpo2.Text = "GPO 2 (Vergleich):"
    $lblGpo2.Location = New-Object System.Drawing.Point(410, 16)
    $lblGpo2.AutoSize = $true
    $lblGpo2.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $comboGpo2 = New-Object System.Windows.Forms.ComboBox
    $comboGpo2.Location = New-Object System.Drawing.Point(540, 13)
    $comboGpo2.Size = New-Object System.Drawing.Size(290, 25)
    $comboGpo2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

    $btnCompare = New-Object System.Windows.Forms.Button
    $btnCompare.Text = "Vergleichen"
    $btnCompare.Location = New-Object System.Drawing.Point(840, 10)
    $btnCompare.Size = New-Object System.Drawing.Size(110, 30)
    $btnCompare.BackColor = [System.Drawing.Color]::FromArgb(225, 238, 255)

    $btnCompareDdpBaseline = New-Object System.Windows.Forms.Button
    $btnCompareDdpBaseline.Text = "DDP vs. MS-Standard"
    $btnCompareDdpBaseline.Location = New-Object System.Drawing.Point(960, 10)
    $btnCompareDdpBaseline.Size = New-Object System.Drawing.Size(170, 30)
    $btnCompareDdpBaseline.BackColor = [System.Drawing.Color]::FromArgb(255, 243, 224)
    $btnCompareDdpBaseline.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $chkOnlyDiffs = New-Object System.Windows.Forms.CheckBox
    $chkOnlyDiffs.Text = "Nur Unterschiede anzeigen"
    $chkOnlyDiffs.Location = New-Object System.Drawing.Point(12, 52)
    $chkOnlyDiffs.AutoSize = $true
    $chkOnlyDiffs.Checked = $true
    $chkOnlyDiffs.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $btnExportCompareCsv = New-Object System.Windows.Forms.Button
    $btnExportCompareCsv.Text = "Diff CSV Export"
    $btnExportCompareCsv.Location = New-Object System.Drawing.Point(235, 48)
    $btnExportCompareCsv.Size = New-Object System.Drawing.Size(130, 30)
    $btnExportCompareCsv.BackColor = [System.Drawing.Color]::FromArgb(230, 245, 230)

    $lblCompareStatus = New-Object System.Windows.Forms.Label
    $lblCompareStatus.Text = "Waehlen Sie zwei GPOs fuer den Vergleich."
    $lblCompareStatus.Location = New-Object System.Drawing.Point(380, 55)
    $lblCompareStatus.AutoSize = $true
    $lblCompareStatus.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Italic)

    $panelCompareTop.Controls.Add($lblGpo1)
    $panelCompareTop.Controls.Add($comboGpo1)
    $panelCompareTop.Controls.Add($lblGpo2)
    $panelCompareTop.Controls.Add($comboGpo2)
    $panelCompareTop.Controls.Add($btnCompare)
    $panelCompareTop.Controls.Add($btnCompareDdpBaseline)
    $panelCompareTop.Controls.Add($chkOnlyDiffs)
    $panelCompareTop.Controls.Add($btnExportCompareCsv)
    $panelCompareTop.Controls.Add($lblCompareStatus)

    $panelCompareMain = New-Object System.Windows.Forms.Panel
    $panelCompareMain.Dock = [System.Windows.Forms.DockStyle]::Fill
    $panelCompareMain.Padding = New-Object System.Windows.Forms.Padding(10, 8, 10, 10)

    $lblCompareLegend = New-Object System.Windows.Forms.Label
    $lblCompareLegend.Text = "Legende:  [Gruen] Identische Einstellung  |  [Gelb/Orange] Abweichender Wert / Status  |  [Rot] Nur in GPO 1  |  [Blau] Nur in GPO 2"
    $lblCompareLegend.Dock = [System.Windows.Forms.DockStyle]::Top
    $lblCompareLegend.Height = 28
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
    $gridCompare.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    $gridCompare.ColumnHeadersDefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::False
    $gridCompare.ColumnHeadersHeight = 34
    $gridCompare.RowTemplate.Height = 28
    $gridCompare.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(249, 251, 254)

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
    $gridCompare.Columns["colCmpState1"].Width = 130
    $gridCompare.Columns["colCmpVal1"].Width = 210
    $gridCompare.Columns["colCmpState2"].Width = 130
    $gridCompare.Columns["colCmpVal2"].Width = 210
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
    $form.Controls.Add($panelBottomStatus)
    $panelBottomStatus.SendToBack()
    $tabControl.BringToFront()

    # Lokale Datencontainer
    $rawOverviewList = [System.Collections.Generic.List[PSCustomObject]]::new()
    $rawBackupList   = [System.Collections.Generic.List[PSCustomObject]]::new()
    $gpoLinksCache   = @{}
    $allGposCache    = [System.Collections.Generic.List[Microsoft.GroupPolicy.Gpo]]::new()
    $rawSettingsList = [System.Collections.Generic.List[PSCustomObject]]::new()
    $rawCompareList  = [System.Collections.Generic.List[PSCustomObject]]::new()

    # =========================================================================
    # REFERENZ: MICROSOFT DEFAULT DOMAIN POLICY WERKSZUSTAND
    # =========================================================================
    function Get-DefaultDomainPolicyBaseline {
        $dateStr = (Get-Date -Format "yyyyMMdd")
        $timeStr = (Get-Date -Format "HHmm")
        $baseline = [System.Collections.Generic.List[PSCustomObject]]::new()

        $addBase = {
            param($cat, $name, $state, $val, $explain)
            $baseline.Add([PSCustomObject]@{
                Scope     = "Computer"
                Category  = $cat
                Name      = $name
                Value     = $val
                State     = $state
                Supported = "Windows 2000 und hoeher"
                Explain   = $explain
                GpoName   = $script:DdpBaselineName
                Datum     = $dateStr
                Uhrzeit   = $timeStr
            })
        }

        & $addBase "Sicherheitseinstellungen / Kontorichtlinien / Kennwortrichtlinie" "Kennwortchronik erzwingen (Password History)" "Aktiviert" "24 gespeicherte Kennwoerter" "Bestimmt die Anzahl neuer Kennwoerter, die verwendet werden muessen, bevor ein altes Kennwort wiederverwendet werden kann. Microsoft Standard: 24."
        & $addBase "Sicherheitseinstellungen / Kontorichtlinien / Kennwortrichtlinie" "Maximales Kennwortalter (Maximum Password Age)" "Aktiviert" "42 Tage" "Bestimmt den Zeitraum in Tagen, fuer den ein Kennwort verwendet werden kann, bevor das System den Benutzer zum Aendern auffordert. Microsoft Standard: 42 Tage."
        & $addBase "Sicherheitseinstellungen / Kontorichtlinien / Kennwortrichtlinie" "Minimales Kennwortalter (Minimum Password Age)" "Aktiviert" "1 Tag" "Bestimmt den Zeitraum in Tagen, fuer den ein Kennwort verwendet werden muss, bevor der Benutzer es aendern kann. Microsoft Standard: 1 Tag."
        & $addBase "Sicherheitseinstellungen / Kontorichtlinien / Kennwortrichtlinie" "Mindestkennwortlaenge (Minimum Password Length)" "Aktiviert" "7 Zeichen" "Bestimmt die Mindestanzahl von Zeichen, die das Kennwort eines Benutzerkontos enthalten muss. Microsoft Standard: 7 Zeichen."
        & $addBase "Sicherheitseinstellungen / Kontorichtlinien / Kennwortrichtlinie" "Kennwort muss Komplexitaetsanforderungen entsprechen" "Aktiviert" "Aktiviert" "Kennwoerter muessen Zeichen aus mindestens 3 der 4 Kategorien enthalten. Microsoft Standard: Aktiviert."
        & $addBase "Sicherheitseinstellungen / Kontorichtlinien / Kennwortrichtlinie" "Kennwoerter mit umkehrbarer Verschluesselung speichern" "Deaktiviert" "Deaktiviert" "Bestimmt, ob das Betriebssystem Kennwoerter unter Verwendung umkehrbarer Verschluesselung speichert. Microsoft Standard: Deaktiviert."
        & $addBase "Sicherheitseinstellungen / Kontorichtlinien / Kontosperrungsrichtlinie" "Kontosperrungsschwelle (Account Lockout Threshold)" "Deaktiviert" "0 ungueltige Anmeldeversuche (Keine Kontosperrung)" "Bestimmt die Anzahl fehlerhafter Anmeldeversuche, nach denen ein Benutzerkonto gesperrt wird. Microsoft Standard: 0."
        & $addBase "Sicherheitseinstellungen / Kontorichtlinien / Kontosperrungsrichtlinie" "Kontosperrdauer (Account Lockout Duration)" "Nicht definiert" "Nicht definiert" "Bestimmt die Dauer einer Kontosperre. Microsoft Standard: Nicht definiert."
        & $addBase "Sicherheitseinstellungen / Kontorichtlinien / Kontosperrungsrichtlinie" "Zuruecksetzungsdauer des Kontosperrzaehlers" "Nicht definiert" "Nicht definiert" "Bestimmt die Zeit bis zum Zuruecksetzen des Zaehlers. Microsoft Standard: Nicht definiert."
        & $addBase "Sicherheitseinstellungen / Kontorichtlinien / Kerberos-Richtlinie" "Maximale Toleranz fuer Synchronisierung der Computeruhr" "Aktiviert" "5 Minuten" "Maximal zulaessige Zeitdifferenz zwischen Client und DC. Microsoft Standard: 5 Minuten."
        & $addBase "Sicherheitseinstellungen / Kontorichtlinien / Kerberos-Richtlinie" "Maximale Lebensdauer fuer Benutzerticket (TGT)" "Aktiviert" "10 Stunden" "Maximale Gueltigkeitsdauer eines TGT. Microsoft Standard: 10 Stunden."
        & $addBase "Sicherheitseinstellungen / Kontorichtlinien / Kerberos-Richtlinie" "Maximale Lebensdauer fuer Serviceticket" "Aktiviert" "600 Minuten" "Maximale Gueltigkeitsdauer eines Servicetickets. Microsoft Standard: 600 Minuten."
        & $addBase "Sicherheitseinstellungen / Kontorichtlinien / Kerberos-Richtlinie" "Maximale Lebensdauer fuer Benutzer-Ticket-Erneuerung" "Aktiviert" "7 Tage" "Zeitraum zur TGT-Erneuerung. Microsoft Standard: 7 Tage."
        & $addBase "Sicherheitseinstellungen / Kontorichtlinien / Kerberos-Richtlinie" "Benutzeranmeldeeinschraenkungen erzwingen" "Aktiviert" "Aktiviert" "Ueberpruefung der Anmelderechte bei Ticketanforderung. Microsoft Standard: Aktiviert."

        return $baseline
    }

    function Get-NormalizedPolicyKey ($item) {
        $n = "$($item.Name)".ToLower()
        if ($n -match "passwordhistory" -or $n -match "kennwortchronik" -or $n -match "password history") { return "$($item.Scope)|PasswordHistory" }
        if ($n -match "maximumpasswordage" -or $n -match "maximales kennwortalter" -or $n -match "maximum password age") { return "$($item.Scope)|MaxPasswordAge" }
        if ($n -match "minimumpasswordage" -or $n -match "minimales kennwortalter" -or $n -match "minimum password age") { return "$($item.Scope)|MinPasswordAge" }
        if ($n -match "minpasswordlength" -or $n -match "mindestkennwort" -or $n -match "minimum password length") { return "$($item.Scope)|MinPasswordLength" }
        if ($n -match "passwordcomplexity" -or $n -match "komplexit" -or $n -match "complexity requirements") { return "$($item.Scope)|PasswordComplexity" }
        if ($n -match "cleartextpassword" -or $n -match "umkehrbar" -or $n -match "reversible encryption") { return "$($item.Scope)|ReversibleEncryption" }
        if ($n -match "lockoutbadcount" -or $n -match "kontosperrungsschwelle" -or $n -match "lockout threshold") { return "$($item.Scope)|LockoutThreshold" }
        if ($n -match "lockoutduration" -or $n -match "kontosperrdauer" -or $n -match "lockout duration") { return "$($item.Scope)|LockoutDuration" }
        if ($n -match "resetlockoutcount" -or $n -match "zuruecksetzungsdauer" -or $n -match "reset account lockout") { return "$($item.Scope)|ResetLockoutCount" }
        if ($n -match "maxclockskew" -or $n -match "synchronisierung der computeruhr" -or $n -match "clock synchronization") { return "$($item.Scope)|MaxClockSkew" }
        if ($n -match "maxticketage" -or $n -match "lebensdauer.*benutzerticket" -or $n -match "lifetime for user ticket") { return "$($item.Scope)|MaxTicketAge" }
        if ($n -match "maxserviceage" -or $n -match "lebensdauer.*serviceticket" -or $n -match "lifetime for service ticket") { return "$($item.Scope)|MaxServiceAge" }
        if ($n -match "maxrenewage" -or $n -match "erneuerung von benutzertickets" -or $n -match "user ticket renewal") { return "$($item.Scope)|MaxRenewAge" }
        if ($n -match "ticketvalidateclient" -or $n -match "benutzeranmeldeeinschraenkungen" -or $n -match "user logon restrictions") { return "$($item.Scope)|TicketValidateClient" }
        return "$($item.Scope)|$($item.Category)|$($item.Name)"
    }

    # =========================================================================
    # PRAEZISER XML-PARSER (Trennt Parameter-Werte strikt von Explain-Texten)
    # =========================================================================
    function Get-ParsedGpoSettings {
        param([string]$GpoId, [string]$GpoDisplayName)

        $dateStr = Get-Date -Format "yyyyMMdd"
        $timeStr = Get-Date -Format "HHmm"
        $list = [System.Collections.Generic.List[PSCustomObject]]::new()

        [xml]$xml = Get-GPOReport -Guid $GpoId -ReportType Xml -ErrorAction Stop

        function Parse-GpoSection ($sectionNode, $scope) {
            if ($null -eq $sectionNode -or -not $sectionNode.ExtensionData) { return }

            foreach ($ext in $sectionNode.ExtensionData.Extension) {
                $extType = if ($ext.type) { $ext.type } else { $ext.LocalName }
                $extCategory = if ($ext.Name) { $ext.Name } else { "Erweiterung" }

                # 1. Administrative Vorlagen (<Policy>)
                $policies = $ext.SelectNodes(".//*[local-name()='Policy']")
                if ($policies -and $policies.Count -gt 0) {
                    foreach ($p in $policies) {
                        $pName = if ($p.SelectSingleNode("./*[local-name()='Name']")) { $p.SelectSingleNode("./*[local-name()='Name']").InnerText.Trim() } elseif ($p.Name) { $p.Name.Trim() } else { "Unbenannte Richtlinie" }
                        $rawState = if ($p.SelectSingleNode("./*[local-name()='State']")) { $p.SelectSingleNode("./*[local-name()='State']").InnerText.Trim() } elseif ($p.State) { $p.State.Trim() } else { "Enabled" }
                        $pState = switch ($rawState) { "Enabled" { "Aktiviert" } "Disabled" { "Deaktiviert" } default { $rawState } }
                        $pCat = if ($p.SelectSingleNode("./*[local-name()='Category']")) { $p.SelectSingleNode("./*[local-name()='Category']").InnerText.Trim() } else { $extCategory }
                        $pSupported = if ($p.SelectSingleNode("./*[local-name()='Supported' or local-name()='SupportedOn']")) { $p.SelectSingleNode("./*[local-name()='Supported' or local-name()='SupportedOn']").InnerText.Trim() } else { "Keine Angabe" }
                        $pExplain = if ($p.SelectSingleNode("./*[local-name()='Explain' or local-name()='ExplainText']")) { $p.SelectSingleNode("./*[local-name()='Explain' or local-name()='ExplainText']").InnerText.Trim() } else { "Keine Erklaerung hinterlegt." }

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
                                    if (-not [string]::IsNullOrWhiteSpace($vn.InnerText)) { $extractedList += $vn.InnerText.Trim() }
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
                                    $paramValues += "$($optLabel): $($joinedVals)"
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

                # 2. Systemdienste
                $services = $ext.SelectNodes(".//*[local-name()='SystemServices'] | .//*[local-name()='Service']")
                if ($services -and $services.Count -gt 0) {
                    foreach ($svc in $services) {
                        $svcName = if ($svc.SelectSingleNode("./*[local-name()='Display']/*[local-name()='Name']")) { $svc.SelectSingleNode("./*[local-name()='Display']/*[local-name()='Name']").InnerText.Trim() } elseif ($svc.SelectSingleNode("./*[local-name()='Name']")) { $svc.SelectSingleNode("./*[local-name()='Name']").InnerText.Trim() } else { "Systemdienst" }
                        $mode = if ($svc.SelectSingleNode("./*[local-name()='StartupMode']")) { $svc.SelectSingleNode("./*[local-name()='StartupMode']").InnerText.Trim() } elseif ($svc.SelectSingleNode("./*[local-name()='Mode']")) { $svc.SelectSingleNode("./*[local-name()='Mode']").InnerText.Trim() } else { "Konfiguriert" }
                        $modeDE = switch ($mode) { "Disabled" { "Deaktiviert" } "Automatic" { "Automatisch" } "Manual" { "Manuell" } default { $mode } }

                        $list.Add([PSCustomObject]@{
                            Scope     = $scope
                            Category  = "Sicherheitseinstellungen / Systemdienste"
                            Name      = $svcName
                            Value     = "Starttyp: $modeDE"
                            State     = $modeDE
                            Supported = "Windows Systemdienste"
                            Explain   = "Startmodus fuer den Dienst '$svcName': $modeDE"
                            GpoName   = $GpoDisplayName
                            Datum     = $dateStr
                            Uhrzeit   = $timeStr
                        })
                    }
                }

                # 3. GPP (Praeferenzen)
                $gppNodes = $ext.SelectNodes(".//*[local-name()='Properties']")
                if ($gppNodes -and $gppNodes.Count -gt 0) {
                    foreach ($prop in $gppNodes) {
                        $parent = $prop.ParentNode
                        $itemType = $parent.LocalName
                        $gppCategory = "Praeferenzen (GPP) / $itemType"

                        $itemName = if ($parent.Attributes["name"]) { $parent.Attributes["name"].Value } elseif ($prop.Attributes["name"]) { $prop.Attributes["name"].Value } elseif ($prop.Attributes["path"]) { $prop.Attributes["path"].Value } elseif ($prop.Attributes["letter"]) { "Laufwerk $($prop.Attributes['letter'].Value):" } else { "GPP $itemType" }
                        $actionCode = if ($prop.Attributes["action"]) { $prop.Attributes["action"].Value } else { "U" }
                        $actionText = switch ($actionCode) { "C" { "Erstellen" } "U" { "Aktualisieren" } "R" { "Ersetzen" } "D" { "Loeschen" } default { $actionCode } }

                        $valParts = @("Aktion: $actionText")
                        if ($prop.Attributes["path"])       { $valParts += "Pfad: $($prop.Attributes['path'].Value)" }
                        if ($prop.Attributes["targetPath"]) { $valParts += "Ziel: $($prop.Attributes['targetPath'].Value)" }
                        if ($prop.Attributes["fromPath"])   { $valParts += "Quelle: $($prop.Attributes['fromPath'].Value)" }
                        if ($prop.Attributes["value"])      { $valParts += "Wert: $($prop.Attributes['value'].Value)" }
                        $cleanGppVal = $valParts -join " | "

                        $descLines = @("GPP OBJEKT: $itemName", "TYP: $itemType", "AKTION: $actionText", "--------------------------------------------------")
                        foreach ($att in $prop.Attributes) { $descLines += " - $($att.Name): $($att.Value)" }

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
                        $secName = if ($sec.SelectSingleNode("./*[local-name()='Display']/*[local-name()='Name']")) { $sec.SelectSingleNode("./*[local-name()='Display']/*[local-name()='Name']").InnerText.Trim() } elseif ($sec.SelectSingleNode("./*[local-name()='Name']")) { $sec.SelectSingleNode("./*[local-name()='Name']").InnerText.Trim() } else { $sec.LocalName }
                        $secVal = if ($sec.SelectSingleNode("./*[local-name()='Display']/*[local-name()='DisplayString']")) { $sec.SelectSingleNode("./*[local-name()='Display']/*[local-name()='DisplayString']").InnerText.Trim() } elseif ($sec.SelectSingleNode("./*[local-name()='SettingNumber']")) { $sec.SelectSingleNode("./*[local-name()='SettingNumber']").InnerText.Trim() } elseif ($sec.SelectSingleNode("./*[local-name()='SettingBoolean']")) { $sec.SelectSingleNode("./*[local-name()='SettingBoolean']").InnerText.Trim() } else { $sec.InnerText.Trim() }

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

                # 5. Benutzerrechte
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
                            Explain   = "Registry: $keyPath\$valName = $valData ($regType)"
                            GpoName   = $GpoDisplayName
                            Datum     = $dateStr
                            Uhrzeit   = $timeStr
                        })
                    }
                }
            }
        }

        Parse-GpoSection $xml.GPO.Computer "Computer"
        Parse-GpoSection $xml.GPO.User "User"

        return $list
    }

    # =========================================================================
    # LOGIK TAB 1: ADSI GPO-Uebersicht & DDP-Integritaetspruefung
    # =========================================================================
    function Update-OverviewDisplay {
        if ($isClosing -or $form.IsDisposed -or $gridOvMaster.IsDisposed) { return }
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

        $lblLegendOverview.Text = "Status: $($arr.Count) von $($rawOverviewList.Count) GPOs  |  [Blau] Echte Default GPO  |  [Orange] DDP-Integritaetswarnung  |  [Gruen] OK  |  [Rot] Nicht OK"
    }

    function Invoke-LoadOverview {
        if ($isClosing -or $form.IsDisposed) { return }
        
        $pbarGlobal.Visible = $true
        $pbarGlobal.Minimum = 0
        $pbarGlobal.Value = 0
        $lblProgressInfo.Text = "Lade Active Directory Struktur..."
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        [System.Windows.Forms.Application]::DoEvents()

        $wmiSearcher = $null; $wmiResults = $null
        $linkSearcher = $null; $ouResults = $null
        $gpoSearcher = $null; $gpoResults = $null
        $wmiRoot = $null; $linkRoot = $null; $gpoRoot = $null

        try {
            $gpoLinksCache.Clear()
            $rawOverviewList.Clear()

            # 1. WMI Filter
            $lblProgressInfo.Text = "Lese WMI-Filter ein..."
            [System.Windows.Forms.Application]::DoEvents()

            $wmiMap = @{}
            $wmiRoot = [System.DirectoryServices.DirectoryEntry]::new("LDAP://CN=SOM,CN=WMIPolicy,CN=System,$domainDN")
            $wmiSearcher = [System.DirectoryServices.DirectorySearcher]::new($wmiRoot)
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
            $lblProgressInfo.Text = "Lese Verknuepfungen (OUs, Domaene, Sites) ein..."
            [System.Windows.Forms.Application]::DoEvents()

            $linkRoot = [System.DirectoryServices.DirectoryEntry]::new("LDAP://$domainDN")
            $linkSearcher = [System.DirectoryServices.DirectorySearcher]::new($linkRoot)
            $linkSearcher.Filter = "(|(objectClass=organizationalUnit)(objectClass=domainDNS))"
            $linkSearcher.PropertiesToLoad.AddRange(@("distinguishedName", "gPLink", "name"))
            $linkSearcher.SearchScope = [System.DirectoryServices.SearchScope]::Subtree
            
            $ouResults = $linkSearcher.FindAll()
            foreach ($ou in $ouResults) {
                if ($ou.Properties["gplink"]) {
                    $rawGpLink = $ou.Properties["gplink"][0]
                    $targetDN = $ou.Properties["distinguishedname"][0]
                    $matches = [regex]::Matches($rawGpLink, "cn=({?[a-fA-F0-9-]+}?)", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                    foreach ($m in $matches) {
                        $cleanGuid = $m.Groups[1].Value.Trim('{','}').ToUpper()
                        if (-not $gpoLinksCache.ContainsKey($cleanGuid)) {
                            $gpoLinksCache[$cleanGuid] = [System.Collections.Generic.List[string]]::new()
                        }
                        $gpoLinksCache[$cleanGuid].Add($targetDN)
                    }
                }
            }

            # 2b. Site-Verlinkungen
            try {
                $configDN = ([ADSI]"LDAP://RootDSE").configurationNamingContext.Value
                $siteRoot = [System.DirectoryServices.DirectoryEntry]::new("LDAP://CN=Sites,$configDN")
                $siteSearcher = [System.DirectoryServices.DirectorySearcher]::new($siteRoot)
                $siteSearcher.Filter = "(objectClass=site)"
                $siteSearcher.PropertiesToLoad.AddRange(@("distinguishedName", "gPLink", "name"))
                $siteResults = $siteSearcher.FindAll()
                foreach ($site in $siteResults) {
                    if ($site.Properties["gplink"]) {
                        $rawGpLink = $site.Properties["gplink"][0]
                        $targetDN = $site.Properties["distinguishedname"][0]
                        $matches = [regex]::Matches($rawGpLink, "cn=({?[a-fA-F0-9-]+}?)", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                        foreach ($m in $matches) {
                            $cleanGuid = $m.Groups[1].Value.Trim('{','}').ToUpper()
                            if (-not $gpoLinksCache.ContainsKey($cleanGuid)) {
                                $gpoLinksCache[$cleanGuid] = [System.Collections.Generic.List[string]]::new()
                            }
                            $gpoLinksCache[$cleanGuid].Add($targetDN)
                        }
                    }
                }
                $siteResults.Dispose(); $siteSearcher.Dispose(); $siteRoot.Dispose()
            } catch {}

            # 3. GPO Container & DDP-Integritaetspruefung
            $gpoRoot = [System.DirectoryServices.DirectoryEntry]::new("LDAP://CN=Policies,CN=System,$domainDN")
            $gpoSearcher = [System.DirectoryServices.DirectorySearcher]::new($gpoRoot)
            $gpoSearcher.Filter = "(objectClass=groupPolicyContainer)"
            $gpoSearcher.PropertiesToLoad.AddRange(@("displayName", "name", "flags", "gPCWQLFilter", "whenCreated", "whenChanged"))

            $gpoResults = $gpoSearcher.FindAll()
            $totalGpos = $gpoResults.Count
            $pbarGlobal.Maximum = [Math]::Max(1, $totalGpos)
            $currentIndex = 0
            $ddpConflictWarnings = [System.Collections.Generic.List[string]]::new()

            foreach ($g in $gpoResults) {
                $currentIndex++
                $rawGuid = $g.Properties["name"][0]
                $cleanGuid = $rawGuid.Trim('{','}').ToUpper()
                $displayName = if ($g.Properties["displayname"]) { $g.Properties["displayname"][0] } else { "{$cleanGuid}" }
                $flags = if ($g.Properties["flags"]) { [int]$g.Properties["flags"][0] } else { 0 }

                $pbarGlobal.Value = $currentIndex
                $lblProgressInfo.Text = "Lese GPO ($currentIndex / $totalGpos): $displayName"
                if ($currentIndex % 4 -eq 0) { [System.Windows.Forms.Application]::DoEvents() }

                $userStatus = if (($flags -band 1) -eq 1) { "Deaktiviert" } else { "Aktiviert" }
                $compStatus = if (($flags -band 2) -eq 2) { "Deaktiviert" } else { "Aktiviert" }

                # Strikte Pruefung auf die echte Microsoft DDP- & DDCP-Standard GUID
                $overallStatus = ""
                if ($cleanGuid -eq $script:StandardDdpGuid) {
                    if ($displayName -eq "Default Domain Policy") {
                        $overallStatus = "Sonderstellung (Default GPO)"
                    } else {
                        $overallStatus = "WARNUNG: Original DDP-GUID, aber umbenannt!"
                        $ddpConflictWarnings.Add("Original-GUID {$cleanGuid} ist umbenannt in: '$displayName'")
                    }
                }
                elseif ($displayName -like "*Default Domain Policy*") {
                    $overallStatus = "WARNUNG: Namensduplikat (Keine Standard-GUID!)"
                    $ddpConflictWarnings.Add("GPO '$displayName' traegt DDP-Namen, hat aber fremde GUID: {$cleanGuid}")
                }
                elseif ($cleanGuid -eq $script:StandardDdcpGuid) {
                    if ($displayName -eq "Default Domain Controllers Policy") {
                        $overallStatus = "Sonderstellung (Default GPO)"
                    } else {
                        $overallStatus = "WARNUNG: Original DDCP-GUID, aber umbenannt!"
                        $ddpConflictWarnings.Add("Original-DDCP {$cleanGuid} ist umbenannt in: '$displayName'")
                    }
                }
                elseif ($displayName -like "*Default Domain Controllers Policy*") {
                    $overallStatus = "WARNUNG: DDCP-Namensduplikat (Keine Standard-GUID!)"
                    $ddpConflictWarnings.Add("GPO '$displayName' traegt DDCP-Namen, hat aber fremde GUID: {$cleanGuid}")
                }
                else {
                    switch ($flags) {
                        1 { $overallStatus = "OK (Nur Computer)" }
                        2 { $overallStatus = "OK (Nur Benutzer)" }
                        0 { $overallStatus = "Nicht OK (Beide aktiviert)" }
                        3 { $overallStatus = "Nicht OK (Vollstaendig deaktiviert)" }
                        default { $overallStatus = "Nicht OK (Unbekannt: $flags)" }
                    }
                }

                $isLinked = ($gpoLinksCache.ContainsKey($cleanGuid) -and $gpoLinksCache[$cleanGuid].Count -gt 0)
                $linkedCount = if ($isLinked) { $gpoLinksCache[$cleanGuid].Count } else { 0 }

                $wmiFilterName = "-"
                $wmiFilterQuery = "-"
                if ($g.Properties["gpcwqlfilter"]) {
                    $rawWmi = $g.Properties["gpcwqlfilter"][0]
                    if ($rawWmi -match "({?[a-fA-F0-9-]+}?)") {
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
                    "GUID"          = "{$cleanGuid}"
                    "Erstellt am"   = $created
                    "Geaendert am"  = $changed
                })
            }

            Update-OverviewDisplay
            $lblProgressInfo.Text = "GPO-Einlesen abgeschlossen: $($rawOverviewList.Count) GPOs erfolgreich geladen."

            # Warnung bei erkannten DDP-Konflikten anzeigen
            if ($ddpConflictWarnings.Count -gt 0) {
                $warnMsg  = "ACHTUNG: Es wurden Unregelmaessigkeiten bei den Standard-Gruppenrichtlinien festgestellt!`r`n`r`n"
                $warnMsg += "Im Active Directory besitzt nur genau eine GPO die Microsoft-Standard-GUID '{31B2F340-016D-11D2-945F-00C04FB984F9}'.`r`n`r`n"
                $warnMsg += "Gefundene Konflikte:`r`n"
                foreach ($w in $ddpConflictWarnings) { $warnMsg += " - $w`r`n" }
                $warnMsg += "`r`nBetroffene Richtlinien sind in Register 1 farblich in ORANGE als WARNUNG hervorgehoben."
                
                [System.Windows.Forms.MessageBox]::Show($warnMsg, "Integritaets-Warnung Default Domain Policy", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            }
        } catch {
            $lblLegendOverview.Text = "Fehler: $($_.Exception.Message)"
            $lblProgressInfo.Text = "Fehler beim Einlesen: $($_.Exception.Message)"
        } finally {
            if ($wmiResults)   { $wmiResults.Dispose() }
            if ($wmiSearcher)  { $wmiSearcher.Dispose() }
            if ($wmiRoot)      { $wmiRoot.Dispose() }
            if ($ouResults)    { $ouResults.Dispose() }
            if ($linkSearcher) { $linkSearcher.Dispose() }
            if ($linkRoot)     { $linkRoot.Dispose() }
            if ($gpoResults)   { $gpoResults.Dispose() }
            if ($gpoSearcher)  { $gpoSearcher.Dispose() }
            if ($gpoRoot)      { $gpoRoot.Dispose() }
            $pbarGlobal.Visible = $false
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }

    # Snapshot speichern
    $btnSaveSnapshot.Add_Click({
        if ($rawOverviewList.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Keine GPO-Daten vorhanden. Bitte lesen Sie zuerst die GPOs ein.", "Hinweis", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Title = "GPO Snapshot speichern"
        $sfd.Filter = "GPO Snapshot (*.gposnap)|*.gposnap|XML Dateien (*.xml)|*.xml"
        $sfd.FileName = "GPO_Snapshot_${domainName}_$((Get-Date).ToString('yyyyMMdd_HHmm')).gposnap"
        $sfd.InitialDirectory = $txtBackupTargetDir.Text.Trim()

        if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                $snapshotData = @{
                    Version       = $script:ToolVersion
                    Timestamp     = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                    Domain        = $domainName
                    OverviewList  = @($rawOverviewList)
                    BackupList    = @($rawBackupList)
                    GpoLinksCache = $gpoLinksCache
                }
                $snapshotData | Export-Clixml -Path $sfd.FileName -Depth 5
                $lblProgressInfo.Text = "Snapshot erfolgreich gespeichert: $($sfd.FileName)"
                [System.Windows.Forms.MessageBox]::Show("GPO Snapshot mit $($rawOverviewList.Count) GPOs erfolgreich gespeichert!`n`nDatei: $($sfd.FileName)", "Snapshot gespeichert", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Fehler beim Speichern des Snapshots: $_", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    })

    # Snapshot laden
    $btnLoadSnapshot.Add_Click({
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Title = "GPO Snapshot oeffnen"
        $ofd.Filter = "GPO Snapshot (*.gposnap;*.xml)|*.gposnap;*.xml|Alle Dateien (*.*)|*.*"
        $ofd.InitialDirectory = $txtBackupTargetDir.Text.Trim()

        if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                $pbarGlobal.Visible = $true
                $pbarGlobal.Value = 20
                $lblProgressInfo.Text = "Lade Snapshot-Datei: $($ofd.FileName)..."
                $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
                [System.Windows.Forms.Application]::DoEvents()

                $snapshotData = Import-Clixml -Path $ofd.FileName -ErrorAction Stop

                if ($null -eq $snapshotData.OverviewList -or $null -eq $snapshotData.GpoLinksCache) {
                    throw "Die ausgewaehlte Datei enthaelt keine gueltigen GPO-Snapshot-Daten."
                }

                $rawOverviewList.Clear()
                foreach ($item in $snapshotData.OverviewList) { [void]$rawOverviewList.Add($item) }

                $gpoLinksCache.Clear()
                foreach ($k in $snapshotData.GpoLinksCache.Keys) {
                    $gpoLinksCache[$k] = $snapshotData.GpoLinksCache[$k]
                }

                $rawBackupList.Clear()
                if ($snapshotData.BackupList) {
                    foreach ($b in $snapshotData.BackupList) { [void]$rawBackupList.Add($b) }
                }

                $pbarGlobal.Value = 70
                [System.Windows.Forms.Application]::DoEvents()

                Update-OverviewDisplay
                Update-BackupGridDisplay
                Update-SettingsGpoDropdown

                $comboGpo1.Items.Clear()
                $comboGpo2.Items.Clear()
                foreach ($it in $rawOverviewList) {
                    [void]$comboGpo1.Items.Add($it."GPO Name")
                    [void]$comboGpo2.Items.Add($it."GPO Name")
                }
                [void]$comboGpo1.Items.Add($script:DdpBaselineName)
                [void]$comboGpo2.Items.Add($script:DdpBaselineName)

                if ($comboGpo1.Items.Count -gt 0) { $comboGpo1.SelectedIndex = 0 }
                if ($comboGpo2.Items.Count -gt 1) { $comboGpo2.SelectedIndex = 1 }

                $pbarGlobal.Value = 100
                $lblProgressInfo.Text = "Snapshot geladen ($($snapshotData.Timestamp)): $($rawOverviewList.Count) GPOs"
                [System.Windows.Forms.MessageBox]::Show("Snapshot erfolgreich geladen!`n`nErstellt am: $($snapshotData.Timestamp)`nDomaene:     $($snapshotData.Domain)`nGPOs:        $($rawOverviewList.Count)", "Snapshot aktiv", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Fehler beim Laden des Snapshots: $_", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            } finally {
                $pbarGlobal.Visible = $false
                $form.Cursor = [System.Windows.Forms.Cursors]::Default
            }
        }
    })

    $gridOvMaster.Add_DataBindingComplete({
        if ($isClosing -or $form.IsDisposed -or $gridOvMaster.IsDisposed) { return }
        foreach ($row in $gridOvMaster.Rows) {
            $status = [string]$row.Cells["Gesamt-Status"].Value

            if ($status -match "^WARNUNG") {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 238, 204)
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(180, 50, 0)
                $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(255, 210, 160)
                $row.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
                $row.DefaultCellStyle.Font = New-Object System.Drawing.Font($gridOvMaster.Font, [System.Drawing.FontStyle]::Bold)
            }
            elseif ($status -match "Sonderstellung") {
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

    $btnLoadAdGpos.Add_Click({ Invoke-LoadOverview })
    $comboViewMode.Add_SelectedIndexChanged({ Update-OverviewDisplay })
    $txtOverviewSearch.Add_TextChanged({ Update-OverviewDisplay })

    $gridOvMaster.Add_SelectionChanged({
        if ($isClosing -or $form.IsDisposed -or $gridOvMaster.IsDisposed -or $gridOvDetails.IsDisposed) { return }
        if ($gridOvMaster.SelectedRows.Count -gt 0) {
            $selectedRow = $gridOvMaster.SelectedRows[0]
            $guid = [string]$selectedRow.Cells["GUID"].Value
            $cleanGuid = $guid.Trim('{','}').ToUpper()
            $gName = [string]$selectedRow.Cells["GPO Name"].Value

            $lblOvDetailsTitle.Text = "Verlinkungsziele fuer:`r`n[$gName]"

            $arrDetails = [System.Collections.ArrayList]::new()
            if ($cleanGuid -and $gpoLinksCache.ContainsKey($cleanGuid) -and $gpoLinksCache[$cleanGuid].Count -gt 0) {
                foreach ($dn in $gpoLinksCache[$cleanGuid]) {
                    $type = "Organizational Unit (OU)"
                    $simpleName = $dn
                    if ($dn -match "^OU=([^,]+)") {
                        $simpleName = $matches[1]
                        $type = "OU"
                    } elseif ($dn -match "^DC=") {
                        $type = "Domaenen-Root"
                        $simpleName = $domainName
                    } elseif ($dn -match "^CN=([^,]+),CN=Sites") {
                        $type = "Active Directory Site"
                        $simpleName = $matches[1]
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
            $cleanGuid = [string]$row.GUID.Trim('{','}').ToUpper()
            $linksStr = if ($gpoLinksCache.ContainsKey($cleanGuid)) { ($gpoLinksCache[$cleanGuid] -join " | ") } else { "Keine" }
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
    # LOGIK TAB 2: GPO Settings Inspector & Lazy Loading
    # =========================================================================
    $btnBrowseSettingsDir.Add_Click({
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description = "Waehlen Sie das Export-Verzeichnis fuer die Richtlinien aus:"
        $dialog.SelectedPath = $txtSettingsExportDir.Text.Trim()
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtSettingsExportDir.Text = $dialog.SelectedPath
            $txtBackupTargetDir.Text = $dialog.SelectedPath
        }
    })

    function Update-SettingsGpoDropdown {
        if ($isClosing -or $form.IsDisposed -or $comboGpo.IsDisposed) { return }
        $mode = $comboSettingsViewMode.SelectedItem

        $matchingGpos = $allGposCache | Where-Object {
            $cleanGuid = $_.Id.ToString().Trim('{','}').ToUpper()
            $isLinked = ($gpoLinksCache.ContainsKey($cleanGuid) -and $gpoLinksCache[$cleanGuid].Count -gt 0)
            switch ($mode) {
                "Nur verlinkte GPOs"              { $isLinked }
                "Nicht verlinkte GPOs (Unlinked)" { -not $isLinked }
                default                           { $true }
            }
        }

        $comboGpo.Items.Clear()
        $summaryLabel = switch ($mode) {
            "Nur verlinkte GPOs"              { "-- ALLE verlinkten GPOs laden ($($matchingGpos.Count)) --" }
            "Nicht verlinkte GPOs (Unlinked)" { "-- ALLE ungelinkten GPOs laden ($($matchingGpos.Count)) --" }
            default                           { "-- ALLE GPOs laden ($($matchingGpos.Count)) --" }
        }

        [void]$comboGpo.Items.Add($summaryLabel)
        foreach ($g in $matchingGpos) {
            [void]$comboGpo.Items.Add($g.DisplayName)
        }

        # Bevorzuge die echte Default Domain Policy oder die 1. Einzel-GPO statt '-- ALLE --'
        if ($comboGpo.Items.Count -gt 1) {
            $defaultIdx = -1
            for ($i = 1; $i -lt $comboGpo.Items.Count; $i++) {
                if ($comboGpo.Items[$i] -eq "Default Domain Policy") { $defaultIdx = $i; break }
            }
            if ($defaultIdx -gt 0) { $comboGpo.SelectedIndex = $defaultIdx } else { $comboGpo.SelectedIndex = 1 }
        } elseif ($comboGpo.Items.Count -gt 0) {
            $comboGpo.SelectedIndex = 0
        }
    }

    $comboSettingsViewMode.Add_SelectedIndexChanged({ Update-SettingsGpoDropdown })

    function Update-SettingsGridDisplay {
        if ($isClosing -or $form.IsDisposed -or $gridSettings.IsDisposed) { return }
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

    function Invoke-LoadSettings {
        if ($isClosing -or $form.IsDisposed) { return }
        $selectedOption = $comboGpo.SelectedItem
        if ([string]::IsNullOrWhiteSpace($selectedOption)) { return }

        $lblSettingsStatus.Text = "Lese Einstellungen ein..."
        $form.Refresh()

        $rawSettingsList.Clear()

        if ($selectedOption.StartsWith("-- ALLE")) {
            $mode = $comboSettingsViewMode.SelectedItem
            $targetGpos = @($allGposCache | Where-Object {
                $cleanGuid = $_.Id.ToString().Trim('{','}').ToUpper()
                $isLinked = ($gpoLinksCache.ContainsKey($cleanGuid) -and $gpoLinksCache[$cleanGuid].Count -gt 0)
                switch ($mode) {
                    "Nur verlinkte GPOs"              { $isLinked }
                    "Nicht verlinkte GPOs (Unlinked)" { -not $isLinked }
                    default                           { $true }
                }
            })

            $pbarGlobal.Visible = $true
            $pbarGlobal.Minimum = 0
            $pbarGlobal.Maximum = [Math]::Max(1, $targetGpos.Count)
            $pbarGlobal.Value = 0

            $curr = 0
            foreach ($g in $targetGpos) {
                $curr++
                $pbarGlobal.Value = $curr
                $lblProgressInfo.Text = "Lese Richtlinien-Details ($curr / $($targetGpos.Count)): $($g.DisplayName)"
                [System.Windows.Forms.Application]::DoEvents()
                try {
                    $items = Get-ParsedGpoSettings -GpoId $g.Id -GpoDisplayName $g.DisplayName
                    foreach ($item in $items) { $rawSettingsList.Add($item) }
                } catch {}
            }

            $pbarGlobal.Visible = $false
            $lblProgressInfo.Text = "Bereit."
        } else {
            try {
                $gpo = $allGposCache | Where-Object { $_.DisplayName -eq $selectedOption } | Select-Object -First 1
                if ($gpo) {
                    $items = Get-ParsedGpoSettings -GpoId $gpo.Id -GpoDisplayName $gpo.DisplayName
                    foreach ($item in $items) { $rawSettingsList.Add($item) }
                }
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Fehler beim Abruf: $_", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }

        Update-SettingsGridDisplay
    }

    $btnLoadSettings.Add_Click({ Invoke-LoadSettings })
    $txtFilter.Add_TextChanged({ Update-SettingsGridDisplay })

    $gridSettings.Add_SelectionChanged({
        if ($isClosing -or $form.IsDisposed -or $gridSettings.IsDisposed -or $txtDescription.IsDisposed) { return }
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
        $targetBase = $txtSettingsExportDir.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($targetBase)) { $targetBase = "C:\Install\Backup\GPO" }
        if (-not (Test-Path $targetBase)) { New-Item -ItemType Directory -Path $targetBase -Force | Out-Null }

        $dateStr = Get-Date -Format "yyyyMMdd"
        $timeStr = Get-Date -Format "HHmm"
        $csvFile = Join-Path $targetBase "GPO_Settings_Export_${dateStr}_${timeStr}.csv"

        $rawSettingsList | Export-Csv -Path $csvFile -Delimiter ";" -NoTypeInformation -Encoding UTF8
        [System.Windows.Forms.MessageBox]::Show("Erfolgreich $($rawSettingsList.Count) Eintraege exportiert!`n`nPfad: $csvFile", "Export abgeschlossen", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    })

    # =========================================================================
    # LOGIK TAB 3: GPO Backup & Audit (Schneller LDAP-Cache)
    # =========================================================================
    $btnBrowseFolder.Add_Click({
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description = "Waehlen Sie das Basis-Backup-Verzeichnis fuer die GPOs aus:"
        $dialog.SelectedPath = $txtBackupTargetDir.Text.Trim()
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtBackupTargetDir.Text = $dialog.SelectedPath
            $txtSettingsExportDir.Text = $dialog.SelectedPath
        }
    })

    function Update-BackupGridDisplay {
        if ($isClosing -or $form.IsDisposed -or $gridGpos.IsDisposed) { return }
        $mode = $comboBackupFilter.SelectedItem
        $filterText = $txtBackupSearch.Text.Trim()

        $filtered = $rawBackupList | Where-Object {
            $item = $_
            $matchMode = switch ($mode) {
                "Nur verlinkte GPOs"              { $item."Link-Anzahl" -gt 0 }
                "Nicht verlinkte GPOs (Unlinked)" { $item."Link-Anzahl" -eq 0 }
                default                           { $true }
            }
            $matchSearch = if ([string]::IsNullOrWhiteSpace($filterText)) { $true } else {
                $item."GPO Name" -like "*$filterText*" -or $item."Status" -like "*$filterText*" -or $item."GPO ID (GUID)" -like "*$filterText*"
            }
            $matchMode -and $matchSearch
        }

        $arr = [System.Collections.ArrayList]::new()
        foreach ($it in $filtered) { [void]$arr.Add($it) }
        $gridGpos.DataSource = $arr

        if ($gridGpos.Columns["Link-Anzahl"]) {
            $gridGpos.Columns["Link-Anzahl"].Visible = $false
        }

        if ($mode -ne "Alle GPOs" -or (-not [string]::IsNullOrWhiteSpace($filterText))) {
            $btnBackupAll.Text = "Gefilterte GPOs sichern ($($arr.Count))"
            $btnExportAllHtml.Text = "HTML (Gefilterte: $($arr.Count))"
        } else {
            $btnBackupAll.Text = "ALLE GPOs sichern ($($arr.Count))"
            $btnExportAllHtml.Text = "HTML (Alle: $($arr.Count))"
        }

        $lblGpoGrid.Text = "1. Gruppenrichtlinien der Domaene ($($arr.Count) angezeigt):"
    }

    $comboBackupFilter.Add_SelectedIndexChanged({ Update-BackupGridDisplay })
    $txtBackupSearch.Add_TextChanged({ Update-BackupGridDisplay })

    # Laedt blitzschnell ueber den ADSI-Cache ohne XML-Reports abzufragen
    function Invoke-LoadGpos {
        if ($isClosing -or $form.IsDisposed -or $gridGpos.IsDisposed) { return }
        $gridGpos.DataSource = $null
        $rawBackupList.Clear()

        foreach ($g in $allGposCache) {
            $cleanGuid = $g.Id.ToString().Trim('{','}').ToUpper()
            $isLinked = ($gpoLinksCache.ContainsKey($cleanGuid) -and $gpoLinksCache[$cleanGuid].Count -gt 0)
            $linkCount = if ($isLinked) { $gpoLinksCache[$cleanGuid].Count } else { 0 }
            $isLinkedText = if ($linkCount -gt 0) { "JA ($linkCount)" } else { "NEIN (Unlinked)" }

            $rawBackupList.Add([PSCustomObject]@{
                "GPO Name"      = $g.DisplayName
                "Status"        = $g.GpoStatus.ToString()
                "Verknuepft"    = $isLinkedText
                "Link-Anzahl"   = $linkCount
                "GPO ID (GUID)" = $g.Id.ToString()
            })
        }

        Update-BackupGridDisplay
    }

    $btnLoadGpos.Add_Click({ Invoke-LoadGpos })

    $gridGpos.Add_SelectionChanged({
        if ($isClosing -or $form.IsDisposed -or $gridGpos.IsDisposed -or $gridLinks.IsDisposed) { return }
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
                $validLinks = @($links) | Where-Object { $null -ne $_ -and (-not [string]::IsNullOrWhiteSpace($_.SOMPath)) }

                if ($validLinks -and @($validLinks).Count -gt 0) {
                    foreach ($l in $validLinks) {
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
        $validLinks = @($links) | Where-Object { $null -ne $_ -and (-not [string]::IsNullOrWhiteSpace($_.SOMPath)) }
        $timeFormatted = Get-Date -Format "dd.MM.yyyy HH:mm:ss"

        "==================================================" | Out-File -FilePath $txtPath -Encoding UTF8
        " GPO AUDIT & BACKUP REPORT"                         | Out-File -FilePath $txtPath -Append -Encoding UTF8
        "==================================================" | Out-File -FilePath $txtPath -Append -Encoding UTF8
        "GPO Name        : $gpoName"                          | Out-File -FilePath $txtPath -Append -Encoding UTF8
        "GPO ID (GUID)   : $gpoGuid"                          | Out-File -FilePath $txtPath -Append -Encoding UTF8
        "Zeitpunkt       : $timeFormatted"                    | Out-File -FilePath $txtPath -Append -Encoding UTF8
        "Ziel-Ordner     : $targetPath"                       | Out-File -FilePath $txtPath -Append -Encoding UTF8
        "--------------------------------------------------" | Out-File -FilePath $txtPath -Append -Encoding UTF8
        if ($validLinks -and @($validLinks).Count -gt 0) {
            "Verknuepfungen:" | Out-File -FilePath $txtPath -Append -Encoding UTF8
            foreach ($l in $validLinks) {
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

        $targetItems = @($gridGpos.DataSource)
        if ($null -eq $targetItems -or $targetItems.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Keine GPOs in der aktuellen Ansicht zum Sichern vorhanden.", "Hinweis", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        $filterMode = $comboBackupFilter.SelectedItem
        $confirmMsg = "Moechten Sie wirklich die $($targetItems.Count) Richtlinien der aktuellen Ansicht ('$filterMode') nach '$basePath' sichern?"
        $confirm = [System.Windows.Forms.MessageBox]::Show($confirmMsg, "Bestaetigung Sicherung", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        $txtLog.AppendText("[$((Get-Date).ToString('HH:mm:ss'))] === Starte Sicherung von $($targetItems.Count) GPOs ('$filterMode') ===`r`n")
        $form.Refresh()

        foreach ($g in $targetItems) {
            try {
                Backup-SingleGPOWithLog -gpoGuid $g."GPO ID (GUID)" -gpoName $g."GPO Name" -basePath $basePath
            } catch {
                $txtLog.AppendText("[$((Get-Date).ToString('HH:mm:ss'))] [FEHLER] bei '$($g."GPO Name")': $_`r`n")
            }
        }
        $txtLog.AppendText("[$((Get-Date).ToString('HH:mm:ss'))] === Sicherung abgeschlossen! ===`r`n")
        [System.Windows.Forms.MessageBox]::Show("Sicherung von $($targetItems.Count) GPOs erfolgreich abgeschlossen!", "Fertig", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    })

    $btnExportBackupCsv.Add_Click({
        $targetItems = @($gridGpos.DataSource)
        if ($null -eq $targetItems -or $targetItems.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Keine Daten zum Exportieren vorhanden.", "Hinweis", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        $targetBase = $txtBackupTargetDir.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($targetBase)) { $targetBase = "C:\Install\Backup\GPO" }
        if (-not (Test-Path $targetBase)) { New-Item -ItemType Directory -Path $targetBase -Force | Out-Null }

        $dateStr = Get-Date -Format "yyyyMMdd"
        $timeStr = Get-Date -Format "HHmm"
        $csvFile = Join-Path $targetBase "GPO_Backup_List_${dateStr}_${timeStr}.csv"

        $targetItems | Export-Csv -Path $csvFile -Delimiter ";" -NoTypeInformation -Encoding UTF8
        [System.Windows.Forms.MessageBox]::Show("GPO-Liste erfolgreich exportiert!`n`nPfad: $csvFile", "Export abgeschlossen", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    })

    # HTML Report fuer die einzeln markierte GPO
    $btnExportSelectedHtml.Add_Click({
        $targetBase = $txtBackupTargetDir.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($targetBase)) { $targetBase = "C:\Install\Backup\GPO" }
        if (-not (Test-Path $targetBase)) { New-Item -ItemType Directory -Path $targetBase -Force | Out-Null }

        if ($gridGpos.SelectedRows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Bitte waehlen Sie zuerst eine GPO aus der Tabelle aus.", "Hinweis", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        $selectedItem = $gridGpos.SelectedRows[0].DataBoundItem
        $gpoName = $selectedItem."GPO Name"
        $guid = $selectedItem."GPO ID (GUID)"
        $safeName = $gpoName -replace '[\\/:*?"<>|]', '_'
        $timestamp = Get-Date -Format "yyyyMMdd_HHmm"

        $htmlDir = Join-Path $targetBase "HTML_Reports"
        if (-not (Test-Path $htmlDir)) { New-Item -ItemType Directory -Path $htmlDir -Force | Out-Null }

        $htmlFile = Join-Path $htmlDir "${safeName}_${timestamp}.html"
        try {
            Get-GPOReport -Guid $guid -ReportType Html -Path $htmlFile -ErrorAction Stop
            $txtLog.AppendText("[$((Get-Date).ToString('HH:mm:ss'))] [OK] HTML-Bericht fuer '$gpoName' erstellt: $htmlFile`r`n")
            $openChoice = [System.Windows.Forms.MessageBox]::Show("HTML-Bericht fuer '$gpoName' erfolgreich erstellt!`n`nPfad: $htmlFile`n`nMoechten Sie den Bericht jetzt im Browser oeffnen?", "HTML Export", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Information)
            if ($openChoice -eq [System.Windows.Forms.DialogResult]::Yes) {
                Start-Process -FilePath $htmlFile
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Erstellen des HTML-Berichts: $_", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })

    # HTML Report fuer alle GPOs der aktuellen Ansicht
    $btnExportAllHtml.Add_Click({
        $targetBase = $txtBackupTargetDir.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($targetBase)) { $targetBase = "C:\Install\Backup\GPO" }
        if (-not (Test-Path $targetBase)) { New-Item -ItemType Directory -Path $targetBase -Force | Out-Null }

        $displayedItems = @($gridGpos.DataSource)
        if ($null -eq $displayedItems -or $displayedItems.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Keine GPOs in der aktuellen Ansicht zum Exportieren vorhanden.", "Hinweis", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        $filterMode = $comboBackupFilter.SelectedItem
        $timestamp = Get-Date -Format "yyyyMMdd_HHmm"

        $confirm = [System.Windows.Forms.MessageBox]::Show("Moechten Sie fuer alle $($displayedItems.Count) GPOs der aktuellen Ansicht ('$filterMode') jeweils einen HTML-Bericht erstellen?", "HTML Massenexport", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        $modeFolder = switch ($filterMode) {
            "Nicht verlinkte GPOs (Unlinked)" { "Unlinked" }
            "Nur verlinkte GPOs"              { "Linked" }
            default                           { "All" }
        }

        $htmlDir = Join-Path $targetBase "GPO_HTML_Reports_${modeFolder}_${timestamp}"
        if (-not (Test-Path $htmlDir)) { New-Item -ItemType Directory -Path $htmlDir -Force | Out-Null }

        $pbarGlobal.Visible = $true
        $pbarGlobal.Minimum = 0
        $pbarGlobal.Maximum = [Math]::Max(1, $displayedItems.Count)
        $pbarGlobal.Value = 0

        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $txtLog.AppendText("[$((Get-Date).ToString('HH:mm:ss'))] === Starte HTML-Massenexport fuer $($displayedItems.Count) GPOs ('$filterMode') ===`r`n")
        $form.Refresh()

        $exportedCount = 0
        $cur = 0
        foreach ($item in $displayedItems) {
            $cur++
            $pbarGlobal.Value = $cur
            $lblProgressInfo.Text = "Exportiere HTML ($cur / $($displayedItems.Count)): $($item.'GPO Name')"
            [System.Windows.Forms.Application]::DoEvents()
            try {
                $safeName = $item."GPO Name" -replace '[\\/:*?"<>|]', '_'
                $guid = $item."GPO ID (GUID)"
                $htmlFile = Join-Path $htmlDir "${safeName}.html"
                Get-GPOReport -Guid $guid -ReportType Html -Path $htmlFile -ErrorAction Stop
                $exportedCount++
            } catch {
                $txtLog.AppendText("[$((Get-Date).ToString('HH:mm:ss'))] [FEHLER] HTML fuer '$($item.'GPO Name')': $_`r`n")
            }
        }

        $pbarGlobal.Visible = $false
        $lblProgressInfo.Text = "Bereit."
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $txtLog.AppendText("[$((Get-Date).ToString('HH:mm:ss'))] === $exportedCount HTML-Berichte erfolgreich erstellt in: $htmlDir ===`r`n")

        $openChoice = [System.Windows.Forms.MessageBox]::Show("$exportedCount HTML-Berichte erfolgreich erstellt!`n`nOrdner: $htmlDir`n`nMoechten Sie den Ordner im Explorer oeffnen?", "HTML Massenexport abgeschlossen", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Information)
        if ($openChoice -eq [System.Windows.Forms.DialogResult]::Yes) {
            Start-Process -FilePath "explorer.exe" -ArgumentList $htmlDir
        }
    })

    # =========================================================================
    # LOGIK TAB 4: GPO-Vergleich (Diff Engine)
    # =========================================================================
    function Update-CompareGridDisplay {
        if ($isClosing -or $form.IsDisposed -or $gridCompare.IsDisposed) { return }
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

    function Invoke-GpoCompare {
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
            $list1 = if ($gpoName1 -eq $script:DdpBaselineName) {
                Get-DefaultDomainPolicyBaseline
            } else {
                $gpoObj1 = $allGposCache | Where-Object { $_.DisplayName -eq $gpoName1 } | Select-Object -First 1
                if (-not $gpoObj1) { $gpoObj1 = Get-GPO -Name $gpoName1 -ErrorAction Stop }
                Get-ParsedGpoSettings -GpoId $gpoObj1.Id -GpoDisplayName $gpoName1
            }

            $list2 = if ($gpoName2 -eq $script:DdpBaselineName) {
                Get-DefaultDomainPolicyBaseline
            } else {
                $gpoObj2 = $allGposCache | Where-Object { $_.DisplayName -eq $gpoName2 } | Select-Object -First 1
                if (-not $gpoObj2) { $gpoObj2 = Get-GPO -Name $gpoName2 -ErrorAction Stop }
                Get-ParsedGpoSettings -GpoId $gpoObj2.Id -GpoDisplayName $gpoName2
            }

            $dict1 = @{}
            foreach ($item in $list1) { 
                $normKey = Get-NormalizedPolicyKey $item
                $dict1[$normKey] = $item 
            }

            $dict2 = @{}
            foreach ($item in $list2) { 
                $normKey = Get-NormalizedPolicyKey $item
                $dict2[$normKey] = $item 
            }

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
    }

    $btnCompare.Add_Click({ Invoke-GpoCompare })

    $btnCompareDdpBaseline.Add_Click({
        # Bevorzuge die Richtlinie mit der echten Standard-DDP GUID
        $realDdp = $allGposCache | Where-Object { $_.Id.ToString().Trim('{','}').ToUpper() -eq $script:StandardDdpGuid } | Select-Object -First 1
        if ($realDdp) {
            $comboGpo1.SelectedItem = $realDdp.DisplayName
        } else {
            $ddpCandidate = $comboGpo1.Items | Where-Object { $_ -match "^Default Domain Policy$" } | Select-Object -First 1
            if ($ddpCandidate) { $comboGpo1.SelectedItem = $ddpCandidate } else { $comboGpo1.SelectedIndex = 0 }
        }

        $comboGpo2.SelectedItem = $script:DdpBaselineName
        $chkOnlyDiffs.Checked = $false
        Invoke-GpoCompare
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
        $txtBackupTargetDir.SelectionStart = 0
        $txtBackupTargetDir.SelectionLength = 0
        $txtSettingsExportDir.SelectionStart = 0
        $txtSettingsExportDir.SelectionLength = 0

        # Alle GPOs einmalig cachen
        $allGposCache.Clear()
        $gpos = Get-GPO -All | Sort-Object DisplayName
        foreach ($g in $gpos) { [void]$allGposCache.Add($g) }

        # Dropdowns fuer Tab 4 befuellen (inkl. MS DDP Baseline)
        $comboGpo1.Items.Clear()
        $comboGpo2.Items.Clear()
        foreach ($g in $allGposCache) {
            [void]$comboGpo1.Items.Add($g.DisplayName)
            [void]$comboGpo2.Items.Add($g.DisplayName)
        }
        [void]$comboGpo1.Items.Add($script:DdpBaselineName)
        [void]$comboGpo2.Items.Add($script:DdpBaselineName)

        if ($comboGpo1.Items.Count -gt 0) { $comboGpo1.SelectedIndex = 0 }
        if ($comboGpo2.Items.Count -gt 1) { $comboGpo2.SelectedIndex = 1 }

        # 1. Register 1 laden (Schnelle LDAP-Struktur)
        Invoke-LoadOverview

        # 2. Register 2 vorbereiten (Lazy Loading: laedt nur die ausgewaehlte Einzel-GPO, kein Freeze!)
        Update-SettingsGpoDropdown
        Invoke-LoadSettings

        # 3. Register 3 vorbereiten (Sofortige Bestandsliste ueber LDAP-Cache, 0.01 s statt 60 s!)
        Invoke-LoadGpos
    })

    $form.Add_FormClosing({
        $isClosing = $true
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
