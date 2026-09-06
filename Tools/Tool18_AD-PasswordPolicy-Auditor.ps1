# ==============================================================================
# TOOL 18: AD KENNWORT-RICHTLINIEN, FINE-GRAINED PSO & GPO AUDITOR
# ==============================================================================

if ($host.Name -eq 'Windows PowerShell ISE Host') {
    if ($PSCommandPath) {
        Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
        return
    }
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.DirectoryServices

if (-not [System.Windows.Forms.Application]::RenderWithVisualStyles) {
    try { [System.Windows.Forms.Application]::EnableVisualStyles() } catch {}
}

# STANDARD CACHE-DATEINAME
$script:ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { [System.AppDomain]::CurrentDomain.BaseDirectory }
$script:DefaultCacheFile = Join-Path $script:ScriptDir "AD_PasswordPolicy_Cache.json"
$script:ActiveLoadedSnapshotPath = "Kein Snapshot geladen (Live-Modus)"

# THEME & SCHRIFTARTEN
$script:Theme = @{
    FontFamily         = "Segoe UI"
    TitleFont          = New-Object System.Drawing.Font("Segoe UI", 12.0, [System.Drawing.FontStyle]::Bold)
    SubTitleFont       = New-Object System.Drawing.Font("Segoe UI", 9.0, [System.Drawing.FontStyle]::Regular)
    BaseFont           = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Regular)
    BoldFont           = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    ButtonFont         = New-Object System.Drawing.Font("Segoe UI", 9.0, [System.Drawing.FontStyle]::Bold)
    GridHeaderFont     = New-Object System.Drawing.Font("Segoe UI", 9.0, [System.Drawing.FontStyle]::Bold)
    GridCellFont       = New-Object System.Drawing.Font("Segoe UI", 9.0, [System.Drawing.FontStyle]::Regular)
    GridHeaderHeight   = 36
    GridRowHeight      = 30
    HeaderBackColor    = [System.Drawing.Color]::FromArgb(241, 245, 249)
    HeaderForeColor    = [System.Drawing.Color]::FromArgb(30, 41, 59)
    GridLineColor      = [System.Drawing.Color]::FromArgb(226, 232, 240)
    RowAlternateColor  = [System.Drawing.Color]::FromArgb(248, 250, 252)
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
                if ($null -eq $val -or $val -eq "") { return "" }
                if ($val -match '^\s*(\d+)') { return [int64]$Matches[1] }
                if ($val -match '^\d{2}\.\d{2}\.\d{4}') {
                    try { return [datetime]::ParseExact($val.ToString().Trim(), @("dd.MM.yyyy HH:mm:ss", "dd.MM.yyyy HH:mm", "dd.MM.yyyy"), $null) } catch { return $val }
                }
                return $val.ToString()
            }
            Descending = (-not $state.Asc)
        }

        $targetGrid.SuspendLayout()
        $arr = [System.Collections.ArrayList]::new()
        foreach ($item in $sorted) { [void]$arr.Add($item) }
        $targetGrid.DataSource = $null
        $targetGrid.DataSource = $arr

        foreach ($c in $targetGrid.Columns) {
            $c.HeaderCell.SortGlyphDirection = [System.Windows.Forms.SortOrder]::None
        }
        $targetGrid.Columns[$e.ColumnIndex].HeaderCell.SortGlyphDirection = `
            $(if ($state.Asc) { [System.Windows.Forms.SortOrder]::Ascending } else { [System.Windows.Forms.SortOrder]::Descending })

        $targetGrid.ResumeLayout()
    })
}

function Start-PasswordPolicyAuditor {
    $isOnline = $false
    $domDN = ""
    $domainName = "OFFLINE"
    try {
        $rootDSE = [ADSI]"LDAP://RootDSE"
        $domDN = $rootDSE.defaultNamingContext.ToString()
        $domainName = ($domDN -replace 'DC=','' -replace ',','.').ToUpper()
        if ($domDN) { $isOnline = $true }
    } catch {
        $isOnline = $false
    }

    if (-not $isOnline -and -not (Test-Path $script:DefaultCacheFile)) {
        $diag = [System.Windows.Forms.MessageBox]::Show(
            "Keine Active Directory Verbindung verfügbar und keine Snapshot-Datei gefunden.`n`n$(Join-Path $script:ScriptDir 'AD_PasswordPolicy_Cache.json')`n`nMöchten Sie eine bestehende Snapshot-Datei manuell auswählen?",
            "Offline - Snapshot laden",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($diag -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Tool 18: Active Directory Kennwortrichtlinien, PSO & GPO Auditor"
    $form.Size = New-Object System.Drawing.Size(1380, 890)
    $form.MinimumSize = New-Object System.Drawing.Size(1080, 690)
    $form.StartPosition = "CenterScreen"
    $form.Font = $script:Theme.BaseFont
    $form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

    try {
        # WARNBANNER (Ganz oben)
        $pnlWarningBanner = New-Object System.Windows.Forms.Panel
        $pnlWarningBanner.Dock = [System.Windows.Forms.DockStyle]::Top
        $pnlWarningBanner.Height = 36
        $pnlWarningBanner.BackColor = [System.Drawing.Color]::FromArgb(254, 242, 242)
        $pnlWarningBanner.Visible = $false
        $form.Controls.Add($pnlWarningBanner)

        $lblWarningBanner = New-Object System.Windows.Forms.Label
        $lblWarningBanner.Dock = [System.Windows.Forms.DockStyle]::Fill
        $lblWarningBanner.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
        $lblWarningBanner.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
        $lblWarningBanner.ForeColor = [System.Drawing.Color]::FromArgb(185, 28, 28)
        $lblWarningBanner.Padding = New-Object System.Windows.Forms.Padding(16, 0, 0, 0)
        $pnlWarningBanner.Controls.Add($lblWarningBanner)

        # HEADER PANEL
        $headerPanel = New-Object System.Windows.Forms.Panel
        $headerPanel.Dock = [System.Windows.Forms.DockStyle]::Top
        $headerPanel.Height = 84
        $headerPanel.BackColor = [System.Drawing.Color]::White
        $form.Controls.Add($headerPanel)

        $lblTitle = New-Object System.Windows.Forms.Label
        $lblTitle.Text = "AD Kennwortrichtlinien, PSOs & GPO Audit"
        $lblTitle.Font = $script:Theme.TitleFont
        $lblTitle.ForeColor = [System.Drawing.Color]::FromArgb(15, 23, 42)
        $lblTitle.Location = New-Object System.Drawing.Point(18, 10)
        $lblTitle.AutoSize = $true
        $headerPanel.Controls.Add($lblTitle)

        $lblSubTitle = New-Object System.Windows.Forms.Label
        $lblSubTitle.Text = if ($isOnline) { "Domäne: $domainName ($domDN) | Angemeldet: $env:USERDOMAIN\$env:USERNAME" } else { "Modus: OFFLINE / SNAPSHOT-ANALYSE" }
        $lblSubTitle.Font = $script:Theme.SubTitleFont
        $lblSubTitle.ForeColor = if ($isOnline) { [System.Drawing.Color]::FromArgb(100, 116, 139) } else { [System.Drawing.Color]::FromArgb(194, 65, 12) }
        $lblSubTitle.Location = New-Object System.Drawing.Point(20, 36)
        $lblSubTitle.AutoSize = $true
        $headerPanel.Controls.Add($lblSubTitle)

        $lblSnapshotInfo = New-Object System.Windows.Forms.Label
        $lblSnapshotInfo.Text = "Snapshot: $script:ActiveLoadedSnapshotPath"
        $lblSnapshotInfo.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Regular)
        $lblSnapshotInfo.ForeColor = [System.Drawing.Color]::FromArgb(14, 116, 144)
        $lblSnapshotInfo.Location = New-Object System.Drawing.Point(20, 58)
        $lblSnapshotInfo.AutoSize = $true
        $headerPanel.Controls.Add($lblSnapshotInfo)

        # Action-Buttons
        $flowHeaderBtns = New-Object System.Windows.Forms.FlowLayoutPanel
        $flowHeaderBtns.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
        $flowHeaderBtns.Location = New-Object System.Drawing.Point(680, 16)
        $flowHeaderBtns.Size = New-Object System.Drawing.Size(670, 52)
        $flowHeaderBtns.FlowDirection = [System.Windows.Forms.FlowDirection]::RightToLeft
        $flowHeaderBtns.WrapContents = $false
        $headerPanel.Controls.Add($flowHeaderBtns)

        $btnRefresh = New-Object System.Windows.Forms.Button
        $btnRefresh.Text = "Live-Scan (DC)"
        $btnRefresh.Height = 36
        $btnRefresh.Width = 120
        $btnRefresh.BackColor = [System.Drawing.Color]::FromArgb(71, 85, 105)
        $btnRefresh.ForeColor = [System.Drawing.Color]::White
        $btnRefresh.Font = $script:Theme.ButtonFont
        $btnRefresh.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnRefresh.FlatAppearance.BorderSize = 0
        $btnRefresh.Margin = New-Object System.Windows.Forms.Padding(6, 0, 0, 0)
        $flowHeaderBtns.Controls.Add($btnRefresh)

        $btnSaveSnapshot = New-Object System.Windows.Forms.Button
        $btnSaveSnapshot.Text = "Snapshot speichern"
        $btnSaveSnapshot.Height = 36
        $btnSaveSnapshot.Width = 150
        $btnSaveSnapshot.BackColor = [System.Drawing.Color]::FromArgb(14, 116, 144)
        $btnSaveSnapshot.ForeColor = [System.Drawing.Color]::White
        $btnSaveSnapshot.Font = $script:Theme.ButtonFont
        $btnSaveSnapshot.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnSaveSnapshot.FlatAppearance.BorderSize = 0
        $btnSaveSnapshot.Margin = New-Object System.Windows.Forms.Padding(6, 0, 0, 0)
        $flowHeaderBtns.Controls.Add($btnSaveSnapshot)

        $btnLoadSnapshot = New-Object System.Windows.Forms.Button
        $btnLoadSnapshot.Text = "Snapshot laden"
        $btnLoadSnapshot.Height = 36
        $btnLoadSnapshot.Width = 135
        $btnLoadSnapshot.BackColor = [System.Drawing.Color]::FromArgb(100, 116, 139)
        $btnLoadSnapshot.ForeColor = [System.Drawing.Color]::White
        $btnLoadSnapshot.Font = $script:Theme.ButtonFont
        $btnLoadSnapshot.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnLoadSnapshot.FlatAppearance.BorderSize = 0
        $btnLoadSnapshot.Margin = New-Object System.Windows.Forms.Padding(6, 0, 0, 0)
        $flowHeaderBtns.Controls.Add($btnLoadSnapshot)

        $btnExport = New-Object System.Windows.Forms.Button
        $btnExport.Text = "CSV Export (Tab)"
        $btnExport.Height = 36
        $btnExport.Width = 130
        $btnExport.BackColor = [System.Drawing.Color]::FromArgb(16, 124, 65)
        $btnExport.ForeColor = [System.Drawing.Color]::White
        $btnExport.Font = $script:Theme.ButtonFont
        $btnExport.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnExport.FlatAppearance.BorderSize = 0
        $btnExport.Margin = New-Object System.Windows.Forms.Padding(0)
        $flowHeaderBtns.Controls.Add($btnExport)

        # FOOTER
        $bottomPanel = New-Object System.Windows.Forms.Panel
        $bottomPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
        $bottomPanel.Height = 32
        $bottomPanel.BackColor = [System.Drawing.Color]::FromArgb(241, 245, 249)
        $form.Controls.Add($bottomPanel)

        $lblStatus = New-Object System.Windows.Forms.Label
        $lblStatus.Location = New-Object System.Drawing.Point(15, 6)
        $lblStatus.AutoSize = $true
        $lblStatus.Font = $script:Theme.BaseFont
        $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(51, 65, 85)
        $lblStatus.Text = "Bereit."
        $bottomPanel.Controls.Add($lblStatus)

        $progressBar = New-Object System.Windows.Forms.ProgressBar
        $progressBar.Size = New-Object System.Drawing.Size(180, 16)
        $progressBar.Location = New-Object System.Drawing.Point(1170, 8)
        $progressBar.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
        $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
        $progressBar.MarqueeAnimationSpeed = 30
        $progressBar.Visible = $false
        $bottomPanel.Controls.Add($progressBar)

        # TAB CONTROL
        $tabControl = New-Object System.Windows.Forms.TabControl
        $tabControl.Dock = [System.Windows.Forms.DockStyle]::Fill
        $tabControl.Padding = New-Object System.Drawing.Point(14, 8)
        $tabControl.ItemSize = New-Object System.Drawing.Size(185, 30)
        $tabControl.Font = $script:Theme.BoldFont
        $form.Controls.Add($tabControl)
        $tabControl.BringToFront()

        $styleGrid = {
            param($grid)
            $grid.Dock = [System.Windows.Forms.DockStyle]::Fill
            $grid.BackgroundColor = [System.Drawing.Color]::White
            $grid.BorderStyle = [System.Windows.Forms.BorderStyle]::None
            $grid.CellBorderStyle = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
            $grid.GridColor = $script:Theme.GridLineColor
            $grid.ReadOnly = $true
            $grid.AllowUserToAddRows = $false
            $grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
            $grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
            $grid.RowHeadersVisible = $false
            $grid.EnableHeadersVisualStyles = $false
            $grid.ColumnHeadersDefaultCellStyle.BackColor = $script:Theme.HeaderBackColor
            $grid.ColumnHeadersDefaultCellStyle.ForeColor = $script:Theme.HeaderForeColor
            $grid.ColumnHeadersDefaultCellStyle.Font = $script:Theme.GridHeaderFont
            $grid.ColumnHeadersHeight = $script:Theme.GridHeaderHeight
            $grid.DefaultCellStyle.Font = $script:Theme.GridCellFont
            $grid.RowTemplate.Height = $script:Theme.GridRowHeight
            $grid.AlternatingRowsDefaultCellStyle.BackColor = $script:Theme.RowAlternateColor
        }

        # --- TAB 1: Standard Domänenrichtlinie (MIT QUELL-INFOBOX) ---
        $tabDomain = New-Object System.Windows.Forms.TabPage
        $tabDomain.Text = "Standard Domänenrichtlinie"
        $tabDomain.BackColor = [System.Drawing.Color]::White
        $tabDomain.Padding = New-Object System.Windows.Forms.Padding(6)
        $tabControl.TabPages.Add($tabDomain)

        $pnlDomainSource = New-Object System.Windows.Forms.Panel
        $pnlDomainSource.Dock = [System.Windows.Forms.DockStyle]::Top
        $pnlDomainSource.Height = 44
        $pnlDomainSource.BackColor = [System.Drawing.Color]::FromArgb(240, 253, 244)
        $pnlDomainSource.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
        $tabDomain.Controls.Add($pnlDomainSource)

        $lblDomainSourceInfo = New-Object System.Windows.Forms.Label
        $lblDomainSourceInfo.Dock = [System.Windows.Forms.DockStyle]::Fill
        $lblDomainSourceInfo.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
        $lblDomainSourceInfo.Font = New-Object System.Drawing.Font("Segoe UI", 9.0, [System.Drawing.FontStyle]::Bold)
        $lblDomainSourceInfo.ForeColor = [System.Drawing.Color]::FromArgb(22, 101, 52)
        $lblDomainSourceInfo.Padding = New-Object System.Windows.Forms.Padding(12, 0, 0, 0)
        $lblDomainSourceInfo.Text = "Richtlinien-Quelle: Lese Quell-Informationen..."
        $pnlDomainSource.Controls.Add($lblDomainSourceInfo)

        $gridDomain = New-Object System.Windows.Forms.DataGridView
        & $styleGrid $gridDomain
        $tabDomain.Controls.Add($gridDomain)
        $gridDomain.BringToFront()

        # --- TAB 2: Fine-Grained PSOs ---
        $tabPSO = New-Object System.Windows.Forms.TabPage
        $tabPSO.Text = "Fine-Grained Policies (PSO)"
        $tabPSO.BackColor = [System.Drawing.Color]::White
        $tabPSO.Padding = New-Object System.Windows.Forms.Padding(6)
        $tabControl.TabPages.Add($tabPSO)

        $gridPSO = New-Object System.Windows.Forms.DataGridView
        & $styleGrid $gridPSO
        $tabPSO.Controls.Add($gridPSO)

        # --- TAB 3: Effektive Benutzerrichtlinie ---
        $tabUser = New-Object System.Windows.Forms.TabPage
        $tabUser.Text = "Effektive Benutzerrichtlinie"
        $tabUser.BackColor = [System.Drawing.Color]::White
        $tabUser.Padding = New-Object System.Windows.Forms.Padding(6)
        $tabControl.TabPages.Add($tabUser)

        $userTopPanel = New-Object System.Windows.Forms.FlowLayoutPanel
        $userTopPanel.Dock = [System.Windows.Forms.DockStyle]::Top
        $userTopPanel.Height = 56
        $userTopPanel.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 252)
        $userTopPanel.Padding = New-Object System.Windows.Forms.Padding(12, 12, 12, 8)
        $userTopPanel.WrapContents = $false
        $tabUser.Controls.Add($userTopPanel)

        $lblUser = New-Object System.Windows.Forms.Label
        $lblUser.Text = "Benutzername (sAMAccountName oder UPN):"
        $lblUser.AutoSize = $true
        $lblUser.Font = $script:Theme.BoldFont
        $lblUser.Margin = New-Object System.Windows.Forms.Padding(0, 6, 10, 0)
        $userTopPanel.Controls.Add($lblUser)

        $txtUserCheck = New-Object System.Windows.Forms.TextBox
        $txtUserCheck.Size = New-Object System.Drawing.Size(260, 26)
        $txtUserCheck.Text = $env:USERNAME
        $txtUserCheck.Font = $script:Theme.BaseFont
        $txtUserCheck.Margin = New-Object System.Windows.Forms.Padding(0, 2, 12, 0)
        $userTopPanel.Controls.Add($txtUserCheck)

        $btnCheckUser = New-Object System.Windows.Forms.Button
        $btnCheckUser.Text = "Richtlinie prüfen"
        $btnCheckUser.Size = New-Object System.Drawing.Size(150, 30)
        $btnCheckUser.BackColor = [System.Drawing.Color]::FromArgb(15, 118, 110)
        $btnCheckUser.ForeColor = [System.Drawing.Color]::White
        $btnCheckUser.Font = $script:Theme.ButtonFont
        $btnCheckUser.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnCheckUser.FlatAppearance.BorderSize = 0
        $userTopPanel.Controls.Add($btnCheckUser)

        $gridUserPolicy = New-Object System.Windows.Forms.DataGridView
        & $styleGrid $gridUserPolicy
        $tabUser.Controls.Add($gridUserPolicy)
        $gridUserPolicy.BringToFront()

        # --- TAB 4: GPO Kennwort-Audit ---
        $tabGPO = New-Object System.Windows.Forms.TabPage
        $tabGPO.Text = "GPO Kennwort-Audit"
        $tabGPO.BackColor = [System.Drawing.Color]::White
        $tabGPO.Padding = New-Object System.Windows.Forms.Padding(6)
        $tabControl.TabPages.Add($tabGPO)

        $gpoTopPanel = New-Object System.Windows.Forms.FlowLayoutPanel
        $gpoTopPanel.Dock = [System.Windows.Forms.DockStyle]::Top
        $gpoTopPanel.Height = 56
        $gpoTopPanel.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 252)
        $gpoTopPanel.Padding = New-Object System.Windows.Forms.Padding(12, 12, 12, 8)
        $gpoTopPanel.WrapContents = $false
        $tabGPO.Controls.Add($gpoTopPanel)

        $lblGpoSearch = New-Object System.Windows.Forms.Label
        $lblGpoSearch.Text = "GPO filtern / gezielt prüfen:"
        $lblGpoSearch.AutoSize = $true
        $lblGpoSearch.Font = $script:Theme.BoldFont
        $lblGpoSearch.Margin = New-Object System.Windows.Forms.Padding(0, 6, 10, 0)
        $gpoTopPanel.Controls.Add($lblGpoSearch)

        $txtGpoFilter = New-Object System.Windows.Forms.TextBox
        $txtGpoFilter.Size = New-Object System.Drawing.Size(260, 26)
        $txtGpoFilter.Text = "Passwortrichtlinie"
        $txtGpoFilter.Font = $script:Theme.BaseFont
        $txtGpoFilter.Margin = New-Object System.Windows.Forms.Padding(0, 2, 12, 0)
        $gpoTopPanel.Controls.Add($txtGpoFilter)

        $btnFilterGpo = New-Object System.Windows.Forms.Button
        $btnFilterGpo.Text = "GPO prüfen"
        $btnFilterGpo.Size = New-Object System.Drawing.Size(125, 30)
        $btnFilterGpo.BackColor = [System.Drawing.Color]::FromArgb(37, 99, 235)
        $btnFilterGpo.ForeColor = [System.Drawing.Color]::White
        $btnFilterGpo.Font = $script:Theme.ButtonFont
        $btnFilterGpo.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnFilterGpo.FlatAppearance.BorderSize = 0
        $gpoTopPanel.Controls.Add($btnFilterGpo)

        $btnResetGpo = New-Object System.Windows.Forms.Button
        $btnResetGpo.Text = "Alle anzeigen"
        $btnResetGpo.Size = New-Object System.Drawing.Size(120, 30)
        $btnResetGpo.BackColor = [System.Drawing.Color]::FromArgb(100, 116, 139)
        $btnResetGpo.ForeColor = [System.Drawing.Color]::White
        $btnResetGpo.Font = $script:Theme.ButtonFont
        $btnResetGpo.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnResetGpo.FlatAppearance.BorderSize = 0
        $btnResetGpo.Margin = New-Object System.Windows.Forms.Padding(8, 0, 0, 0)
        $gpoTopPanel.Controls.Add($btnResetGpo)

        $gridGPOAudit = New-Object System.Windows.Forms.DataGridView
        & $styleGrid $gridGPOAudit
        $tabGPO.Controls.Add($gridGPOAudit)
        $gridGPOAudit.BringToFront()

        $gridGPOAudit.Add_DataBindingComplete({
            foreach ($row in $gridGPOAudit.Rows) {
                $st = [string]$row.Cells["Status"].Value
                if ($st -like "*Umbenannte Default Domain Policy*") {
                    $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(238, 242, 255)
                    $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(67, 56, 202)
                    $row.DefaultCellStyle.Font = New-Object System.Drawing.Font($gridGPOAudit.Font, [System.Drawing.FontStyle]::Bold)
                } elseif ($st -like "*Kein Original*") {
                    $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 237, 213)
                    $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(194, 65, 12)
                    $row.DefaultCellStyle.Font = New-Object System.Drawing.Font($gridGPOAudit.Font, [System.Drawing.FontStyle]::Bold)
                } elseif ($st -like "*Fehlkonfiguriert*") {
                    $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(254, 243, 199)
                    $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(146, 64, 14)
                } elseif ($st -like "*Wirkungslos*") {
                    $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(254, 226, 226)
                    $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(153, 27, 27)
                } elseif ($st -like "*Aktiv*") {
                    $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(240, 253, 244)
                    $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(22, 101, 52)
                    $row.DefaultCellStyle.Font = New-Object System.Drawing.Font($gridGPOAudit.Font, [System.Drawing.FontStyle]::Bold)
                }
            }
        })

        # --- TAB 5: EXPERT INFO (VERTIKALE MASTER-DETAIL AUFTEILUNG) ---
        $tabExpert = New-Object System.Windows.Forms.TabPage
        $tabExpert.Text = "Expert Info"
        $tabExpert.BackColor = [System.Drawing.Color]::White
        $tabExpert.Padding = New-Object System.Windows.Forms.Padding(6)
        $tabControl.TabPages.Add($tabExpert)

        $splitExpert = New-Object System.Windows.Forms.SplitContainer
        $splitExpert.Dock = [System.Windows.Forms.DockStyle]::Fill
        $splitExpert.Orientation = [System.Windows.Forms.Orientation]::Horizontal
        $splitExpert.SplitterDistance = 240
        $splitExpert.SplitterWidth = 6
        $tabExpert.Controls.Add($splitExpert)

        # Oben: Tabelle
        $pnlExpTop = New-Object System.Windows.Forms.Panel
        $pnlExpTop.Dock = [System.Windows.Forms.DockStyle]::Fill
        $splitExpert.Panel1.Controls.Add($pnlExpTop)

        $lblExpTableTitle = New-Object System.Windows.Forms.Label
        $lblExpTableTitle.Text = "Diagnose-Themen & Befunde (Klick für Details):"
        $lblExpTableTitle.Dock = [System.Windows.Forms.DockStyle]::Top
        $lblExpTableTitle.Height = 24
        $lblExpTableTitle.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
        $pnlExpTop.Controls.Add($lblExpTableTitle)

        $gridExpert = New-Object System.Windows.Forms.DataGridView
        & $styleGrid $gridExpert
        $gridExpert.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
        $gridExpert.MultiSelect = $false
        $pnlExpTop.Controls.Add($gridExpert)
        $gridExpert.BringToFront()

        # Unten: Detailbereich
        $pnlExpBottom = New-Object System.Windows.Forms.Panel
        $pnlExpBottom.Dock = [System.Windows.Forms.DockStyle]::Fill
        $pnlExpBottom.Padding = New-Object System.Windows.Forms.Padding(0, 4, 0, 0)
        $splitExpert.Panel2.Controls.Add($pnlExpBottom)

        $lblExpDetailTitle = New-Object System.Windows.Forms.Label
        $lblExpDetailTitle.Text = "Detaildiagnose & Handlungsempfehlung:"
        $lblExpDetailTitle.Dock = [System.Windows.Forms.DockStyle]::Top
        $lblExpDetailTitle.Height = 24
        $lblExpDetailTitle.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
        $pnlExpBottom.Controls.Add($lblExpDetailTitle)

        $txtExpertDetail = New-Object System.Windows.Forms.TextBox
        $txtExpertDetail.Dock = [System.Windows.Forms.DockStyle]::Fill
        $txtExpertDetail.Multiline = $true
        $txtExpertDetail.ReadOnly = $true
        $txtExpertDetail.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
        $txtExpertDetail.Font = New-Object System.Drawing.Font("Consolas", 9.5)
        $txtExpertDetail.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 252)
        $txtExpertDetail.ForeColor = [System.Drawing.Color]::FromArgb(15, 23, 42)
        $pnlExpBottom.Controls.Add($txtExpertDetail)
        $txtExpertDetail.BringToFront()

        Enable-UniversalGridSorting -Grid $gridDomain
        Enable-UniversalGridSorting -Grid $gridPSO
        Enable-UniversalGridSorting -Grid $gridUserPolicy
        Enable-UniversalGridSorting -Grid $gridGPOAudit
        Enable-UniversalGridSorting -Grid $gridExpert

        $convertTimeSpan = {
            param($val, $asDays = $false)
            if ($null -eq $val -or $val -eq 0 -or $val -eq "") { return "Nie ablaufend / Deaktiviert" }
            try {
                $ticks = 0
                if ($val -is [System.TimeSpan]) {
                    $ts = $val
                } else {
                    if ($val.GetType().Name -eq "__ComObject" -or $val.GetType().FullName -like "*LargeInteger*") {
                        $high = [int64]$val.GetType().InvokeMember("HighPart", [System.Reflection.BindingFlags]::GetProperty, $null, $val, $null)
                        $low  = [int64]$val.GetType().InvokeMember("LowPart", [System.Reflection.BindingFlags]::GetProperty, $null, $val, $null)
                        $ticks = ($high -shl 32) + ($low -band 0xFFFFFFFF)
                    } else {
                        $ticks = [int64]$val
                    }
                    if ($ticks -eq 0 -or $ticks -eq [int64]::MinValue) { return "Nie ablaufend / Deaktiviert" }
                    if ($ticks -lt 0) { $ticks = -$ticks }
                    $ts = [TimeSpan]::FromTicks($ticks)
                }

                if ($asDays) {
                    $days = [Math]::Round($ts.TotalDays, 1)
                    return "$days Tage"
                } else {
                    if ($ts.TotalMinutes -ge 60) {
                        $hours = [Math]::Round($ts.TotalHours, 1)
                        return "$hours Stunden"
                    } else {
                        $mins = [Math]::Round($ts.TotalMinutes, 0)
                        return "$mins Minuten"
                    }
                }
            } catch {
                return $val.ToString()
            }
        }

        # CACHE VARIABLEN
        $script:cachedDomainPolicies            = @()
        $script:cachedPSOs                     = @()
        $script:cachedUserPolicy               = @()
        $script:cachedGPOAudits                = @()
        $script:cachedFullPSOs                 = @()
        $script:cachedDefaultPolicyGpoSettings = @()
        $script:cachedExpertTopics             = @()
        $script:lastCheckedUser                = $env:USERNAME

        # HILFSFUNKTION: TAB 1 QUELL-INFOBOX AKTUALISIEREN
        $updateDomainSourceBox = {
            $origGpo = $script:cachedGPOAudits | Where-Object { $_."GPO GUID" -eq "{31B2F340-016D-11D2-945F-00C04FB984F9}" } | Select-Object -First 1
            $activeDomainGpo = $script:cachedGPOAudits | Where-Object { $_.Status -like "*Aktiv*" -and $_."Verknüpft an (SOM)" -like "*Domänen-Root*" } | Select-Object -First 1

            $gpoName = if ($activeDomainGpo) { $activeDomainGpo.'GPO Name' } elseif ($origGpo) { $origGpo.'GPO Name' } else { "Unbekannt" }
            $gpoGuid = if ($activeDomainGpo) { $activeDomainGpo.'GPO GUID' } elseif ($origGpo) { $origGpo.'GPO GUID' } else { "{31B2F340-016D-11D2-945F-00C04FB984F9}" }
            
            $isRenamed = ($origGpo -and $origGpo."GPO Name" -ne "Default Domain Policy")

            if ($script:cachedDefaultPolicyGpoSettings.Count -gt 0) {
                if ($isRenamed) {
                    $pnlDomainSource.BackColor = [System.Drawing.Color]::FromArgb(254, 243, 199)
                    $lblDomainSourceInfo.ForeColor = [System.Drawing.Color]::FromArgb(146, 64, 14)
                    $lblDomainSourceInfo.Text = "Werte stammen aus GPO: '$gpoName' (Original Default Domain Policy, umbenannt!) | GUID: $gpoGuid"
                } else {
                    $pnlDomainSource.BackColor = [System.Drawing.Color]::FromArgb(240, 253, 244)
                    $lblDomainSourceInfo.ForeColor = [System.Drawing.Color]::FromArgb(22, 101, 52)
                    $lblDomainSourceInfo.Text = "Werte stammen aus GPO: '$gpoName' (Original Default Domain Policy an Domänen-Wurzel) | GUID: $gpoGuid"
                }
            } else {
                $pnlDomainSource.BackColor = [System.Drawing.Color]::FromArgb(255, 237, 213)
                $lblDomainSourceInfo.ForeColor = [System.Drawing.Color]::FromArgb(194, 65, 12)
                $lblDomainSourceInfo.Text = "Werte stammen aus LDAP-Attributen des Domänenkopfs ($domDN), da in GPO '$gpoName' nicht definiert! (Erklärt z. B. 37 Min.)"
            }
        }

        # EXPERT-TOPICS GENERIEREN
        $buildExpertTopics = {
            $topics = [System.Collections.Generic.List[PSCustomObject]]::new()

            $origGpo = $script:cachedGPOAudits | Where-Object { $_."GPO GUID" -eq "{31B2F340-016D-11D2-945F-00C04FB984F9}" } | Select-Object -First 1
            $isRenamed = ($origGpo -and $origGpo."GPO Name" -ne "Default Domain Policy")

            $gpoStatusBrief = if (-not $origGpo) {
                "Original nicht gefunden"
            } elseif ($isRenamed) {
                "UMBENANNT in '$($origGpo.'GPO Name')'"
            } else {
                "Original vorhanden & korrekt benannt"
            }

            $curGpoName = if ($origGpo) { $origGpo.'GPO Name' } else { 'Nicht gefunden' }
            $curGpoGuid = if ($origGpo) { $origGpo.'GPO GUID' } else { '-' }
            $curGpoSom  = if ($origGpo) { $origGpo.'Verknüpft an (SOM)' } else { '-' }
            $curGpoHint = if ($isRenamed) {
                "ACHTUNG: Die Richtlinie wurde umbenannt! Es wird dringend empfohlen, den Namen in der GPMC wieder auf 'Default Domain Policy' zurückzusetzen."
            } else {
                "Alles in Ordnung. Die Original-GPO existiert und trägt ihren Standardnamen."
            }

            $gpoDetailText = @"
ECHTHEIT DER DEFAULT DOMAIN POLICY:
======================================================================
Original-System-GUID von Microsoft:
  {31B2F340-016D-11D2-945F-00C04FB984F9}

BEFUND:
  * Aktueller Name im AD:  '$curGpoName'
  * Originaler Name:       'Default Domain Policy'
  * Gefundene GUID:        $curGpoGuid
  * Verknüpfungsort:       $curGpoSom

HINWEIS:
$curGpoHint
"@

            $topics.Add([PSCustomObject]@{
                "Kategorie" = "GPO-Architektur"
                "Thema"     = "Status Default Domain Policy"
                "Ergebnis"  = $gpoStatusBrief
                "Details"   = $gpoDetailText
            })

            $lockoutDetailText = @"
HERKUNFT DER '37 MINUTEN' (BZW. KRUMMER ZEITEN):
======================================================================
Active Directory speichert Sperrzeiten intern in 100-Nanosekunden-Ticks:
  -22.295.600.000 Ticks = 2.229,56 Sekunden = ~37,16 Minuten

WARUM STEHT DAS IN KEINER GPO?
Steht in der 'Default Domain Policy' die Kontosperrung auf 'Nicht definiert',
greift Active Directory auf die alten LDAP-Rohwerte am Domänenkopf zurück.

LÖSUNG:
In gpmc.msc -> Default Domain Policy -> Computerkonfiguration -> Sicherheitseinstellungen -> Kontorichtlinien -> Kontosperrungsrichtlinie feste Werte eintragen (z. B. 30 Minuten).
"@

            $topics.Add([PSCustomObject]@{
                "Kategorie" = "Kontosperrung"
                "Thema"     = "Dauer der Kontosperrung (37 Minuten)"
                "Ergebnis"  = "LDAP-Rohwert am Domänenkopf (-22.295.600.000 Ticks)"
                "Details"   = $lockoutDetailText
            })

            $defGpoValBrief = if ($script:cachedDefaultPolicyGpoSettings.Count -gt 0) {
                "$($script:cachedDefaultPolicyGpoSettings.Count) Einstellung(en) im XML definiert"
            } else {
                "Keine expliziten Werte (Fallback auf LDAP)"
            }

            $formattedXmlSettings = if ($script:cachedDefaultPolicyGpoSettings.Count -gt 0) {
                "Konfigurierte Werte laut GPO-XML:`r`n`r`n" + (($script:cachedDefaultPolicyGpoSettings | ForEach-Object { "  * $_" }) -join "`r`n")
            } else {
                "In der GPO sind keine expliziten Werte für Kontorichtlinien definiert (Status: 'Nicht definiert'). Es greifen die LDAP-Werte der Domänenwurzel."
            }

            $topics.Add([PSCustomObject]@{
                "Kategorie" = "GPO-Konfiguration"
                "Thema"     = "Werte in Default Domain Policy"
                "Ergebnis"  = $defGpoValBrief
                "Details"   = $formattedXmlSettings
            })

            $topics.Add([PSCustomObject]@{
                "Kategorie" = "Richtlinien-Vorrang"
                "Thema"     = "GPO vs. Fine-Grained PSOs"
                "Ergebnis"  = "PSO überschreibt GPO vollständig"
                "Details"   = "PSOs (msDS-PasswordSettings) besitzen immer Vorrang vor GPOs an der Domänenwurzel für die zugewiesenen Konten und Gruppen."
            })

            $topics.Add([PSCustomObject]@{
                "Kategorie" = "Best Practice"
                "Thema"     = "OU-Verlinkte Kennwort-GPOs"
                "Ergebnis"  = "Wirkungslos für Domänenbenutzer"
                "Details"   = "Kennwort-GPOs an OUs wirken NIE auf Domänenbenutzer, sondern ausschließlich auf lokale SAM-Konten der dort liegenden Rechner."
            })

            $script:cachedExpertTopics = $topics

            $arrExp = [System.Collections.ArrayList]::new()
            foreach ($t in $topics) { [void]$arrExp.Add($t) }
            $gridExpert.DataSource = $arrExp

            if ($gridExpert.Columns["Details"]) {
                $gridExpert.Columns["Details"].Visible = $false
            }

            if ($gridExpert.Rows.Count -gt 0) {
                $gridExpert.Rows[0].Selected = $true
                $txtExpertDetail.Text = [string]$gridExpert.Rows[0].Cells["Details"].Value
            }
        }

        $gridExpert.Add_SelectionChanged({
            if ($gridExpert.SelectedRows.Count -gt 0) {
                $txtExpertDetail.Text = [string]$gridExpert.SelectedRows[0].Cells["Details"].Value
            }
        })

        $gridExpert.Add_DataBindingComplete({
            foreach ($row in $gridExpert.Rows) {
                $erg = [string]$row.Cells["Ergebnis"].Value
                if ($erg -like "*UMBENANNT*" -or $erg -like "*nicht gefunden*") {
                    $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(254, 226, 226)
                    $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(185, 28, 28)
                    $row.DefaultCellStyle.Font = New-Object System.Drawing.Font($gridExpert.Font, [System.Drawing.FontStyle]::Bold)
                } elseif ($erg -like "*korrekt benannt*" -or $erg -like "*Vorrang*") {
                    $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(240, 253, 244)
                    $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(22, 101, 52)
                }
            }
        })

        # SNAPSHOT-LOGIK
        $saveSnapshotToFile = {
            param([string]$FilePath)
            try {
                $exportPackage = [ordered]@{
                    "SnapshotVersion"               = "1.2"
                    "DomainName"                    = $domainName
                    "DomainDN"                      = $domDN
                    "CreatedAt"                     = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                    "SourceFile"                    = [System.IO.Path]::GetFileName($FilePath)
                    "DomainPolicies"                = $script:cachedDomainPolicies
                    "PSOs"                          = $script:cachedPSOs
                    "FullPSOs"                      = $script:cachedFullPSOs
                    "GPOAudits"                     = $script:cachedGPOAudits
                    "DefaultPolicyGpoSettings"      = $script:cachedDefaultPolicyGpoSettings
                }
                $json = $exportPackage | ConvertTo-Json -Depth 6
                [System.IO.File]::WriteAllText($FilePath, $json, [System.Text.Encoding]::UTF8)
                $script:ActiveLoadedSnapshotPath = "$([System.IO.Path]::GetFileName($FilePath)) ($((Get-Date).ToString('HH:mm:ss')))"
                $lblSnapshotInfo.Text = "Snapshot: $script:ActiveLoadedSnapshotPath"
                return $true
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Fehler beim Speichern:`n$($_.Exception.Message)", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return $false
            }
        }

        $loadSnapshotFromFile = {
            param([string]$FilePath)
            if (-not (Test-Path $FilePath)) { return $false }
            try {
                $rawJson = [System.IO.File]::ReadAllText($FilePath, [System.Text.Encoding]::UTF8)
                $pkg = $rawJson | ConvertFrom-Json

                $script:cachedDomainPolicies            = @($pkg.DomainPolicies)
                $script:cachedPSOs                     = @($pkg.PSOs)
                $script:cachedFullPSOs                 = @($pkg.FullPSOs)
                $script:cachedGPOAudits                = @($pkg.GPOAudits)
                $script:cachedDefaultPolicyGpoSettings = @($pkg.DefaultPolicyGpoSettings)

                if ($pkg.DomainName) { $domainName = $pkg.DomainName }
                if ($pkg.DomainDN)   { $domDN = $pkg.DomainDN }

                $arrD = [System.Collections.ArrayList]::new()
                foreach ($r in $script:cachedDomainPolicies) { [void]$arrD.Add($r) }
                $gridDomain.DataSource = $arrD

                $arrP = [System.Collections.ArrayList]::new()
                foreach ($p in $script:cachedPSOs) { [void]$arrP.Add($p) }
                $gridPSO.DataSource = $arrP

                $arrG = [System.Collections.ArrayList]::new()
                foreach ($g in $script:cachedGPOAudits) { [void]$arrG.Add($g) }
                $gridGPOAudit.DataSource = $arrG

                $fileInfo = Get-Item $FilePath
                $script:ActiveLoadedSnapshotPath = "$([System.IO.Path]::GetFileName($FilePath)) | Erstellt: $($pkg.CreatedAt) | $([Math]::Round($fileInfo.Length / 1KB, 1)) KB"
                $lblSnapshotInfo.Text = "Snapshot geladen: $script:ActiveLoadedSnapshotPath"
                $lblSubTitle.Text = "Modus: SNAPSHOT-CACHE ($($pkg.CreatedAt)) | Domäne: $domainName"
                $lblSubTitle.ForeColor = [System.Drawing.Color]::FromArgb(14, 116, 144)
                $lblStatus.Text = "Snapshot geladen aus '$([System.IO.Path]::GetFileName($FilePath))'."

                $origGpo = $script:cachedGPOAudits | Where-Object { $_."GPO GUID" -eq "{31B2F340-016D-11D2-945F-00C04FB984F9}" } | Select-Object -First 1
                if ($origGpo -and $origGpo."GPO Name" -ne "Default Domain Policy") {
                    $pnlWarningBanner.Visible = $true
                    $lblWarningBanner.Text = "ACHTUNG: Die 'Default Domain Policy' wurde umbenannt in '$($origGpo.'GPO Name')'! (GUID: $($origGpo.'GPO GUID'))"
                } else {
                    $pnlWarningBanner.Visible = $false
                }

                & $updateDomainSourceBox
                & $buildExpertTopics
                return $true
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Fehler beim Laden:`n$($_.Exception.Message)", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return $false
            }
        }

        # LIVE SCAN ROUTINE
        $loadAllPoliciesLive = {
            if (-not $isOnline) {
                [System.Windows.Forms.MessageBox]::Show("Keine Verbindung zum DC. System läuft offline.", "Offline", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }

            $progressBar.Visible = $true
            $lblStatus.Text = "Lese Richtlinien live aus dem Active Directory..."
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            $form.Refresh()

            $script:cachedDomainPolicies = @()
            $script:cachedPSOs = @()
            $script:cachedFullPSOs = @()
            $script:cachedGPOAudits = @()
            $script:cachedDefaultPolicyGpoSettings = @()

            try {
                $domEntry = [ADSI]"LDAP://$domDN"

                # 1. Standard-Domänenrichtlinie
                $rawMinLen    = $domEntry.Properties["minPwdLength"].Value
                $rawPwdHist   = $domEntry.Properties["pwdHistoryLength"].Value
                $rawPwdProps  = $domEntry.Properties["pwdProperties"].Value
                $rawMaxAge    = $domEntry.Properties["maxPwdAge"].Value
                $rawMinAge    = $domEntry.Properties["minPwdAge"].Value
                $rawLockThr   = $domEntry.Properties["lockoutThreshold"].Value
                $rawLockDur   = $domEntry.Properties["lockoutDuration"].Value
                $rawLockWin   = $domEntry.Properties["lockOutObservationWindow"].Value

                $minLen = if ($null -ne $rawMinLen) { "$rawMinLen Zeichen" } else { "0 Zeichen" }
                $pwdHist = if ($null -ne $rawPwdHist) { "$rawPwdHist Passwörter gespeichert" } else { "0" }

                $pwdPropsInt = 0
                if ($null -ne $rawPwdProps) { [int]::TryParse($rawPwdProps.ToString(), [ref]$pwdPropsInt) | Out-Null }
                $complexity = ($pwdPropsInt -band 1) -eq 1
                $reversible = ($pwdPropsInt -band 16) -eq 16

                $domList = @(
                    [PSCustomObject]@{ "Kategorie" = "Kennwort-Richtlinie"; "Eigenschaft" = "Minimale Kennwortlänge"; "Wert" = $minLen },
                    [PSCustomObject]@{ "Kategorie" = "Kennwort-Richtlinie"; "Eigenschaft" = "Kennworthistorie (Chronik)"; "Wert" = $pwdHist },
                    [PSCustomObject]@{ "Kategorie" = "Kennwort-Richtlinie"; "Eigenschaft" = "Kennwortkomplexität"; "Wert" = if ($complexity) { "Aktiviert" } else { "Deaktiviert" } },
                    [PSCustomObject]@{ "Kategorie" = "Kennwort-Richtlinie"; "Eigenschaft" = "Umkehrbare Verschlüsselung"; "Wert" = if ($reversible) { "Aktiviert (Unsicher!)" } else { "Deaktiviert (Sicher)" } },
                    [PSCustomObject]@{ "Kategorie" = "Kennwort-Richtlinie"; "Eigenschaft" = "Maximales Kennwortalter (Ablauf)"; "Wert" = (& $convertTimeSpan $rawMaxAge $true) },
                    [PSCustomObject]@{ "Kategorie" = "Kennwort-Richtlinie"; "Eigenschaft" = "Minimales Kennwortalter"; "Wert" = (& $convertTimeSpan $rawMinAge $true) },
                    [PSCustomObject]@{ "Kategorie" = "Kontosperrung"; "Eigenschaft" = "Kontosperrungsschwelle"; "Wert" = if ([int]$rawLockThr -gt 0) { "$rawLockThr ungültige Versuche" } else { "Keine Sperre (0)" } },
                    [PSCustomObject]@{ "Kategorie" = "Kontosperrung"; "Eigenschaft" = "Dauer der Kontosperrung"; "Wert" = (& $convertTimeSpan $rawLockDur $false) },
                    [PSCustomObject]@{ "Kategorie" = "Kontosperrung"; "Eigenschaft" = "Sperrzähler-Zurücksetzung"; "Wert" = (& $convertTimeSpan $rawLockWin $false) }
                )

                $script:cachedDomainPolicies = $domList
                $arrD = [System.Collections.ArrayList]::new()
                foreach ($row in $domList) { [void]$arrD.Add($row) }
                $gridDomain.DataSource = $arrD

                # 2. PSOs
                $psoSearcher = New-Object System.DirectoryServices.DirectorySearcher($domEntry)
                $psoSearcher.Filter = "(objectClass=msDS-PasswordSettings)"
                $psoSearcher.SearchScope = [System.DirectoryServices.SearchScope]::Subtree
                $psoSearcher.PageSize = 200
                $psoResults = $psoSearcher.FindAll()

                if ($psoResults -and $psoResults.Count -gt 0) {
                    foreach ($res in $psoResults) {
                        $entry = $res.GetDirectoryEntry()
                        $pName = if ($entry.Properties["name"].Value) { $entry.Properties["name"].Value } else { "Unbenannt" }
                        $pPrec = [int]$entry.Properties["msDS-PasswordSettingsPrecedence"].Value
                        $pMinLen = "$($entry.Properties['msDS-MinimumPasswordLength'].Value) Zeichen"
                        $pHist = "$($entry.Properties['msDS-PasswordHistoryLength'].Value)"
                        $pComp = if ([bool]$entry.Properties["msDS-PasswordComplexityEnabled"].Value) { "Aktiviert" } else { "Deaktiviert" }
                        $pMaxA = & $convertTimeSpan $entry.Properties["msDS-MaxPasswordAge"].Value $true
                        $pMinA = & $convertTimeSpan $entry.Properties["msDS-MinPasswordAge"].Value $true
                        $pThr  = "$($entry.Properties['msDS-LockoutThreshold'].Value) Versuche"
                        $pDur  = & $convertTimeSpan $entry.Properties["msDS-LockoutDuration"].Value $false
                        $pDN   = $entry.Properties["distinguishedName"].Value

                        $applies = @()
                        $rawApplies = @()
                        if ($entry.Properties["msDS-PSOAppliesTo"]) {
                            foreach ($app in $entry.Properties["msDS-PSOAppliesTo"]) {
                                $rawApplies += $app.ToString()
                                $applies += ($app -split ',')[0] -replace '^CN=', ''
                            }
                        }
                        $appliesStr = if ($applies.Count -gt 0) { $applies -join ', ' } else { "Keine Zuweisung" }

                        $script:cachedPSOs += [PSCustomObject]@{
                            "PSO Name"         = $pName
                            "Priorität (Rang)" = $pPrec
                            "Min. Länge"       = $pMinLen
                            "Chronik"          = $pHist
                            "Komplexität"      = $pComp
                            "Max. Alter"       = $pMaxA
                            "Min. Alter"       = $pMinA
                            "Sperrschwelle"    = $pThr
                            "Sperrdauer"       = $pDur
                            "Zugewiesen an"    = $appliesStr
                        }

                        $script:cachedFullPSOs += [PSCustomObject]@{
                            Name       = $pName
                            DN         = $pDN
                            Precedence = $pPrec
                            MinLength  = $pMinLen
                            History    = $pHist
                            Complexity = $pComp
                            MaxAge     = $pMaxA
                            MinAge     = $pMinA
                            Threshold  = $pThr
                            Duration   = $pDur
                            AppliesTo  = $rawApplies
                        }
                    }
                    $arrPSO = [System.Collections.ArrayList]::new()
                    foreach ($p in ($script:cachedPSOs | Sort-Object "Priorität (Rang)")) { [void]$arrPSO.Add($p) }
                    $gridPSO.DataSource = $arrPSO
                }

                # 3. GPO-Audit
                $gpoAuditList = [System.Collections.Generic.List[PSCustomObject]]::new()
                try {
                    Import-Module GroupPolicy -ErrorAction Stop
                    $allGpos = Get-GPO -All -ErrorAction SilentlyContinue
                    $defaultDomainPolicyGuid = "{31B2F340-016D-11D2-945F-00C04FB984F9}"

                    foreach ($g in $allGpos) {
                        try { [xml]$xml = Get-GPOReport -Guid $g.Id -ReportType Xml -ErrorAction Stop } catch { continue }

                        $secExt = $xml.GPO.Computer.ExtensionData.Extension | Where-Object { 
                            $_.type -match "SecuritySettings" -or $_.LocalName -eq "SecuritySettings" 
                        }

                        $settingsFound = @()
                        if ($secExt) {
                            $accountNodes = $secExt.SelectNodes(".//*[local-name()='Account']/*")
                            if ($accountNodes) {
                                foreach ($node in $accountNodes) {
                                    $pName = if ($node.SelectSingleNode("./*[local-name()='Display']/*[local-name()='Name']")) { 
                                        $node.SelectSingleNode("./*[local-name()='Display']/*[local-name()='Name']").InnerText.Trim() 
                                    } elseif ($node.Name) { $node.Name } else { $node.LocalName }

                                    $pVal = if ($node.SelectSingleNode("./*[local-name()='Display']/*[local-name()='DisplayString']")) { 
                                        $node.SelectSingleNode("./*[local-name()='Display']/*[local-name()='DisplayString']").InnerText.Trim() 
                                    } elseif ($node.SettingNumber) { $node.SettingNumber } else { $node.InnerText.Trim() }

                                    if (-not [string]::IsNullOrWhiteSpace($pVal)) { $settingsFound += "$pName = $pVal" }
                                }
                            }
                        }

                        $currentGuidStr = ("{" + $g.Id.ToString().Trim('{}') + "}").ToUpper()
                        $isRealDefaultPolicyByGuid = ($currentGuidStr -eq $defaultDomainPolicyGuid)
                        $hasDefaultPolicyName     = ($g.DisplayName -eq "Default Domain Policy")

                        if ($settingsFound.Count -eq 0) {
                            if (-not $isRealDefaultPolicyByGuid -and $g.DisplayName -notlike "*Passwort*" -and $g.DisplayName -notlike "*Password*") {
                                continue
                            }
                        }

                        if ($isRealDefaultPolicyByGuid) {
                            $script:cachedDefaultPolicyGpoSettings = $settingsFound
                        }

                        $links = $xml.GPO.LinksTo
                        $compEnabled = ($xml.GPO.Computer.Enabled -eq "true")

                        $valString = if ($settingsFound.Count -gt 0) { $settingsFound -join " | " } else { "Keine Kennworteinstellungen im XML definiert (Nicht konfiguriert)" }

                        if (-not $links -or $links.Count -eq 0) {
                            $gpoAuditList.Add([PSCustomObject]@{
                                "GPO Name"            = $g.DisplayName
                                "Status"              = "Wirkungslos"
                                "Problem / Diagnose"  = "GPO ist nirgendwo verlinkt (Unlinked GPO)"
                                "Verknüpft an (SOM)"  = "Keine Verlinkung"
                                "GPO GUID"            = $currentGuidStr
                                "Computer-Knoten"     = if ($compEnabled) { "Aktiviert" } else { "Deaktiviert" }
                                "Konfigurierte Werte" = $valString
                            })
                        } else {
                            foreach ($link in $links) {
                                $somPath     = $link.SOMPath
                                $linkEnabled = ($link.Enabled -eq "true")
                                $isDomainRoot = ($somPath -eq $domDN -or $somPath -like "DC=*")
                                $isOU         = ($somPath -match "^OU=")

                                if (-not $compEnabled) {
                                    $gpoAuditList.Add([PSCustomObject]@{
                                        "GPO Name"            = $g.DisplayName
                                        "Status"              = "Wirkungslos"
                                        "Problem / Diagnose"  = "Computerknoten der GPO ist DEAKTIVIERT"
                                        "Verknüpft an (SOM)"  = $somPath
                                        "GPO GUID"            = $currentGuidStr
                                        "Computer-Knoten"     = "Deaktiviert"
                                        "Konfigurierte Werte" = $valString
                                    })
                                } elseif (-not $linkEnabled) {
                                    $gpoAuditList.Add([PSCustomObject]@{
                                        "GPO Name"            = $g.DisplayName
                                        "Status"              = "Wirkungslos"
                                        "Problem / Diagnose"  = "Verknüpfung (Link) ist DEAKTIVIERT"
                                        "Verknüpft an (SOM)"  = $somPath
                                        "GPO GUID"            = $currentGuidStr
                                        "Computer-Knoten"     = "Aktiviert"
                                        "Konfigurierte Werte" = $valString
                                    })
                                } elseif ($isOU) {
                                    $ouName = ($somPath -split ',')[0] -replace '^OU=', ''
                                    $gpoAuditList.Add([PSCustomObject]@{
                                        "GPO Name"            = $g.DisplayName
                                        "Status"              = "Fehlkonfiguriert"
                                        "Problem / Diagnose"  = "An OU verlinkt! Ignoriert Domänenbenutzer (wirkt nur auf lokale PC-SAM)."
                                        "Verknüpft an (SOM)"  = "OU: $ouName"
                                        "GPO GUID"            = $currentGuidStr
                                        "Computer-Knoten"     = "Aktiviert"
                                        "Konfigurierte Werte" = $valString
                                    })
                                } elseif ($isDomainRoot) {
                                    $statusText = ""
                                    $diagText   = ""
                                    if ($isRealDefaultPolicyByGuid -and $hasDefaultPolicyName) {
                                        $statusText = "Aktiv (Original Default Domain Policy)"
                                        $diagText   = "Echte Original-GPO an Domänenwurzel gebunden (GUID bestätigt, gilt primär)"
                                    } elseif ($isRealDefaultPolicyByGuid -and -not $hasDefaultPolicyName) {
                                        $statusText = "Umbenannte Default Domain Policy"
                                        $diagText   = "ACHTUNG: Dies ist die ECHTE Default Domain Policy ({31B2F340...}), wurde aber in '$($g.DisplayName)' umbenannt!"
                                    } elseif (-not $isRealDefaultPolicyByGuid -and $hasDefaultPolicyName) {
                                        $statusText = "Kein Original (Nur Namensgleichheit!)"
                                        $diagText   = "WARNUNG: Heißt 'Default Domain Policy', hat aber eine andere GUID! Gilt nur bei höherer Link-Order."
                                    } else {
                                        $statusText = "Aktiv (Zusätzliche Domänen-GPO)"
                                        $diagText   = "An Domänenwurzel gebunden (Überschreibt Default Policy, falls Link-Order höher)"
                                    }

                                    $gpoAuditList.Add([PSCustomObject]@{
                                        "GPO Name"            = $g.DisplayName
                                        "Status"              = $statusText
                                        "Problem / Diagnose"  = $diagText
                                        "Verknüpft an (SOM)"  = "Domänen-Root ($domainName)"
                                        "GPO GUID"            = $currentGuidStr
                                        "Computer-Knoten"     = "Aktiviert"
                                        "Konfigurierte Werte" = $valString
                                    })
                                }
                            }
                        }
                    }
                } catch {}

                $script:cachedGPOAudits = $gpoAuditList
                $arrGPO = [System.Collections.ArrayList]::new()
                foreach ($item in ($gpoAuditList | Sort-Object "Status" -Descending)) { [void]$arrGPO.Add($item) }
                $gridGPOAudit.DataSource = $arrGPO

                $script:ActiveLoadedSnapshotPath = "Live-Scan aus AD ($((Get-Date).ToString('yyyy-MM-dd HH:mm:ss')))"
                $lblSnapshotInfo.Text = "Datenquelle: $script:ActiveLoadedSnapshotPath"

                $origGpo = $script:cachedGPOAudits | Where-Object { $_."GPO GUID" -eq "{31B2F340-016D-11D2-945F-00C04FB984F9}" } | Select-Object -First 1
                if ($origGpo -and $origGpo."GPO Name" -ne "Default Domain Policy") {
                    $pnlWarningBanner.Visible = $true
                    $lblWarningBanner.Text = "ACHTUNG: Die 'Default Domain Policy' wurde umbenannt in '$($origGpo.'GPO Name')'! (GUID: $($origGpo.'GPO GUID'))"
                } else {
                    $pnlWarningBanner.Visible = $false
                }

                & $updateDomainSourceBox
                & $buildExpertTopics
                & $saveSnapshotToFile $script:DefaultCacheFile
                $lblStatus.Text = "Live-Abfrage erfolgreich abgeschlossen."

            } catch {
                $lblStatus.Text = "Fehler: $($_.Exception.Message)"
            } finally {
                $progressBar.Visible = $false
                $form.Cursor = [System.Windows.Forms.Cursors]::Default
            }
        }

        # BENUTZER-CHECK LOGIK
        $checkUserAction = {
            $uName = $txtUserCheck.Text.Trim()
            if (-not $uName) { return }

            $lblStatus.Text = "Prüfe Richtlinie für '$uName'..."
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

            try {
                $userPSO = $null
                $psoFoundVia = ""

                if ($isOnline) {
                    $searcher = New-Object System.DirectoryServices.DirectorySearcher([ADSI]"LDAP://$domDN")
                    $searcher.Filter = "(&(objectCategory=person)(objectClass=user)(|(sAMAccountName=$uName)(userPrincipalName=$uName)))"
                    $searcher.PropertiesToLoad.AddRange(@("distinguishedName", "msDS-ResultantPSO", "sAMAccountName", "memberOf"))
                    $uRes = $searcher.FindOne()

                    if ($uRes) {
                        $userEntry = $uRes.GetDirectoryEntry()
                        $foundSam  = $userEntry.Properties["sAMAccountName"].Value
                        $userDN    = $userEntry.Properties["distinguishedName"].Value
                        $script:lastCheckedUser = $foundSam

                        $userPSO = $userEntry.Properties["msDS-ResultantPSO"].Value
                        $psoFoundVia = "Direkte Attribut-Auswertung (msDS-ResultantPSO)"

                        if (-not $userPSO) {
                            $userGroupsDN = [System.Collections.Generic.List[string]]::new()
                            if ($userEntry.Properties["memberOf"]) {
                                foreach ($m in $userEntry.Properties["memberOf"]) {
                                    $userGroupsDN.Add($m.ToString().Trim())
                                }
                            }

                            $matchingPSOs = @()
                            foreach ($p in $script:cachedFullPSOs) {
                                $appliesTo = @($p.AppliesTo)
                                $isMatch = $false
                                $matchedGroup = ""

                                if ($appliesTo -contains $userDN) {
                                    $isMatch = $true
                                    $matchedGroup = "Direkte Zuweisung an Benutzer"
                                } else {
                                    foreach ($grp in $userGroupsDN) {
                                        if ($appliesTo -contains $grp) {
                                            $isMatch = $true
                                            $matchedGroup = ($grp -split ',')[0] -replace '^CN=', ''
                                            break
                                        }
                                    }
                                }

                                if ($isMatch) {
                                    $matchingPSOs += [PSCustomObject]@{
                                        DN         = $p.DN
                                        Precedence = [int]$p.Precedence
                                        MatchVia   = $matchedGroup
                                    }
                                }
                            }

                            if ($matchingPSOs.Count -gt 0) {
                                $bestMatch = $matchingPSOs | Sort-Object Precedence | Select-Object -First 1
                                $userPSO = $bestMatch.DN
                                $psoFoundVia = "Gruppenzuweisung über: '$($bestMatch.MatchVia)' (Rang $($bestMatch.Precedence))"
                            }
                        }
                    }
                }

                $effList = @()
                if ($userPSO) {
                    $psoObj = $script:cachedFullPSOs | Where-Object { $_.DN -eq $userPSO } | Select-Object -First 1
                    if ($psoObj) {
                        $effList = @(
                            [PSCustomObject]@{ "Kategorie" = "Angewendete Richtlinie"; "Eigenschaft" = "Richtlinien-Typ"; "Wert" = "Fine-Grained Password Policy (PSO)" },
                            [PSCustomObject]@{ "Kategorie" = "Angewendete Richtlinie"; "Eigenschaft" = "PSO Name"; "Wert" = $psoObj.Name },
                            [PSCustomObject]@{ "Kategorie" = "Angewendete Richtlinie"; "Eigenschaft" = "Ermittlungsweg"; "Wert" = $psoFoundVia },
                            [PSCustomObject]@{ "Kategorie" = "Angewendete Richtlinie"; "Eigenschaft" = "Distinguished Name"; "Wert" = $userPSO },
                            [PSCustomObject]@{ "Kategorie" = "Kennwort-Richtlinie"; "Eigenschaft" = "Minimale Kennwortlänge"; "Wert" = $psoObj.MinLength },
                            [PSCustomObject]@{ "Kategorie" = "Kennwort-Richtlinie"; "Eigenschaft" = "Kennworthistorie"; "Wert" = "$($psoObj.History) Passwörter" },
                            [PSCustomObject]@{ "Kategorie" = "Kennwort-Richtlinie"; "Eigenschaft" = "Kennwortkomplexität"; "Wert" = $psoObj.Complexity },
                            [PSCustomObject]@{ "Kategorie" = "Kennwort-Richtlinie"; "Eigenschaft" = "Maximales Kennwortalter"; "Wert" = $psoObj.MaxAge },
                            [PSCustomObject]@{ "Kategorie" = "Kennwort-Richtlinie"; "Eigenschaft" = "Minimales Kennwortalter"; "Wert" = $psoObj.MinAge },
                            [PSCustomObject]@{ "Kategorie" = "Kontosperrung"; "Eigenschaft" = "Kontosperrschwelle"; "Wert" = $psoObj.Threshold },
                            [PSCustomObject]@{ "Kategorie" = "Kontosperrung"; "Eigenschaft" = "Kontosperrdauer"; "Wert" = $psoObj.Duration }
                        )
                    }
                }

                if ($effList.Count -eq 0) {
                    $effList += [PSCustomObject]@{ "Kategorie" = "Angewendete Richtlinie"; "Eigenschaft" = "Richtlinien-Typ"; "Wert" = "Standard Domänen-Kennwortrichtlinie (Default Domain Policy)" }
                    $effList += [PSCustomObject]@{ "Kategorie" = "Angewendete Richtlinie"; "Eigenschaft" = "Ermittlungsweg"; "Wert" = "Weder direkt noch über Gruppen ein PSO zugewiesen" }
                    foreach ($row in $script:cachedDomainPolicies) { $effList += $row }
                }

                $script:cachedUserPolicy = $effList
                $arrU = [System.Collections.ArrayList]::new()
                foreach ($el in $effList) { [void]$arrU.Add($el) }
                $gridUserPolicy.DataSource = $arrU
                $lblStatus.Text = "Effektive Richtlinie für '$uName' ermittelt."
            } catch {
                $lblStatus.Text = "Fehler: $($_.Exception.Message)"
            } finally {
                $form.Cursor = [System.Windows.Forms.Cursors]::Default
            }
        }

        # GEZIELTE GPO-PRÜFUNG
        $applyGpoFilterAction = {
            $f = $txtGpoFilter.Text.Trim()
            
            if ($isOnline -and $f) {
                try {
                    $liveGpo = Get-GPO -Name $f -ErrorAction SilentlyContinue
                    if (-not $liveGpo) {
                        $liveGpo = Get-GPO -All | Where-Object { $_.DisplayName -like "*$f*" } | Select-Object -First 1
                    }

                    if ($liveGpo) {
                        $existsInCache = $script:cachedGPOAudits | Where-Object { $_."GPO GUID" -eq ("{" + $liveGpo.Id.ToString().Trim('{}') + "}").ToUpper() }
                        if (-not $existsInCache) {
                            [xml]$xml = Get-GPOReport -Guid $liveGpo.Id -ReportType Xml -ErrorAction SilentlyContinue
                            $secExt = $xml.GPO.Computer.ExtensionData.Extension | Where-Object { $_.type -match "SecuritySettings" -or $_.LocalName -eq "SecuritySettings" }
                            $settingsFound = @()
                            if ($secExt) {
                                foreach ($node in $secExt.SelectNodes(".//*[local-name()='Account']/*")) {
                                    $pName = if ($node.SelectSingleNode("./*[local-name()='Display']/*[local-name()='Name']")) { $node.SelectSingleNode("./*[local-name()='Display']/*[local-name()='Name']").InnerText.Trim() } elseif ($node.Name) { $node.Name } else { $node.LocalName }
                                    $pVal = if ($node.SelectSingleNode("./*[local-name()='Display']/*[local-name()='DisplayString']")) { $node.SelectSingleNode("./*[local-name()='Display']/*[local-name()='DisplayString']").InnerText.Trim() } elseif ($node.SettingNumber) { $node.SettingNumber } else { $node.InnerText.Trim() }
                                    if (-not [string]::IsNullOrWhiteSpace($pVal)) { $settingsFound += "$pName = $pVal" }
                                }
                            }
                            $links = $xml.GPO.LinksTo
                            $compEnabled = ($xml.GPO.Computer.Enabled -eq "true")
                            $valString = if ($settingsFound.Count -gt 0) { $settingsFound -join " | " } else { "Keine Kennworteinstellungen definiert (Nicht konfiguriert)" }
                            $currentGuidStr = ("{" + $liveGpo.Id.ToString().Trim('{}') + "}").ToUpper()

                            $linkPath = if ($links) { ($links | ForEach-Object { $_.SOMPath }) -join ", " } else { "Keine Verlinkung (Unlinked)" }
                            $status = if (-not $links) { "Wirkungslos" } elseif ($linkPath -match "^OU=") { "Fehlkonfiguriert" } else { "Aktiv" }
                            $diag = if (-not $links) { "GPO ist nirgendwo verlinkt" } elseif ($linkPath -match "^OU=") { "An OU verlinkt! Ignoriert Domänenbenutzer." } else { "An Domänenwurzel gebunden" }

                            $script:cachedGPOAudits += [PSCustomObject]@{
                                "GPO Name"            = $liveGpo.DisplayName
                                "Status"              = $status
                                "Problem / Diagnose"  = $diag
                                "Verknüpft an (SOM)"  = $linkPath
                                "GPO GUID"            = $currentGuidStr
                                "Computer-Knoten"     = if ($compEnabled) { "Aktiviert" } else { "Deaktiviert" }
                                "Konfigurierte Werte" = $valString
                            }
                        }
                    }
                } catch {}
            }

            $filtered = if ($f) {
                $script:cachedGPOAudits | Where-Object { 
                    $_."GPO Name" -like "*$f*" -or 
                    $_."Status" -like "*Default Domain Policy*" -or
                    $_."GPO GUID" -eq "{31B2F340-016D-11D2-945F-00C04FB984F9}"
                }
            } else {
                $script:cachedGPOAudits
            }

            $arrGPO = [System.Collections.ArrayList]::new()
            foreach ($item in ($filtered | Sort-Object "Status" -Descending)) { [void]$arrGPO.Add($item) }
            $gridGPOAudit.DataSource = $arrGPO
            $lblStatus.Text = "$($arrGPO.Count) GPO(s) angezeigt (Filter: '$f')."

            if ($f) {
                $specificMatch = $script:cachedGPOAudits | Where-Object { $_."GPO Name" -eq $f -or $_."GPO Name" -like "*$f*" } | Select-Object -First 1
                if ($specificMatch) {
                    $reasonMsg = @"
AUDIT-ERGEBNIS FÜR: '$($specificMatch.'GPO Name')'
GUID: $($specificMatch.'GPO GUID')
STATUS: $($specificMatch.'Status')
VERLINKT AN: $($specificMatch.'Verknüpft an (SOM)')

KONFIGURIERTE EINSTELLUNGEN:
$($specificMatch.'Konfigurierte Werte')

DIAGNOSE & AUSWIRKUNG:
$($specificMatch.'Problem / Diagnose')
"@
                    [System.Windows.Forms.MessageBox]::Show($reasonMsg, "GPO Audit Diagnose", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                } else {
                    [System.Windows.Forms.MessageBox]::Show("Die GPO '$f' wurde im AD-System nicht gefunden.", "Nicht gefunden", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                }
            }
        }

        $btnFilterGpo.Add_Click($applyGpoFilterAction)
        $txtGpoFilter.Add_KeyDown({
            if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
                $_.SuppressKeyPress = $true
                & $applyGpoFilterAction
            }
        })
        $btnResetGpo.Add_Click({
            $txtGpoFilter.Text = ""
            $arrGPO = [System.Collections.ArrayList]::new()
            foreach ($item in ($script:cachedGPOAudits | Sort-Object "Status" -Descending)) { [void]$arrGPO.Add($item) }
            $gridGPOAudit.DataSource = $arrGPO
            $lblStatus.Text = "Alle $($arrGPO.Count) GPOs angezeigt."
        })

        # EXPORT & SNAPSHOT BUTTONS
        $btnSaveSnapshot.Add_Click({
            $sfd = New-Object System.Windows.Forms.SaveFileDialog
            $sfd.Filter = "JSON Snapshot-Datei (*.json)|*.json"
            $sfd.FileName = "AD_PasswordPolicy_Snapshot_${domainName}_$((Get-Date).ToString('yyyyMMdd_HHmm')).json"
            if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                if (& $saveSnapshotToFile $sfd.FileName) {
                    [System.Windows.Forms.MessageBox]::Show("Snapshot gespeichert:`n$($sfd.FileName)", "Gespeichert", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                }
            }
        })

        $btnLoadSnapshot.Add_Click({
            $ofd = New-Object System.Windows.Forms.OpenFileDialog
            $ofd.Filter = "JSON Snapshot-Datei (*.json)|*.json"
            if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                & $loadSnapshotFromFile $ofd.FileName
            }
        })

        $btnRefresh.Add_Click({ & $loadAllPoliciesLive })
        $btnCheckUser.Add_Click($checkUserAction)
        $txtUserCheck.Add_KeyDown({
            if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
                $_.SuppressKeyPress = $true
                & $checkUserAction
            }
        })

        $btnExport.Add_Click({
            $timestamp = (Get-Date).ToString("yyyyMMdd_HHmm")
            $sfd = New-Object System.Windows.Forms.SaveFileDialog
            $sfd.Filter = "CSV-Datei (*.csv)|*.csv"

            switch ($tabControl.SelectedIndex) {
                0 { $exportData = $script:cachedDomainPolicies; $sfd.FileName = "AD_Kennwortrichtlinie_DomainDefault_${domainName}_${timestamp}.csv" }
                1 { $exportData = $script:cachedPSOs; $sfd.FileName = "AD_Kennwortrichtlinie_PSO_Uebersicht_${domainName}_${timestamp}.csv" }
                2 { $exportData = $script:cachedUserPolicy; $uClean = ($script:lastCheckedUser -replace '[\\/:*?"<>|]', '_'); $sfd.FileName = "AD_Kennwortrichtlinie_Effektiv_${uClean}_${timestamp}.csv" }
                3 { $exportData = $gridGPOAudit.DataSource; $sfd.FileName = "AD_Kennwortrichtlinie_GPO_Audit_${domainName}_${timestamp}.csv" }
                4 { $exportData = $script:cachedExpertTopics; $sfd.FileName = "AD_Kennwortrichtlinie_ExpertInfo_${domainName}_${timestamp}.csv" }
            }

            if ($null -eq $exportData -or $exportData.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("Keine Daten zum Exportieren vorhanden.", "Export", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                return
            }

            if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                try {
                    $exportData | Export-Csv -Path $sfd.FileName -NoTypeInformation -Delimiter ";" -Encoding UTF8
                    [System.Windows.Forms.MessageBox]::Show("Export gespeichert:`n$($sfd.FileName)", "Export", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                } catch {
                    [System.Windows.Forms.MessageBox]::Show("Fehler beim Speichern: $($_.Exception.Message)", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                }
            }
        })

        # INITIALISIERUNG
        $form.Add_Shown({
            if (Test-Path $script:DefaultCacheFile) {
                & $loadSnapshotFromFile $script:DefaultCacheFile
            } elseif ($isOnline) {
                & $loadAllPoliciesLive
            } else {
                & $buildExpertTopics
            }
            & $checkUserAction
        })

        [void]$form.ShowDialog()

    } finally {
        if ($null -ne $form) { $form.Dispose() }
        $script:cachedDomainPolicies            = @()
        $script:cachedPSOs                     = @()
        $script:cachedUserPolicy               = @()
        $script:cachedGPOAudits                = @()
        $script:cachedFullPSOs                 = @()
        $script:cachedDefaultPolicyGpoSettings = @()
        $script:cachedExpertTopics             = @()

        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

Start-PasswordPolicyAuditor
