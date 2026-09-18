<#
================================================================================
 TOOL 25: SERVER WINRM & WINDOWS DIENSTE AUDIT (OPTIMIERTES BUTTON-LAYOUT)
 - "⚡ Dienste abfragen"-Button direkt bei "1. Server WinRM Status" platziert
 - "Auswahl"-Spalte sauber lesbar (65px)
 - Sauberes FlowLayoutPanel gegen Verschieben des Server-Filterfeldes
 - Robuste SplitterDistance-Initialisierung (keine Panel2MinSize-Exceptions)
 - Vollständiger ISE Anti-Freeze Standard (Out-of-Process & Session-Cleanup)
================================================================================
#>

# ------------------------------------------------------------------------------
# 1. ISE OUT-OF-PROCESS SCHUTZ & SYSTEM-STYLES
# ------------------------------------------------------------------------------
if ($host.Name -eq 'Windows PowerShell ISE Host' -and -not [string]::IsNullOrEmpty($PSCommandPath)) {
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    return
}

if (-not [System.Windows.Forms.Application]::RenderWithVisualStyles) {
    try { [System.Windows.Forms.Application]::EnableVisualStyles() } catch {}
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.DirectoryServices

# ==============================================================================
# 2. ZENTRALES THEME
# ==============================================================================
$script:UITheme = @{
    FontFamily         = "Segoe UI"
    HeaderHeight       = 34
    RowHeight          = 26
    HeaderFontSize     = 9.0
    CellFontSize       = 9.0
    HeaderBackColor    = [System.Drawing.Color]::FromArgb(235, 241, 250)
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
    $hdrStyle.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, $script:UITheme.HeaderFontSize, [System.Drawing.FontStyle]::Bold)
    $hdrStyle.BackColor = $script:UITheme.HeaderBackColor
    $hdrStyle.ForeColor = $script:UITheme.HeaderForeColor
    $hdrStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleLeft
    $hdrStyle.Padding = New-Object System.Windows.Forms.Padding(6, 0, 6, 0)
    $grid.ColumnHeadersDefaultCellStyle = $hdrStyle

    $cellStyle = New-Object System.Windows.Forms.DataGridViewCellStyle
    $cellStyle.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, $script:UITheme.CellFontSize, [System.Drawing.FontStyle]::Regular)
    $cellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleLeft
    $cellStyle.Padding = New-Object System.Windows.Forms.Padding(6, 1, 6, 1)
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
    $grid.RowHeadersVisible = $false
    $grid.AllowUserToResizeRows = $false
}

function Show-Tool25-ServiceAudit {
    $script:rawServerList = [System.Collections.Generic.List[PSCustomObject]]::new()
    $script:allServices   = [System.Collections.Generic.List[PSCustomObject]]::new()
    $script:sortCol = ""
    $script:sortAsc = $true

    # --------------------------------------------------------------------------
    # 3. HAUPTFENSTER
    # --------------------------------------------------------------------------
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Tool 25: Server WinRM & Windows Dienste Audit (Umstellungshilfe)"
    $form.Size = New-Object System.Drawing.Size(1480, 860)
    $form.MinimumSize = New-Object System.Drawing.Size(1050, 640)
    $form.StartPosition = "CenterScreen"
    $form.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9)
    $form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

    # Obere Steuerleiste
    $pnlTop = New-Object System.Windows.Forms.Panel
    $pnlTop.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlTop.AutoSize = $true
    $pnlTop.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $pnlTop.BackColor = [System.Drawing.Color]::White
    $pnlTop.Padding = New-Object System.Windows.Forms.Padding(10, 8, 10, 8)
    $form.Controls.Add($pnlTop)

    $tblTop = New-Object System.Windows.Forms.TableLayoutPanel
    $tblTop.Dock = [System.Windows.Forms.DockStyle]::Top
    $tblTop.AutoSize = $true
    $tblTop.ColumnCount = 1
    $tblTop.RowCount = 2
    $tblTop.Margin = New-Object System.Windows.Forms.Padding(0)
    $pnlTop.Controls.Add($tblTop)

    # --- ZEILE 1: Server-Filter & Aktionsbuttons (Dienste-Button nach links versetzt) ---
    $flwRow1 = New-Object System.Windows.Forms.FlowLayoutPanel
    $flwRow1.Dock = [System.Windows.Forms.DockStyle]::Top
    $flwRow1.AutoSize = $true
    $flwRow1.WrapContents = $false
    $flwRow1.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 6)
    $tblTop.Controls.Add($flwRow1, 0, 0)

    $lblTarget = New-Object System.Windows.Forms.Label
    $lblTarget.Text = "Server / Filter:"
    $lblTarget.AutoSize = $true
    $lblTarget.Margin = New-Object System.Windows.Forms.Padding(0, 6, 6, 0)
    $lblTarget.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9, [System.Drawing.FontStyle]::Bold)
    $flwRow1.Controls.Add($lblTarget)

    $txtTarget = New-Object System.Windows.Forms.TextBox
    $txtTarget.Text = "*"
    $txtTarget.Size = New-Object System.Drawing.Size(130, 23)
    $txtTarget.Margin = New-Object System.Windows.Forms.Padding(0, 3, 10, 0)
    $flwRow1.Controls.Add($txtTarget)

    $btnLoadAD = New-Object System.Windows.Forms.Button
    $btnLoadAD.Text = "AD-Server suchen"
    $btnLoadAD.AutoSize = $true
    $btnLoadAD.Height = 26
    $btnLoadAD.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
    $btnLoadAD.ForeColor = [System.Drawing.Color]::White
    $btnLoadAD.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnLoadAD.FlatAppearance.BorderSize = 0
    $btnLoadAD.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 8.5, [System.Drawing.FontStyle]::Bold)
    $btnLoadAD.Margin = New-Object System.Windows.Forms.Padding(0, 2, 8, 0)
    $flwRow1.Controls.Add($btnLoadAD)

    $btnExport = New-Object System.Windows.Forms.Button
    $btnExport.Text = "CSV Export (;)"
    $btnExport.AutoSize = $true
    $btnExport.Height = 26
    $btnExport.BackColor = [System.Drawing.Color]::FromArgb(71, 85, 105)
    $btnExport.ForeColor = [System.Drawing.Color]::White
    $btnExport.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnExport.FlatAppearance.BorderSize = 0
    $btnExport.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 8.5, [System.Drawing.FontStyle]::Bold)
    $btnExport.Margin = New-Object System.Windows.Forms.Padding(0, 2, 14, 0)
    $flwRow1.Controls.Add($btnExport)

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Text = "Bereit. Server laden oder Suchmuster eingeben."
    $lblStatus.AutoSize = $true
    $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(30, 58, 138)
    $lblStatus.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9, [System.Drawing.FontStyle]::Bold)
    $lblStatus.Margin = New-Object System.Windows.Forms.Padding(0, 6, 0, 0)
    $flwRow1.Controls.Add($lblStatus)

    # --- ZEILE 2: Filterleiste für Dienste ---
    $flwRow2 = New-Object System.Windows.Forms.FlowLayoutPanel
    $flwRow2.Dock = [System.Windows.Forms.DockStyle]::Top
    $flwRow2.AutoSize = $true
    $flwRow2.WrapContents = $false
    $flwRow2.Margin = New-Object System.Windows.Forms.Padding(0)
    $tblTop.Controls.Add($flwRow2, 0, 1)

    $lblCat = New-Object System.Windows.Forms.Label
    $lblCat.Text = "Kategorie (Konto):"
    $lblCat.AutoSize = $true
    $lblCat.Margin = New-Object System.Windows.Forms.Padding(0, 6, 6, 0)
    $lblCat.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9, [System.Drawing.FontStyle]::Bold)
    $flwRow2.Controls.Add($lblCat)

    $cmbAccountCat = New-Object System.Windows.Forms.ComboBox
    $cmbAccountCat.Size = New-Object System.Drawing.Size(185, 23)
    $cmbAccountCat.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $cmbAccountCat.Margin = New-Object System.Windows.Forms.Padding(0, 2, 16, 0)
    [void]$cmbAccountCat.Items.AddRange(@(
        "Alle Konten",
        "Nur Benutzer-Konten (User)",
        "Nur Lokales System (SYSTEM)",
        "Nur Lokaler Dienst (LocalService)",
        "Nur Netzwerkdienst (NetworkService)",
        "Virtuelle Dienstkonten (NT SERVICE)"
    ))
    $cmbAccountCat.SelectedIndex = 0
    $flwRow2.Controls.Add($cmbAccountCat)

    $lblStartMode = New-Object System.Windows.Forms.Label
    $lblStartMode.Text = "Starttyp:"
    $lblStartMode.AutoSize = $true
    $lblStartMode.Margin = New-Object System.Windows.Forms.Padding(0, 6, 6, 0)
    $lblStartMode.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9, [System.Drawing.FontStyle]::Bold)
    $flwRow2.Controls.Add($lblStartMode)

    $cmbStartMode = New-Object System.Windows.Forms.ComboBox
    $cmbStartMode.Size = New-Object System.Drawing.Size(90, 23)
    $cmbStartMode.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $cmbStartMode.Margin = New-Object System.Windows.Forms.Padding(0, 2, 16, 0)
    [void]$cmbStartMode.Items.AddRange(@("Alle", "Auto", "Manual", "Disabled"))
    $cmbStartMode.SelectedIndex = 0
    $flwRow2.Controls.Add($cmbStartMode)

    $lblState = New-Object System.Windows.Forms.Label
    $lblState.Text = "Status:"
    $lblState.AutoSize = $true
    $lblState.Margin = New-Object System.Windows.Forms.Padding(0, 6, 6, 0)
    $lblState.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9, [System.Drawing.FontStyle]::Bold)
    $flwRow2.Controls.Add($lblState)

    $cmbState = New-Object System.Windows.Forms.ComboBox
    $cmbState.Size = New-Object System.Drawing.Size(85, 23)
    $cmbState.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $cmbState.Margin = New-Object System.Windows.Forms.Padding(0, 2, 16, 0)
    [void]$cmbState.Items.AddRange(@("Alle", "Running", "Stopped"))
    $cmbState.SelectedIndex = 0
    $flwRow2.Controls.Add($cmbState)

    $lblSearch = New-Object System.Windows.Forms.Label
    $lblSearch.Text = "Volltextsuche:"
    $lblSearch.AutoSize = $true
    $lblSearch.Margin = New-Object System.Windows.Forms.Padding(0, 6, 6, 0)
    $lblSearch.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9, [System.Drawing.FontStyle]::Bold)
    $flwRow2.Controls.Add($lblSearch)

    $txtFilter = New-Object System.Windows.Forms.TextBox
    $txtFilter.Size = New-Object System.Drawing.Size(200, 23)
    $txtFilter.Margin = New-Object System.Windows.Forms.Padding(0, 3, 0, 0)
    $flwRow2.Controls.Add($txtFilter)

    # Trennlinie
    $pnlLine = New-Object System.Windows.Forms.Panel
    $pnlLine.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlLine.Height = 1
    $pnlLine.BackColor = [System.Drawing.Color]::FromArgb(226, 232, 240)
    $form.Controls.Add($pnlLine)
    $pnlLine.BringToFront()

    # Splitter für Links und Rechts
    $split = New-Object System.Windows.Forms.SplitContainer
    $split.Dock = [System.Windows.Forms.DockStyle]::Fill
    $split.Orientation = [System.Windows.Forms.Orientation]::Vertical
    $form.Controls.Add($split)
    $split.BringToFront()

    # --------------------------------------------------------------------------
    # 4. BEREICH LINKS: SERVERLISTE & DIENSTE-BUTTON NEBENEINANDER
    # --------------------------------------------------------------------------
    $grpServers = New-Object System.Windows.Forms.GroupBox
    $grpServers.Text = "1. Server WinRM Status"
    $grpServers.Dock = [System.Windows.Forms.DockStyle]::Fill
    $grpServers.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9, [System.Drawing.FontStyle]::Bold)
    $split.Panel1.Controls.Add($grpServers)

    # Toolbar oberhalb des Server-Grids
    $pnlServerTop = New-Object System.Windows.Forms.Panel
    $pnlServerTop.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlServerTop.Height = 66
    $pnlServerTop.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 252)
    $grpServers.Controls.Add($pnlServerTop)

    $flwServerBar = New-Object System.Windows.Forms.FlowLayoutPanel
    $flwServerBar.Dock = [System.Windows.Forms.DockStyle]::Fill
    $flwServerBar.Padding = New-Object System.Windows.Forms.Padding(6, 4, 6, 2)
    $flwServerBar.WrapContents = $true
    $pnlServerTop.Controls.Add($flwServerBar)

    # 1. BUTTON: Dienste abfragen (NEUER PLATZ)
    $btnScanServices = New-Object System.Windows.Forms.Button
    $btnScanServices.Text = "⚡ Dienste abfragen"
    $btnScanServices.AutoSize = $true
    $btnScanServices.Height = 24
    $btnScanServices.BackColor = [System.Drawing.Color]::FromArgb(16, 124, 65)
    $btnScanServices.ForeColor = [System.Drawing.Color]::White
    $btnScanServices.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnScanServices.FlatAppearance.BorderSize = 0
    $btnScanServices.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 8.5, [System.Drawing.FontStyle]::Bold)
    $btnScanServices.Margin = New-Object System.Windows.Forms.Padding(0, 0, 6, 3)
    $flwServerBar.Controls.Add($btnScanServices)

    # 2. BUTTON: Alle
    $btnSelectAll = New-Object System.Windows.Forms.Button
    $btnSelectAll.Text = "✓ Alle"
    $btnSelectAll.Size = New-Object System.Drawing.Size(55, 24)
    $btnSelectAll.BackColor = [System.Drawing.Color]::FromArgb(235, 241, 250)
    $btnSelectAll.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnSelectAll.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(203, 213, 225)
    $btnSelectAll.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 8.5)
    $btnSelectAll.Margin = New-Object System.Windows.Forms.Padding(0, 0, 4, 3)
    $flwServerBar.Controls.Add($btnSelectAll)

    # 3. BUTTON: Kein
    $btnDeselectAll = New-Object System.Windows.Forms.Button
    $btnDeselectAll.Text = "✕ Kein"
    $btnDeselectAll.Size = New-Object System.Drawing.Size(55, 24)
    $btnDeselectAll.BackColor = [System.Drawing.Color]::FromArgb(235, 241, 250)
    $btnDeselectAll.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnDeselectAll.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(203, 213, 225)
    $btnDeselectAll.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 8.5)
    $btnDeselectAll.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 3)
    $flwServerBar.Controls.Add($btnDeselectAll)

    # Zeilenumbruch zur Filterbox
    $flwServerBar.SetFlowBreak($btnDeselectAll, $true)

    $lblSrvSearch = New-Object System.Windows.Forms.Label
    $lblSrvSearch.Text = "Server filtern:"
    $lblSrvSearch.AutoSize = $true
    $lblSrvSearch.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 8.5)
    $lblSrvSearch.Margin = New-Object System.Windows.Forms.Padding(0, 5, 6, 0)
    $flwServerBar.Controls.Add($lblSrvSearch)

    $txtServerFilter = New-Object System.Windows.Forms.TextBox
    $txtServerFilter.Size = New-Object System.Drawing.Size(185, 23)
    $txtServerFilter.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 8.5)
    $txtServerFilter.Margin = New-Object System.Windows.Forms.Padding(0, 2, 0, 0)
    $flwServerBar.Controls.Add($txtServerFilter)

    $gridServers = New-Object System.Windows.Forms.DataGridView
    $gridServers.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridServers.AllowUserToAddRows = $false
    $gridServers.AllowUserToDeleteRows = $false
    $gridServers.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $gridServers.EditMode = [System.Windows.Forms.DataGridViewEditMode]::EditOnEnter
    Apply-StandardGridTheme $gridServers
    $grpServers.Controls.Add($gridServers)
    $gridServers.BringToFront()

    # Spaltenkonfiguration: "Auswahl" voll lesbar mit 65 px
    $colChk = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $colChk.HeaderText = "Auswahl"
    $colChk.Name = "Select"
    $colChk.Width = 65
    $colChk.Resizable = [System.Windows.Forms.DataGridViewTriState]::False
    [void]$gridServers.Columns.Add($colChk)

    $colSrvName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colSrvName.HeaderText = "Servername"
    $colSrvName.Name = "ServerName"
    $colSrvName.Width = 155
    $colSrvName.ReadOnly = $true
    [void]$gridServers.Columns.Add($colSrvName)

    $colWinRm = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colWinRm.HeaderText = "WinRM"
    $colWinRm.Name = "WinRM"
    $colWinRm.Width = 95
    $colWinRm.ReadOnly = $true
    [void]$gridServers.Columns.Add($colWinRm)

    # Sofortiger Commit bei Checkbox-Klick
    $gridServers.Add_CurrentCellDirtyStateChanged({
        if ($gridServers.IsCurrentCellDirty) {
            $gridServers.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
        }
    })
    $gridServers.Add_CellContentClick({
        param($s, $e)
        if ($e.RowIndex -ge 0 -and $e.ColumnIndex -eq 0) {
            $gridServers.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
        }
    })

    # --------------------------------------------------------------------------
    # 5. BEREICH RECHTS: GEFILTERTE DIENSTE (MAXIMALE BREITE)
    # --------------------------------------------------------------------------
    $grpServices = New-Object System.Windows.Forms.GroupBox
    $grpServices.Text = "2. Gefundene Windows-Dienste (gefiltert)"
    $grpServices.Dock = [System.Windows.Forms.DockStyle]::Fill
    $grpServices.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9, [System.Drawing.FontStyle]::Bold)
    $split.Panel2.Controls.Add($grpServices)

    $gridServices = New-Object System.Windows.Forms.DataGridView
    $gridServices.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridServices.ReadOnly = $true
    $gridServices.AllowUserToAddRows = $false
    $gridServices.AllowUserToDeleteRows = $false
    $gridServices.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $gridServices.AutoGenerateColumns = $false
    Apply-StandardGridTheme $gridServices
    $grpServices.Controls.Add($gridServices)
    $gridServices.BringToFront()

    $cols = @(
        @{ Name="Server"; Header="Server"; Width=120 },
        @{ Name="ServiceName"; Header="Dienstname"; Width=135 },
        @{ Name="DisplayName"; Header="Anzeigename"; Width=210 },
        @{ Name="State"; Header="Status"; Width=80 },
        @{ Name="StartMode"; Header="Starttyp"; Width=80 },
        @{ Name="AccountCategory"; Header="Kategorie (Konto)"; Width=150 },
        @{ Name="StartName"; Header="Anmelden als (Konto)"; Width=190 },
        @{ Name="Description"; Header="Beschreibung"; Width=450 }
    )

    foreach ($c in $cols) {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = $c.Name
        $col.DataPropertyName = $c.Name
        $col.HeaderText = $c.Header
        $col.Width = $c.Width
        [void]$gridServices.Columns.Add($col)
    }

    # --------------------------------------------------------------------------
    # 6. ROW PAINT FARBEN
    # --------------------------------------------------------------------------
    $gridServers.Add_RowPrePaint({
        param($s, $e)
        if ($e.RowIndex -ge 0 -and $e.RowIndex -lt $gridServers.Rows.Count) {
            $row = $gridServers.Rows[$e.RowIndex]
            $winrm = [string]$row.Cells["WinRM"].Value
            if ($winrm -like "*OK*") {
                $row.Cells["WinRM"].Style.ForeColor = [System.Drawing.Color]::FromArgb(22, 101, 52)
                $row.Cells["WinRM"].Style.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 8.5, [System.Drawing.FontStyle]::Bold)
            } else {
                $row.Cells["WinRM"].Style.ForeColor = [System.Drawing.Color]::FromArgb(185, 28, 28)
                $row.Cells["WinRM"].Style.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 8.5, [System.Drawing.FontStyle]::Bold)
            }
        }
    })

    $gridServices.Add_RowPrePaint({
        param($s, $e)
        if ($e.RowIndex -ge 0 -and $e.RowIndex -lt $gridServices.Rows.Count) {
            $row = $gridServices.Rows[$e.RowIndex]
            $cat = [string]$row.Cells["AccountCategory"].Value
            $state = [string]$row.Cells["State"].Value

            if ($cat -eq "User / Domain Account") {
                $row.DefaultCellStyle.BackColor          = [System.Drawing.Color]::FromArgb(254, 243, 199)
                $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(253, 230, 138)
                $row.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::FromArgb(120, 53, 15)
                $row.Cells["AccountCategory"].Style.ForeColor = [System.Drawing.Color]::FromArgb(180, 83, 9)
                $row.Cells["AccountCategory"].Style.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 8.5, [System.Drawing.FontStyle]::Bold)
            }

            if ($state -eq "Error") {
                $row.DefaultCellStyle.BackColor          = [System.Drawing.Color]::FromArgb(254, 226, 226)
                $row.DefaultCellStyle.ForeColor          = [System.Drawing.Color]::FromArgb(153, 27, 27)
                $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(252, 165, 165)
                $row.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::FromArgb(127, 29, 29)
            }
        }
    })

    # --------------------------------------------------------------------------
    # 7. LOGIK: SERVER ERMITTELN & LIVE-FILTER
    # --------------------------------------------------------------------------
    $btnLoadAD.Add_Click({
        $script:rawServerList.Clear()
        $gridServers.Rows.Clear()
        $pattern = $txtTarget.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($pattern)) { $pattern = "*" }

        $lblStatus.Text = "Suche Server im AD..."
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        [System.Windows.Forms.Application]::DoEvents()

        $serverList = @()
        try {
            $rootDSE = [ADSI]"LDAP://RootDSE"
            $defNC = $rootDSE.defaultNamingContext.ToString()
            $searcher = [System.DirectoryServices.DirectorySearcher]::new([System.DirectoryServices.DirectoryEntry]"LDAP://$defNC")
            $searcher.Filter = "(&(objectCategory=computer)(operatingsystem=*Server*)(name=$pattern))"
            $searcher.PageSize = 500
            [void]$searcher.PropertiesToLoad.Add("dnshostname")
            [void]$searcher.PropertiesToLoad.Add("samaccountname")

            $results = $searcher.FindAll()
            foreach ($r in $results) {
                $h = if ($r.Properties["dnshostname"].Count -gt 0) { $r.Properties["dnshostname"][0] } else { $r.Properties["samaccountname"][0].ToString().TrimEnd('$') }
                $serverList += $h
            }
            $searcher.Dispose()
        } catch {
            if ($pattern -ne "*") {
                $serverList = $pattern -split "[,;\s]+" | Where-Object { $_ }
            } else {
                $serverList = @($env:COMPUTERNAME)
            }
        }

        $lblStatus.Text = "Prüfe WinRM auf $($serverList.Count) Servern..."
        [System.Windows.Forms.Application]::DoEvents()

        foreach ($srv in $serverList) {
            $winrmOk = $false
            try {
                $ws = Test-WSMan -ComputerName $srv -ErrorAction Stop
                if ($ws) { $winrmOk = $true }
            } catch {
                $winrmOk = $false
            }

            $statusText = if ($winrmOk) { "Online (OK)" } else { "Offline" }
            $script:rawServerList.Add([PSCustomObject]@{
                Select     = $winrmOk
                ServerName = $srv
                WinRM      = $statusText
            })
        }

        & $applyServerFilter
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $lblStatus.Text = "$($serverList.Count) Server ermittelt."
    })

    # Server Live-Filter
    $applyServerFilter = {
        $stxt = $txtServerFilter.Text.Trim()
        $gridServers.Rows.Clear()
        
        $matched = if ([string]::IsNullOrEmpty($stxt)) {
            $script:rawServerList
        } else {
            $script:rawServerList | Where-Object { $_.ServerName -like "*$stxt*" }
        }

        foreach ($srv in $matched) {
            [void]$gridServers.Rows.Add($srv.Select, $srv.ServerName, $srv.WinRM)
        }
        $gridServers.ClearSelection()
    }
    $txtServerFilter.Add_TextChanged($applyServerFilter)

    # Schnell-Buttons
    $btnSelectAll.Add_Click({
        foreach ($row in $gridServers.Rows) { 
            $row.Cells["Select"].Value = $true 
            $srvName = $row.Cells["ServerName"].Value
            $obj = $script:rawServerList | Where-Object { $_.ServerName -eq $srvName } | Select-Object -First 1
            if ($obj) { $obj.Select = $true }
        }
    })
    $btnDeselectAll.Add_Click({
        foreach ($row in $gridServers.Rows) { 
            $row.Cells["Select"].Value = $false 
            $srvName = $row.Cells["ServerName"].Value
            $obj = $script:rawServerList | Where-Object { $_.ServerName -eq $srvName } | Select-Object -First 1
            if ($obj) { $obj.Select = $false }
        }
    })

    # --------------------------------------------------------------------------
    # 8. LOGIK: DIENSTE SAMMELN & KATEGORISIEREN
    # --------------------------------------------------------------------------
    $btnScanServices.Add_Click({
        $selectedServers = @()
        foreach ($row in $gridServers.Rows) {
            if ($row.Cells["Select"].Value -eq $true) {
                $selectedServers += $row.Cells["ServerName"].Value
            }
        }

        if ($selectedServers.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Bitte mindestens einen Server in der linken Tabelle auswählen.", "Kein Server gewählt", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        $script:allServices.Clear()
        $lblStatus.Text = "Lese Dienste von $($selectedServers.Count) Server(n) ein..."
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        [System.Windows.Forms.Application]::DoEvents()

        foreach ($srv in $selectedServers) {
            $lblStatus.Text = "Frage $srv ab..."
            [System.Windows.Forms.Application]::DoEvents()

            try {
                $cimOpt = New-CimSessionOption -Protocol Wsman
                $sess = New-CimSession -ComputerName $srv -SessionOption $cimOpt -OperationTimeoutSec 15 -ErrorAction Stop
                $rawServices = Get-CimInstance -CimSession $sess -ClassName Win32_Service -Property Name, DisplayName, State, StartMode, StartName, Description -ErrorAction Stop
                Remove-CimSession $sess -ErrorAction SilentlyContinue

                foreach ($svc in $rawServices) {
                    $startName = if ($svc.StartName) { $svc.StartName } else { "N/A" }
                    
                    $cat = "User / Domain Account"
                    if ($startName -match "LocalSystem" -or $startName -eq ".\LocalSystem" -or $startName -eq "LocalSystem") {
                        $cat = "Lokales System"
                    } elseif ($startName -match "LocalService") {
                        $cat = "Lokaler Dienst"
                    } elseif ($startName -match "NetworkService") {
                        $cat = "Netzwerkdienst"
                    } elseif ($startName -match "^NT SERVICE\\") {
                        $cat = "Virtuelles Dienstkonto"
                    }

                    $script:allServices.Add([PSCustomObject]@{
                        Server          = $srv
                        ServiceName     = $svc.Name
                        DisplayName     = $svc.DisplayName
                        State           = $svc.State
                        StartMode       = $svc.StartMode
                        AccountCategory = $cat
                        StartName       = $startName
                        Description     = if ($svc.Description) { $svc.Description } else { "-" }
                    })
                }
            } catch {
                $script:allServices.Add([PSCustomObject]@{
                    Server          = $srv
                    ServiceName     = "FEHLER"
                    DisplayName     = "Dienste konnten nicht gelesen werden: $($_.Exception.Message)"
                    State           = "Error"
                    StartMode       = "-"
                    AccountCategory = "Fehler"
                    StartName       = "-"
                    Description     = "-"
                })
            }
        }

        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        & $applyFilter
    })

    # --------------------------------------------------------------------------
    # 9. LIVE-FILTER FÜR DIENSTE
    # --------------------------------------------------------------------------
    $applyFilter = {
        if ($null -eq $script:allServices -or $script:allServices.Count -eq 0) {
            $gridServices.DataSource = $null
            return
        }

        $selectedCat = $cmbAccountCat.SelectedItem
        $selectedMode = $cmbStartMode.SelectedItem
        $selectedState = $cmbState.SelectedItem
        $searchTxt = $txtFilter.Text.Trim()

        $filtered = $script:allServices | Where-Object {
            $item = $_

            if ($selectedCat -eq "Nur Benutzer-Konten (User)" -and $item.AccountCategory -ne "User / Domain Account") { return $false }
            if ($selectedCat -eq "Nur Lokales System (SYSTEM)" -and $item.AccountCategory -ne "Lokales System") { return $false }
            if ($selectedCat -eq "Nur Lokaler Dienst (LocalService)" -and $item.AccountCategory -ne "Lokaler Dienst") { return $false }
            if ($selectedCat -eq "Nur Netzwerkdienst (NetworkService)" -and $item.AccountCategory -ne "Netzwerkdienst") { return $false }
            if ($selectedCat -eq "Virtuelle Dienstkonten (NT SERVICE)" -and $item.AccountCategory -ne "Virtuelles Dienstkonto") { return $false }

            if ($selectedMode -ne "Alle" -and $item.StartMode -ne $selectedMode) { return $false }
            if ($selectedState -ne "Alle" -and $item.State -ne $selectedState) { return $false }

            if (-not [string]::IsNullOrEmpty($searchTxt)) {
                $match = ($item.ServiceName -match [regex]::Escape($searchTxt)) -or 
                         ($item.DisplayName -match [regex]::Escape($searchTxt)) -or 
                         ($item.StartName -match [regex]::Escape($searchTxt)) -or 
                         ($item.Description -match [regex]::Escape($searchTxt))
                if (-not $match) { return $false }
            }

            return $true
        }

        $arr = [System.Collections.ArrayList]::new()
        foreach ($f in $filtered) { [void]$arr.Add($f) }
        
        $gridServices.SuspendLayout()
        $gridServices.DataSource = $arr
        $gridServices.ClearSelection()
        $gridServices.ResumeLayout()

        $lblStatus.Text = "Gefiltert: $($arr.Count) von $($script:allServices.Count) Diensten."
    }

    $cmbAccountCat.Add_SelectedIndexChanged($applyFilter)
    $cmbStartMode.Add_SelectedIndexChanged($applyFilter)
    $cmbState.Add_SelectedIndexChanged($applyFilter)
    $txtFilter.Add_TextChanged($applyFilter)

    # --------------------------------------------------------------------------
    # 10. SPALTENSORTIERUNG
    # --------------------------------------------------------------------------
    $gridServices.Add_ColumnHeaderMouseClick({
        param($s, $e)
        $prop = $gridServices.Columns[$e.ColumnIndex].DataPropertyName
        if (-not $prop) { return }

        if ($script:sortCol -eq $prop) { 
            $script:sortAsc = -not $script:sortAsc 
        } else { 
            $script:sortCol = $prop
            $script:sortAsc = $true 
        }

        $current = @($gridServices.DataSource)
        if ($current.Count -le 1) { return }

        $sorted = $current | Sort-Object -Property @{ Expression = { $_.$prop }; Descending = (-not $script:sortAsc) }
        $arr = [System.Collections.ArrayList]::new()
        foreach ($item in $sorted) { [void]$arr.Add($item) }
        
        $gridServices.SuspendLayout()
        $gridServices.DataSource = $arr
        $gridServices.ClearSelection()
        $gridServices.ResumeLayout()
    })

    # --------------------------------------------------------------------------
    # 11. CSV-EXPORT
    # --------------------------------------------------------------------------
    $btnExport.Add_Click({
        if (-not $gridServices.DataSource -or $gridServices.Rows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Keine Daten zum Exportieren vorhanden.", "Hinweis", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "CSV-Datei (*.csv)|*.csv"
        $sfd.FileName = "Server_Services_Audit_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $gridServices.DataSource | Export-Csv -Path $sfd.FileName -Delimiter ';' -NoTypeInformation -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show("Export erfolgreich erstellt!`n`nDatei: $($sfd.FileName)", "Export abgeschlossen", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    })

    # --------------------------------------------------------------------------
    # 12. SICHERE INITIALISIERUNG DES SPLITTERS
    # --------------------------------------------------------------------------
    $form.Add_Shown({
        try {
            if ($split.Width -gt 500) {
                $split.SplitterDistance = 350
            }
        } catch {}
    })

    # --------------------------------------------------------------------------
    # 13. CLEANUP & ISE-SCHUTZ
    # --------------------------------------------------------------------------
    $form.Add_FormClosing({
        $gridServices.DataSource = $null
        $gridServers.Rows.Clear()
    })

    try {
        [void]$form.ShowDialog()
    } finally {
        if ($null -ne $form) { $form.Dispose() }
        if ($script:allServices) { $script:allServices.Clear() }
        if ($script:rawServerList) { $script:rawServerList.Clear() }

        Get-PSSession | Remove-PSSession -ErrorAction SilentlyContinue
        Get-CimSession | Remove-CimSession -ErrorAction SilentlyContinue

        [System.Windows.Forms.Application]::ExitThread()
        [System.GC]::Collect()
    }
}

# Start
Show-Tool25-ServiceAudit
