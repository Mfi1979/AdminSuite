Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

function Show-ADObjectCompareTool {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Tool 19: AD Universal Objekt-Vergleich & Attribut-Inspektor"
    $form.Size = New-Object System.Drawing.Size(1460, 910)
    $form.MinimumSize = New-Object System.Drawing.Size(1200, 760)
    $form.StartPosition = "CenterScreen"
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

    # -------------------------------------------------------------
    # TYP-SPEZIFISCHE DEFAULT-ATTRIBUTE
    # -------------------------------------------------------------
    $script:TypeDefaultProps = @{
        "User (Benutzer)" = [string[]]@("department", "mail", "title", "telephoneNumber", "userPrincipalName")
        "Computer"        = [string[]]@("operatingSystem", "operatingSystemVersion", "dNSHostName", "whenCreated", "lastLogonTimestamp", "description")
        "Group (Gruppen)" = [string[]]@("description", "mail", "groupType", "whenCreated", "info")
    }

    $script:SelectedOverviewProps = [System.Collections.Generic.List[string]]::new($script:TypeDefaultProps["User (Benutzer)"])

    # -------------------------------------------------------------
    # TOP PANEL MIT DYNAMISCHEM LAYOUT
    # -------------------------------------------------------------
    $pnlTop = New-Object System.Windows.Forms.Panel
    $pnlTop.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlTop.AutoSize = $true
    $pnlTop.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $pnlTop.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 252)
    $pnlTop.Padding = New-Object System.Windows.Forms.Padding(15, 10, 15, 10)
    $form.Controls.Add($pnlTop)

    $stackLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $stackLayout.Dock = [System.Windows.Forms.DockStyle]::Top
    $stackLayout.AutoSize = $true
    $stackLayout.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $stackLayout.ColumnCount = 1
    $stackLayout.RowCount = 4
    [void]$stackLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    $pnlTop.Controls.Add($stackLayout)

    # -------------------------------------------------------------
    # ZEILE 1: Objekttyp & Erweiterter Ladefilter (Enthält / Beginn / Ende)
    # -------------------------------------------------------------
    $flowType = New-Object System.Windows.Forms.FlowLayoutPanel
    $flowType.AutoSize = $true
    $flowType.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
    $flowType.WrapContents = $false
    $flowType.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 6)

    $lblType = New-Object System.Windows.Forms.Label
    $lblType.Text = "Objekttyp:"
    $lblType.AutoSize = $true
    $lblType.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    $lblType.Margin = New-Object System.Windows.Forms.Padding(0, 4, 6, 0)
    $flowType.Controls.Add($lblType)

    $cmbType = New-Object System.Windows.Forms.ComboBox
    $cmbType.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$cmbType.Items.AddRange(@("User (Benutzer)", "Computer", "Group (Gruppen)"))
    $cmbType.SelectedIndex = 0
    $cmbType.Size = New-Object System.Drawing.Size(140, 25)
    $cmbType.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
    $flowType.Controls.Add($cmbType)

    $lblMode = New-Object System.Windows.Forms.Label
    $lblMode.Text = "Modus:"
    $lblMode.AutoSize = $true
    $lblMode.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblMode.Margin = New-Object System.Windows.Forms.Padding(0, 4, 4, 0)
    $flowType.Controls.Add($lblMode)

    $cmbMatchMode = New-Object System.Windows.Forms.ComboBox
    $cmbMatchMode.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$cmbMatchMode.Items.AddRange(@("Enthält (*Wert*)", "Beginnt mit (Wert*)", "Endet mit (*Wert)"))
    $cmbMatchMode.SelectedIndex = 0
    $cmbMatchMode.Size = New-Object System.Drawing.Size(145, 25)
    $cmbMatchMode.Margin = New-Object System.Windows.Forms.Padding(0, 0, 6, 0)
    $flowType.Controls.Add($cmbMatchMode)

    $txtObjFilter = New-Object System.Windows.Forms.TextBox
    $txtObjFilter.Size = New-Object System.Drawing.Size(150, 24)
    $txtObjFilter.Margin = New-Object System.Windows.Forms.Padding(0, 1, 10, 0)
    $flowType.Controls.Add($txtObjFilter)

    $btnLoad = New-Object System.Windows.Forms.Button
    $btnLoad.Text = "🔍 Objekte laden"
    $btnLoad.Size = New-Object System.Drawing.Size(125, 27)
    $btnLoad.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $btnLoad.ForeColor = [System.Drawing.Color]::White
    $btnLoad.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnLoad.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
    $flowType.Controls.Add($btnLoad)

    $btnAllOverview = New-Object System.Windows.Forms.Button
    $btnAllOverview.Text = "📋 Alle Objekte (Kompakt)"
    $btnAllOverview.Size = New-Object System.Drawing.Size(185, 27)
    $btnAllOverview.BackColor = [System.Drawing.Color]::FromArgb(70, 80, 95)
    $btnAllOverview.ForeColor = [System.Drawing.Color]::White
    $btnAllOverview.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnAllOverview.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
    $flowType.Controls.Add($btnAllOverview)

    $btnCustomizeCols = New-Object System.Windows.Forms.Button
    $btnCustomizeCols.Text = "⚙️ Spalten anpassen"
    $btnCustomizeCols.Size = New-Object System.Drawing.Size(150, 27)
    $btnCustomizeCols.BackColor = [System.Drawing.Color]::FromArgb(235, 240, 248)
    $btnCustomizeCols.ForeColor = [System.Drawing.Color]::FromArgb(30, 41, 59)
    $btnCustomizeCols.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $flowType.Controls.Add($btnCustomizeCols)

    [void]$stackLayout.Controls.Add($flowType, 0, 0)

    # -------------------------------------------------------------
    # ZEILE 2: 2 Kacheln & Auswerten/Vergleichen-Button
    # -------------------------------------------------------------
    $tblBoxes = New-Object System.Windows.Forms.TableLayoutPanel
    $tblBoxes.Dock = [System.Windows.Forms.DockStyle]::Fill
    $tblBoxes.AutoSize = $true
    $tblBoxes.ColumnCount = 3
    $tblBoxes.RowCount = 1
    [void]$tblBoxes.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 43)))
    [void]$tblBoxes.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 43)))
    [void]$tblBoxes.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 14)))
    $tblBoxes.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 6)

    # Box Links
    $grpLeft = New-Object System.Windows.Forms.GroupBox
    $grpLeft.Text = " Basis-Objekt 1 (Links) "
    $grpLeft.Dock = [System.Windows.Forms.DockStyle]::Fill
    $grpLeft.Height = 110
    $grpLeft.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
    $grpLeft.Padding = New-Object System.Windows.Forms.Padding(10, 6, 10, 6)

    $tblBox1 = New-Object System.Windows.Forms.TableLayoutPanel
    $tblBox1.Dock = [System.Windows.Forms.DockStyle]::Fill
    $tblBox1.ColumnCount = 2
    $tblBox1.RowCount = 2
    [void]$tblBox1.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 68)))
    [void]$tblBox1.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))

    $lblFilterL = New-Object System.Windows.Forms.Label
    $lblFilterL.Text = "Filter:"
    $lblFilterL.Anchor = [System.Windows.Forms.AnchorStyles]::Left; $lblFilterL.AutoSize = $true
    $lblFilterL.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
    [void]$tblBox1.Controls.Add($lblFilterL, 0, 0)

    $txtFilterL = New-Object System.Windows.Forms.TextBox
    $txtFilterL.Dock = [System.Windows.Forms.DockStyle]::Fill
    $txtFilterL.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
    [void]$tblBox1.Controls.Add($txtFilterL, 1, 0)

    $lblSelL = New-Object System.Windows.Forms.Label
    $lblSelL.Text = "Objekt:"
    $lblSelL.Anchor = [System.Windows.Forms.AnchorStyles]::Left; $lblSelL.AutoSize = $true
    $lblSelL.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
    [void]$tblBox1.Controls.Add($lblSelL, 0, 1)

    $cmbObj1 = New-Object System.Windows.Forms.ComboBox
    $cmbObj1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $cmbObj1.Dock = [System.Windows.Forms.DockStyle]::Fill
    $cmbObj1.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
    [void]$tblBox1.Controls.Add($cmbObj1, 1, 1)
    $grpLeft.Controls.Add($tblBox1)
    [void]$tblBoxes.Controls.Add($grpLeft, 0, 0)

    # Box Rechts
    $grpRight = New-Object System.Windows.Forms.GroupBox
    $grpRight.Text = " Vergleichs-Objekt 2 (Rechts) "
    $grpRight.Dock = [System.Windows.Forms.DockStyle]::Fill
    $grpRight.Height = 110
    $grpRight.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
    $grpRight.Padding = New-Object System.Windows.Forms.Padding(10, 6, 10, 6)

    $tblBox2 = New-Object System.Windows.Forms.TableLayoutPanel
    $tblBox2.Dock = [System.Windows.Forms.DockStyle]::Fill
    $tblBox2.ColumnCount = 2
    $tblBox2.RowCount = 2
    [void]$tblBox2.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 68)))
    [void]$tblBox2.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))

    $lblFilterR = New-Object System.Windows.Forms.Label
    $lblFilterR.Text = "Filter:"
    $lblFilterR.Anchor = [System.Windows.Forms.AnchorStyles]::Left; $lblFilterR.AutoSize = $true
    $lblFilterR.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
    [void]$tblBox2.Controls.Add($lblFilterR, 0, 0)

    $txtFilterR = New-Object System.Windows.Forms.TextBox
    $txtFilterR.Dock = [System.Windows.Forms.DockStyle]::Fill
    $txtFilterR.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
    [void]$tblBox2.Controls.Add($txtFilterR, 1, 0)

    $lblSelR = New-Object System.Windows.Forms.Label
    $lblSelR.Text = "Objekt:"
    $lblSelR.Anchor = [System.Windows.Forms.AnchorStyles]::Left; $lblSelR.AutoSize = $true
    $lblSelR.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
    [void]$tblBox2.Controls.Add($lblSelR, 0, 1)

    $cmbObj2 = New-Object System.Windows.Forms.ComboBox
    $cmbObj2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $cmbObj2.Dock = [System.Windows.Forms.DockStyle]::Fill
    $cmbObj2.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
    [void]$tblBox2.Controls.Add($cmbObj2, 1, 1)
    $grpRight.Controls.Add($tblBox2)
    [void]$tblBoxes.Controls.Add($grpRight, 1, 0)

    $btnCompare = New-Object System.Windows.Forms.Button
    $btnCompare.Text = "Auswerten /`nVergleichen"
    $btnCompare.Dock = [System.Windows.Forms.DockStyle]::Fill
    $btnCompare.Margin = New-Object System.Windows.Forms.Padding(8, 6, 0, 0)
    $btnCompare.BackColor = [System.Drawing.Color]::FromArgb(16, 124, 65)
    $btnCompare.ForeColor = [System.Drawing.Color]::White
    $btnCompare.Font = New-Object System.Drawing.Font("Segoe UI", 10.5, [System.Drawing.FontStyle]::Bold)
    $btnCompare.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    [void]$tblBoxes.Controls.Add($btnCompare, 2, 0)

    [void]$stackLayout.Controls.Add($tblBoxes, 0, 1)

    # -------------------------------------------------------------
    # ZEILE 3: Zweifacher Filter (Attribut + Wert) & Aktionen
    # -------------------------------------------------------------
    $flowFilter = New-Object System.Windows.Forms.FlowLayoutPanel
    $flowFilter.AutoSize = $true
    $flowFilter.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
    $flowFilter.WrapContents = $false
    $flowFilter.Margin = New-Object System.Windows.Forms.Padding(0, 4, 0, 4)

    $lblAttrFilter = New-Object System.Windows.Forms.Label
    $lblAttrFilter.Text = "Attribut enthält:"
    $lblAttrFilter.AutoSize = $true
    $lblAttrFilter.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblAttrFilter.Margin = New-Object System.Windows.Forms.Padding(0, 4, 4, 0)
    $flowFilter.Controls.Add($lblAttrFilter)

    $txtAttrFilter = New-Object System.Windows.Forms.TextBox
    $txtAttrFilter.Size = New-Object System.Drawing.Size(150, 24)
    $txtAttrFilter.Margin = New-Object System.Windows.Forms.Padding(0, 1, 16, 0)
    $flowFilter.Controls.Add($txtAttrFilter)

    $lblValFilter = New-Object System.Windows.Forms.Label
    $lblValFilter.Text = "Wert enthält:"
    $lblValFilter.AutoSize = $true
    $lblValFilter.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblValFilter.Margin = New-Object System.Windows.Forms.Padding(0, 4, 4, 0)
    $flowFilter.Controls.Add($lblValFilter)

    $txtValFilter = New-Object System.Windows.Forms.TextBox
    $txtValFilter.Size = New-Object System.Drawing.Size(170, 24)
    $txtValFilter.Margin = New-Object System.Windows.Forms.Padding(0, 1, 16, 0)
    $flowFilter.Controls.Add($txtValFilter)

    $chkOnlyDiff = New-Object System.Windows.Forms.CheckBox
    $chkOnlyDiff.Text = "Nur Unterschiede"
    $chkOnlyDiff.AutoSize = $true
    $chkOnlyDiff.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $chkOnlyDiff.Margin = New-Object System.Windows.Forms.Padding(0, 3, 16, 0)
    $flowFilter.Controls.Add($chkOnlyDiff)

    $btnExport = New-Object System.Windows.Forms.Button
    $btnExport.Text = "📥 CSV Export"
    $btnExport.Size = New-Object System.Drawing.Size(120, 26)
    $btnExport.BackColor = [System.Drawing.Color]::FromArgb(70, 80, 95)
    $btnExport.ForeColor = [System.Drawing.Color]::White
    $btnExport.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnExport.Margin = New-Object System.Windows.Forms.Padding(0, 0, 16, 0)
    $flowFilter.Controls.Add($btnExport)

    $lblCount = New-Object System.Windows.Forms.Label
    $lblCount.Text = ""
    $lblCount.AutoSize = $true
    $lblCount.ForeColor = [System.Drawing.Color]::FromArgb(70, 80, 90)
    $lblCount.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
    $lblCount.Margin = New-Object System.Windows.Forms.Padding(0, 4, 0, 0)
    $flowFilter.Controls.Add($lblCount)

    [void]$stackLayout.Controls.Add($flowFilter, 0, 2)

    # -------------------------------------------------------------
    # ZEILE 4: Legende
    # -------------------------------------------------------------
    $lblLegend = New-Object System.Windows.Forms.Label
    $lblLegend.Text = "💡 Klick auf Spaltenkopf sortiert die Tabelle | Rechtsklick auf Zeile wählt Objekt für Vergleich | Legende: [Grün] Identisch | [Gelb/Orange] Abweichend"
    $lblLegend.AutoSize = $true
    $lblLegend.ForeColor = [System.Drawing.Color]::FromArgb(100, 110, 120)
    $lblLegend.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
    [void]$stackLayout.Controls.Add($lblLegend, 0, 3)

    # -------------------------------------------------------------
    # DATAGRIDVIEW
    # -------------------------------------------------------------
    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = [System.Windows.Forms.DockStyle]::Fill
    $grid.ReadOnly = $true
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grid.BackgroundColor = [System.Drawing.Color]::White
    $grid.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $grid.CellBorderStyle = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
    $grid.GridColor = [System.Drawing.Color]::FromArgb(226, 232, 240)
    $grid.RowHeadersVisible = $false

    $grid.EnableHeadersVisualStyles = $false
    $grid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
    $grid.ColumnHeadersHeight = 36
    $grid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(235, 241, 250)
    $grid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(30, 41, 59)
    $grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    $grid.ColumnHeadersDefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(8, 0, 8, 0)

    $grid.RowTemplate.Height = 28
    $grid.DefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9.0)
    $grid.DefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(8, 2, 8, 2)
    $grid.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(210, 230, 250)
    $grid.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black

    # -------------------------------------------------------------
    # RECHTSKLICK-KONTEXTMENÜ (Objekt 1 / Objekt 2 zuweisen)
    # -------------------------------------------------------------
    $contextMenu = New-Object System.Windows.Forms.ContextMenuStrip

    $menuItemSetLeft = New-Object System.Windows.Forms.ToolStripMenuItem
    $menuItemSetLeft.Text = "👉 Als Basis-Objekt 1 setzen (Links)"
    $contextMenu.Items.Add($menuItemSetLeft) | Out-Null

    $menuItemSetRight = New-Object System.Windows.Forms.ToolStripMenuItem
    $menuItemSetRight.Text = "👉 Als Vergleichs-Objekt 2 setzen (Rechts)"
    $contextMenu.Items.Add($menuItemSetRight) | Out-Null

    $grid.ContextMenuStrip = $contextMenu

    # Rechtsklick selektiert direkt die Zeile unter dem Mauszeiger
    $grid.Add_CellMouseDown({
        param($sender, $e)
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right -and $e.RowIndex -ge 0) {
            $grid.ClearSelection()
            $grid.Rows[$e.RowIndex].Selected = $true
        }
    })

    # Helferfunktion: DisplayName des aktuell markierten Objekts ermitteln
    function Get-SelectedObjectDisplayName {
        if ($grid.SelectedRows.Count -eq 0) { return $null }
        $row = $grid.SelectedRows[0]

        # 1. Übersichtsmodus: Spalte "DisplayName"
        if ($row.Cells["DisplayName"] -and $row.Cells["DisplayName"].Value) {
            return [string]$row.Cells["DisplayName"].Value
        }
        # 2. Vergleichsmodus: Fallback auf Name/SAM
        if ($row.Cells["Name"] -and $row.Cells["Name"].Value) {
            $n = [string]$row.Cells["Name"].Value
            $s = if ($row.Cells["SAM"]) { [string]$row.Cells["SAM"].Value } else { "" }
            return if ($s) { "$n ($s)" } else { $n }
        }
        return $null
    }

    # Menü-Aktion: Links setzen
    $menuItemSetLeft.Add_Click({
        $objName = Get-SelectedObjectDisplayName
        if (-not $objName) { return }

        $foundIdx = $cmbObj1.FindStringExact($objName)
        if ($foundIdx -ge 0) {
            $cmbObj1.SelectedIndex = $foundIdx
        } else {
            # Bei Teilübereinstimmung
            for ($i = 0; $i -lt $cmbObj1.Items.Count; $i++) {
                if ($cmbObj1.Items[$i].ToString() -like "*$objName*") {
                    $cmbObj1.SelectedIndex = $i
                    break
                }
            }
        }
    })

    # Menü-Aktion: Rechts setzen
    $menuItemSetRight.Add_Click({
        $objName = Get-SelectedObjectDisplayName
        if (-not $objName) { return }

        $foundIdx = $cmbObj2.FindStringExact($objName)
        if ($foundIdx -ge 0) {
            $cmbObj2.SelectedIndex = $foundIdx
        } else {
            for ($i = 0; $i -lt $cmbObj2.Items.Count; $i++) {
                if ($cmbObj2.Items[$i].ToString() -like "*$objName*") {
                    $cmbObj2.SelectedIndex = $i
                    break
                }
            }
        }
    })

    $form.Controls.Add($grid)
    $grid.BringToFront()
    $pnlTop.SendToBack()

    # -------------------------------------------------------------
    # LOGIK & DATEN
    # -------------------------------------------------------------
    $script:LoadedObjects = @()
    $script:CurrentDisplayList = [System.Collections.Generic.List[PSCustomObject]]::new()
    $script:ViewMode = "Overview"

    # Spalten-Klicksortierung
    $grid.Add_ColumnHeaderMouseClick({
        param($sender, $e)
        try {
            $targetGrid = $sender
            if ($null -eq $targetGrid -or $null -eq $targetGrid.Columns) { return }
            if ($e.ColumnIndex -lt 0 -or $e.ColumnIndex -ge $targetGrid.Columns.Count) { return }

            $col = $targetGrid.Columns[$e.ColumnIndex]
            if ($null -eq $col) { return }
            $propName = if ($col.DataPropertyName) { $col.DataPropertyName } else { $col.HeaderText }

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
                    if ($null -eq $val) { return "" }
                    if ($val -as [int]) { return [int]$val }
                    return $val
                }
                Descending = (-not $state.Asc)
            }

            $targetGrid.SuspendLayout()
            $arr = [System.Collections.ArrayList]::new()
            foreach ($it in $sorted) { [void]$arr.Add($it) }
            $targetGrid.DataSource = $null
            $targetGrid.DataSource = $arr

            foreach ($c in $targetGrid.Columns) {
                $c.HeaderCell.SortGlyphDirection = [System.Windows.Forms.SortOrder]::None
                $c.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Programmatic
            }
            $targetGrid.Columns[$e.ColumnIndex].HeaderCell.SortGlyphDirection = `
                $(if ($state.Asc) { [System.Windows.Forms.SortOrder]::Ascending } else { [System.Windows.Forms.SortOrder]::Descending })

            $targetGrid.ResumeLayout()
        } catch { }
    })

    # Farbindikatoren
    $grid.Add_RowPrePaint({
        param($s, $e)
        if ($script:ViewMode -ne "Compare") { return }
        if ($e.RowIndex -ge 0 -and $e.RowIndex -lt $grid.Rows.Count) {
            $row = $grid.Rows[$e.RowIndex]
            $st = [string]$row.Cells["Status"].Value
            switch ($st) {
                "Identisch" {
                    $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(238, 247, 238)
                    $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(20, 100, 30)
                }
                "Abweichend" {
                    $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(254, 249, 231)
                    $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(160, 90, 0)
                }
                "Nur bei Objekt 1" {
                    $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(253, 237, 237)
                    $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(150, 20, 20)
                }
                "Nur bei Objekt 2" {
                    $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(237, 244, 253)
                    $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(20, 50, 140)
                }
                "Info (Einzelansicht)" {
                    $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::White
                    $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(30, 41, 59)
                }
            }
        }
    })

    # Intelligente Konvertierung
    function Convert-ADValToString {
        param($propName, $propVal)
        if ($null -eq $propVal) { return "" }

        if ($propName -eq "grouptype" -and ($propVal -as [int64])) {
            $gt = [int64]$propVal
            $isSec = (($gt -band 0x80000000) -ne 0)
            $typeStr = if ($isSec) { "Sicherheit" } else { "Verteilung" }
            $scopeStr = if (($gt -band 2) -ne 0) { "Global" }
                        elseif (($gt -band 4) -ne 0) { "Domänenlokal" }
                        elseif (($gt -band 8) -ne 0) { "Universell" }
                        else { "Lokal" }
            return "$scopeStr ($typeStr) [$gt]"
        }

        if ($propVal -is [System.Byte[]]) {
            try { return ([System.Security.Principal.SecurityIdentifier]::new($propVal, 0)).Value }
            catch { return [System.BitConverter]::ToString($propVal) }
        }
        if ($propVal -is [System.Collections.IEnumerable] -and -not ($propVal -is [string])) {
            return ($propVal | ForEach-Object { Convert-ADValToString $propName $_ }) -join " ; "
        }
        return $propVal.ToString()
    }

    # AD-Objekte mit exakter Klammer-Syntax abfragen
    $loadObjectsAction = {
        $type = $cmbType.SelectedItem.ToString()

        $rawTerm = $txtObjFilter.Text.Trim()
        $termClean = $rawTerm.Trim('*')

        $pattern = "*"
        if ($termClean) {
            switch ($cmbMatchMode.SelectedIndex) {
                0 { $pattern = "*$termClean*" }
                1 { $pattern = "$termClean*" }
                2 { $pattern = "*$termClean" }
            }
        }

        $finalFilter = if ($pattern -eq "*") {
            switch -Wildcard ($type) {
                "User*"     { "(&(objectCategory=person)(objectClass=user))" }
                "Computer*" { "(objectCategory=computer)" }
                "Group*"    { "(objectCategory=group)" }
            }
        } else {
            switch -Wildcard ($type) {
                "User*"     { "(&(objectCategory=person)(objectClass=user)(|(name=$pattern)(sAMAccountName=$pattern)))" }
                "Computer*" { "(&(objectCategory=computer)(|(name=$pattern)(sAMAccountName=$pattern)))" }
                "Group*"    { "(&(objectCategory=group)(|(name=$pattern)(sAMAccountName=$pattern)))" }
            }
        }

        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $root = [ADSI]"LDAP://RootDSE"
            $searcher = [System.DirectoryServices.DirectorySearcher]::new([ADSI]"LDAP://$($root.defaultNamingContext)")
            $searcher.Filter = $finalFilter
            $searcher.PageSize = 1000
            
            [string[]]$propsToLoad = (@("name", "sAMAccountName", "distinguishedName", "l") + $script:SelectedOverviewProps) | Select-Object -Unique
            $searcher.PropertiesToLoad.AddRange($propsToLoad)

            $script:LoadedObjects = @($searcher.FindAll() | ForEach-Object {
                $sam = if ($_.Properties["samaccountname"]) { $_.Properties["samaccountname"][0] } else { $_.Properties["name"][0] }
                $loc = if ($_.Properties["l"]) { $_.Properties["l"][0] } else { "-" }

                $dynObj = [ordered]@{
                    DisplayName = "$($_.Properties['name'][0]) ($sam)"
                    Name        = [string]$_.Properties["name"][0]
                    SAM         = [string]$sam
                    Location    = [string]$loc
                    DN          = [string]$_.Properties["distinguishedname"][0]
                }
                foreach ($p in $script:SelectedOverviewProps) {
                    if ($p -ne "l") {
                        $pKey = $p.ToLower()
                        $dynObj[$p] = if ($_.Properties[$pKey]) { Convert-ADValToString $pKey $_.Properties[$pKey][0] } else { "-" }
                    }
                }
                [PSCustomObject]$dynObj
            } | Sort-Object DisplayName)

            & $updateDropdowns

            if ($script:ViewMode -eq "Overview" -or $script:CurrentDisplayList.Count -eq 0) {
                $script:ViewMode = "Overview"
                $script:CurrentDisplayList.Clear()
                foreach ($o in $script:LoadedObjects) {
                    $rowHash = [ordered]@{
                        "DisplayName" = $o.DisplayName
                        "SAM"         = $o.SAM
                        "Location"    = $o.Location
                    }
                    foreach ($p in $script:SelectedOverviewProps) {
                        if ($p -ne "l") {
                            $rowHash[$p] = $o.$p
                        }
                    }
                    $rowHash["DN"] = $o.DN
                    $script:CurrentDisplayList.Add([PSCustomObject]$rowHash)
                }
                & $renderGrid
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Laden: $($_.Exception.Message)", "Fehler", "OK", "Error")
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }

    # Dropdowns aktualisieren
    $updateDropdowns = {
        $fL = $txtFilterL.Text.Trim()
        $fR = $txtFilterR.Text.Trim()

        $cmbObj1.Items.Clear()
        $cmbObj2.Items.Clear()

        [void]$cmbObj2.Items.Add("(Kein Vergleich - Nur Objekt 1 anzeigen)")

        $list1 = if ($fL) { @($script:LoadedObjects | Where-Object { $_.DisplayName -like "*$fL*" }) } else { $script:LoadedObjects }
        $list2 = if ($fR) { @($script:LoadedObjects | Where-Object { $_.DisplayName -like "*$fR*" }) } else { $script:LoadedObjects }

        foreach ($o in $list1) { [void]$cmbObj1.Items.Add($o.DisplayName) }
        foreach ($o in $list2) { [void]$cmbObj2.Items.Add($o.DisplayName) }

        if ($cmbObj1.Items.Count -gt 0) { $cmbObj1.SelectedIndex = 0 }
        if ($cmbObj2.Items.Count -gt 2) { $cmbObj2.SelectedIndex = 2 }
        elseif ($cmbObj2.Items.Count -gt 0) { $cmbObj2.SelectedIndex = 0 }
    }

    # Typ-Wechsel Event
    $cmbType.Add_SelectedIndexChanged({
        $curType = $cmbType.SelectedItem.ToString()
        $script:SelectedOverviewProps.Clear()
        if ($script:TypeDefaultProps.ContainsKey($curType)) {
            $script:SelectedOverviewProps.AddRange($script:TypeDefaultProps[$curType])
        }
        $script:ViewMode = "Overview"
        & $loadObjectsAction
    })

    # Such-Events
    $btnLoad.Add_Click({
        $script:ViewMode = "Overview"
        & $loadObjectsAction
    })

    $txtObjFilter.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $_.SuppressKeyPress = $true
            $script:ViewMode = "Overview"
            & $loadObjectsAction
        }
    })

    $txtFilterL.Add_TextChanged($updateDropdowns)
    $txtFilterR.Add_TextChanged($updateDropdowns)

    # Render Grid Engine
    $renderGrid = {
        $attrPattern = $txtAttrFilter.Text.Trim()
        $valPattern  = $txtValFilter.Text.Trim()

        if ($script:ViewMode -eq "Compare") {
            $onlyDiff = $chkOnlyDiff.Checked

            $filtered = $script:CurrentDisplayList | Where-Object {
                if ($onlyDiff -and $_.Status -eq "Identisch") { return $false }

                if ($attrPattern -and -not ($_."AD Attribut" -like "*$attrPattern*")) { return $false }

                if ($valPattern) {
                    $m1 = $_."Wert bei Objekt 1" -like "*$valPattern*"
                    $m2 = $_."Wert bei Objekt 2" -like "*$valPattern*"
                    if (-not ($m1 -or $m2)) { return $false }
                }
                return $true
            }

            $arr = [System.Collections.ArrayList]::new(@($filtered))
            $grid.DataSource = $null
            $grid.DataSource = $arr

            if ($grid.Columns["AD Attribut"]) { $grid.Columns["AD Attribut"].Width = 240 }
            if ($grid.Columns["Wert bei Objekt 1"]) { $grid.Columns["Wert bei Objekt 1"].Width = 380 }
            if ($grid.Columns["Wert bei Objekt 2"]) { $grid.Columns["Wert bei Objekt 2"].Width = 380 }
            if ($grid.Columns["Status"]) { $grid.Columns["Status"].AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill }

            $lblCount.Text = "$($arr.Count) von $($script:CurrentDisplayList.Count) Attributen gelistet"
        }
        else {
            $filtered = $script:CurrentDisplayList | Where-Object {
                if ($valPattern) {
                    $match = $false
                    foreach ($prop in $_.PSObject.Properties.Name) {
                        if ($_.$prop -like "*$valPattern*") { $match = $true; break }
                    }
                    if (-not $match) { return $false }
                }
                return $true
            }

            $arr = [System.Collections.ArrayList]::new(@($filtered))
            $grid.DataSource = $null
            $grid.DataSource = $arr

            if ($grid.Columns["DisplayName"]) { $grid.Columns["DisplayName"].Width = 260 }
            if ($grid.Columns["SAM"])         { $grid.Columns["SAM"].Width = 120 }
            if ($grid.Columns["Location"])    { $grid.Columns["Location"].Width = 160 }
            if ($grid.Columns["DN"])          { $grid.Columns["DN"].AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill }

            $lblCount.Text = "$($arr.Count) von $($script:CurrentDisplayList.Count) Objekten gelistet"
        }

        foreach ($c in $grid.Columns) { $c.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Programmatic }
        $grid.ClearSelection()
    }

    # 1. VERGLEICH ODER EINZELBETRACHTUNG
    $btnCompare.Add_Click({
        if ($cmbObj1.SelectedIndex -lt 0) {
            [System.Windows.Forms.MessageBox]::Show("Bitte mindestens für Objekt 1 eine Auswahl treffen.", "Auswahl fehlt", "OK", "Warning")
            return
        }

        $isSingleView = ($cmbObj2.SelectedIndex -le 0)
        $obj1Meta = $script:LoadedObjects | Where-Object { $_.DisplayName -eq $cmbObj1.SelectedItem.ToString() } | Select-Object -First 1
        $obj2Meta = if (-not $isSingleView) { $script:LoadedObjects | Where-Object { $_.DisplayName -eq $cmbObj2.SelectedItem.ToString() } | Select-Object -First 1 } else { $null }

        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $script:ViewMode = "Compare"
            $script:CurrentDisplayList.Clear()

            $e1 = [ADSI]"LDAP://$($obj1Meta.DN)"
            $props1 = @{}
            foreach ($p in $e1.Properties.PropertyNames) {
                $val = Convert-ADValToString $p.ToLower() $e1.Properties[$p].Value
                if (-not [string]::IsNullOrWhiteSpace($val)) { $props1[$p] = $val }
            }

            if ($isSingleView) {
                foreach ($attr in ($props1.Keys | Sort-Object)) {
                    $script:CurrentDisplayList.Add([PSCustomObject]@{
                        "AD Attribut"        = $attr
                        "Wert bei Objekt 1"  = $props1[$attr]
                        "Wert bei Objekt 2"  = "-- [Nicht abgefragt] --"
                        "Status"             = "Info (Einzelansicht)"
                    })
                }
            } else {
                $e2 = [ADSI]"LDAP://$($obj2Meta.DN)"
                $props2 = @{}
                foreach ($p in $e2.Properties.PropertyNames) {
                    $val = Convert-ADValToString $p.ToLower() $e2.Properties[$p].Value
                    if (-not [string]::IsNullOrWhiteSpace($val)) { $props2[$p] = $val }
                }

                $allProps = ($props1.Keys + $props2.Keys) | Select-Object -Unique | Sort-Object
                foreach ($attr in $allProps) {
                    $has1 = $props1.ContainsKey($attr)
                    $has2 = $props2.ContainsKey($attr)
                    $v1 = if ($has1) { $props1[$attr] } else { "-- [Kein Wert] --" }
                    $v2 = if ($has2) { $props2[$attr] } else { "-- [Kein Wert] --" }

                    $status = ""
                    if ($has1 -and $has2) {
                        $status = if ($v1 -eq $v2) { "Identisch" } else { "Abweichend" }
                    } elseif ($has1) {
                        $status = "Nur bei Objekt 1"
                    } else {
                        $status = "Nur bei Objekt 2"
                    }

                    $script:CurrentDisplayList.Add([PSCustomObject]@{
                        "AD Attribut"        = $attr
                        "Wert bei Objekt 1"  = $v1
                        "Wert bei Objekt 2"  = $v2
                        "Status"             = $status
                    })
                }
            }

            & $renderGrid
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Abruf: $($_.Exception.Message)", "Fehler", "OK", "Error")
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    })

    # 2. GESAMTÜBERSICHT ALLER OBJEKTE
    $btnAllOverview.Add_Click({
        if ($script:LoadedObjects.Count -eq 0) { return }

        $script:ViewMode = "Overview"
        $script:CurrentDisplayList.Clear()

        foreach ($o in $script:LoadedObjects) {
            $rowHash = [ordered]@{
                "DisplayName" = $o.DisplayName
                "SAM"         = $o.SAM
                "Location"    = $o.Location
            }
            foreach ($p in $script:SelectedOverviewProps) {
                if ($p -ne "l") {
                    $rowHash[$p] = $o.$p
                }
            }
            $rowHash["DN"] = $o.DN
            $script:CurrentDisplayList.Add([PSCustomObject]$rowHash)
        }

        & $renderGrid
    })

    # 3. SPALTEN- & ATTRIBUT-AUSWAHLDIALOG
    $btnCustomizeCols.Add_Click({
        $diag = New-Object System.Windows.Forms.Form
        $diag.Text = "Spalten & Attribute konfigurieren ($($cmbType.SelectedItem))"
        $diag.Size = New-Object System.Drawing.Size(460, 580)
        $diag.StartPosition = "CenterParent"
        $diag.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        $diag.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $diag.MaximizeBox = $false
        $diag.MinimizeBox = $false

        $curType = $cmbType.SelectedItem.ToString()

        $liveScannedProps = @()
        if ($script:LoadedObjects.Count -gt 0) {
            $targetSample = if ($cmbObj1.SelectedIndex -ge 0) {
                $script:LoadedObjects | Where-Object { $_.DisplayName -eq $cmbObj1.SelectedItem.ToString() } | Select-Object -First 1
            } else {
                $script:LoadedObjects[0]
            }
            if ($targetSample) {
                try {
                    $sampleEntry = [ADSI]"LDAP://$($targetSample.DN)"
                    $liveScannedProps = @($sampleEntry.Properties.PropertyNames)
                } catch { }
            }
        }

        $diagLbl = New-Object System.Windows.Forms.Label
        $diagLbl.Text = "Wähle Attribute für die Übersicht (Location ist fix aktiv):"
        $diagLbl.Location = "15, 12"; $diagLbl.AutoSize = $true
        $diag.Controls.Add($diagLbl)

        $txtDiagSearch = New-Object System.Windows.Forms.TextBox
        $txtDiagSearch.Location = "15, 34"; $txtDiagSearch.Size = "410, 24"
        $diag.Controls.Add($txtDiagSearch)

        $chkList = New-Object System.Windows.Forms.CheckedListBox
        $chkList.Location = "15, 64"; $chkList.Size = "410, 320"
        $chkList.CheckOnClick = $true
        $diag.Controls.Add($chkList)

        $suggested = if ($script:TypeDefaultProps.ContainsKey($curType)) { $script:TypeDefaultProps[$curType] } else { @() }
        $masterList = ($suggested + $liveScannedProps + $script:SelectedOverviewProps) | Where-Object { $_ -ne "l" -and -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique | Sort-Object

        $fillCheckList = {
            param([string]$filterText = "")
            $chkList.Items.Clear()
            foreach ($prop in $masterList) {
                if ($filterText -and -not ($prop -like "*$filterText*")) { continue }
                $isChk = $script:SelectedOverviewProps.Contains($prop)
                [void]$chkList.Items.Add($prop, $isChk)
            }
        }

        & $fillCheckList ""

        $txtDiagSearch.Add_TextChanged({
            & $fillCheckList $txtDiagSearch.Text.Trim()
        })

        $lblCustom = New-Object System.Windows.Forms.Label
        $lblCustom.Text = "Manuelles Attribut eintragen:"
        $lblCustom.Location = "15, 395"; $lblCustom.AutoSize = $true
        $diag.Controls.Add($lblCustom)

        $txtCustom = New-Object System.Windows.Forms.TextBox
        $txtCustom.Location = "15, 417"; $txtCustom.Size = "290, 24"
        $diag.Controls.Add($txtCustom)

        $btnAddCustom = New-Object System.Windows.Forms.Button
        $btnAddCustom.Text = "Hinzufügen"
        $btnAddCustom.Location = "315, 416"; $btnAddCustom.Size = "110, 26"
        $btnAddCustom.Add_Click({
            $customVal = $txtCustom.Text.Trim()
            if ($customVal -and -not $chkList.Items.Contains($customVal)) {
                [void]$chkList.Items.Add($customVal, $true)
                if (-not $script:SelectedOverviewProps.Contains($customVal)) {
                    $script:SelectedOverviewProps.Add($customVal)
                }
                $txtCustom.Clear()
            }
        })
        $diag.Controls.Add($btnAddCustom)

        $btnOk = New-Object System.Windows.Forms.Button
        $btnOk.Text = "Speichern & Neu laden"
        $btnOk.Location = "265, 480"; $btnOk.Size = "160, 32"
        $btnOk.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
        $btnOk.ForeColor = [System.Drawing.Color]::White
        $btnOk.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $diag.Controls.Add($btnOk)

        if ($diag.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
            $script:SelectedOverviewProps.Clear()
            foreach ($item in $chkList.CheckedItems) {
                $script:SelectedOverviewProps.Add($item.ToString())
            }
            $script:ViewMode = "Overview"
            & $loadObjectsAction
        }
    })

    # Filter-Events
    $chkOnlyDiff.Add_CheckedChanged({ & $renderGrid })
    $txtAttrFilter.Add_TextChanged({ & $renderGrid })
    $txtValFilter.Add_TextChanged({ & $renderGrid })

    # CSV Export
    $btnExport.Add_Click({
        $itemsToExport = @($grid.DataSource)
        if ($null -eq $itemsToExport -or $itemsToExport.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Keine Daten zum Exportieren vorhanden.", "Export", "OK", "Warning")
            return
        }

        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "CSV-Datei (*.csv)|*.csv"
        $modePrefix = if ($script:ViewMode -eq "Compare") { "AD_Objektvergleich" } else { "AD_Objektuebersicht" }
        $sfd.FileName = "${modePrefix}_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"

        if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                $itemsToExport | Export-Csv -Path $sfd.FileName -Delimiter ';' -NoTypeInformation -Encoding UTF8
                [System.Windows.Forms.MessageBox]::Show("Erfolgreich $($itemsToExport.Count) Zeilen exportiert nach:`n$($sfd.FileName)", "Export erfolgreich", "OK", "Information")
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Fehler beim Exportieren: $($_.Exception.Message)", "Fehler", "OK", "Error")
            }
        }
    })

    $form.Add_Shown({ 
        $script:ViewMode = "Overview"
        & $loadObjectsAction 
    })
    [void]$form.ShowDialog()
}

Show-ADObjectCompareTool
