# =========================================================================
# Tab3_Backup.ps1 - GPO Backup & HTML Report Export
# =========================================================================

function Build-Tab3_Backup {
    param($tabControl)

    $tabBackup = New-Object System.Windows.Forms.TabPage
    $tabBackup.Text = "3. GPO Backup & Verknuepfungs-Status"
    $tabBackup.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
    $script:tabBackup = $tabBackup

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

    $script:txtBackupTargetDir = New-Object System.Windows.Forms.TextBox
    $script:txtBackupTargetDir.Location = New-Object System.Drawing.Point(140, 14)
    $script:txtBackupTargetDir.Size = New-Object System.Drawing.Size(360, 25)
    $script:txtBackupTargetDir.Text = "C:\Install\Backup\GPO"

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

    $script:btnExportAllHtml = New-Object System.Windows.Forms.Button
    $script:btnExportAllHtml.Text = "HTML (Gefilterte)"
    $script:btnExportAllHtml.Location = New-Object System.Drawing.Point(895, 11)
    $script:btnExportAllHtml.Size = New-Object System.Drawing.Size(150, 30)
    $script:btnExportAllHtml.BackColor = [System.Drawing.Color]::FromArgb(255, 238, 220)

    $lblBackupViewFilter = New-Object System.Windows.Forms.Label
    $lblBackupViewFilter.Text = "Ansicht:"
    $lblBackupViewFilter.Location = New-Object System.Drawing.Point(12, 58)
    $lblBackupViewFilter.AutoSize = $true
    $lblBackupViewFilter.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $script:comboBackupFilter = New-Object System.Windows.Forms.ComboBox
    $script:comboBackupFilter.Location = New-Object System.Drawing.Point(72, 55)
    $script:comboBackupFilter.Size = New-Object System.Drawing.Size(200, 25)
    $script:comboBackupFilter.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$script:comboBackupFilter.Items.AddRange(@("Alle GPOs", "Nur verlinkte GPOs", "Nicht verlinkte GPOs (Unlinked)"))
    $script:comboBackupFilter.SelectedIndex = 0

    $lblBackupSearch = New-Object System.Windows.Forms.Label
    $lblBackupSearch.Text = "Suche:"
    $lblBackupSearch.Location = New-Object System.Drawing.Point(282, 58)
    $lblBackupSearch.AutoSize = $true
    $lblBackupSearch.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    # Im Script-Scope registrieren, damit .Trim() im Update-Block nicht auf $null laeuft
    $script:txtBackupSearch = New-Object System.Windows.Forms.TextBox
    $script:txtBackupSearch.Location = New-Object System.Drawing.Point(335, 55)
    $script:txtBackupSearch.Size = New-Object System.Drawing.Size(140, 25)

    $btnLoadGpos = New-Object System.Windows.Forms.Button
    $btnLoadGpos.Text = "GPO-Liste laden"
    $btnLoadGpos.Location = New-Object System.Drawing.Point(485, 52)
    $btnLoadGpos.Size = New-Object System.Drawing.Size(120, 32)

    $btnBackupSelected = New-Object System.Windows.Forms.Button
    $btnBackupSelected.Text = "Ausgewaehlte GPO sichern"
    $btnBackupSelected.Location = New-Object System.Drawing.Point(615, 52)
    $btnBackupSelected.Size = New-Object System.Drawing.Size(190, 32)
    $btnBackupSelected.BackColor = [System.Drawing.Color]::FromArgb(230, 245, 230)

    $script:btnBackupAll = New-Object System.Windows.Forms.Button
    $script:btnBackupAll.Text = "ALLE GPOs sichern"
    $script:btnBackupAll.Location = New-Object System.Drawing.Point(815, 52)
    $script:btnBackupAll.Size = New-Object System.Drawing.Size(210, 32)
    $script:btnBackupAll.BackColor = [System.Drawing.Color]::FromArgb(255, 245, 230)

    $panelBackupTop.Controls.AddRange(@(
        $lblTargetPath, $script:txtBackupTargetDir, $btnBrowseFolder, $btnExportBackupCsv, 
        $btnExportSelectedHtml, $script:btnExportAllHtml, $lblBackupViewFilter, 
        $script:comboBackupFilter, $lblBackupSearch, $script:txtBackupSearch, 
        $btnLoadGpos, $btnBackupSelected, $script:btnBackupAll
    ))

    $splitBackupMain = New-Object System.Windows.Forms.SplitContainer
    $splitBackupMain.Dock = [System.Windows.Forms.DockStyle]::Fill
    $splitBackupMain.SplitterDistance = 750
    $splitBackupMain.SplitterWidth = 6

    $panelGpoLeft = New-Object System.Windows.Forms.Panel
    $panelGpoLeft.Dock = [System.Windows.Forms.DockStyle]::Fill
    $panelGpoLeft.Padding = New-Object System.Windows.Forms.Padding(10, 8, 4, 10)

    $script:lblGpoGrid = New-Object System.Windows.Forms.Label
    $script:lblGpoGrid.Text = "1. Gruppenrichtlinien der Domaene:"
    $script:lblGpoGrid.Dock = [System.Windows.Forms.DockStyle]::Top
    $script:lblGpoGrid.Height = 28
    $script:lblGpoGrid.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)

    $script:gridGpos = New-Object System.Windows.Forms.DataGridView
    $script:gridGpos.Dock = [System.Windows.Forms.DockStyle]::Fill
    $script:gridGpos.ReadOnly = $true
    $script:gridGpos.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
    $script:gridGpos.AllowUserToAddRows = $false
    $script:gridGpos.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $script:gridGpos.MultiSelect = $false
    $script:gridGpos.RowHeadersVisible = $false
    $script:gridGpos.BackgroundColor = [System.Drawing.Color]::White
    $script:gridGpos.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $script:gridGpos.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(230, 236, 245)
    $script:gridGpos.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    $script:gridGpos.ColumnHeadersHeight = 34
    $script:gridGpos.RowTemplate.Height = 28
    $script:gridGpos.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(249, 251, 254)

    $panelGpoLeft.Controls.Add($script:gridGpos)
    $panelGpoLeft.Controls.Add($script:lblGpoGrid)
    $splitBackupMain.Panel1.Controls.Add($panelGpoLeft)

    $splitBackupRight = New-Object System.Windows.Forms.SplitContainer
    $splitBackupRight.Dock = [System.Windows.Forms.DockStyle]::Fill
    $splitBackupRight.Orientation = [System.Windows.Forms.Orientation]::Horizontal
    $splitBackupRight.SplitterDistance = 280
    $splitBackupRight.SplitterWidth = 6

    $panelLinks = New-Object System.Windows.Forms.Panel
    $panelLinks.Dock = [System.Windows.Forms.DockStyle]::Fill
    $panelLinks.Padding = New-Object System.Windows.Forms.Padding(4, 8, 10, 4)

    $script:lblLinks = New-Object System.Windows.Forms.Label
    $script:lblLinks.Text = "2. Verknuepfungs-Ziele (OUs / Sites):"
    $script:lblLinks.Dock = [System.Windows.Forms.DockStyle]::Top
    $script:lblLinks.Height = 28
    $script:lblLinks.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)

    $script:gridLinks = New-Object System.Windows.Forms.DataGridView
    $script:gridLinks.Dock = [System.Windows.Forms.DockStyle]::Fill
    $script:gridLinks.ReadOnly = $true
    $script:gridLinks.RowHeadersVisible = $false
    $script:gridLinks.BackgroundColor = [System.Drawing.Color]::White
    $script:gridLinks.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $script:gridLinks.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(230, 236, 245)
    $script:gridLinks.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    $script:gridLinks.ColumnHeadersHeight = 34
    $script:gridLinks.RowTemplate.Height = 28
    $script:gridLinks.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(249, 251, 254)
    $script:gridLinks.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill

    $panelLinks.Controls.Add($script:gridLinks)
    $panelLinks.Controls.Add($script:lblLinks)
    $splitBackupRight.Panel1.Controls.Add($panelLinks)

    $panelLog = New-Object System.Windows.Forms.Panel
    $panelLog.Dock = [System.Windows.Forms.DockStyle]::Fill
    $panelLog.Padding = New-Object System.Windows.Forms.Padding(4, 4, 10, 10)

    $lblLog = New-Object System.Windows.Forms.Label
    $lblLog.Text = "3. Backup- & Aktivitaets-Protokoll:"
    $lblLog.Dock = [System.Windows.Forms.DockStyle]::Top
    $lblLog.Height = 28
    $lblLog.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)

    $script:txtLog = New-Object System.Windows.Forms.TextBox
    $script:txtLog.Dock = [System.Windows.Forms.DockStyle]::Fill
    $script:txtLog.Multiline = $true
    $script:txtLog.ReadOnly = $true
    $script:txtLog.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $script:txtLog.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 252)
    $script:txtLog.Font = New-Object System.Drawing.Font("Consolas", 9)

    $panelLog.Controls.Add($script:txtLog)
    $panelLog.Controls.Add($lblLog)
    $splitBackupRight.Panel2.Controls.Add($panelLog)

    $splitBackupMain.Panel2.Controls.Add($splitBackupRight)
    $tabBackup.Controls.Add($splitBackupMain)
    $tabBackup.Controls.Add($panelBackupTop)
    $tabControl.TabPages.Add($tabBackup)

    Enable-GridSorting -Grid $script:gridGpos

    $btnBrowseFolder.Add_Click({
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.SelectedPath = $script:txtBackupTargetDir.Text.Trim()
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $script:txtBackupTargetDir.Text = $dialog.SelectedPath
            if ($script:txtSettingsExportDir) { $script:txtSettingsExportDir.Text = $dialog.SelectedPath }
        }
    })

    # Abgesichertes Rendering: kein Aufruf von .Trim() auf $null
    $script:Update_BackupGridDisplay = {
        if ($script:isClosing -or $null -eq $script:gridGpos -or $script:gridGpos.IsDisposed) { return }
        $mode = if ($script:comboBackupFilter -and -not $script:comboBackupFilter.IsDisposed) { $script:comboBackupFilter.SelectedItem } else { "Alle GPOs" }
        $filterText = if ($script:txtBackupSearch -and -not $script:txtBackupSearch.IsDisposed) { "$($script:txtBackupSearch.Text)".Trim() } else { "" }

        $filtered = $script:rawBackupList | Where-Object {
            $matchMode = switch ($mode) { "Nur verlinkte GPOs" { $_."Link-Anzahl" -gt 0 } "Nicht verlinkte GPOs (Unlinked)" { $_."Link-Anzahl" -eq 0 } default { $true } }
            $matchSearch = if ([string]::IsNullOrWhiteSpace($filterText)) { $true } else { $_."GPO Name" -like "*$filterText*" -or $_."Status" -like "*$filterText*" }
            $matchMode -and $matchSearch
        }

        $arr = [System.Collections.ArrayList]::new()
        foreach ($it in $filtered) { [void]$arr.Add($it) }
        $script:gridGpos.DataSource = $arr
        if ($script:gridGpos.Columns["Link-Anzahl"]) { $script:gridGpos.Columns["Link-Anzahl"].Visible = $false }

        if ($script:btnBackupAll -and -not $script:btnBackupAll.IsDisposed) {
            $script:btnBackupAll.Text = if ($mode -ne "Alle GPOs" -or (-not [string]::IsNullOrWhiteSpace($filterText))) { "Gefilterte GPOs sichern ($($arr.Count))" } else { "ALLE GPOs sichern ($($arr.Count))" }
        }
        if ($script:btnExportAllHtml -and -not $script:btnExportAllHtml.IsDisposed) {
            $script:btnExportAllHtml.Text = if ($mode -ne "Alle GPOs" -or (-not [string]::IsNullOrWhiteSpace($filterText))) { "HTML (Gefilterte: $($arr.Count))" } else { "HTML (Alle: $($arr.Count))" }
        }
        if ($script:lblGpoGrid -and -not $script:lblGpoGrid.IsDisposed) {
            $script:lblGpoGrid.Text = "1. Gruppenrichtlinien der Domaene ($($arr.Count) angezeigt):"
        }
    }

    $script:comboBackupFilter.Add_SelectedIndexChanged({ & $script:Update_BackupGridDisplay })
    $script:txtBackupSearch.Add_TextChanged({ & $script:Update_BackupGridDisplay })

    $script:Invoke_LoadGpos = {
        if ($script:isClosing -or $null -eq $script:gridGpos -or $script:gridGpos.IsDisposed) { return }
        $script:rawBackupList.Clear()
        foreach ($g in $script:allGposCache) {
            $cleanGuid = if ($g.Id) { "$($g.Id)".Trim('{','}').ToUpper() } else { "" }
            $isLinked = ($script:gpoLinksCache.ContainsKey($cleanGuid) -and $script:gpoLinksCache[$cleanGuid].Count -gt 0)
            $linkCount = if ($isLinked) { $script:gpoLinksCache[$cleanGuid].Count } else { 0 }

            $script:rawBackupList.Add([PSCustomObject]@{
                "GPO Name"      = $g.DisplayName
                "Status"        = $g.GpoStatus.ToString()
                "Verknuepft"    = if ($linkCount -gt 0) { "JA ($linkCount)" } else { "NEIN (Unlinked)" }
                "Link-Anzahl"   = $linkCount
                "GPO ID (GUID)" = "$($g.Id)"
            })
        }
        & $script:Update_BackupGridDisplay
    }

    $btnLoadGpos.Add_Click({ & $script:Invoke_LoadGpos })

    $script:gridGpos.Add_SelectionChanged({
        if ($script:isClosing -or $null -eq $script:gridGpos -or $script:gridGpos.IsDisposed -or $null -eq $script:gridLinks -or $script:gridLinks.IsDisposed) { return }
        if ($script:gridGpos.SelectedRows.Count -gt 0) {
            $guid = $script:gridGpos.SelectedRows[0].Cells["GPO ID (GUID)"].Value
            $tableLinks = New-Object System.Data.DataTable
            [void]$tableLinks.Columns.AddRange(@("Verknuepfungs-Ziel (SOMPath)", "Aktiviert", "Enforced"))

            try {
                $report = [xml](Get-GPOReport -Guid $guid -ReportType Xml)
                foreach ($l in @($report.GPO.LinksTo | Where-Object { -not [string]::IsNullOrWhiteSpace($_.SOMPath) })) {
                    [void]$tableLinks.Rows.Add($l.SOMPath, $l.Enabled, $l.NoOverride)
                }
                if ($tableLinks.Rows.Count -eq 0) { [void]$tableLinks.Rows.Add("-- Keine Verknuepfung --", "-", "-") }
            } catch {
                [void]$tableLinks.Rows.Add("Fehler beim Abruf", "-", "-")
            }
            $script:gridLinks.DataSource = $tableLinks
        }
    })

    function Backup-SingleGPOWithLog ($guid, $name, $basePath) {
        $target = Join-Path (Join-Path $basePath ($name -replace '[\\/:*?"<>|]', '_')) (Get-Date -Format "yyyyMMdd_HHmm")
        if (-not (Test-Path $target)) { New-Item -ItemType Directory -Path $target -Force | Out-Null }
        Backup-GPO -Guid $guid -Path $target | Out-Null
        if ($script:txtLog -and -not $script:txtLog.IsDisposed) {
            $script:txtLog.AppendText("[$((Get-Date).ToString('HH:mm:ss'))] [OK] Gesichert: '$name' -> $target`r`n")
        }
    }

    $btnBackupSelected.Add_Click({
        if ($script:gridGpos.SelectedRows.Count -gt 0) {
            $g = $script:gridGpos.SelectedRows[0]
            Backup-SingleGPOWithLog $g.Cells["GPO ID (GUID)"].Value $g.Cells["GPO Name"].Value $script:txtBackupTargetDir.Text.Trim()
        }
    })

    $script:btnBackupAll.Add_Click({
        $items = @($script:gridGpos.DataSource)
        if ($items.Count -eq 0) { return }
        foreach ($g in $items) {
            try { Backup-SingleGPOWithLog $g."GPO ID (GUID)" $g."GPO Name" $script:txtBackupTargetDir.Text.Trim() } catch {}
        }
    })

    $btnExportSelectedHtml.Add_Click({
        if ($script:gridGpos.SelectedRows.Count -gt 0) {
            $g = $script:gridGpos.SelectedRows[0]
            $htmlDir = Join-Path $script:txtBackupTargetDir.Text.Trim() "HTML_Reports"
            if (-not (Test-Path $htmlDir)) { New-Item -ItemType Directory -Path $htmlDir -Force | Out-Null }
            $file = Join-Path $htmlDir "$($g.Cells['GPO Name'].Value -replace '[\\/:*?`"<>|]','_')_$((Get-Date).ToString('yyyyMMdd_HHmm')).html"
            Get-GPOReport -Guid $g.Cells["GPO ID (GUID)"].Value -ReportType Html -Path $file -ErrorAction Stop
            Start-Process $file
        }
    })

    $script:btnExportAllHtml.Add_Click({
        $items = @($script:gridGpos.DataSource)
        if ($items.Count -eq 0) { return }
        $htmlDir = Join-Path $script:txtBackupTargetDir.Text.Trim() "GPO_HTML_Reports_$((Get-Date).ToString('yyyyMMdd_HHmm'))"
        if (-not (Test-Path $htmlDir)) { New-Item -ItemType Directory -Path $htmlDir -Force | Out-Null }
        foreach ($it in $items) {
            try {
                $file = Join-Path $htmlDir "$($it.'GPO Name' -replace '[\\/:*?`"<>|]','_').html"
                Get-GPOReport -Guid $it."GPO ID (GUID)" -ReportType Html -Path $file -ErrorAction Stop
            } catch {}
        }
        Start-Process -FilePath "explorer.exe" -ArgumentList $htmlDir
    })
}