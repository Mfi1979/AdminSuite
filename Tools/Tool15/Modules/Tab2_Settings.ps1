# =========================================================================
# Tab2_Settings.ps1 - GPO Richtlinien-Inspektor (Lazy Loading)
# =========================================================================

function Build-Tab2_Settings {
    param($tabControl)

    $tabSettings = New-Object System.Windows.Forms.TabPage
    $tabSettings.Text = "2. GPO Richtlinien-Einstellungen & Inspektor"
    $tabSettings.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)

    $panelSettingsTop = New-Object System.Windows.Forms.Panel
    $panelSettingsTop.Dock = [System.Windows.Forms.DockStyle]::Top
    $panelSettingsTop.Height = 110
    $panelSettingsTop.BackColor = [System.Drawing.Color]::FromArgb(242, 245, 250)
    $panelSettingsTop.Padding = New-Object System.Windows.Forms.Padding(10)

    $lblSettingsTargetPath = New-Object System.Windows.Forms.Label
    $lblSettingsTargetPath.Text = "Export Ziel-Pfad:"
    $lblSettingsTargetPath.Location = New-Object System.Drawing.Point(12, 17)
    $lblSettingsTargetPath.AutoSize = $true
    $lblSettingsTargetPath.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $txtSettingsExportDir = New-Object System.Windows.Forms.TextBox
    $txtSettingsExportDir.Location = New-Object System.Drawing.Point(135, 14)
    $txtSettingsExportDir.Size = New-Object System.Drawing.Size(440, 25)
    $txtSettingsExportDir.Text = "C:\Install\Backup\GPO"
    $script:txtSettingsExportDir = $txtSettingsExportDir

    $btnBrowseSettingsDir = New-Object System.Windows.Forms.Button
    $btnBrowseSettingsDir.Text = "Durchsuchen..."
    $btnBrowseSettingsDir.Location = New-Object System.Drawing.Point(585, 11)
    $btnBrowseSettingsDir.Size = New-Object System.Drawing.Size(115, 30)

    $btnExportCsv = New-Object System.Windows.Forms.Button
    $btnExportCsv.Text = "CSV Export"
    $btnExportCsv.Location = New-Object System.Drawing.Point(710, 11)
    $btnExportCsv.Size = New-Object System.Drawing.Size(110, 30)
    $btnExportCsv.BackColor = [System.Drawing.Color]::FromArgb(230, 245, 230)

    $script:lblSettingsStatus = New-Object System.Windows.Forms.Label
    $script:lblSettingsStatus.Text = "Bereit."
    $script:lblSettingsStatus.Location = New-Object System.Drawing.Point(835, 17)
    $script:lblSettingsStatus.AutoSize = $true
    $script:lblSettingsStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)

    $lblSettingsViewFilter = New-Object System.Windows.Forms.Label
    $lblSettingsViewFilter.Text = "Ansicht:"
    $lblSettingsViewFilter.Location = New-Object System.Drawing.Point(12, 58)
    $lblSettingsViewFilter.AutoSize = $true
    $lblSettingsViewFilter.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    # 1. ViewMode Dropdown sauber im Script-Scope
    $script:comboSettingsViewMode = New-Object System.Windows.Forms.ComboBox
    $script:comboSettingsViewMode.Location = New-Object System.Drawing.Point(75, 55)
    $script:comboSettingsViewMode.Size = New-Object System.Drawing.Size(210, 25)
    $script:comboSettingsViewMode.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$script:comboSettingsViewMode.Items.AddRange(@("Alle GPOs", "Nur verlinkte GPOs", "Nicht verlinkte GPOs (Unlinked)"))
    $script:comboSettingsViewMode.SelectedIndex = 0

    $lblGpo = New-Object System.Windows.Forms.Label
    $lblGpo.Text = "GPO:"
    $lblGpo.Location = New-Object System.Drawing.Point(295, 58)
    $lblGpo.AutoSize = $true
    $lblGpo.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    # 2. GPO Dropdown sauber im Script-Scope
    $script:comboGpoSettings = New-Object System.Windows.Forms.ComboBox
    $script:comboGpoSettings.Location = New-Object System.Drawing.Point(340, 55)
    $script:comboGpoSettings.Size = New-Object System.Drawing.Size(360, 25)
    $script:comboGpoSettings.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

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

    $script:txtFilterSettings = New-Object System.Windows.Forms.TextBox
    $script:txtFilterSettings.Location = New-Object System.Drawing.Point(885, 55)
    $script:txtFilterSettings.Size = New-Object System.Drawing.Size(200, 25)

    $panelSettingsTop.Controls.AddRange(@(
        $lblSettingsTargetPath, $txtSettingsExportDir, $btnBrowseSettingsDir, $btnExportCsv, 
        $script:lblSettingsStatus, $lblSettingsViewFilter, $script:comboSettingsViewMode, 
        $lblGpo, $script:comboGpoSettings, $btnLoadSettings, $lblFilter, $script:txtFilterSettings
    ))

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

    $script:gridSettings = New-Object System.Windows.Forms.DataGridView
    $script:gridSettings.Dock = [System.Windows.Forms.DockStyle]::Fill
    $script:gridSettings.ReadOnly = $true
    $script:gridSettings.AllowUserToAddRows = $false
    $script:gridSettings.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $script:gridSettings.MultiSelect = $false
    $script:gridSettings.RowHeadersVisible = $false
    $script:gridSettings.BackgroundColor = [System.Drawing.Color]::White
    $script:gridSettings.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $script:gridSettings.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(230, 236, 245)
    $script:gridSettings.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    $script:gridSettings.ColumnHeadersHeight = 34
    $script:gridSettings.RowTemplate.Height = 28
    $script:gridSettings.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(249, 251, 254)

    [void]$script:gridSettings.Columns.Add("colScope", "Bereich")
    [void]$script:gridSettings.Columns.Add("colCategory", "Kategorie / Pfad")
    [void]$script:gridSettings.Columns.Add("colName", "Einstellung (Name)")
    [void]$script:gridSettings.Columns.Add("colValue", "Konfigurierter Wert")
    [void]$script:gridSettings.Columns.Add("colState", "Status")
    [void]$script:gridSettings.Columns.Add("colSupported", "Unterstuetzt ab")
    [void]$script:gridSettings.Columns.Add("colExplain", "Erklaerung")
    [void]$script:gridSettings.Columns.Add("colGpo", "GPO")

    $script:gridSettings.Columns["colScope"].Width = 90
    $script:gridSettings.Columns["colCategory"].Width = 240
    $script:gridSettings.Columns["colName"].Width = 300
    $script:gridSettings.Columns["colValue"].AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
    $script:gridSettings.Columns["colState"].Visible = $false
    $script:gridSettings.Columns["colSupported"].Visible = $false
    $script:gridSettings.Columns["colExplain"].Visible = $false
    $script:gridSettings.Columns["colGpo"].Visible = $false

    $panelSettingsLeft.Controls.Add($script:gridSettings)
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

    $script:txtDescription = New-Object System.Windows.Forms.TextBox
    $script:txtDescription.Dock = [System.Windows.Forms.DockStyle]::Fill
    $script:txtDescription.Multiline = $true
    $script:txtDescription.ReadOnly = $true
    $script:txtDescription.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $script:txtDescription.BackColor = [System.Drawing.Color]::FromArgb(252, 252, 254)
    $script:txtDescription.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)

    $panelSettingsRight.Controls.Add($script:txtDescription)
    $panelSettingsRight.Controls.Add($lblDescHeader)
    $splitSettings.Panel2.Controls.Add($panelSettingsRight)

    $tabSettings.Controls.Add($splitSettings)
    $tabSettings.Controls.Add($panelSettingsTop)
    $tabControl.TabPages.Add($tabSettings)

    $btnBrowseSettingsDir.Add_Click({
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.SelectedPath = $txtSettingsExportDir.Text.Trim()
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtSettingsExportDir.Text = $dialog.SelectedPath
            if ($script:txtBackupTargetDir) { $script:txtBackupTargetDir.Text = $dialog.SelectedPath }
        }
    })

    # Dropdown-Update mit Absicherung gegen Null-Referenzen
    $script:Update_SettingsGpoDropdown = {
        if ($script:isClosing) { return }
        if ($null -eq $script:comboSettingsViewMode -or $script:comboSettingsViewMode.IsDisposed) { return }
        if ($null -eq $script:comboGpoSettings -or $script:comboGpoSettings.IsDisposed) { return }

        $mode = $script:comboSettingsViewMode.SelectedItem

        $matchingGpos = $script:allGposCache | Where-Object {
            $cleanGuid = $_.Id.ToString().Trim('{','}').ToUpper()
            $isLinked = ($script:gpoLinksCache.ContainsKey($cleanGuid) -and $script:gpoLinksCache[$cleanGuid].Count -gt 0)
            switch ($mode) {
                "Nur verlinkte GPOs"              { $isLinked }
                "Nicht verlinkte GPOs (Unlinked)" { -not $isLinked }
                default                           { $true }
            }
        }

        $script:comboGpoSettings.Items.Clear()

        $headerText = switch ($mode) {
            "Nur verlinkte GPOs"              { "-- ALLE verlinkten GPOs laden ($($matchingGpos.Count)) --" }
            "Nicht verlinkte GPOs (Unlinked)" { "-- ALLE ungelinkten GPOs laden ($($matchingGpos.Count)) --" }
            default                           { "-- ALLE GPOs laden ($($matchingGpos.Count)) --" }
        }
        [void]$script:comboGpoSettings.Items.Add($headerText)

        foreach ($g in $matchingGpos) { 
            [void]$script:comboGpoSettings.Items.Add($g.DisplayName) 
        }

        if ($script:comboGpoSettings.Items.Count -gt 1) {
            $defIdx = -1
            for ($i = 1; $i -lt $script:comboGpoSettings.Items.Count; $i++) {
                if ($script:comboGpoSettings.Items[$i] -eq "Default Domain Policy") { $defIdx = $i; break }
            }
            $script:comboGpoSettings.SelectedIndex = if ($defIdx -gt 0) { $defIdx } else { 1 }
        } elseif ($script:comboGpoSettings.Items.Count -gt 0) { 
            $script:comboGpoSettings.SelectedIndex = 0 
        }
    }

    $script:comboSettingsViewMode.Add_SelectedIndexChanged({ & $script:Update_SettingsGpoDropdown })

    # Tabellen-Update mit Absicherung gegen Null-Referenzen
    $script:Update_SettingsGridDisplay = {
        if ($script:isClosing) { return }
        if ($null -eq $script:gridSettings -or $script:gridSettings.IsDisposed) { return }

        $filterText = if ($script:txtFilterSettings -and -not $script:txtFilterSettings.IsDisposed) { $script:txtFilterSettings.Text.Trim() } else { "" }
        $script:gridSettings.Rows.Clear()

        $filtered = if ([string]::IsNullOrWhiteSpace($filterText)) { $script:rawSettingsList } else {
            $script:rawSettingsList | Where-Object { $_.Name -like "*$filterText*" -or $_.Category -like "*$filterText*" -or $_.Value -like "*$filterText*" }
        }

        foreach ($it in @($filtered)) {
            [void]$script:gridSettings.Rows.Add($it.Scope, $it.Category, $it.Name, $it.Value, $it.State, $it.Supported, $it.Explain, $it.GpoName)
        }
        if ($script:lblSettingsStatus -and -not $script:lblSettingsStatus.IsDisposed) {
            $script:lblSettingsStatus.Text = "$(@($filtered).Count) Einstellung(en) geladen."
        }
        if ($script:gridSettings.Rows.Count -gt 0) { $script:gridSettings.Rows[0].Selected = $true }
    }

    $script:Invoke_LoadSettings = {
        if ($script:isClosing -or $null -eq $script:comboGpoSettings -or $script:comboGpoSettings.IsDisposed) { return }

        $selectedOption = $script:comboGpoSettings.SelectedItem
        if ([string]::IsNullOrWhiteSpace($selectedOption)) { return }

        if ($script:lblSettingsStatus -and -not $script:lblSettingsStatus.IsDisposed) { 
            $script:lblSettingsStatus.Text = "Lese Richtlinien..." 
        }
        $script:rawSettingsList.Clear()

        if ($selectedOption.StartsWith("-- ALLE")) {
            $mode = if ($script:comboSettingsViewMode) { $script:comboSettingsViewMode.SelectedItem } else { "Alle GPOs" }
            $targets = @($script:allGposCache | Where-Object {
                $cleanGuid = $_.Id.ToString().Trim('{','}').ToUpper()
                $isLinked = ($script:gpoLinksCache.ContainsKey($cleanGuid) -and $script:gpoLinksCache[$cleanGuid].Count -gt 0)
                switch ($mode) { "Nur verlinkte GPOs" { $isLinked } "Nicht verlinkte GPOs (Unlinked)" { -not $isLinked } default { $true } }
            })

            if ($script:pbarGlobal) {
                $script:pbarGlobal.Visible = $true
                $script:pbarGlobal.Maximum = [Math]::Max(1, $targets.Count)
            }
            $c = 0
            foreach ($g in $targets) {
                $c++
                if ($script:pbarGlobal) { $script:pbarGlobal.Value = $c }
                if ($script:lblProgressInfo) { $script:lblProgressInfo.Text = "Lese GPO ($c / $($targets.Count)): $($g.DisplayName)" }
                [System.Windows.Forms.Application]::DoEvents()
                try {
                    $items = Get-ParsedGpoSettings -GpoId $g.Id.ToString() -GpoDisplayName $g.DisplayName
                    foreach ($it in $items) { $script:rawSettingsList.Add($it) }
                } catch {}
            }
            if ($script:pbarGlobal) { $script:pbarGlobal.Visible = $false }
        } else {
            $gpo = $script:allGposCache | Where-Object { $_.DisplayName -eq $selectedOption } | Select-Object -First 1
            if ($gpo) {
                try {
                    $items = Get-ParsedGpoSettings -GpoId $gpo.Id.ToString() -GpoDisplayName $gpo.DisplayName
                    foreach ($it in $items) { $script:rawSettingsList.Add($it) }
                } catch {}
            }
        }
        & $script:Update_SettingsGridDisplay
    }

    $btnLoadSettings.Add_Click({ & $script:Invoke_LoadSettings })
    $script:txtFilterSettings.Add_TextChanged({ & $script:Update_SettingsGridDisplay })

    $script:gridSettings.Add_SelectionChanged({
        if ($script:isClosing -or $null -eq $script:gridSettings -or $script:gridSettings.IsDisposed) { return }
        if ($script:gridSettings.SelectedRows.Count -gt 0) {
            $row = $script:gridSettings.SelectedRows[0]
            if ($script:txtDescription -and -not $script:txtDescription.IsDisposed) {
                $script:txtDescription.Text = "RICHTLINIE : $($row.Cells['colName'].Value)`r`nGPO        : $($row.Cells['colGpo'].Value)`r`nBEREICH    : $($row.Cells['colScope'].Value)`r`nPFAD       : $($row.Cells['colCategory'].Value)`r`nWERT       : $($row.Cells['colValue'].Value)`r`n`r`n[ERLAEUTERUNG]`r`n$($row.Cells['colExplain'].Value)"
            }
        }
    })

    $btnExportCsv.Add_Click({
        if ($script:rawSettingsList.Count -eq 0) { return }
        $targetBase = $txtSettingsExportDir.Text.Trim()
        if (-not (Test-Path $targetBase)) { New-Item -ItemType Directory -Path $targetBase -Force | Out-Null }
        $csvFile = Join-Path $targetBase "GPO_Settings_Export_$((Get-Date).ToString('yyyyMMdd_HHmm')).csv"
        $script:rawSettingsList | Export-Csv -Path $csvFile -Delimiter ";" -NoTypeInformation -Encoding UTF8
        [System.Windows.Forms.MessageBox]::Show("Exportiert nach: $csvFile", "Erfolg", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    })
}