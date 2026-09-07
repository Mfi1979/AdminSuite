# =========================================================================
# Tab0_Dashboard.ps1 - Master-Detail Status-Ueberblick
# =========================================================================

function Build-Tab0_Dashboard {
    param($tabControl)

    $tabDashboard = New-Object System.Windows.Forms.TabPage
    $tabDashboard.Text = "0. Status Ueberblick"
    $tabDashboard.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
    $tabDashboard.BackColor = [System.Drawing.Color]::FromArgb(246, 248, 252)

    $splitDashboard = New-Object System.Windows.Forms.SplitContainer
    $splitDashboard.Dock = [System.Windows.Forms.DockStyle]::Fill
    $splitDashboard.SplitterDistance = 680
    $splitDashboard.SplitterWidth = 6

    # Links: Master-Tabelle
    $pnlDashLeft = New-Object System.Windows.Forms.Panel
    $pnlDashLeft.Dock = [System.Windows.Forms.DockStyle]::Fill
    $pnlDashLeft.Padding = New-Object System.Windows.Forms.Padding(10, 10, 5, 10)

    $lblDashTableTitle = New-Object System.Windows.Forms.Label
    $lblDashTableTitle.Text = "Status-Pruefungen (Klick auf eine Zeile zeigt rechts Details):"
    $lblDashTableTitle.Dock = [System.Windows.Forms.DockStyle]::Top
    $lblDashTableTitle.Height = 28
    $lblDashTableTitle.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)

    $script:gridDashMaster = New-Object System.Windows.Forms.DataGridView
    $script:gridDashMaster.Dock = [System.Windows.Forms.DockStyle]::Fill
    $script:gridDashMaster.ReadOnly = $true
    $script:gridDashMaster.AllowUserToAddRows = $false
    $script:gridDashMaster.AllowUserToDeleteRows = $false
    $script:gridDashMaster.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $script:gridDashMaster.MultiSelect = $false
    $script:gridDashMaster.RowHeadersVisible = $false
    $script:gridDashMaster.ColumnHeadersVisible = $true
    $script:gridDashMaster.EnableHeadersVisualStyles = $false
    $script:gridDashMaster.BackgroundColor = [System.Drawing.Color]::White
    $script:gridDashMaster.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $script:gridDashMaster.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(230, 236, 245)
    $script:gridDashMaster.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    $script:gridDashMaster.ColumnHeadersHeight = 34
    $script:gridDashMaster.RowTemplate.Height = 34
    $script:gridDashMaster.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill

    $pnlDashLeft.Controls.Add($script:gridDashMaster)
    $pnlDashLeft.Controls.Add($lblDashTableTitle)
    $splitDashboard.Panel1.Controls.Add($pnlDashLeft)

    # Rechts: Detailansicht
    $pnlDashRight = New-Object System.Windows.Forms.Panel
    $pnlDashRight.Dock = [System.Windows.Forms.DockStyle]::Fill
    $pnlDashRight.Padding = New-Object System.Windows.Forms.Padding(5, 10, 10, 10)

    $lblDashDetailTitle = New-Object System.Windows.Forms.Label
    $lblDashDetailTitle.Text = "Diagnose- & Detailansicht:"
    $lblDashDetailTitle.Dock = [System.Windows.Forms.DockStyle]::Top
    $lblDashDetailTitle.Height = 28
    $lblDashDetailTitle.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)

    $script:pnlDashBadge = New-Object System.Windows.Forms.Panel
    $script:pnlDashBadge.Dock = [System.Windows.Forms.DockStyle]::Top
    $script:pnlDashBadge.Height = 42
    $script:pnlDashBadge.BackColor = [System.Drawing.Color]::FromArgb(235, 247, 235)
    $script:pnlDashBadge.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

    $script:lblDashBadgeText = New-Object System.Windows.Forms.Label
    $script:lblDashBadgeText.Dock = [System.Windows.Forms.DockStyle]::Fill
    $script:lblDashBadgeText.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $script:lblDashBadgeText.Text = "Waehlen Sie links eine Pruefung aus."
    $script:lblDashBadgeText.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $script:lblDashBadgeText.ForeColor = [System.Drawing.Color]::DarkGreen
    $script:pnlDashBadge.Controls.Add($script:lblDashBadgeText)

    $script:pnlDashActions = New-Object System.Windows.Forms.Panel
    $script:pnlDashActions.Dock = [System.Windows.Forms.DockStyle]::Top
    $script:pnlDashActions.Height = 44
    $script:pnlDashActions.BackColor = [System.Drawing.Color]::Transparent
    $script:pnlDashActions.Visible = $false

    $btnExportUnlinkedCsv = New-Object System.Windows.Forms.Button
    $btnExportUnlinkedCsv.Text = "Unlinked CSV Export"
    $btnExportUnlinkedCsv.Location = New-Object System.Drawing.Point(0, 6)
    $btnExportUnlinkedCsv.Size = New-Object System.Drawing.Size(160, 32)
    $btnExportUnlinkedCsv.BackColor = [System.Drawing.Color]::FromArgb(230, 245, 230)

    $btnSwitchToBackup = New-Object System.Windows.Forms.Button
    $btnSwitchToBackup.Text = "Zu Register 3 (Backup / HTML) wechseln"
    $btnSwitchToBackup.Location = New-Object System.Drawing.Point(170, 6)
    $btnSwitchToBackup.Size = New-Object System.Drawing.Size(260, 32)
    $btnSwitchToBackup.BackColor = [System.Drawing.Color]::FromArgb(255, 245, 230)

    $script:pnlDashActions.Controls.Add($btnExportUnlinkedCsv)
    $script:pnlDashActions.Controls.Add($btnSwitchToBackup)

    $script:txtDashDetailText = New-Object System.Windows.Forms.TextBox
    $script:txtDashDetailText.Dock = [System.Windows.Forms.DockStyle]::Fill
    $script:txtDashDetailText.Multiline = $true
    $script:txtDashDetailText.ReadOnly = $true
    $script:txtDashDetailText.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $script:txtDashDetailText.BackColor = [System.Drawing.Color]::White
    $script:txtDashDetailText.Font = New-Object System.Drawing.Font("Consolas", 9.5)

    $pnlDashRight.Controls.Add($script:txtDashDetailText)
    $pnlDashRight.Controls.Add($script:pnlDashActions)
    $pnlDashRight.Controls.Add($script:pnlDashBadge)
    $pnlDashRight.Controls.Add($lblDashDetailTitle)
    $splitDashboard.Panel2.Controls.Add($pnlDashRight)

    $tabDashboard.Controls.Add($splitDashboard)
    $tabControl.TabPages.Add($tabDashboard)

    # Dashboard-Logik
    $script:Invoke_UpdateDashboard = {
        if ($script:isClosing -or $null -eq $script:gridDashMaster -or $script:gridDashMaster.IsDisposed) { return }

        $ddpReal = $script:rawOverviewList | Where-Object { [string]$_."GUID" -match $script:StandardDdpGuid } | Select-Object -First 1
        $ddpByName = @($script:rawOverviewList | Where-Object { $_."GPO Name" -like "*Default Domain Policy*" })

        $ddpLines = [System.Collections.Generic.List[string]]::new()
        $ddpIsHealthy = $true

        if ($ddpReal) {
            $ddpLines.Add("Original Microsoft Standard-GUID vorhanden:")
            $ddpLines.Add(" - Aktueller Name    : '$($ddpReal.'GPO Name')'")
            $ddpLines.Add(" - Standard-GUID     : {$script:StandardDdpGuid}")
            $ddpLines.Add(" - GPO-Status        : $($ddpReal.'Gesamt-Status')")
            $ddpLines.Add(" - Verlinkt          : $($ddpReal.'Verlinkt') ($($ddpReal.'Link-Anzahl') Ziel(e))")

            if ($ddpReal."GPO Name" -ne "Default Domain Policy") {
                $ddpIsHealthy = $false
                $ddpLines.Add(" ! HINWEIS: Richtlinie wurde umbenannt (Original: 'Default Domain Policy')")
            }
            if ($ddpReal."Verlinkt" -ne "Ja") {
                $ddpIsHealthy = $false
                $ddpLines.Add(" ! KRITISCH: Die echte DDP ist aktuell NICHT verlinkt!")
            }
        } else {
            $ddpIsHealthy = $false
            $ddpLines.Add("! KRITISCH: Keine GPO mit der Standard-DDP-GUID {$script:StandardDdpGuid} gefunden!")
        }

        if ($ddpByName.Count -gt 1 -or ($ddpByName.Count -eq 1 -and -not $ddpReal)) {
            $ddpIsHealthy = $false
            $ddpLines.Add("")
            $ddpLines.Add("! GEFUNDENE NAMENSDUPLIKATE MIT FREMDER GUID:")
            foreach ($dup in $ddpByName) {
                if ($dup.GUID -notmatch $script:StandardDdpGuid) {
                    $ddpLines.Add(" - '$($dup.'GPO Name')' -> GUID: $($dup.GUID)")
                }
            }
        }

        $script:dashDetailCache["DDP"] = @{
            Healthy    = $ddpIsHealthy
            Badge      = if ($ddpIsHealthy) { "[OK] Default Domain Policy ist sauber konfiguriert" } else { "[WARNUNG] Unregelmaessigkeiten bei der Default Domain Policy!" }
            Text       = $ddpLines -join "`r`n"
            IsUnlinked = $false
        }

        # DDCP Check
        $ddcpReal = $script:rawOverviewList | Where-Object { [string]$_."GUID" -match $script:StandardDdcpGuid } | Select-Object -First 1
        $ddcpByName = @($script:rawOverviewList | Where-Object { $_."GPO Name" -like "*Default Domain Controllers Policy*" })

        $ddcpLines = [System.Collections.Generic.List[string]]::new()
        $ddcpIsHealthy = $true

        if ($ddcpReal) {
            $ddcpLines.Add("Original Microsoft Standard-GUID vorhanden:")
            $ddcpLines.Add(" - Aktueller Name    : '$($ddcpReal.'GPO Name')'")
            $ddcpLines.Add(" - Standard-GUID     : {$script:StandardDdcpGuid}")
            $ddcpLines.Add(" - GPO-Status        : $($ddcpReal.'Gesamt-Status')")
            $ddcpLines.Add(" - Verlinkt          : $($ddcpReal.'Verlinkt') ($($ddcpReal.'Link-Anzahl') Ziel(e))")

            if ($ddcpReal."GPO Name" -ne "Default Domain Controllers Policy") {
                $ddcpIsHealthy = $false
                $ddcpLines.Add(" ! HINWEIS: Richtlinie wurde umbenannt (Original: 'Default Domain Controllers Policy')")
            }
            if ($ddcpReal."Verlinkt" -ne "Ja") {
                $ddcpIsHealthy = $false
                $ddcpLines.Add(" ! KRITISCH: Die echte DDCP ist aktuell NICHT verlinkt!")
            }
        } else {
            $ddcpIsHealthy = $false
            $ddcpLines.Add("! KRITISCH: Keine GPO mit der Standard-DDCP-GUID {$script:StandardDdcpGuid} gefunden!")
        }

        if ($ddcpByName.Count -gt 1 -or ($ddcpByName.Count -eq 1 -and -not $ddcpReal)) {
            $ddcpIsHealthy = $false
            $ddcpLines.Add("")
            $ddcpLines.Add("! GEFUNDENE NAMENSDUPLIKATE MIT FREMDER GUID:")
            foreach ($dup in $ddcpByName) {
                if ($dup.GUID -notmatch $script:StandardDdcpGuid) {
                    $ddcpLines.Add(" - '$($dup.'GPO Name')' -> GUID: $($dup.GUID)")
                }
            }
        }

        $script:dashDetailCache["DDCP"] = @{
            Healthy    = $ddcpIsHealthy
            Badge      = if ($ddcpIsHealthy) { "[OK] Default Domain Controllers Policy ist sauber konfiguriert" } else { "[WARNUNG] Unregelmaessigkeiten bei der DDCP!" }
            Text       = $ddcpLines -join "`r`n"
            IsUnlinked = $false
        }

        # Unlinked GPOs
        $unlinkedItems = @($script:rawOverviewList | Where-Object { $_."Verlinkt" -eq "Nein" })
        $unlinkedLines = [System.Collections.Generic.List[string]]::new()
        $unlinkedLines.Add("BESTANDSAUFNAHME DER NICHT VERLINKTEN RICHTLINIEN:")
        $unlinkedLines.Add("Aktuell sind $($unlinkedItems.Count) von $($script:rawOverviewList.Count) Gruppenrichtlinien nirgendwo verlinkt (Unlinked).`r`n")
        foreach ($u in $unlinkedItems) { $unlinkedLines.Add(" - $($u.'GPO Name') [$($u.'Gesamt-Status')] -> GUID: $($u.GUID)") }

        $script:dashDetailCache["UNLINKED"] = @{
            Healthy    = ($unlinkedItems.Count -eq 0)
            Badge      = if ($unlinkedItems.Count -eq 0) { "[OK] Keine ungelinkten GPOs vorhanden" } else { "[HINWEIS] $($unlinkedItems.Count) nicht verlinkte GPOs gefunden" }
            Text       = $unlinkedLines -join "`r`n"
            IsUnlinked = $true
        }

        $tableData = [System.Collections.ArrayList]::new()
        [void]$tableData.Add([PSCustomObject]@{
            "Key"                  = "DDP"
            "Komponente / Bereich" = "1. Default Domain Policy (DDP)"
            "Status"               = if ($ddpIsHealthy) { "OK (Gueltig)" } else { "WARNUNG (Konflikt)" }
            "Befund / Details"     = if ($ddpReal) { "$($ddpReal.'GPO Name') (Verlinkt: $($ddpReal.'Verlinkt'))" } else { "GUID fehlt!" }
        })
        [void]$tableData.Add([PSCustomObject]@{
            "Key"                  = "DDCP"
            "Komponente / Bereich" = "2. Default Domain Controllers Policy (DDCP)"
            "Status"               = if ($ddcpIsHealthy) { "OK (Gueltig)" } else { "WARNUNG (Konflikt)" }
            "Befund / Details"     = if ($ddcpReal) { "$($ddcpReal.'GPO Name') (Verlinkt: $($ddcpReal.'Verlinkt'))" } else { "GUID fehlt!" }
        })
        [void]$tableData.Add([PSCustomObject]@{
            "Key"                  = "UNLINKED"
            "Komponente / Bereich" = "3. Nicht verlinkte GPOs (Unlinked)"
            "Status"               = if ($unlinkedItems.Count -eq 0) { "OK (0 GPOs)" } else { "Hinweis ($($unlinkedItems.Count) GPOs)" }
            "Befund / Details"     = "$($unlinkedItems.Count) von $($script:rawOverviewList.Count) GPOs sind nicht verlinkt"
        })

        $script:gridDashMaster.DataSource = $tableData
        if ($script:gridDashMaster.Columns["Key"]) { $script:gridDashMaster.Columns["Key"].Visible = $false }
        if ($script:gridDashMaster.Rows.Count -gt 0) { $script:gridDashMaster.Rows[0].Selected = $true }
    }

    $script:gridDashMaster.Add_DataBindingComplete({
        if ($script:isClosing -or $null -eq $script:gridDashMaster -or $script:gridDashMaster.IsDisposed) { return }
        foreach ($row in $script:gridDashMaster.Rows) {
            $status = [string]$row.Cells["Status"].Value
            if ($status -match "^WARNUNG") {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 238, 204)
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(180, 50, 0)
                $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(255, 210, 160)
                $row.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
                $row.DefaultCellStyle.Font = New-Object System.Drawing.Font($script:gridDashMaster.Font, [System.Drawing.FontStyle]::Bold)
            } elseif ($status -match "^Hinweis") {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 248, 225)
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(160, 90, 0)
                $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(255, 230, 170)
                $row.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
            } else {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(235, 247, 235)
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkGreen
                $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(200, 235, 200)
                $row.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
            }
        }
    })

    $script:gridDashMaster.Add_SelectionChanged({
        if ($script:isClosing -or $null -eq $script:gridDashMaster -or $script:gridDashMaster.IsDisposed) { return }
        if ($script:gridDashMaster.SelectedRows.Count -gt 0) {
            $key = [string]$script:gridDashMaster.SelectedRows[0].Cells["Key"].Value
            if ($script:dashDetailCache.ContainsKey($key)) {
                $entry = $script:dashDetailCache[$key]
                if ($script:lblDashBadgeText -and -not $script:lblDashBadgeText.IsDisposed) {
                    $script:lblDashBadgeText.Text = $entry.Badge
                }
                if ($script:txtDashDetailText -and -not $script:txtDashDetailText.IsDisposed) {
                    $script:txtDashDetailText.Text = $entry.Text
                }
                if ($script:pnlDashActions -and -not $script:pnlDashActions.IsDisposed) {
                    $script:pnlDashActions.Visible = $entry.IsUnlinked
                }
                if ($script:pnlDashBadge -and -not $script:pnlDashBadge.IsDisposed) {
                    if ($entry.Healthy) {
                        $script:pnlDashBadge.BackColor = [System.Drawing.Color]::FromArgb(235, 247, 235)
                        if ($script:lblDashBadgeText) { $script:lblDashBadgeText.ForeColor = [System.Drawing.Color]::DarkGreen }
                    } else {
                        $script:pnlDashBadge.BackColor = [System.Drawing.Color]::FromArgb(255, 238, 204)
                        if ($script:lblDashBadgeText) { $script:lblDashBadgeText.ForeColor = [System.Drawing.Color]::FromArgb(180, 50, 0) }
                    }
                }
            }
        }
    })

    $btnSwitchToBackup.Add_Click({
        if ($script:comboBackupFilter) { $script:comboBackupFilter.SelectedItem = "Nicht verlinkte GPOs (Unlinked)" }
        $tabControl.SelectedTab = $script:tabBackup
    })

    $btnExportUnlinkedCsv.Add_Click({
        $unlinkedItems = @($script:rawOverviewList | Where-Object { $_."Verlinkt" -eq "Nein" })
        if ($unlinkedItems.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Keine ungelinkten GPOs vorhanden.", "Hinweis", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }

        $targetBase = if ($script:txtBackupTargetDir) { $script:txtBackupTargetDir.Text.Trim() } else { "C:\Install\Backup\GPO" }
        if (-not (Test-Path $targetBase)) { New-Item -ItemType Directory -Path $targetBase -Force | Out-Null }
        $csvFile = Join-Path $targetBase "GPO_Unlinked_Dashboard_$((Get-Date).ToString('yyyyMMdd_HHmm')).csv"
        $unlinkedItems | Export-Csv -Path $csvFile -Delimiter ";" -NoTypeInformation -Encoding UTF8
        [System.Windows.Forms.MessageBox]::Show("Exportiert nach: $csvFile", "Export fertig", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    })
}