# =========================================================================
# Tab4_Compare.ps1 - GPO Vergleichs-Engine & Baseline-Diff
# =========================================================================

# Hilfsfunktion: Zweisprachiger Zahlen- & Statusvergleich (DE / EN)
function Test-PolicyValueMatch ($v1, $v2) {
    if ([string]::IsNullOrWhiteSpace($v1) -and [string]::IsNullOrWhiteSpace($v2)) { return $true }
    $s1 = "$v1".Trim().ToLower()
    $s2 = "$v2".Trim().ToLower()

    if ($s1 -eq $s2) { return $true }

    # Boolean & Status (DE / EN)
    if (($s1 -match "^(enabled|aktiviert)") -and ($s2 -match "^(enabled|aktiviert)")) { return $true }
    if (($s1 -match "^(disabled|deaktiviert)") -and ($s2 -match "^(disabled|deaktiviert)")) { return $true }
    if (($s1 -match "not defined|nicht definiert") -and ($s2 -match "not defined|nicht definiert")) { return $true }

    # Numerische Werte mit Einheiten vergleichen (z.B. "24 passwords remembered" vs "24 gespeicherte kennwoerter")
    $m1 = [regex]::Match($s1, "\b(\d+)\b")
    $m2 = [regex]::Match($s2, "\b(\d+)\b")
    if ($m1.Success -and $m2.Success) {
        if ($m1.Groups[1].Value -eq $m2.Groups[1].Value) {
            $u1 = if ($s1 -match "day|tag") { "d" } elseif ($s1 -match "hour|stund") { "h" } elseif ($s1 -match "min") { "m" } elseif ($s1 -match "passw|kennw") { "p" } elseif ($s1 -match "char|zeich") { "c" } else { "" }
            $u2 = if ($s2 -match "day|tag") { "d" } elseif ($s2 -match "hour|stund") { "h" } elseif ($s2 -match "min") { "m" } elseif ($s2 -match "passw|kennw") { "p" } elseif ($s2 -match "char|zeich") { "c" } else { "" }
            if ($u1 -eq $u2) { return $true }
        }
    }

    return $false
}

function Build-Tab4_Compare {
    param($tabControl)

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

    $script:comboCompareGpo1 = New-Object System.Windows.Forms.ComboBox
    $script:comboCompareGpo1.Location = New-Object System.Drawing.Point(115, 13)
    $script:comboCompareGpo1.Size = New-Object System.Drawing.Size(280, 25)
    $script:comboCompareGpo1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

    $lblGpo2 = New-Object System.Windows.Forms.Label
    $lblGpo2.Text = "GPO 2 (Vergleich):"
    $lblGpo2.Location = New-Object System.Drawing.Point(410, 16)
    $lblGpo2.AutoSize = $true
    $lblGpo2.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $script:comboCompareGpo2 = New-Object System.Windows.Forms.ComboBox
    $script:comboCompareGpo2.Location = New-Object System.Drawing.Point(540, 13)
    $script:comboCompareGpo2.Size = New-Object System.Drawing.Size(290, 25)
    $script:comboCompareGpo2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

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

    $script:chkOnlyDiffs = New-Object System.Windows.Forms.CheckBox
    $script:chkOnlyDiffs.Text = "Nur Unterschiede anzeigen"
    $script:chkOnlyDiffs.Location = New-Object System.Drawing.Point(12, 52)
    $script:chkOnlyDiffs.AutoSize = $true
    $script:chkOnlyDiffs.Checked = $true
    $script:chkOnlyDiffs.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $btnExportCompareCsv = New-Object System.Windows.Forms.Button
    $btnExportCompareCsv.Text = "Diff CSV Export"
    $btnExportCompareCsv.Location = New-Object System.Drawing.Point(235, 48)
    $btnExportCompareCsv.Size = New-Object System.Drawing.Size(130, 30)
    $btnExportCompareCsv.BackColor = [System.Drawing.Color]::FromArgb(230, 245, 230)

    $script:lblCompareStatus = New-Object System.Windows.Forms.Label
    $script:lblCompareStatus.Text = "Waehlen Sie zwei GPOs fuer den Vergleich."
    $script:lblCompareStatus.Location = New-Object System.Drawing.Point(380, 55)
    $script:lblCompareStatus.AutoSize = $true
    $script:lblCompareStatus.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Italic)

    $panelCompareTop.Controls.AddRange(@(
        $lblGpo1, $script:comboCompareGpo1, $lblGpo2, $script:comboCompareGpo2, 
        $btnCompare, $btnCompareDdpBaseline, $script:chkOnlyDiffs, 
        $btnExportCompareCsv, $script:lblCompareStatus
    ))

    $panelCompareMain = New-Object System.Windows.Forms.Panel
    $panelCompareMain.Dock = [System.Windows.Forms.DockStyle]::Fill
    $panelCompareMain.Padding = New-Object System.Windows.Forms.Padding(10, 8, 10, 10)

    $lblCompareLegend = New-Object System.Windows.Forms.Label
    $lblCompareLegend.Text = "Legende:  [Gruen] Identische Einstellung  |  [Gelb/Orange] Abweichender Wert / Status  |  [Rot] Nur in GPO 1  |  [Blau] Nur in GPO 2"
    $lblCompareLegend.Dock = [System.Windows.Forms.DockStyle]::Top
    $lblCompareLegend.Height = 28
    $lblCompareLegend.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Italic)

    $script:gridCompare = New-Object System.Windows.Forms.DataGridView
    $script:gridCompare.Dock = [System.Windows.Forms.DockStyle]::Fill
    $script:gridCompare.ReadOnly = $true
    $script:gridCompare.AllowUserToAddRows = $false
    $script:gridCompare.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $script:gridCompare.MultiSelect = $false
    $script:gridCompare.RowHeadersVisible = $false
    $script:gridCompare.BackgroundColor = [System.Drawing.Color]::White
    $script:gridCompare.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $script:gridCompare.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(230, 236, 245)
    $script:gridCompare.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    $script:gridCompare.ColumnHeadersHeight = 34
    $script:gridCompare.RowTemplate.Height = 28
    $script:gridCompare.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(249, 251, 254)

    [void]$script:gridCompare.Columns.Add("colCmpScope", "Bereich")
    [void]$script:gridCompare.Columns.Add("colCmpCategory", "Kategorie / Pfad")
    [void]$script:gridCompare.Columns.Add("colCmpName", "Einstellung (Name)")
    [void]$script:gridCompare.Columns.Add("colCmpState1", "Status (GPO 1)")
    [void]$script:gridCompare.Columns.Add("colCmpVal1", "Wert in GPO 1")
    [void]$script:gridCompare.Columns.Add("colCmpState2", "Status (GPO 2)")
    [void]$script:gridCompare.Columns.Add("colCmpVal2", "Wert in GPO 2")
    [void]$script:gridCompare.Columns.Add("colCmpStatus", "Vergleichs-Status")

    $script:gridCompare.Columns["colCmpScope"].Width = 85
    $script:gridCompare.Columns["colCmpCategory"].Width = 220
    $script:gridCompare.Columns["colCmpName"].Width = 260
    $script:gridCompare.Columns["colCmpState1"].Width = 110
    $script:gridCompare.Columns["colCmpVal1"].Width = 210
    $script:gridCompare.Columns["colCmpState2"].Width = 110
    $script:gridCompare.Columns["colCmpVal2"].Width = 210
    $script:gridCompare.Columns["colCmpStatus"].AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill

    $panelCompareMain.Controls.Add($script:gridCompare)
    $tabCompare.Controls.Add($panelCompareMain)
    $tabCompare.Controls.Add($panelCompareTop)
    $tabControl.TabPages.Add($tabCompare)

    # Rendering-Funktion
    $script:Update_CompareGridDisplay = {
        if ($script:isClosing -or $null -eq $script:gridCompare -or $script:gridCompare.IsDisposed) { return }
        $script:gridCompare.Rows.Clear()
        $onlyDiffs = if ($script:chkOnlyDiffs) { $script:chkOnlyDiffs.Checked } else { $true }
        $diffCount = 0
        $totalCount = $script:rawCompareList.Count

        foreach ($item in $script:rawCompareList) {
            if ($item.DiffStatus -ne "Identisch") { $diffCount++ }
        }

        $suppress = ($diffCount -eq 0 -and $totalCount -gt 0)

        foreach ($item in $script:rawCompareList) {
            $isDiff = $item.DiffStatus -ne "Identisch"
            if ($onlyDiffs -and -not $isDiff -and -not $suppress) { continue }

            $idx = $script:gridCompare.Rows.Add($item.Scope, $item.Category, $item.Name, $item.StateGPO1, $item.ValueGPO1, $item.StateGPO2, $item.ValueGPO2, $item.DiffStatus)
            $row = $script:gridCompare.Rows[$idx]

            switch -Regex ($item.DiffStatus) {
                "^Identisch"   { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(235, 247, 235); $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkGreen }
                "^Abweichend"  { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::LemonChiffon; $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkOrange; $row.DefaultCellStyle.Font = New-Object System.Drawing.Font($script:gridCompare.Font, [System.Drawing.FontStyle]::Bold) }
                "Nur in GPO 1" { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::MistyRose; $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkRed }
                "Nur in GPO 2" { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::AliceBlue; $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkBlue }
            }
        }

        if ($script:lblCompareStatus -and -not $script:lblCompareStatus.IsDisposed) {
            $script:lblCompareStatus.Text = if ($diffCount -eq 0 -and $totalCount -gt 0) { "[OK] 0 Unterschiede! Alle $totalCount Richtlinien sind identisch." } else { "$diffCount Unterschied(e) bei $totalCount untersuchten Richtlinien." }
        }
    }

    # Vergleichs-Funktion
    $script:Invoke_GpoCompare = {
        if ($null -eq $script:comboCompareGpo1 -or $null -eq $script:comboCompareGpo2) { return }
        $gpoName1 = $script:comboCompareGpo1.SelectedItem
        $gpoName2 = $script:comboCompareGpo2.SelectedItem
        if ([string]::IsNullOrWhiteSpace($gpoName1) -or [string]::IsNullOrWhiteSpace($gpoName2)) { return }

        if ($script:lblCompareStatus) { $script:lblCompareStatus.Text = "Analysiere & Vergleiche..." }
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $form.Refresh()
        $script:rawCompareList.Clear()

        try {
            $list1 = if ($gpoName1 -eq $script:DdpBaselineName) { Get-DefaultDomainPolicyBaseline } else {
                $gObj1 = $script:allGposCache | Where-Object { $_.DisplayName -eq $gpoName1 } | Select-Object -First 1
                if (-not $gObj1) { $gObj1 = Get-GPO -Name $gpoName1 -ErrorAction Stop }
                Get-ParsedGpoSettings -GpoId $gObj1.Id.ToString() -GpoDisplayName $gpoName1
            }

            $list2 = if ($gpoName2 -eq $script:DdpBaselineName) { Get-DefaultDomainPolicyBaseline } else {
                $gObj2 = $script:allGposCache | Where-Object { $_.DisplayName -eq $gpoName2 } | Select-Object -First 1
                if (-not $gObj2) { $gObj2 = Get-GPO -Name $gpoName2 -ErrorAction Stop }
                Get-ParsedGpoSettings -GpoId $gObj2.Id.ToString() -GpoDisplayName $gpoName2
            }

            $dict1 = @{}; foreach ($it in $list1) { $dict1[(Get-NormalizedPolicyKey $it)] = $it }
            $dict2 = @{}; foreach ($it in $list2) { $dict2[(Get-NormalizedPolicyKey $it)] = $it }

            $allKeys = [System.Collections.Generic.HashSet[string]]::new()
            foreach ($k in $dict1.Keys) { [void]$allKeys.Add($k) }
            foreach ($k in $dict2.Keys) { [void]$allKeys.Add($k) }

            foreach ($k in ($allKeys | Sort-Object)) {
                $item1 = $dict1[$k]
                $item2 = $dict2[$k]

                $diffStatus = ""
                if ($null -ne $item1 -and $null -ne $item2) {
                    $valuesMatch = Test-PolicyValueMatch $item1.Value $item2.Value
                    $statesMatch = Test-PolicyValueMatch $item1.State $item2.State

                    if ($valuesMatch -and $statesMatch) {
                        $diffStatus = "Identisch"
                    } elseif (-not $valuesMatch) {
                        $diffStatus = "Abweichender Wert"
                    } else {
                        $diffStatus = "Abweichender Status ($($item1.State) vs $($item2.State))"
                    }
                } elseif ($null -ne $item1) {
                    $diffStatus = "Nur in GPO 1 ($gpoName1)"
                } else {
                    $diffStatus = "Nur in GPO 2 ($gpoName2)"
                }

                $script:rawCompareList.Add([PSCustomObject]@{
                    Scope      = if ($item1) { $item1.Scope } else { $item2.Scope }
                    Category   = if ($item1) { $item1.Category } else { $item2.Category }
                    Name       = if ($item1) { $item1.Name } else { $item2.Name }
                    StateGPO1  = if ($item1) { $item1.State } else { "-" }
                    ValueGPO1  = if ($item1) { $item1.Value } else { "-- [Nicht konfiguriert] --" }
                    StateGPO2  = if ($item2) { $item2.State } else { "-" }
                    ValueGPO2  = if ($item2) { $item2.Value } else { "-- [Nicht konfiguriert] --" }
                    DiffStatus = $diffStatus
                    GPO1       = $gpoName1
                    GPO2       = $gpoName2
                })
            }

            & $script:Update_CompareGridDisplay
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Vergleich: $($_.Exception.Message)", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }

    $btnCompare.Add_Click({ & $script:Invoke_GpoCompare })

    $btnCompareDdpBaseline.Add_Click({
        $realDdp = $script:allGposCache | Where-Object { $_.Id.ToString().Trim('{','}').ToUpper() -eq $script:StandardDdpGuid } | Select-Object -First 1
        if ($realDdp) { 
            $script:comboCompareGpo1.SelectedItem = $realDdp.DisplayName 
        } else {
            $cand = $script:comboCompareGpo1.Items | Where-Object { $_ -match "^Default Domain Policy$" } | Select-Object -First 1
            if ($cand) { $script:comboCompareGpo1.SelectedItem = $cand } else { $script:comboCompareGpo1.SelectedIndex = 0 }
        }

        if (-not $script:comboCompareGpo2.Items.Contains($script:DdpBaselineName)) { 
            [void]$script:comboCompareGpo2.Items.Add($script:DdpBaselineName) 
        }
        $script:comboCompareGpo2.SelectedItem = $script:DdpBaselineName
        $script:chkOnlyDiffs.Checked = $false
        & $script:Invoke_GpoCompare
    })

    $script:chkOnlyDiffs.Add_CheckedChanged({ & $script:Update_CompareGridDisplay })

    $btnExportCompareCsv.Add_Click({
        if ($script:rawCompareList.Count -eq 0) { return }
        $targetBase = if ($script:txtBackupTargetDir) { $script:txtBackupTargetDir.Text.Trim() } else { "C:\Install\Backup\GPO" }
        if (-not (Test-Path $targetBase)) { New-Item -ItemType Directory -Path $targetBase -Force | Out-Null }
        $csvFile = Join-Path $targetBase "GPO_Compare_$((Get-Date).ToString('yyyyMMdd_HHmm')).csv"
        $script:rawCompareList | Export-Csv -Path $csvFile -Delimiter ";" -NoTypeInformation -Encoding UTF8
        [System.Windows.Forms.MessageBox]::Show("Exportiert nach: $csvFile", "Export abgeschlossen", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    })
}