<#
================================================================================
 TOOL 1: MULTI-DC LASTLOGON ANALYSE
 Datei: C:\install\AdminSuite\Tools\Tool01_LastLogon.ps1
 Features:
  - Versionsnummer nur im Fenstertitel (v2.2)
  - 100% ASCII-sichere Beschriftungen (keine Umlautprobleme)
  - Lokale AD-Site Ermittlung (standardmaessig nur lokale DCs vorausgewaehlt)
  - Sofort schaltbare DC-Checkboxen & Schnellauswahl-Buttons
  - Filter: AD-Konto-Status (Aktiviert / Deaktiviert)
  - Filter: Login-Aktivitaet (<= 90 Tage aktiv / > 90 Tage inaktiv)
  - Spalten: Status, Inaktiv seit (Tage), Neuester Login & DC-Einzelwerte
  - Spaltensortierung per Klick auf Spaltenkopf (Auf-/Absteigend)
  - Farbmarkierung: Hellgruen (Aktiv) und Hellrot (Inaktiv)
================================================================================
#>

function Open-ToolLastLogon {
    $toolVersion = "v2.2"

    if (Get-Command Assert-DomainJoined -ErrorAction SilentlyContinue) {
        if (-not (Assert-DomainJoined)) { return }
    }

    $subForm = New-Object System.Windows.Forms.Form
    $subForm.Text = "Tool 1: Multi-DC LastLogon Analyse ($toolVersion)"
    $subForm.Size = New-Object System.Drawing.Size(1280, 820)
    $subForm.StartPosition = "CenterParent"
    $subForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $subForm.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

    # -------------------------------------------------------------
    # 1. TOP PANEL: SUCHFELD & FILTER
    # -------------------------------------------------------------
    $pnlTop = New-Object System.Windows.Forms.Panel
    $pnlTop.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlTop.Height = 85
    $pnlTop.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
    $subForm.Controls.Add($pnlTop)

    # Zeile 1: Suchfeld & Sortierung
    $lblSearch = New-Object System.Windows.Forms.Label
    $lblSearch.Location = New-Object System.Drawing.Point(15, 14)
    $lblSearch.Text = "Computer (Wildcard):"
    $lblSearch.Size = New-Object System.Drawing.Size(130, 20)
    $pnlTop.Controls.Add($lblSearch)

    $txtSearch = New-Object System.Windows.Forms.TextBox
    $txtSearch.Location = New-Object System.Drawing.Point(150, 11)
    $txtSearch.Size = New-Object System.Drawing.Size(160, 23)
    $txtSearch.Text = "*"
    $pnlTop.Controls.Add($txtSearch)

    $lblSort = New-Object System.Windows.Forms.Label
    $lblSort.Location = New-Object System.Drawing.Point(325, 14)
    $lblSort.Text = "Sortieren nach:"
    $lblSort.Size = New-Object System.Drawing.Size(100, 20)
    $pnlTop.Controls.Add($lblSort)

    $cmbSortBy = New-Object System.Windows.Forms.ComboBox
    $cmbSortBy.Location = New-Object System.Drawing.Point(435, 11)
    $cmbSortBy.Size = New-Object System.Drawing.Size(155, 23)
    $cmbSortBy.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$cmbSortBy.Items.AddRange(@("Computer", "Status", "Inaktiv seit (Tage)", "Neuester Login"))
    $cmbSortBy.SelectedIndex = 0
    $pnlTop.Controls.Add($cmbSortBy)

    # Zeile 2: AD-Kontostatus & Login-Aktivitaets-Filter
    $lblEnabled = New-Object System.Windows.Forms.Label
    $lblEnabled.Location = New-Object System.Drawing.Point(15, 48)
    $lblEnabled.Text = "Konto-Status:"
    $lblEnabled.Size = New-Object System.Drawing.Size(130, 20)
    $pnlTop.Controls.Add($lblEnabled)

    $cmbEnabledFilter = New-Object System.Windows.Forms.ComboBox
    $cmbEnabledFilter.Location = New-Object System.Drawing.Point(150, 45)
    $cmbEnabledFilter.Size = New-Object System.Drawing.Size(160, 23)
    $cmbEnabledFilter.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$cmbEnabledFilter.Items.AddRange(@("Alle Status", "Nur Aktivierte", "Nur Deaktivierte"))
    $cmbEnabledFilter.SelectedIndex = 0
    $pnlTop.Controls.Add($cmbEnabledFilter)

    $lblActivity = New-Object System.Windows.Forms.Label
    $lblActivity.Location = New-Object System.Drawing.Point(325, 48)
    $lblActivity.Text = "Login-Aktivitaet:"
    $lblActivity.Size = New-Object System.Drawing.Size(105, 20)
    $pnlTop.Controls.Add($lblActivity)

    $cmbActivityFilter = New-Object System.Windows.Forms.ComboBox
    $cmbActivityFilter.Location = New-Object System.Drawing.Point(435, 45)
    $cmbActivityFilter.Size = New-Object System.Drawing.Size(155, 23)
    $cmbActivityFilter.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$cmbActivityFilter.Items.AddRange(@("Alle Computer", "Nur Aktive (<= 90 Tage)", "Nur Inaktive (> 90 Tage / Nie)"))
    $cmbActivityFilter.SelectedIndex = 0
    $pnlTop.Controls.Add($cmbActivityFilter)

    # Start-Button
    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Location = New-Object System.Drawing.Point(605, 11)
    $btnRun.Size = New-Object System.Drawing.Size(130, 58)
    $btnRun.Text = "Abfragen"
    $btnRun.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $btnRun.ForeColor = [System.Drawing.Color]::White
    $btnRun.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnRun.FlatAppearance.BorderSize = 0
    $btnRun.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    $pnlTop.Controls.Add($btnRun)

    # Statusanzeige
    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = New-Object System.Drawing.Point(750, 11)
    $lblStatus.Size = New-Object System.Drawing.Size(510, 58)
    $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(30, 58, 138)
    $lblStatus.Text = "Ermittle lokale AD-Site und Domain Controller..."
    $pnlTop.Controls.Add($lblStatus)

    # -------------------------------------------------------------
    # 2. DC-AUSWAHL GROUPBOX MIT INTERAKTIVEM GRID
    # -------------------------------------------------------------
    $grpDCs = New-Object System.Windows.Forms.GroupBox
    $grpDCs.Text = "Verfuegbare Domain Controller (Waehlen Sie die abzufragenden DCs aus)"
    $grpDCs.Dock = [System.Windows.Forms.DockStyle]::Top
    $grpDCs.Height = 160
    $grpDCs.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
    $subForm.Controls.Add($grpDCs)

    # Schnellauswahl-Buttons
    $pnlDcBtns = New-Object System.Windows.Forms.Panel
    $pnlDcBtns.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlDcBtns.Height = 28
    $grpDCs.Controls.Add($pnlDcBtns)

    $btnSelAll = New-Object System.Windows.Forms.Button
    $btnSelAll.Text = "Alle auswaehlen"
    $btnSelAll.Location = New-Object System.Drawing.Point(8, 2)
    $btnSelAll.Size = New-Object System.Drawing.Size(120, 23)
    $btnSelAll.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $pnlDcBtns.Controls.Add($btnSelAll)

    $btnDeselAll = New-Object System.Windows.Forms.Button
    $btnDeselAll.Text = "Alle abwaehlen"
    $btnDeselAll.Location = New-Object System.Drawing.Point(135, 2)
    $btnDeselAll.Size = New-Object System.Drawing.Size(120, 23)
    $btnDeselAll.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $pnlDcBtns.Controls.Add($btnDeselAll)

    $gridDCs = New-Object System.Windows.Forms.DataGridView
    $gridDCs.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridDCs.AllowUserToAddRows = $false
    $gridDCs.AllowUserToDeleteRows = $false
    $gridDCs.RowHeadersVisible = $false
    $gridDCs.BackgroundColor = [System.Drawing.Color]::White
    $gridDCs.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $gridDCs.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $gridDCs.EditMode = [System.Windows.Forms.DataGridViewEditMode]::EditOnEnter
    $grpDCs.Controls.Add($gridDCs)
    $gridDCs.BringToFront()

    $colCheck = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $colCheck.Name = "Check"
    $colCheck.HeaderText = "Abfragen"
    $colCheck.Width = 70
    $colCheck.ReadOnly = $false
    [void]$gridDCs.Columns.Add($colCheck)

    $colHost = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colHost.Name = "HostName"
    $colHost.HeaderText = "Domain Controller"
    $colHost.ReadOnly = $true
    [void]$gridDCs.Columns.Add($colHost)

    $colSite = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colSite.Name = "Site"
    $colSite.HeaderText = "AD Site"
    $colSite.ReadOnly = $true
    [void]$gridDCs.Columns.Add($colSite)

    $gridDCs.Columns[0].FillWeight = 15
    $gridDCs.Columns[1].FillWeight = 55
    $gridDCs.Columns[2].FillWeight = 30

    $gridDCs.Add_CurrentCellDirtyStateChanged({
        if ($gridDCs.IsCurrentCellDirty) {
            $gridDCs.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
        }
    })

    # -------------------------------------------------------------
    # 3. STATUS-LEGENDE AM UNTEREN RAND
    # -------------------------------------------------------------
    $pnlLegend = New-Object System.Windows.Forms.Panel
    $pnlLegend.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $pnlLegend.Height = 36
    $pnlLegend.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
    $subForm.Controls.Add($pnlLegend)

    $lblLegTitle = New-Object System.Windows.Forms.Label
    $lblLegTitle.Location = New-Object System.Drawing.Point(15, 9)
    $lblLegTitle.Text = "Status-Legende:"
    $lblLegTitle.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
    $lblLegTitle.AutoSize = $true
    $pnlLegend.Controls.Add($lblLegTitle)

    $pnlBoxGreen = New-Object System.Windows.Forms.Panel
    $pnlBoxGreen.Location = New-Object System.Drawing.Point(120, 10); $pnlBoxGreen.Size = New-Object System.Drawing.Size(12, 12)
    $pnlBoxGreen.BackColor = [System.Drawing.Color]::FromArgb(220, 252, 231); $pnlBoxGreen.BorderStyle = "FixedSingle"
    $pnlLegend.Controls.Add($pnlBoxGreen)

    $lblLegGreen = New-Object System.Windows.Forms.Label
    $lblLegGreen.Location = New-Object System.Drawing.Point(137, 9)
    $lblLegGreen.Text = "Login <= 90 Tage (Aktiv)"
    $lblLegGreen.AutoSize = $true
    $pnlLegend.Controls.Add($lblLegGreen)

    $pnlBoxRed = New-Object System.Windows.Forms.Panel
    $pnlBoxRed.Location = New-Object System.Drawing.Point(290, 10); $pnlBoxRed.Size = New-Object System.Drawing.Size(12, 12)
    $pnlBoxRed.BackColor = [System.Drawing.Color]::FromArgb(254, 226, 226); $pnlBoxRed.BorderStyle = "FixedSingle"
    $pnlLegend.Controls.Add($pnlBoxRed)

    $lblLegRed = New-Object System.Windows.Forms.Label
    $lblLegRed.Location = New-Object System.Drawing.Point(307, 9)
    $lblLegRed.Text = "Login > 90 Tage / Nie (Inaktiv)"
    $lblLegRed.AutoSize = $true
    $pnlLegend.Controls.Add($lblLegRed)

    $lblLegInfo = New-Object System.Windows.Forms.Label
    $lblLegInfo.Location = New-Object System.Drawing.Point(500, 9)
    $lblLegInfo.Text = "|  Status 'Deaktiviert' = AD-Konto deaktiviert  |  Klick auf Spaltenkopf sortiert die Tabelle"
    $lblLegInfo.ForeColor = [System.Drawing.Color]::DarkRed
    $lblLegInfo.AutoSize = $true
    $pnlLegend.Controls.Add($lblLegInfo)

    # -------------------------------------------------------------
    # 4. DATA GRID VIEW ERGEBNISSE
    # -------------------------------------------------------------
    $gridResults = New-Object System.Windows.Forms.DataGridView
    $gridResults.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridResults.ReadOnly = $true
    $gridResults.AllowUserToAddRows = $false
    $gridResults.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $gridResults.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
    $gridResults.BackgroundColor = [System.Drawing.Color]::White
    $gridResults.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $gridResults.RowHeadersVisible = $false
    $gridResults.EnableHeadersVisualStyles = $false
    $gridResults.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(241, 245, 249)
    $gridResults.ColumnHeadersHeight = 32
    $subForm.Controls.Add($gridResults)

    # Z-Order Korrektur
    $gridResults.BringToFront()
    $pnlLegend.SendToBack()
    $grpDCs.SendToBack()
    $pnlTop.SendToBack()

    # -------------------------------------------------------------
    # LOGIK & EVENTS
    # -------------------------------------------------------------
    $btnSelAll.Add_Click({ foreach ($r in $gridDCs.Rows) { $r.Cells["Check"].Value = $true } })
    $btnDeselAll.Add_Click({ foreach ($r in $gridDCs.Rows) { $r.Cells["Check"].Value = $false } })

    $script:rawComputerResults = @()

    $applyRowColors = {
        $ninetyDaysAgo = (Get-Date).AddDays(-90)
        foreach ($row in $gridResults.Rows) {
            $statusVal = $row.Cells["Status"].Value
            $latestVal = $row.Cells["Neuester Login"].Value

            if ($statusVal -eq "Deaktiviert") {
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Gray
                $row.Cells["Status"].Style.ForeColor = [System.Drawing.Color]::DarkRed
                $row.Cells["Status"].Style.Font = New-Object System.Drawing.Font($gridResults.Font, [System.Drawing.FontStyle]::Bold)
            } else {
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Black
            }

            if ($latestVal -and $latestVal -ne "Nie") {
                try {
                    $parsed = [datetime]::ParseExact($latestVal, "dd.MM.yyyy HH:mm:ss", $null)
                    if ($parsed -ge $ninetyDaysAgo) {
                        $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(220, 252, 231)
                    } else {
                        $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(254, 226, 226)
                    }
                } catch {
                    $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::White
                }
            } else {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(254, 226, 226)
            }
        }
    }

    $enableColumnSortMode = {
        foreach ($c in $gridResults.Columns) {
            $c.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Programmatic
        }
    }

    $applyFilterAndSort = {
        if (-not $script:rawComputerResults -or $script:rawComputerResults.Count -eq 0) { return }

        $statusSel   = $cmbEnabledFilter.SelectedIndex
        $activitySel = $cmbActivityFilter.SelectedIndex
        $sortProp    = $cmbSortBy.SelectedItem.ToString()

        $filtered = $script:rawComputerResults | Where-Object {
            $matchStatus = switch ($statusSel) {
                1 { $_.Status -eq "Aktiviert" }
                2 { $_.Status -eq "Deaktiviert" }
                Default { $true }
            }

            $matchActivity = switch ($activitySel) {
                1 { $_."Inaktiv seit (Tage)" -ne "Nie" -and [int]$_."Inaktiv seit (Tage)" -le 90 }
                2 { $_."Inaktiv seit (Tage)" -eq "Nie" -or [int]$_."Inaktiv seit (Tage)" -gt 90 }
                Default { $true }
            }

            $matchStatus -and $matchActivity
        }

        $sorted = $filtered | Sort-Object {
            if ($sortProp -eq "Inaktiv seit (Tage)") {
                $val = $_.$sortProp
                if ($val -eq "Nie") { [int]::MaxValue } else { [int]$val }
            } else {
                $_.$sortProp
            }
        }

        $arr = [System.Collections.ArrayList]::new()
        foreach ($item in $sorted) { [void]$arr.Add($item) }
        $gridResults.DataSource = $null
        $gridResults.DataSource = $arr
        
        & $enableColumnSortMode
        & $applyRowColors

        $lblStatus.Text = "Bereit: $($arr.Count) von $($script:rawComputerResults.Count) Computer(n) angezeigt."
    }

    # Spaltenkopf-Klick Sortierung
    $script:sortCol = ""
    $script:sortAsc = $true

    $gridResults.Add_ColumnHeaderMouseClick({
        param($sender, $e)

        $col = $gridResults.Columns[$e.ColumnIndex]
        if (-not $col) { return }
        $propName = if ($col.DataPropertyName) { $col.DataPropertyName } else { $col.HeaderText }

        if ($script:sortCol -eq $propName) {
            $script:sortAsc = -not $script:sortAsc
        } else {
            $script:sortCol = $propName
            $script:sortAsc = $true
        }

        $currentData = @($gridResults.DataSource)
        if ($null -eq $currentData -or $currentData.Count -le 1) { return }

        $sorted = $currentData | Sort-Object -Property @{
            Expression = {
                $val = $_.$propName
                if ($null -eq $val -or $val -eq "") { return "" }

                if ($propName -like "*Tage*" -or $propName -like "*Inaktiv*") {
                    if ($val -eq "Nie" -or $val -eq "N/A") { return [int]::MaxValue }
                    if ($val -as [int]) { return [int]$val }
                }

                if ($val -match '^\d{2}\.\d{2}\.\d{4}') {
                    try {
                        return [datetime]::ParseExact($val.ToString().Trim(), @("dd.MM.yyyy HH:mm:ss", "dd.MM.yyyy"), $null)
                    } catch { return $val }
                }

                return $val
            }
            Descending = (-not $script:sortAsc)
        }

        $arr = [System.Collections.ArrayList]::new()
        foreach ($item in $sorted) { [void]$arr.Add($item) }
        $gridResults.DataSource = $null
        $gridResults.DataSource = $arr

        & $enableColumnSortMode
        foreach ($c in $gridResults.Columns) {
            $c.HeaderCell.SortGlyphDirection = [System.Windows.Forms.SortOrder]::None
        }
        $gridResults.Columns[$e.ColumnIndex].HeaderCell.SortGlyphDirection = `
            $(if ($script:sortAsc) { [System.Windows.Forms.SortOrder]::Ascending } else { [System.Windows.Forms.SortOrder]::Descending })

        & $applyRowColors
    })

    # DCs laden (Standardmaessig nur lokale Site vorausgewaehlt)
    $subForm.Add_Shown({
        $lblStatus.Text = "Ermittle lokale Site und Domain Controller..."
        $subForm.Refresh()
        $gridDCs.Rows.Clear()

        $localSiteName = ""
        try {
            $localSiteName = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().Name
        } catch {
            $localSiteName = ""
        }

        try {
            $domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
            $dcs = $domain.DomainControllers
            
            $matchedCount = 0
            foreach ($dc in $dcs) {
                $isSameSite = ($localSiteName -ne "" -and $dc.SiteName -eq $localSiteName)
                if ($isSameSite) { $matchedCount++ }
                
                [void]$gridDCs.Rows.Add($isSameSite, $dc.Name, $dc.SiteName)
            }

            if ($matchedCount -eq 0 -and $gridDCs.Rows.Count -gt 0) {
                foreach ($row in $gridDCs.Rows) { $row.Cells["Check"].Value = $true }
                $lblStatus.Text = "Bereit. $($dcs.Count) DC(s) geladen (Keine Site-Uebereinstimmung -> alle aktiviert)."
            } else {
                $lblStatus.Text = "Bereit. Lokale Site '$localSiteName': $matchedCount von $($dcs.Count) DC(s) vorausgewaehlt."
            }
        } catch {
            $fallback = if ($env:LOGONSERVER) { $env:LOGONSERVER.Replace("\", "") } else { "localhost" }
            [void]$gridDCs.Rows.Add($true, $fallback, "Default-Site")
            $lblStatus.Text = "Bereit - Fallback DC geladen."
        }
    })

    # Abfrage starten
    $btnRun.Add_Click({
        $selectedDCs = @()
        foreach ($row in $gridDCs.Rows) {
            if ($row.Cells["Check"].Value -eq $true) {
                $selectedDCs += $row.Cells["HostName"].Value
            }
        }

        if ($selectedDCs.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Bitte mindestens einen Domain Controller auswaehlen.", "Hinweis", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        $lblStatus.Text = "Suche Computer im Active Directory..."
        $subForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        [System.Windows.Forms.Application]::DoEvents()

        try {
            $namePattern = $txtSearch.Text.Trim()
            if ([string]::IsNullOrEmpty($namePattern)) { $namePattern = "*" }
            
            $ldapFilter = if ($namePattern -eq "*") {
                "(&(objectCategory=computer))"
            } else {
                "(&(objectCategory=computer)(name=$namePattern))"
            }

            $found = Search-NativeLdap -LdapFilter $ldapFilter -PropertiesToLoad @("name","distinguishedName","userAccountControl")

            if (-not $found -or $found.Count -eq 0) {
                $lblStatus.Text = "Keine Computer-Objekte gefunden."
                $gridResults.DataSource = $null
                $script:rawComputerResults = @()
                return
            }

            if ($found.Count -gt 35) {
                $msgResult = [System.Windows.Forms.MessageBox]::Show(
                    "Es wurden $($found.Count) Computer gefunden.`n`nMulti-DC-Abfrage fuer $($selectedDCs.Count) DC(s) starten?",
                    "Hinweis: Grosse Treffermenge",
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Question
                )
                if ($msgResult -ne [System.Windows.Forms.DialogResult]::Yes) {
                    $lblStatus.Text = "Abfrage abgebrochen."
                    return
                }
            }

            $results = [System.Collections.ArrayList]::new()
            $i = 0
            $now = Get-Date

            foreach ($c in $found) {
                $i++
                $cName = $c.Properties["name"][0]
                $cDn   = $c.Properties["distinguishedname"][0]

                $isEnabled = $true
                if ($c.Properties["useraccountcontrol"].Count -gt 0) {
                    $uac = [int]$c.Properties["useraccountcontrol"][0]
                    if (($uac -band 2) -eq 2) { $isEnabled = $false }
                }

                $lblStatus.Text = "Frage DCs ab ($i von $($found.Count)): $cName"
                [System.Windows.Forms.Application]::DoEvents()

                $latest = [datetime]::MinValue
                $bestDC = "N/A"
                $dcValues = [ordered]@{}

                foreach ($dc in $selectedDCs) {
                    try {
                        $res = Search-NativeLdap -Server $dc -LdapFilter "(distinguishedName=$cDn)" -PropertiesToLoad @("lastLogon")
                        if ($res -and $res[0].Properties["lastlogon"].Count -gt 0 -and $res[0].Properties["lastlogon"][0] -gt 0) {
                            $dt = [DateTime]::FromFileTime($res[0].Properties["lastlogon"][0])
                            $dcValues[$dc] = $dt.ToString("dd.MM.yyyy HH:mm:ss")
                            if ($dt -gt $latest) {
                                $latest = $dt
                                $bestDC = $dc
                            }
                        } else {
                            $dcValues[$dc] = "Nie"
                        }
                    } catch {
                        $dcValues[$dc] = "Nicht erreichbar"
                    }
                }

                $inactivityDays = if ($latest -eq [datetime]::MinValue) {
                    "Nie"
                } else {
                    [int][math]::Round(($now - $latest).TotalDays, 0)
                }

                $rowMap = [ordered]@{
                    "Computer"            = $cName
                    "Status"              = if ($isEnabled) { "Aktiviert" } else { "Deaktiviert" }
                    "Inaktiv seit (Tage)" = $inactivityDays
                    "Neuester Login"      = if ($latest -eq [datetime]::MinValue) { "Nie" } else { $latest.ToString("dd.MM.yyyy HH:mm:ss") }
                    "Neuester DC"         = $bestDC
                }

                foreach ($k in $dcValues.Keys) {
                    $rowMap[$k] = $dcValues[$k]
                }

                [void]$results.Add([PSCustomObject]$rowMap)
            }

            $script:rawComputerResults = $results
            & $applyFilterAndSort

        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Scan: $($_.Exception.Message)", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $lblStatus.Text = "Fehler: " + $_.Exception.Message
        } finally {
            $subForm.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    })

    # Dropdown-Events
    $cmbEnabledFilter.Add_SelectedIndexChanged({ & $applyFilterAndSort })
    $cmbActivityFilter.Add_SelectedIndexChanged({ & $applyFilterAndSort })
    $cmbSortBy.Add_SelectedIndexChanged({ & $applyFilterAndSort })

    $txtSearch.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $_.SuppressKeyPress = $true
            $btnRun.PerformClick()
        }
    })

    [void]$subForm.ShowDialog()
}