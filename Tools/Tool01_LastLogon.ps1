<#
================================================================================
 TOOL 1: MULTI-DC LASTLOGON ANTWORT- & INAKTIVITÄTSANALYSE
================================================================================
#>
function Open-ToolLastLogon {
    if (-not (Assert-DomainJoined)) { return }

    $subForm = New-Object System.Windows.Forms.Form
    $subForm.Text = "Tool 1: Multi-DC LastLogon Analyse"
    $subForm.Size = New-Object System.Drawing.Size(1280, 800)
    $subForm.StartPosition = "CenterParent"
    $subForm.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9)

    $pnlTop = New-Object System.Windows.Forms.Panel
    $pnlTop.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlTop.Height = 85
    $subForm.Controls.Add($pnlTop)

    $lblSearch = New-Object System.Windows.Forms.Label
    $lblSearch.Location = New-Object System.Drawing.Point(15, 15)
    $lblSearch.Text = "Computer (Wildcard):"
    $lblSearch.Size = New-Object System.Drawing.Size(130, 20)
    $pnlTop.Controls.Add($lblSearch)

    $txtSearch = New-Object System.Windows.Forms.TextBox
    $txtSearch.Location = New-Object System.Drawing.Point(150, 12)
    $txtSearch.Size = New-Object System.Drawing.Size(150, 23)
    $txtSearch.Text = "*"
    $pnlTop.Controls.Add($txtSearch)

    $lblSort = New-Object System.Windows.Forms.Label
    $lblSort.Location = New-Object System.Drawing.Point(315, 15)
    $lblSort.Text = "Sortieren nach:"
    $lblSort.Size = New-Object System.Drawing.Size(95, 20)
    $pnlTop.Controls.Add($lblSort)

    $cmbSortBy = New-Object System.Windows.Forms.ComboBox
    $cmbSortBy.Location = New-Object System.Drawing.Point(415, 12)
    $cmbSortBy.Size = New-Object System.Drawing.Size(140, 23)
    $cmbSortBy.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$cmbSortBy.Items.AddRange(@("Computer", "Status", "Inaktiv seit (Tage)"))
    $cmbSortBy.SelectedIndex = 0
    $pnlTop.Controls.Add($cmbSortBy)

    $lblEnabled = New-Object System.Windows.Forms.Label
    $lblEnabled.Location = New-Object System.Drawing.Point(15, 48)
    $lblEnabled.Text = "Konto-Status:"
    $lblEnabled.Size = New-Object System.Drawing.Size(130, 20)
    $pnlTop.Controls.Add($lblEnabled)

    $cmbEnabledFilter = New-Object System.Windows.Forms.ComboBox
    $cmbEnabledFilter.Location = New-Object System.Drawing.Point(150, 45)
    $cmbEnabledFilter.Size = New-Object System.Drawing.Size(150, 23)
    $cmbEnabledFilter.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$cmbEnabledFilter.Items.AddRange(@("Alle Status", "Nur Aktivierte", "Nur Deaktivierte"))
    $cmbEnabledFilter.SelectedIndex = 0
    $pnlTop.Controls.Add($cmbEnabledFilter)

    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Location = New-Object System.Drawing.Point(575, 12)
    $btnRun.Size = New-Object System.Drawing.Size(120, 56)
    $btnRun.Text = "Abfragen"
    $btnRun.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $btnRun.ForeColor = [System.Drawing.Color]::White
    $btnRun.Font = New-Object System.Drawing.Font($subForm.Font.FontFamily, 9.5, [System.Drawing.FontStyle]::Bold)
    $pnlTop.Controls.Add($btnRun)

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = New-Object System.Drawing.Point(710, 12)
    $lblStatus.Size = New-Object System.Drawing.Size(540, 60)
    $lblStatus.ForeColor = [System.Drawing.Color]::DarkBlue
    $pnlTop.Controls.Add($lblStatus)

    $grpDCs = New-Object System.Windows.Forms.GroupBox
    $grpDCs.Text = "Verfügbare Domain Controller (Wählen Sie die abzufragenden DCs aus)"
    $grpDCs.Dock = [System.Windows.Forms.DockStyle]::Top
    $grpDCs.Height = 140
    $subForm.Controls.Add($grpDCs)

    $gridDCs = New-Object System.Windows.Forms.DataGridView
    $gridDCs.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridDCs.AllowUserToAddRows = $false
    $gridDCs.AllowUserToDeleteRows = $false
    $gridDCs.RowHeadersVisible = $false
    $gridDCs.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    Apply-StandardGridTheme -Grid $gridDCs
    $grpDCs.Controls.Add($gridDCs)

    $colCheck = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $colCheck.Name = "Check"; $colCheck.HeaderText = "Abfragen"; $colCheck.Width = 70
    [void]$gridDCs.Columns.Add($colCheck)
    [void]$gridDCs.Columns.Add("HostName", "Domain Controller")
    [void]$gridDCs.Columns.Add("Site", "AD Site")

    $pnlLegend = New-Object System.Windows.Forms.Panel
    $pnlLegend.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlLegend.Height = 32
    $pnlLegend.BackColor = [System.Drawing.Color]::FromArgb(240, 243, 246)
    $subForm.Controls.Add($pnlLegend)

    $lblLegendTitle = New-Object System.Windows.Forms.Label
    $lblLegendTitle.Text = "Status-Legende:"
    $lblLegendTitle.Location = New-Object System.Drawing.Point(15, 7)
    $lblLegendTitle.AutoSize = $true
    $lblLegendTitle.Font = New-Object System.Drawing.Font($subForm.Font.FontFamily, 8.5, [System.Drawing.FontStyle]::Bold)
    $pnlLegend.Controls.Add($lblLegendTitle)

    $pnlGreenBox = New-Object System.Windows.Forms.Panel
    $pnlGreenBox.Location = New-Object System.Drawing.Point(130, 8); $pnlGreenBox.Size = New-Object System.Drawing.Size(16, 16)
    $pnlGreenBox.BackColor = [System.Drawing.Color]::FromArgb(232, 248, 232); $pnlGreenBox.BorderStyle = "FixedSingle"
    $pnlLegend.Controls.Add($pnlGreenBox)

    $lblGreenDesc = New-Object System.Windows.Forms.Label
    $lblGreenDesc.Text = "Login <= 90 Tage (Aktiv)"
    $lblGreenDesc.Location = New-Object System.Drawing.Point(152, 7); $lblGreenDesc.AutoSize = $true
    $pnlLegend.Controls.Add($lblGreenDesc)

    $pnlRedBox = New-Object System.Windows.Forms.Panel
    $pnlRedBox.Location = New-Object System.Drawing.Point(320, 8); $pnlRedBox.Size = New-Object System.Drawing.Size(16, 16)
    $pnlRedBox.BackColor = [System.Drawing.Color]::FromArgb(253, 232, 232); $pnlRedBox.BorderStyle = "FixedSingle"
    $pnlLegend.Controls.Add($pnlRedBox)

    $lblRedDesc = New-Object System.Windows.Forms.Label
    $lblRedDesc.Text = "Login > 90 Tage / Nie (Inaktiv)"
    $lblRedDesc.Location = New-Object System.Drawing.Point(342, 7); $lblRedDesc.AutoSize = $true
    $pnlLegend.Controls.Add($lblRedDesc)

    $lblDisabledDesc = New-Object System.Windows.Forms.Label
    $lblDisabledDesc.Text = "|   Status 'Deaktiviert' = AD-Konto deaktiviert  |  💡 Klick auf Spaltenkopf sortiert die Tabelle"
    $lblDisabledDesc.Location = New-Object System.Drawing.Point(530, 7); $lblDisabledDesc.AutoSize = $true
    $lblDisabledDesc.ForeColor = [System.Drawing.Color]::DarkRed
    $pnlLegend.Controls.Add($lblDisabledDesc)

    $gridResults = New-Object System.Windows.Forms.DataGridView
    $gridResults.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridResults.ReadOnly = $true
    $gridResults.AllowUserToAddRows = $false
    $gridResults.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
    $gridResults.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    Apply-StandardGridTheme -Grid $gridResults
    $subForm.Controls.Add($gridResults)

    $gridResults.BringToFront()
    $pnlLegend.SendToBack()
    $grpDCs.SendToBack()
    $pnlTop.SendToBack()

    # Dynamische Zeilenfarbgebung
    $gridResults.Add_RowPrePaint({
        param($sender, $e)
        if ($e.RowIndex -ge 0 -and $e.RowIndex -lt $gridResults.Rows.Count) {
            $row = $gridResults.Rows[$e.RowIndex]
            $latestVal = $row.Cells["NeuesterLogin"].Value
            $ninetyDaysAgo = (Get-Date).AddDays(-90)
            $isWithin90Days = $false

            if ($latestVal -and $latestVal -ne "Nie") {
                try {
                    $parsedDate = [datetime]::ParseExact($latestVal, "dd.MM.yyyy HH:mm:ss", $null)
                    if ($parsedDate -ge $ninetyDaysAgo) { $isWithin90Days = $true }
                } catch {}
            }

            if ($isWithin90Days) {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(232, 248, 232)
            } else {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(253, 232, 232)
            }

            if ($row.Cells["Status"].Value -eq "Deaktiviert") {
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
                $row.Cells["Status"].Style.ForeColor = [System.Drawing.Color]::DarkRed
                $row.Cells["Status"].Style.Font = New-Object System.Drawing.Font($gridResults.Font, [System.Drawing.FontStyle]::Bold)
            } else {
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Black
            }
        }
    })

    $subForm.Add_Shown({
        $lblStatus.Text = "Lade Domain Controller & Sites..."
        $subForm.Refresh()
        try {
            $domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
            $dcs = $domain.DomainControllers
            foreach ($dc in $dcs) {
                [void]$gridDCs.Rows.Add($true, $dc.Name, $dc.SiteName)
            }
            $lblStatus.Text = "Bereit. $($dcs.Count) Domain Controller geladen."
        } catch {
            $lblStatus.Text = "Fehler beim Laden der Domain Controller."
        }
    })

    $btnRun.Add_Click({
        $selectedDCs = @()
        foreach ($row in $gridDCs.Rows) {
            if ($row.Cells["Check"].Value -eq $true) {
                $selectedDCs += $row.Cells["HostName"].Value
            }
        }

        if ($selectedDCs.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Bitte mindestens einen Domain Controller auswählen.", "Hinweis", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        $lblStatus.Text = "Suche Computer im Active Directory..."
        $subForm.Refresh()

        $namePattern = $txtSearch.Text -replace '\*', '*'
        $selectedEnabledIndex = $cmbEnabledFilter.SelectedIndex

        $ldapBase = "(&(objectCategory=computer)(name=$namePattern)"
        if ($selectedEnabledIndex -eq 1) {
            $ldapBase += "(!(userAccountControl:1.2.840.113556.1.4.803:=2))"
        } elseif ($selectedEnabledIndex -eq 2) {
            $ldapBase += "(userAccountControl:1.2.840.113556.1.4.803:=2)"
        }
        $ldapBase += ")"

        $found = Search-NativeLdap -LdapFilter $ldapBase -PropertiesToLoad @("name","distinguishedName","userAccountControl")

        if (-not $found -or $found.Count -eq 0) {
            $lblStatus.Text = "Keine Computer-Objekte mit den angegebenen Kriterien gefunden."
            $gridResults.DataSource = $null
            return
        }

        if ($found.Count -gt 30) {
            $msgResult = [System.Windows.Forms.MessageBox]::Show(
                "Es wurden $($found.Count) Computer gefunden.`n`nMulti-DC-Abfrage für $($selectedDCs.Count) DC(s) starten?",
                "Hinweis: Größere Treffermenge",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )
            if ($msgResult -ne [System.Windows.Forms.DialogResult]::Yes) {
                $lblStatus.Text = "Abfrage abgebrochen."
                return
            }
        }

        $results = @()
        $i = 0
        $today = Get-Date
        $ninetyDaysAgo = $today.AddDays(-90)
        $countActive90 = 0
        $countInactive90 = 0
        $countDisabled = 0

        foreach ($c in $found) {
            $i++
            $cName = $c.Properties["name"][0]
            $cDn   = $c.Properties["distinguishedname"][0]
            
            $isEnabled = $true
            if ($c.Properties["useraccountcontrol"].Count -gt 0) {
                $uac = [int]$c.Properties["useraccountcontrol"][0]
                if (($uac -band 2) -eq 2) { 
                    $isEnabled = $false 
                }
            }
            $statusText = if ($isEnabled) { "Aktiviert" } else { "Deaktiviert" }
            if (-not $isEnabled) { $countDisabled++ }

            $lblStatus.Text = "Frage DCs ab... ($i von $($found.Count)): $cName"
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
                    $dcValues[$dc] = "Fehler / Nicht erreichbar"
                }
            }

            $daysInactive = if ($latest -eq [datetime]::MinValue) { 
                "Nie" 
            } else { 
                [int]($today - $latest).TotalDays 
            }

            if ($latest -ne [datetime]::MinValue -and $latest -ge $ninetyDaysAgo) {
                $countActive90++
            } else {
                $countInactive90++
            }

            $map = [ordered]@{ 
                "Computer"            = $cName
                "Status"              = $statusText
                "Inaktiv seit (Tage)" = $daysInactive
                "NeuesterLogin"       = if ($latest -eq [datetime]::MinValue) { "Nie" } else { $latest.ToString("dd.MM.yyyy HH:mm:ss") }
                "NeuesterDC"          = $bestDC
            }

            foreach ($k in $dcValues.Keys) {
                $map[$k] = $dcValues[$k]
            }

            $results += [PSCustomObject]$map
        }

        $selectedSort = $cmbSortBy.SelectedItem.ToString()
        $sortedResults = switch ($selectedSort) {
            "Status" { 
                $results | Sort-Object "Status", "Computer" 
            }
            "Inaktiv seit (Tage)" { 
                $results | Sort-Object @{
                    Expression = { 
                        if ($_."Inaktiv seit (Tage)" -eq "Nie") { [int]::MaxValue } else { [int]$_."Inaktiv seit (Tage)" } 
                    }; Descending = $false 
                }, "Computer"
            }
            Default { 
                $results | Sort-Object "Computer" 
            }
        }

        $arr = [System.Collections.ArrayList]::new()
        foreach ($r in $sortedResults) { [void]$arr.Add($r) }
        $gridResults.DataSource = $null
        $gridResults.DataSource = $arr

        $lblStatus.Text = "Ergebnis: $($sortedResults.Count) Computer gefunden (🟢 Aktiv: $countActive90 | 🔴 Inaktiv (>90d): $countInactive90 | ⚪ Deaktiviert: $countDisabled)."
    })

    [void]$subForm.ShowDialog()
}
