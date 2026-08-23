# ===================================================================
# MODULE TOOL 1: MULTI-DC LASTLOGON ANALYSE (KORRIGIERTE DC-AUSWAHL)
# ===================================================================
function Open-ToolLastLogon {
    if (-not (Assert-DomainJoined)) { return }

    $subForm = New-Object System.Windows.Forms.Form
    $subForm.Text = "Tool 1: Multi-DC LastLogon Analyse"
    $subForm.Size = New-Object System.Drawing.Size(1050, 750)
    $subForm.StartPosition = "CenterParent"
    $subForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $subForm.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

    # 1. TOP PANEL: Suchfeld & Start-Button
    $pnlTop = New-Object System.Windows.Forms.Panel
    $pnlTop.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlTop.Height = 75
    $subForm.Controls.Add($pnlTop)

    $lblSearch = New-Object System.Windows.Forms.Label
    $lblSearch.Location = New-Object System.Drawing.Point(15, 18)
    $lblSearch.Text = "Computer-Name (Wildcard):"
    $lblSearch.Size = New-Object System.Drawing.Size(160, 20)
    $pnlTop.Controls.Add($lblSearch)

    $txtSearch = New-Object System.Windows.Forms.TextBox
    $txtSearch.Location = New-Object System.Drawing.Point(180, 15)
    $txtSearch.Size = New-Object System.Drawing.Size(200, 23)
    $txtSearch.Text = "*"
    $pnlTop.Controls.Add($txtSearch)

    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Location = New-Object System.Drawing.Point(400, 12)
    $btnRun.Size = New-Object System.Drawing.Size(130, 48)
    $btnRun.Text = "Abfragen"
    $btnRun.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $btnRun.ForeColor = [System.Drawing.Color]::White
    $btnRun.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnRun.FlatAppearance.BorderSize = 0
    $btnRun.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    $pnlTop.Controls.Add($btnRun)

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = New-Object System.Drawing.Point(545, 15)
    $lblStatus.Size = New-Object System.Drawing.Size(470, 45)
    $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(30, 58, 138)
    $pnlTop.Controls.Add($lblStatus)

    # 2. DC-AUSWAHL PANEL MIT BUTTONS & INTERAKTIVER TABELLE
    $grpDCs = New-Object System.Windows.Forms.GroupBox
    $grpDCs.Text = "Verfügbare Domain Controller (Wählen Sie die abzufragenden DCs aus)"
    $grpDCs.Dock = [System.Windows.Forms.DockStyle]::Top
    $grpDCs.Height = 175
    $subForm.Controls.Add($grpDCs)

    # Button-Leiste über der DC-Tabelle
    $pnlDcBtns = New-Object System.Windows.Forms.Panel
    $pnlDcBtns.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlDcBtns.Height = 30
    $grpDCs.Controls.Add($pnlDcBtns)

    $btnSelAll = New-Object System.Windows.Forms.Button
    $btnSelAll.Text = "Alle DCs auswählen"
    $btnSelAll.Location = New-Object System.Drawing.Point(10, 3)
    $btnSelAll.Size = New-Object System.Drawing.Size(135, 24)
    $btnSelAll.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnSelAll.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
    $pnlDcBtns.Controls.Add($btnSelAll)

    $btnDeselAll = New-Object System.Windows.Forms.Button
    $btnDeselAll.Text = "Alle abwählen"
    $btnDeselAll.Location = New-Object System.Drawing.Point(155, 3)
    $btnDeselAll.Size = New-Object System.Drawing.Size(120, 24)
    $btnDeselAll.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnDeselAll.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
    $pnlDcBtns.Controls.Add($btnDeselAll)

    # Grid für DCs
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
    $colCheck.Width = 80
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

    $gridDCs.Columns[0].FillWeight = 20
    $gridDCs.Columns[1].FillWeight = 50
    $gridDCs.Columns[2].FillWeight = 30

    # Sofort-Commit beim Anklicken der Checkbox
    $gridDCs.Add_CurrentCellDirtyStateChanged({
        if ($gridDCs.IsCurrentCellDirty) {
            $gridDCs.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
        }
    })

    # Direkter Klick-Toggle für die Checkbox-Zelle
    $gridDCs.Add_CellContentClick({
        param($sender, $e)
        if ($e.RowIndex -ge 0 -and $e.ColumnIndex -eq 0) {
            $gridDCs.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
        }
    })

    # 3. ERGEBNISTABELLE (UNTEN)
    $gridResults = New-Object System.Windows.Forms.DataGridView
    $gridResults.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridResults.ReadOnly = $true
    $gridResults.AllowUserToAddRows = $false
    $gridResults.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
    $gridResults.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $gridResults.BackgroundColor = [System.Drawing.Color]::White
    $gridResults.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $gridResults.RowHeadersVisible = $false
    $gridResults.EnableHeadersVisualStyles = $false
    $gridResults.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(241, 245, 249)
    $gridResults.ColumnHeadersHeight = 32
    $subForm.Controls.Add($gridResults)

    # Z-Order Korrektur
    $gridResults.BringToFront()
    $grpDCs.SendToBack()
    $pnlTop.SendToBack()

    # Buttons: Alle an-/abwählen
    $btnSelAll.Add_Click({
        foreach ($row in $gridDCs.Rows) { $row.Cells["Check"].Value = $true }
    })
    $btnDeselAll.Add_Click({
        foreach ($row in $gridDCs.Rows) { $row.Cells["Check"].Value = $false }
    })

    # Laden der DCs beim Öffnen
    $subForm.Add_Shown({
        $lblStatus.Text = "Lade Domain Controller & Sites..."
        $subForm.Refresh()
        $gridDCs.Rows.Clear()
        try {
            $domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
            $dcs = $domain.DomainControllers
            foreach ($dc in $dcs) {
                [void]$gridDCs.Rows.Add($true, $dc.Name, $dc.SiteName)
            }
            $lblStatus.Text = "Bereit. $($dcs.Count) Domain Controller geladen."
        } catch {
            $fallback = if ($env:LOGONSERVER) { $env:LOGONSERVER.Replace("\", "") } else { "localhost" }
            [void]$gridDCs.Rows.Add($true, $fallback, "Default-Site")
            $lblStatus.Text = "Bereit (Fallback DC geladen)."
        }
    })

    # Scan ausführen
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
                return
            }

            if ($found.Count -gt 30) {
                $msgResult = [System.Windows.Forms.MessageBox]::Show(
                    "Es wurden $($found.Count) Computer gefunden.`n`nMulti-DC-Abfrage für $($selectedDCs.Count) DC(s) starten?",
                    "Hinweis: Große Treffermenge",
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
            foreach ($c in $found) {
                $i++
                $cName = $c.Properties["name"][0]
                $cDn   = $c.Properties["distinguishedname"][0]
                $lblStatus.Text = "Frage DCs ab ($i von $($found.Count)): $cName"
                [System.Windows.Forms.Application]::DoEvents()

                $map = [ordered]@{ "Computer" = $cName }
                $latest = [datetime]::MinValue
                $bestDC = "N/A"

                foreach ($dc in $selectedDCs) {
                    try {
                        $res = Search-NativeLdap -Server $dc -LdapFilter "(distinguishedName=$cDn)" -PropertiesToLoad @("lastLogon")
                        if ($res -and $res[0].Properties["lastlogon"].Count -gt 0 -and $res[0].Properties["lastlogon"][0] -gt 0) {
                            $dt = [DateTime]::FromFileTime($res[0].Properties["lastlogon"][0])
                            $map[$dc] = $dt.ToString("dd.MM.yyyy HH:mm:ss")
                            if ($dt -gt $latest) {
                                $latest = $dt
                                $bestDC = $dc
                            }
                        } else {
                            $map[$dc] = "Nie"
                        }
                    } catch {
                        $map[$dc] = "Nicht erreichbar"
                    }
                }

                $map["Neuester Login"] = if ($latest -eq [datetime]::MinValue) { "Nie" } else { $latest.ToString("dd.MM.yyyy HH:mm:ss") }
                $map["Neuester DC"]    = $bestDC
                $results += [PSCustomObject]$map
            }

            $arr = [System.Collections.ArrayList]::new()
            foreach ($r in $results) { [void]$arr.Add($r) }
            $gridResults.DataSource = $arr
            $lblStatus.Text = "Abgeschlossen: $($results.Count) Objekte über $($selectedDCs.Count) DC(s) abgefragt."
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Scan: $($_.Exception.Message)", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $lblStatus.Text = "Fehler: " + $_.Exception.Message
        } finally {
            $subForm.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    })

    $txtSearch.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $_.SuppressKeyPress = $true
            $btnRun.PerformClick()
        }
    })

    [void]$subForm.ShowDialog()
}