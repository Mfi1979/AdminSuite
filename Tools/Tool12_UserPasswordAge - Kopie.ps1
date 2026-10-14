<#
================================================================================
 TOOL 12: AD BENUTZER KENNWORTALTER & MULTI-DC LASTLOGON AUDIT
 100% ASCII-sicher (kein Zeichensalat), robustes COM LargeInteger-Parsing
 und 3s-Timeout fuer nicht erreichbare Domain Controller.
================================================================================
#>

function Open-ToolUserPasswordAge {
    if (Get-Command Assert-DomainJoined -ErrorAction SilentlyContinue) {
        if (-not (Assert-DomainJoined)) { return }
    }

    $isDE = ($script:CurrentLang -eq "DE")

    $subForm = New-Object System.Windows.Forms.Form
    $subForm.Text = if ($isDE) { "Tool 12: AD Benutzer Kennwortalter & Multi-DC LastLogon" } else { "Tool 12: AD User Password Age & Multi-DC LastLogon" }
    $subForm.Size = New-Object System.Drawing.Size(1200, 780)
    $subForm.StartPosition = "CenterParent"
    $subForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $subForm.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

    # 1. TOP PANEL
    $pnlTop = New-Object System.Windows.Forms.Panel
    $pnlTop.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlTop.Height = 85
    $pnlTop.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
    $subForm.Controls.Add($pnlTop)

    $lblSearch = New-Object System.Windows.Forms.Label
    $lblSearch.Location = New-Object System.Drawing.Point(15, 14)
    $lblSearch.Text = if ($isDE) { "Benutzer (Wildcard):" } else { "User (Wildcard):" }
    $lblSearch.Size = New-Object System.Drawing.Size(130, 20)
    $pnlTop.Controls.Add($lblSearch)

    $txtSearch = New-Object System.Windows.Forms.TextBox
    $txtSearch.Location = New-Object System.Drawing.Point(150, 11)
    $txtSearch.Size = New-Object System.Drawing.Size(170, 23)
    $txtSearch.Text = "*"
    $pnlTop.Controls.Add($txtSearch)

    $lblSort = New-Object System.Windows.Forms.Label
    $lblSort.Location = New-Object System.Drawing.Point(340, 14)
    $lblSort.Text = if ($isDE) { "Sortieren nach:" } else { "Sort by:" }
    $lblSort.Size = New-Object System.Drawing.Size(95, 20)
    $pnlTop.Controls.Add($lblSort)

    $cmbSort = New-Object System.Windows.Forms.ComboBox
    $cmbSort.Location = New-Object System.Drawing.Point(440, 11)
    $cmbSort.Size = New-Object System.Drawing.Size(150, 23)
    $cmbSort.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $cmbSort.Items.AddRange(@("Benutzername", "Anzeigename", "Status", "PW-Alter (Tage)", "Neuester Login"))
    $cmbSort.SelectedIndex = 0
    $pnlTop.Controls.Add($cmbSort)

    $lblStatusFilter = New-Object System.Windows.Forms.Label
    $lblStatusFilter.Location = New-Object System.Drawing.Point(15, 48)
    $lblStatusFilter.Text = if ($isDE) { "Konto-Status:" } else { "Account Status:" }
    $lblStatusFilter.Size = New-Object System.Drawing.Size(130, 20)
    $pnlTop.Controls.Add($lblStatusFilter)

    $cmbStatusFilter = New-Object System.Windows.Forms.ComboBox
    $cmbStatusFilter.Location = New-Object System.Drawing.Point(150, 45)
    $cmbStatusFilter.Size = New-Object System.Drawing.Size(170, 23)
    $cmbStatusFilter.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $cmbStatusFilter.Items.AddRange(@("Alle Status", "Nur Aktiv", "Nur Deaktiviert"))
    $cmbStatusFilter.SelectedIndex = 0
    $pnlTop.Controls.Add($cmbStatusFilter)

    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Location = New-Object System.Drawing.Point(610, 11)
    $btnRun.Size = New-Object System.Drawing.Size(130, 56)
    $btnRun.Text = if ($isDE) { "Abfragen" } else { "Query" }
    $btnRun.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $btnRun.ForeColor = [System.Drawing.Color]::White
    $btnRun.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnRun.FlatAppearance.BorderSize = 0
    $btnRun.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    $pnlTop.Controls.Add($btnRun)

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = New-Object System.Drawing.Point(760, 14)
    $lblStatus.Size = New-Object System.Drawing.Size(420, 50)
    $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(30, 58, 138)
    $lblStatus.Text = if ($isDE) { "Lade Domain Controller Sites..." } else { "Loading Domain Controller Sites..." }
    $pnlTop.Controls.Add($lblStatus)

    # 2. DC PANEL (ASCII-konforme Beschriftung)
    $grpDCs = New-Object System.Windows.Forms.GroupBox
    $grpDCs.Text = if ($isDE) { "Verfuegbare Domain Controller (Waehlen Sie die abzufragenden DCs aus)" } else { "Available Domain Controllers (Select DCs to query)" }
    $grpDCs.Dock = [System.Windows.Forms.DockStyle]::Top
    $grpDCs.Height = 150
    $grpDCs.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
    $subForm.Controls.Add($grpDCs)

    $gridDCs = New-Object System.Windows.Forms.DataGridView
    $gridDCs.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridDCs.AllowUserToAddRows = $false
    $gridDCs.AllowUserToDeleteRows = $false
    $gridDCs.RowHeadersVisible = $false
    $gridDCs.BackgroundColor = [System.Drawing.Color]::White
    $gridDCs.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $gridDCs.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $grpDCs.Controls.Add($gridDCs)

    $colCheck = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $colCheck.Name = "Check"
    $colCheck.HeaderText = if ($isDE) { "Abfragen" } else { "Query" }
    $colCheck.Width = 80
    [void]$gridDCs.Columns.Add($colCheck)
    [void]$gridDCs.Columns.Add("HostName", "Domain Controller")
    [void]$gridDCs.Columns.Add("Site", "AD Site")
    $gridDCs.Columns[0].FillWeight = 20
    $gridDCs.Columns[1].FillWeight = 50
    $gridDCs.Columns[2].FillWeight = 30

    # 3. STATUS-LEGENDE
    $pnlLegend = New-Object System.Windows.Forms.Panel
    $pnlLegend.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $pnlLegend.Height = 40
    $pnlLegend.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
    $subForm.Controls.Add($pnlLegend)

    $lblLegTitle = New-Object System.Windows.Forms.Label
    $lblLegTitle.Location = New-Object System.Drawing.Point(15, 12)
    $lblLegTitle.Text = "Status-Legende:"
    $lblLegTitle.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
    $lblLegTitle.AutoSize = $true
    $pnlLegend.Controls.Add($lblLegTitle)

    $pnlBoxGreen = New-Object System.Windows.Forms.Panel
    $pnlBoxGreen.Location = New-Object System.Drawing.Point(115, 13)
    $pnlBoxGreen.Size = New-Object System.Drawing.Size(12, 12)
    $pnlBoxGreen.BackColor = [System.Drawing.Color]::FromArgb(220, 252, 231)
    $pnlBoxGreen.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $pnlLegend.Controls.Add($pnlBoxGreen)

    $lblLegGreen = New-Object System.Windows.Forms.Label
    $lblLegGreen.Location = New-Object System.Drawing.Point(132, 12)
    $lblLegGreen.Text = "Login <= 90 Tage (Aktiv)"
    $lblLegGreen.AutoSize = $true
    $pnlLegend.Controls.Add($lblLegGreen)

    $pnlBoxRed = New-Object System.Windows.Forms.Panel
    $pnlBoxRed.Location = New-Object System.Drawing.Point(280, 13)
    $pnlBoxRed.Size = New-Object System.Drawing.Size(12, 12)
    $pnlBoxRed.BackColor = [System.Drawing.Color]::FromArgb(254, 226, 226)
    $pnlBoxRed.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $pnlLegend.Controls.Add($pnlBoxRed)

    $lblLegRed = New-Object System.Windows.Forms.Label
    $lblLegRed.Location = New-Object System.Drawing.Point(297, 12)
    $lblLegRed.Text = "Login > 90 Tage / Nie (Inaktiv)"
    $lblLegRed.AutoSize = $true
    $pnlLegend.Controls.Add($lblLegRed)

    $lblLegInfo = New-Object System.Windows.Forms.Label
    $lblLegInfo.Location = New-Object System.Drawing.Point(490, 12)
    $lblLegInfo.Text = "|   Status 'Deaktiviert' = AD-Konto deaktiviert   |   Klick auf Spaltenkopf sortiert die Tabelle"
    $lblLegInfo.ForeColor = [System.Drawing.Color]::DarkRed
    $lblLegInfo.AutoSize = $true
    $pnlLegend.Controls.Add($lblLegInfo)

    # 4. RESULT DATA GRID VIEW
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

    $gridResults.BringToFront()
    $pnlLegend.SendToBack()
    $grpDCs.SendToBack()
    $pnlTop.SendToBack()

    # -------------------------------------------------------------
    # HILFSFUNKTIONEN: EXAKTE LARGEINTEGER- UND TIME-CONVERSION
    # -------------------------------------------------------------
    $testDcReachability = {
        param([string]$dcHost, [int]$timeoutMs = 3000)
        try {
            $tcpClient = New-Object System.Net.Sockets.TcpClient
            $asyncResult = $tcpClient.BeginConnect($dcHost, 389, $null, $null)
            $success = $asyncResult.AsyncWaitHandle.WaitOne($timeoutMs, $false)
            if ($success -and $tcpClient.Connected) {
                $tcpClient.EndConnect($asyncResult)
                $tcpClient.Close()
                return $true
            } else {
                $tcpClient.Close()
                return $false
            }
        } catch { return $false }
    }

    # 100% funktionierende Extraktion aus ADSI COM-Objekten
    $convertLargeInt = {
        param($val)
        if ($null -eq $val) { return [int64]0 }
        if ($val -is [int64]) { return $val }
        if ($val -is [int32] -or $val -is [int] -or $val -is [double]) { return [int64]$val }
        
        # COM IADsLargeInteger per Reflection abfragen
        try {
            $type = $val.GetType()
            $high = [int64]$type.InvokeMember("HighPart", [System.Reflection.BindingFlags]::GetProperty, $null, $val, $null)
            $low  = [int64]$type.InvokeMember("LowPart", [System.Reflection.BindingFlags]::GetProperty, $null, $val, $null)
            if ($low -lt 0) { $low = [int64]([uint32]$low) }
            return [int64](($high -shl 32) + $low)
        } catch {}

        # Direkte Eigenschaftsabfrage (Powershell COM-Wrapper)
        try {
            $high = [int64]$val.HighPart
            $low  = [int64]$val.LowPart
            if ($low -lt 0) { $low = [int64]([uint32]$low) }
            return [int64](($high -shl 32) + $low)
        } catch {}

        try {
            [int64]$parsed = 0
            if ([int64]::TryParse($val.ToString(), [ref]$parsed)) { return $parsed }
        } catch {}

        return [int64]0
    }

    $formatFileTime = {
        param([int64]$ticks)
        if ($ticks -le 0 -or $ticks -ge [DateTime]::MaxValue.ToFileTime()) { return $null }
        try { return [DateTime]::FromFileTime($ticks) } catch { return $null }
    }

    # Initiales Laden der Domain Controller
    $subForm.Add_Shown({
        $lblStatus.Text = "Lade Domain Controller Sites..."
        $subForm.Refresh()
        $gridDCs.Rows.Clear()

        try {
            $dom = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
            foreach ($dc in $dom.DomainControllers) {
                $sName = if ($dc.SiteName) { $dc.SiteName } else { "Unbekannt" }
                [void]$gridDCs.Rows.Add($true, $dc.Name, $sName)
            }
            $lblStatus.Text = "Bereit. $($dom.DomainControllers.Count) Domain Controller geladen."
        } catch {
            $fallback = if ($env:LOGONSERVER) { $env:LOGONSERVER.Replace("\", "") } else { "localhost" }
            [void]$gridDCs.Rows.Add($true, $fallback, "Default-Site")
            $lblStatus.Text = "Bereit (Fallback DC)."
        }
    })

    $script:rawUserResults = @()

    $applyRowColors = {
        $ninetyDaysAgo = (Get-Date).AddDays(-90)
        foreach ($row in $gridResults.Rows) {
            $statusVal = $row.Cells["Status"].Value
            $lastLoginVal = $row.Cells["Neuester Login"].Value

            if ($statusVal -eq "Deaktiviert") {
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Gray
            } else {
                if ($lastLoginVal -and $lastLoginVal -ne "Nie" -and $lastLoginVal -ne "Offline / Nicht erreichbar") {
                    try {
                        $parsed = [datetime]::ParseExact($lastLoginVal, "dd.MM.yyyy HH:mm", $null)
                        if ($parsed -ge $ninetyDaysAgo) {
                            $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(220, 252, 231)
                        } else {
                            $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(254, 226, 226)
                        }
                    } catch {}
                } else {
                    $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(254, 226, 226)
                }
            }
        }
    }

    $applyFilterAndSort = {
        if (-not $script:rawUserResults -or $script:rawUserResults.Count -eq 0) { return }

        $statusSel = $cmbStatusFilter.SelectedIndex
        $sortProp  = $cmbSort.SelectedItem.ToString()

        $filtered = $script:rawUserResults | Where-Object {
            switch ($statusSel) {
                1 { $_.Status -eq "Aktiv" }
                2 { $_.Status -eq "Deaktiviert" }
                Default { $true }
            }
        }

        $sorted = $filtered | Sort-Object {
            if ($sortProp -eq "PW-Alter (Tage)") {
                $v = $_.$sortProp
                if ($v -is [int]) { $v } else { 99999 }
            } else {
                $_.$sortProp
            }
        }

        $arr = [System.Collections.ArrayList]::new()
        foreach ($item in $sorted) { [void]$arr.Add($item) }
        $gridResults.DataSource = $arr
        & $applyRowColors
        $lblStatus.Text = "Abgeschlossen: $($arr.Count) von $($script:rawUserResults.Count) Benutzer(n) angezeigt."
    }

    $gridResults.Add_ColumnHeaderMouseClick({
        param($sender, $e)
        $col = $gridResults.Columns[$e.ColumnIndex]
        if (-not $col) { return }
        $propName = if ($col.DataPropertyName) { $col.DataPropertyName } else { $col.HeaderText }

        if (-not $gridResults.Tag -or -not ($gridResults.Tag -is [hashtable])) {
            $gridResults.Tag = @{ LastCol = ""; Asc = $true }
        }
        $state = $gridResults.Tag
        if ($state.LastCol -eq $propName) {
            $state.Asc = -not $state.Asc
        } else {
            $state.LastCol = $propName
            $state.Asc = $true
        }

        if ($gridResults.DataSource) {
            $currentData = @($gridResults.DataSource)
            $sorted = if ($state.Asc) {
                $currentData | Sort-Object { $_.$propName }
            } else {
                $currentData | Sort-Object { $_.$propName } -Descending
            }
            $arr = [System.Collections.ArrayList]::new()
            foreach ($item in $sorted) { [void]$arr.Add($item) }
            $gridResults.DataSource = $arr
            & $applyRowColors
        }
    })

    # Abfrage durchfuehren
    $btnRun.Add_Click({
        $subForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $lblStatus.Text = "Pruefe Erreichbarkeit der ausgewaehlten Domain Controller..."
        [System.Windows.Forms.Application]::DoEvents()

        $activeDCs = @()

        # Timeout-Check (max 3s pro DC)
        foreach ($row in $gridDCs.Rows) {
            if ($row.Cells["Check"].Value -eq $true) {
                $dcHost = $row.Cells["HostName"].Value
                $lblStatus.Text = "Pruefe Verbindung zu $dcHost (max 3s)..."
                [System.Windows.Forms.Application]::DoEvents()

                $isOnline = & $testDcReachability $dcHost 3000

                if ($isOnline) {
                    $activeDCs += $dcHost
                    $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::White
                    $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Black
                } else {
                    $row.Cells["Check"].Value = $false
                    $row.Cells["Site"].Value = "Nicht erreichbar (Timeout > 3s)"
                    $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(254, 226, 226)
                    $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkRed
                }
            }
        }

        if ($activeDCs.Count -eq 0) {
            $subForm.Cursor = [System.Windows.Forms.Cursors]::Default
            [System.Windows.Forms.MessageBox]::Show("Keiner der ausgewaehlten Domain Controller konnte erreicht werden.", "Verbindungsfehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $lblStatus.Text = "Abfrage abgebrochen: Keine erreichbaren DCs."
            return
        }

        $lblStatus.Text = "Suche Benutzer im Active Directory..."
        [System.Windows.Forms.Application]::DoEvents()

        try {
            $rootEntry = [System.DirectoryServices.DirectoryEntry]::new("LDAP://RootDSE")
            $defNC = $rootEntry.defaultNamingContext.ToString()
            $now = Get-Date

            $searchPattern = $txtSearch.Text.Trim()
            if ([string]::IsNullOrEmpty($searchPattern)) { $searchPattern = "*" }

            $ldapFilter = if ($searchPattern -eq "*") {
                "(&(objectCategory=person)(objectClass=user))"
            } else {
                "(&(objectCategory=person)(objectClass=user)(|(sAMAccountName=$searchPattern)(cn=$searchPattern)))"
            }

            # 1. Benutzerstamm abfragen
            $primaryDC = $activeDCs[0]
            $domEntry = [System.DirectoryServices.DirectoryEntry]::new("LDAP://$primaryDC/$defNC")
            $searcher = [System.DirectoryServices.DirectorySearcher]::new($domEntry)
            $searcher.Filter = $ldapFilter
            $searcher.PageSize = 2500
            $searcher.PropertiesToLoad.AddRange(@("samaccountname", "displayname", "pwdlastset", "useraccountcontrol", "distinguishedname"))
            $allUsers = $searcher.FindAll()

            if (-not $allUsers -or $allUsers.Count -eq 0) {
                $lblStatus.Text = "Keine Benutzerkonten gefunden."
                $gridResults.DataSource = $null
                return
            }

            # 2. Multi-DC lastLogon Map aufbauen (nur erreichbare DCs)
            $dcLoginData = @{}
            foreach ($dc in $activeDCs) {
                $dcLoginData[$dc] = @{}
                $lblStatus.Text = "Lese LastLogon von DC: $dc..."
                [System.Windows.Forms.Application]::DoEvents()

                try {
                    $dcEntry = [System.DirectoryServices.DirectoryEntry]::new("LDAP://$dc/$defNC")
                    $dcSearcher = [System.DirectoryServices.DirectorySearcher]::new($dcEntry)
                    $dcSearcher.Filter = $ldapFilter
                    $dcSearcher.PageSize = 2500
                    $dcSearcher.PropertiesToLoad.AddRange(@("samaccountname", "lastlogon"))
                    $dcResults = $dcSearcher.FindAll()

                    foreach ($res in $dcResults) {
                        $sName = if ($res.Properties["samaccountname"].Count -gt 0) { $res.Properties["samaccountname"][0].ToString().ToLower() } else { $null }
                        if ($sName) {
                            $rawLL = if ($res.Properties["lastlogon"].Count -gt 0) { $res.Properties["lastlogon"][0] } else { 0 }
                            $ticksLL = & $convertLargeInt $rawLL
                            $dcLoginData[$dc][$sName] = $ticksLL
                        }
                    }
                } catch {}
            }

            # 3. Konsolidierung der Ergebnisse
            $resultsList = [System.Collections.ArrayList]::new()

            foreach ($u in $allUsers) {
                $p = $u.Properties
                $sam = if ($p["samaccountname"].Count -gt 0) { $p["samaccountname"][0].ToString() } else { "" }
                $disp = if ($p["displayname"].Count -gt 0) { $p["displayname"][0].ToString() } else { "" }
                $dn = if ($p["distinguishedname"].Count -gt 0) { $p["distinguishedname"][0].ToString() } else { "" }

                $ou = "Users"
                if ($dn -match 'OU=([^,]+)') { $ou = $Matches[1] }

                $uac = if ($p["useraccountcontrol"].Count -gt 0) { [int]$p["useraccountcontrol"][0] } else { 0 }
                $isDisabled = ($uac -band 2) -eq 2
                $dontExpire = ($uac -band 65536) -eq 65536

                # pwdLastSet Parsing
                $rawPwd = if ($p["pwdlastset"].Count -gt 0) { $p["pwdlastset"][0] } else { 0 }
                $pwdTicks = & $convertLargeInt $rawPwd

                $pwdDateStr = ""
                $pwdAgeDays = 9999

                if ($pwdTicks -eq 0) {
                    $pwdDateStr = "Muss geaendert werden (pwdLastSet=0)"
                } else {
                    $dtPwd = & $formatFileTime $pwdTicks
                    if ($null -ne $dtPwd) {
                        $pwdDateStr = $dtPwd.ToString("dd.MM.yyyy HH:mm")
                        $pwdAgeDays = [int][math]::Round(($now - $dtPwd).TotalDays, 0)
                    } else {
                        $pwdDateStr = "Nie"
                    }
                }

                # Multi-DC Logins ermitteln
                $samLower = $sam.ToLower()
                $maxLogonTicks = 0
                $latestDC = "N/A"
                $dcValues = [ordered]@{}

                foreach ($row in $gridDCs.Rows) {
                    $dc = $row.Cells["HostName"].Value
                    
                    if ($activeDCs -contains $dc) {
                        $llTicks = 0
                        if ($dcLoginData.ContainsKey($dc) -and $dcLoginData[$dc].ContainsKey($samLower)) {
                            $llTicks = $dcLoginData[$dc][$samLower]
                        }

                        $dtLogon = & $formatFileTime $llTicks
                        if ($null -ne $dtLogon) {
                            $dcValues[$dc] = $dtLogon.ToString("dd.MM.yyyy HH:mm")
                            if ($llTicks -gt $maxLogonTicks) {
                                $maxLogonTicks = $llTicks
                                $latestDC = $dc
                            }
                        } else {
                            $dcValues[$dc] = "Nie"
                        }
                    } else {
                        $dcValues[$dc] = "Offline / Nicht erreichbar"
                    }
                }

                $latestLoginStr = "Nie"
                $inactDays = 9999

                if ($maxLogonTicks -gt 0) {
                    $dtMax = & $formatFileTime $maxLogonTicks
                    if ($null -ne $dtMax) {
                        $latestLoginStr = $dtMax.ToString("dd.MM.yyyy HH:mm")
                        $inactDays = [int][math]::Round(($now - $dtMax).TotalDays, 0)
                    }
                }

                $rowMap = [ordered]@{
                    "Benutzername"             = $sam
                    "Anzeigename"              = $disp
                    "Status"                   = if ($isDisabled) { "Deaktiviert" } else { "Aktiv" }
                    "Letzte Kennwortaenderung" = $pwdDateStr
                    "PW-Alter (Tage)"          = if ($pwdDateStr -eq "Muss geaendert werden (pwdLastSet=0)") { 0 } else { $pwdAgeDays }
                    "PW laeuft nie ab"         = if ($dontExpire) { "Ja" } else { "Nein" }
                    "Neuester Login"           = $latestLoginStr
                    "Neuester DC"              = $latestDC
                    "Inaktiv seit (Tage)"      = if ($latestLoginStr -eq "Nie") { "Nie" } else { $inactDays }
                    "OU Pfad"                  = $ou
                }

                foreach ($k in $dcValues.Keys) {
                    $rowMap["DC: $k"] = $dcValues[$k]
                }

                [void]$resultsList.Add([PSCustomObject]$rowMap)
            }

            $script:rawUserResults = $resultsList
            & $applyFilterAndSort

        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Scan: $($_.Exception.Message)", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $lblStatus.Text = "Fehler: " + $_.Exception.Message
        } finally {
            $subForm.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    })

    $cmbStatusFilter.Add_SelectedIndexChanged({ & $applyFilterAndSort })
    $cmbSort.Add_SelectedIndexChanged({ & $applyFilterAndSort })

    $txtSearch.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $_.SuppressKeyPress = $true
            $btnRun.PerformClick()
        }
    })

    [void]$subForm.ShowDialog()
}