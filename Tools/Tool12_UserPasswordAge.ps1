<#
================================================================================
 TOOL 12: AD BENUTZER KENNWORTALTER & MULTI-DC LASTLOGON AUDIT
 Exaktes Layout & DC-Auswahl-Tabelle wie in Tool 1
================================================================================
#>

function Open-ToolUserPasswordAge {
    if (Get-Command Assert-DomainJoined -ErrorAction SilentlyContinue) {
        if (-not (Assert-DomainJoined)) { return }
    }

    $isDE = ($script:CurrentLang -eq "DE")
    $t = @{
        "Title"         = if ($isDE) { "Tool 12: AD Benutzer Kennwortalter & Multi-DC LastLogon" } else { "Tool 12: AD User Password Age & Multi-DC LastLogon" }
        "ColUser"       = if ($isDE) { "Benutzername" } else { "Username" }
        "ColName"       = if ($isDE) { "Anzeigename" } else { "Display Name" }
        "ColStatus"     = if ($isDE) { "Status" } else { "Status" }
        "ColPwdLast"    = if ($isDE) { "Letzte Kennwortaenderung" } else { "Last Password Change" }
        "ColPwdAge"     = if ($isDE) { "PW-Alter (Tage)" } else { "PW-Age (Days)" }
        "ColNeverExp"   = if ($isDE) { "PW laeuft nie ab" } else { "PW Never Expires" }
        "ColLatestLogin"= if ($isDE) { "Neuester Login" } else { "Latest Login" }
        "ColLatestDC"   = if ($isDE) { "Neuester DC" } else { "Latest DC" }
        "ColInactDays"  = if ($isDE) { "Inaktiv seit (Tage)" } else { "Inactive (Days)" }
        "ColOU"         = if ($isDE) { "OU Pfad" } else { "OU Path" }
        "LblSearch"     = if ($isDE) { "Benutzername (Wildcard):" } else { "Username (Wildcard):" }
        "BtnStart"      = if ($isDE) { "Abfragen" } else { "Query" }
        "BtnExport"     = if ($isDE) { "Export CSV" } else { "Export CSV" }
        "GrpDCs"        = if ($isDE) { "Verfuegbare Domain Controller (Wählen Sie die abzufragenden DCs aus)" } else { "Available Domain Controllers (Select DCs to query)" }
        "StatusReady"   = if ($isDE) { "Bereit." } else { "Ready." }
        "StatusScanning"= if ($isDE) { "Scanne Active Directory Benutzer & Domain Controller..." } else { "Scanning Active Directory Users & Domain Controllers..." }
        "MustChange"    = if ($isDE) { "Muss geaendert werden (pwdLastSet=0)" } else { "Must change (pwdLastSet=0)" }
        "Never"         = if ($isDE) { "Nie" } else { "Never" }
        "Active"        = if ($isDE) { "Aktiv" } else { "Enabled" }
        "Disabled"      = if ($isDE) { "Deaktiviert" } else { "Disabled" }
    }

    $subForm = New-Object System.Windows.Forms.Form
    $subForm.Text = $t["Title"]
    $subForm.Size = New-Object System.Drawing.Size(1250, 800)
    $subForm.StartPosition = "CenterParent"
    $subForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $subForm.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

    # 1. TOP PANEL (Suchfeld & Start-Button wie Tool 1)[cite: 1]
    $pnlTop = New-Object System.Windows.Forms.Panel
    $pnlTop.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlTop.Height = 75
    $pnlTop.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
    $subForm.Controls.Add($pnlTop)

    $lblSearch = New-Object System.Windows.Forms.Label
    $lblSearch.Location = New-Object System.Drawing.Point(15, 18)
    $lblSearch.Text = $t["LblSearch"]
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
    $btnRun.Text = $t["BtnStart"]
    $btnRun.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $btnRun.ForeColor = [System.Drawing.Color]::White
    $btnRun.Font = New-Object System.Drawing.Font($subForm.Font.FontFamily, 9.5, [System.Drawing.FontStyle]::Bold)
    $pnlTop.Controls.Add($btnRun)

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = New-Object System.Drawing.Point(545, 15)
    $lblStatus.Size = New-Object System.Drawing.Size(480, 45)
    $lblStatus.ForeColor = [System.Drawing.Color]::DarkBlue
    $pnlTop.Controls.Add($lblStatus)

    # 2. DC-Auswahl GroupBox mit Checkbox-Tabelle (Exakt wie Tool 1)[cite: 1]
    $grpDCs = New-Object System.Windows.Forms.GroupBox
    $grpDCs.Text = $t["GrpDCs"]
    $grpDCs.Dock = [System.Windows.Forms.DockStyle]::Top
    $grpDCs.Height = 150
    $subForm.Controls.Add($grpDCs)

    $gridDCs = New-Object System.Windows.Forms.DataGridView
    $gridDCs.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridDCs.AllowUserToAddRows = $false
    $gridDCs.AllowUserToDeleteRows = $false
    $gridDCs.RowHeadersVisible = $false
    $gridDCs.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $gridDCs.BackgroundColor = [System.Drawing.Color]::White
    $grpDCs.Controls.Add($gridDCs)

    $colCheck = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $colCheck.Name = "Check"
    $colCheck.HeaderText = "Abfragen"
    $colCheck.Width = 70
    [void]$gridDCs.Columns.Add($colCheck)
    [void]$gridDCs.Columns.Add("HostName", "Domain Controller")
    [void]$gridDCs.Columns.Add("Site", "AD Site")

    # 3. BOTTOM PANEL (Export Button)
    $pnlBottom = New-Object System.Windows.Forms.Panel
    $pnlBottom.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $pnlBottom.Height = 45
    $pnlBottom.BackColor = [System.Drawing.Color]::White
    $subForm.Controls.Add($pnlBottom)

    $btnExport = New-Object System.Windows.Forms.Button
    $btnExport.Text = $t["BtnExport"]
    $btnExport.Size = New-Object System.Drawing.Size(130, 28)
    $btnExport.Location = New-Object System.Drawing.Point(1090, 8)
    $btnExport.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
    $btnExport.BackColor = [System.Drawing.Color]::FromArgb(37, 99, 235)
    $btnExport.ForeColor = [System.Drawing.Color]::White
    $btnExport.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnExport.FlatAppearance.BorderSize = 0
    $pnlBottom.Controls.Add($btnExport)

    # 4. Untere Ergebnistabelle (DataGridView)[cite: 1]
    $gridResults = New-Object System.Windows.Forms.DataGridView
    $gridResults.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridResults.ReadOnly = $true
    $gridResults.AllowUserToAddRows = $false
    $gridResults.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
    $gridResults.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $gridResults.BackgroundColor = [System.Drawing.Color]::White
    $subForm.Controls.Add($gridResults)

    # Z-Order Korrektur
    $gridResults.BringToFront()
    $grpDCs.SendToBack()
    $pnlTop.SendToBack()
    $pnlBottom.BringToFront()

    # Universeller Spaltenkopf-Sortierer
    $gridResults.Add_ColumnHeaderMouseClick({
        param($sender, $e)
        $col = $gridResults.Columns[$e.ColumnIndex]
        if (-not $col) { return }

        $propName = if ($col.DataPropertyName) { $col.DataPropertyName } else { $col.HeaderText }
        if (-not $propName) { return }

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

        if ($script:rawScanResults -and $script:rawScanResults.Count -gt 0) {
            $sorted = if ($state.Asc) {
                $script:rawScanResults | Sort-Object { $_.$propName }
            } else {
                $script:rawScanResults | Sort-Object { $_.$propName } -Descending
            }

            $arr = [System.Collections.ArrayList]::new()
            foreach ($item in $sorted) { [void]$arr.Add($item) }
            $gridResults.DataSource = $arr
            & $applyRowColors
        }
    })

    # Hilfsfunktion: LargeInteger / ComObject Konverter
    $convertLargeInt = {
        param($val)
        if ($null -eq $val) { return 0 }
        try {
            if ($val -is [int64] -or $val -is [int] -or $val -is [double]) { return [int64]$val }
            if ($val.GetType().Name -match "ComObject|__ComObject|LargeInteger") {
                $h = [int64]$val.GetType().InvokeMember("HighPart", [System.Reflection.BindingFlags]::GetProperty, $null, $val, $null)
                $l = [int64]$val.GetType().InvokeMember("LowPart", [System.Reflection.BindingFlags]::GetProperty, $null, $val, $null)
                return (($h -shl 32) + ($l -band 0xFFFFFFFF))
            }
            $parsed = 0
            if ([int64]::TryParse($val.ToString(), [ref]$parsed)) { return $parsed }
            return 0
        } catch { return 0 }
    }

    # DCs beim Start laden (wie in Tool 1)[cite: 2]
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
            $gridDCs.Rows.Add($true, $env:LOGONSERVER.Replace("\", ""), "Default-Site")
        }
    })

    $script:rawScanResults = @()

    # Farb-Hervorhebung
    $applyRowColors = {
        $ninetyDaysAgo = (Get-Date).AddDays(-90)
        foreach ($row in $gridResults.Rows) {
            $statusVal = $row.Cells[$t["ColStatus"]].Value
            $lastLoginVal = $row.Cells[$t["ColLatestLogin"]].Value

            if ($statusVal -eq $t["Disabled"]) {
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Gray
            } else {
                if ($lastLoginVal -and $lastLoginVal -ne $t["Never"]) {
                    try {
                        $parsed = [datetime]::ParseExact($lastLoginVal, "dd.MM.yyyy HH:mm:ss", $null)
                        if ($parsed -ge $ninetyDaysAgo) {
                            $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(236, 253, 245) # Hellgruen
                        } else {
                            $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(254, 242, 242) # Hellrot
                        }
                    } catch {}
                } else {
                    $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(254, 242, 242)
                }
            }
        }
    }

    # Multi-DC Scan Engine
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

        $lblStatus.Text = $t["StatusScanning"]
        $subForm.Refresh()

        try {
            $rootEntry = [System.DirectoryServices.DirectoryEntry]::new("LDAP://RootDSE")
            $defNC = $rootEntry.defaultNamingContext.ToString()
            $now = Get-Date
            $searchPattern = $txtSearch.Text.Trim()
            if ([string]::IsNullOrEmpty($searchPattern)) { $searchPattern = "*" }

            # 1. Benutzerstamm abfragen
            $primaryDC = $selectedDCs[0]
            $domEntry = [System.DirectoryServices.DirectoryEntry]::new("LDAP://$primaryDC/$defNC")
            $searcher = [System.DirectoryServices.DirectorySearcher]::new($domEntry)
            
            if ($searchPattern -eq "*") {
                $searcher.Filter = "(&(objectCategory=person)(objectClass=user))"
            } else {
                $searcher.Filter = "(&(objectCategory=person)(objectClass=user)(|(sAMAccountName=$searchPattern)(cn=$searchPattern)))"
            }
            $searcher.PageSize = 2000
            $searcher.PropertiesToLoad.AddRange(@("samaccountname", "displayname", "pwdlastset", "useraccountcontrol", "distinguishedname"))
            $allUsers = $searcher.FindAll()

            # 2. Multi-DC lastLogon Map aufbauen
            $dcLoginData = @{}
            foreach ($dc in $selectedDCs) {
                $dcLoginData[$dc] = @{}
                try {
                    $dcEntry = [System.DirectoryServices.DirectoryEntry]::new("LDAP://$dc/$defNC")
                    $dcSearcher = [System.DirectoryServices.DirectorySearcher]::new($dcEntry)
                    $dcSearcher.Filter = "(&(objectCategory=person)(objectClass=user))"
                    $dcSearcher.PageSize = 2000
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

            # 3. Aggregation
            $resultsList = [System.Collections.ArrayList]::new()
            $i = 0

            foreach ($u in $allUsers) {
                $i++
                $p = $u.Properties
                $sam = if ($p["samaccountname"].Count -gt 0) { $p["samaccountname"][0].ToString() } else { "" }
                $disp = if ($p["displayname"].Count -gt 0) { $p["displayname"][0].ToString() } else { "" }
                $dn = if ($p["distinguishedname"].Count -gt 0) { $p["distinguishedname"][0].ToString() } else { "" }

                if ($i % 50 -eq 0) {
                    $lblStatus.Text = "Verarbeite Benutzer... ($i von $($allUsers.Count))"
                    [System.Windows.Forms.Application]::DoEvents()
                }

                $ou = "Users"
                if ($dn -match 'OU=([^,]+)') { $ou = $Matches[1] }

                $uac = if ($p["useraccountcontrol"].Count -gt 0) { [int]$p["useraccountcontrol"][0] } else { 0 }
                $isDisabled = ($uac -band 2) -eq 2
                $dontExpire = ($uac -band 65536) -eq 65536

                $rawPwd = if ($p["pwdlastset"].Count -gt 0) { $p["pwdlastset"][0] } else { 0 }
                $pwdTicks = & $convertLargeInt $rawPwd

                $pwdDateStr = ""
                $pwdAgeDays = -1

                if ($pwdTicks -eq 0) {
                    $pwdDateStr = $t["MustChange"]
                    $pwdAgeDays = 9999
                } elseif ($pwdTicks -lt 0 -or $pwdTicks -ge [DateTime]::MaxValue.ToFileTime()) {
                    $pwdDateStr = $t["Never"]
                    $pwdAgeDays = 9999
                } else {
                    $dtPwd = [DateTime]::FromFileTime($pwdTicks)
                    $pwdDateStr = $dtPwd.ToString("dd.MM.yyyy HH:mm:ss")
                    $pwdAgeDays = [math]::Round(($now - $dtPwd).TotalDays, 0)
                }

                $samLower = $sam.ToLower()
                $maxLogonTicks = 0
                $latestDC = "N/A"
                $dcValues = [ordered]@{}

                foreach ($dc in $selectedDCs) {
                    $llTicks = 0
                    if ($dcLoginData.ContainsKey($dc) -and $dcLoginData[$dc].ContainsKey($samLower)) {
                        $llTicks = $dcLoginData[$dc][$samLower]
                    }

                    if ($llTicks -gt 0 -and $llTicks -lt [DateTime]::MaxValue.ToFileTime()) {
                        $dtLogon = [DateTime]::FromFileTime($llTicks)
                        $dcValues[$dc] = $dtLogon.ToString("dd.MM.yyyy HH:mm:ss")
                        if ($llTicks -gt $maxLogonTicks) {
                            $maxLogonTicks = $llTicks
                            $latestDC = $dc
                        }
                    } else {
                        $dcValues[$dc] = $t["Never"]
                    }
                }

                $latestLoginStr = $t["Never"]
                $inactDays = 9999

                if ($maxLogonTicks -gt 0) {
                    $dtMax = [DateTime]::FromFileTime($maxLogonTicks)
                    $latestLoginStr = $dtMax.ToString("dd.MM.yyyy HH:mm:ss")
                    $inactDays = [math]::Round(($now - $dtMax).TotalDays, 0)
                }

                $rowMap = [ordered]@{
                    $t["ColUser"]        = $sam
                    $t["ColName"]        = $disp
                    $t["ColStatus"]      = if ($isDisabled) { $t["Disabled"] } else { $t["Active"] }
                    $t["ColPwdLast"]     = $pwdDateStr
                    $t["ColPwdAge"]      = [int]$pwdAgeDays
                    $t["ColNeverExp"]    = if ($dontExpire) { "Ja" } else { "Nein" }
                    $t["ColLatestLogin"] = $latestLoginStr
                    $t["ColLatestDC"]    = $latestDC
                    $t["ColInactDays"]   = if ($latestLoginStr -eq $t["Never"]) { "Nie" } else { [int]$inactDays }
                    $t["ColOU"]          = $ou
                }

                foreach ($k in $dcValues.Keys) {
                    $rowMap["DC: $k"] = $dcValues[$k]
                }

                [void]$resultsList.Add([PSCustomObject]$rowMap)
            }

            $script:rawScanResults = $resultsList
            $arr = [System.Collections.ArrayList]::new()
            foreach ($item in $resultsList) { [void]$arr.Add($item) }
            $gridResults.DataSource = $arr
            & $applyRowColors

            $lblStatus.Text = "Abgeschlossen: $($resultsList.Count) Benutzerkonten über $($selectedDCs.Count) DC(s) abgefragt."

        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Scan: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $lblStatus.Text = "Fehler: " + $_.Exception.Message
        }
    })

    $txtSearch.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $_.SuppressKeyPress = $true
            $btnRun.PerformClick()
        }
    })

    $btnExport.Add_Click({
        if (-not $gridResults.DataSource) { return }
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "CSV-Datei (*.csv)|*.csv"
        $sfd.FileName = "AD_User_Password_MultiDC_Audit_$((Get-Date).ToString('yyyyMMdd')).csv"
        if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $gridResults.DataSource | Export-Csv -Path $sfd.FileName -NoTypeInformation -Delimiter ";" -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show("Export erfolgreich:`n$($sfd.FileName)", "Export OK", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    })

    [void]$subForm.ShowDialog()
}