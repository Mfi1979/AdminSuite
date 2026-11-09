<#
================================================================================
 TOOL 7: AD SECURITY & OS SUPPORT AUDIT (EOL & BUILD-ERKENNUNG)
 HELPER: DYNAMIC OS SUPPORT & EOL EVALUATION (INKL. NEAR-EOL-MONATE PARAMETER)
================================================================================
#>
function Analyze-OSSupportDetails {
    param(
        [string]$OSName,
        [string]$OSVersion,
        [string]$OSServicePack,
        [int]$NearEolMonths = 12
    )

    $status        = "Supported"
    $eolDateString = "Unbekannt"
    $clientVersion = "N/A"
    $buildNumber   = "Unbekannt"
    $targetDate    = $null

    # 1. BUILD-NUMMER EXTRAHIEREN
    if ($OSVersion -match "(\d{5})") {
        $buildNumber = $Matches[1]
    } elseif ($OSVersion -match "10\.0\.(\d+)") {
        $buildNumber = $Matches[1]
    } elseif ($OSVersion) {
        $buildNumber = $OSVersion
    }

    # 2. FEATURE UPDATE & CLIENT VERSION ZUORDNEN
    if ($buildNumber -ne "Unbekannt") {
        $buildInt = 0
        [int]::TryParse($buildNumber, [ref]$buildInt) | Out-Null

        $isLTSC = ($OSName -like "*LTSC*" -or $OSName -like "*LTSB*" -or $OSServicePack -like "*LTSC*" -or $OSServicePack -like "*LTSB*")

        switch ($buildInt) {
            # Windows 11
            26300 { $clientVersion = "26H2" }
            26200 { $clientVersion = "25H2" }
            26100 { 
                if ($OSName -like "*Server*") { $clientVersion = "Server 2025" }
                elseif ($isLTSC) { $clientVersion = "24H2 / LTSC 2024" }
                else { $clientVersion = "24H2" }
            }
            22631 { $clientVersion = "23H2" }
            22621 { $clientVersion = "22H2" }
            22000 { $clientVersion = "21H2" }

            # Windows 10
            19045 { $clientVersion = "22H2" }
            19044 { if ($isLTSC) { $clientVersion = "21H2 / LTSC 2021" } else { $clientVersion = "21H2" } }
            19043 { $clientVersion = "21H1" }
            19042 { $clientVersion = "20H2" }
            19041 { $clientVersion = "2004" }
            18363 { $clientVersion = "1909" }
            17763 { 
                if ($OSName -like "*Server*") { $clientVersion = "Server 2019" } 
                else { $clientVersion = if ($isLTSC) { "1809 / LTSC 2019" } else { "1809" } } 
            }
            14393 { 
                if ($OSName -like "*Server*") { $clientVersion = "Server 2016" } 
                else { $clientVersion = if ($isLTSC) { "1607 / LTSB 2016" } else { "1607" } } 
            }
            10240 { $clientVersion = "1507 / LTSB 2015" }

            # Windows Server
            20348 { $clientVersion = "Server 2022" }

            default {
                if ($buildInt -gt 26300) { $clientVersion = "Insider Build ($buildInt)" }
                elseif ($OSVersion) { $clientVersion = $OSVersion }
                else { $clientVersion = "Build $buildNumber" }
            }
        }
    } else {
        if ($OSServicePack) { $clientVersion = $OSServicePack }
        elseif ($OSVersion) { $clientVersion = $OSVersion }
    }

    # Editionen prüfen (Enterprise / Education erhalten 36 Monate Support)
    $isEnterpriseOrEdu = ($OSName -like "*Enterprise*" -or $OSName -like "*Education*")
    $isLTSCOrLTSB     = ($OSName -like "*LTSC*" -or $OSName -like "*LTSB*" -or $clientVersion -like "*LTSC*" -or $clientVersion -like "*LTSB*")

    # 3. EOL-DATUM PRÜFEN
    if ($OSName -like "*Windows 11*") {
        if ($clientVersion -like "*21H2*" -or $buildNumber -eq "22000") {
            $dateStr = if ($isEnterpriseOrEdu) { "08.10.2024" } else { "10.10.2023" }
            $eolDateString = $dateStr; $targetDate = [datetime]::ParseExact($dateStr, "dd.MM.yyyy", $null)
        } 
        elseif ($clientVersion -like "*22H2*" -or $buildNumber -eq "22621") {
            $dateStr = if ($isEnterpriseOrEdu) { "14.10.2025" } else { "08.10.2024" }
            $eolDateString = $dateStr; $targetDate = [datetime]::ParseExact($dateStr, "dd.MM.yyyy", $null)
        } 
        elseif ($clientVersion -like "*23H2*" -or $buildNumber -eq "22631") {
            $dateStr = if ($isEnterpriseOrEdu) { "10.11.2026" } else { "11.11.2025" }
            $eolDateString = $dateStr; $targetDate = [datetime]::ParseExact($dateStr, "dd.MM.yyyy", $null)
        } 
        elseif ($clientVersion -like "*24H2*" -or $buildNumber -eq "26100") {
            if ($isLTSCOrLTSB) {
                $dateStr = if ($OSName -like "*IoT*") { "10.10.2034" } else { "09.10.2029" }
            } else {
                $dateStr = if ($isEnterpriseOrEdu) { "12.10.2027" } else { "13.10.2026" }
            }
            $eolDateString = $dateStr; $targetDate = [datetime]::ParseExact($dateStr, "dd.MM.yyyy", $null)
        } 
        elseif ($clientVersion -like "*25H2*" -or $buildNumber -eq "26200") {
            $dateStr = if ($isEnterpriseOrEdu) { "10.10.2028" } else { "12.10.2027" }
            $eolDateString = $dateStr; $targetDate = [datetime]::ParseExact($dateStr, "dd.MM.yyyy", $null)
        }
        elseif ($clientVersion -like "*26H2*" -or $buildNumber -eq "26300") {
            $dateStr = if ($isEnterpriseOrEdu) { "09.10.2029" } else { "10.10.2028" }
            $eolDateString = $dateStr; $targetDate = [datetime]::ParseExact($dateStr, "dd.MM.yyyy", $null)
        } 
        else {
            $eolDateString = "Support aktiv"
        }
    }
    elseif ($OSName -like "*Windows 10*") {
        if ($isLTSCOrLTSB) {
            if ($buildNumber -eq "19044" -or $clientVersion -like "*2021*") {
                $dateStr = if ($OSName -like "*IoT*") { "13.01.2032" } else { "12.01.2027" }
                $eolDateString = $dateStr; $targetDate = [datetime]::ParseExact($dateStr, "dd.MM.yyyy", $null)
            }
            elseif ($buildNumber -eq "17763" -or $clientVersion -like "*2019*") {
                $eolDateString = "09.01.2029"; $targetDate = [datetime]::ParseExact("09.01.2029", "dd.MM.yyyy", $null)
            }
            elseif ($buildNumber -eq "14393" -or $clientVersion -like "*2016*") {
                $eolDateString = "13.10.2026"; $targetDate = [datetime]::ParseExact("13.10.2026", "dd.MM.yyyy", $null)
            }
            elseif ($buildNumber -eq "10240" -or $clientVersion -like "*2015*") {
                $eolDateString = "14.10.2025"; $targetDate = [datetime]::ParseExact("14.10.2025", "dd.MM.yyyy", $null)
            }
            else {
                $eolDateString = "12.01.2027"; $targetDate = [datetime]::ParseExact("12.01.2027", "dd.MM.yyyy", $null)
            }
        } else {
            $eolDateString = "14.10.2025"; $targetDate = [datetime]::ParseExact("14.10.2025", "dd.MM.yyyy", $null)
        }
    }
    elseif ($OSName -like "*Windows 7*" -or $OSName -like "*Windows 8*" -or $OSName -like "*Windows XP*" -or $OSName -like "*Windows Vista*") {
        $eolDateString = "Ausgelaufen"
        $status = "Out of Support / EOL"
    }
    elseif ($OSName -like "*Server 2003*" -or $OSName -like "*Server 2008*") {
        $eolDateString = "Ausgelaufen"
        $status = "Out of Support / EOL"
    }
    elseif ($OSName -like "*Server 2012*") {
        $eolDateString = "10.10.2023"; $targetDate = [datetime]::ParseExact("10.10.2023", "dd.MM.yyyy", $null)
    }
    elseif ($OSName -like "*Server 2016*") {
        $eolDateString = "12.01.2027"; $targetDate = [datetime]::ParseExact("12.01.2027", "dd.MM.yyyy", $null)
    }
    elseif ($OSName -like "*Server 2019*") {
        $eolDateString = "09.01.2029"; $targetDate = [datetime]::ParseExact("09.01.2029", "dd.MM.yyyy", $null)
    }
    elseif ($OSName -like "*Server 2022*") {
        $eolDateString = "14.10.2031"; $targetDate = [datetime]::ParseExact("14.10.2031", "dd.MM.yyyy", $null)
    }
    elseif ($OSName -like "*Server 2025*") {
        $eolDateString = "14.10.2034"; $targetDate = [datetime]::ParseExact("14.10.2034", "dd.MM.yyyy", $null)
    }

    # 4. STATUS BERECHNUNG (Near EOL = variabel anhand des Parameters NearEolMonths)
    if ($targetDate) {
        $today = Get-Date
        if ($today -ge $targetDate) {
            $status = "Out of Support / EOL"
        } elseif ($today.AddMonths($NearEolMonths) -ge $targetDate) {
            $status = "Near EOL"
        } else {
            $status = "Supported"
        }
    }

    return [PSCustomObject]@{
        ClientVersion = $clientVersion
        BuildNumber   = $buildNumber
        Status        = $status
        EOLDate       = $eolDateString
    }
}

<#
================================================================================
 TOOL 7: AD SECURITY & OS SUPPORT AUDIT (MIT DEFINIERBAREN NEAR-EOL-MONATEN)
================================================================================
#>
function Open-ToolOSSupportAudit {
    if (-not (Assert-DomainJoined)) { return }

    $subForm = New-Object System.Windows.Forms.Form
    $subForm.Text = "Tool 7: AD Security & OS Support Audit"
    $subForm.Size = New-Object System.Drawing.Size(1440, 860)
    $subForm.StartPosition = "CenterParent"
    $subForm.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9)

    $pnlTop = New-Object System.Windows.Forms.Panel
    $pnlTop.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlTop.Height = 110
    $subForm.Controls.Add($pnlTop)

    # 1. Systemtyp Filter
    $lblType = New-Object System.Windows.Forms.Label
    $lblType.Location = "15, 15"; $lblType.Size = "80, 20"; $lblType.Text = "Systemtyp:"
    $pnlTop.Controls.Add($lblType)

    $cmbSystemType = New-Object System.Windows.Forms.ComboBox
    $cmbSystemType.Location = "100, 12"; $cmbSystemType.Size = "130, 23"
    $cmbSystemType.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$cmbSystemType.Items.AddRange(@("All / Alle", "Clients Only", "Servers Only"))
    $cmbSystemType.SelectedIndex = 0
    $pnlTop.Controls.Add($cmbSystemType)

    # 2. Betriebssystem Dropdown
    $lblOSFilter = New-Object System.Windows.Forms.Label
    $lblOSFilter.Location = "245, 15"; $lblOSFilter.Size = "100, 20"; $lblOSFilter.Text = "Betriebssystem:"
    $pnlTop.Controls.Add($lblOSFilter)

    $cmbOSFilter = New-Object System.Windows.Forms.ComboBox
    $cmbOSFilter.Location = "345, 12"; $cmbOSFilter.Size = "220, 23"
    $cmbOSFilter.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$cmbOSFilter.Items.Add("All OS / Alle OS")
    $cmbOSFilter.SelectedIndex = 0
    $pnlTop.Controls.Add($cmbOSFilter)

    # 3. Computername (Wildcard)
    $lblName = New-Object System.Windows.Forms.Label
    $lblName.Location = "15, 48"; $lblName.Size = "80, 20"; $lblName.Text = "Computer:"
    $pnlTop.Controls.Add($lblName)

    $txtNameSearch = New-Object System.Windows.Forms.TextBox
    $txtNameSearch.Location = "100, 45"; $txtNameSearch.Size = "130, 23"; $txtNameSearch.Text = "*"
    $pnlTop.Controls.Add($txtNameSearch)

    # 4. Konto-Status (Enabled)
    $lblEnabled = New-Object System.Windows.Forms.Label
    $lblEnabled.Location = "245, 48"; $lblEnabled.Size = "100, 20"; $lblEnabled.Text = "Konto-Status:"
    $pnlTop.Controls.Add($lblEnabled)

    $cmbEnabledFilter = New-Object System.Windows.Forms.ComboBox
    $cmbEnabledFilter.Location = "345, 45"; $cmbEnabledFilter.Size = "220, 23"
    $cmbEnabledFilter.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$cmbEnabledFilter.Items.AddRange(@("Alle Status", "Nur Aktivierte (True)", "Nur Deaktivierte (False)"))
    $cmbEnabledFilter.SelectedIndex = 0
    $pnlTop.Controls.Add($cmbEnabledFilter)

    # 5. Support-Status Mehrfachauswahl
    $lblStatusGroup = New-Object System.Windows.Forms.Label
    $lblStatusGroup.Location = "585, 15"; $lblStatusGroup.Size = "95, 20"; $lblStatusGroup.Text = "Support-Status:"
    $pnlTop.Controls.Add($lblStatusGroup)

    $chkListStatus = New-Object System.Windows.Forms.CheckedListBox
    $chkListStatus.Location = "685, 12"; $chkListStatus.Size = "165, 80"; $chkListStatus.CheckOnClick = $true
    [void]$chkListStatus.Items.Add("Supported", $true)
    [void]$chkListStatus.Items.Add("Near EOL", $true)
    [void]$chkListStatus.Items.Add("Out of Support / EOL", $true)
    [void]$chkListStatus.Items.Add("Unbekannt / Sonstige", $true)
    $pnlTop.Controls.Add($chkListStatus)

    # 6. NEU: Near EOL Schwellenwert in Monaten
    $lblNearMonths = New-Object System.Windows.Forms.Label
    $lblNearMonths.Location = "865, 15"; $lblNearMonths.Size = "115, 20"; $lblNearMonths.Text = "Near EOL (Monate):"
    $pnlTop.Controls.Add($lblNearMonths)

    $numNearMonths = New-Object System.Windows.Forms.NumericUpDown
    $numNearMonths.Location = "985, 12"; $numNearMonths.Size = "60, 23"
    $numNearMonths.Minimum = 1
    $numNearMonths.Maximum = 60
    $numNearMonths.Value = 12
    $pnlTop.Controls.Add($numNearMonths)

    # 7. Checkbox Sortieren
    $chkGroupByOS = New-Object System.Windows.Forms.CheckBox
    $chkGroupByOS.Location = "100, 78"; $chkGroupByOS.Size = "250, 22"; $chkGroupByOS.Text = "Nach Operating System sortieren"
    $chkGroupByOS.Checked = $true
    $pnlTop.Controls.Add($chkGroupByOS)

    # 8. Scan Button
    $btnRunAudit = New-Object System.Windows.Forms.Button
    $btnRunAudit.Location = "1065, 10"; $btnRunAudit.Size = "130, 35"
    $btnRunAudit.Text = "AD Scannen"
    $btnRunAudit.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $btnRunAudit.ForeColor = [System.Drawing.Color]::White
    $btnRunAudit.Font = New-Object System.Drawing.Font($subForm.Font.FontFamily, 9, [System.Drawing.FontStyle]::Bold)
    $pnlTop.Controls.Add($btnRunAudit)

    # Status Label
    $lblAuditStatus = New-Object System.Windows.Forms.Label
    $lblAuditStatus.Location = "865, 52"; $lblAuditStatus.Size = "540, 45"; $lblAuditStatus.ForeColor = [System.Drawing.Color]::DarkBlue
    $pnlTop.Controls.Add($lblAuditStatus)

    # DataGridView Result Panel
    $gridOSAudit = New-Object System.Windows.Forms.DataGridView
    $gridOSAudit.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridOSAudit.ReadOnly = $true
    $gridOSAudit.AllowUserToAddRows = $false
    $gridOSAudit.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
    $gridOSAudit.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    Apply-StandardGridTheme -Grid $gridOSAudit
    $subForm.Controls.Add($gridOSAudit)
    $gridOSAudit.BringToFront()

    # Farbgebung über RowPrePaint
    $gridOSAudit.Add_RowPrePaint({
        param($sender, $e)
        if ($e.RowIndex -ge 0 -and $e.RowIndex -lt $gridOSAudit.Rows.Count) {
            $row = $gridOSAudit.Rows[$e.RowIndex]
            $stVal = [string]$row.Cells["Support Status"].Value
            if ($stVal -like "*Out of Support*") {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 235, 235)
            } elseif ($stVal -eq "Near EOL") {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 248, 220)
            } elseif ($stVal -like "*Supported*") {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(240, 255, 240)
            }
        }
    })

    # FILTER-AKTION
    $applyFilterAction = {
        if (-not $script:RawComputersList) { return }
        $lblAuditStatus.Text = "Wende Filter an..."
        $subForm.Refresh()

        $selectedOS = $cmbOSFilter.SelectedItem
        $selectedTypeIndex = $cmbSystemType.SelectedIndex
        $nameSearchRaw = $txtNameSearch.Text.Trim()
        $selectedEnabled = $cmbEnabledFilter.SelectedIndex
        $selectedStatuses = @($chkListStatus.CheckedItems | ForEach-Object { $_.ToString() })
        $customNearMonths = [int]$numNearMonths.Value

        # Regex für Wildcard-Suche
        $regexPattern = if ([string]::IsNullOrWhiteSpace($nameSearchRaw) -or $nameSearchRaw -eq "*") {
            ".*"
        } else {
            "^" + [regex]::Escape($nameSearchRaw).Replace("\*", ".*").Replace("\?", ".") + "$"
        }

        $auditResults = @()

        foreach ($c in $script:RawComputersList) {
            $cName = [string]$c.Properties["name"][0]
            
            # Filter: Computername
            if ($cName -notmatch "(?i)$regexPattern") { continue }

            $osName = if ($c.Properties["operatingsystem"].Count -gt 0) { [string]$c.Properties["operatingsystem"][0] } else { "Unspecified OS" }
            $osVer  = if ($c.Properties["operatingsystemversion"].Count -gt 0) { [string]$c.Properties["operatingsystemversion"][0] } else { "" }
            $osSP   = if ($c.Properties["operatingsystemservicepack"].Count -gt 0) { [string]$c.Properties["operatingsystemservicepack"][0] } else { "" }

            $isServer = ($osName -match "Server")
            $sysTypeStr = if ($isServer) { "Server" } else { "Client" }

            # Filter: Systemtyp
            if ($selectedTypeIndex -eq 1 -and $isServer) { continue }
            if ($selectedTypeIndex -eq 2 -and -not $isServer) { continue }

            # Filter: OS Dropdown
            if ($selectedOS -and $selectedOS -ne "All OS / Alle OS" -and $osName -ne $selectedOS) { continue }

            # UserAccountControl-Auswertung (Bit 2 = ACCOUNTDISABLE)
            $uacVal = 0
            if ($c.Properties["useraccountcontrol"].Count -gt 0) {
                try {
                    $uacVal = [Convert]::ToInt32($c.Properties["useraccountcontrol"][0])
                } catch {
                    $uacVal = [int]$c.Properties["useraccountcontrol"][0]
                }
            }
            $isEnabled = (($uacVal -band 2) -ne 2)

            # Filter: Konto-Status
            if ($selectedEnabled -eq 1 -and -not $isEnabled) { continue }
            if ($selectedEnabled -eq 2 -and $isEnabled) { continue }

            # Support-Details analysieren (mit dynamischem Monatsfenster)
            $eval = Analyze-OSSupportDetails -OSName $osName -OSVersion $osVer -OSServicePack $osSP -NearEolMonths $customNearMonths
            $supportStatus = $eval.Status
            $eolDate       = $eval.EOLDate
            $clientVer     = $eval.ClientVersion
            $buildNum      = $eval.BuildNumber

            # Filter: Support-Status
            $matchesStatus = $false
            foreach ($st in $selectedStatuses) {
                if ($st -eq "Supported" -and $supportStatus -like "*Supported*") { $matchesStatus = $true; break }
                if ($st -eq "Near EOL" -and $supportStatus -eq "Near EOL") { $matchesStatus = $true; break }
                if ($st -eq "Out of Support / EOL" -and $supportStatus -like "*Out of Support*") { $matchesStatus = $true; break }
                if ($st -eq "Unbekannt / Sonstige" -and $supportStatus -notlike "*Supported*" -and $supportStatus -ne "Near EOL" -and $supportStatus -notlike "*Out of Support*") { $matchesStatus = $true; break }
            }
            if (-not $matchesStatus) { continue }

            $dnVal = [string]$c.Properties["distinguishedname"][0]
            $ouPath = if ($dnVal -match "OU=.*") { $dnVal.Substring($dnVal.IndexOf("OU=")) } else { $dnVal }

            $auditResults += [PSCustomObject]@{
                "Operating System"        = $osName
                "Version / Release"       = $clientVer
                "Build Number"            = $buildNum
                "Computer Name"           = $cName
                "System Type"             = $sysTypeStr
                "Support Status"          = $supportStatus
                "EOL Date / Support-Ende" = $eolDate
                "Enabled"                 = $isEnabled
                "OU Path"                 = $ouPath
            }
        }

        $sortedResults = if ($chkGroupByOS.Checked) {
            $auditResults | Sort-Object "Operating System", "Computer Name"
        } else {
            $auditResults | Sort-Object "Computer Name"
        }

        $arrA = [System.Collections.ArrayList]::new()
        foreach ($item in $sortedResults) { [void]$arrA.Add($item) }

        $gridOSAudit.DataSource = $null
        $gridOSAudit.DataSource = $arrA

        $lblAuditStatus.Text = "Ergebnis: $($arrA.Count) Computer gefunden (Near EOL Schwelle: $customNearMonths Monate)."
    }

    # Event-Verdrahtung
    $btnRunAudit.Add_Click({ & $applyFilterAction })
    $cmbEnabledFilter.Add_SelectedIndexChanged({ & $applyFilterAction })
    $cmbSystemType.Add_SelectedIndexChanged({ & $applyFilterAction })
    $cmbOSFilter.Add_SelectedIndexChanged({ & $applyFilterAction })
    $chkGroupByOS.Add_CheckedChanged({ & $applyFilterAction })
    $chkListStatus.Add_SelectedIndexChanged({ & $applyFilterAction })
    $numNearMonths.Add_ValueChanged({ & $applyFilterAction })

    $txtNameSearch.Add_KeyDown({
        param($s, $e)
        if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $e.SuppressKeyPress = $true
            & $applyFilterAction
        }
    })

    # Initiales Laden
    $subForm.Add_Shown({
        $lblAuditStatus.Text = "Lade AD Computerkonten..."
        $subForm.Refresh()
        $script:RawComputersList = Search-NativeLdap -LdapFilter "(objectCategory=computer)" -PropertiesToLoad @("name","operatingSystem","operatingSystemVersion","operatingSystemServicePack","userAccountControl","distinguishedName")
        
        $osSet = [System.Collections.Generic.HashSet[string]]::new()
        foreach ($c in $script:RawComputersList) {
            $osVal = if ($c.Properties["operatingsystem"].Count -gt 0) { [string]$c.Properties["operatingsystem"][0] } else { "Unspecified OS" }
            [void]$osSet.Add($osVal)
        }
        foreach ($osItem in ($osSet | Sort-Object)) { [void]$cmbOSFilter.Items.Add($osItem) }
        $lblAuditStatus.Text = "$($script:RawComputersList.Count) Computer geladen. Bereit."
        & $applyFilterAction
    })

    [void]$subForm.ShowDialog()
}
