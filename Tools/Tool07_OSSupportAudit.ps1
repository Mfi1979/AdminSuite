<#
================================================================================
 TOOL 7: AD SECURITY & OS SUPPORT AUDIT (EOL & BUILD-ERKENNUNG)
================================================================================
#>
function Open-ToolOSSupportAudit {
    if (-not (Assert-DomainJoined)) { return }

    $subForm = New-Object System.Windows.Forms.Form
    $subForm.Text = "Tool 7: AD Security & OS Support Audit"
    $subForm.Size = New-Object System.Drawing.Size(1420, 860)
    $subForm.StartPosition = "CenterParent"
    $subForm.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9)

    $pnlTop = New-Object System.Windows.Forms.Panel
    $pnlTop.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlTop.Height = 110
    $subForm.Controls.Add($pnlTop)

    $lblType = New-Object System.Windows.Forms.Label
    $lblType.Location = "15, 15"; $lblType.Size = "80, 20"; $lblType.Text = "Systemtyp:"
    $pnlTop.Controls.Add($lblType)

    $cmbSystemType = New-Object System.Windows.Forms.ComboBox
    $cmbSystemType.Location = "100, 12"; $cmbSystemType.Size = "130, 23"
    $cmbSystemType.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$cmbSystemType.Items.AddRange(@("All / Alle", "Clients Only", "Servers Only"))
    $cmbSystemType.SelectedIndex = 0
    $pnlTop.Controls.Add($cmbSystemType)

    $lblOSFilter = New-Object System.Windows.Forms.Label
    $lblOSFilter.Location = "245, 15"; $lblOSFilter.Size = "100, 20"; $lblOSFilter.Text = "Betriebssystem:"
    $pnlTop.Controls.Add($lblOSFilter)

    $cmbOSFilter = New-Object System.Windows.Forms.ComboBox
    $cmbOSFilter.Location = "345, 12"; $cmbOSFilter.Size = "220, 23"
    $cmbOSFilter.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$cmbOSFilter.Items.Add("All OS / Alle OS")
    $cmbOSFilter.SelectedIndex = 0
    $pnlTop.Controls.Add($cmbOSFilter)

    $lblName = New-Object System.Windows.Forms.Label
    $lblName.Location = "15, 48"; $lblName.Size = "80, 20"; $lblName.Text = "Computer:"
    $pnlTop.Controls.Add($lblName)

    $txtNameSearch = New-Object System.Windows.Forms.TextBox
    $txtNameSearch.Location = "100, 45"; $txtNameSearch.Size = "130, 23"; $txtNameSearch.Text = "*"
    $pnlTop.Controls.Add($txtNameSearch)

    $lblEnabled = New-Object System.Windows.Forms.Label
    $lblEnabled.Location = "245, 48"; $lblEnabled.Size = "100, 20"; $lblEnabled.Text = "Konto-Status:"
    $pnlTop.Controls.Add($lblEnabled)

    $cmbEnabledFilter = New-Object System.Windows.Forms.ComboBox
    $cmbEnabledFilter.Location = "345, 45"; $cmbEnabledFilter.Size = "220, 23"
    $cmbEnabledFilter.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$cmbEnabledFilter.Items.AddRange(@("Alle Status", "Nur Aktivierte (True)", "Nur Deaktivierte (False)"))
    $cmbEnabledFilter.SelectedIndex = 0
    $pnlTop.Controls.Add($cmbEnabledFilter)

    $lblStatusGroup = New-Object System.Windows.Forms.Label
    $lblStatusGroup.Location = "585, 15"; $lblStatusGroup.Size = "100, 20"; $lblStatusGroup.Text = "Support-Status:"
    $pnlTop.Controls.Add($lblStatusGroup)

    $chkListStatus = New-Object System.Windows.Forms.CheckedListBox
    $chkListStatus.Location = "685, 12"; $chkListStatus.Size = "180, 80"; $chkListStatus.CheckOnClick = $true
    [void]$chkListStatus.Items.Add("Supported", $true)
    [void]$chkListStatus.Items.Add("Near EOL", $true)
    [void]$chkListStatus.Items.Add("Out of Support / EOL", $true)
    [void]$chkListStatus.Items.Add("Unbekannt / Sonstige", $true)
    $pnlTop.Controls.Add($chkListStatus)

    $chkGroupByOS = New-Object System.Windows.Forms.CheckBox
    $chkGroupByOS.Location = "100, 78"; $chkGroupByOS.Size = "250, 22"; $chkGroupByOS.Text = "Nach Operating System sortieren"
    $chkGroupByOS.Checked = $true
    $pnlTop.Controls.Add($chkGroupByOS)

    $btnRunAudit = New-Object System.Windows.Forms.Button
    $btnRunAudit.Location = "885, 12"; $btnRunAudit.Size = "140, 40"; $btnRunAudit.Text = "AD Scannen"
    $btnRunAudit.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $btnRunAudit.ForeColor = [System.Drawing.Color]::White
    $btnRunAudit.Font = New-Object System.Drawing.Font($subForm.Font.FontFamily, 9, [System.Drawing.FontStyle]::Bold)
    $pnlTop.Controls.Add($btnRunAudit)

    $lblAuditStatus = New-Object System.Windows.Forms.Label
    $lblAuditStatus.Location = "885, 58"; $lblAuditStatus.Size = "500, 30"; $lblAuditStatus.ForeColor = [System.Drawing.Color]::DarkBlue
    $pnlTop.Controls.Add($lblAuditStatus)

    $gridOSAudit = New-Object System.Windows.Forms.DataGridView
    $gridOSAudit.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridOSAudit.ReadOnly = $true
    $gridOSAudit.AllowUserToAddRows = $false
    $gridOSAudit.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
    $gridOSAudit.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    Apply-StandardGridTheme -Grid $gridOSAudit
    $subForm.Controls.Add($gridOSAudit); $gridOSAudit.BringToFront()

    # Farbgebung über RowPrePaint
    $gridOSAudit.Add_RowPrePaint({
        param($sender, $e)
        if ($e.RowIndex -ge 0 -and $e.RowIndex -lt $gridOSAudit.Rows.Count) {
            $row = $gridOSAudit.Rows[$e.RowIndex]
            $stVal = $row.Cells["Support Status"].Value
            if ($stVal -like "*Out of Support*") {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 235, 235)
            } elseif ($stVal -eq "Near EOL") {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 248, 220)
            } elseif ($stVal -like "*Supported*") {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(240, 255, 240)
            }
        }
    })

    $subForm.Add_Shown({
        $lblAuditStatus.Text = "Lade AD Computerkonten..."
        $subForm.Refresh()
        $script:RawComputersList = Search-NativeLdap -LdapFilter "(objectCategory=computer)" -PropertiesToLoad @("name","operatingSystem","operatingSystemVersion","operatingSystemServicePack","userAccountControl","distinguishedName")
        
        $osSet = [System.Collections.Generic.HashSet[string]]::new()
        foreach ($c in $script:RawComputersList) {
            $osVal = if ($c.Properties["operatingsystem"].Count -gt 0) { $c.Properties["operatingsystem"][0] } else { "Unspecified OS" }
            [void]$osSet.Add($osVal)
        }
        foreach ($osItem in ($osSet | Sort-Object)) { [void]$cmbOSFilter.Items.Add($osItem) }
        $lblAuditStatus.Text = "$($script:RawComputersList.Count) Computer geladen. Bereit."
        $btnRunAudit.PerformClick()
    })

    $btnRunAudit.Add_Click({
        if (-not $script:RawComputersList) { return }
        $lblAuditStatus.Text = "Wende Filter an..."
        $subForm.Refresh()

        $selectedOS = $cmbOSFilter.SelectedItem
        $selectedTypeIndex = $cmbSystemType.SelectedIndex
        $namePattern = $txtNameSearch.Text -replace '\*', '.*'
        $selectedEnabled = $cmbEnabledFilter.SelectedIndex
        $selectedStatuses = @($chkListStatus.CheckedItems | ForEach-Object { $_.ToString() })

        $auditResults = @()

        foreach ($c in $script:RawComputersList) {
            $cName = $c.Properties["name"][0]
            if ($namePattern -and $cName -notmatch "(?i)^$namePattern$") { continue }

            $osName = if ($c.Properties["operatingsystem"].Count -gt 0) { $c.Properties["operatingsystem"][0] } else { "Unspecified OS" }
            $osVer  = if ($c.Properties["operatingsystemversion"].Count -gt 0) { $c.Properties["operatingsystemversion"][0] } else { "" }
            $osSP   = if ($c.Properties["operatingsystemservicepack"].Count -gt 0) { $c.Properties["operatingsystemservicepack"][0] } else { "" }

            $isServer = $osName -match "Server"
            $sysTypeStr = if ($isServer) { "Server" } else { "Client" }

            if ($selectedTypeIndex -eq 1 -and $isServer) { continue }
            if ($selectedTypeIndex -eq 2 -and -not $isServer) { continue }
            if ($selectedOS -and $selectedOS -ne "All OS / Alle OS" -and $osName -ne $selectedOS) { continue }

            $isEnabled = $true
            if ($c.Properties["useraccountcontrol"].Count -gt 0) {
                if (([int]$c.Properties["useraccountcontrol"][0] -band 2) -eq 2) { $isEnabled = $false }
            }

            if ($selectedEnabled -eq 1 -and -not $isEnabled) { continue }
            if ($selectedEnabled -eq 2 -and $isEnabled) { continue }

            $eval = Analyze-OSSupportDetails -OSName $osName -OSVersion $osVer -OSServicePack $osSP
            $supportStatus = $eval.Status
            $eolDate       = $eval.EOLDate
            $clientVer     = $eval.ClientVersion
            $buildNum      = $eval.BuildNumber

            $matchesStatus = $false
            foreach ($st in $selectedStatuses) {
                if ($st -eq "Supported" -and $supportStatus -like "*Supported*") { $matchesStatus = $true; break }
                if ($st -eq "Near EOL" -and $supportStatus -eq "Near EOL") { $matchesStatus = $true; break }
                if ($st -eq "Out of Support / EOL" -and $supportStatus -like "*Out of Support*") { $matchesStatus = $true; break }
                if ($st -eq "Unbekannt / Sonstige" -and $supportStatus -notlike "*Supported*" -and $supportStatus -ne "Near EOL" -and $supportStatus -notlike "*Out of Support*") { $matchesStatus = $true; break }
            }
            if (-not $matchesStatus) { continue }

            $dnVal = $c.Properties["distinguishedname"][0]
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

        $lblAuditStatus.Text = "Ergebnis: $($arrA.Count) Computer gefunden (aus $($script:RawComputersList.Count) AD-Konten)."
    })

    [void]$subForm.ShowDialog()
}
