<#
================================================================================
 TOOL 10: ACTIVE DIRECTORY OU & GRUPPEN FINDER (INKL. ADMINCOUNT & UPN)
================================================================================
#>
function Show-Tool10-OUGroupFinder {
    if (-not (Assert-DomainJoined)) { return }

    $Form = New-Object System.Windows.Forms.Form
    $Form.Text = "Tool 10: Active Directory OU & Gruppen Finder (inkl. adminCount & UPN)"
    $Form.Size = New-Object System.Drawing.Size(1300, 750)
    $Form.StartPosition = "CenterScreen"
    $Form.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9)
    $Form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

    $script:CurrentObjectGroups = @()

    # --- OBERES STEUERUNGS-PANEL (LAYOUT & ABSTÄNDE) ---
    $pnlTop = New-Object System.Windows.Forms.Panel
    $pnlTop.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlTop.Height = 75
    $pnlTop.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
    $Form.Controls.Add($pnlTop)

    # 1. Objekttyp
    $LblType = New-Object System.Windows.Forms.Label
    $LblType.Text = "Objekttyp:"
    $LblType.Location = New-Object System.Drawing.Point(15, 15)
    $LblType.AutoSize = $true
    $LblType.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9, [System.Drawing.FontStyle]::Bold)
    $pnlTop.Controls.Add($LblType)

    $CmbType = New-Object System.Windows.Forms.ComboBox
    $CmbType.Location = New-Object System.Drawing.Point(95, 12)
    $CmbType.Size = New-Object System.Drawing.Size(175, 24)
    $CmbType.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$CmbType.Items.Add("Clients (ohne Server)")
    [void]$CmbType.Items.Add("Server (nur Server OS)")
    [void]$CmbType.Items.Add("User (Benutzerkonten)")
    $CmbType.SelectedIndex = 0
    $pnlTop.Controls.Add($CmbType)

    # 2. Suchbegriff
    $LblSearch = New-Object System.Windows.Forms.Label
    $LblSearch.Text = "Suchbegriff:"
    $LblSearch.Location = New-Object System.Drawing.Point(290, 15)
    $LblSearch.AutoSize = $true
    $LblSearch.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9, [System.Drawing.FontStyle]::Bold)
    $pnlTop.Controls.Add($LblSearch)

    $TxtSearch = New-Object System.Windows.Forms.TextBox
    $TxtSearch.Location = New-Object System.Drawing.Point(375, 12)
    $TxtSearch.Size = New-Object System.Drawing.Size(300, 24)
    $TxtSearch.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9)
    $TxtSearch.Text = "*"
    $pnlTop.Controls.Add($TxtSearch)

    # 3. Suchen-Button
    $BtnSearch = New-Object System.Windows.Forms.Button
    $BtnSearch.Text = "Suchen"
    $BtnSearch.Location = New-Object System.Drawing.Point(690, 10)
    $BtnSearch.Size = New-Object System.Drawing.Size(110, 28)
    $BtnSearch.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $BtnSearch.ForeColor = [System.Drawing.Color]::White
    $BtnSearch.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $BtnSearch.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9, [System.Drawing.FontStyle]::Bold)
    $pnlTop.Controls.Add($BtnSearch)

    # Statuszeile
    $LblStatus = New-Object System.Windows.Forms.Label
    $LblStatus.Text = "Bereit."
    $LblStatus.Location = New-Object System.Drawing.Point(15, 46)
    $LblStatus.AutoSize = $true
    $LblStatus.ForeColor = [System.Drawing.Color]::FromArgb(50, 70, 90)
    $pnlTop.Controls.Add($LblStatus)

    # --- SPLITCONTAINER (LINKS: DATAGRID, RECHTS: GRUPPEN) ---
    $splitContainer = New-Object System.Windows.Forms.SplitContainer
    $splitContainer.Dock = [System.Windows.Forms.DockStyle]::Fill
    $splitContainer.SplitterDistance = 840
    $splitContainer.Panel1.Padding = New-Object System.Windows.Forms.Padding(15, 5, 5, 15)
    $splitContainer.Panel2.Padding = New-Object System.Windows.Forms.Padding(5, 5, 15, 15)
    $Form.Controls.Add($splitContainer)
    $splitContainer.BringToFront()
    $pnlTop.SendToBack()

    # Tabelle links
    $DataGrid = New-Object System.Windows.Forms.DataGridView
    $DataGrid.Dock = [System.Windows.Forms.DockStyle]::Fill
    $DataGrid.ReadOnly = $true
    $DataGrid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $DataGrid.MultiSelect = $false
    $DataGrid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    Apply-StandardGridTheme -Grid $DataGrid
    $splitContainer.Panel1.Controls.Add($DataGrid)

    # Gruppen-Container rechts
    $pnlGroupHeader = New-Object System.Windows.Forms.Panel
    $pnlGroupHeader.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlGroupHeader.Height = 65
    $pnlGroupHeader.BackColor = [System.Drawing.Color]::Transparent
    $splitContainer.Panel2.Controls.Add($pnlGroupHeader)

    $LblGroups = New-Object System.Windows.Forms.Label
    $LblGroups.Text = "Gruppenmitgliedschaften:"
    $LblGroups.Location = New-Object System.Drawing.Point(0, 5)
    $LblGroups.AutoSize = $true
    $LblGroups.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9, [System.Drawing.FontStyle]::Bold)
    $pnlGroupHeader.Controls.Add($LblGroups)

    $TxtGroupFilter = New-Object System.Windows.Forms.TextBox
    $TxtGroupFilter.Location = New-Object System.Drawing.Point(0, 30)
    $TxtGroupFilter.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $pnlGroupHeader.Controls.Add($TxtGroupFilter)

    $PlaceholderText = "Gruppen filtern... (z.B. *Admin*)"
    $TxtGroupFilter.Text = $PlaceholderText
    $TxtGroupFilter.ForeColor = [System.Drawing.Color]::Gray

    $TxtGroupFilter.Add_GotFocus({
        if ($TxtGroupFilter.Text -eq $PlaceholderText) {
            $TxtGroupFilter.Text = ""
            $TxtGroupFilter.ForeColor = [System.Drawing.Color]::Black
        }
    })

    $TxtGroupFilter.Add_LostFocus({
        if ([string]::IsNullOrWhiteSpace($TxtGroupFilter.Text)) {
            $TxtGroupFilter.Text = $PlaceholderText
            $TxtGroupFilter.ForeColor = [System.Drawing.Color]::Gray
        }
    })

    $LstGroups = New-Object System.Windows.Forms.ListBox
    $LstGroups.Dock = [System.Windows.Forms.DockStyle]::Fill
    $LstGroups.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $LstGroups.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9)
    $splitContainer.Panel2.Controls.Add($LstGroups)
    $LstGroups.BringToFront()

    # Farb-Highlighting für adminCount = 1
    $DataGrid.Add_RowPrePaint({
        param($sender, $e)
        if ($e.RowIndex -ge 0 -and $e.RowIndex -lt $DataGrid.Rows.Count) {
            $row = $DataGrid.Rows[$e.RowIndex]
            $adminCountVal = $row.Cells["adminCount"].Value
            if ($adminCountVal -eq 1 -or $adminCountVal -eq "1") {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 243, 205)
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(140, 70, 0)
                $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(245, 225, 170)
                $row.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
            } else {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::White
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Black
            }
        }
    })

    # Filter-Routine für die rechte Gruppenliste
    $FilterGroupsList = {
        $LstGroups.Items.Clear()
        if (-not $script:CurrentObjectGroups -or $script:CurrentObjectGroups.Count -eq 0) { return }

        $Filter = $TxtGroupFilter.Text.Trim()

        if ($Filter -eq $PlaceholderText -or [string]::IsNullOrWhiteSpace($Filter)) {
            $Filtered = $script:CurrentObjectGroups
        } else {
            $Pattern = if ($Filter -notlike "*\**") { "*$Filter*" } else { $Filter }
            $Filtered = $script:CurrentObjectGroups | Where-Object { $_ -like $Pattern }
        }

        $LblGroups.Text = "Gruppen ($($Filtered.Count) von $($script:CurrentObjectGroups.Count)):"
        foreach ($Group in $Filtered) {
            [void]$LstGroups.Items.Add($Group)
        }
    }

    # Hauptsuche (Native LDAP)
    $PerformSearch = {
        $SearchTerm = $TxtSearch.Text.Trim()
        $LstGroups.Items.Clear()
        $script:CurrentObjectGroups = @()
        $LblGroups.Text = "Gruppenmitgliedschaften:"
        
        if ([string]::IsNullOrWhiteSpace($SearchTerm)) {
            $SearchTerm = "*"
        }

        $FilterPattern = if ($SearchTerm -eq "*" -or $SearchTerm -eq "") {
            "*"
        } elseif ($SearchTerm -notlike "*\**") {
            "*$SearchTerm*"
        } else {
            $SearchTerm
        }

        $LblStatus.Text = "Suche läuft..."
        $LblStatus.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
        $Form.Refresh()

        $ResultsList = [System.Collections.Generic.List[PSObject]]::new()

        try {
            switch ($CmbType.SelectedIndex) {
                0 { # Clients (ohne Server)
                    $ldapFilter = if ($FilterPattern -eq "*") {
                        "(&(objectCategory=computer)(!(operatingSystem=*Server*)))"
                    } else {
                        "(&(objectCategory=computer)(name=$FilterPattern)(!(operatingSystem=*Server*)))"
                    }
                    $raw = Search-NativeLdap -LdapFilter $ldapFilter -PropertiesToLoad @("name","distinguishedName","userAccountControl","operatingSystem","adminCount","userPrincipalName")
                    foreach ($Item in $raw) {
                        $p = $Item.Properties
                        $dn = [string]$p["distinguishedname"][0]
                        $dnComps = $dn -split '(?<!\\),'
                        $ou = ($dnComps[1..($dnComps.Length - 1)]) -join ','
                        $uac = if ($p["useraccountcontrol"].Count -gt 0) { [int]$p["useraccountcontrol"][0] } else { 0 }
                        $enabled = ($uac -band 2) -ne 2
                        $aCount = if ($p["admincount"].Count -gt 0 -and $p["admincount"][0] -eq 1) { 1 } else { 0 }
                        $upn = if ($p["userprincipalname"].Count -gt 0) { $p["userprincipalname"][0] } else { "" }
                        $os = if ($p["operatingsystem"].Count -gt 0) { $p["operatingsystem"][0] } else { "Unbekannt" }

                        $ResultsList.Add([PSCustomObject]@{
                            "Name"              = $p["name"][0]
                            "OU Pfad"           = $ou
                            "UPN"               = $upn
                            "adminCount"        = $aCount
                            "Aktiv"             = $enabled
                            "Betriebssystem"    = $os
                            "DistinguishedName" = $dn
                        })
                    }
                }
                1 { # Server
                    $ldapFilter = if ($FilterPattern -eq "*") {
                        "(&(objectCategory=computer)(operatingSystem=*Server*))"
                    } else {
                        "(&(objectCategory=computer)(name=$FilterPattern)(operatingSystem=*Server*))"
                    }
                    $raw = Search-NativeLdap -LdapFilter $ldapFilter -PropertiesToLoad @("name","distinguishedName","userAccountControl","operatingSystem","adminCount","userPrincipalName")
                    foreach ($Item in $raw) {
                        $p = $Item.Properties
                        $dn = [string]$p["distinguishedname"][0]
                        $dnComps = $dn -split '(?<!\\),'
                        $ou = ($dnComps[1..($dnComps.Length - 1)]) -join ','
                        $uac = if ($p["useraccountcontrol"].Count -gt 0) { [int]$p["useraccountcontrol"][0] } else { 0 }
                        $enabled = ($uac -band 2) -ne 2
                        $aCount = if ($p["admincount"].Count -gt 0 -and $p["admincount"][0] -eq 1) { 1 } else { 0 }
                        $upn = if ($p["userprincipalname"].Count -gt 0) { $p["userprincipalname"][0] } else { "" }
                        $os = if ($p["operatingsystem"].Count -gt 0) { $p["operatingsystem"][0] } else { "Unbekannt" }

                        $ResultsList.Add([PSCustomObject]@{
                            "Name"              = $p["name"][0]
                            "OU Pfad"           = $ou
                            "UPN"               = $upn
                            "adminCount"        = $aCount
                            "Aktiv"             = $enabled
                            "Betriebssystem"    = $os
                            "DistinguishedName" = $dn
                        })
                    }
                }
                2 { # User
                    $ldapFilter = if ($FilterPattern -eq "*") {
                        "(&(objectCategory=person)(objectClass=user))"
                    } else {
                        "(&(objectCategory=person)(objectClass=user)(|(sAMAccountName=$FilterPattern)(displayName=$FilterPattern)(userPrincipalName=$FilterPattern)))"
                    }
                    $raw = Search-NativeLdap -LdapFilter $ldapFilter -PropertiesToLoad @("sAMAccountName","displayName","distinguishedName","userAccountControl","adminCount","userPrincipalName","mail","department")
                    foreach ($Item in $raw) {
                        $p = $Item.Properties
                        $dn = [string]$p["distinguishedname"][0]
                        $dnComps = $dn -split '(?<!\\),'
                        $ou = ($dnComps[1..($dnComps.Length - 1)]) -join ','
                        $uac = if ($p["useraccountcontrol"].Count -gt 0) { [int]$p["useraccountcontrol"][0] } else { 0 }
                        $enabled = ($uac -band 2) -ne 2
                        $aCount = if ($p["admincount"].Count -gt 0 -and $p["admincount"][0] -eq 1) { 1 } else { 0 }
                        $upn = if ($p["userprincipalname"].Count -gt 0) { $p["userprincipalname"][0] } else { "" }
                        $sAM = if ($p["samaccountname"].Count -gt 0) { $p["samaccountname"][0] } else { "" }
                        $dName = if ($p["displayname"].Count -gt 0) { $p["displayname"][0] } else { "" }
                        $mail = if ($p["mail"].Count -gt 0) { $p["mail"][0] } else { "" }
                        $dept = if ($p["department"].Count -gt 0) { $p["department"][0] } else { "" }

                        $ResultsList.Add([PSCustomObject]@{
                            "Anmeldename"        = $sAM
                            "Vollständiger Name" = $dName
                            "UPN"                = $upn
                            "OU Pfad"            = $ou
                            "adminCount"         = $aCount
                            "E-Mail"             = $mail
                            "Abteilung"          = $dept
                            "Aktiv"              = $enabled
                            "DistinguishedName"  = $dn
                        })
                    }
                }
            }

            $DataGrid.DataSource = [System.Collections.ArrayList]::new($ResultsList)
            $LblStatus.Text = "Gefundene Objekte: $($ResultsList.Count)"
            $LblStatus.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 215)

            if ($DataGrid.Columns["DistinguishedName"]) { $DataGrid.Columns["DistinguishedName"].Visible = $false }
            if ($DataGrid.Columns["adminCount"])        { $DataGrid.Columns["adminCount"].FillWeight = 40 }
            if ($DataGrid.Columns["OU Pfad"])           { $DataGrid.Columns["OU Pfad"].FillWeight = 160 }

        } catch {
            $LblStatus.Text = "Fehler bei der Abfrage."
            $LblStatus.ForeColor = [System.Drawing.Color]::Red
            [System.Windows.Forms.MessageBox]::Show("Fehler bei AD-Abfrage: $($_.Exception.Message)", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }

    # Selection changed: Gruppen laden
    $DataGrid.Add_SelectionChanged({
        $LstGroups.Items.Clear()
        $script:CurrentObjectGroups = @()

        if ($DataGrid.SelectedRows.Count -gt 0) {
            $SelectedRow = $DataGrid.SelectedRows[0]
            $DN = [string]$SelectedRow.Cells["DistinguishedName"].Value

            if ($DN) {
                $LblGroups.Text = "Gruppen werden geladen..."
                $Form.Refresh()

                try {
                    $grpList = @()
                    $objEntry = [ADSI]"LDAP://$DN"
                    if ($objEntry.Properties["memberOf"].Count -gt 0) {
                        foreach ($mDN in $objEntry.Properties["memberOf"]) {
                            if ($mDN -match "^CN=([^,]+)") { $grpList += $Matches[1] } else { $grpList += $mDN }
                        }
                    }
                    $script:CurrentObjectGroups = ($grpList | Sort-Object)
                    & $FilterGroupsList
                }
                catch {
                    $LblGroups.Text = "Fehler beim Laden der Gruppen."
                    [void]$LstGroups.Items.Add("Fehler: $($_.Exception.Message)")
                }
            }
        }
    })

    # Event-Bindungen
    $TxtGroupFilter.Add_KeyUp({ & $FilterGroupsList })
    $BtnSearch.Add_Click($PerformSearch)
    $TxtSearch.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $_.SuppressKeyPress = $true
            & $PerformSearch
        }
    })

    $Form.Add_Shown({
        $BtnSearch.PerformClick()
    })

    [void]$Form.ShowDialog()
}
