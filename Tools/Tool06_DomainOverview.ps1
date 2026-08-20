<#
================================================================================
 TOOL 6: DOMAIN OVERVIEW & PRIVILEGED ADMIN AUDIT (3 TABS)
================================================================================
#>
function Open-ToolDomainOverview {
    if (-not (Assert-DomainJoined)) { return }

    $subForm = New-Object System.Windows.Forms.Form
    $subForm.Text = "Tool 6 - Active Directory Domänen-Übersicht & Admin-Audit"
    $subForm.Size = New-Object System.Drawing.Size(1200, 780)
    $subForm.StartPosition = "CenterParent"
    $subForm.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9)
    $subForm.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Dock = [System.Windows.Forms.DockStyle]::Fill
    $tabControl.Padding = New-Object System.Drawing.Point(12, 6)

    # --- TAB 1: DOMÄNENCONTROLLER (DCs) ---
    $tabDCs = New-Object System.Windows.Forms.TabPage
    $tabDCs.Text = if ($script:CurrentLang -eq "DE") { "  🖥️ Domänencontroller (DCs)  " } else { "  🖥️ Domain Controllers (DCs)  " }
    $tabDCs.BackColor = [System.Drawing.Color]::White

    $pnlTopDC = New-Object System.Windows.Forms.Panel
    $pnlTopDC.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlTopDC.Height = 44
    $pnlTopDC.BackColor = [System.Drawing.Color]::FromArgb(248, 249, 250)

    $lblDCCount = New-Object System.Windows.Forms.Label
    $lblDCCount.Text = if ($script:CurrentLang -eq "DE") { "Erkannte Domain Controller in der Gesamtstruktur:" } else { "Discovered Domain Controllers in Forest:" }
    $lblDCCount.Location = New-Object System.Drawing.Point(12, 13)
    $lblDCCount.AutoSize = $true
    $lblDCCount.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9.5, [System.Drawing.FontStyle]::Bold)

    $btnRefreshDC = New-Object System.Windows.Forms.Button
    $btnRefreshDC.Text = if ($script:CurrentLang -eq "DE") { "Aktualisieren" } else { "Refresh" }
    $btnRefreshDC.Location = New-Object System.Drawing.Point(420, 8)
    $btnRefreshDC.Size = New-Object System.Drawing.Size(120, 28)
    $btnRefreshDC.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $btnRefreshDC.ForeColor = [System.Drawing.Color]::White
    $btnRefreshDC.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat

    $pnlTopDC.Controls.AddRange(@($lblDCCount, $btnRefreshDC))

    $gridDCs = New-Object System.Windows.Forms.DataGridView
    $gridDCs.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridDCs.ReadOnly = $true
    $gridDCs.AllowUserToAddRows = $false
    $gridDCs.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    Apply-StandardGridTheme -Grid $gridDCs -EnableAlternatingRowColor

    $tabDCs.Controls.Add($gridDCs)
    $tabDCs.Controls.Add($pnlTopDC)
    $pnlTopDC.SendToBack()
    $gridDCs.BringToFront()

    # --- TAB 2: FSMO ROLLEN & DOMÄNENSTATUS ---
    $tabFSMO = New-Object System.Windows.Forms.TabPage
    $tabFSMO.Text = if ($script:CurrentLang -eq "DE") { "  ⚙️ FSMO-Rollen & Domänenstatus  " } else { "  ⚙️ FSMO Roles & Domain Status  " }
    $tabFSMO.BackColor = [System.Drawing.Color]::White

    $pnlTopFSMO = New-Object System.Windows.Forms.Panel
    $pnlTopFSMO.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlTopFSMO.Height = 44
    $pnlTopFSMO.BackColor = [System.Drawing.Color]::FromArgb(248, 249, 250)

    $lblFSMOHead = New-Object System.Windows.Forms.Label
    $lblFSMOHead.Text = if ($script:CurrentLang -eq "DE") { "Zentrale FSMO Rolleninhaber & Gesamtstruktur-Ebenen:" } else { "FSMO Role Holders & Forest Levels:" }
    $lblFSMOHead.Location = New-Object System.Drawing.Point(12, 13)
    $lblFSMOHead.AutoSize = $true
    $lblFSMOHead.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9.5, [System.Drawing.FontStyle]::Bold)
    $pnlTopFSMO.Controls.Add($lblFSMOHead)

    $gridFSMO = New-Object System.Windows.Forms.DataGridView
    $gridFSMO.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridFSMO.ReadOnly = $true
    $gridFSMO.AllowUserToAddRows = $false
    $gridFSMO.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    Apply-StandardGridTheme -Grid $gridFSMO -EnableAlternatingRowColor

    $tabFSMO.Controls.Add($gridFSMO)
    $tabFSMO.Controls.Add($pnlTopFSMO)
    $pnlTopFSMO.SendToBack()
    $gridFSMO.BringToFront()

    # --- TAB 3: PRIVILEGIERTE ADMINS & DELEGIERUNG ---
    $tabAdmins = New-Object System.Windows.Forms.TabPage
    $tabAdmins.Text = if ($script:CurrentLang -eq "DE") { "  🛡️ Privilegierte Konten & Delegierung  " } else { "  🛡️ Privileged Accounts & Delegation  " }
    $tabAdmins.BackColor = [System.Drawing.Color]::White

    $pnlAdminTop = New-Object System.Windows.Forms.Panel
    $pnlAdminTop.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlAdminTop.Height = 48
    $pnlAdminTop.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

    $lblAdminFilter = New-Object System.Windows.Forms.Label
    $lblAdminFilter.Text = if ($script:CurrentLang -eq "DE") { "Live-Filter:" } else { "Live Filter:" }
    $lblAdminFilter.Location = New-Object System.Drawing.Point(12, 14)
    $lblAdminFilter.AutoSize = $true
    $lblAdminFilter.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9.0, [System.Drawing.FontStyle]::Bold)

    $txtAdminFilter = New-Object System.Windows.Forms.TextBox
    $txtAdminFilter.Location = New-Object System.Drawing.Point(95, 11)
    $txtAdminFilter.Size = New-Object System.Drawing.Size(260, 24)

    $lblAdminCount = New-Object System.Windows.Forms.Label
    $lblAdminCount.Location = New-Object System.Drawing.Point(375, 14)
    $lblAdminCount.AutoSize = $true
    $lblAdminCount.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9, [System.Drawing.FontStyle]::Regular)
    $lblAdminCount.ForeColor = [System.Drawing.Color]::FromArgb(50, 70, 90)

    $pnlAdminTop.Controls.AddRange(@($lblAdminFilter, $txtAdminFilter, $lblAdminCount))

    $pnlAdminLegend = New-Object System.Windows.Forms.Panel
    $pnlAdminLegend.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $pnlAdminLegend.Height = 32
    $pnlAdminLegend.BackColor = [System.Drawing.Color]::FromArgb(246, 248, 250)

    $lblAdminLegend = New-Object System.Windows.Forms.Label
    $lblAdminLegend.Dock = [System.Windows.Forms.DockStyle]::Fill
    $lblAdminLegend.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $lblAdminLegend.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 8.5)
    $lblAdminLegend.Text = " 🟥 Rot: Schema-Admins (Sollte im Normalbetrieb leer sein!)  |  ⚠️ Gelb: Delegierung erlaubt (Konto nicht vor Kerberos-Delegierung geschützt)"
    $pnlAdminLegend.Controls.Add($lblAdminLegend)

    $gridAdmins = New-Object System.Windows.Forms.DataGridView
    $gridAdmins.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridAdmins.ReadOnly = $true
    $gridAdmins.AllowUserToAddRows = $false
    $gridAdmins.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    Apply-StandardGridTheme -Grid $gridAdmins -EnableAlternatingRowColor

    $tabAdmins.Controls.Add($gridAdmins)
    $tabAdmins.Controls.Add($pnlAdminTop)
    $tabAdmins.Controls.Add($pnlAdminLegend)
    $pnlAdminTop.SendToBack()
    $pnlAdminLegend.SendToBack()
    $gridAdmins.BringToFront()

    $tabControl.TabPages.AddRange(@($tabDCs, $tabFSMO, $tabAdmins))
    $subForm.Controls.Add($tabControl)

    $script:adminList = [System.Collections.Generic.List[PSCustomObject]]::new()
    $script:dcList    = [System.Collections.Generic.List[PSCustomObject]]::new()
    $script:fsmoList  = [System.Collections.Generic.List[PSCustomObject]]::new()
    $script:SchemaAdminSid = ""

    $gridAdmins.Add_RowPrePaint({
        param($sender, $e)
        if ($e.RowIndex -ge 0 -and $e.RowIndex -lt $gridAdmins.Rows.Count) {
            $row  = $gridAdmins.Rows[$e.RowIndex]
            $grp  = [string]$row.Cells["Admin-Gruppe"].Value
            $sid  = [string]$row.Cells["Gruppen-SID"].Value
            $del  = [string]$row.Cells["Delegierungsschutz"].Value
            $stat = [string]$row.Cells["Status"].Value

            if (($script:SchemaAdminSid -and $sid -eq $script:SchemaAdminSid) -or ($grp -match "Schema[- ]?Admins")) {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 224, 224)
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(180, 0, 0)
                $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(245, 190, 190)
                $row.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
            }
            elseif ($del -like "*Ungeschützt*") {
                $row.Cells["Delegierungsschutz"].Style.BackColor = [System.Drawing.Color]::FromArgb(255, 243, 205)
                $row.Cells["Delegierungsschutz"].Style.ForeColor = [System.Drawing.Color]::FromArgb(160, 80, 0)
                $row.Cells["Delegierungsschutz"].Style.SelectionBackColor = [System.Drawing.Color]::FromArgb(245, 225, 170)
                $row.Cells["Delegierungsschutz"].Style.SelectionForeColor = [System.Drawing.Color]::Black
            }

            if ($stat -eq "Deaktiviert") {
                $row.Cells["Status"].Style.ForeColor = [System.Drawing.Color]::Gray
            }
        }
    })

    $loadDomainData = {
        try {
            $domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
            $forest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()

            # 1. DC Liste
            $script:dcList.Clear()
            foreach ($dc in $domain.DomainControllers) {
                $os = "Windows Server"
                try {
                    $de = $dc.GetDirectoryEntry()
                    if ($de.operatingSystem) { $os = "$($de.operatingSystem) $($de.operatingSystemServicePack)" }
                } catch {}

                $roles = @()
                if ($dc.Name -eq $domain.PdcRoleOwner.Name) { $roles += "PDC" }
                if ($dc.Name -eq $domain.RidRoleOwner.Name) { $roles += "RID" }
                if ($dc.Name -eq $domain.InfrastructureRoleOwner.Name) { $roles += "Infra" }
                if ($dc.Name -eq $forest.SchemaRoleOwner.Name) { $roles += "Schema" }
                if ($dc.Name -eq $forest.NamingRoleOwner.Name) { $roles += "Naming" }
                if ($roles.Count -eq 0) { $roles += "Standard DC" }

                $script:dcList.Add([PSCustomObject]@{
                    "DC Hostname"     = $dc.Name
                    "AD Site"         = $dc.SiteName
                    "IP-Adresse"      = $dc.IPAddress
                    "Global Catalog"  = if ($dc.IsGlobalCatalog()) { "Ja / Active" } else { "Nein" }
                    "Inhaber Rollen"  = ($roles -join ", ")
                    "Betriebssystem"  = $os
                })
            }
            $gridDCs.DataSource = [System.Collections.ArrayList]::new($script:dcList)
            $lblDCCount.Text = if ($script:CurrentLang -eq "DE") { "Gefundene Domain Controller ($($script:dcList.Count)):" } else { "Discovered Domain Controllers ($($script:dcList.Count)):" }

            # 2. FSMO & Domänenstatus
            $recycleBin = "Deaktiviert / Nicht konfiguriert"
            try {
                $rootDSE = [ADSI]"LDAP://RootDSE"
                $partDN = "CN=Partitions," + $rootDSE.configurationNamingContext.ToString()
                $partEntry = [ADSI]"LDAP://$partDN"
                if ($partEntry.Properties["msDS-EnabledFeature"].Count -gt 0) {
                    $recycleBin = "Aktiviert (Active Directory Recycle Bin ON)"
                }
            } catch {}

            $script:fsmoList.Clear()
            $script:fsmoList.Add([PSCustomObject]@{"Parameter / FSMO Rolle" = "PDC Emulator"; "Inhaber / Status" = $domain.PdcRoleOwner.Name; "Geltungsbereich" = "Domäne" })
            $script:fsmoList.Add([PSCustomObject]@{"Parameter / FSMO Rolle" = "RID Master"; "Inhaber / Status" = $domain.RidRoleOwner.Name; "Geltungsbereich" = "Domäne" })
            $script:fsmoList.Add([PSCustomObject]@{"Parameter / FSMO Rolle" = "Infrastructure Master"; "Inhaber / Status" = $domain.InfrastructureRoleOwner.Name; "Geltungsbereich" = "Domäne" })
            $script:fsmoList.Add([PSCustomObject]@{"Parameter / FSMO Rolle" = "Schema Master"; "Inhaber / Status" = $forest.SchemaRoleOwner.Name; "Geltungsbereich" = "Gesamtstruktur" })
            $script:fsmoList.Add([PSCustomObject]@{"Parameter / FSMO Rolle" = "Domain Naming Master"; "Inhaber / Status" = $forest.NamingRoleOwner.Name; "Geltungsbereich" = "Gesamtstruktur" })
            $script:fsmoList.Add([PSCustomObject]@{"Parameter / FSMO Rolle" = "Domänen-Funktionsebene"; "Inhaber / Status" = $domain.DomainMode.ToString(); "Geltungsbereich" = "Domäne" })
            $script:fsmoList.Add([PSCustomObject]@{"Parameter / FSMO Rolle" = "Gesamtstruktur-Funktionsebene"; "Inhaber / Status" = $forest.ForestMode.ToString(); "Geltungsbereich" = "Gesamtstruktur" })
            $script:fsmoList.Add([PSCustomObject]@{"Parameter / FSMO Rolle" = "AD Papierkorb (Recycle Bin)"; "Inhaber / Status" = $recycleBin; "Geltungsbereich" = "Gesamtstruktur" })

            $gridFSMO.DataSource = [System.Collections.ArrayList]::new($script:fsmoList)

            # 3. Privilegierte Admin-Gruppen & Delegierung
            $script:adminList.Clear()
            $domDN = Get-DomainDN
            $domainEntry = [ADSI]"LDAP://$domDN"
            $domainSID = (New-Object System.Security.Principal.SecurityIdentifier($domainEntry.Properties["objectSid"][0], 0)).Value
            $script:SchemaAdminSid = "$domainSID-518"

            $groupsToQuery = @(
                @{ Name = "Domänen-Admins";       SID = "$domainSID-512" },
                @{ Name = "Organisations-Admins"; SID = "$domainSID-519" },
                @{ Name = "Schema-Admins";        SID = $script:SchemaAdminSid },
                @{ Name = "Administratoren (AD)"; SID = "S-1-5-32-544" },
                @{ Name = "Konten-Operatoren";    SID = "$domainSID-517" },
                @{ Name = "Server-Operatoren";    SID = "$domainSID-549" },
                @{ Name = "Sicherungs-Operatoren";SID = "$domainSID-551" },
                @{ Name = "GPO-Ersteller";        SID = "$domainSID-520" }
            )

            foreach ($g in $groupsToQuery) {
                $gSearcher = New-Object DirectoryServices.DirectorySearcher($domainEntry)
                $gSearcher.Filter = "(&(objectCategory=group)(objectsid=$($g.SID)))"
                $groupObj = $gSearcher.FindOne()
                if ($groupObj) {
                    $groupName = if ($groupObj.Properties["samaccountname"].Count -gt 0) { $groupObj.Properties["samaccountname"][0] } else { $g.Name }
                    
                    foreach ($mDN in $groupObj.Properties["member"]) {
                        try {
                            $mSearcher = New-Object DirectoryServices.DirectorySearcher($domainEntry)
                            $mSearcher.Filter = "(distinguishedName=$mDN)"
                            $mObj = $mSearcher.FindOne()
                            
                            if ($mObj) {
                                $sAM = [string]$mObj.Properties["samaccountname"][0]
                                $uac = if ($mObj.Properties["useraccountcontrol"].Count -gt 0) { [int]$mObj.Properties["useraccountcontrol"][0] } else { 0 }
                                $enabled = if (($uac -band 2) -eq 2) { "Deaktiviert" } else { "Aktiv" }

                                $isNotDelegated = ($uac -band 0x100000) -eq 0x100000
                                $delegationStatus = if ($isNotDelegated) { 
                                    "Geschützt (Keine Delegierung)" 
                                } else { 
                                    "⚠️ Ungeschützt (Delegierung erlaubt)" 
                                }

                                $script:adminList.Add([PSCustomObject]@{
                                    "Admin-Gruppe"        = $groupName
                                    "Gruppen-SID"         = $g.SID
                                    "Account Name"        = $sAM
                                    "Status"              = $enabled
                                    "Delegierungsschutz"  = $delegationStatus
                                    "DistinguishedName"   = $mDN
                                })
                            }
                        } catch {}
                    }
                }
            }

            $gridAdmins.DataSource = [System.Collections.ArrayList]::new($script:adminList)
            $lblAdminCount.Text = "Gefundene privilegierte Konten: $($script:adminList.Count)"
            if ($gridAdmins.Columns["Gruppen-SID"]) { $gridAdmins.Columns["Gruppen-SID"].Visible = $false }
            if ($gridAdmins.Columns["DistinguishedName"]) { $gridAdmins.Columns["DistinguishedName"].FillWeight = 140 }

        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim LDAP-Abruf: $($_.Exception.Message)", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }

    $btnRefreshDC.Add_Click({ & $loadDomainData })

    $txtAdminFilter.Add_TextChanged({
        $filterText = $txtAdminFilter.Text.Trim().ToLower()
        if ([string]::IsNullOrWhiteSpace($filterText)) {
            $gridAdmins.DataSource = [System.Collections.ArrayList]::new($script:adminList)
            $lblAdminCount.Text = "Gefundene privilegierte Konten: $($script:adminList.Count)"
        } else {
            $filtered = $script:adminList | Where-Object {
                ($_."Admin-Gruppe" -and $_."Admin-Gruppe".ToLower().Contains($filterText)) -or
                ($_."Account Name" -and $_."Account Name".ToLower().Contains($filterText)) -or
                ($_."Delegierungsschutz" -and $_."Delegierungsschutz".ToLower().Contains($filterText)) -or
                ($_."DistinguishedName" -and $_."DistinguishedName".ToLower().Contains($filterText))
            }
            $arrF = [System.Collections.ArrayList]::new()
            foreach ($item in $filtered) { [void]$arrF.Add($item) }
            $gridAdmins.DataSource = $arrF
            $lblAdminCount.Text = "Gefiltert: $($arrF.Count) von $($script:adminList.Count)"
        }
    })

    $subForm.Add_Shown({ & $loadDomainData })
    [void]$subForm.ShowDialog()
}
