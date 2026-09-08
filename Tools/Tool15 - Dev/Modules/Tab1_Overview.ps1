# =========================================================================
# Tab1_Overview.ps1 - GPO Uebersicht & LDAP / Snapshot-Engine
# =========================================================================

function Build-Tab1_Overview {
    param($tabControl, $domainDN, $domainName)

    if ($domainDN)   { $script:domainDN = $domainDN }
    if ($domainName) { $script:domainName = $domainName }

    $tabOverview = New-Object System.Windows.Forms.TabPage
    $tabOverview.Text = "1. GPO Uebersicht & Verlinkungs-Analyse"
    $tabOverview.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)

    $panelOverviewTop = New-Object System.Windows.Forms.Panel
    $panelOverviewTop.Dock = [System.Windows.Forms.DockStyle]::Top
    $panelOverviewTop.Height = 85
    $panelOverviewTop.BackColor = [System.Drawing.Color]::FromArgb(245, 248, 252)
    $panelOverviewTop.Padding = New-Object System.Windows.Forms.Padding(10)

    $lblViewFilter = New-Object System.Windows.Forms.Label
    $lblViewFilter.Text = "Ansicht:"
    $lblViewFilter.Location = New-Object System.Drawing.Point(12, 16)
    $lblViewFilter.AutoSize = $true
    $lblViewFilter.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    # Im Script-Scope registriert, damit der spätere Refresh-Block Zugriff hat
    $script:comboViewMode = New-Object System.Windows.Forms.ComboBox
    $script:comboViewMode.Location = New-Object System.Drawing.Point(75, 13)
    $script:comboViewMode.Size = New-Object System.Drawing.Size(200, 25)
    $script:comboViewMode.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$script:comboViewMode.Items.AddRange(@("Alle GPOs", "GPOs mit WMI-Filter", "Nicht verlinkte GPOs (Unlinked)"))
    $script:comboViewMode.SelectedIndex = 0

    $lblOverviewSearch = New-Object System.Windows.Forms.Label
    $lblOverviewSearch.Text = "Suche:"
    $lblOverviewSearch.Location = New-Object System.Drawing.Point(285, 16)
    $lblOverviewSearch.AutoSize = $true
    $lblOverviewSearch.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    # Im Script-Scope registriert, damit .Trim() nie auf $null läuft
    $script:txtOverviewSearch = New-Object System.Windows.Forms.TextBox
    $script:txtOverviewSearch.Location = New-Object System.Drawing.Point(335, 13)
    $script:txtOverviewSearch.Size = New-Object System.Drawing.Size(130, 25)

    $btnLoadAdGpos = New-Object System.Windows.Forms.Button
    $btnLoadAdGpos.Text = "GPOs einlesen"
    $btnLoadAdGpos.Location = New-Object System.Drawing.Point(475, 10)
    $btnLoadAdGpos.Size = New-Object System.Drawing.Size(120, 30)
    $btnLoadAdGpos.BackColor = [System.Drawing.Color]::FromArgb(225, 238, 255)
    $btnLoadAdGpos.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $btnSaveSnapshot = New-Object System.Windows.Forms.Button
    $btnSaveSnapshot.Text = "Snapshot speichern"
    $btnSaveSnapshot.Location = New-Object System.Drawing.Point(602, 10)
    $btnSaveSnapshot.Size = New-Object System.Drawing.Size(140, 30)
    $btnSaveSnapshot.BackColor = [System.Drawing.Color]::FromArgb(255, 245, 230)

    $btnLoadSnapshot = New-Object System.Windows.Forms.Button
    $btnLoadSnapshot.Text = "Snapshot laden"
    $btnLoadSnapshot.Location = New-Object System.Drawing.Point(748, 10)
    $btnLoadSnapshot.Size = New-Object System.Drawing.Size(125, 30)
    $btnLoadSnapshot.BackColor = [System.Drawing.Color]::FromArgb(235, 245, 255)

    $btnExportOverviewCsv = New-Object System.Windows.Forms.Button
    $btnExportOverviewCsv.Text = "CSV Export"
    $btnExportOverviewCsv.Location = New-Object System.Drawing.Point(880, 10)
    $btnExportOverviewCsv.Size = New-Object System.Drawing.Size(100, 30)
    $btnExportOverviewCsv.BackColor = [System.Drawing.Color]::FromArgb(230, 245, 230)

    $script:lblLegendOverview = New-Object System.Windows.Forms.Label
    $script:lblLegendOverview.Text = "Legende:  [Blau] Echte Default GPO  |  [Orange] DDP-Integritaetswarnung  |  [Gruen] OK (1 Seite aktiv)  |  [Rot] Nicht OK"
    $script:lblLegendOverview.Location = New-Object System.Drawing.Point(12, 52)
    $script:lblLegendOverview.AutoSize = $true
    $script:lblLegendOverview.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Italic)

    $panelOverviewTop.Controls.AddRange(@(
        $lblViewFilter, $script:comboViewMode, $lblOverviewSearch, $script:txtOverviewSearch, 
        $btnLoadAdGpos, $btnSaveSnapshot, $btnLoadSnapshot, $btnExportOverviewCsv, $script:lblLegendOverview
    ))

    $splitOverview = New-Object System.Windows.Forms.SplitContainer
    $splitOverview.Dock = [System.Windows.Forms.DockStyle]::Fill
    $splitOverview.SplitterDistance = 980
    $splitOverview.SplitterWidth = 6

    $panelOvLeft = New-Object System.Windows.Forms.Panel
    $panelOvLeft.Dock = [System.Windows.Forms.DockStyle]::Fill
    $panelOvLeft.Padding = New-Object System.Windows.Forms.Padding(10, 8, 4, 10)

    $script:gridOvMaster = New-Object System.Windows.Forms.DataGridView
    $script:gridOvMaster.Dock = [System.Windows.Forms.DockStyle]::Fill
    $script:gridOvMaster.ReadOnly = $true
    $script:gridOvMaster.AllowUserToAddRows = $false
    $script:gridOvMaster.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $script:gridOvMaster.MultiSelect = $false
    $script:gridOvMaster.RowHeadersVisible = $false
    $script:gridOvMaster.BackgroundColor = [System.Drawing.Color]::White
    $script:gridOvMaster.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $script:gridOvMaster.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(230, 236, 245)
    $script:gridOvMaster.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    $script:gridOvMaster.ColumnHeadersHeight = 34
    $script:gridOvMaster.RowTemplate.Height = 28
    $script:gridOvMaster.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(249, 251, 254)
    $script:gridOvMaster.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells

    $panelOvLeft.Controls.Add($script:gridOvMaster)
    $splitOverview.Panel1.Controls.Add($panelOvLeft)

    $panelOvRight = New-Object System.Windows.Forms.Panel
    $panelOvRight.Dock = [System.Windows.Forms.DockStyle]::Fill
    $panelOvRight.Padding = New-Object System.Windows.Forms.Padding(4, 8, 10, 10)

    $script:lblOvDetailsTitle = New-Object System.Windows.Forms.Label
    $script:lblOvDetailsTitle.Text = "Verlinkungsziele der GPO (OUs / Domaene):"
    $script:lblOvDetailsTitle.Dock = [System.Windows.Forms.DockStyle]::Top
    $script:lblOvDetailsTitle.Height = 46
    $script:lblOvDetailsTitle.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)

    $script:gridOvDetails = New-Object System.Windows.Forms.DataGridView
    $script:gridOvDetails.Dock = [System.Windows.Forms.DockStyle]::Fill
    $script:gridOvDetails.ReadOnly = $true
    $script:gridOvDetails.RowHeadersVisible = $false
    $script:gridOvDetails.BackgroundColor = [System.Drawing.Color]::White
    $script:gridOvDetails.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $script:gridOvDetails.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(230, 236, 245)
    $script:gridOvDetails.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    $script:gridOvDetails.ColumnHeadersHeight = 34
    $script:gridOvDetails.RowTemplate.Height = 28
    $script:gridOvDetails.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(249, 251, 254)
    $script:gridOvDetails.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill

    $panelOvRight.Controls.Add($script:gridOvDetails)
    $panelOvRight.Controls.Add($script:lblOvDetailsTitle)
    $splitOverview.Panel2.Controls.Add($panelOvRight)

    $tabOverview.Controls.Add($splitOverview)
    $tabOverview.Controls.Add($panelOverviewTop)
    $tabControl.TabPages.Add($tabOverview)

    Enable-GridSorting -Grid $script:gridOvMaster
    Enable-GridSorting -Grid $script:gridOvDetails

    # Rendering mit vollständiger Null-Absicherung
    $script:Update_OverviewDisplay = {
        if ($script:isClosing -or $null -eq $script:gridOvMaster -or $script:gridOvMaster.IsDisposed) { return }

        $mode = if ($script:comboViewMode -and -not $script:comboViewMode.IsDisposed -and $script:comboViewMode.SelectedItem) {
            $script:comboViewMode.SelectedItem
        } else {
            "Alle GPOs"
        }

        $filterText = if ($script:txtOverviewSearch -and -not $script:txtOverviewSearch.IsDisposed -and $script:txtOverviewSearch.Text) {
            "$($script:txtOverviewSearch.Text)".Trim()
        } else {
            ""
        }

        $filtered = $script:rawOverviewList | Where-Object {
            $item = $_
            $matchMode = switch ($mode) {
                "GPOs mit WMI-Filter"              { $item."WMI-Filter" -ne "-" }
                "Nicht verlinkte GPOs (Unlinked)" { $item."Verlinkt" -eq "Nein" }
                default                           { $true }
            }
            $matchSearch = if ([string]::IsNullOrWhiteSpace($filterText)) { $true } else {
                $item."GPO Name" -like "*$filterText*" -or $item."WMI-Filter" -like "*$filterText*" -or $item."Gesamt-Status" -like "*$filterText*"
            }
            $matchMode -and $matchSearch
        }

        $arr = [System.Collections.ArrayList]::new()
        foreach ($it in $filtered) { [void]$arr.Add($it) }
        $script:gridOvMaster.DataSource = $arr

        if ($script:gridOvMaster.Columns["GUID"]) { $script:gridOvMaster.Columns["GUID"].Visible = $false }
        if ($script:gridOvMaster.Columns["WMI Query"]) { $script:gridOvMaster.Columns["WMI Query"].Visible = $false }

        if ($script:lblLegendOverview -and -not $script:lblLegendOverview.IsDisposed) {
            $script:lblLegendOverview.Text = "Status: $($arr.Count) von $($script:rawOverviewList.Count) GPOs  |  [Blau] Echte Default GPO  |  [Orange] DDP-Integritaetswarnung  |  [Gruen] OK  |  [Rot] Nicht OK"
        }

        if ($script:gridOvMaster.Rows.Count -gt 0) {
            $script:gridOvMaster.Rows[0].Selected = $true
        }
    }

    $script:Invoke_LoadOverview = {
        if ($script:isClosing -or $form.IsDisposed) { return }

        $domainTargetDN = if ($script:domainDN) { $script:domainDN } else { ([ADSI]"LDAP://RootDSE").defaultNamingContext.Value }

        if ($script:pbarGlobal) {
            $script:pbarGlobal.Visible = $true
            $script:pbarGlobal.Minimum = 0
            $script:pbarGlobal.Value = 0
        }
        if ($script:lblProgressInfo) { $script:lblProgressInfo.Text = "Lade Active Directory Struktur..." }
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        [System.Windows.Forms.Application]::DoEvents()

        try {
            $script:gpoLinksCache.Clear()
            $script:rawOverviewList.Clear()

            # 1. WMI-Filter
            $wmiMap = @{}
            try {
                $wmiRoot = [System.DirectoryServices.DirectoryEntry]::new("LDAP://CN=SOM,CN=WMIPolicy,CN=System,$domainTargetDN")
                $wmiSearcher = [System.DirectoryServices.DirectorySearcher]::new($wmiRoot)
                $wmiSearcher.Filter = "(objectClass=msWMI-Som)"
                $wmiSearcher.PropertiesToLoad.AddRange(@("msWMI-Name", "msWMI-ID", "msWMI-Parm2"))
                $wmiResults = $wmiSearcher.FindAll()
                foreach ($w in $wmiResults) {
                    $wmiMap[$w.Properties["mswmi-id"][0]] = [PSCustomObject]@{
                        Name  = $w.Properties["mswmi-name"][0]
                        Query = if ($w.Properties["mswmi-parm2"]) { $w.Properties["mswmi-parm2"][0] } else { "" }
                    }
                }
            } catch {}

            # 2. OU & Domain Verlinkungen
            $linkRoot = [System.DirectoryServices.DirectoryEntry]::new("LDAP://$domainTargetDN")
            $linkSearcher = [System.DirectoryServices.DirectorySearcher]::new($linkRoot)
            $linkSearcher.Filter = "(|(objectClass=organizationalUnit)(objectClass=domainDNS))"
            $linkSearcher.PropertiesToLoad.AddRange(@("distinguishedName", "gPLink"))
            $linkSearcher.SearchScope = [System.DirectoryServices.SearchScope]::Subtree
            $ouResults = $linkSearcher.FindAll()

            foreach ($ou in $ouResults) {
                if ($ou.Properties["gplink"]) {
                    $targetDN = $ou.Properties["distinguishedname"][0]
                    $matches = [regex]::Matches($ou.Properties["gplink"][0], "cn=({?[a-fA-F0-9-]+}?)", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                    foreach ($m in $matches) {
                        $cleanGuid = $m.Groups[1].Value.Trim('{','}').ToUpper()
                        if (-not $script:gpoLinksCache.ContainsKey($cleanGuid)) {
                            $script:gpoLinksCache[$cleanGuid] = [System.Collections.Generic.List[string]]::new()
                        }
                        $script:gpoLinksCache[$cleanGuid].Add($targetDN)
                    }
                }
            }

            # 2b. Sites
            try {
                $configDN = ([ADSI]"LDAP://RootDSE").configurationNamingContext.Value
                $siteRoot = [System.DirectoryServices.DirectoryEntry]::new("LDAP://CN=Sites,$configDN")
                $siteSearcher = [System.DirectoryServices.DirectorySearcher]::new($siteRoot)
                $siteSearcher.Filter = "(objectClass=site)"
                $siteSearcher.PropertiesToLoad.AddRange(@("distinguishedName", "gPLink"))
                $siteResults = $siteSearcher.FindAll()
                foreach ($site in $siteResults) {
                    if ($site.Properties["gplink"]) {
                        $targetDN = $site.Properties["distinguishedname"][0]
                        $matches = [regex]::Matches($site.Properties["gplink"][0], "cn=({?[a-fA-F0-9-]+}?)", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                        foreach ($m in $matches) {
                            $cleanGuid = $m.Groups[1].Value.Trim('{','}').ToUpper()
                            if (-not $script:gpoLinksCache.ContainsKey($cleanGuid)) {
                                $script:gpoLinksCache[$cleanGuid] = [System.Collections.Generic.List[string]]::new()
                            }
                            $script:gpoLinksCache[$cleanGuid].Add($targetDN)
                        }
                    }
                }
            } catch {}

            # 3. GPOs auslesen
            $gpoRoot = [System.DirectoryServices.DirectoryEntry]::new("LDAP://CN=Policies,CN=System,$domainTargetDN")
            $gpoSearcher = [System.DirectoryServices.DirectorySearcher]::new($gpoRoot)
            $gpoSearcher.Filter = "(objectClass=groupPolicyContainer)"
            $gpoSearcher.PropertiesToLoad.AddRange(@("displayName", "name", "flags", "gPCWQLFilter", "whenCreated", "whenChanged"))
            $gpoResults = $gpoSearcher.FindAll()

            if ($script:pbarGlobal) { $script:pbarGlobal.Maximum = [Math]::Max(1, $gpoResults.Count) }
            $currentIndex = 0

            foreach ($g in $gpoResults) {
                $currentIndex++
                $rawGuid = if ($g.Properties["name"]) { "$($g.Properties['name'][0])" } else { "" }
                $cleanGuid = if ($rawGuid) { $rawGuid.Trim('{','}').ToUpper() } else { "" }
                $displayName = if ($g.Properties["displayname"]) { "$($g.Properties['displayname'][0])" } else { "{$cleanGuid}" }
                $flags = if ($g.Properties["flags"]) { [int]$g.Properties["flags"][0] } else { 0 }

                if ($script:pbarGlobal) { $script:pbarGlobal.Value = $currentIndex }
                if ($script:lblProgressInfo) { $script:lblProgressInfo.Text = "Lese GPO ($currentIndex / $($gpoResults.Count)): $displayName" }
                if ($currentIndex % 4 -eq 0) { [System.Windows.Forms.Application]::DoEvents() }

                $overallStatus = ""
                if ($cleanGuid -eq $script:StandardDdpGuid) {
                    $overallStatus = if ($displayName -eq "Default Domain Policy") { "Sonderstellung (Default GPO)" } else { "WARNUNG: Original DDP-GUID, aber umbenannt!" }
                } elseif ($displayName -like "*Default Domain Policy*") {
                    $overallStatus = "WARNUNG: Namensduplikat (Keine Standard-GUID!)"
                } elseif ($cleanGuid -eq $script:StandardDdcpGuid) {
                    $overallStatus = if ($displayName -eq "Default Domain Controllers Policy") { "Sonderstellung (Default GPO)" } else { "WARNUNG: Original DDCP-GUID, aber umbenannt!" }
                } elseif ($displayName -like "*Default Domain Controllers Policy*") {
                    $overallStatus = "WARNUNG: DDCP-Namensduplikat (Keine Standard-GUID!)"
                } else {
                    $overallStatus = switch ($flags) {
                        1 { "OK (Nur Computer)" }
                        2 { "OK (Nur Benutzer)" }
                        0 { "Nicht OK (Beide aktiviert)" }
                        default { "Nicht OK (Vollstaendig deaktiviert)" }
                    }
                }

                $isLinked = ($script:gpoLinksCache.ContainsKey($cleanGuid) -and $script:gpoLinksCache[$cleanGuid].Count -gt 0)
                $linkedCount = if ($isLinked) { $script:gpoLinksCache[$cleanGuid].Count } else { 0 }

                $wmiFilterName = "-"
                $wmiFilterQuery = "-"
                if ($g.Properties["gpcwqlfilter"] -and "$($g.Properties['gpcwqlfilter'][0])" -match "({?[a-fA-F0-9-]+}?)") {
                    $wGuid = $matches[1]
                    if ($wmiMap.ContainsKey($wGuid)) {
                        $wmiFilterName = $wmiMap[$wGuid].Name
                        $wmiFilterQuery = $wmiMap[$wGuid].Query
                    } else { $wmiFilterName = $wGuid }
                }

                $script:rawOverviewList.Add([PSCustomObject]@{
                    "GPO Name"      = $displayName
                    "Gesamt-Status" = $overallStatus
                    "Verlinkt"      = if ($isLinked) { "Ja" } else { "Nein" }
                    "Link-Anzahl"   = $linkedCount
                    "Benutzer"      = if (($flags -band 1) -eq 1) { "Deaktiviert" } else { "Aktiviert" }
                    "Computer"      = if (($flags -band 2) -eq 2) { "Deaktiviert" } else { "Aktiviert" }
                    "WMI-Filter"    = $wmiFilterName
                    "WMI Query"     = $wmiFilterQuery
                    "GUID"          = "{$cleanGuid}"
                    "Erstellt am"   = if ($g.Properties["whencreated"]) { (Get-Date $g.Properties["whencreated"][0]).ToString("dd.MM.yyyy HH:mm") } else { "-" }
                    "Geaendert am"  = if ($g.Properties["whenchanged"]) { (Get-Date $g.Properties["whenchanged"][0]).ToString("dd.MM.yyyy HH:mm") } else { "-" }
                })
            }

            & $script:Update_OverviewDisplay
            if ($script:lblProgressInfo) { $script:lblProgressInfo.Text = "GPO-Einlesen abgeschlossen: $($script:rawOverviewList.Count) Richtlinien geladen." }
        } catch {
            if ($script:lblProgressInfo) { $script:lblProgressInfo.Text = "Fehler: $($_.Exception.Message)" }
        } finally {
            # Offene LDAP-Verbindungen freigeben
            if ($wmiSearcher)  { $wmiSearcher.Dispose() }
            if ($wmiRoot)      { $wmiRoot.Dispose() }
            if ($linkSearcher) { $linkSearcher.Dispose() }
            if ($linkRoot)     { $linkRoot.Dispose() }
            if ($gpoSearcher)  { $gpoSearcher.Dispose() }
            if ($gpoRoot)      { $gpoRoot.Dispose() }

            if ($script:pbarGlobal) { $script:pbarGlobal.Visible = $false }
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            
            # Speicher freigeben
            [System.GC]::Collect()
        }    }

    $script:gridOvMaster.Add_DataBindingComplete({
        if ($script:isClosing -or $null -eq $script:gridOvMaster -or $script:gridOvMaster.IsDisposed) { return }
        foreach ($row in $script:gridOvMaster.Rows) {
            $status = [string]$row.Cells["Gesamt-Status"].Value
            if ($status -match "^WARNUNG") {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 238, 204)
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(180, 50, 0)
                $row.DefaultCellStyle.Font = New-Object System.Drawing.Font($script:gridOvMaster.Font, [System.Drawing.FontStyle]::Bold)
            } elseif ($status -match "Sonderstellung") {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(232, 238, 255)
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(0, 45, 135)
                $row.DefaultCellStyle.Font = New-Object System.Drawing.Font($script:gridOvMaster.Font, [System.Drawing.FontStyle]::Bold)
            } elseif ($status -match "^Nicht OK") {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::MistyRose
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkRed
            } else {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(235, 247, 235)
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkGreen
            }
        }
    })

    $btnLoadAdGpos.Add_Click({
        & $script:Invoke_LoadOverview
        if ($script:Invoke_UpdateDashboard) { & $script:Invoke_UpdateDashboard }
    })

    $script:comboViewMode.Add_SelectedIndexChanged({ & $script:Update_OverviewDisplay })
    $script:txtOverviewSearch.Add_TextChanged({ & $script:Update_OverviewDisplay })

    $script:gridOvMaster.Add_SelectionChanged({
        if ($script:isClosing -or $null -eq $script:gridOvMaster -or $script:gridOvMaster.IsDisposed -or $null -eq $script:gridOvDetails -or $script:gridOvDetails.IsDisposed) { return }
        if ($script:gridOvMaster.SelectedRows.Count -gt 0) {
            $rawGuid = [string]$script:gridOvMaster.SelectedRows[0].Cells["GUID"].Value
            $cleanGuid = if ($rawGuid) { $rawGuid.Trim('{','}').ToUpper() } else { "" }
            $gName = [string]$script:gridOvMaster.SelectedRows[0].Cells["GPO Name"].Value
            if ($script:lblOvDetailsTitle) { $script:lblOvDetailsTitle.Text = "Verlinkungsziele fuer:`r`n[$gName]" }

            $arrDetails = [System.Collections.ArrayList]::new()
            if ($cleanGuid -and $script:gpoLinksCache.ContainsKey($cleanGuid) -and $script:gpoLinksCache[$cleanGuid].Count -gt 0) {
                foreach ($dn in $script:gpoLinksCache[$cleanGuid]) {
                    [void]$arrDetails.Add([PSCustomObject]@{
                        "Typ"               = if ($dn -match "^OU=") { "OU" } elseif ($dn -match "^DC=") { "Domaenen-Root" } else { "Site" }
                        "Name / Ziel"       = if ($dn -match "^[A-Z]+=([^,]+)") { $matches[1] } else { $dn }
                        "DistinguishedName" = $dn
                    })
                }
            } else {
                [void]$arrDetails.Add([PSCustomObject]@{
                    "Typ"               = "Info"
                    "Name / Ziel"       = "-- Keine Verknuepfung --"
                    "DistinguishedName" = "[Hinweis] Diese GPO ist nicht verlinkt (Unlinked)."
                })
            }
            $script:gridOvDetails.DataSource = $arrDetails
        }
    })

    $btnSaveSnapshot.Add_Click({
        if ($script:rawOverviewList.Count -eq 0) { return }
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "GPO Snapshot (*.gposnap)|*.gposnap"
        $sfd.FileName = "GPO_Snapshot_$($script:domainName)_$((Get-Date).ToString('yyyyMMdd_HHmm')).gposnap"
        if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            @{
                Version       = $script:ToolVersion
                Timestamp     = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                Domain        = $script:domainName
                OverviewList  = @($script:rawOverviewList)
                GpoLinksCache = $script:gpoLinksCache
            } | Export-Clixml -Path $sfd.FileName -Depth 4
            [System.Windows.Forms.MessageBox]::Show("Snapshot gespeichert!", "Erfolg", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    })

    $btnLoadSnapshot.Add_Click({
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Filter = "GPO Snapshot (*.gposnap)|*.gposnap"
        if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $data = Import-Clixml -Path $ofd.FileName
            $script:rawOverviewList.Clear()
            foreach ($it in $data.OverviewList) { [void]$script:rawOverviewList.Add($it) }
            $script:gpoLinksCache.Clear()
            foreach ($k in $data.GpoLinksCache.Keys) { $script:gpoLinksCache[$k] = $data.GpoLinksCache[$k] }

            & $script:Update_OverviewDisplay
            if ($script:Invoke_UpdateDashboard) { & $script:Invoke_UpdateDashboard }
            [System.Windows.Forms.MessageBox]::Show("Snapshot geladen: $($script:rawOverviewList.Count) GPOs", "Aktiv", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    })

    $btnExportOverviewCsv.Add_Click({
        if ($null -eq $script:gridOvMaster -or $script:gridOvMaster.Rows.Count -eq 0) { return }
        $targetBase = if ($script:txtBackupTargetDir) { $script:txtBackupTargetDir.Text.Trim() } else { "C:\Install\Backup\GPO" }
        if (-not (Test-Path $targetBase)) { New-Item -ItemType Directory -Path $targetBase -Force | Out-Null }
        $csvFile = Join-Path $targetBase "GPO_Overview_$((Get-Date).ToString('yyyyMMdd_HHmm')).csv"
        @($script:gridOvMaster.DataSource) | Export-Csv -Path $csvFile -Delimiter ";" -NoTypeInformation -Encoding UTF8
        [System.Windows.Forms.MessageBox]::Show("Exportiert nach: $csvFile", "Export abgeschlossen", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    })
}