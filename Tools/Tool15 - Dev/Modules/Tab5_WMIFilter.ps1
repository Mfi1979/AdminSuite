# =========================================================================
# Tab5_WmiFilter.ps1 - WMI Filter Analyse & Zuordnungs-Inspektor
# =========================================================================

# Parser-Funktion auf Script-Ebene (fuer alle Scriptbloecke sichtbar)
function Parse-WmiParm2 ([string]$parm2) {
    if ([string]::IsNullOrWhiteSpace($parm2)) {
        return @{ Namespace = "root\cimv2"; Query = "-- Keine Query hinterlegt --" }
    }

    # Format: 1;1;<len>;WQL;root\cimv2;SELECT * FROM ...
    if ($parm2 -match '(?i)WQL;(?<ns>[^;]+);(?<q>SELECT[\s\S]+)') {
        return @{
            Namespace = $matches["ns"].Trim()
            Query     = $matches["q"].Trim()
        }
    }

    # Direkte WQL Query
    if ($parm2 -match '(?i)SELECT\s+') {
        return @{
            Namespace = "root\cimv2"
            Query     = $parm2.Trim()
        }
    }

    return @{
        Namespace = "root\cimv2"
        Query     = $parm2.Trim()
    }
}

function Build-Tab5_WmiFilter {
    param($tabControl, $domainDN, $domainName)

    $tabWmi = New-Object System.Windows.Forms.TabPage
    $tabWmi.Text = "5. WMI Filter Analyse"
    $tabWmi.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
    $tabWmi.BackColor = [System.Drawing.Color]::FromArgb(246, 248, 252)

    # ---------------------------------------------------------------------
    # Top Panel
    # ---------------------------------------------------------------------
    $panelTop = New-Object System.Windows.Forms.Panel
    $panelTop.Dock = [System.Windows.Forms.DockStyle]::Top
    $panelTop.Height = 85
    $panelTop.BackColor = [System.Drawing.Color]::FromArgb(242, 245, 250)
    $panelTop.Padding = New-Object System.Windows.Forms.Padding(10)

    $lblFilterMode = New-Object System.Windows.Forms.Label
    $lblFilterMode.Text = "Ansicht:"
    $lblFilterMode.Location = New-Object System.Drawing.Point(12, 16)
    $lblFilterMode.AutoSize = $true
    $lblFilterMode.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $script:comboWmiViewMode = New-Object System.Windows.Forms.ComboBox
    $script:comboWmiViewMode.Location = New-Object System.Drawing.Point(75, 13)
    $script:comboWmiViewMode.Size = New-Object System.Drawing.Size(200, 25)
    $script:comboWmiViewMode.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$script:comboWmiViewMode.Items.AddRange(@("Alle WMI-Filter", "Nur genutzte Filter", "Ungenutzte Filter (Verwaist)"))
    $script:comboWmiViewMode.SelectedIndex = 0

    $lblSearch = New-Object System.Windows.Forms.Label
    $lblSearch.Text = "Suche:"
    $lblSearch.Location = New-Object System.Drawing.Point(285, 16)
    $lblSearch.AutoSize = $true
    $lblSearch.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $script:txtWmiSearch = New-Object System.Windows.Forms.TextBox
    $script:txtWmiSearch.Location = New-Object System.Drawing.Point(335, 13)
    $script:txtWmiSearch.Size = New-Object System.Drawing.Size(160, 25)

    $btnLoadWmi = New-Object System.Windows.Forms.Button
    $btnLoadWmi.Text = "WMI-Filter laden"
    $btnLoadWmi.Location = New-Object System.Drawing.Point(505, 10)
    $btnLoadWmi.Size = New-Object System.Drawing.Size(130, 30)
    $btnLoadWmi.BackColor = [System.Drawing.Color]::FromArgb(225, 238, 255)
    $btnLoadWmi.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $btnExportWmiCsv = New-Object System.Windows.Forms.Button
    $btnExportWmiCsv.Text = "CSV Export"
    $btnExportWmiCsv.Location = New-Object System.Drawing.Point(645, 10)
    $btnExportWmiCsv.Size = New-Object System.Drawing.Size(100, 30)
    $btnExportWmiCsv.BackColor = [System.Drawing.Color]::FromArgb(230, 245, 230)

    $script:lblWmiStatus = New-Object System.Windows.Forms.Label
    $script:lblWmiStatus.Text = "Bereit."
    $script:lblWmiStatus.Location = New-Object System.Drawing.Point(12, 52)
    $script:lblWmiStatus.AutoSize = $true
    $script:lblWmiStatus.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Italic)

    $panelTop.Controls.AddRange(@(
        $lblFilterMode, $script:comboWmiViewMode, $lblSearch, $script:txtWmiSearch,
        $btnLoadWmi, $btnExportWmiCsv, $script:lblWmiStatus
    ))

    # ---------------------------------------------------------------------
    # Main SplitContainer
    # ---------------------------------------------------------------------
    $splitMain = New-Object System.Windows.Forms.SplitContainer
    $splitMain.Dock = [System.Windows.Forms.DockStyle]::Fill
    $splitMain.SplitterDistance = 640
    $splitMain.SplitterWidth = 6

    # Links: Master-Tabelle
    $panelLeft = New-Object System.Windows.Forms.Panel
    $panelLeft.Dock = [System.Windows.Forms.DockStyle]::Fill
    $panelLeft.Padding = New-Object System.Windows.Forms.Padding(10, 8, 4, 10)

    $lblMasterTitle = New-Object System.Windows.Forms.Label
    $lblMasterTitle.Text = "Gefundene WMI-Filter der Domaene:"
    $lblMasterTitle.Dock = [System.Windows.Forms.DockStyle]::Top
    $lblMasterTitle.Height = 28
    $lblMasterTitle.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)

    $script:gridWmiMaster = New-Object System.Windows.Forms.DataGridView
    $script:gridWmiMaster.Dock = [System.Windows.Forms.DockStyle]::Fill
    $script:gridWmiMaster.ReadOnly = $true
    $script:gridWmiMaster.AllowUserToAddRows = $false
    $script:gridWmiMaster.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $script:gridWmiMaster.MultiSelect = $false
    $script:gridWmiMaster.RowHeadersVisible = $false
    $script:gridWmiMaster.BackgroundColor = [System.Drawing.Color]::White
    $script:gridWmiMaster.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $script:gridWmiMaster.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(230, 236, 245)
    $script:gridWmiMaster.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    $script:gridWmiMaster.ColumnHeadersHeight = 34
    $script:gridWmiMaster.RowTemplate.Height = 28
    $script:gridWmiMaster.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(249, 251, 254)
    $script:gridWmiMaster.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells

    $panelLeft.Controls.Add($script:gridWmiMaster)
    $panelLeft.Controls.Add($lblMasterTitle)
    $splitMain.Panel1.Controls.Add($panelLeft)

    # Rechts: Geteilt in Query-Details und GPO-Zuordnungen
    $splitRight = New-Object System.Windows.Forms.SplitContainer
    $splitRight.Dock = [System.Windows.Forms.DockStyle]::Fill
    $splitRight.Orientation = [System.Windows.Forms.Orientation]::Horizontal
    $splitRight.SplitterDistance = 290
    $splitRight.SplitterWidth = 6

    # Rechts oben: WQL Query Details
    $panelQuery = New-Object System.Windows.Forms.Panel
    $panelQuery.Dock = [System.Windows.Forms.DockStyle]::Fill
    $panelQuery.Padding = New-Object System.Windows.Forms.Padding(4, 8, 10, 4)

    $panelQueryHeader = New-Object System.Windows.Forms.Panel
    $panelQueryHeader.Dock = [System.Windows.Forms.DockStyle]::Top
    $panelQueryHeader.Height = 34
    $panelQueryHeader.Padding = New-Object System.Windows.Forms.Padding(0, 2, 0, 4)

    $btnTestQuery = New-Object System.Windows.Forms.Button
    $btnTestQuery.Text = "Abfrage lokal testen"
    $btnTestQuery.Dock = [System.Windows.Forms.DockStyle]::Right
    $btnTestQuery.Width = 160
    $btnTestQuery.BackColor = [System.Drawing.Color]::FromArgb(235, 245, 255)
    $btnTestQuery.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $lblQueryTitle = New-Object System.Windows.Forms.Label
    $lblQueryTitle.Text = "WMI WQL-Abfrage & Namespace:"
    $lblQueryTitle.Dock = [System.Windows.Forms.DockStyle]::Left
    $lblQueryTitle.AutoSize = $true
    $lblQueryTitle.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)

    $panelQueryHeader.Controls.Add($btnTestQuery)
    $panelQueryHeader.Controls.Add($lblQueryTitle)

    $script:txtWmiQueryDetails = New-Object System.Windows.Forms.TextBox
    $script:txtWmiQueryDetails.Dock = [System.Windows.Forms.DockStyle]::Fill
    $script:txtWmiQueryDetails.Multiline = $true
    $script:txtWmiQueryDetails.ReadOnly = $true
    $script:txtWmiQueryDetails.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $script:txtWmiQueryDetails.BackColor = [System.Drawing.Color]::FromArgb(252, 252, 254)
    $script:txtWmiQueryDetails.Font = New-Object System.Drawing.Font("Consolas", 9.5)

    $panelQuery.Controls.Add($script:txtWmiQueryDetails)
    $panelQuery.Controls.Add($panelQueryHeader)
    $splitRight.Panel1.Controls.Add($panelQuery)

    # Rechts unten: Zugeordnete GPOs
    $panelGpo = New-Object System.Windows.Forms.Panel
    $panelGpo.Dock = [System.Windows.Forms.DockStyle]::Fill
    $panelGpo.Padding = New-Object System.Windows.Forms.Padding(4, 4, 10, 10)

    $script:lblLinkedGposTitle = New-Object System.Windows.Forms.Label
    $script:lblLinkedGposTitle.Text = "Verknuepfte Gruppenrichtlinien (GPOs):"
    $script:lblLinkedGposTitle.Dock = [System.Windows.Forms.DockStyle]::Top
    $script:lblLinkedGposTitle.Height = 28
    $script:lblLinkedGposTitle.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)

    $script:gridWmiLinkedGpos = New-Object System.Windows.Forms.DataGridView
    $script:gridWmiLinkedGpos.Dock = [System.Windows.Forms.DockStyle]::Fill
    $script:gridWmiLinkedGpos.ReadOnly = $true
    $script:gridWmiLinkedGpos.AllowUserToAddRows = $false
    $script:gridWmiLinkedGpos.RowHeadersVisible = $false
    $script:gridWmiLinkedGpos.BackgroundColor = [System.Drawing.Color]::White
    $script:gridWmiLinkedGpos.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $script:gridWmiLinkedGpos.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(230, 236, 245)
    $script:gridWmiLinkedGpos.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    $script:gridWmiLinkedGpos.ColumnHeadersHeight = 34
    $script:gridWmiLinkedGpos.RowTemplate.Height = 28
    $script:gridWmiLinkedGpos.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(249, 251, 254)
    $script:gridWmiLinkedGpos.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill

    $panelGpo.Controls.Add($script:gridWmiLinkedGpos)
    $panelGpo.Controls.Add($script:lblLinkedGposTitle)
    $splitRight.Panel2.Controls.Add($panelGpo)

    $splitMain.Panel2.Controls.Add($splitRight)
    $tabWmi.Controls.Add($splitMain)
    $tabWmi.Controls.Add($panelTop)
    $tabControl.TabPages.Add($tabWmi)

    Enable-GridSorting -Grid $script:gridWmiMaster
    Enable-GridSorting -Grid $script:gridWmiLinkedGpos

    # ---------------------------------------------------------------------
    # Interne Datenstrukturen
    # ---------------------------------------------------------------------
    $script:rawWmiList = [System.Collections.Generic.List[PSCustomObject]]::new()
    $script:wmiQueryMap = @{}
    $script:wmiLinkedGpoMap = @{}

    # ---------------------------------------------------------------------
    # Anzeige aktualisieren (Filtern & Suchen)
    # ---------------------------------------------------------------------
    $script:Update_WmiDisplay = {
        if ($script:isClosing -or $null -eq $script:gridWmiMaster -or $script:gridWmiMaster.IsDisposed) { return }

        $mode = if ($script:comboWmiViewMode -and -not $script:comboWmiViewMode.IsDisposed) {
            $script:comboWmiViewMode.SelectedItem
        } else { "Alle WMI-Filter" }

        $search = if ($script:txtWmiSearch -and -not $script:txtWmiSearch.IsDisposed) {
            "$($script:txtWmiSearch.Text)".Trim()
        } else { "" }

        $filtered = $script:rawWmiList | Where-Object {
            $item = $_
            $matchMode = switch ($mode) {
                "Nur genutzte Filter"          { [int]$item."GPO-Anzahl" -gt 0 }
                "Ungenutzte Filter (Verwaist)" { [int]$item."GPO-Anzahl" -eq 0 }
                default                        { $true }
            }
            $matchSearch = if ([string]::IsNullOrWhiteSpace($search)) { $true } else {
                $item."Filter-Name" -like "*$search*" -or $item."Beschreibung" -like "*$search*" -or $item."WQL-Query" -like "*$search*"
            }
            $matchMode -and $matchSearch
        }

        $arr = [System.Collections.ArrayList]::new()
        foreach ($it in $filtered) { [void]$arr.Add($it) }
        $script:gridWmiMaster.DataSource = $arr

        if ($script:gridWmiMaster.Columns["WQL-Query"]) { $script:gridWmiMaster.Columns["WQL-Query"].Visible = $false }
        if ($script:gridWmiMaster.Columns["Namespace"]) { $script:gridWmiMaster.Columns["Namespace"].Visible = $false }

        if ($script:lblWmiStatus -and -not $script:lblWmiStatus.IsDisposed) {
            $orphanCount = @($script:rawWmiList | Where-Object { [int]$_."GPO-Anzahl" -eq 0 }).Count
            $script:lblWmiStatus.Text = "Status: $($arr.Count) von $($script:rawWmiList.Count) WMI-Filtern angezeigt | $orphanCount ungenutzt (verwaist)."
        }

        if ($script:gridWmiMaster.Rows.Count -gt 0) {
            $script:gridWmiMaster.Rows[0].Selected = $true
        } else {
            $script:txtWmiQueryDetails.Clear()
            $script:gridWmiLinkedGpos.DataSource = $null
        }
    }

    # ---------------------------------------------------------------------
    # Farbliche Kennzeichnung
    # ---------------------------------------------------------------------
    $script:gridWmiMaster.Add_DataBindingComplete({
        if ($script:isClosing -or $null -eq $script:gridWmiMaster -or $script:gridWmiMaster.IsDisposed) { return }
        foreach ($row in $script:gridWmiMaster.Rows) {
            $gpoCount = [int]$row.Cells["GPO-Anzahl"].Value
            if ($gpoCount -eq 0) {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 248, 225)
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(180, 90, 0)
            } else {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(235, 247, 235)
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkGreen
            }
        }
    })

    # ---------------------------------------------------------------------
    # Laderoutine: Dual-Resolution (Native GPO API + LDAP)
    # ---------------------------------------------------------------------
    $script:Invoke_LoadWmiFilters = {
        if ($script:isClosing -or $form.IsDisposed) { return }

        $targetDN = if ($domainDN) { $domainDN } else { ([ADSI]"LDAP://RootDSE").defaultNamingContext.Value }

        if ($script:pbarGlobal) {
            $script:pbarGlobal.Visible = $true
            $script:pbarGlobal.Minimum = 0
            $script:pbarGlobal.Value = 0
        }
        if ($script:lblProgressInfo) { $script:lblProgressInfo.Text = "Lese WMI-Filter und GPO-Verknuepfungen aus..." }
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        [System.Windows.Forms.Application]::DoEvents()

        $wmiRoot = $null
        $wmiSearcher = $null
        $gpoRoot = $null
        $gpoSearcher = $null

        try {
            $script:rawWmiList.Clear()
            $script:wmiQueryMap.Clear()
            $script:wmiLinkedGpoMap.Clear()

            $gpoUsageByGuid = @{}
            $gpoUsageByName = @{}

            function Register-GpoLink ($key, $gpoObj, $isGuid = $false) {
                if ([string]::IsNullOrWhiteSpace($key)) { return }
                $cleanKey = if ($isGuid) {
                    if ($key -match '(?i)([a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12})') { $matches[1].ToUpper() } else { $key.Trim('{','}').ToUpper() }
                } else {
                    $key.Trim().ToLower()
                }

                $dict = if ($isGuid) { $gpoUsageByGuid } else { $gpoUsageByName }
                if (-not $dict.ContainsKey($cleanKey)) {
                    $dict[$cleanKey] = [System.Collections.Generic.List[PSCustomObject]]::new()
                }
                $exists = $dict[$cleanKey] | Where-Object { $_."GPO GUID" -eq $gpoObj."GPO GUID" }
                if (-not $exists) {
                    $dict[$cleanKey].Add($gpoObj)
                }
            }

            # 1. Native Abfrage ueber Get-GPO -All (hoechste Genauigkeit)
            try {
                $allGposNative = Get-GPO -All -ErrorAction Stop
                foreach ($g in $allGposNative) {
                    if ($g.WmiFilter) {
                        $gpoItem = [PSCustomObject]@{
                            "GPO Name" = $g.DisplayName
                            "GPO GUID" = "{$($g.Id.ToString().Trim('{','}').ToUpper())}"
                        }
                        if ($g.WmiFilter.Name) { Register-GpoLink $g.WmiFilter.Name $gpoItem $false }
                        if ($g.WmiFilter.Id)   { Register-GpoLink $g.WmiFilter.Id.ToString() $gpoItem $true }
                    }
                }
            } catch {
                # 2. Fallback: LDAP-Abfrage auf gPCWQLFilter
                try {
                    $gpoRoot = [System.DirectoryServices.DirectoryEntry]::new("LDAP://CN=Policies,CN=System,$targetDN")
                    $gpoSearcher = [System.DirectoryServices.DirectorySearcher]::new($gpoRoot)
                    $gpoSearcher.PageSize = 1000
                    $gpoSearcher.Filter = "(&(objectClass=groupPolicyContainer)(gPCWQLFilter=*))"
                    $gpoSearcher.PropertiesToLoad.AddRange(@("displayName", "name", "gPCWQLFilter"))
                    $foundGpos = $gpoSearcher.FindAll()

                    foreach ($g in $foundGpos) {
                        $gName = if ($g.Properties["displayname"]) { "$($g.Properties['displayname'][0])" } else { "$($g.Properties['name'][0])" }
                        $gGuid = "{$($g.Properties['name'][0].ToString().Trim('{','}').ToUpper())}"
                        $rawFilter = if ($g.Properties["gpcwqlfilter"]) { "$($g.Properties['gpcwqlfilter'][0])" } else { "" }

                        $gpoItem = [PSCustomObject]@{ "GPO Name" = $gName; "GPO GUID" = $gGuid }

                        if ($rawFilter -match '\[\s*[^;]+;\s*(?<ref>[^;]+);\s*\d+\s*\]') {
                            $ref = $matches["ref"].Trim()
                            Register-GpoLink $ref $gpoItem $false
                            if ($ref -match '(?i)[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}') {
                                Register-GpoLink $ref $gpoItem $true
                            }
                        }
                        if ($rawFilter -match '(?i)([a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12})') {
                            Register-GpoLink $matches[1] $gpoItem $true
                        }
                    }
                } catch {}
            }

            # 3. WMI-Filter abfragen (msWMI-Som)
            $wmiRoot = [System.DirectoryServices.DirectoryEntry]::new("LDAP://CN=SOM,CN=WMIPolicy,CN=System,$targetDN")
            $wmiSearcher = [System.DirectoryServices.DirectorySearcher]::new($wmiRoot)
            $wmiSearcher.PageSize = 1000
            $wmiSearcher.Filter = "(objectClass=msWMI-Som)"
            $wmiSearcher.PropertiesToLoad.AddRange(@("msWMI-Name", "msWMI-ID", "msWMI-Parm1", "msWMI-Parm2", "msWMI-Author", "whenCreated", "whenChanged", "name"))
            $wmiResults = $wmiSearcher.FindAll()

            if ($script:pbarGlobal) { $script:pbarGlobal.Maximum = [Math]::Max(1, $wmiResults.Count) }
            $idx = 0

            foreach ($w in $wmiResults) {
                $idx++
                if ($script:pbarGlobal) { $script:pbarGlobal.Value = $idx }

                $fName = if ($w.Properties["mswmi-name"]) { "$($w.Properties['mswmi-name'][0])" } else { "Unbenannter Filter" }
                
                $rawId = if ($w.Properties["mswmi-id"]) { "$($w.Properties['mswmi-id'][0])" } elseif ($w.Properties["name"]) { "$($w.Properties['name'][0])" } else { "" }
                $cleanId = if ($rawId -match '(?i)([a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12})') {
                    $matches[1].ToUpper()
                } else {
                    $rawId.Trim('{','}').ToUpper()
                }

                $desc = if ($w.Properties["mswmi-parm1"]) { "$($w.Properties['mswmi-parm1'][0])" } else { "-" }
                $author = if ($w.Properties["mswmi-author"]) { "$($w.Properties['mswmi-author'][0])" } else { "-" }
                $created = if ($w.Properties["whencreated"]) { (Get-Date $w.Properties["whencreated"][0]).ToString("dd.MM.yyyy HH:mm") } else { "-" }
                $changed = if ($w.Properties["whenchanged"]) { (Get-Date $w.Properties["whenchanged"][0]).ToString("dd.MM.yyyy HH:mm") } else { "-" }

                $rawParm2 = if ($w.Properties["mswmi-parm2"]) { "$($w.Properties['mswmi-parm2'][0])" } else { "" }
                $parsedQ = Parse-WmiParm2 $rawParm2

                # Verknuepfte GPOs ueber GUID UND Name zusammenfuehren
                $linkedList = [System.Collections.Generic.List[PSCustomObject]]::new()
                if (-not [string]::IsNullOrWhiteSpace($cleanId) -and $gpoUsageByGuid.ContainsKey($cleanId)) {
                    foreach ($gp in $gpoUsageByGuid[$cleanId]) {
                        if (-not ($linkedList | Where-Object { $_."GPO GUID" -eq $gp."GPO GUID" })) { $linkedList.Add($gp) }
                    }
                }
                $kName = $fName.Trim().ToLower()
                if (-not [string]::IsNullOrWhiteSpace($kName) -and $gpoUsageByName.ContainsKey($kName)) {
                    foreach ($gp in $gpoUsageByName[$kName]) {
                        if (-not ($linkedList | Where-Object { $_."GPO GUID" -eq $gp."GPO GUID" })) { $linkedList.Add($gp) }
                    }
                }

                $finalGpos = @($linkedList | Sort-Object "GPO Name")
                if (-not [string]::IsNullOrWhiteSpace($cleanId)) { $script:wmiLinkedGpoMap[$cleanId] = $finalGpos }
                if (-not [string]::IsNullOrWhiteSpace($kName))   { $script:wmiLinkedGpoMap[$kName]   = $finalGpos }
                if (-not [string]::IsNullOrWhiteSpace($cleanId)) { $script:wmiQueryMap[$cleanId]     = $parsedQ }

                $script:rawWmiList.Add([PSCustomObject]@{
                    "Filter-Name"   = $fName
                    "GPO-Anzahl"    = $finalGpos.Count
                    "Nutzung"       = if ($finalGpos.Count -gt 0) { "Verknuepft ($($finalGpos.Count) GPOs)" } else { "Ungenutzt (Verwaist)" }
                    "Beschreibung"  = $desc
                    "Autor"         = $author
                    "Filter GUID"   = "{$cleanId}"
                    "Erstellt am"   = $created
                    "Geaendert am"  = $changed
                    "Namespace"     = $parsedQ.Namespace
                    "WQL-Query"     = $parsedQ.Query
                })
            }

            & $script:Update_WmiDisplay
            if ($script:lblProgressInfo) { $script:lblProgressInfo.Text = "WMI-Filter geladen: $($script:rawWmiList.Count) Filter gefunden." }
        } catch {
            if ($script:lblProgressInfo) { $script:lblProgressInfo.Text = "Fehler: $($_.Exception.Message)" }
        } finally {
            if ($wmiSearcher)  { $wmiSearcher.Dispose() }
            if ($wmiRoot)      { $wmiRoot.Dispose() }
            if ($gpoSearcher)  { $gpoSearcher.Dispose() }
            if ($gpoRoot)      { $gpoRoot.Dispose() }
            if ($script:pbarGlobal) { $script:pbarGlobal.Visible = $false }
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }

    # ---------------------------------------------------------------------
    # Selektionswechsel (Sichere Schluesselabfrage ohne Null-Werte)
    # ---------------------------------------------------------------------
    $script:gridWmiMaster.Add_SelectionChanged({
        if ($script:isClosing -or $null -eq $script:gridWmiMaster -or $script:gridWmiMaster.IsDisposed) { return }
        if ($script:gridWmiMaster.SelectedRows.Count -gt 0) {
            $row = $script:gridWmiMaster.SelectedRows[0]
            $cellVal = "$($row.Cells['Filter GUID'].Value)"
            $fGuid = if ($cellVal -match '(?i)([a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12})') {
                $matches[1].ToUpper()
            } else {
                $cellVal.Trim('{','}').ToUpper()
            }
            $fName = "$($row.Cells['Filter-Name'].Value)"
            $kName = $fName.Trim().ToLower()

            if (-not [string]::IsNullOrWhiteSpace($fGuid) -and $script:wmiQueryMap.ContainsKey($fGuid)) {
                $qInfo = $script:wmiQueryMap[$fGuid]
                $script:txtWmiQueryDetails.Text = "FILTER-NAME : $fName`r`nNAMESPACE   : $($qInfo.Namespace)`r`nGUID        : {$fGuid}`r`n`r`n[WQL ABFRAGE]`r`n$($qInfo.Query)"
            }

            # Sichere Auflösung der verknüpften GPOs
            $gpos = @()
            if (-not [string]::IsNullOrWhiteSpace($fGuid) -and $script:wmiLinkedGpoMap.ContainsKey($fGuid)) {
                $gpos = $script:wmiLinkedGpoMap[$fGuid]
            } elseif (-not [string]::IsNullOrWhiteSpace($kName) -and $script:wmiLinkedGpoMap.ContainsKey($kName)) {
                $gpos = $script:wmiLinkedGpoMap[$kName]
            }

            $script:lblLinkedGposTitle.Text = "Verknuepfte Gruppenrichtlinien fuer '$fName' ($($gpos.Count)): "

            $arrGpo = [System.Collections.ArrayList]::new()
            if ($gpos.Count -gt 0) {
                foreach ($gp in $gpos) { [void]$arrGpo.Add($gp) }
            } else {
                [void]$arrGpo.Add([PSCustomObject]@{
                    "GPO Name" = "-- Keine GPO verknuepft --"
                    "GPO GUID" = "[Hinweis] Dieser WMI-Filter wird derzeit von keiner Richtlinie verwendet."
                })
            }
            $script:gridWmiLinkedGpos.DataSource = $arrGpo
        }
    })

    # ---------------------------------------------------------------------
    # Lokaler Test der WMI-Query
    # ---------------------------------------------------------------------
    $btnTestQuery.Add_Click({
        if ($script:gridWmiMaster.SelectedRows.Count -eq 0) { return }
        $cellVal = "$($script:gridWmiMaster.SelectedRows[0].Cells['Filter GUID'].Value)"
        $fGuid = if ($cellVal -match '(?i)([a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12})') {
            $matches[1].ToUpper()
        } else {
            $cellVal.Trim('{','}').ToUpper()
        }

        if ([string]::IsNullOrWhiteSpace($fGuid) -or -not $script:wmiQueryMap.ContainsKey($fGuid)) { return }

        $qInfo = $script:wmiQueryMap[$fGuid]
        $ns = $qInfo.Namespace
        $query = $qInfo.Query

        if ([string]::IsNullOrWhiteSpace($query) -or $query.StartsWith("--")) {
            [System.Windows.Forms.MessageBox]::Show("Keine valide WQL-Abfrage vorhanden.", "Hinweis", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        try {
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            $results = @(Get-CimInstance -Namespace $ns -Query $query -ErrorAction Stop)
            $form.Cursor = [System.Windows.Forms.Cursors]::Default

            $matchText = if ($results.Count -gt 0) {
                "[POSITIV - Match]`r`nDie Abfrage trifft auf dieses System zu!`r`nErgebnis: $($results.Count) Objekt(e) zurueckgegeben."
            } else {
                "[NEGATIV - Kein Match]`r`nDie Abfrage liefert 0 Treffer auf diesem System (GPO wuerde hier NICHT greifen)."
            }

            [System.Windows.Forms.MessageBox]::Show("Namespace: $ns`r`nAbfrage: $query`r`n`r`nErgebnis auf diesem Host:`r`n$matchText", "WMI Test-Ergebnis", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } catch {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Ausfuehren der WMI-Abfrage:`r`n$($_.Exception.Message)", "WMI Syntax-/Abfragefehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })

    # ---------------------------------------------------------------------
    # Events
    # ---------------------------------------------------------------------
    $btnLoadWmi.Add_Click({ & $script:Invoke_LoadWmiFilters })
    $script:comboWmiViewMode.Add_SelectedIndexChanged({ & $script:Update_WmiDisplay })
    $script:txtWmiSearch.Add_TextChanged({ & $script:Update_WmiDisplay })

    $btnExportWmiCsv.Add_Click({
        if ($script:rawWmiList.Count -eq 0) { return }
        $targetBase = if ($script:txtBackupTargetDir) { $script:txtBackupTargetDir.Text.Trim() } else { "C:\Install\Backup\GPO" }
        if (-not (Test-Path $targetBase)) { New-Item -ItemType Directory -Path $targetBase -Force | Out-Null }
        $csvFile = Join-Path $targetBase "WMI_Filter_Analyse_$((Get-Date).ToString('yyyyMMdd_HHmm')).csv"
        $script:rawWmiList | Export-Csv -Path $csvFile -Delimiter ";" -NoTypeInformation -Encoding UTF8
        [System.Windows.Forms.MessageBox]::Show("Exportiert nach: $csvFile", "Export fertig", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    })
}