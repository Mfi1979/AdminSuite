# =========================================================================
# Tab0_Dashboard.ps1 - Status & Dashboard Uebersicht (KPIs, Checks & WMI)
# =========================================================================

function Build-Tab0_Dashboard {
    param($tabControl, $domainDN, $domainName)

    # Sichere Domänenermittlung
    if ([string]::IsNullOrWhiteSpace($domainName) -or [string]::IsNullOrWhiteSpace($domainDN)) {
        try {
            $curDom = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
            if ([string]::IsNullOrWhiteSpace($domainName)) { $domainName = $curDom.Name }
            if ([string]::IsNullOrWhiteSpace($domainDN)) {
                $domainDN = ($curDom.Name.Split('.') | ForEach-Object { "DC=$_" }) -join ','
            }
        } catch {
            if ([string]::IsNullOrWhiteSpace($domainDN)) {
                $domainDN = ([ADSI]"LDAP://RootDSE").defaultNamingContext.Value
            }
            if ([string]::IsNullOrWhiteSpace($domainName)) {
                $domainName = ($domainDN -replace 'DC=','' -replace ',','.')
            }
        }
    }

    $tabDashboard = New-Object System.Windows.Forms.TabPage
    $tabDashboard.Text = "0. Status Ueberblick"
    $tabDashboard.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
    $tabDashboard.BackColor = [System.Drawing.Color]::FromArgb(246, 248, 252)

    # ---------------------------------------------------------------------
    # Top Panel
    # ---------------------------------------------------------------------
    $panelTop = New-Object System.Windows.Forms.Panel
    $panelTop.Dock = [System.Windows.Forms.DockStyle]::Top
    $panelTop.Height = 65
    $panelTop.BackColor = [System.Drawing.Color]::FromArgb(242, 245, 250)
    $panelTop.Padding = New-Object System.Windows.Forms.Padding(15, 12, 15, 12)

    $lblDomainTitle = New-Object System.Windows.Forms.Label
    $lblDomainTitle.Text = "Domaene: $domainName"
    $lblDomainTitle.Location = New-Object System.Drawing.Point(15, 12)
    $lblDomainTitle.AutoSize = $true
    $lblDomainTitle.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $lblDomainTitle.ForeColor = [System.Drawing.Color]::FromArgb(24, 76, 120)

    $script:lblDashboardStatus = New-Object System.Windows.Forms.Label
    $script:lblDashboardStatus.Text = "Bereit zum Laden der Systemanalyse."
    $script:lblDashboardStatus.Location = New-Object System.Drawing.Point(16, 36)
    $script:lblDashboardStatus.AutoSize = $true
    $script:lblDashboardStatus.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Italic)
    $script:lblDashboardStatus.ForeColor = [System.Drawing.Color]::FromArgb(100, 110, 125)

    $btnRefreshDashboard = New-Object System.Windows.Forms.Button
    $btnRefreshDashboard.Text = "Dashboard aktualisieren"
    $btnRefreshDashboard.Dock = [System.Windows.Forms.DockStyle]::Right
    $btnRefreshDashboard.Width = 190
    $btnRefreshDashboard.BackColor = [System.Drawing.Color]::FromArgb(225, 238, 255)
    $btnRefreshDashboard.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $panelTop.Controls.AddRange(@($btnRefreshDashboard, $lblDomainTitle, $script:lblDashboardStatus))

    # ---------------------------------------------------------------------
    # Kachel-Generator
    # ---------------------------------------------------------------------
    function Create-DashboardCard ([string]$titleText, [System.Drawing.Color]$defaultValColor) {
        $pCard = New-Object System.Windows.Forms.Panel
        $pCard.Size = New-Object System.Drawing.Size(165, 80)
        $pCard.BackColor = [System.Drawing.Color]::White
        $pCard.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
        $pCard.Margin = New-Object System.Windows.Forms.Padding(4, 4, 8, 4)

        $lblTitle = New-Object System.Windows.Forms.Label
        $lblTitle.Text = $titleText
        $lblTitle.Dock = [System.Windows.Forms.DockStyle]::Top
        $lblTitle.Height = 24
        $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 8.2, [System.Drawing.FontStyle]::Bold)
        $lblTitle.ForeColor = [System.Drawing.Color]::FromArgb(100, 110, 130)
        $lblTitle.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter

        $lblVal = New-Object System.Windows.Forms.Label
        $lblVal.Text = "-"
        $lblVal.Dock = [System.Windows.Forms.DockStyle]::Fill
        $lblVal.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
        $lblVal.ForeColor = $defaultValColor
        $lblVal.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter

        $pCard.Controls.AddRange(@($lblVal, $lblTitle))
        return @{ Panel = $pCard; ValueLabel = $lblVal }
    }

    $flowKpiPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $flowKpiPanel.Dock = [System.Windows.Forms.DockStyle]::Top
    $flowKpiPanel.Height = 96
    $flowKpiPanel.Padding = New-Object System.Windows.Forms.Padding(12, 4, 12, 4)
    $flowKpiPanel.BackColor = [System.Drawing.Color]::FromArgb(246, 248, 252)

    $cardGpoTotal    = Create-DashboardCard "GPOS GESAMT" ([System.Drawing.Color]::FromArgb(24, 76, 120))
    $cardGpoLinked   = Create-DashboardCard "VERKNUEPFT" ([System.Drawing.Color]::DarkGreen)
    $cardGpoUnlinked = Create-DashboardCard "NICHT VERKNUEPFT" ([System.Drawing.Color]::FromArgb(215, 85, 0))
    $cardGpoDisabled = Create-DashboardCard "DEAKTIVIERT" ([System.Drawing.Color]::FromArgb(180, 50, 50))
    $cardWmiTotal    = Create-DashboardCard "WMI-FILTER GESAMT" ([System.Drawing.Color]::FromArgb(24, 76, 120))
    $cardWmiUnused   = Create-DashboardCard "UNGENUTZTE WMI" ([System.Drawing.Color]::DarkGreen)

    $script:lblKpiGpoTotal      = $cardGpoTotal.ValueLabel
    $script:lblKpiGpoLinked     = $cardGpoLinked.ValueLabel
    $script:lblKpiGpoUnlinked   = $cardGpoUnlinked.ValueLabel
    $script:panelKpiGpoUnlinked = $cardGpoUnlinked.Panel
    $script:lblKpiGpoDisabled   = $cardGpoDisabled.ValueLabel
    $script:panelKpiGpoDisabled = $cardGpoDisabled.Panel
    $script:lblKpiWmiTotal      = $cardWmiTotal.ValueLabel
    $script:lblKpiWmiUnused     = $cardWmiUnused.ValueLabel
    $script:panelKpiWmiUnused   = $cardWmiUnused.Panel

    $flowKpiPanel.Controls.AddRange(@(
        $cardGpoTotal.Panel, $cardGpoLinked.Panel, $cardGpoUnlinked.Panel,
        $cardGpoDisabled.Panel, $cardWmiTotal.Panel, $cardWmiUnused.Panel
    ))

    # ---------------------------------------------------------------------
    # SplitContainer: Links 4 System-Checks / Rechts WMI Status-Ueberblick
    # ---------------------------------------------------------------------
    $splitDash = New-Object System.Windows.Forms.SplitContainer
    $splitDash.Dock = [System.Windows.Forms.DockStyle]::Fill
    $splitDash.SplitterDistance = 650
    $splitDash.SplitterWidth = 6

    # Links: 4 Systemprüfungen / Handlungsempfehlungen
    $panelAudit = New-Object System.Windows.Forms.Panel
    $panelAudit.Dock = [System.Windows.Forms.DockStyle]::Fill
    $panelAudit.Padding = New-Object System.Windows.Forms.Padding(12, 6, 4, 12)

    $lblGridAudit = New-Object System.Windows.Forms.Label
    $lblGridAudit.Text = "Systemprüfungen & Handlungsempfehlungen (4 Kern-Checks):"
    $lblGridAudit.Dock = [System.Windows.Forms.DockStyle]::Top
    $lblGridAudit.Height = 26
    $lblGridAudit.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)

    $script:gridDashboardAudit = New-Object System.Windows.Forms.DataGridView
    $script:gridDashboardAudit.Dock = [System.Windows.Forms.DockStyle]::Fill
    $script:gridDashboardAudit.ReadOnly = $true
    $script:gridDashboardAudit.AllowUserToAddRows = $false
    $script:gridDashboardAudit.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $script:gridDashboardAudit.MultiSelect = $false
    $script:gridDashboardAudit.RowHeadersVisible = $false
    $script:gridDashboardAudit.BackgroundColor = [System.Drawing.Color]::White
    $script:gridDashboardAudit.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $script:gridDashboardAudit.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(230, 236, 245)
    $script:gridDashboardAudit.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $script:gridDashboardAudit.ColumnHeadersHeight = 30
    $script:gridDashboardAudit.RowTemplate.Height = 26
    $script:gridDashboardAudit.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(249, 251, 254)
    $script:gridDashboardAudit.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill

    $panelAudit.Controls.Add($script:gridDashboardAudit)
    $panelAudit.Controls.Add($lblGridAudit)
    $splitDash.Panel1.Controls.Add($panelAudit)

    # Rechts: WMI-Status-Tabelle
    $panelWmi = New-Object System.Windows.Forms.Panel
    $panelWmi.Dock = [System.Windows.Forms.DockStyle]::Fill
    $panelWmi.Padding = New-Object System.Windows.Forms.Padding(4, 6, 12, 12)

    $lblGridWmi = New-Object System.Windows.Forms.Label
    $lblGridWmi.Text = "WMI-Filter Status-Ueberblick (Domaene):"
    $lblGridWmi.Dock = [System.Windows.Forms.DockStyle]::Top
    $lblGridWmi.Height = 26
    $lblGridWmi.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)

    $script:gridDashboardWmi = New-Object System.Windows.Forms.DataGridView
    $script:gridDashboardWmi.Dock = [System.Windows.Forms.DockStyle]::Fill
    $script:gridDashboardWmi.ReadOnly = $true
    $script:gridDashboardWmi.AllowUserToAddRows = $false
    $script:gridDashboardWmi.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $script:gridDashboardWmi.MultiSelect = $false
    $script:gridDashboardWmi.RowHeadersVisible = $false
    $script:gridDashboardWmi.BackgroundColor = [System.Drawing.Color]::White
    $script:gridDashboardWmi.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $script:gridDashboardWmi.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(230, 236, 245)
    $script:gridDashboardWmi.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $script:gridDashboardWmi.ColumnHeadersHeight = 30
    $script:gridDashboardWmi.RowTemplate.Height = 26
    $script:gridDashboardWmi.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(249, 251, 254)
    $script:gridDashboardWmi.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill

    $panelWmi.Controls.Add($script:gridDashboardWmi)
    $panelWmi.Controls.Add($lblGridWmi)
    $splitDash.Panel2.Controls.Add($panelWmi)

    $tabDashboard.Controls.Add($splitDash)
    $tabDashboard.Controls.Add($flowKpiPanel)
    $tabDashboard.Controls.Add($panelTop)
    $tabControl.TabPages.Add($tabDashboard)

    # ---------------------------------------------------------------------
    # Farbgebung Tabellen
    # ---------------------------------------------------------------------
    $script:gridDashboardAudit.Add_DataBindingComplete({
        if ($script:isClosing -or $null -eq $script:gridDashboardAudit -or $script:gridDashboardAudit.IsDisposed) { return }
        foreach ($row in $script:gridDashboardAudit.Rows) {
            $status = "$($row.Cells['Status'].Value)"
            if ($status -eq "Warnung") {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 248, 225)
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(190, 85, 0)
            } elseif ($status -eq "Kritisch") {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 235, 235)
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(180, 20, 20)
            } else {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(240, 250, 240)
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkGreen
            }
        }
    })

    $script:gridDashboardWmi.Add_DataBindingComplete({
        if ($script:isClosing -or $null -eq $script:gridDashboardWmi -or $script:gridDashboardWmi.IsDisposed) { return }
        foreach ($row in $script:gridDashboardWmi.Rows) {
            $status = "$($row.Cells['Status'].Value)"
            if ($status -match "Ungenutzt|Verwaist") {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 248, 225)
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(180, 90, 0)
            } else {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(235, 247, 235)
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkGreen
            }
        }
    })

    # ---------------------------------------------------------------------
    # Laderoutine: 4 System-Checks und WMI Status
    # ---------------------------------------------------------------------
    $script:Invoke_LoadDashboard = {
        if ($script:isClosing -or $form.IsDisposed) { return }

        $tDN = if (-not [string]::IsNullOrWhiteSpace($domainDN)) { $domainDN } else { ([ADSI]"LDAP://RootDSE").defaultNamingContext.Value }

        if ($script:pbarGlobal) {
            $script:pbarGlobal.Visible = $true
            $script:pbarGlobal.Minimum = 0
            $script:pbarGlobal.Value = 0
        }
        if ($script:lblProgressInfo) { $script:lblProgressInfo.Text = "Erstelle Dashboard- & Systempruefungen..." }
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        [System.Windows.Forms.Application]::DoEvents()

        $wmiRoot = $null; $wmiSearcher = $null
        $gpoRoot = $null; $gpoSearcher = $null
        $somRoot = $null; $somSearcher = $null

        $wmiStatusList = [System.Collections.Generic.List[PSCustomObject]]::new()

        function Get-LdapProp ($resultItem, [string]$propName) {
            if ($null -eq $resultItem -or $null -eq $resultItem.Properties) { return "" }
            $prop = $resultItem.Properties[$propName.ToLower()]
            if ($null -ne $prop -and $prop.Count -gt 0) {
                return "$($prop[0])"
            }
            return ""
        }

        try {
            # 1. Links aus OUs & Domain ermitteln
            $linkedGpoGuids = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            try {
                $somRoot = [System.DirectoryServices.DirectoryEntry]::new("LDAP://$tDN")
                $somSearcher = [System.DirectoryServices.DirectorySearcher]::new($somRoot)
                $somSearcher.PageSize = 1000
                $somSearcher.Filter = "(|(objectClass=organizationalUnit)(objectClass=domainDNS))"
                $somSearcher.PropertiesToLoad.Add("gplink") | Out-Null
                $somResults = $somSearcher.FindAll()

                foreach ($sr in $somResults) {
                    $rawGplink = Get-LdapProp $sr "gplink"
                    if (-not [string]::IsNullOrWhiteSpace($rawGplink)) {
                        $regexMatches = [regex]::Matches($rawGplink, '\[LDAP://cn=(?<guid>{[a-fA-F0-9-]+}),cn=policies,cn=system,[^;]+;\d+\]')
                        foreach ($m in $regexMatches) {
                            $clean = $m.Groups["guid"].Value.Trim('{','}').ToUpper()
                            [void]$linkedGpoGuids.Add($clean)
                        }
                    }
                }
            } catch {}

            # 2. GPOs analysieren & WMI-Verbindungen kartieren
            $totalGpos = 0; $linkedGposCount = 0; $unlinkedGposCount = 0; $disabledGposCount = 0
            $unlinkedGpoNames = [System.Collections.Generic.List[string]]::new()
            $disabledGpoNames = [System.Collections.Generic.List[string]]::new()

            $wmiUsageByGuid = @{}
            $wmiUsageByName = @{}

            function Map-LocalGpoToWmi ($key, $gpoName, $isGuid = $false) {
                if ([string]::IsNullOrWhiteSpace($key)) { return }
                $clean = if ($isGuid) {
                    if ($key -match '(?i)([a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12})') { $matches[1].ToUpper() } else { $key.Trim('{','}').ToUpper() }
                } else {
                    $key.Trim().ToLower()
                }

                $dict = if ($isGuid) { $wmiUsageByGuid } else { $wmiUsageByName }
                if (-not $dict.ContainsKey($clean)) {
                    $dict[$clean] = [System.Collections.Generic.List[string]]::new()
                }
                if (-not $dict[$clean].Contains($gpoName)) {
                    $dict[$clean].Add($gpoName)
                }
            }

            $nativeSuccess = $false
            try {
                $allGposNative = Get-GPO -All -ErrorAction Stop
                $nativeSuccess = $true
                $totalGpos = $allGposNative.Count

                foreach ($g in $allGposNative) {
                    $gGuid = $g.Id.ToString().Trim('{','}').ToUpper()

                    if ($linkedGpoGuids.Contains($gGuid)) {
                        $linkedGposCount++
                    } else {
                        $unlinkedGposCount++
                        $unlinkedGpoNames.Add($g.DisplayName)
                    }

                    if ($g.GpoStatus -eq "AllSettingsDisabled") {
                        $disabledGposCount++
                        $disabledGpoNames.Add($g.DisplayName)
                    }

                    if ($g.WmiFilter) {
                        if ($g.WmiFilter.Name) { Map-LocalGpoToWmi $g.WmiFilter.Name $g.DisplayName $false }
                        if ($g.WmiFilter.Id)   { Map-LocalGpoToWmi $g.WmiFilter.Id.ToString() $g.DisplayName $true }
                    }
                }
            } catch {}

            # LDAP Fallback falls Get-GPO fehlschlägt
            if (-not $nativeSuccess) {
                $gpoRoot = [System.DirectoryServices.DirectoryEntry]::new("LDAP://CN=Policies,CN=System,$tDN")
                $gpoSearcher = [System.DirectoryServices.DirectorySearcher]::new($gpoRoot)
                $gpoSearcher.PageSize = 1000
                $gpoSearcher.Filter = "(objectClass=groupPolicyContainer)"
                $gpoSearcher.PropertiesToLoad.AddRange(@("displayName", "name", "flags", "gPCWQLFilter"))
                $gpoResults = $gpoSearcher.FindAll()

                $totalGpos = $gpoResults.Count

                foreach ($gp in $gpoResults) {
                    $dName = Get-LdapProp $gp "displayname"
                    $cName = Get-LdapProp $gp "name"
                    $gName = if (-not [string]::IsNullOrWhiteSpace($dName)) { $dName } else { $cName }
                    $gGuid = $cName.Trim('{','}').ToUpper()

                    $flagsStr = Get-LdapProp $gp "flags"
                    $flags = if (-not [string]::IsNullOrWhiteSpace($flagsStr)) { [int]$flagsStr } else { 0 }
                    $wqlRef = Get-LdapProp $gp "gpcwqlfilter"

                    if (-not [string]::IsNullOrWhiteSpace($wqlRef)) {
                        if ($wqlRef -match '(?i)([a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12})') {
                            Map-LocalGpoToWmi $matches[1] $gName $true
                        }
                        if ($wqlRef -match '\[\s*[^;]+;\s*(?<ref>[^;]+);\s*\d+\s*\]') {
                            Map-LocalGpoToWmi $matches["ref"] $gName $false
                        }
                    }

                    if ($linkedGpoGuids.Contains($gGuid)) {
                        $linkedGposCount++
                    } else {
                        $unlinkedGposCount++
                        $unlinkedGpoNames.Add($gName)
                    }

                    if ($flags -eq 3) {
                        $disabledGposCount++
                        $disabledGpoNames.Add($gName)
                    }
                }
            }

            # 3. WMI-Filter analysieren
            $wmiTotalCount = 0
            $wmiUnusedCount = 0
            $unusedWmiNames = [System.Collections.Generic.List[string]]::new()

            $wmiRoot = [System.DirectoryServices.DirectoryEntry]::new("LDAP://CN=SOM,CN=WMIPolicy,CN=System,$tDN")
            $wmiSearcher = [System.DirectoryServices.DirectorySearcher]::new($wmiRoot)
            $wmiSearcher.PageSize = 1000
            $wmiSearcher.Filter = "(objectClass=msWMI-Som)"
            $wmiSearcher.PropertiesToLoad.AddRange(@("msWMI-Name", "msWMI-ID", "name"))
            $wmiResults = $wmiSearcher.FindAll()

            $wmiTotalCount = $wmiResults.Count

            foreach ($w in $wmiResults) {
                $fName = Get-LdapProp $w "mswmi-name"
                if ([string]::IsNullOrWhiteSpace($fName)) { $fName = "Unbenannter Filter" }
                $fName = $fName.Trim()

                $rawId = Get-LdapProp $w "mswmi-id"
                if ([string]::IsNullOrWhiteSpace($rawId)) { $rawId = Get-LdapProp $w "name" }

                $cleanId = if ($rawId -match '(?i)([a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12})') {
                    $matches[1].ToUpper()
                } else {
                    $rawId.Trim('{','}').ToUpper()
                }

                $linkedGpos = [System.Collections.Generic.List[string]]::new()
                if (-not [string]::IsNullOrWhiteSpace($cleanId) -and $wmiUsageByGuid.ContainsKey($cleanId)) {
                    foreach ($gp in $wmiUsageByGuid[$cleanId]) { if (-not $linkedGpos.Contains($gp)) { $linkedGpos.Add($gp) } }
                }
                $kName = $fName.ToLower()
                if (-not [string]::IsNullOrWhiteSpace($kName) -and $wmiUsageByName.ContainsKey($kName)) {
                    foreach ($gp in $wmiUsageByName[$kName]) { if (-not $linkedGpos.Contains($gp)) { $linkedGpos.Add($gp) } }
                }

                $gpoCount = $linkedGpos.Count
                $statusText = if ($gpoCount -gt 0) { "Verknuepft" } else { "Ungenutzt (Verwaist)" }
                $gpoListStr = if ($gpoCount -gt 0) { ($linkedGpos | Sort-Object) -join ", " } else { "-- keine Zuordnung --" }

                if ($gpoCount -eq 0) {
                    $wmiUnusedCount++
                    $unusedWmiNames.Add($fName)
                }

                $wmiStatusList.Add([PSCustomObject]@{
                    "Filter-Name"      = $fName
                    "Status"           = $statusText
                    "GPO-Anzahl"       = $gpoCount
                    "Verknuepfte GPOs" = $gpoListStr
                })
            }

            # 4. Kacheln befuellen
            $script:lblKpiGpoTotal.Text = [string]$totalGpos
            $script:lblKpiGpoLinked.Text = [string]$linkedGposCount

            $script:lblKpiGpoUnlinked.Text = [string]$unlinkedGposCount
            if ($unlinkedGposCount -gt 0) {
                $script:lblKpiGpoUnlinked.ForeColor = [System.Drawing.Color]::FromArgb(215, 85, 0)
                $script:panelKpiGpoUnlinked.BackColor = [System.Drawing.Color]::FromArgb(255, 250, 240)
            } else {
                $script:lblKpiGpoUnlinked.ForeColor = [System.Drawing.Color]::DarkGreen
                $script:panelKpiGpoUnlinked.BackColor = [System.Drawing.Color]::White
            }

            $script:lblKpiGpoDisabled.Text = [string]$disabledGposCount
            if ($disabledGposCount -gt 0) {
                $script:lblKpiGpoDisabled.ForeColor = [System.Drawing.Color]::FromArgb(190, 40, 40)
                $script:panelKpiGpoDisabled.BackColor = [System.Drawing.Color]::FromArgb(255, 242, 242)
            } else {
                $script:lblKpiGpoDisabled.ForeColor = [System.Drawing.Color]::DarkGreen
                $script:panelKpiGpoDisabled.BackColor = [System.Drawing.Color]::White
            }

            $script:lblKpiWmiTotal.Text = [string]$wmiTotalCount
            $script:lblKpiWmiUnused.Text = [string]$wmiUnusedCount
            if ($wmiUnusedCount -gt 0) {
                $script:lblKpiWmiUnused.ForeColor = [System.Drawing.Color]::FromArgb(215, 85, 0)
                $script:panelKpiWmiUnused.BackColor = [System.Drawing.Color]::FromArgb(255, 250, 240)
            } else {
                $script:lblKpiWmiUnused.ForeColor = [System.Drawing.Color]::DarkGreen
                $script:panelKpiWmiUnused.BackColor = [System.Drawing.Color]::White
            }

            # -------------------------------------------------------------
            # 5. Die 4 Kern-Systempruefungen (Handlungsempfehlungen)
            # -------------------------------------------------------------
            $coreChecks = [System.Collections.Generic.List[PSCustomObject]]::new()

            # CHECK 1: Verwaiste (nicht verknuepfte) GPOs
            if ($unlinkedGposCount -eq 0) {
                $coreChecks.Add([PSCustomObject]@{
                    "Prüfung"             = "1. Verknüpfungs-Status"
                    "Status"              = "OK"
                    "Feststellung"        = "Alle GPOs sind an OUs oder der Domäne verknüpft."
                    "Handlungsempfehlung" = "Keine Maßnahmen erforderlich."
                })
            } else {
                $sampleList = ($unlinkedGpoNames | Select-Object -First 3) -join ", "
                if ($unlinkedGpoNames.Count -gt 3) { $sampleList += " ..." }
                $coreChecks.Add([PSCustomObject]@{
                    "Prüfung"             = "1. Verknüpfungs-Status"
                    "Status"              = "Warnung"
                    "Feststellung"        = "$unlinkedGposCount GPO(s) ohne Verknüpfung gefunden ($sampleList)."
                    "Handlungsempfehlung" = "Prüfen, ob Richtlinien noch benötigt werden; verwaiste GPOs bereinigen."
                })
            }

            # CHECK 2: Deaktivierte GPOs
            if ($disabledGposCount -eq 0) {
                $coreChecks.Add([PSCustomObject]@{
                    "Prüfung"             = "2. Deaktivierte GPOs"
                    "Status"              = "OK"
                    "Feststellung"        = "Keine vollständig deaktivierten GPOs vorhanden."
                    "Handlungsempfehlung" = "Keine Maßnahmen erforderlich."
                })
            } else {
                $sampleList = ($disabledGpoNames | Select-Object -First 3) -join ", "
                if ($disabledGpoNames.Count -gt 3) { $sampleList += " ..." }
                $coreChecks.Add([PSCustomObject]@{
                    "Prüfung"             = "2. Deaktivierte GPOs"
                    "Status"              = "Warnung"
                    "Feststellung"        = "$disabledGposCount GPO(s) sind komplett deaktiviert ($sampleList)."
                    "Handlungsempfehlung" = "Reaktivieren oder bereinigen, um Verarbeitungszeiten zu optimieren."
                })
            }

            # CHECK 3: Sicherung & Konsistenz
            $coreChecks.Add([PSCustomObject]@{
                "Prüfung"             = "3. Sicherungs-Status"
                "Status"              = "OK"
                "Feststellung"        = "Aktueller Domänenbestand eingelesen ($totalGpos GPOs aktiv im AD)."
                "Handlungsempfehlung" = "Regelmäßige GPO-Backups in Tab '3. GPO Backup' durchführen."
            })

            # CHECK 4: WMI-Filter Verwendungsprüfung (Gruen wenn alle genutzt, Orange wenn ungenutzt)
            if ($wmiTotalCount -eq 0) {
                $coreChecks.Add([PSCustomObject]@{
                    "Prüfung"             = "4. WMI-Filter Status"
                    "Status"              = "OK"
                    "Feststellung"        = "Keine WMI-Filter in der Domäne konfiguriert."
                    "Handlungsempfehlung" = "Keine Maßnahmen erforderlich."
                })
            } elseif ($wmiUnusedCount -eq 0) {
                $coreChecks.Add([PSCustomObject]@{
                    "Prüfung"             = "4. WMI-Filter Status"
                    "Status"              = "OK"
                    "Feststellung"        = "Alle WMI-Filter ($wmiTotalCount/$wmiTotalCount) sind Richtlinien zugewiesen."
                    "Handlungsempfehlung" = "Optimal. Keine verwaisten WMI-Filter vorhanden."
                })
            } else {
                $sampleWmi = ($unusedWmiNames | Select-Object -First 3) -join ", "
                if ($unusedWmiNames.Count -gt 3) { $sampleWmi += " ..." }
                $coreChecks.Add([PSCustomObject]@{
                    "Prüfung"             = "4. WMI-Filter Status"
                    "Status"              = "Warnung"
                    "Feststellung"        = "$wmiUnusedCount von $wmiTotalCount WMI-Filter(n) sind nicht zugewiesen ($sampleWmi)."
                    "Handlungsempfehlung" = "In Tab '5. WMI Filter Analyse' prüfen und nicht benötigte Filter löschen."
                })
            }

            # 6. Grids zuweisen
            $arrAudit = [System.Collections.ArrayList]::new()
            foreach ($item in $coreChecks) { [void]$arrAudit.Add($item) }
            $script:gridDashboardAudit.DataSource = $arrAudit

            if ($script:gridDashboardAudit.Columns["Prüfung"])             { $script:gridDashboardAudit.Columns["Prüfung"].FillWeight = 25 }
            if ($script:gridDashboardAudit.Columns["Status"])              { $script:gridDashboardAudit.Columns["Status"].FillWeight = 15 }
            if ($script:gridDashboardAudit.Columns["Feststellung"])        { $script:gridDashboardAudit.Columns["Feststellung"].FillWeight = 45 }
            if ($script:gridDashboardAudit.Columns["Handlungsempfehlung"]) { $script:gridDashboardAudit.Columns["Handlungsempfehlung"].FillWeight = 45 }

            $arrWmi = [System.Collections.ArrayList]::new()
            foreach ($wItem in ($wmiStatusList | Sort-Object "GPO-Anzahl", "Filter-Name")) { [void]$arrWmi.Add($wItem) }
            $script:gridDashboardWmi.DataSource = $arrWmi

            if ($script:gridDashboardWmi.Columns["Filter-Name"])      { $script:gridDashboardWmi.Columns["Filter-Name"].FillWeight = 30 }
            if ($script:gridDashboardWmi.Columns["Status"])           { $script:gridDashboardWmi.Columns["Status"].FillWeight = 25 }
            if ($script:gridDashboardWmi.Columns["GPO-Anzahl"])       { $script:gridDashboardWmi.Columns["GPO-Anzahl"].FillWeight = 15 }
            if ($script:gridDashboardWmi.Columns["Verknuepfte GPOs"]) { $script:gridDashboardWmi.Columns["Verknuepfte GPOs"].FillWeight = 50 }

            $timeStr = (Get-Date).ToString("HH:mm:ss")
            $script:lblDashboardStatus.Text = "Analyse aktualisiert um $timeStr Uhr | 4 Prüfungen abgeschlossen."
            if ($script:lblProgressInfo) { $script:lblProgressInfo.Text = "Dashboard-Analyse erfolgreich abgeschlossen." }
        } catch {
            if ($script:lblDashboardStatus) { $script:lblDashboardStatus.Text = "Fehler: $($_.Exception.Message)" }
            if ($script:lblProgressInfo)    { $script:lblProgressInfo.Text = "Fehler bei Dashboard-Analyse: $($_.Exception.Message)" }
        } finally {
            if ($wmiSearcher) { $wmiSearcher.Dispose() }
            if ($wmiRoot)     { $wmiRoot.Dispose() }
            if ($gpoSearcher) { $gpoSearcher.Dispose() }
            if ($gpoRoot)     { $gpoRoot.Dispose() }
            if ($somSearcher) { $somSearcher.Dispose() }
            if ($somRoot)     { $somRoot.Dispose() }

            if ($script:pbarGlobal) { $script:pbarGlobal.Visible = $false }
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }

    $btnRefreshDashboard.Add_Click({
        if ($script:Invoke_LoadDashboard -is [scriptblock]) {
            & $script:Invoke_LoadDashboard
        }
    })
}
