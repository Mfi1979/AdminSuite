<#
.SYNOPSIS
    Modul 14: GPO Analyse & CSV Export Tool (Standalone & Suite-kompatibel)
.DESCRIPTION
    Analysiert alle Gruppenrichtlinien (GPOs) via LDAP/ADSI.
    - Spaltenstruktur:
      * Verlinkt: "Ja" / "Nein"
      * Link-Anzahl: [int] (0 wenn Nein)
      * Benutzer-Konfig: "Aktiviert" / "Deaktiviert"
      * Computer-Konfig: "Aktiviert" / "Deaktiviert"
      * Gesamt-Status: "Aktiviert", "Benutzer deaktiviert", "Computer deaktiviert", "Vollständig deaktiviert"
    - Master-Detail-Ansicht: Klick auf GPO links zeigt rechts alle Verlinkungsziele
    - CSV-Export und Spaltensortierung
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.DirectoryServices

function Show-GPOReportModule {
    [CmdletBinding()]
    param()

    # --- Prüfen auf Domänenzugehörigkeit ---
    try {
        $domainInfo = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
        $domainDN = ([ADSI]"LDAP://RootDSE").defaultNamingContext.Value
        $domainName = $domainInfo.Name
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Dieser Computer ist nicht mit einer Active Directory Domäne verbunden oder der Domain Controller ist nicht erreichbar.",
            "Keine Domänenverbindung",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    # --- GUI Fenster Definition ---
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Modul 14: GPO Report & Verlinkungs-Analyse ($domainName)"
    $form.Size = New-Object System.Drawing.Size(1350, 800)
    $form.MinimumSize = New-Object System.Drawing.Size(1050, 600)
    $form.StartPosition = "CenterScreen"
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    # --- TOP PANEL: Status & Steuerungs-Buttons ---
    $topPanel = New-Object System.Windows.Forms.Panel
    $topPanel.Dock = [System.Windows.Forms.DockStyle]::Top
    $topPanel.Height = 70
    $topPanel.BackColor = [System.Drawing.Color]::FromArgb(245, 248, 252)
    $form.Controls.Add($topPanel)

    $lblSummary = New-Object System.Windows.Forms.Label
    $lblSummary.Text = "Status: Lade Daten..."
    $lblSummary.Location = New-Object System.Drawing.Point(15, 12)
    $lblSummary.AutoSize = $true
    $lblSummary.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $topPanel.Controls.Add($lblSummary)

    $lblSubText = New-Object System.Windows.Forms.Label
    $lblSubText.Text = "GPOs werden analysiert (Links & WMI-Filter)..."
    $lblSubText.Location = New-Object System.Drawing.Point(15, 36)
    $lblSubText.AutoSize = $true
    $lblSubText.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
    $topPanel.Controls.Add($lblSubText)

    # Button: CSV Export
    $btnExportTab = New-Object System.Windows.Forms.Button
    $btnExportTab.Text = "CSV Export (Aktiver Tab)"
    $btnExportTab.Size = New-Object System.Drawing.Size(175, 34)
    $btnExportTab.Location = New-Object System.Drawing.Point(1030, 16)
    $btnExportTab.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
    $btnExportTab.BackColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
    $btnExportTab.ForeColor = [System.Drawing.Color]::White
    $btnExportTab.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $topPanel.Controls.Add($btnExportTab)

    # Button: Aktualisieren
    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = "Neu laden"
    $btnRefresh.Size = New-Object System.Drawing.Size(100, 34)
    $btnRefresh.Location = New-Object System.Drawing.Point(1215, 16)
    $btnRefresh.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
    $topPanel.Controls.Add($btnRefresh)

    # TabControl als Haupt-Container
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Dock = [System.Windows.Forms.DockStyle]::Fill
    $form.Controls.Add($tabControl)
    $topPanel.SendToBack()

    # Cache für Verlinkungsziele
    $script:GpoLinksCache = @{}

    # --- Hilfsfunktion zum Erstellen eines Master-Detail SplitContainers ---
    function New-GpoSplitContainer {
        param(
            [string]$LeftTitle,
            [string]$RightTitle
        )

        $split = New-Object System.Windows.Forms.SplitContainer
        $split.Dock = [System.Windows.Forms.DockStyle]::Fill
        $split.Orientation = [System.Windows.Forms.Orientation]::Vertical
        $split.SplitterDistance = 860

        # Linke Box & Grid
        $grpLeft = New-Object System.Windows.Forms.GroupBox
        $grpLeft.Text = $LeftTitle
        $grpLeft.Dock = [System.Windows.Forms.DockStyle]::Fill
        $grpLeft.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
        $split.Panel1.Controls.Add($grpLeft)

        $gridMaster = New-Object System.Windows.Forms.DataGridView
        $gridMaster.Dock = [System.Windows.Forms.DockStyle]::Fill
        $gridMaster.ReadOnly = $true
        $gridMaster.AllowUserToAddRows = $false
        $gridMaster.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
        $gridMaster.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
        $gridMaster.BackgroundColor = [System.Drawing.Color]::White
        $gridMaster.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Regular)
        $grpLeft.Controls.Add($gridMaster)

        # Rechte Box & Grid (Verlinkungsziele)
        $grpRight = New-Object System.Windows.Forms.GroupBox
        $grpRight.Text = $RightTitle
        $grpRight.Dock = [System.Windows.Forms.DockStyle]::Fill
        $grpRight.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
        $split.Panel2.Controls.Add($grpRight)

        $gridDetails = New-Object System.Windows.Forms.DataGridView
        $gridDetails.Dock = [System.Windows.Forms.DockStyle]::Fill
        $gridDetails.ReadOnly = $true
        $gridDetails.AllowUserToAddRows = $false
        $gridDetails.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
        $gridDetails.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
        $gridDetails.BackgroundColor = [System.Drawing.Color]::White
        $gridDetails.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Regular)
        $grpRight.Controls.Add($gridDetails)

        return @{
            SplitContainer = $split
            GridMaster     = $gridMaster
            GridDetails    = $gridDetails
            GrpRight       = $grpRight
        }
    }

    # --- Tab 1: Alle GPOs (Master-Detail) ---
    $tabAll = New-Object System.Windows.Forms.TabPage
    $tabAll.Text = "Alle GPOs (0)"
    $tabAll.Padding = New-Object System.Windows.Forms.Padding(4)
    $controlsAll = New-GpoSplitContainer -LeftTitle "Gruppenrichtlinien (GPOs)" -RightTitle "Verlinkungsziele (OUs & Domäne)"
    $tabAll.Controls.Add($controlsAll.SplitContainer)
    $tabControl.TabPages.Add($tabAll)

    # --- Tab 2: GPOs mit WMI-Filter (Master-Detail) ---
    $tabWmi = New-Object System.Windows.Forms.TabPage
    $tabWmi.Text = "GPOs mit WMI-Filter (0)"
    $tabWmi.Padding = New-Object System.Windows.Forms.Padding(4)
    $controlsWmi = New-GpoSplitContainer -LeftTitle "WMI-gefilterte GPOs" -RightTitle "Verlinkungsziele der gefilterten GPO"
    $tabWmi.Controls.Add($controlsWmi.SplitContainer)
    $tabControl.TabPages.Add($tabWmi)

    # --- Tab 3: Nicht verlinkte GPOs ---
    $tabUnlinked = New-Object System.Windows.Forms.TabPage
    $tabUnlinked.Text = "Nicht verlinkte GPOs (0)"
    $tabUnlinked.Padding = New-Object System.Windows.Forms.Padding(4)

    $gridUnlinked = New-Object System.Windows.Forms.DataGridView
    $gridUnlinked.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridUnlinked.ReadOnly = $true
    $gridUnlinked.AllowUserToAddRows = $false
    $gridUnlinked.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
    $gridUnlinked.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $gridUnlinked.BackgroundColor = [System.Drawing.Color]::White
    $tabUnlinked.Controls.Add($gridUnlinked)
    $tabControl.TabPages.Add($tabUnlinked)

    # --- Details-Tabelle bei Zeilenauswahl befüllen ---
    function Update-LinkDetailGrid {
        param(
            [System.Windows.Forms.DataGridView]$SourceGrid,
            [System.Windows.Forms.DataGridView]$TargetGrid,
            [System.Windows.Forms.GroupBox]$TargetBox
        )

        if ($SourceGrid.SelectedRows.Count -eq 0) {
            $TargetGrid.DataSource = $null
            $TargetBox.Text = "Verlinkungsziele: Keine Auswahl"
            return
        }

        $selectedRow = $SourceGrid.SelectedRows[0]
        $guid = $selectedRow.Cells["GUID"].Value
        $gpoName = $selectedRow.Cells["GPO Name"].Value

        if (-not $guid -or -not $script:GpoLinksCache.ContainsKey($guid) -or $script:GpoLinksCache[$guid].Count -eq 0) {
            $arrEmpty = [System.Collections.ArrayList]::new()
            [void]$arrEmpty.Add([PSCustomObject]@{
                "Typ"                = "Info"
                "Name / OU"          = "-"
                "DistinguishedName"  = "⚠️ Diese Richtlinie ist an kein AD-Objekt verlinkt."
            })
            $TargetGrid.DataSource = $arrEmpty
            $TargetBox.Text = "Verlinkungsziele: $gpoName (0 Links)"
            return
        }

        $linkTargets = $script:GpoLinksCache[$guid]
        $arrDetails = [System.Collections.ArrayList]::new()

        foreach ($dn in $linkTargets) {
            $type = "Organizational Unit (OU)"
            $simpleName = $dn

            if ($dn -match "^OU=([^,]+)") {
                $simpleName = $matches[1]
                $type = "OU"
            } elseif ($dn -match "^DC=") {
                $type = "Domänen-Root (DomainDNS)"
                $simpleName = $domainName
            }

            [void]$arrDetails.Add([PSCustomObject]@{
                "Typ"                = $type
                "Name / OU"          = $simpleName
                "DistinguishedName"  = $dn
            })
        }

        $TargetGrid.DataSource = $arrDetails
        $TargetBox.Text = "Verlinkungsziele: $gpoName ($($arrDetails.Count) Link(s))"
    }

    $controlsAll.GridMaster.Add_SelectionChanged({
        Update-LinkDetailGrid -SourceGrid $controlsAll.GridMaster -TargetGrid $controlsAll.GridDetails -TargetBox $controlsAll.GrpRight
    })

    $controlsWmi.GridMaster.Add_SelectionChanged({
        Update-LinkDetailGrid -SourceGrid $controlsWmi.GridMaster -TargetGrid $controlsWmi.GridDetails -TargetBox $controlsWmi.GrpRight
    })

    # --- Spaltensortierung per Mausklick ---
    function Enable-GridSorting {
        param([System.Windows.Forms.DataGridView]$Grid)

        $Grid.Add_ColumnHeaderMouseClick({
            param($sender, $e)
            $targetGrid = $sender
            $colProp = $targetGrid.Columns[$e.ColumnIndex].DataPropertyName
            if (-not $colProp) { $colProp = $targetGrid.Columns[$e.ColumnIndex].HeaderText }

            if ($targetGrid.Tag -and $targetGrid.Tag.Column -eq $colProp) {
                $asc = -not $targetGrid.Tag.Ascending
            } else {
                $asc = $true
            }
            $targetGrid.Tag = @{ Column = $colProp; Ascending = $asc }

            $data = @($targetGrid.DataSource)
            if ($null -eq $data -or $data.Count -le 1) { return }

            $sorted = $data | Sort-Object -Property @{
                Expression = {
                    $val = $_.$colProp
                    if ($null -eq $val) { return "" }
                    if ($colProp -eq "Link-Anzahl" -and ($val -as [int])) { return [int]$val }
                    return $val
                }
                Descending = (-not $asc)
            }

            $arr = [System.Collections.ArrayList]::new()
            foreach ($item in $sorted) { [void]$arr.Add($item) }
            $targetGrid.DataSource = $arr

            foreach ($col in $targetGrid.Columns) {
                $col.HeaderCell.SortGlyphDirection = [System.Windows.Forms.SortOrder]::None
            }
            $targetGrid.Columns[$e.ColumnIndex].HeaderCell.SortGlyphDirection = if ($asc) { 
                [System.Windows.Forms.SortOrder]::Ascending 
            } else { 
                [System.Windows.Forms.SortOrder]::Descending 
            }
        })
    }

    Enable-GridSorting -Grid $controlsAll.GridMaster
    Enable-GridSorting -Grid $controlsAll.GridDetails
    Enable-GridSorting -Grid $controlsWmi.GridMaster
    Enable-GridSorting -Grid $controlsWmi.GridDetails
    Enable-GridSorting -Grid $gridUnlinked

    # --- Datenabfrage & Analyse ---
    $loadData = {
        $lblSummary.Text = "Lade GPOs und Links aus Active Directory..."
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $form.Refresh()

        try {
            $script:GpoLinksCache = @{}

            # 1. WMI-Filter auflösen
            $wmiMap = @{}
            $wmiSearcher = [System.DirectoryServices.DirectorySearcher]::new(
                [System.DirectoryServices.DirectoryEntry]::new("LDAP://CN=SOM,CN=WMIPolicy,CN=System,$domainDN")
            )
            $wmiSearcher.Filter = "(objectClass=msWMI-Som)"
            $wmiSearcher.PropertiesToLoad.AddRange(@("msWMI-Name", "msWMI-ID", "msWMI-Parm2"))
            
            try {
                $wmiResults = $wmiSearcher.FindAll()
                foreach ($w in $wmiResults) {
                    $id = $w.Properties["mswmi-id"][0]
                    $wName = $w.Properties["mswmi-name"][0]
                    $wQuery = if ($w.Properties["mswmi-parm2"]) { $w.Properties["mswmi-parm2"][0] } else { "" }
                    $wmiMap[$id] = [PSCustomObject]@{ Name = $wName; Query = $wQuery }
                }
            } catch { }

            # 2. Verlinkungen (gPLink) an OUs & Domänen-Root auslesen
            $linkSearcher = [System.DirectoryServices.DirectorySearcher]::new(
                [System.DirectoryServices.DirectoryEntry]::new("LDAP://$domainDN")
            )
            $linkSearcher.Filter = "(|(objectClass=organizationalUnit)(objectClass=domainDNS))"
            $linkSearcher.PropertiesToLoad.AddRange(@("distinguishedName", "gPLink", "name"))
            $linkSearcher.SearchScope = [System.DirectoryServices.SearchScope]::Subtree
            $ouResults = $linkSearcher.FindAll()

            foreach ($ou in $ouResults) {
                if ($ou.Properties["gplink"]) {
                    $rawGpLink = $ou.Properties["gplink"][0]
                    $targetDN = $ou.Properties["distinguishedname"][0]
                    
                    $regexMatches = [regex]::Matches($rawGpLink, "cn=({[a-fA-F0-9-]+})", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                    foreach ($match in $regexMatches) {
                        $gpoGuid = $match.Groups[1].Value.ToUpper()
                        if (-not $script:GpoLinksCache.ContainsKey($gpoGuid)) {
                            $script:GpoLinksCache[$gpoGuid] = [System.Collections.Generic.List[string]]::new()
                        }
                        $script:GpoLinksCache[$gpoGuid].Add($targetDN)
                    }
                }
            }

            # 3. GPO Container aus CN=Policies,CN=System abfragen
            $gpoSearcher = [System.DirectoryServices.DirectorySearcher]::new(
                [System.DirectoryServices.DirectoryEntry]::new("LDAP://CN=Policies,CN=System,$domainDN")
            )
            $gpoSearcher.Filter = "(objectClass=groupPolicyContainer)"
            $gpoSearcher.PropertiesToLoad.AddRange(@("displayName", "name", "flags", "gPCWQLFilter", "whenCreated", "whenChanged"))
            $gpoResults = $gpoSearcher.FindAll()

            $listAll = [System.Collections.ArrayList]::new()
            $listUnlinked = [System.Collections.ArrayList]::new()
            $listWmi = [System.Collections.ArrayList]::new()

            foreach ($g in $gpoResults) {
                $guid = $g.Properties["name"][0].ToUpper()
                $displayName = if ($g.Properties["displayname"]) { $g.Properties["displayname"][0] } else { $guid }
                $flags = if ($g.Properties["flags"]) { [int]$g.Properties["flags"][0] } else { 0 }
                
                # flags: 0 = Alle aktiv, 1 = User disabled, 2 = Computer disabled, 3 = Alle disabled
                $userStatus = if (($flags -band 1) -eq 1) { "Deaktiviert" } else { "Aktiviert" }
                $compStatus = if (($flags -band 2) -eq 2) { "Deaktiviert" } else { "Aktiviert" }

                $overallStatus = switch ($flags) {
                    0 { "Aktiviert" }
                    1 { "Benutzer deaktiviert" }
                    2 { "Computer deaktiviert" }
                    3 { "Vollständig deaktiviert" }
                    default { "Unbekannt ($flags)" }
                }

                # Verlinkung Ja/Nein + Anzahl
                $isLinked = $false
                $linkedCount = 0

                if ($script:GpoLinksCache.ContainsKey($guid) -and $script:GpoLinksCache[$guid].Count -gt 0) {
                    $isLinked = $true
                    $linkedCount = $script:GpoLinksCache[$guid].Count
                }

                # WMI Filter
                $hasWmi = $false
                $wmiFilterName = "-"
                $wmiFilterQuery = "-"

                if ($g.Properties["gpcwqlfilter"]) {
                    $rawWmi = $g.Properties["gpcwqlfilter"][0]
                    $hasWmi = $true
                    
                    if ($rawWmi -match "({[a-fA-F0-9-]+})") {
                        $wmiGuid = $matches[1]
                        if ($wmiMap.ContainsKey($wmiGuid)) {
                            $wmiFilterName = $wmiMap[$wmiGuid].Name
                            $wmiFilterQuery = $wmiMap[$wmiGuid].Query
                        } else {
                            $wmiFilterName = $wmiGuid
                        }
                    } else {
                        $wmiFilterName = $rawWmi
                    }
                }

                $created = if ($g.Properties["whencreated"]) { $g.Properties["whencreated"][0] } else { "-" }
                $changed = if ($g.Properties["whenchanged"]) { $g.Properties["whenchanged"][0] } else { "-" }

                $item = [PSCustomObject]@{
                    "GPO Name"         = $displayName
                    "Verlinkt"         = if ($isLinked) { "Ja" } else { "Nein" }
                    "Link-Anzahl"      = $linkedCount
                    "Benutzer-Konfig"  = $userStatus
                    "Computer-Konfig"  = $compStatus
                    "Gesamt-Status"    = $overallStatus
                    "WMI-Filter"       = $wmiFilterName
                    "WMI Query"        = $wmiFilterQuery
                    "GUID"             = $guid
                    "Erstellt am"      = $created
                    "Geändert am"      = $changed
                }

                [void]$listAll.Add($item)

                if (-not $isLinked) {
                    [void]$listUnlinked.Add($item)
                }

                if ($hasWmi) {
                    [void]$listWmi.Add($item)
                }
            }

            # Grids binden
            $controlsAll.GridMaster.DataSource = $listAll
            $controlsWmi.GridMaster.DataSource = $listWmi
            $gridUnlinked.DataSource = $listUnlinked

            # Registerkarten-Titel aktualisieren
            $tabAll.Text = "Alle GPOs ($($listAll.Count))"
            $tabWmi.Text = "GPOs mit WMI-Filter ($($listWmi.Count))"
            $tabUnlinked.Text = if ($listUnlinked.Count -gt 0) { "⚠️ Nicht verlinkte GPOs ($($listUnlinked.Count))" } else { "Nicht verlinkte GPOs (0)" }

            $lblSummary.Text = "Gesamt: $($listAll.Count) GPOs | Nicht verlinkt: $($listUnlinked.Count) | Mit WMI-Filter: $($listWmi.Count)"
            $lblSubText.Text = "Domäne: $domainName | Klicken Sie links eine GPO an, um rechts alle Verknüpfungs-OUs zu sehen."

            Update-LinkDetailGrid -SourceGrid $controlsAll.GridMaster -TargetGrid $controlsAll.GridDetails -TargetBox $controlsAll.GrpRight

        } catch {
            $lblSummary.Text = "Fehler beim Laden der GPO-Informationen"
            $lblSubText.Text = $_.Exception.Message
            [System.Windows.Forms.MessageBox]::Show(
                "Fehler beim Abfragen des Active Directory:`r`n$($_.Exception.Message)",
                "Abfragefehler",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }

    # --- CSV Export Event ---
    $btnExportTab.Add_Click({
        $activeGrid = switch ($tabControl.SelectedIndex) {
            0 { $controlsAll.GridMaster }
            1 { $controlsWmi.GridMaster }
            2 { $gridUnlinked }
            default { $controlsAll.GridMaster }
        }

        $tabNameClean = switch ($tabControl.SelectedIndex) {
            0 { "All_GPOs" }
            1 { "WMI_Filtered_GPOs" }
            2 { "Unlinked_GPOs" }
            default { "GPO_Report" }
        }

        $exportData = @($activeGrid.DataSource)
        if ($null -eq $exportData -or $exportData.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Keine Daten in der aktuellen Ansicht zum Exportieren vorhanden.", "Export", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }

        # Für CSV-Export: Verlinkungsziele als lesbare Textspalte anfügen
        $enrichedExport = foreach ($row in $exportData) {
            $gGuid = $row.GUID
            $linksStr = if ($script:GpoLinksCache.ContainsKey($gGuid)) { ($script:GpoLinksCache[$gGuid] -join " | ") } else { "Keine (Nicht verlinkt)" }
            
            [PSCustomObject]@{
                "GPO Name"         = $row."GPO Name"
                "Verlinkt"         = $row."Verlinkt"
                "Link-Anzahl"      = $row."Link-Anzahl"
                "Benutzer-Konfig"  = $row."Benutzer-Konfig"
                "Computer-Konfig"  = $row."Computer-Konfig"
                "Gesamt-Status"    = $row."Gesamt-Status"
                "Verlinkungsziele" = $linksStr
                "WMI-Filter"       = $row."WMI-Filter"
                "WMI Query"        = $row."WMI Query"
                "GUID"             = $row."GUID"
                "Erstellt am"      = $row."Erstellt am"
                "Geändert am"      = $row."Geändert am"
            }
        }

        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "CSV-Datei (*.csv)|*.csv"
        $sfd.FileName = "GPO_Report_${tabNameClean}_$($domainName)_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"

        if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                $enrichedExport | Export-Csv -Path $sfd.FileName -NoTypeInformation -Encoding UTF8 -Delimiter ";"
                [System.Windows.Forms.MessageBox]::Show("Erfolgreich $($enrichedExport.Count) Einträge exportiert nach:`r`n$($sfd.FileName)", "Export Abgeschlossen", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Fehler beim Speichern der CSV-Datei:`r`n$($_.Exception.Message)", "Exportfehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    })

    $btnRefresh.Add_Click({ & $loadData })
    $form.Add_Shown({ & $loadData })

    [void]$form.ShowDialog()
}

# --- Standalone-Aufruf ---
if ($MyInvocation.InvocationName -ne '.' -and $MyInvocation.Line -notmatch 'Open-Tool') {
    Show-GPOReportModule
}