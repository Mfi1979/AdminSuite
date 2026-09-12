<#
================================================================================
 TOOL 21: LAPS VERSION & STATUS AUDIT (COMPUTERS & SERVERS)
 Version: v1.3.3 (Direktabfrage am Schema-Master zur Umgehung von Replikations-Lags)
================================================================================
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.DirectoryServices

if (-not [System.Windows.Forms.Application]::RenderWithVisualStyles) {
    try { [System.Windows.Forms.Application]::EnableVisualStyles() } catch {}
}

function Open-ToolLapsInventoryAudit {
    $toolVersion = "v1.3.3"
    $isClosing = $false

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Tool 21: Domaenenweiter LAPS-Versionen & Status Audit ($toolVersion)"
    $form.Size = New-Object System.Drawing.Size(1420, 860)
    $form.StartPosition = "CenterScreen"
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

    try {
        $pnlHeader = New-Object System.Windows.Forms.Panel
        $pnlHeader.Dock = [System.Windows.Forms.DockStyle]::Top
        $pnlHeader.Height = 145
        $pnlHeader.BackColor = [System.Drawing.Color]::FromArgb(238, 242, 246)
        $form.Controls.Add($pnlHeader)

        # ZEILE 1: AD-Schema Aktivierung
        $pnlAdStatus = New-Object System.Windows.Forms.Panel
        $pnlAdStatus.Location = New-Object System.Drawing.Point(15, 10)
        $pnlAdStatus.Size = New-Object System.Drawing.Size(1370, 36)
        $pnlAdStatus.BackColor = [System.Drawing.Color]::FromArgb(255, 255, 255)
        $pnlAdStatus.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
        $pnlHeader.Controls.Add($pnlAdStatus)

        $lblSchemaTitle = New-Object System.Windows.Forms.Label
        $lblSchemaTitle.Text = "AD-Schema Aktivierung:"
        $lblSchemaTitle.Location = New-Object System.Drawing.Point(10, 8)
        $lblSchemaTitle.AutoSize = $true
        $lblSchemaTitle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $pnlAdStatus.Controls.Add($lblSchemaTitle)

        $lblSchemaLegacy = New-Object System.Windows.Forms.Label
        $lblSchemaLegacy.Text = "Legacy LAPS Schema: Pruefe..."
        $lblSchemaLegacy.Location = New-Object System.Drawing.Point(190, 8)
        $lblSchemaLegacy.AutoSize = $true
        $lblSchemaLegacy.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $pnlAdStatus.Controls.Add($lblSchemaLegacy)

        $lblSchemaModern = New-Object System.Windows.Forms.Label
        $lblSchemaModern.Text = "Modern Windows LAPS Schema: Pruefe..."
        $lblSchemaModern.Location = New-Object System.Drawing.Point(490, 8)
        $lblSchemaModern.AutoSize = $true
        $lblSchemaModern.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $pnlAdStatus.Controls.Add($lblSchemaModern)

        $lblActiveSummary = New-Object System.Windows.Forms.Label
        $lblActiveSummary.Text = "AD-Status: Pruefe..."
        $lblActiveSummary.Location = New-Object System.Drawing.Point(870, 8)
        $lblActiveSummary.AutoSize = $true
        $lblActiveSummary.ForeColor = [System.Drawing.Color]::FromArgb(0, 70, 150)
        $lblActiveSummary.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
        $pnlAdStatus.Controls.Add($lblActiveSummary)

        # ZEILE 2: Gefundene Geräte Dashboard
        $pnlDashSub = New-Object System.Windows.Forms.Panel
        $pnlDashSub.Location = New-Object System.Drawing.Point(15, 52)
        $pnlDashSub.Size = New-Object System.Drawing.Size(1370, 36)
        $pnlDashSub.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 252)
        $pnlDashSub.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
        $pnlHeader.Controls.Add($pnlDashSub)

        $lblDashTitle = New-Object System.Windows.Forms.Label
        $lblDashTitle.Text = "Gefundene Geraete:"
        $lblDashTitle.Location = New-Object System.Drawing.Point(10, 8)
        $lblDashTitle.AutoSize = $true
        $lblDashTitle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $pnlDashSub.Controls.Add($lblDashTitle)

        $lblDashTotal = New-Object System.Windows.Forms.Label
        $lblDashTotal.Text = "Gesamt: 0"
        $lblDashTotal.Location = New-Object System.Drawing.Point(150, 8)
        $lblDashTotal.AutoSize = $true
        $pnlDashSub.Controls.Add($lblDashTotal)

        $lblDashLegacy = New-Object System.Windows.Forms.Label
        $lblDashLegacy.Text = "[!] Mit Legacy LAPS: 0"
        $lblDashLegacy.Location = New-Object System.Drawing.Point(300, 8)
        $lblDashLegacy.AutoSize = $true
        $lblDashLegacy.ForeColor = [System.Drawing.Color]::FromArgb(170, 100, 0)
        $lblDashLegacy.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $pnlDashSub.Controls.Add($lblDashLegacy)

        $lblDashModern = New-Object System.Windows.Forms.Label
        $lblDashModern.Text = "[+] Mit Modern LAPS: 0"
        $lblDashModern.Location = New-Object System.Drawing.Point(520, 8)
        $lblDashModern.AutoSize = $true
        $lblDashModern.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 0)
        $lblDashModern.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $pnlDashSub.Controls.Add($lblDashModern)

        $lblDashNone = New-Object System.Windows.Forms.Label
        $lblDashNone.Text = "[-] Ohne LAPS: 0"
        $lblDashNone.Location = New-Object System.Drawing.Point(740, 8)
        $lblDashNone.AutoSize = $true
        $lblDashNone.ForeColor = [System.Drawing.Color]::FromArgb(180, 0, 0)
        $lblDashNone.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $pnlDashSub.Controls.Add($lblDashNone)

        # ZEILE 3: Filter & Buttons
        $lblFilter = New-Object System.Windows.Forms.Label
        $lblFilter.Text = "Suche:"
        $lblFilter.Location = New-Object System.Drawing.Point(15, 104)
        $lblFilter.AutoSize = $true
        $lblFilter.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $pnlHeader.Controls.Add($lblFilter)

        $txtFilter = New-Object System.Windows.Forms.TextBox
        $txtFilter.Location = New-Object System.Drawing.Point(65, 101)
        $txtFilter.Size = New-Object System.Drawing.Size(120, 23)
        $txtFilter.Text = "*"
        $pnlHeader.Controls.Add($txtFilter)

        $lblObjType = New-Object System.Windows.Forms.Label
        $lblObjType.Text = "Typ:"
        $lblObjType.Location = New-Object System.Drawing.Point(195, 104)
        $lblObjType.AutoSize = $true
        $lblObjType.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $pnlHeader.Controls.Add($lblObjType)

        $cmbObjType = New-Object System.Windows.Forms.ComboBox
        $cmbObjType.Location = New-Object System.Drawing.Point(230, 101)
        $cmbObjType.Size = New-Object System.Drawing.Size(95, 23)
        $cmbObjType.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
        [void]$cmbObjType.Items.Add("Alle")
        [void]$cmbObjType.Items.Add("Computer")
        [void]$cmbObjType.Items.Add("Server")
        $cmbObjType.SelectedIndex = 0
        $pnlHeader.Controls.Add($cmbObjType)

        $lblStatusFilter = New-Object System.Windows.Forms.Label
        $lblStatusFilter.Text = "Status:"
        $lblStatusFilter.Location = New-Object System.Drawing.Point(335, 104)
        $lblStatusFilter.AutoSize = $true
        $lblStatusFilter.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $pnlHeader.Controls.Add($lblStatusFilter)

        $cmbStatusFilter = New-Object System.Windows.Forms.ComboBox
        $cmbStatusFilter.Location = New-Object System.Drawing.Point(385, 101)
        $cmbStatusFilter.Size = New-Object System.Drawing.Size(115, 23)
        $cmbStatusFilter.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
        [void]$cmbStatusFilter.Items.Add("Alle Status")
        [void]$cmbStatusFilter.Items.Add("Nur Aktivierte")
        [void]$cmbStatusFilter.Items.Add("Nur Deaktivierte")
        $cmbStatusFilter.SelectedIndex = 0
        $pnlHeader.Controls.Add($cmbStatusFilter)

        $btnRun = New-Object System.Windows.Forms.Button
        $btnRun.Text = "LAPS Scan starten"
        $btnRun.Location = New-Object System.Drawing.Point(510, 99)
        $btnRun.Size = New-Object System.Drawing.Size(125, 28)
        $btnRun.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
        $btnRun.ForeColor = [System.Drawing.Color]::White
        $btnRun.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnRun.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $pnlHeader.Controls.Add($btnRun)

        $btnExport = New-Object System.Windows.Forms.Button
        $btnExport.Text = "CSV Export (;)"
        $btnExport.Location = New-Object System.Drawing.Point(642, 99)
        $btnExport.Size = New-Object System.Drawing.Size(100, 28)
        $btnExport.BackColor = [System.Drawing.Color]::FromArgb(16, 124, 65)
        $btnExport.ForeColor = [System.Drawing.Color]::White
        $btnExport.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnExport.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $pnlHeader.Controls.Add($btnExport)

        $btnCheckSchemaUpdate = New-Object System.Windows.Forms.Button
        $btnCheckSchemaUpdate.Text = "Schema Update Check"
        $btnCheckSchemaUpdate.Location = New-Object System.Drawing.Point(750, 99)
        $btnCheckSchemaUpdate.Size = New-Object System.Drawing.Size(155, 28)
        $btnCheckSchemaUpdate.BackColor = [System.Drawing.Color]::FromArgb(216, 107, 0)
        $btnCheckSchemaUpdate.ForeColor = [System.Drawing.Color]::White
        $btnCheckSchemaUpdate.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnCheckSchemaUpdate.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $pnlHeader.Controls.Add($btnCheckSchemaUpdate)

        $lblStatus = New-Object System.Windows.Forms.Label
        $lblStatus.Location = New-Object System.Drawing.Point(915, 104)
        $lblStatus.AutoSize = $true
        $lblStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(0, 70, 150)
        $lblStatus.Text = "Bereit ($toolVersion)."
        $pnlHeader.Controls.Add($lblStatus)

        # 2. Container Panel fuer Tabelle
        $pnlGridContainer = New-Object System.Windows.Forms.Panel
        $pnlGridContainer.Dock = [System.Windows.Forms.DockStyle]::Fill
        $pnlGridContainer.Padding = New-Object System.Windows.Forms.Padding(15, 10, 15, 15)
        $form.Controls.Add($pnlGridContainer)
        $pnlGridContainer.BringToFront()

        # 3. DataGridView
        $grid = New-Object System.Windows.Forms.DataGridView
        $grid.Dock = [System.Windows.Forms.DockStyle]::Fill
        $grid.ReadOnly = $true
        $grid.RowHeadersVisible = $false
        $grid.ColumnHeadersVisible = $true
        $grid.EnableHeadersVisualStyles = $false
        $grid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
        $grid.ColumnHeadersHeight = 35
        $grid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(215, 228, 242)
        $grid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(20, 40, 70)
        $grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
        $grid.BackgroundColor = [System.Drawing.Color]::White
        $grid.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
        $grid.AutoGenerateColumns = $false
        $pnlGridContainer.Controls.Add($grid)

        $col1 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $col1.HeaderText = "Computer / Server";  $col1.DataPropertyName = "Computer";       $col1.Width = 230; [void]$grid.Columns.Add($col1)
        $col2 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $col2.HeaderText = "Konto-Status";       $col2.DataPropertyName = "KontoStatus";    $col2.Width = 110; [void]$grid.Columns.Add($col2)
        $col3 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $col3.HeaderText = "Betriebssystem";     $col3.DataPropertyName = "Betriebssystem";  $col3.Width = 240; [void]$grid.Columns.Add($col3)
        $col4 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $col4.HeaderText = "LAPS Variante";      $col4.DataPropertyName = "LapsVariante";   $col4.Width = 180; [void]$grid.Columns.Add($col4)
        $col5 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $col5.HeaderText = "Kennwort-Status";    $col5.DataPropertyName = "KennwortStatus"; $col5.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill; [void]$grid.Columns.Add($col5)
        $col6 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $col6.HeaderText = "Legacy Zeitstempel"; $col6.DataPropertyName = "LegacyZeit";     $col6.Width = 140; [void]$grid.Columns.Add($col6)
        $col7 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $col7.HeaderText = "Modern Zeitstempel"; $col7.DataPropertyName = "ModernZeit";     $col7.Width = 140; [void]$grid.Columns.Add($col7)

        $script:inventoryResults = [System.Collections.Generic.List[PSCustomObject]]::new()
        $script:sortCol = ""
        $script:sortAsc = $true

        $grid.Add_ColumnHeaderMouseClick({
            param($sender, $e)
            if ($isClosing -or $form.IsDisposed -or $grid.IsDisposed) { return }
            $colIndex = $e.ColumnIndex
            if ($colIndex -lt 0) { return }

            $col = $grid.Columns[$colIndex]
            $propName = $col.DataPropertyName
            if (-not $propName) { return }

            if ($script:sortCol -eq $propName) {
                $script:sortAsc = -not $script:sortAsc
            } else {
                $script:sortCol = $propName
                $script:sortAsc = $true
            }

            $currentData = @($grid.DataSource)
            if ($null -eq $currentData -or $currentData.Count -le 1) { return }

            $sorted = $currentData | Sort-Object -Property @{
                Expression = {
                    $val = $_.$propName
                    if ($null -eq $val) { return "" }
                    return $val
                }
                Descending = (-not $script:sortAsc)
            }

            $arr = [System.Collections.ArrayList]::new()
            foreach ($item in $sorted) { [void]$arr.Add($item) }

            $grid.DataSource = $null
            $grid.DataSource = $arr

            foreach ($c in $grid.Columns) {
                $c.HeaderCell.SortGlyphDirection = [System.Windows.Forms.SortOrder]::None
            }
            $grid.Columns[$colIndex].HeaderCell.SortGlyphDirection = $(if ($script:sortAsc) { [System.Windows.Forms.SortOrder]::Ascending } else { [System.Windows.Forms.SortOrder]::Descending })
        })

        # Scan-Logik
        $runInventoryScan = {
            if ($isClosing -or $form.IsDisposed) { return }
            $btnRun.Enabled = $false
            $btnCheckSchemaUpdate.Enabled = $false
            $lblStatus.Text = "Pruefe AD-Schema und Geraete ($toolVersion)..."
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            $script:inventoryResults.Clear()
            $form.Refresh()

            $rootDSE = $null; $schemaEntry = $null; $searcherSchema = $null; $searcher = $null; $computers = $null

            try {
                $rootDSE = [ADSI]"LDAP://RootDSE"
                $schemaDN = $rootDSE.schemaNamingContext.ToString()
                $defNC = $rootDSE.defaultNamingContext.ToString()

                # Schema-Master ermitteln für zuverlässige Schema-Prüfung
                $schemaEntry = [ADSI]"LDAP://$schemaDN"
                $fsmOwner = $schemaEntry.Properties["fSMORoleOwner"].Value
                $serverEntry = [ADSI]"LDAP://$fsmOwner"
                $schemaMasterDC = $serverEntry.Properties["dNSHostName"].Value
                if (-not $schemaMasterDC) { $schemaMasterDC = ($fsmOwner -split ',')[1] -replace 'CN=' }

                # Abfrage direkt gegen Schema-Master (umgeht Replikations-Lag)
                $schemaDirect = [ADSI]"LDAP://$schemaMasterDC/$schemaDN"
                $searcherSchema = [System.DirectoryServices.DirectorySearcher]::new($schemaDirect)

                $searcherSchema.Filter = "(cn=ms-Mcs-AdmPwd)"
                $resLegacy = $searcherSchema.FindOne()
                $hasLegacySchema = ($null -ne $resLegacy)

                # Prüft nach allen Varianten (msLAPS-*, ms-LAPS-*)
                $searcherSchema.Filter = "(|(name=msLAPS*)(name=ms-LAPS*)(cn=msLAPS*)(cn=ms-LAPS*))"
                $resModern = $searcherSchema.FindOne()
                $hasModernSchema = ($null -ne $resModern)

                $lblSchemaLegacy.Text = "Legacy Schema: " + $(if ($hasLegacySchema) { "JA (Installiert)" } else { "NEIN" })
                $lblSchemaLegacy.ForeColor = $(if ($hasLegacySchema) { [System.Drawing.Color]::FromArgb(0, 120, 0) } else { [System.Drawing.Color]::Gray })

                $lblSchemaModern.Text = "Modern Windows LAPS Schema: " + $(if ($hasModernSchema) { "JA (Installiert)" } else { "NEIN" })
                $lblSchemaModern.ForeColor = $(if ($hasModernSchema) { [System.Drawing.Color]::FromArgb(0, 120, 0) } else { [System.Drawing.Color]::Gray })

                $activeSummaryText = if ($hasModernSchema -and $hasLegacySchema) {
                    "Beide Schemata vorhanden (Migration moeglich)"
                } elseif ($hasModernSchema) {
                    "Modern Windows LAPS aktiv"
                } elseif ($hasLegacySchema) {
                    "Legacy LAPS aktiv (AdmPwd)"
                } else {
                    "Kein LAPS-Schema installiert"
                }
                $lblActiveSummary.Text = "AD-Status: $activeSummaryText"

                # 2. Geraete-Abfrage
                $rawInput = $txtFilter.Text.Trim()
                if ([string]::IsNullOrEmpty($rawInput)) { $rawInput = "*" }
                $searchPattern = if ($rawInput -notmatch "\*") { "*$rawInput*" } else { $rawInput }

                $searcher = [System.DirectoryServices.DirectorySearcher]::new([System.DirectoryServices.DirectoryEntry]"LDAP://$defNC")
                $searcher.Filter = "(&(objectCategory=computer)(name=$searchPattern))"
                $searcher.PageSize = 1000
                $searcher.PropertiesToLoad.AddRange(@(
                    "samaccountname", "dnshostname", "operatingsystem", "userAccountControl",
                    "ms-mcs-admPwd", "ms-mcs-admtime", "msLAPS-Password", "ms-LAPS-Password",
                    "msLAPS-PasswordExpirationTime", "ms-LAPS-PasswordExpirationTime"
                ))

                $computers = $searcher.FindAll()
                $lblStatus.Text = "Analysiere $($computers.Count) Computer ($toolVersion)..."
                $form.Refresh()

                $selectedTypeFilter = $cmbObjType.SelectedItem
                $selectedStatusFilter = $cmbStatusFilter.SelectedItem

                $countTotal = 0; $countModern = 0; $countLegacy = 0; $countNone = 0

                foreach ($comp in $computers) {
                    $compName = if ($comp.Properties["dnshostname"].Count -gt 0) { $comp.Properties["dnshostname"][0].ToString() } else { $comp.Properties["samaccountname"][0].ToString() }
                    $osName = if ($comp.Properties["operatingsystem"].Count -gt 0) { $comp.Properties["operatingsystem"][0].ToString() } else { "Unbekannt" }

                    $uac = if ($comp.Properties["useraccountcontrol"].Count -gt 0) { [int]$comp.Properties["useraccountcontrol"][0] } else { 0 }
                    $isDisabled = ($uac -band 2) -eq 2
                    $accountStatusText = if ($isDisabled) { "Deaktiviert" } else { "Aktiviert" }

                    $isServer = $osName -match "Server"
                    if ($selectedTypeFilter -eq "Server" -and -not $isServer) { continue }
                    if ($selectedTypeFilter -eq "Computer" -and $isServer) { continue }

                    if ($selectedStatusFilter -eq "Nur Aktivierte" -and $isDisabled) { continue }
                    if ($selectedStatusFilter -eq "Nur Deaktivierte" -and -not $isDisabled) { continue }

                    $legacyTime = $null
                    $hasLegacyValue = ($comp.Properties["ms-mcs-admPwd"].Count -gt 0)
                    if ($comp.Properties["ms-mcs-admtime"].Count -gt 0) {
                        try { $legacyTime = [DateTime]::FromFileTime([int64]$comp.Properties["ms-mcs-admtime"][0]) } catch {}
                    }

                    $modernTime = $null
                    $hasModernValue = ($comp.Properties["mslaps-password"].Count -gt 0 -or $comp.Properties["ms-laps-password"].Count -gt 0)
                    
                    $modExpProp = if ($comp.Properties["mslaps-passwordexpirationtime"].Count -gt 0) {
                        $comp.Properties["mslaps-passwordexpirationtime"][0]
                    } elseif ($comp.Properties["ms-laps-passwordexpirationtime"].Count -gt 0) {
                        $comp.Properties["ms-laps-passwordexpirationtime"][0]
                    } else { $null }

                    if ($null -ne $modExpProp) {
                        try {
                            if ($modExpProp -is [int64]) { $modernTime = [DateTime]::FromFileTime($modExpProp) }
                        } catch {}
                    }

                    $lapsVariant = "Nicht im AD konfiguriert"
                    $passwordStatus = "Kein Kennwort im AD hinterlegt"

                    if ($hasModernValue -or $null -ne $modernTime) {
                        $lapsVariant = "Modernes Windows LAPS"
                        $passwordStatus = if ($modernTime) { "Aktiv (Laeuft ab: $modernTime)" } else { "Aktiv (Kennwort gesetzt)" }
                        $countModern++
                    } elseif ($hasLegacyValue -or $null -ne $legacyTime) {
                        $lapsVariant = "Legacy LAPS (AdmPwd)"
                        $passwordStatus = if ($legacyTime) { "Aktiv (Letzte Aenderung: $legacyTime)" } else { "Aktiv (Kennwort gesetzt)" }
                        $countLegacy++
                    } else {
                        $countNone++
                    }

                    $countTotal++

                    $script:inventoryResults.Add([PSCustomObject]@{
                        Computer        = $compName
                        KontoStatus     = $accountStatusText
                        Betriebssystem  = $osName
                        LapsVariante    = $lapsVariant
                        KennwortStatus  = $passwordStatus
                        LegacyZeit      = if ($legacyTime) { $legacyTime.ToString("dd.MM.yyyy HH:mm") } else { "-" }
                        ModernZeit      = if ($modernTime) { $modernTime.ToString("dd.MM.yyyy HH:mm") } else { "-" }
                    })
                }

                $lblDashTotal.Text  = "Gesamt: $countTotal"
                $lblDashLegacy.Text = "[!] Mit Legacy LAPS: $countLegacy"
                $lblDashModern.Text = "[+] Mit Modern LAPS: $countModern"
                $lblDashNone.Text   = "[-] Ohne LAPS: $countNone"

                $grid.DataSource = [System.Collections.ArrayList]::new($script:inventoryResults)
                $lblStatus.Text = "Scan abgeschlossen ($toolVersion). $($script:inventoryResults.Count) Eintraege geprueft."

            } catch {
                [System.Windows.Forms.MessageBox]::Show("Fehler beim Scan: $($_.Exception.Message)", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $lblStatus.Text = "Fehler aufgetreten ($toolVersion)."
            } finally {
                if ($computers)      { $computers.Dispose() }
                if ($searcher)       { $searcher.Dispose() }
                if ($searcherSchema) { $searcherSchema.Dispose() }
                if ($schemaEntry)    { $schemaEntry.Dispose() }
                if ($rootDSE)        { $rootDSE.Dispose() }

                $form.Cursor = [System.Windows.Forms.Cursors]::Default
                $btnRun.Enabled = $true
                $btnCheckSchemaUpdate.Enabled = $true
            }
        }

        # ----------------------------------------------------------------------
        # AKTION: SCHEMA-UPDATE-DIALOG (v1.3.3)
        # ----------------------------------------------------------------------
        $checkSchemaUpdateAction = {
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            $lblStatus.Text = "Fuehre Vorabpruefung direkt am Schema-Master durch..."
            $form.Refresh()

            # 1. Modul / Cmdlet
            $cmdAvailable = [bool](Get-Command Update-LapsADSchema -ErrorAction SilentlyContinue)

            # 2. Schema Master ermitteln & Erreichbarkeit testen
            $schemaMaster = $null
            $isPingOk = $false
            try {
                $root = [ADSI]"LDAP://RootDSE"
                $schemaEntry = [ADSI]"LDAP://$($root.schemaNamingContext)"
                $fsmOwner = $schemaEntry.Properties["fSMORoleOwner"].Value
                $serverEntry = [ADSI]"LDAP://$fsmOwner"
                $schemaMaster = $serverEntry.Properties["dNSHostName"].Value
                if (-not $schemaMaster) { $schemaMaster = ($fsmOwner -split ',')[1] -replace 'CN=' }
                
                if ($schemaMaster) {
                    $isPingOk = [bool](Test-Connection -ComputerName $schemaMaster -Count 1 -Quiet)
                }
            } catch {}

            # 3. Schema Admin Gruppen-Mitgliedschaft (Token-Prüfung)
            $isSchemaAdmin = $false
            $isLocalAdmin = $false
            try {
                $root = [ADSI]"LDAP://RootDSE"
                $domEntry = [ADSI]"LDAP://$($root.defaultNamingContext)"
                $domSid = (New-Object System.Security.Principal.SecurityIdentifier($domEntry.Properties["objectSid"][0], 0)).Value
                $schemaGroupSid = "$domSid-518"

                $currentUserIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
                $isSchemaAdmin = [bool]($currentUserIdentity.Claims | Where-Object { $_.Value -eq $schemaGroupSid })
                $wp = New-Object System.Security.Principal.WindowsPrincipal($currentUserIdentity)
                $isLocalAdmin = $wp.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
            } catch {}

            # 4. Schema-Status: Explizite Direktabfrage am Schema-Master DC
            $existingLapsAttrCount = 0
            try {
                $targetDc = if ($schemaMaster) { $schemaMaster } else { "localhost" }
                $rootDirect = [ADSI]"LDAP://$targetDc/RootDSE"
                $schemaNC = $rootDirect.schemaNamingContext.ToString()
                $searcherDirect = [System.DirectoryServices.DirectorySearcher]::new([ADSI]"LDAP://$targetDc/$schemaNC")
                # Sucht nach Attributen und Klassen mit LAPS im Namen
                $searcherDirect.Filter = "(|(name=*laps*)(cn=*laps*))"
                $searcherDirect.PageSize = 100
                $attrs = $searcherDirect.FindAll()
                
                # Zählen der Windows LAPS Schema-Objekte (ohne ms-Mcs Legacy)
                foreach ($a in $attrs) {
                    $n = [string]$a.Properties["name"][0]
                    if ($n -match "ms-?LAPS") {
                        $existingLapsAttrCount++
                    }
                }
                if ($attrs) { $attrs.Dispose() }
                $searcherDirect.Dispose()
            } catch {}

            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            $lblStatus.Text = "Vorabpruefung abgeschlossen."

            $isAlreadyUpdated = ($existingLapsAttrCount -gt 0)
            $isReady = ($cmdAvailable -and $isPingOk -and $isSchemaAdmin)

            # DIALOG-AUFBAU
            $dlg = New-Object System.Windows.Forms.Form
            $dlg.Text = "LAPS AD Schema Update - Vorabpruefungsbericht ($toolVersion)"
            $dlg.Size = New-Object System.Drawing.Size(860, 690)
            $dlg.StartPosition = "CenterParent"
            $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
            $dlg.MaximizeBox = $false
            $dlg.MinimizeBox = $false
            $dlg.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
            $dlg.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 253)

            $pnlHero = New-Object System.Windows.Forms.Panel
            $pnlHero.Dock = [System.Windows.Forms.DockStyle]::Top
            $pnlHero.Height = 85
            $pnlHero.Padding = New-Object System.Windows.Forms.Padding(20, 15, 20, 15)
            
            if ($isAlreadyUpdated) {
                $pnlHero.BackColor = [System.Drawing.Color]::FromArgb(235, 248, 238)
            } elseif ($isReady) {
                $pnlHero.BackColor = [System.Drawing.Color]::FromArgb(238, 246, 255)
            } else {
                $pnlHero.BackColor = [System.Drawing.Color]::FromArgb(254, 242, 242)
            }
            $dlg.Controls.Add($pnlHero)

            $lblHeroTitle = New-Object System.Windows.Forms.Label
            $lblHeroTitle.Dock = [System.Windows.Forms.DockStyle]::Top
            $lblHeroTitle.Height = 28
            $lblHeroTitle.Font = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)
            
            if ($isAlreadyUpdated) {
                $lblHeroTitle.ForeColor = [System.Drawing.Color]::FromArgb(22, 101, 52)
                $lblHeroTitle.Text = "SCHEMA-UPDATE BEREITS ERFOLGREICH DURCHGEFUEHRT"
            } elseif ($isReady) {
                $lblHeroTitle.ForeColor = [System.Drawing.Color]::FromArgb(20, 80, 160)
                $lblHeroTitle.Text = "BEREIT FUER DAS SCHEMA-UPDATE"
            } else {
                $lblHeroTitle.ForeColor = [System.Drawing.Color]::FromArgb(153, 27, 27)
                $lblHeroTitle.Text = "VORAUSSETZUNGEN NICHT VOLLSTAENDIG ERFUELLT"
            }
            $pnlHero.Controls.Add($lblHeroTitle)

            $lblHeroSub = New-Object System.Windows.Forms.Label
            $lblHeroSub.Dock = [System.Windows.Forms.DockStyle]::Fill
            $lblHeroSub.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
            $lblHeroSub.ForeColor = [System.Drawing.Color]::FromArgb(55, 65, 81)
            $lblHeroSub.Text = $(if ($isAlreadyUpdated) {
                "Das Active Directory Schema auf $schemaMaster enthaelt bereits die modernen Windows LAPS-Definitionen ($existingLapsAttrCount Schema-Objekte)."
            } elseif ($isReady) {
                "Alle Voraussetzungen sind erfuellt. Sie koennen den Update-Befehl nun am Schema-Master ausfuehren."
            } else {
                "Bitte beachten Sie die rot markierten Warnhinweise unten, bevor Sie das Update anstossen."
            })
            $pnlHero.Controls.Add($lblHeroSub)

            $pnlCards = New-Object System.Windows.Forms.TableLayoutPanel
            $pnlCards.Dock = [System.Windows.Forms.DockStyle]::Fill
            $pnlCards.ColumnCount = 1
            $pnlCards.RowCount = 4
            $pnlCards.Padding = New-Object System.Windows.Forms.Padding(20, 15, 20, 15)
            for ($i=0; $i -lt 4; $i++) {
                [void]$pnlCards.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 105)))
            }
            $dlg.Controls.Add($pnlCards)
            $pnlCards.BringToFront()

            function Add-StatusCard {
                param(
                    [string]$StepNumber,
                    [string]$Title,
                    [string]$StatusBadge,
                    [System.Drawing.Color]$BadgeColor,
                    [string]$MainText,
                    [string]$DetailHint,
                    [System.Drawing.Color]$DetailColor = [System.Drawing.Color]::FromArgb(107, 114, 128)
                )

                $card = New-Object System.Windows.Forms.Panel
                $card.Dock = [System.Windows.Forms.DockStyle]::Fill
                $card.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)
                $card.BackColor = [System.Drawing.Color]::White
                $card.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
                $card.Padding = New-Object System.Windows.Forms.Padding(12, 8, 12, 8)

                $pHead = New-Object System.Windows.Forms.Panel
                $pHead.Dock = [System.Windows.Forms.DockStyle]::Top
                $pHead.Height = 26
                $card.Controls.Add($pHead)

                $lblT = New-Object System.Windows.Forms.Label
                $lblT.Text = "[$StepNumber]  $Title"
                $lblT.Dock = [System.Windows.Forms.DockStyle]::Left
                $lblT.AutoSize = $true
                $lblT.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
                $lblT.ForeColor = [System.Drawing.Color]::FromArgb(31, 41, 55)
                $pHead.Controls.Add($lblT)

                $lblB = New-Object System.Windows.Forms.Label
                $lblB.Text = " $StatusBadge "
                $lblB.Dock = [System.Windows.Forms.DockStyle]::Right
                $lblB.AutoSize = $true
                $lblB.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
                $lblB.ForeColor = $BadgeColor
                $pHead.Controls.Add($lblB)

                $lblMain = New-Object System.Windows.Forms.Label
                $lblMain.Text = $MainText
                $lblMain.Dock = [System.Windows.Forms.DockStyle]::Top
                $lblMain.Height = 22
                $lblMain.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
                $lblMain.ForeColor = [System.Drawing.Color]::FromArgb(55, 65, 81)
                $card.Controls.Add($lblMain)

                $lblDet = New-Object System.Windows.Forms.Label
                $lblDet.Text = $DetailHint
                $lblDet.Dock = [System.Windows.Forms.DockStyle]::Fill
                $lblDet.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
                $lblDet.ForeColor = $DetailColor
                $card.Controls.Add($lblDet)

                $lblDet.BringToFront()
                $lblMain.BringToFront()
                $pHead.BringToFront()

                $pnlCards.Controls.Add($card)
            }

            if ($cmdAvailable) {
                Add-StatusCard -StepNumber "1/4" -Title "Windows LAPS PowerShell-Cmdlet" `
                    -StatusBadge "VERFUEGBAR [OK]" -BadgeColor ([System.Drawing.Color]::FromArgb(22, 101, 52)) `
                    -MainText "Befehl 'Update-LapsADSchema' ist auf diesem System vorhanden." `
                    -DetailHint "Natives Management-Modul geladen."
            } else {
                Add-StatusCard -StepNumber "1/4" -Title "Windows LAPS PowerShell-Cmdlet" `
                    -StatusBadge "NICHT GEFUNDEN" -BadgeColor ([System.Drawing.Color]::FromArgb(185, 28, 28)) `
                    -MainText "Das Cmdlet 'Update-LapsADSchema' ist auf dieser Station nicht installiert." `
                    -DetailHint "Aktion: Windows Updates installieren oder RSAT nachruesten." `
                    -DetailColor ([System.Drawing.Color]::FromArgb(185, 28, 28))
            }

            if ($schemaMaster -and $isPingOk) {
                Add-StatusCard -StepNumber "2/4" -Title "Schema-Master FSMO-Rolle" `
                    -StatusBadge "ONLINE [OK]" -BadgeColor ([System.Drawing.Color]::FromArgb(22, 101, 52)) `
                    -MainText "DC: $schemaMaster" `
                    -DetailHint "Netzwerk- und LDAP-Verbindung zum Schema-Master erfolgreich hergestellt."
            } else {
                $dcName = if ($schemaMaster) { $schemaMaster } else { "Nicht ermittelbar" }
                Add-StatusCard -StepNumber "2/4" -Title "Schema-Master FSMO-Rolle" `
                    -StatusBadge "NICHT ERREICHBAR" -BadgeColor ([System.Drawing.Color]::FromArgb(185, 28, 28)) `
                    -MainText "DC: $dcName" `
                    -DetailHint "Verbindungstest zum Schema-Master ist fehlgeschlagen." `
                    -DetailColor ([System.Drawing.Color]::FromArgb(185, 28, 28))
            }

            if ($isSchemaAdmin) {
                $adminText = "Aktiver Benutzer besitzt gueltige Schema-Admin-Rechte (SID -518)."
                if (-not $isLocalAdmin) { $adminText += " (Hinweis: Konsole nicht erhoeht als Admin gestartet)" }
                Add-StatusCard -StepNumber "3/4" -Title "Gruppenmitgliedschaft 'Schema-Admins'" `
                    -StatusBadge "BERECHTIGT [OK]" -BadgeColor ([System.Drawing.Color]::FromArgb(22, 101, 52)) `
                    -MainText $adminText `
                    -DetailHint "Kerberos-Token enthaelt den Sicherheitsanspruch fuer Schema-Aenderungen."
            } else {
                Add-StatusCard -StepNumber "3/4" -Title "Gruppenmitgliedschaft 'Schema-Admins'" `
                    -StatusBadge "KEINE RECHTE" -BadgeColor ([System.Drawing.Color]::FromArgb(185, 28, 28)) `
                    -MainText "Der Benutzer besitzt im aktuellen Kerberos-Token KEINE Schema-Admin-Rechte (SID -518)." `
                    -DetailHint "Falls die Gruppe kuerzlich zugewiesen wurde: Abmelden und neu anmelden!" `
                    -DetailColor ([System.Drawing.Color]::FromArgb(185, 28, 28))
            }

            if ($isAlreadyUpdated) {
                Add-StatusCard -StepNumber "4/4" -Title "Aktueller LAPS-Schema-Status im AD" `
                    -StatusBadge "BEREITS INSTALLIERT [OK]" -BadgeColor ([System.Drawing.Color]::FromArgb(22, 101, 52)) `
                    -MainText "Gefundene Windows LAPS Attribute am Schema-Master: $existingLapsAttrCount Stk." `
                    -DetailHint "Das AD-Schema ist aktiv fuer Modern Windows LAPS erweitert. Kein Update noetig."
            } else {
                Add-StatusCard -StepNumber "4/4" -Title "Aktueller LAPS-Schema-Status im AD" `
                    -StatusBadge "UPDATE AUSSTEHEND" -BadgeColor ([System.Drawing.Color]::FromArgb(2, 132, 199)) `
                    -MainText "Das AD-Schema enthaelt noch keine 'ms-LAPS-*'-Attribute." `
                    -DetailHint "Fuehren Sie den Befehl 'Update-LapsADSchema -Verbose' aus, um das Schema zu erweitern."
            }

            $pnlFoot = New-Object System.Windows.Forms.Panel
            $pnlFoot.Dock = [System.Windows.Forms.DockStyle]::Bottom
            $pnlFoot.Height = 60
            $pnlFoot.BackColor = [System.Drawing.Color]::FromArgb(238, 242, 246)
            $pnlFoot.Padding = New-Object System.Windows.Forms.Padding(20, 12, 20, 12)
            $dlg.Controls.Add($pnlFoot)

            $btnCopyScript = New-Object System.Windows.Forms.Button
            $btnCopyScript.Text = "Befehl kopieren: Update-LapsADSchema"
            $btnCopyScript.Dock = [System.Windows.Forms.DockStyle]::Left
            $btnCopyScript.Width = 280
            $btnCopyScript.BackColor = [System.Drawing.Color]::FromArgb(240, 244, 248)
            $btnCopyScript.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            $btnCopyScript.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
            $btnCopyScript.Add_Click({
                [System.Windows.Forms.Clipboard]::SetText("Update-LapsADSchema -Verbose")
                [System.Windows.Forms.MessageBox]::Show("Befehl 'Update-LapsADSchema -Verbose' in die Zwischenablage kopiert!`r`n`r`nFuehren Sie diesen in einer administrativen PowerShell-Sitzung aus.", "Kopiert", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            })
            $pnlFoot.Controls.Add($btnCopyScript)

            $btnCloseDlg = New-Object System.Windows.Forms.Button
            $btnCloseDlg.Text = "Schliessen"
            $btnCloseDlg.Dock = [System.Windows.Forms.DockStyle]::Right
            $btnCloseDlg.Width = 120
            $btnCloseDlg.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
            $btnCloseDlg.ForeColor = [System.Drawing.Color]::White
            $btnCloseDlg.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            $btnCloseDlg.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
            $btnCloseDlg.Add_Click({ $dlg.Close() })
            $pnlFoot.Controls.Add($btnCloseDlg)

            [void]$dlg.ShowDialog($form)
            $dlg.Dispose()
        }

        # Event-Wiring
        $btnRun.Add_Click($runInventoryScan)
        $btnCheckSchemaUpdate.Add_Click($checkSchemaUpdateAction)

        $txtFilter.Add_KeyDown({
            param($sender, $e)
            if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
                $e.SuppressKeyPress = $true
                & $runInventoryScan
            }
        })

        $btnExport.Add_Click({
            if (-not $grid.DataSource -or $script:inventoryResults.Count -eq 0) { return }
            $sfd = New-Object System.Windows.Forms.SaveFileDialog
            $sfd.Filter = "CSV-Datei (*.csv)|*.csv"
            $sfd.FileName = "LAPS_Inventory_Audit_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
            if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $script:inventoryResults | Export-Csv -Path $sfd.FileName -Delimiter ';' -NoTypeInformation -Encoding UTF8
                [System.Windows.Forms.MessageBox]::Show("Erfolgreich nach '$($sfd.FileName)' exportiert ($toolVersion)!", "Export OK", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        })

        $form.Add_FormClosing({
            $isClosing = $true
        })

        $form.Add_Shown({ & $runInventoryScan })
        [void]$form.ShowDialog()

    } finally {
        if ($null -ne $form) { $form.Dispose() }
        if ($script:inventoryResults) { $script:inventoryResults.Clear() }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

if ($MyInvocation.InvocationName -ne '.' -and $MyInvocation.Line -notmatch 'Open-Tool') {
    Open-ToolLapsInventoryAudit
}