<#
================================================================================
 TOOL 26: WinRM - Status aller Domänen-Computer
 Abfragefilterung direkt im LDAP-Query integriert
================================================================================
#>

# ISE-Freeze-Schutz
if ($psISE -and ($MyInvocation.InvocationName -ne '.')) {
    if ($PSCommandPath) {
        Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
        return
    }
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.DirectoryServices

if (-not [System.Windows.Forms.Application]::RenderWithVisualStyles) {
    try { [System.Windows.Forms.Application]::EnableVisualStyles() } catch {}
}

$script:UITheme = @{
    FontDefault        = New-Object System.Drawing.Font("Segoe UI", 9.0)
    FontBold           = New-Object System.Drawing.Font("Segoe UI", 9.0, [System.Drawing.FontStyle]::Bold)
    FontHeader         = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    HeaderHeight       = 36
    RowHeight          = 28
    HeaderBackColor    = [System.Drawing.Color]::FromArgb(238, 242, 248)
    HeaderForeColor    = [System.Drawing.Color]::FromArgb(30, 41, 59)
    GridLineColor      = [System.Drawing.Color]::FromArgb(226, 232, 240)
    RowBackColor       = [System.Drawing.Color]::White
    RowAltBackColor    = [System.Drawing.Color]::FromArgb(248, 250, 252)
    SelectionBackColor = [System.Drawing.Color]::FromArgb(203, 228, 249)
    SelectionForeColor = [System.Drawing.Color]::Black
}

function Show-Tool26-WinRMStatus {
    [CmdletBinding()]
    param()

    # Domänen-Prüfung
    $isPartOfDomain = $false
    try {
        $sysInfo = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue
        $isPartOfDomain = $sysInfo.PartOfDomain
    } catch {
        $isPartOfDomain = $false
    }

    if (-not $isPartOfDomain) {
        [System.Windows.Forms.MessageBox]::Show(
            "Dieses Werkzeug erfordert eine Active Directory Domänenmitgliedschaft.",
            "Keine Domäne gefunden",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    # 1. Hauptformular
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Tool 26: WinRM-Status aller Domänen-Computer"
    $form.Size = New-Object System.Drawing.Size(1260, 820)
    $form.MinimumSize = New-Object System.Drawing.Size(1020, 620)
    $form.StartPosition = "CenterScreen"
    $form.Font = $script:UITheme.FontDefault
    $form.BackColor = [System.Drawing.Color]::FromArgb(243, 245, 249)
    $form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi

    $script:rawComputerList = [System.Collections.Generic.List[PSCustomObject]]::new()
    $script:runspacePool    = $null
    $script:tasks           = [System.Collections.Generic.List[PSCustomObject]]::new()
    $script:isScanning      = $false

    # 2. Responsiver Kopfbereich
    $pnlHeader = New-Object System.Windows.Forms.Panel
    $pnlHeader.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlHeader.Height = 110
    $pnlHeader.BackColor = [System.Drawing.Color]::White
    $pnlHeader.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $form.Controls.Add($pnlHeader)

    # Zeile 1: Aktionen & Dashboard
    $flowActions = New-Object System.Windows.Forms.FlowLayoutPanel
    $flowActions.Dock = [System.Windows.Forms.DockStyle]::Top
    $flowActions.Height = 44
    $flowActions.Padding = New-Object System.Windows.Forms.Padding(8, 6, 8, 2)
    $flowActions.WrapContents = $false
    $pnlHeader.Controls.Add($flowActions)

    $btnScan = New-Object System.Windows.Forms.Button
    $btnScan.Text = "Scan starten"
    $btnScan.Size = New-Object System.Drawing.Size(125, 30)
    $btnScan.BackColor = [System.Drawing.Color]::FromArgb(16, 185, 129)
    $btnScan.ForeColor = [System.Drawing.Color]::White
    $btnScan.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnScan.Font = $script:UITheme.FontBold
    $btnScan.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0)
    $flowActions.Controls.Add($btnScan)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Abbrechen"
    $btnCancel.Size = New-Object System.Drawing.Size(110, 30)
    $btnCancel.BackColor = [System.Drawing.Color]::FromArgb(239, 68, 68)
    $btnCancel.ForeColor = [System.Drawing.Color]::White
    $btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnCancel.Font = $script:UITheme.FontBold
    $btnCancel.Enabled = $false
    $btnCancel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0)
    $flowActions.Controls.Add($btnCancel)

    $btnExport = New-Object System.Windows.Forms.Button
    $btnExport.Text = "CSV Export"
    $btnExport.Size = New-Object System.Drawing.Size(115, 30)
    $btnExport.BackColor = [System.Drawing.Color]::FromArgb(59, 130, 246)
    $btnExport.ForeColor = [System.Drawing.Color]::White
    $btnExport.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnExport.Font = $script:UITheme.FontBold
    $btnExport.Margin = New-Object System.Windows.Forms.Padding(0, 0, 15, 0)
    $flowActions.Controls.Add($btnExport)

    $lblStats = New-Object System.Windows.Forms.Label
    $lblStats.AutoSize = $true
    $lblStats.Font = $script:UITheme.FontBold
    $lblStats.ForeColor = [System.Drawing.Color]::FromArgb(30, 41, 59)
    $lblStats.Text = "Bereit. Parameter definieren und 'Scan starten' klicken."
    $lblStats.Margin = New-Object System.Windows.Forms.Padding(0, 7, 0, 0)
    $flowActions.Controls.Add($lblStats)

    # Zeile 2: LDAP-Abfragefilter
    $flowFilters = New-Object System.Windows.Forms.FlowLayoutPanel
    $flowFilters.Dock = [System.Windows.Forms.DockStyle]::Top
    $flowFilters.Height = 52
    $flowFilters.Padding = New-Object System.Windows.Forms.Padding(8, 8, 8, 4)
    $flowFilters.WrapContents = $false
    $pnlHeader.Controls.Add($flowFilters)
    $flowFilters.BringToFront()

    # LDAP Name Filter
    $lblName = New-Object System.Windows.Forms.Label
    $lblName.Text = "Name:"
    $lblName.AutoSize = $true
    $lblName.Font = $script:UITheme.FontBold
    $lblName.Margin = New-Object System.Windows.Forms.Padding(0, 5, 4, 0)
    $flowFilters.Controls.Add($lblName)

    $txtNamePattern = New-Object System.Windows.Forms.TextBox
    $txtNamePattern.Size = New-Object System.Drawing.Size(130, 25)
    $txtNamePattern.Margin = New-Object System.Windows.Forms.Padding(0, 2, 6, 0)
    $flowFilters.Controls.Add($txtNamePattern)

    $cmbMatchMode = New-Object System.Windows.Forms.ComboBox
    $cmbMatchMode.Size = New-Object System.Drawing.Size(135, 25)
    $cmbMatchMode.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$cmbMatchMode.Items.AddRange(@(
        "Enthält (*Wert*)",
        "Beginnt mit (Wert*)",
        "Endet mit (*Wert)",
        "Exakt (=Wert)",
        "Alle (*)"
    ))
    $cmbMatchMode.SelectedIndex = 0
    $cmbMatchMode.Margin = New-Object System.Windows.Forms.Padding(0, 2, 16, 0)
    $flowFilters.Controls.Add($cmbMatchMode)

    # LDAP OS Typ Filter
    $lblFilterType = New-Object System.Windows.Forms.Label
    $lblFilterType.Text = "Typ:"
    $lblFilterType.AutoSize = $true
    $lblFilterType.Font = $script:UITheme.FontBold
    $lblFilterType.Margin = New-Object System.Windows.Forms.Padding(0, 5, 4, 0)
    $flowFilters.Controls.Add($lblFilterType)

    $cmbType = New-Object System.Windows.Forms.ComboBox
    $cmbType.Size = New-Object System.Drawing.Size(120, 25)
    $cmbType.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$cmbType.Items.AddRange(@("Alle Typen", "Nur Server", "Nur Clients"))
    $cmbType.SelectedIndex = 0
    $cmbType.Margin = New-Object System.Windows.Forms.Padding(0, 2, 16, 0)
    $flowFilters.Controls.Add($cmbType)

    # Live-Filter WinRM Ergebnis
    $lblFilterWinRM = New-Object System.Windows.Forms.Label
    $lblFilterWinRM.Text = "WinRM:"
    $lblFilterWinRM.AutoSize = $true
    $lblFilterWinRM.Font = $script:UITheme.FontBold
    $lblFilterWinRM.Margin = New-Object System.Windows.Forms.Padding(0, 5, 4, 0)
    $flowFilters.Controls.Add($lblFilterWinRM)

    $cmbWinRM = New-Object System.Windows.Forms.ComboBox
    $cmbWinRM.Size = New-Object System.Drawing.Size(120, 25)
    $cmbWinRM.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$cmbWinRM.Items.AddRange(@("Alle Status", "WinRM OK", "WinRM Fehler", "Ping Offline"))
    $cmbWinRM.SelectedIndex = 0
    $cmbWinRM.Margin = New-Object System.Windows.Forms.Padding(0, 2, 16, 0)
    $flowFilters.Controls.Add($cmbWinRM)

    # LDAP UAC Filter
    $chkActiveOnly = New-Object System.Windows.Forms.CheckBox
    $chkActiveOnly.Text = "Nur aktive Computer"
    $chkActiveOnly.AutoSize = $true
    $chkActiveOnly.Checked = $true
    $chkActiveOnly.Font = $script:UITheme.FontBold
    $chkActiveOnly.Margin = New-Object System.Windows.Forms.Padding(0, 5, 0, 0)
    $flowFilters.Controls.Add($chkActiveOnly)

    # Progressbar
    $progBar = New-Object System.Windows.Forms.ProgressBar
    $progBar.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $progBar.Height = 5
    $progBar.Visible = $false
    $pnlHeader.Controls.Add($progBar)

    # 3. DataGridView
    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = [System.Windows.Forms.DockStyle]::Fill
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.ReadOnly = $true
    $grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grid.MultiSelect = $false
    $grid.BackgroundColor = [System.Drawing.Color]::White
    $grid.RowHeadersVisible = $false
    $grid.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $grid.CellBorderStyle = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
    $grid.GridColor = $script:UITheme.GridLineColor

    $grid.EnableHeadersVisualStyles = $false
    $grid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
    $grid.ColumnHeadersHeight = $script:UITheme.HeaderHeight
    $grid.RowTemplate.Height = $script:UITheme.RowHeight

    $hdrStyle = New-Object System.Windows.Forms.DataGridViewCellStyle
    $hdrStyle.Font = $script:UITheme.FontHeader
    $hdrStyle.BackColor = $script:UITheme.HeaderBackColor
    $hdrStyle.ForeColor = $script:UITheme.HeaderForeColor
    $hdrStyle.Alignment = [System.Drawing.ContentAlignment]::MiddleLeft
    $hdrStyle.Padding = New-Object System.Windows.Forms.Padding(6, 0, 6, 0)
    $grid.ColumnHeadersDefaultCellStyle = $hdrStyle

    $cellStyle = New-Object System.Windows.Forms.DataGridViewCellStyle
    $cellStyle.Font = $script:UITheme.FontDefault
    $cellStyle.SelectionBackColor = $script:UITheme.SelectionBackColor
    $cellStyle.SelectionForeColor = $script:UITheme.SelectionForeColor
    $cellStyle.Padding = New-Object System.Windows.Forms.Padding(6, 0, 6, 0)
    $grid.DefaultCellStyle = $cellStyle
    $grid.AlternatingRowsDefaultCellStyle.BackColor = $script:UITheme.RowAltBackColor

    $form.Controls.Add($grid)
    $grid.BringToFront()

    [void]$grid.Columns.Add("ComputerName", "Computername")
    [void]$grid.Columns.Add("DNSHostName", "FQDN")
    [void]$grid.Columns.Add("Type", "Typ")
    [void]$grid.Columns.Add("OperatingSystem", "Betriebssystem")
    [void]$grid.Columns.Add("Ping", "Ping (ICMP)")
    [void]$grid.Columns.Add("WinRMStatus", "WinRM Status")
    [void]$grid.Columns.Add("WinRMDetails", "WinRM Details / Fehler")
    [void]$grid.Columns.Add("AccountStatus", "AD-Konto")

    $grid.Columns["ComputerName"].Width = 140
    $grid.Columns["DNSHostName"].Width = 190
    $grid.Columns["Type"].Width = 85
    $grid.Columns["OperatingSystem"].Width = 220
    $grid.Columns["Ping"].Width = 105
    $grid.Columns["WinRMStatus"].Width = 115
    $grid.Columns["AccountStatus"].Width = 95
    $grid.Columns["WinRMDetails"].AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill

    $grid.Add_CellFormatting({
        param($s, $e)
        if ($e.RowIndex -lt 0) { return }
        $colName = $grid.Columns[$e.ColumnIndex].Name
        $val = [string]$e.Value

        if ($colName -eq "WinRMStatus") {
            if ($val -eq "OK") {
                $e.CellStyle.ForeColor = [System.Drawing.Color]::DarkGreen
                $e.CellStyle.Font = $script:UITheme.FontBold
            } elseif ($val -eq "Fehler") {
                $e.CellStyle.ForeColor = [System.Drawing.Color]::FromArgb(220, 38, 38)
            }
        }
        if ($colName -eq "Ping") {
            if ($val -eq "Online") {
                $e.CellStyle.ForeColor = [System.Drawing.Color]::DarkGreen
            } elseif ($val -eq "Offline") {
                $e.CellStyle.ForeColor = [System.Drawing.Color]::FromArgb(217, 119, 6)
            }
        }
    })

    # 4. LDAP-Abfrage mit integriertem Filter
    function Get-ADComputersFilteredLDAP {
        param(
            [string]$NamePattern,
            [string]$MatchMode,
            [string]$TypeFilter,
            [bool]$ActiveOnly
        )

        $filterParts = [System.Collections.Generic.List[string]]::new()
        $filterParts.Add("(objectCategory=computer)")

        # UAC-Filter
        if ($ActiveOnly) {
            $filterParts.Add("(!(userAccountControl:1.2.840.113556.1.4.803:=2))")
        }

        # Typ-Filter (Server / Client)
        if ($TypeFilter -eq "Nur Server") {
            $filterParts.Add("(operatingSystem=*Server*)")
        } elseif ($TypeFilter -eq "Nur Clients") {
            $filterParts.Add("(!(operatingSystem=*Server*))")
        }

        # Namensmuster-Filter
        $cleanPattern = $NamePattern.Trim()
        if (-not [string]::IsNullOrWhiteSpace($cleanPattern) -and $MatchMode -ne "Alle (*)") {
            $ldapNameTerm = switch ($MatchMode) {
                "Enthält (*Wert*)"    { "*$cleanPattern*" }
                "Beginnt mit (Wert*)" { "$cleanPattern*" }
                "Endet mit (*Wert)"   { "*$cleanPattern" }
                "Exakt (=Wert)"       { "$cleanPattern" }
                default               { "*$cleanPattern*" }
            }
            # Auf Name oder DNSHostName matchen
            $filterParts.Add("(|(name=$ldapNameTerm)(dNSHostName=$ldapNameTerm))")
        }

        # Vollständigen LDAP-Filter zusammensetzen
        $finalFilter = "(&" + ($filterParts -join "") + ")"

        $rootDSE = [System.DirectoryServices.DirectoryEntry]::new("LDAP://RootDSE")
        $defNC   = $rootDSE.Properties["defaultNamingContext"][0]
        $rootDSE.Dispose()

        $entry    = [System.DirectoryServices.DirectoryEntry]::new("LDAP://$defNC")
        $searcher = [System.DirectoryServices.DirectorySearcher]::new($entry)
        $searcher.Filter = $finalFilter
        $searcher.PageSize = 1000
        [void]$searcher.PropertiesToLoad.AddRange(@("name", "dnshostname", "operatingsystem", "useraccountcontrol"))

        $results = $searcher.FindAll()
        $list = [System.Collections.Generic.List[PSCustomObject]]::new()

        foreach ($r in $results) {
            $cName = if ($r.Properties["name"].Count -gt 0) { [string]$r.Properties["name"][0] } else { "N/A" }
            $dns   = if ($r.Properties["dnshostname"].Count -gt 0) { [string]$r.Properties["dnshostname"][0] } else { $cName }
            $os    = if ($r.Properties["operatingsystem"].Count -gt 0) { [string]$r.Properties["operatingsystem"][0] } else { "Unbekannt" }
            $uac   = if ($r.Properties["useraccountcontrol"].Count -gt 0) { [int]$r.Properties["useraccountcontrol"][0] } else { 0 }
            
            $isServer   = ($os -like "*Server*")
            $isDisabled = ($uac -band 2) -ne 0

            $list.Add([PSCustomObject]@{
                ComputerName    = $cName
                DNSHostName     = $dns
                Type            = if ($isServer) { "Server" } else { "Client" }
                OperatingSystem = $os
                Ping            = "Ausstehend"
                WinRMStatus     = "Ausstehend"
                WinRMDetails    = ""
                AccountStatus   = if ($isDisabled) { "Deaktiviert" } else { "Aktiv" }
            })
        }

        $searcher.Dispose()
        $entry.Dispose()
        return $list
    }

    # 5. Grid Synchronisation (WinRM Dropdown)
    $applyWinRMFilter = {
        $winrmFilter = [string]$cmbWinRM.SelectedItem

        $grid.SuspendLayout()
        $grid.Rows.Clear()

        $matchedCount = 0
        $okCount = 0

        foreach ($item in $script:rawComputerList) {
            if ($item.WinRMStatus -eq "OK") { $okCount++ }

            $matchWinRM = switch ($winrmFilter) {
                "WinRM OK"     { $item.WinRMStatus -eq "OK" }
                "WinRM Fehler" { $item.WinRMStatus -eq "Fehler" }
                "Ping Offline" { $item.Ping -eq "Offline" }
                default        { $true }
            }

            if ($matchWinRM) {
                $matchedCount++
                [void]$grid.Rows.Add(
                    $item.ComputerName,
                    $item.DNSHostName,
                    $item.Type,
                    $item.OperatingSystem,
                    $item.Ping,
                    $item.WinRMStatus,
                    $item.WinRMDetails,
                    $item.AccountStatus
                )
            }
        }

        $grid.ResumeLayout()
        $lblStats.Text = "LDAP-Treffer: $($script:rawComputerList.Count) | Angezeigt: $matchedCount | WinRM OK: $okCount"
    }

    $cmbWinRM.Add_SelectedIndexChanged({ & $applyWinRMFilter })

    # Enter-Taste im Suchfeld startet direkt
    $txtNamePattern.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $_.SuppressKeyPress = $true
            $btnScan.PerformClick()
        }
    })

    # 6. Multi-Threading Scan
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 250

    $workerBlock = {
        param($ComputerName, $DnsName)
        $pingOk = $false
        $winrmOk = $false
        $details = ""

        try {
            $ping = [System.Net.NetworkInformation.Ping]::new()
            $reply = $ping.Send($DnsName, 650)
            if ($reply -and $reply.Status -eq [System.Net.NetworkInformation.IPStatus]::Success) {
                $pingOk = $true
            }
            $ping.Dispose()
        } catch {
            $pingOk = $false
        }

        if ($pingOk) {
            try {
                $wsTest = Test-WSMan -ComputerName $DnsName -ErrorAction Stop
                if ($null -ne $wsTest) {
                    $winrmOk = $true
                    $details = "WSMan Stack: $($wsTest.ProductVersion) (Port OK)"
                }
            } catch {
                $winrmOk = $false
                $msg = $_.Exception.Message
                if ($msg -match "Access is denied" -or $msg -match "Zugriff verweigert") {
                    $details = "Port erreichbar, Authentifizierung verweigert"
                } elseif ($msg -match "timed out" -or $msg -match "Zeitüberschreitung") {
                    $details = "Port 5985 blockiert / Firewall"
                } else {
                    $details = ($msg -replace "[\r\n]+", " ").Trim()
                    if ($details.Length -gt 55) { $details = $details.Substring(0, 55) + "..." }
                }
            }
        } else {
            $details = "Host nicht erreichbar (Ping Timeout)"
        }

        return [PSCustomObject]@{
            ComputerName = $ComputerName
            Ping         = if ($pingOk) { "Online" } else { "Offline" }
            WinRMStatus  = if ($winrmOk) { "OK" } else { "Fehler" }
            WinRMDetails = $details
        }
    }

    $timer.Add_Tick({
        if (-not $script:isScanning) { return }

        $finishedTasks = [System.Collections.Generic.List[PSCustomObject]]::new()

        foreach ($t in $script:tasks) {
            if ($t.AsyncResult.IsCompleted) {
                $finishedTasks.Add($t)
                try {
                    $output = $t.PowerShell.EndInvoke($t.AsyncResult)
                    if ($output -and $output.Count -gt 0) {
                        $res = $output[0]
                        if ($null -ne $res) {
                            $target = $script:rawComputerList | Where-Object { $_.ComputerName -eq $res.ComputerName }
                            if ($target) {
                                $target.Ping         = $res.Ping
                                $target.WinRMStatus  = $res.WinRMStatus
                                $target.WinRMDetails = $res.WinRMDetails

                                foreach ($row in $grid.Rows) {
                                    if ($row.Cells["ComputerName"].Value -eq $res.ComputerName) {
                                        $row.Cells["Ping"].Value         = $res.Ping
                                        $row.Cells["WinRMStatus"].Value  = $res.WinRMStatus
                                        $row.Cells["WinRMDetails"].Value = $res.WinRMDetails
                                        break
                                    }
                                }
                            }
                        }
                    }
                } catch {
                } finally {
                    $t.PowerShell.Dispose()
                }
            }
        }

        foreach ($ft in $finishedTasks) {
            [void]$script:tasks.Remove($ft)
        }

        $done = $progBar.Maximum - $script:tasks.Count
        $progBar.Value = [Math]::Min([Math]::Max(0, $done), $progBar.Maximum)

        $okNow = ($script:rawComputerList | Where-Object { $_.WinRMStatus -eq "OK" }).Count
        $lblStats.Text = "Prüfe Rechner... $done von $($progBar.Maximum) abgeschlossen. (WinRM OK: $okNow)"

        if ($script:tasks.Count -eq 0) {
            $timer.Stop()
            $script:isScanning = $false

            if ($script:runspacePool) {
                $script:runspacePool.Close()
                $script:runspacePool.Dispose()
                $script:runspacePool = $null
            }

            $btnScan.Enabled       = $true
            $btnCancel.Enabled     = $false
            $txtNamePattern.Enabled= $true
            $cmbMatchMode.Enabled  = $true
            $cmbType.Enabled       = $true
            $chkActiveOnly.Enabled = $true
            $progBar.Visible       = $false
            & $applyWinRMFilter
        }
    })

    $btnScan.Add_Click({
        $btnScan.Enabled       = $false
        $btnCancel.Enabled     = $true
        $txtNamePattern.Enabled= $false
        $cmbMatchMode.Enabled  = $false
        $cmbType.Enabled       = $false
        $chkActiveOnly.Enabled = $false
        $progBar.Visible       = $true
        $progBar.Value         = 0

        $lblStats.Text = "Führe gefilterte LDAP-Abfrage im Active Directory aus..."
        [System.Windows.Forms.Application]::DoEvents()

        # Abfrage direkt mit LDAP-Filtern
        $script:rawComputerList = Get-ADComputersFilteredLDAP `
            -NamePattern $txtNamePattern.Text `
            -MatchMode ([string]$cmbMatchMode.SelectedItem) `
            -TypeFilter ([string]$cmbType.SelectedItem) `
            -ActiveOnly $chkActiveOnly.Checked

        if ($script:rawComputerList.Count -eq 0) {
            $lblStats.Text = "Keine Computer mit den gewählten Filterkriterien im AD gefunden."
            $btnScan.Enabled       = $true
            $btnCancel.Enabled     = $false
            $txtNamePattern.Enabled= $true
            $cmbMatchMode.Enabled  = $true
            $cmbType.Enabled       = $true
            $chkActiveOnly.Enabled = $true
            $progBar.Visible       = $false
            & $applyWinRMFilter
            return
        }

        $progBar.Maximum = $script:rawComputerList.Count
        & $applyWinRMFilter

        # Parallelisierung
        $script:runspacePool = [runspacefactory]::CreateRunspacePool(1, 25)
        $script:runspacePool.Open()
        $script:tasks.Clear()

        foreach ($comp in $script:rawComputerList) {
            $ps = [powershell]::Create()
            $ps.RunspacePool = $script:runspacePool
            [void]$ps.AddScript($workerBlock)
            [void]$ps.AddArgument($comp.ComputerName)
            [void]$ps.AddArgument($comp.DNSHostName)

            $asyncResult = $ps.BeginInvoke()
            $script:tasks.Add([PSCustomObject]@{
                PowerShell  = $ps
                AsyncResult = $asyncResult
                Name        = $comp.ComputerName
            })
        }

        $script:isScanning = $true
        $timer.Start()
    })

    $btnCancel.Add_Click({
        $timer.Stop()
        $script:isScanning = $false

        if ($script:tasks) {
            foreach ($t in $script:tasks) {
                try { $t.PowerShell.Dispose() } catch {}
            }
            $script:tasks.Clear()
        }

        if ($script:runspacePool) {
            $script:runspacePool.Close()
            $script:runspacePool.Dispose()
            $script:runspacePool = $null
        }

        $lblStats.Text = "Scan durch Benutzer abgebrochen."
        $btnScan.Enabled       = $true
        $btnCancel.Enabled     = $false
        $txtNamePattern.Enabled= $true
        $cmbMatchMode.Enabled  = $true
        $cmbType.Enabled       = $true
        $chkActiveOnly.Enabled = $true
        $progBar.Visible       = $false
        & $applyWinRMFilter
    })

    # CSV-Export
    $btnExport.Add_Click({
        if ($script:rawComputerList.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Keine Daten zum Exportieren vorhanden.", "Hinweis", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }

        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "CSV-Datei (*.csv)|*.csv"
        $sfd.FileName = "WinRM_Filtered_Audit_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

        if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                $script:rawComputerList | Export-Csv -Path $sfd.FileName -NoTypeInformation -Encoding UTF8 -Delimiter ";"
                [System.Windows.Forms.MessageBox]::Show("Export erfolgreich gespeichert unter:`n$($sfd.FileName)", "Erfolg", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Fehler beim Exportieren: $($_.Exception.Message)", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    })

    # Clean-up
    $form.Add_FormClosing({
        $timer.Stop()
        $timer.Dispose()
        if ($script:runspacePool) {
            try {
                $script:runspacePool.Close()
                $script:runspacePool.Dispose()
            } catch {}
        }
    })

    try {
        [void]$form.ShowDialog()
    } finally {
        if ($null -ne $form) { $form.Dispose() }
        [System.GC]::Collect()
    }
}

Show-Tool26-WinRMStatus
