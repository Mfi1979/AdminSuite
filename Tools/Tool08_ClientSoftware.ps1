<#
================================================================================
 TOOL 8: CLIENT SOFTWARE & APP ANALYSE (WIN32 & STORE APPS)
================================================================================
#>
function Show-ClientSoftwareAnalysis {
    $subForm = New-Object System.Windows.Forms.Form
    $subForm.Text = "Tool 8 - Client Software & App Analyse ($env:COMPUTERNAME)"
    $subForm.Size = New-Object System.Drawing.Size(1100, 720)
    $subForm.MinimumSize = New-Object System.Drawing.Size(850, 500)
    $subForm.StartPosition = "CenterScreen"
    $subForm.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9)

    $pnlTop = New-Object System.Windows.Forms.Panel
    $pnlTop.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlTop.Height = 90
    $pnlTop.Padding = New-Object System.Windows.Forms.Padding(10)
    $subForm.Controls.Add($pnlTop)

    $lblSearch = New-Object System.Windows.Forms.Label
    $lblSearch.Text = "Suche / Search:"
    $lblSearch.Location = New-Object System.Drawing.Point(12, 15)
    $lblSearch.AutoSize = $true
    $pnlTop.Controls.Add($lblSearch)

    $txtSearch = New-Object System.Windows.Forms.TextBox
    $txtSearch.Location = New-Object System.Drawing.Point(115, 12)
    $txtSearch.Size = New-Object System.Drawing.Size(160, 23)
    $pnlTop.Controls.Add($txtSearch)

    $lblType = New-Object System.Windows.Forms.Label
    $lblType.Text = "Typ:"
    $lblType.Location = New-Object System.Drawing.Point(290, 15)
    $lblType.AutoSize = $true
    $pnlTop.Controls.Add($lblType)

    $cmbType = New-Object System.Windows.Forms.ComboBox
    $cmbType.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $cmbType.Location = New-Object System.Drawing.Point(325, 12)
    $cmbType.Size = New-Object System.Drawing.Size(120, 23)
    [void]$cmbType.Items.AddRange(@("Alle Typen", "Desktop-App", "Store-App"))
    $cmbType.SelectedIndex = 0
    $pnlTop.Controls.Add($cmbType)

    $lblLang = New-Object System.Windows.Forms.Label
    $lblLang.Text = "Sprache:"
    $lblLang.Location = New-Object System.Drawing.Point(460, 15)
    $lblLang.AutoSize = $true
    $pnlTop.Controls.Add($lblLang)

    $cmbLang = New-Object System.Windows.Forms.ComboBox
    $cmbLang.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $cmbLang.Location = New-Object System.Drawing.Point(520, 12)
    $cmbLang.Size = New-Object System.Drawing.Size(145, 23)
    [void]$cmbLang.Items.AddRange(@("Alle Sprachen", "Nur DE (Deutsch)", "Nur EN (Englisch)", "Neutral / Multilingual"))
    $cmbLang.SelectedIndex = 0
    $pnlTop.Controls.Add($cmbLang)

    $btnScan = New-Object System.Windows.Forms.Button
    $btnScan.Text = "Neu laden"
    $btnScan.Location = New-Object System.Drawing.Point(680, 10)
    $btnScan.Size = New-Object System.Drawing.Size(120, 28)
    $btnScan.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $btnScan.ForeColor = [System.Drawing.Color]::White
    $btnScan.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnScan.Font = New-Object System.Drawing.Font($subForm.Font.FontFamily, 9, [System.Drawing.FontStyle]::Bold)
    $pnlTop.Controls.Add($btnScan)

    $btnExport = New-Object System.Windows.Forms.Button
    $btnExport.Text = "CSV Export"
    $btnExport.Location = New-Object System.Drawing.Point(810, 10)
    $btnExport.Size = New-Object System.Drawing.Size(100, 28)
    $pnlTop.Controls.Add($btnExport)

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = New-Object System.Drawing.Point(12, 55)
    $lblStatus.Size = New-Object System.Drawing.Size(950, 22)
    $lblStatus.ForeColor = [System.Drawing.Color]::DarkSlateGray
    $lblStatus.Text = "Initialisiere..."
    $pnlTop.Controls.Add($lblStatus)

    $gridApps = New-Object System.Windows.Forms.DataGridView
    $gridApps.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridApps.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $gridApps.AllowUserToAddRows = $false
    $gridApps.ReadOnly = $true
    $gridApps.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $gridApps.MultiSelect = $false
    Apply-StandardGridTheme -Grid $gridApps -EnableAlternatingRowColor
    $subForm.Controls.Add($gridApps)

    $gridApps.BringToFront()
    $pnlTop.SendToBack()

    $global:allInstalledApps = @()

    $applyFilter = {
        if (-not $global:allInstalledApps) { return }
        $searchPattern = if ([string]::IsNullOrWhiteSpace($txtSearch.Text)) { "*" } else { "*$($txtSearch.Text.Trim())*" }
        $selectedType  = $cmbType.SelectedItem
        $selectedLang  = $cmbLang.SelectedItem

        $filtered = $global:allInstalledApps | Where-Object {
            $nameMatch = $_."App-Name" -like $searchPattern
            $typeMatch = ($selectedType -eq "Alle Typen") -or ($_."Typ" -eq $selectedType)
            $langMatch = $true
            if ($selectedLang -eq "Nur DE (Deutsch)") { $langMatch = $_."Sprache" -like "DE*" }
            elseif ($selectedLang -eq "Nur EN (Englisch)") { $langMatch = $_."Sprache" -like "EN*" }
            elseif ($selectedLang -eq "Neutral / Multilingual") { $langMatch = $_."Sprache" -like "Neutral*" }
            $nameMatch -and $typeMatch -and $langMatch
        }

        $arr = [System.Collections.ArrayList]::new()
        foreach ($item in $filtered) { [void]$arr.Add($item) }
        $gridApps.DataSource = $null
        $gridApps.DataSource = $arr
        $lblStatus.Text = "Gefiltert: $($arr.Count) von $($global:allInstalledApps.Count) Programmen & Apps angezeigt."
    }

    $btnScan.Add_Click({
        $lblStatus.Text = "Lese Registry und Store-Pakete aus..."
        $subForm.Refresh()
        try {
            $regPaths = @(
                "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
                "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
                "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
            )
            $win32Apps = Get-ItemProperty $regPaths -ErrorAction SilentlyContinue |
                Where-Object { $_.DisplayName -and $_.SystemComponent -ne 1 -and $_.ParentDisplayName -eq $null } |
                Select-Object @{N="App-Name"; E={$_.DisplayName}}, 
                              @{N="Version";  E={$_.DisplayVersion}}, 
                              @{N="Sprache";  E={ Convert-LcidToLanguageName $_.Language }}, 
                              @{N="Typ";      E={"Desktop-App"}}

            $storeApps = Get-AppxPackage -ErrorAction SilentlyContinue |
                Where-Object { -not $_.IsFramework -and $_.NonRemovable -ne $true -and $_.SignatureKind -eq "Store" } |
                Select-Object @{N="App-Name"; E={$_.Name}}, 
                              @{N="Version";  E={$_.Version}}, 
                              @{N="Sprache";  E={"Neutral / Multilingual"}}, 
                              @{N="Typ";      E={"Store-App"}}

            $global:allInstalledApps = @($win32Apps + $storeApps) | Sort-Object "App-Name" -Unique
            & $applyFilter
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Abrufen der Apps: $($_.Exception.Message)", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $lblStatus.Text = "Fehler beim Laden."
        }
    })

    $txtSearch.Add_TextChanged({ & $applyFilter })
    $cmbType.Add_SelectedIndexChanged({ & $applyFilter })
    $cmbLang.Add_SelectedIndexChanged({ & $applyFilter })

    $btnExport.Add_Click({
        if (-not $gridApps.DataSource -or $gridApps.Rows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Keine Daten zum Exportieren vorhanden.", "Export", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "CSV-Datei (*.csv)|*.csv"
        $sfd.FileName = "ClientApps_$($env:COMPUTERNAME)_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
        if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $gridApps.DataSource | Export-Csv -Path $sfd.FileName -NoTypeInformation -Encoding UTF8 -Delimiter ";"
            [System.Windows.Forms.MessageBox]::Show("Erfolgreich exportiert nach:`n$($sfd.FileName)", "Export Abgeschlossen", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    })

    $subForm.Add_Shown({
        $btnScan.PerformClick()
    })

    [void]$subForm.ShowDialog()
}
