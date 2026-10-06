# =========================================================================
# Tool15_Launcher.ps1 - Active Directory GPO Enterprise Suite
# =========================================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# -------------------------------------------------------------------------
# 1. Pfad-Erkennung & GitHub-Konfiguration
# -------------------------------------------------------------------------
$gitHubBaseUrl = "https://raw.githubusercontent.com/Mfi1979/AdminSuite/main/Tools/Tool15"

$isWebExecution = [string]::IsNullOrWhiteSpace($PSScriptRoot) -and [string]::IsNullOrWhiteSpace($MyInvocation.MyCommand.Path)

if ($isWebExecution) {
    $localTempBase = Join-Path $env:TEMP "Tool15_GPO_Suite"
    $modulePath    = Join-Path $localTempBase "Modules"
    if (-not (Test-Path $modulePath)) {
        New-Item -ItemType Directory -Path $modulePath -Force | Out-Null
    }
} else {
    $scriptDir  = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
    $modulePath = Join-Path $scriptDir "Modules"
    if (-not (Test-Path $modulePath)) {
        New-Item -ItemType Directory -Path $modulePath -Force | Out-Null
    }
}

# -------------------------------------------------------------------------
# 2. Module herunterladen (falls nötig) & im Skript-Scope dot-sourcen
# -------------------------------------------------------------------------
$moduleList = @(
    "common.ps1",
    "GpoParser.ps1",
    "Tab0_Dashboard.ps1",
    "Tab1_Overview.ps1",
    "Tab2_Settings.ps1",
    "Tab3_Backup.ps1",
    "Tab4_Compare.ps1",
    "Tab5_WMIFilter.ps1"
)

foreach ($mod in $moduleList) {
    $targetFile = Join-Path $modulePath $mod

    if ($isWebExecution -or (-not (Test-Path $targetFile))) {
        $rawUrl = "$gitHubBaseUrl/Modules/$mod"
        try {
            [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12
            Invoke-RestMethod -Uri $rawUrl -OutFile $targetFile -ErrorAction Stop
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Herunterladen von '$mod':`r`n$($_.Exception.Message)", "Download-Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            continue
        }
    }

    if (Test-Path $targetFile) {
        . $targetFile
    } else {
        [System.Windows.Forms.MessageBox]::Show("Modul '$mod' fehlt unter '$targetFile'.", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# -------------------------------------------------------------------------
# 3. Globale Caches initialisieren
# -------------------------------------------------------------------------
$script:isClosing         = $false
$script:rawOverviewList   = [System.Collections.Generic.List[PSCustomObject]]::new()
$script:rawSettingsList   = [System.Collections.Generic.List[PSCustomObject]]::new()
$script:rawBackupList     = [System.Collections.Generic.List[PSCustomObject]]::new()
$script:rawCompareList    = [System.Collections.Generic.List[PSCustomObject]]::new()
$script:rawWmiList        = [System.Collections.Generic.List[PSCustomObject]]::new()
$script:gpoLinksCache     = @{}
$script:allGposCache      = @()

# -------------------------------------------------------------------------
# 4. Domäne ermitteln
# -------------------------------------------------------------------------
$domainName = ""
$domainDN   = ""

try {
    $curDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
    $domainName = $curDomain.Name
    $domainDN   = ($curDomain.Name.Split('.') | ForEach-Object { "DC=$_" }) -join ','
} catch {
    try {
        $rootDSE = [ADSI]"LDAP://RootDSE"
        $domainDN = "$($rootDSE.defaultNamingContext)"
        $domainName = ($domainDN -replace 'DC=','' -replace ',','.')
    } catch {
        $domainName = "Lokal / Unbekannt"
        $domainDN = ""
    }
}

# -------------------------------------------------------------------------
# 5. Grid-Sortierung Hilfsfunktion
# -------------------------------------------------------------------------
function Enable-GridSorting {
    param([System.Windows.Forms.DataGridView]$Grid)

    $Grid.Add_ColumnHeaderMouseClick({
        param($sender, $e)
        
        $colIndex = $e.ColumnIndex
        if ($colIndex -lt 0 -or $colIndex -ge $sender.Columns.Count) { return }
        
        $propName = $sender.Columns[$colIndex].DataPropertyName
        if ([string]::IsNullOrWhiteSpace($propName)) { return }

        $dataSource = $sender.DataSource
        if ($null -eq $dataSource) { return }

        $currentGlyph = $sender.Columns[$colIndex].HeaderCell.SortGlyphDirection
        $direction = if ($currentGlyph -eq [System.Windows.Forms.SortOrder]::Ascending) {
            [System.ComponentModel.ListSortDirection]::Descending
        } else {
            [System.ComponentModel.ListSortDirection]::Ascending
        }

        $list = [System.Collections.ArrayList]::new()
        $sorted = if ($direction -eq [System.ComponentModel.ListSortDirection]::Ascending) {
            $dataSource | Sort-Object -Property @{ Expression = { $_.$propName } }
        } else {
            $dataSource | Sort-Object -Property @{ Expression = { $_.$propName } } -Descending
        }

        foreach ($item in $sorted) { [void]$list.Add($item) }
        $sender.DataSource = $list

        foreach ($c in $sender.Columns) {
            $c.HeaderCell.SortGlyphDirection = [System.Windows.Forms.SortOrder]::None
        }
        
        $targetCol = $sender.Columns[$colIndex]
        if ($null -ne $targetCol) {
            $targetCol.HeaderCell.SortGlyphDirection = if ($direction -eq [System.ComponentModel.ListSortDirection]::Ascending) {
                [System.Windows.Forms.SortOrder]::Ascending
            } else {
                [System.Windows.Forms.SortOrder]::Descending
            }
        }
    })
}

# -------------------------------------------------------------------------
# 6. Hauptfenster initialisieren
# -------------------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Tool 15 - Active Directory GPO Enterprise Suite ($domainName) - v2.0.0"
$form.Size = New-Object System.Drawing.Size(1280, 800)
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$form.MinimumSize = New-Object System.Drawing.Size(1024, 650)
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
$form.BackColor = [System.Drawing.Color]::FromArgb(246, 248, 252)

# Statuszeile & Fortschrittsbalken
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusStrip.Height = 28
$statusStrip.BackColor = [System.Drawing.Color]::FromArgb(238, 242, 248)

$script:lblProgressInfo = New-Object System.Windows.Forms.ToolStripStatusLabel
$script:lblProgressInfo.Text = "Bereit."
$script:lblProgressInfo.Spring = $true
$script:lblProgressInfo.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft

$script:pbarGlobal = New-Object System.Windows.Forms.ToolStripProgressBar
$script:pbarGlobal.Size = New-Object System.Drawing.Size(250, 18)
$script:pbarGlobal.Visible = $false

[void]$statusStrip.Items.Add($script:lblProgressInfo)
[void]$statusStrip.Items.Add($script:pbarGlobal)
$form.Controls.Add($statusStrip)

# TabControl
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = [System.Windows.Forms.DockStyle]::Fill
$tabControl.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($tabControl)

# -------------------------------------------------------------------------
# 7. UI-Tabs aufbauen
# -------------------------------------------------------------------------
Build-Tab0_Dashboard -tabControl $tabControl -domainDN $domainDN -domainName $domainName
Build-Tab1_Overview  -tabControl $tabControl -domainDN $domainDN -domainName $domainName
Build-Tab2_Settings  -tabControl $tabControl
Build-Tab3_Backup    -tabControl $tabControl
Build-Tab4_Compare   -tabControl $tabControl
Build-Tab5_WmiFilter -tabControl $tabControl -domainDN $domainDN -domainName $domainName

# -------------------------------------------------------------------------
# 8. Vollautomatischer Start-Load nach dem Anzeigen des Fensters
# -------------------------------------------------------------------------
function Set-ProgressSafe ([int]$pct, [string]$msg) {
    if ($script:pbarGlobal) {
        $script:pbarGlobal.Visible = $true
        $script:pbarGlobal.Minimum = 0
        $script:pbarGlobal.Maximum = 100
        $script:pbarGlobal.Value = [Math]::Min(100, [Math]::Max(0, $pct))
    }
    if ($script:lblProgressInfo -and -not [string]::IsNullOrWhiteSpace($msg)) {
        $script:lblProgressInfo.Text = $msg
    }
    [System.Windows.Forms.Application]::DoEvents()
}

$form.Add_Shown({
    foreach ($tab in $tabControl.TabPages) {
        $null = $tab.Handle
    }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

    try {
        # 1. TAB 1: GPOs via LDAP einlesen (befüllt $rawOverviewList und $allGposCache)
        Set-ProgressSafe 10 "Lese GPOs der Domäne via LDAP ein (Schritt 1/6)..."
        if ($script:Invoke_LoadOverview -is [scriptblock]) {
            & $script:Invoke_LoadOverview
        }

        # 2. TAB 2 (Inspektor): GPO-Dropdown befüllen & erste GPO laden
        Set-ProgressSafe 35 "Befülle Tab 2 (Inspektor) mit GPOs (Schritt 2/6)..."
        if ($script:Update_SettingsGpoDropdown -is [scriptblock]) {
            & $script:Update_SettingsGpoDropdown
        }
        if ($script:Invoke_LoadSettings -is [scriptblock]) {
            & $script:Invoke_LoadSettings
        }


		# 3. TAB 3 (Backup): GPO-Bestand direkt laden
        Set-ProgressSafe 55 "Befülle Tab 3 (Backup) mit GPOs (Schritt 3/6)..."
        if ($script:Invoke_LoadGposBackup -is [scriptblock]) {
            & $script:Invoke_LoadGposBackup
        }	


        # 4. TAB 4 (Diff): Dropdowns comboCompareGpo1 und comboCompareGpo2 füllen
        Set-ProgressSafe 70 "Synchronisiere Tab 4 (Vergleich) (Schritt 4/6)..."
        if ($script:comboCompareGpo1 -and -not $script:comboCompareGpo1.IsDisposed) {
            $script:comboCompareGpo1.Items.Clear()
            $script:comboCompareGpo2.Items.Clear()
            [void]$script:comboCompareGpo2.Items.Add($script:DdpBaselineName)

            foreach ($g in $script:allGposCache) {
                [void]$script:comboCompareGpo1.Items.Add($g.DisplayName)
                [void]$script:comboCompareGpo2.Items.Add($g.DisplayName)
            }
            if ($script:comboCompareGpo1.Items.Count -gt 0) { $script:comboCompareGpo1.SelectedIndex = 0 }
            if ($script:comboCompareGpo2.Items.Count -gt 0) { $script:comboCompareGpo2.SelectedIndex = 0 }
        }

        # 5. TAB 5 (WMI-Filter): Filter-Analyse laden
        Set-ProgressSafe 85 "Lese Tab 5 (WMI-Filter) ein (Schritt 5/6)..."
        if ($script:Invoke_LoadWmiFilters -is [scriptblock]) {
            & $script:Invoke_LoadWmiFilters
        }

        # 6. TAB 0 (Dashboard): KPIs und die 4 System-Checks aktualisieren
        Set-ProgressSafe 95 "Aktualisiere Dashboard & System-Checks (Schritt 6/6)..."
        if ($script:Invoke_LoadDashboard -is [scriptblock]) {
            & $script:Invoke_LoadDashboard
        }

        Set-ProgressSafe 100 "Bereit. $($script:rawOverviewList.Count) GPOs auf allen Tabs geladen."
    } catch {
        if ($script:lblProgressInfo) {
            $script:lblProgressInfo.Text = "Fehler beim Start-Laden: $($_.Exception.Message)"
        }
    } finally {
        if ($script:pbarGlobal) { $script:pbarGlobal.Visible = $false }
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

# -------------------------------------------------------------------------
# 9. Schliess- & Bereinigungslogik
# -------------------------------------------------------------------------
$form.Add_FormClosing({
    param($sender, $e)
    $script:isClosing = $true
})

try {
    [void]$form.ShowDialog()
}
finally {
    if ($form -and -not $form.IsDisposed) {
        $form.Dispose()
    }

    if ($script:rawOverviewList) { $script:rawOverviewList.Clear() }
    if ($script:rawSettingsList) { $script:rawSettingsList.Clear() }
    if ($script:rawBackupList)   { $script:rawBackupList.Clear() }
    if ($script:rawCompareList)  { $script:rawCompareList.Clear() }
    if ($script:rawWmiList)      { $script:rawWmiList.Clear() }
    if ($script:gpoLinksCache)   { $script:gpoLinksCache.Clear() }

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
