# =========================================================================
# Tool15_Launcher.ps1 - Active Directory GPO Enterprise Suite
# =========================================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# 1. Pfade und Modulverzeichnis festlegen
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$modulePath = Join-Path $scriptDir "Modules"

# 2. Globale Script-Variablen und Caches initialisieren
$script:isClosing         = $false
$script:rawOverviewList   = [System.Collections.Generic.List[PSCustomObject]]::new()
$script:rawSettingsList   = [System.Collections.Generic.List[PSCustomObject]]::new()
$script:rawBackupList     = [System.Collections.Generic.List[PSCustomObject]]::new()
$script:rawCompareList    = [System.Collections.Generic.List[PSCustomObject]]::new()
$script:gpoLinksCache     = @{}
$script:allGposCache      = @()

# 3. Domaenen-Informationen robust ermitteln
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
# 4. Hilfsfunktionen fuer Sortierung
# -------------------------------------------------------------------------
function Enable-GridSorting {
    param([System.Windows.Forms.DataGridView]$Grid)

    $Grid.Add_ColumnHeaderMouseClick({
        param($sender, $e)
        $column = $sender.Columns[$e.ColumnIndex]
        $propName = $column.DataPropertyName
        if ([string]::IsNullOrWhiteSpace($propName)) { return }

        $dataSource = $sender.DataSource
        if ($null -eq $dataSource) { return }

        $direction = [System.ComponentModel.ListSortDirection]::Ascending
        if ($column.HeaderCell.SortGlyphDirection -eq [System.Windows.Forms.SortOrder]::Ascending) {
            $direction = [System.ComponentModel.ListSortDirection]::Descending
        }

        $list = [System.Collections.ArrayList]::new()
        $sorted = if ($direction -eq [System.ComponentModel.ListSortDirection]::Ascending) {
            $dataSource | Sort-Object -Property @{ Expression = { $_.$propName } }
        } else {
            $dataSource | Sort-Object -Property @{ Expression = { $_.$propName } } -Descending
        }

        foreach ($item in $sorted) { [void]$list.Add($item) }
        $sender.DataSource = $list

        foreach ($col in $sender.Columns) {
            $col.HeaderCell.SortGlyphDirection = [System.Windows.Forms.SortOrder]::None
        }
        $column.HeaderCell.SortGlyphDirection = if ($direction -eq [System.ComponentModel.ListSortDirection]::Ascending) {
            [System.Windows.Forms.SortOrder]::Ascending
        } else {
            [System.Windows.Forms.SortOrder]::Descending
        }
    })
}

# -------------------------------------------------------------------------
# 5. Module laden
# -------------------------------------------------------------------------
. (Join-Path $modulePath "GpoParser.ps1")
. (Join-Path $modulePath "Tab0_Dashboard.ps1")
. (Join-Path $modulePath "Tab1_Overview.ps1")
. (Join-Path $modulePath "Tab2_Settings.ps1")
. (Join-Path $modulePath "Tab3_Backup.ps1")
. (Join-Path $modulePath "Tab4_Compare.ps1")
. (Join-Path $modulePath "Tab5_WmiFilter.ps1")

# -------------------------------------------------------------------------
# 6. Hauptfenster & Statusleiste initialisieren
# -------------------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Tool 15 - Active Directory GPO Enterprise Suite ($domainName) - v1.8.9"
$form.Size = New-Object System.Drawing.Size(1280, 800)
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$form.MinimumSize = New-Object System.Drawing.Size(1024, 650)
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
$form.BackColor = [System.Drawing.Color]::FromArgb(246, 248, 252)

# Statuszeile (Bottom)
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusStrip.Height = 28
$statusStrip.BackColor = [System.Drawing.Color]::FromArgb(238, 242, 248)

$script:lblProgressInfo = New-Object System.Windows.Forms.ToolStripStatusLabel
$script:lblProgressInfo.Text = "Bereit."
$script:lblProgressInfo.Spring = $true
$script:lblProgressInfo.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft

$script:pbarGlobal = New-Object System.Windows.Forms.ToolStripProgressBar
$script:pbarGlobal.Size = New-Object System.Drawing.Size(200, 18)
$script:pbarGlobal.Visible = $false

[void]$statusStrip.Items.Add($script:lblProgressInfo)
[void]$statusStrip.Items.Add($script:pbarGlobal)
$form.Controls.Add($statusStrip)

# TabControl (Zentral)
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = [System.Windows.Forms.DockStyle]::Fill
$tabControl.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($tabControl)

# -------------------------------------------------------------------------
# 7. UI-Tabs erstellen
# -------------------------------------------------------------------------
Build-Tab0_Dashboard -tabControl $tabControl -domainDN $domainDN -domainName $domainName
Build-Tab1_Overview  -tabControl $tabControl -domainDN $domainDN -domainName $domainName
Build-Tab2_Settings  -tabControl $tabControl -domainDN $domainDN -domainName $domainName
Build-Tab3_Backup    -tabControl $tabControl -domainDN $domainDN -domainName $domainName
Build-Tab4_Compare   -tabControl $tabControl
Build-Tab5_WmiFilter -tabControl $tabControl -domainDN $domainDN -domainName $domainName

# -------------------------------------------------------------------------
# 8. Initiales Laden beim Programmstart (Typsicher abgesichert)
# -------------------------------------------------------------------------
if ($script:Invoke_LoadDashboard -is [scriptblock]) {
    & $script:Invoke_LoadDashboard
}

# -------------------------------------------------------------------------
# 9. Schliess- & Bereinigungslogik (Verhindert Hänger in der PowerShell ISE)
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