<#
==================================================================================
 Tool 15: Active Directory GPO Enterprise Suite (Modular Launcher)
 Version: 1.8.8
==================================================================================
#>

$modulePath = Join-Path $PSScriptRoot "Modules"
. (Join-Path $modulePath "Common.ps1")

# --- Domaenenpruefung ---
try {
    $domainInfo = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
    $rootDse = [ADSI]"LDAP://RootDSE"
    $domainDN = $rootDse.defaultNamingContext.Value
    $domainName = $domainInfo.Name
    $rootDse.Dispose()
} catch {
    [System.Windows.Forms.MessageBox]::Show(
        "Active Directory Domaene nicht erreichbar oder Computer nicht domaenengebunden.",
        "Fehler",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    return
}

$script:isClosing = $false

# --- Hauptfenster & UI-Container initialisieren ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Tool 15 - Active Directory GPO Enterprise Suite ($domainName) - $script:ToolVersion"
$form.Size = New-Object System.Drawing.Size(1680, 960)
$form.StartPosition = "CenterScreen"
$form.MinimumSize = New-Object System.Drawing.Size(1250, 750)
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# --- Statusleiste unten ---
$panelBottomStatus = New-Object System.Windows.Forms.Panel
$panelBottomStatus.Dock = [System.Windows.Forms.DockStyle]::Bottom
$panelBottomStatus.Height = 32
$panelBottomStatus.BackColor = [System.Drawing.Color]::FromArgb(240, 242, 246)
$panelBottomStatus.Padding = New-Object System.Windows.Forms.Padding(10, 4, 10, 4)

$lblProgressInfo = New-Object System.Windows.Forms.Label
$lblProgressInfo.Dock = [System.Windows.Forms.DockStyle]::Fill
$lblProgressInfo.Text = "Bereit."
$lblProgressInfo.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$lblProgressInfo.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$script:lblProgressInfo = $lblProgressInfo

$pbarGlobal = New-Object System.Windows.Forms.ProgressBar
$pbarGlobal.Dock = [System.Windows.Forms.DockStyle]::Right
$pbarGlobal.Width = 360
$pbarGlobal.Visible = $false
$script:pbarGlobal = $pbarGlobal

$panelBottomStatus.Controls.Add($lblProgressInfo)
$panelBottomStatus.Controls.Add($pbarGlobal)

# --- TabControl ---
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = [System.Windows.Forms.DockStyle]::Fill
$tabControl.Font = New-Object System.Drawing.Font("Segoe UI", 10.5, [System.Drawing.FontStyle]::Bold)
$tabControl.Padding = New-Object System.Drawing.Point(14, 6)

$form.Controls.Add($tabControl)
$form.Controls.Add($panelBottomStatus)
$panelBottomStatus.SendToBack()
$tabControl.BringToFront()

# Lokale Datencontainer
$script:rawOverviewList = [System.Collections.Generic.List[PSCustomObject]]::new()
$script:rawBackupList   = [System.Collections.Generic.List[PSCustomObject]]::new()
$script:gpoLinksCache   = @{}
$script:allGposCache    = [System.Collections.Generic.List[Microsoft.GroupPolicy.Gpo]]::new()
$script:rawSettingsList = [System.Collections.Generic.List[PSCustomObject]]::new()
$script:rawCompareList  = [System.Collections.Generic.List[PSCustomObject]]::new()
$script:dashDetailCache = @{}

# Module laden
. (Join-Path $modulePath "GpoParser.ps1")
. (Join-Path $modulePath "Tab0_Dashboard.ps1")
. (Join-Path $modulePath "Tab1_Overview.ps1")
. (Join-Path $modulePath "Tab2_Settings.ps1")
. (Join-Path $modulePath "Tab3_Backup.ps1")
. (Join-Path $modulePath "Tab4_Compare.ps1")

# Tabs assemblieren
Build-Tab0_Dashboard -tabControl $tabControl
Build-Tab1_Overview -tabControl $tabControl -domainDN $domainDN -domainName $domainName
Build-Tab2_Settings -tabControl $tabControl
Build-Tab3_Backup -tabControl $tabControl
Build-Tab4_Compare -tabControl $tabControl

# Initialer Startablauf
$form.Add_Shown({
    $script:allGposCache.Clear()
    $gpos = Get-GPO -All | Sort-Object DisplayName
    foreach ($g in $gpos) { [void]$script:allGposCache.Add($g) }

    # Tab 4 Dropdowns befuellen
    $script:comboCompareGpo1.Items.Clear()
    $script:comboCompareGpo2.Items.Clear()
    foreach ($g in $script:allGposCache) {
        [void]$script:comboCompareGpo1.Items.Add($g.DisplayName)
        [void]$script:comboCompareGpo2.Items.Add($g.DisplayName)
    }
    [void]$script:comboCompareGpo1.Items.Add($script:DdpBaselineName)
    [void]$script:comboCompareGpo2.Items.Add($script:DdpBaselineName)

    if ($script:comboCompareGpo1.Items.Count -gt 0) { $script:comboCompareGpo1.SelectedIndex = 0 }
    if ($script:comboCompareGpo2.Items.Count -gt 1) { $script:comboCompareGpo2.SelectedIndex = 1 }

    & $script:Invoke_LoadOverview
    & $script:Invoke_UpdateDashboard
    & $script:Update_SettingsGpoDropdown
    & $script:Invoke_LoadSettings
    & $script:Invoke_LoadGpos
})

$form.Add_FormClosing({ $script:isClosing = $true })

# 1. Event VOR dem Anzeigen registrieren (verhindert Hintergrund-Aufrufe beim Schlieﬂen)
$form.Add_FormClosing({
    param($sender, $e)
    $script:isClosing = $true
})

# 2. Fenster anzeigen und im finally-Block restlos bereinigen
try {
    [void]$form.ShowDialog()
}
finally {
    # Formular & Steuerelemente entsorgen
    if ($form -and -not $form.IsDisposed) {
        $form.Dispose()
    }

    # Interne Caches leeren, um Referenzen im ISE-Speicher zu kappen
    if ($script:rawOverviewList) { $script:rawOverviewList.Clear() }
    if ($script:rawSettingsList) { $script:rawSettingsList.Clear() }
    if ($script:rawBackupList)   { $script:rawBackupList.Clear() }
    if ($script:rawCompareList)  { $script:rawCompareList.Clear() }
    if ($script:gpoLinksCache)   { $script:gpoLinksCache.Clear() }

    # Garbage Collection anstoﬂen, um offene LDAP-/COM-Handles sofort freizugeben
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}