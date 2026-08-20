<#
================================================================================
 ACTIVE DIRECTORY & ENTRA ID ADMIN SUITE (VOLLSTÄNDIGER LETZTSTAND)
 Module: Tools 1 bis 9 & 11 | Sprachunterstützung: DE / EN | Native LDAP & Forms
================================================================================
#>

# Windows Forms & Drawing laden
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ==============================================================================
# 0. MULTILANGUAGE DICTIONARY (I18N: DE / EN)
# ==============================================================================
$script:CurrentLang = "DE" # Standard: DE

$script:I18N = @{
    "DE" = @{
        "Title"             = "Active Directory & Entra ID Admin Suite"
        "CategoryClient"    = "Kategorie: Client & Lokale Diagnose (Client Tools)"
        "CategoryAD"        = "Kategorie: Active Directory & Domänen Diagnose (AD Tools)"
        "Computer"          = "Computer"
        "Domain"            = "Domäne"
        "LogonServer"       = "Logonserver"
        "User"              = "Benutzer"
        "EntraStatus"       = "Entra Join Status"
        "Workgroup"         = "Arbeitsgruppe / Kein AD Join"
        "BtnTool1"          = "Tool 1: Multi-DC LastLogon Übersicht`n(Automatischer DC- & Site-Scan mit Checkbox-Auswahl & Live-Abfrage)"
        "BtnTool2"          = "Tool 2: Quick AD Audit`n(Inaktive Computerkonten, Passwort nie abgelaufen, Leere Gruppen)"
        "BtnTool3"          = "Tool 3: Entra ID / Hybrid Join Diagnostic`n(Auslesen von dsregcmd /status & Tenant Details)"
        "BtnTool4"          = "Tool 4: Gruppen & GPO Diagnostik`n(Register: User-Gruppen, PC-Gruppen & angewendete Richtlinien)"
        "BtnTool5"          = "Tool 5: Windows 11 Readiness & OS Details`n(Hardware-Kompatibilität, TPM 2.0, SecureBoot & Patch-Stand)"
        "BtnTool6"          = "Tool 6: Domänen-Übersicht & Admin-Audit`n(FSMO-Rollen, AD-Papierkorb, Objekt-Anzahl & Privilegierte Admins)"
        "BtnTool7"          = "Tool 7: AD Security & OS Support Audit`n(Identifikation veralteter & nicht mehr unterstützter Systeme)"
        "BtnTool8"          = "Tool 8: Client Software & App Analyse`n(Win32- & Store-Apps mit Sprachauswertung und CSV-Export)"
        "BtnTool9"          = "Tool 9: AD ACL & Berechtigungsvergleich`n(Objekt-Berechtigungen von Usern, Gruppen, Computern oder OUs vergleichen)"
        "BtnTool11"         = "Tool 11: AD Kennwortrichtlinien & PSO Audit`n(Default Domain Policy, Fine-Grained PSOs & Benutzer-Check)"
        "ErrNoDomain"       = "Dieses Werkzeug erfordert eine Active Directory Domänenmitgliedschaft."
    }
    "EN" = @{
        "Title"             = "Active Directory & Entra ID Admin Suite"
        "CategoryClient"    = "Category: Client & Local Diagnostics (Client Tools)"
        "CategoryAD"        = "Category: Active Directory & Domain Diagnostics (AD Tools)"
        "Computer"          = "Computer"
        "Domain"            = "Domain"
        "LogonServer"       = "Logon Server"
        "User"              = "User"
        "EntraStatus"       = "Entra Join Status"
        "Workgroup"         = "Workgroup / No AD Join"
        "BtnTool1"          = "Tool 1: Multi-DC LastLogon Overview`n(Automatic DC & Site Scan with Checkbox Selection & Live Query)"
        "BtnTool2"          = "Tool 2: Quick AD Audit`n(Disabled Accounts, Password Never Expires, Empty Groups)"
        "BtnTool3"          = "Tool 3: Entra ID / Hybrid Join Diagnostic`n(Read dsregcmd /status & Tenant Details)"
        "BtnTool4"          = "Tool 4: Groups & GPO Diagnostics`n(Tabs: User Groups, PC Groups & Applied GPOs)"
        "BtnTool5"          = "Tool 5: Windows 11 Readiness & OS Details`n(Hardware Compatibility, TPM 2.0, SecureBoot & Patch Level)"
        "BtnTool6"          = "Tool 6: Domain Overview & Admin Audit`n(FSMO Roles, AD Recycle Bin, Object Counts & Privileged Admins)"
        "BtnTool7"          = "Tool 7: AD Security & Out-of-Support OS Audit`n(Identify Legacy & End-of-Life Windows Systems in AD)"
        "BtnTool8"          = "Tool 8: Client Software & App Analysis`n(Win32 Registry & Store Apps with Language Detection & Export)"
        "BtnTool9"          = "Tool 9: AD ACL & Permission Diff Tool`n(Compare object permissions of Users, Groups, Computers or OUs)"
        "BtnTool11"         = "Tool 11: AD Password Policies & PSO Audit`n(Default Domain Policy, Fine-Grained PSOs & User Check)"
        "ErrNoDomain"       = "This tool requires Active Directory Domain Membership."
    }
}

function Get-Text([string]$key) {
    if ($script:I18N[$script:CurrentLang].ContainsKey($key)) {
        return $script:I18N[$script:CurrentLang][$key]
    }
    return $key
}

# ==============================================================================
# 1. GLOBALE HILFSFUNKTIONEN & LDAP-ENGINE
# ==============================================================================
function Assert-DomainJoined {
    $isDomain = (Get-CimInstance -ClassName Win32_ComputerSystem).PartOfDomain
    if (-not $isDomain) {
        [System.Windows.Forms.MessageBox]::Show((Get-Text "ErrNoDomain"), "Info", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return $false
    }
    return $true
}

# Native LDAP-Suche mit Server-Bindung (wichtig für DC-spezifische Werte wie lastLogon)
function Search-NativeLdap {
    param (
        [string]$LdapFilter = "(objectClass=*)",
        [string[]]$PropertiesToLoad = @("name"),
        [string]$Server = $null
    )
    try {
        $entry = if ($Server) { 
            New-Object System.DirectoryServices.DirectoryEntry("LDAP://$Server") 
        } else { 
            New-Object System.DirectoryServices.DirectoryEntry 
        }
        $searcher = New-Object System.DirectoryServices.DirectorySearcher($entry)
        $searcher.Filter = $LdapFilter
        $searcher.PageSize = 1000
        foreach ($prop in $PropertiesToLoad) { [void]$searcher.PropertiesToLoad.Add($prop) }
        return $searcher.FindAll()
    }
    catch {
        return @()
    }
}

# LCID zu Klarnamen-Übersetzer für Software-Audit
function Convert-LcidToLanguageName {
    param([object]$lcid)
    if ($null -eq $lcid -or $lcid -eq "" -or $lcid -eq 0) { return "Neutral / Multilingual" }
    try {
        $id = [int]$lcid
        switch ($id) {
            1031 { return "DE (de-DE)" }
            2055 { return "DE (de-CH)" }
            3079 { return "DE (de-AT)" }
            1033 { return "EN (en-US)" }
            2057 { return "EN (en-GB)" }
            4105 { return "EN (en-CA)" }
            3081 { return "EN (en-AU)" }
            1036 { return "FR (fr-FR)" }
            1034 { return "ES (es-ES)" }
            1040 { return "IT (it-IT)" }
            default {
                $cult = [System.Globalization.CultureInfo]::GetCultureInfo($id)
                return "$($cult.TwoLetterISOLanguageName.ToUpper()) ($($cult.Name))"
            }
        }
    }
    catch {
        return "Neutral / Multilingual"
    }
}

# Dynamische OS Support & EOL-Erkennungs-Engine für Tool 7
function Analyze-OSSupportDetails {
    param(
        [string]$OSName,
        [string]$OSVersion,
        [string]$OSServicePack
    )

    $status        = "Supported"
    $eolDateString = "Unbekannt"
    $clientVersion = "N/A"
    $buildNumber   = "Unbekannt"
    $targetDate    = $null

    # 1. Build-Nummer extrahieren
    if ($OSVersion -match "(\d{5})") {
        $buildNumber = $Matches[1]
    } elseif ($OSVersion -match "10\.0\.(\d+)") {
        $buildNumber = $Matches[1]
    } elseif ($OSVersion) {
        $buildNumber = $OSVersion
    }

    # 2. Feature Update & LTSC ermitteln
    $isLTSC = ($OSName -like "*LTSC*" -or $OSName -like "*LTSB*" -or $OSServicePack -like "*LTSC*" -or $OSServicePack -like "*LTSB*")

    if ($buildNumber -ne "Unbekannt") {
        $buildInt = 0
        [int]::TryParse($buildNumber, [ref]$buildInt) | Out-Null

        switch ($buildInt) {
            # Windows 11
            26200 { $clientVersion = "25H2" }
            26100 { 
                if ($OSName -like "*Server*") { $clientVersion = "Server 2025" }
                elseif ($isLTSC) { $clientVersion = "24H2 / LTSC 2024" }
                else { $clientVersion = "24H2" }
            }
            22631 { $clientVersion = "23H2" }
            22621 { $clientVersion = "22H2" }
            22000 { $clientVersion = "21H2" }

            # Windows 10
            19045 { $clientVersion = "22H2" }
            19044 { if ($isLTSC) { $clientVersion = "21H2 / LTSC 2021" } else { $clientVersion = "21H2" } }
            19043 { $clientVersion = "21H1" }
            19042 { $clientVersion = "20H2" }
            19041 { $clientVersion = "2004" }
            18363 { $clientVersion = "1909" }
            17763 { 
                if ($OSName -like "*Server*") { $clientVersion = "Server 2019" } 
                else { $clientVersion = if ($isLTSC) { "1809 / LTSC 2019" } else { "1809" } } 
            }
            14393 { 
                if ($OSName -like "*Server*") { $clientVersion = "Server 2016" } 
                else { $clientVersion = if ($isLTSC) { "1607 / LTSB 2016" } else { "1607" } } 
            }
            10240 { $clientVersion = "1507 / LTSB 2015" }

            # Server
            20348 { $clientVersion = "Server 2022" }

            default {
                if ($buildInt -gt 26200) { $clientVersion = "Insider Build" }
                elseif ($OSVersion) { $clientVersion = $OSVersion }
                else { $clientVersion = "Build $buildNumber" }
            }
        }
    } else {
        if ($OSServicePack) { $clientVersion = $OSServicePack }
        elseif ($OSVersion) { $clientVersion = $OSVersion }
    }

    $isEnterpriseOrEdu = ($OSName -like "*Enterprise*" -or $OSName -like "*Education*")
    $isLTSCOrLTSB     = ($isLTSC -or $clientVersion -like "*LTSC*" -or $clientVersion -like "*LTSB*")

    # 3. EOL-Datumsberechnung
    if ($OSName -like "*Windows 11*") {
        if ($clientVersion -like "*21H2*" -or $buildNumber -eq "22000") {
            $dateStr = if ($isEnterpriseOrEdu) { "08.10.2024" } else { "10.10.2023" }
            $eolDateString = $dateStr; $targetDate = [datetime]::ParseExact($dateStr, "dd.MM.yyyy", $null)
        } 
        elseif ($clientVersion -like "*22H2*" -or $buildNumber -eq "22621") {
            $dateStr = if ($isEnterpriseOrEdu) { "14.10.2025" } else { "08.10.2024" }
            $eolDateString = $dateStr; $targetDate = [datetime]::ParseExact($dateStr, "dd.MM.yyyy", $null)
        } 
        elseif ($clientVersion -like "*23H2*" -or $buildNumber -eq "22631") {
            $dateStr = if ($isEnterpriseOrEdu) { "10.11.2026" } else { "11.11.2025" }
            $eolDateString = $dateStr; $targetDate = [datetime]::ParseExact($dateStr, "dd.MM.yyyy", $null)
        } 
        elseif ($clientVersion -like "*24H2*" -or $buildNumber -eq "26100") {
            if ($isLTSCOrLTSB) {
                $dateStr = if ($OSName -like "*IoT*") { "10.10.2034" } else { "09.10.2029" }
            } else {
                $dateStr = if ($isEnterpriseOrEdu) { "12.10.2027" } else { "13.10.2026" }
            }
            $eolDateString = $dateStr; $targetDate = [datetime]::ParseExact($dateStr, "dd.MM.yyyy", $null)
        } 
        elseif ($clientVersion -like "*25H2*" -or $buildNumber -eq "26200") {
            $dateStr = if ($isEnterpriseOrEdu) { "10.10.2028" } else { "12.10.2027" }
            $eolDateString = $dateStr; $targetDate = [datetime]::ParseExact($dateStr, "dd.MM.yyyy", $null)
        } 
        else {
            $eolDateString = "Support aktiv"
        }
    }
    elseif ($OSName -like "*Windows 10*") {
        if ($isLTSCOrLTSB) {
            if ($buildNumber -eq "19044" -or $clientVersion -like "*2021*") {
                $dateStr = if ($OSName -like "*IoT*") { "13.01.2032" } else { "12.01.2027" }
                $eolDateString = $dateStr; $targetDate = [datetime]::ParseExact($dateStr, "dd.MM.yyyy", $null)
            }
            elseif ($buildNumber -eq "17763" -or $clientVersion -like "*2019*") {
                $eolDateString = "09.01.2029"; $targetDate = [datetime]::ParseExact("09.01.2029", "dd.MM.yyyy", $null)
            }
            elseif ($buildNumber -eq "14393" -or $clientVersion -like "*2016*") {
                $eolDateString = "13.10.2026"; $targetDate = [datetime]::ParseExact("13.10.2026", "dd.MM.yyyy", $null)
            }
            elseif ($buildNumber -eq "10240" -or $clientVersion -like "*2015*") {
                $eolDateString = "14.10.2025"; $targetDate = [datetime]::ParseExact("14.10.2025", "dd.MM.yyyy", $null)
            }
            else {
                $eolDateString = "12.01.2027"; $targetDate = [datetime]::ParseExact("12.01.2027", "dd.MM.yyyy", $null)
            }
        } else {
            $eolDateString = "14.10.2025"; $targetDate = [datetime]::ParseExact("14.10.2025", "dd.MM.yyyy", $null)
        }
    }
    elseif ($OSName -like "*Windows 7*" -or $OSName -like "*Windows 8*" -or $OSName -like "*Windows XP*" -or $OSName -like "*Windows Vista*") {
        $eolDateString = "Ausgelaufen"
        $status = "Out of Support / EOL"
    }
    elseif ($OSName -like "*Server 2003*" -or $OSName -like "*Server 2008*") {
        $eolDateString = "Ausgelaufen"
        $status = "Out of Support / EOL"
    }
    elseif ($OSName -like "*Server 2012*") {
        $eolDateString = "10.10.2023"; $targetDate = [datetime]::ParseExact("10.10.2023", "dd.MM.yyyy", $null)
    }
    elseif ($OSName -like "*Server 2016*") {
        $eolDateString = "12.01.2027"; $targetDate = [datetime]::ParseExact("12.01.2027", "dd.MM.yyyy", $null)
    }
    elseif ($OSName -like "*Server 2019*") {
        $eolDateString = "09.01.2029"; $targetDate = [datetime]::ParseExact("09.01.2029", "dd.MM.yyyy", $null)
    }
    elseif ($OSName -like "*Server 2022*") {
        $eolDateString = "14.10.2031"; $targetDate = [datetime]::ParseExact("14.10.2031", "dd.MM.yyyy", $null)
    }
    elseif ($OSName -like "*Server 2025*") {
        $eolDateString = "14.10.2034"; $targetDate = [datetime]::ParseExact("14.10.2034", "dd.MM.yyyy", $null)
    }

    # 4. Dynamischer Statusvergleich
    if ($targetDate) {
        $today = Get-Date
        if ($today -ge $targetDate) {
            $status = "Out of Support / EOL"
        } elseif ($today.AddMonths(12) -ge $targetDate) {
            $status = "Near EOL"
        } else {
            $status = "Supported"
        }
    }

    return [PSCustomObject]@{
        ClientVersion = $clientVersion
        BuildNumber   = $buildNumber
        Status        = $status
        EOLDate       = $eolDateString
    }
}

# Lokale Systemparameter auslesen
$localComputerName = $env:COMPUTERNAME
$localUserName     = $env:USERNAME
$cs = Get-CimInstance -ClassName Win32_ComputerSystem
$localDomainName   = if ($cs.PartOfDomain) { $cs.Domain } else { "WORKGROUP" }
$localLogonServer  = if ($env:LOGONSERVER) { $env:LOGONSERVER.TrimStart('\') } else { "Local" }

$localJoinStatus = "Nicht gekoppelt"
$localAzureAdPrt = "NO"
try {
    $dsreg = dsregcmd /status
    if ($dsreg -match "AzureAdJoined\s*:\s*YES") { $localJoinStatus = "Azure AD Joined" }
    if ($dsreg -match "EnterpriseJoined\s*:\s*YES") { $localJoinStatus = "Hybrid Joined" }
    if ($dsreg -match "AzureAdPrt\s*:\s*YES") { $localAzureAdPrt = "YES" }
} catch {}

$osInfo = Get-CimInstance -ClassName Win32_OperatingSystem
$osNameFormatted   = $osInfo.Caption
$osVersionRelease  = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name DisplayVersion -ErrorAction SilentlyContinue).DisplayVersion
$osBuildFormatted  = "$($osInfo.Version) ($($osInfo.BuildNumber))"
$osArchFormatted   = $osInfo.OSArchitecture
$osInstallDateForm = $osInfo.InstallDate.ToString("dd.MM.yyyy HH:mm")
$osPatchFormatted  = (Get-HotFix | Sort-Object InstalledOn -Descending | Select-Object -First 1).HotFixID

# ===================================================================
# MODULE TOOL 1: MULTI-DC LASTLOGON (INKL. SORTIER-AUSWAHL & SPALTEN-KLICK)
# ===================================================================
function Open-ToolLastLogon {
    if (-not (Assert-DomainJoined)) { return }

    $subForm = New-Object System.Windows.Forms.Form
    $subForm.Text = "Tool 1: Multi-DC LastLogon Analyse"
    $subForm.Size = New-Object System.Drawing.Size(1280, 800)
    $subForm.StartPosition = "CenterParent"
    $subForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $pnlTop = New-Object System.Windows.Forms.Panel
    $pnlTop.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlTop.Height = 85
    $subForm.Controls.Add($pnlTop)

    $lblSearch = New-Object System.Windows.Forms.Label
    $lblSearch.Location = New-Object System.Drawing.Point(15, 15)
    $lblSearch.Text = "Computer (Wildcard):"
    $lblSearch.Size = New-Object System.Drawing.Size(130, 20)
    $pnlTop.Controls.Add($lblSearch)

    $txtSearch = New-Object System.Windows.Forms.TextBox
    $txtSearch.Location = New-Object System.Drawing.Point(150, 12)
    $txtSearch.Size = New-Object System.Drawing.Size(150, 23)
    $txtSearch.Text = "*"
    $pnlTop.Controls.Add($txtSearch)

    $lblSort = New-Object System.Windows.Forms.Label
    $lblSort.Location = New-Object System.Drawing.Point(315, 15)
    $lblSort.Text = "Sortieren nach:"
    $lblSort.Size = New-Object System.Drawing.Size(95, 20)
    $pnlTop.Controls.Add($lblSort)

    $cmbSortBy = New-Object System.Windows.Forms.ComboBox
    $cmbSortBy.Location = New-Object System.Drawing.Point(415, 12)
    $cmbSortBy.Size = New-Object System.Drawing.Size(140, 23)
    $cmbSortBy.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$cmbSortBy.Items.AddRange(@("Computer", "Status", "Inaktiv seit (Tage)"))
    $cmbSortBy.SelectedIndex = 0
    $pnlTop.Controls.Add($cmbSortBy)

    $lblEnabled = New-Object System.Windows.Forms.Label
    $lblEnabled.Location = New-Object System.Drawing.Point(15, 48)
    $lblEnabled.Text = "Konto-Status:"
    $lblEnabled.Size = New-Object System.Drawing.Size(130, 20)
    $pnlTop.Controls.Add($lblEnabled)

    $cmbEnabledFilter = New-Object System.Windows.Forms.ComboBox
    $cmbEnabledFilter.Location = New-Object System.Drawing.Point(150, 45)
    $cmbEnabledFilter.Size = New-Object System.Drawing.Size(150, 23)
    $cmbEnabledFilter.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$cmbEnabledFilter.Items.AddRange(@("Alle Status", "Nur Aktivierte", "Nur Deaktivierte"))
    $cmbEnabledFilter.SelectedIndex = 0
    $pnlTop.Controls.Add($cmbEnabledFilter)

    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Location = New-Object System.Drawing.Point(575, 12)
    $btnRun.Size = New-Object System.Drawing.Size(120, 56)
    $btnRun.Text = "Abfragen"
    $btnRun.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $btnRun.ForeColor = [System.Drawing.Color]::White
    $btnRun.Font = New-Object System.Drawing.Font($subForm.Font.FontFamily, 9.5, [System.Drawing.FontStyle]::Bold)
    $pnlTop.Controls.Add($btnRun)

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = New-Object System.Drawing.Point(710, 12)
    $lblStatus.Size = New-Object System.Drawing.Size(540, 60)
    $lblStatus.ForeColor = [System.Drawing.Color]::DarkBlue
    $pnlTop.Controls.Add($lblStatus)

    $grpDCs = New-Object System.Windows.Forms.GroupBox
    $grpDCs.Text = "Verfügbare Domain Controller (Wählen Sie die abzufragenden DCs aus)"
    $grpDCs.Dock = [System.Windows.Forms.DockStyle]::Top
    $grpDCs.Height = 140
    $subForm.Controls.Add($grpDCs)

    $gridDCs = New-Object System.Windows.Forms.DataGridView
    $gridDCs.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridDCs.AllowUserToAddRows = $false
    $gridDCs.AllowUserToDeleteRows = $false
    $gridDCs.RowHeadersVisible = $false
    $gridDCs.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $gridDCs.BackgroundColor = [System.Drawing.Color]::White
    $grpDCs.Controls.Add($gridDCs)

    $colCheck = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $colCheck.Name = "Check"; $colCheck.HeaderText = "Abfragen"; $colCheck.Width = 70
    [void]$gridDCs.Columns.Add($colCheck)
    [void]$gridDCs.Columns.Add("HostName", "Domain Controller")
    [void]$gridDCs.Columns.Add("Site", "AD Site")

    $pnlLegend = New-Object System.Windows.Forms.Panel
    $pnlLegend.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlLegend.Height = 32
    $pnlLegend.BackColor = [System.Drawing.Color]::FromArgb(240, 243, 246)
    $subForm.Controls.Add($pnlLegend)

    $lblLegendTitle = New-Object System.Windows.Forms.Label
    $lblLegendTitle.Text = "Status-Legende:"
    $lblLegendTitle.Location = New-Object System.Drawing.Point(15, 7)
    $lblLegendTitle.AutoSize = $true
    $lblLegendTitle.Font = New-Object System.Drawing.Font($subForm.Font.FontFamily, 8.5, [System.Drawing.FontStyle]::Bold)
    $pnlLegend.Controls.Add($lblLegendTitle)

    $pnlGreenBox = New-Object System.Windows.Forms.Panel
    $pnlGreenBox.Location = New-Object System.Drawing.Point(130, 8); $pnlGreenBox.Size = New-Object System.Drawing.Size(16, 16)
    $pnlGreenBox.BackColor = [System.Drawing.Color]::FromArgb(232, 248, 232); $pnlGreenBox.BorderStyle = "FixedSingle"
    $pnlLegend.Controls.Add($pnlGreenBox)

    $lblGreenDesc = New-Object System.Windows.Forms.Label
    $lblGreenDesc.Text = "Login <= 90 Tage (Aktiv)"
    $lblGreenDesc.Location = New-Object System.Drawing.Point(152, 7); $lblGreenDesc.AutoSize = $true
    $pnlLegend.Controls.Add($lblGreenDesc)

    $pnlRedBox = New-Object System.Windows.Forms.Panel
    $pnlRedBox.Location = New-Object System.Drawing.Point(320, 8); $pnlRedBox.Size = New-Object System.Drawing.Size(16, 16)
    $pnlRedBox.BackColor = [System.Drawing.Color]::FromArgb(253, 232, 232); $pnlRedBox.BorderStyle = "FixedSingle"
    $pnlLegend.Controls.Add($pnlRedBox)

    $lblRedDesc = New-Object System.Windows.Forms.Label
    $lblRedDesc.Text = "Login > 90 Tage / Nie (Inaktiv)"
    $lblRedDesc.Location = New-Object System.Drawing.Point(342, 7); $lblRedDesc.AutoSize = $true
    $pnlLegend.Controls.Add($lblRedDesc)

    $lblDisabledDesc = New-Object System.Windows.Forms.Label
    $lblDisabledDesc.Text = "|   Status 'Deaktiviert' = AD-Konto deaktiviert  |  💡 Klick auf Spaltenkopf sortiert die Tabelle"
    $lblDisabledDesc.Location = New-Object System.Drawing.Point(530, 7); $lblDisabledDesc.AutoSize = $true
    $lblDisabledDesc.ForeColor = [System.Drawing.Color]::DarkRed
    $pnlLegend.Controls.Add($lblDisabledDesc)

    $gridResults = New-Object System.Windows.Forms.DataGridView
    $gridResults.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridResults.ReadOnly = $true
    $gridResults.AllowUserToAddRows = $false
    $gridResults.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
    $gridResults.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $gridResults.BackgroundColor = [System.Drawing.Color]::White
    $subForm.Controls.Add($gridResults)

    $gridResults.BringToFront()
    $pnlLegend.SendToBack()
    $grpDCs.SendToBack()
    $pnlTop.SendToBack()

    $applyRowColors = {
        $ninetyDaysAgo = (Get-Date).AddDays(-90)
        foreach ($row in $gridResults.Rows) {
            $latestVal = $row.Cells["NeuesterLogin"].Value
            $isWithin90Days = $false

            if ($latestVal -and $latestVal -ne "Nie") {
                try {
                    $parsedDate = [datetime]::ParseExact($latestVal, "dd.MM.yyyy HH:mm:ss", $null)
                    if ($parsedDate -ge $ninetyDaysAgo) {
                        $isWithin90Days = $true
                    }
                } catch {}
            }

            if ($isWithin90Days) {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(232, 248, 232)
            } else {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(253, 232, 232)
            }

            if ($row.Cells["Status"].Value -eq "Deaktiviert") {
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
                $row.Cells["Status"].Style.ForeColor = [System.Drawing.Color]::DarkRed
                $row.Cells["Status"].Style.Font = New-Object System.Drawing.Font($gridResults.Font, [System.Drawing.FontStyle]::Bold)
            } else {
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Black
            }
        }
    }

    $script:sortCol = ""
    $script:sortAsc = $true

    $gridResults.Add_ColumnHeaderMouseClick({
        param($sender, $e)

        $col = $gridResults.Columns[$e.ColumnIndex]
        $propName = if ($col.DataPropertyName) { $col.DataPropertyName } else { $col.HeaderText }

        if ($script:sortCol -eq $propName) {
            $script:sortAsc = -not $script:sortAsc
        } else {
            $script:sortCol = $propName
            $script:sortAsc = $true
        }

        $items = @($gridResults.DataSource)
        if ($null -eq $items -or $items.Count -le 1) { return }

        $sorted = $items | Sort-Object -Property @{
            Expression = {
                $val = $_.$propName
                if ($null -eq $val -or $val -eq "") { return "" }

                if ($propName -like "*Tage*" -or $propName -like "*Inaktiv*") {
                    if ($val -eq "Nie" -or $val -eq "N/A") { return [int]::MaxValue }
                    if ($val -as [int]) { return [int]$val }
                }

                if ($val -match '^\d{2}\.\d{2}\.\d{4}') {
                    try {
                        return [datetime]::ParseExact($val.ToString().Trim(), @("dd.MM.yyyy HH:mm:ss", "dd.MM.yyyy"), $null)
                    } catch { return $val }
                }

                return $val
            }
            Descending = (-not $script:sortAsc)
        }

        $newArr = [System.Collections.ArrayList]::new()
        foreach ($it in $sorted) { [void]$newArr.Add($it) }

        $gridResults.DataSource = $null
        $gridResults.DataSource = $newArr

        foreach ($c in $gridResults.Columns) {
            $c.HeaderCell.SortGlyphDirection = [System.Windows.Forms.SortOrder]::None
        }
        $gridResults.Columns[$e.ColumnIndex].HeaderCell.SortGlyphDirection = `
            $(if ($script:sortAsc) { [System.Windows.Forms.SortOrder]::Ascending } else { [System.Windows.Forms.SortOrder]::Descending })

        & $applyRowColors
    })

    $subForm.Add_Shown({
        $lblStatus.Text = "Lade Domain Controller & Sites..."
        $subForm.Refresh()
        try {
            $domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
            $dcs = $domain.DomainControllers
            foreach ($dc in $dcs) {
                [void]$gridDCs.Rows.Add($true, $dc.Name, $dc.SiteName)
            }
            $lblStatus.Text = "Bereit. $($dcs.Count) Domain Controller geladen."
        } catch {
            $lblStatus.Text = "Fehler beim Laden der Domain Controller."
        }
    })

    $btnRun.Add_Click({
        $selectedDCs = @()
        foreach ($row in $gridDCs.Rows) {
            if ($row.Cells["Check"].Value -eq $true) {
                $selectedDCs += $row.Cells["HostName"].Value
            }
        }

        if ($selectedDCs.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Bitte mindestens einen Domain Controller auswählen.", "Hinweis", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        $lblStatus.Text = "Suche Computer im Active Directory..."
        $subForm.Refresh()

        $namePattern = $txtSearch.Text -replace '\*', '*'
        $selectedEnabledIndex = $cmbEnabledFilter.SelectedIndex

        $ldapBase = "(&(objectCategory=computer)(name=$namePattern)"
        if ($selectedEnabledIndex -eq 1) {
            $ldapBase += "(!(userAccountControl:1.2.840.113556.1.4.803:=2))"
        } elseif ($selectedEnabledIndex -eq 2) {
            $ldapBase += "(userAccountControl:1.2.840.113556.1.4.803:=2)"
        }
        $ldapBase += ")"

        $found = Search-NativeLdap -LdapFilter $ldapBase -PropertiesToLoad @("name","distinguishedName","userAccountControl")

        if (-not $found -or $found.Count -eq 0) {
            $lblStatus.Text = "Keine Computer-Objekte mit den angegebenen Kriterien gefunden."
            $gridResults.DataSource = $null
            return
        }

        if ($found.Count -gt 30) {
            $msgResult = [System.Windows.Forms.MessageBox]::Show(
                "Es wurden $($found.Count) Computer gefunden.`n`nMulti-DC-Abfrage für $($selectedDCs.Count) DC(s) starten?",
                "Hinweis: Größere Treffermenge",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )
            if ($msgResult -ne [System.Windows.Forms.DialogResult]::Yes) {
                $lblStatus.Text = "Abfrage abgebrochen."
                return
            }
        }

        $results = @()
        $i = 0
        $today = Get-Date
        $ninetyDaysAgo = $today.AddDays(-90)
        $countActive90 = 0
        $countInactive90 = 0
        $countDisabled = 0

        foreach ($c in $found) {
            $i++
            $cName = $c.Properties["name"][0]
            $cDn   = $c.Properties["distinguishedname"][0]
            
            $isEnabled = $true
            if ($c.Properties["useraccountcontrol"].Count -gt 0) {
                $uac = [int]$c.Properties["useraccountcontrol"][0]
                if (($uac -band 2) -eq 2) { 
                    $isEnabled = $false 
                }
            }
            $statusText = if ($isEnabled) { "Aktiviert" } else { "Deaktiviert" }
            if (-not $isEnabled) { $countDisabled++ }

            $lblStatus.Text = "Frage DCs ab... ($i von $($found.Count)): $cName"
            [System.Windows.Forms.Application]::DoEvents()

            $latest = [datetime]::MinValue
            $bestDC = "N/A"
            $dcValues = [ordered]@{}

            foreach ($dc in $selectedDCs) {
                try {
                    $res = Search-NativeLdap -Server $dc -LdapFilter "(distinguishedName=$cDn)" -PropertiesToLoad @("lastLogon")
                    if ($res -and $res[0].Properties["lastlogon"].Count -gt 0 -and $res[0].Properties["lastlogon"][0] -gt 0) {
                        $dt = [DateTime]::FromFileTime($res[0].Properties["lastlogon"][0])
                        $dcValues[$dc] = $dt.ToString("dd.MM.yyyy HH:mm:ss")
                        if ($dt -gt $latest) {
                            $latest = $dt
                            $bestDC = $dc
                        }
                    } else {
                        $dcValues[$dc] = "Nie"
                    }
                } catch {
                    $dcValues[$dc] = "Fehler / Nicht erreichbar"
                }
            }

            $daysInactive = if ($latest -eq [datetime]::MinValue) { 
                "Nie" 
            } else { 
                [int]($today - $latest).TotalDays 
            }

            if ($latest -ne [datetime]::MinValue -and $latest -ge $ninetyDaysAgo) {
                $countActive90++
            } else {
                $countInactive90++
            }

            $map = [ordered]@{ 
                "Computer"            = $cName
                "Status"              = $statusText
                "Inaktiv seit (Tage)" = $daysInactive
                "NeuesterLogin"       = if ($latest -eq [datetime]::MinValue) { "Nie" } else { $latest.ToString("dd.MM.yyyy HH:mm:ss") }
                "NeuesterDC"          = $bestDC
            }

            foreach ($k in $dcValues.Keys) {
                $map[$k] = $dcValues[$k]
            }

            $results += [PSCustomObject]$map
        }

        $selectedSort = $cmbSortBy.SelectedItem.ToString()
        $sortedResults = switch ($selectedSort) {
            "Status" { 
                $results | Sort-Object "Status", "Computer" 
            }
            "Inaktiv seit (Tage)" { 
                $results | Sort-Object @{
                    Expression = { 
                        if ($_."Inaktiv seit (Tage)" -eq "Nie") { [int]::MaxValue } else { [int]$_."Inaktiv seit (Tage)" } 
                    }; 
                    Descending = $false 
                }, "Computer"
            }
            Default { 
                $results | Sort-Object "Computer" 
            }
        }

        $arr = [System.Collections.ArrayList]::new()
        foreach ($r in $sortedResults) { [void]$arr.Add($r) }
        $gridResults.DataSource = $null
        $gridResults.DataSource = $arr

        & $applyRowColors

        $lblStatus.Text = "Ergebnis: $($sortedResults.Count) Computer gefunden (🟢 Aktiv: $countActive90 | 🔴 Inaktiv (>90d): $countInactive90 | ⚪ Deaktiviert: $countDisabled)."
    })

    [void]$subForm.ShowDialog()
}        

# ==============================================================================
# TOOL 2: QUICK AD AUDIT & ABFRAGEN
# ==============================================================================
function Open-ToolADAudit {
    if (-not (Assert-DomainJoined)) { return }
    $subForm = New-Object System.Windows.Forms.Form
    $subForm.Text = "Tool 2: Quick AD Audit & Abfragen"
    $subForm.Size = New-Object System.Drawing.Size(1000, 600)
    $subForm.StartPosition = "CenterParent"

    $pnlTop = New-Object System.Windows.Forms.Panel
    $pnlTop.Dock = "Top"; $pnlTop.Height = 60; $subForm.Controls.Add($pnlTop)

    $cmbQuery = New-Object System.Windows.Forms.ComboBox
    $cmbQuery.DropDownStyle = "DropDownList"; $cmbQuery.Location = "15, 18"; $cmbQuery.Width = "300"
    [void]$cmbQuery.Items.AddRange(@("Deaktivierte Computerkonten", "Passwort läuft nie ab (User)", "Leere AD-Gruppen"))
    $cmbQuery.SelectedIndex = 0; $pnlTop.Controls.Add($cmbQuery)

    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Text = "Abfrage starten"; $btnRun.Location = "330, 16"; $btnRun.Size = "130, 28"; $pnlTop.Controls.Add($btnRun)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = "Fill"; $grid.ReadOnly = $true; $grid.AutoSizeColumnsMode = "Fill"; $subForm.Controls.Add($grid); $grid.BringToFront()

    $btnRun.Add_Click({
        $sel = $cmbQuery.SelectedIndex
        $list = @()
        if ($sel -eq 0) {
            $raw = Search-NativeLdap -LdapFilter "(&(objectCategory=computer)(userAccountControl:1.2.840.113556.1.4.803:=2))" -PropertiesToLoad @("name","distinguishedName")
            foreach ($i in $raw) { $list += [PSCustomObject]@{ "Name" = $i.Properties["name"][0]; "DN" = $i.Properties["distinguishedname"][0]; "Typ" = "Computer (Deaktiviert)" } }
        } elseif ($sel -eq 1) {
            $raw = Search-NativeLdap -LdapFilter "(&(objectCategory=person)(objectClass=user)(userAccountControl:1.2.840.113556.1.4.803:=65536))" -PropertiesToLoad @("name","sAMAccountName","distinguishedName")
            foreach ($i in $raw) { $list += [PSCustomObject]@{ "Name" = $i.Properties["name"][0]; "SAMAccount" = $i.Properties["samaccountname"][0]; "DN" = $i.Properties["distinguishedname"][0] } }
        } elseif ($sel -eq 2) {
            $raw = Search-NativeLdap -LdapFilter "(&(objectCategory=group)(!member=*))" -PropertiesToLoad @("name","distinguishedName")
            foreach ($i in $raw) { $list += [PSCustomObject]@{ "Gruppenname" = $i.Properties["name"][0]; "DN" = $i.Properties["distinguishedname"][0]; "Status" = "Keine Mitglieder" } }
        }
        $arr = [System.Collections.ArrayList]::new()
        foreach ($l in $list) { [void]$arr.Add($l) }
        $grid.DataSource = $arr
    })
    [void]$subForm.ShowDialog()
}

# ==============================================================================
# TOOL 3: ENTRA ID / HYBRID JOIN DIAGNOSTIC
# ==============================================================================
function Open-ToolEntraStatus {
    $subForm = New-Object System.Windows.Forms.Form
    $subForm.Text = "Tool 3: Entra ID / Hybrid Join Diagnostic"
    $subForm.Size = New-Object System.Drawing.Size(850, 600)
    $subForm.StartPosition = "CenterParent"

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Dock = "Fill"; $txt.Multiline = $true; $txt.ScrollBars = "Both"
    $txt.Font = New-Object System.Drawing.Font("Consolas", 9.5)
    $subForm.Controls.Add($txt)

    try { $txt.Text = (dsregcmd /status | Out-String) } catch { $txt.Text = "Fehler beim Ausführen von dsregcmd." }
    [void]$subForm.ShowDialog()
}

# ==============================================================================
# TOOL 4: GRUPPEN & GPO DIAGNOSTIK (3 REGISTER)
# ==============================================================================
function Open-ToolGroupsAndGPO {
    if (-not (Assert-DomainJoined)) { return }

    $subForm = New-Object System.Windows.Forms.Form
    $subForm.Text = "Tool 4: Gruppenmitgliedschaften & angewendete GPOs"
    $subForm.Size = New-Object System.Drawing.Size(1000, 700)
    $subForm.StartPosition = "CenterParent"
    $subForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Dock = [System.Windows.Forms.DockStyle]::Fill
    $subForm.Controls.Add($tabControl)

    # Register 1: User-Gruppen
    $tabUserGroup = New-Object System.Windows.Forms.TabPage
    $tabUserGroup.Text = "User-Gruppen ($localUserName)"
    $tabUserGroup.Padding = New-Object System.Windows.Forms.Padding(5)

    $gridUserGroups = New-Object System.Windows.Forms.DataGridView
    $gridUserGroups.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridUserGroups.ReadOnly = $true
    $gridUserGroups.AllowUserToAddRows = $false
    $gridUserGroups.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
    $gridUserGroups.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $gridUserGroups.BackgroundColor = [System.Drawing.Color]::White
    $tabUserGroup.Controls.Add($gridUserGroups)
    $tabControl.TabPages.Add($tabUserGroup)

    # Register 2: PC-Gruppen
    $tabPCGroup = New-Object System.Windows.Forms.TabPage
    $tabPCGroup.Text = "PC-Gruppen ($localComputerName)"
    $tabPCGroup.Padding = New-Object System.Windows.Forms.Padding(5)

    $gridPCGroups = New-Object System.Windows.Forms.DataGridView
    $gridPCGroups.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridPCGroups.ReadOnly = $true
    $gridPCGroups.AllowUserToAddRows = $false
    $gridPCGroups.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
    $gridPCGroups.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $gridPCGroups.BackgroundColor = [System.Drawing.Color]::White
    $tabPCGroup.Controls.Add($gridPCGroups)
    $tabControl.TabPages.Add($tabPCGroup)

    # Register 3: GPOs
    $tabGPOs = New-Object System.Windows.Forms.TabPage
    $tabGPOs.Text = "Angewendete GPOs (gpresult)"
    $tabGPOs.Padding = New-Object System.Windows.Forms.Padding(5)

    $txtGPO = New-Object System.Windows.Forms.TextBox
    $txtGPO.Dock = [System.Windows.Forms.DockStyle]::Fill
    $txtGPO.Multiline = $true
    $txtGPO.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
    $txtGPO.Font = New-Object System.Drawing.Font("Consolas", 9.5)
    $txtGPO.ReadOnly = $true
    $txtGPO.BackColor = [System.Drawing.Color]::White
    $tabGPOs.Controls.Add($txtGPO)
    $tabControl.TabPages.Add($tabGPOs)

    $subForm.Add_Shown({
        try {
            $uRes = Search-NativeLdap -LdapFilter "(&(objectCategory=person)(objectClass=user)(sAMAccountName=$localUserName))" -PropertiesToLoad @("memberOf")
            $uGroups = @()
            if ($uRes -and $uRes[0].Properties["memberof"]) {
                foreach ($dn in $uRes[0].Properties["memberof"]) {
                    $cn = ($dn -split ',')[0] -replace '^CN=', ''
                    $uGroups += [PSCustomObject]@{ "Gruppenname" = $cn; "DistinguishedName" = $dn }
                }
            }
            $arrU = [System.Collections.ArrayList]::new()
            foreach ($g in ($uGroups | Sort-Object Gruppenname)) { [void]$arrU.Add($g) }
            $gridUserGroups.DataSource = $arrU
        } catch { }

        try {
            $cRes = Search-NativeLdap -LdapFilter "(&(objectCategory=computer)(name=$localComputerName))" -PropertiesToLoad @("memberOf")
            $cGroups = @()
            if ($cRes -and $cRes[0].Properties["memberof"]) {
                foreach ($dn in $cRes[0].Properties["memberof"]) {
                    $cn = ($dn -split ',')[0] -replace '^CN=', ''
                    $cGroups += [PSCustomObject]@{ "Gruppenname" = $cn; "DistinguishedName" = $dn }
                }
            }
            $arrC = [System.Collections.ArrayList]::new()
            foreach ($g in ($cGroups | Sort-Object Gruppenname)) { [void]$arrC.Add($g) }
            $gridPCGroups.DataSource = $arrC
        } catch { }

        $txtGPO.Text = "Lade Gruppenrichtlinien via gpresult /r ..."
        $subForm.Refresh()
        try {
            $gpoOutput = gpresult /r 2>&1 | Out-String
            $txtGPO.Text = $gpoOutput
        } catch {
            $txtGPO.Text = "Fehler beim Ausführen von gpresult /r.`r`n$($_.Exception.Message)"
        }
    })

    [void]$subForm.ShowDialog()
}

# ==============================================================================
# TOOL 5: WINDOWS 11 READINESS & HARDWARE CHECK
# ==============================================================================
function Open-ToolWin11Check {
    $subForm = New-Object System.Windows.Forms.Form
    $subForm.Text = "Tool 5: Windows 11 Readiness & OS Details"
    $subForm.Size = New-Object System.Drawing.Size(850, 500)
    $subForm.StartPosition = "CenterParent"

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = "Fill"; $grid.ReadOnly = $true; $grid.AutoSizeColumnsMode = "Fill"; $subForm.Controls.Add($grid)

    $tpm = Get-CimInstance -Namespace "Root\CIMv2\Security\MicrosoftTpm" -ClassName Win32_Tpm -ErrorAction SilentlyContinue
    $tpmOk = ($tpm -and $tpm.SpecVersion -match "2.0")

    $sb = Confirm-SecureBootUEFI -ErrorAction SilentlyContinue
    $ramGB = [Math]::Round((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1GB, 1)
    $cpuCores = (Get-CimInstance Win32_Processor).NumberOfCores

    $checks = @(
        [PSCustomObject]@{ "Prüfkriterium" = "TPM 2.0"; "Status" = if ($tpmOk) { "Erfüllt (TPM 2.0 aktiv)" } else { "Nicht erfüllt / Deaktiviert" }; "Erforderlich" = "TPM 2.0" }
        [PSCustomObject]@{ "Prüfkriterium" = "Secure Boot"; "Status" = if ($sb) { "Aktiviert" } else { "Deaktiviert / Legacy" }; "Erforderlich" = "Aktiviert" }
        [PSCustomObject]@{ "Prüfkriterium" = "Arbeitsspeicher (RAM)"; "Status" = "$ramGB GB"; "Erforderlich" = ">= 4 GB" }
        [PSCustomObject]@{ "Prüfkriterium" = "CPU-Kerne"; "Status" = "$cpuCores Kerne"; "Erforderlich" = ">= 2 Kerne" }
        [PSCustomObject]@{ "Prüfkriterium" = "Installiertes OS"; "Status" = "$osNameFormatted (Build $osBuildFormatted)"; "Erforderlich" = "Windows 10/11" }
    )
    $arr = [System.Collections.ArrayList]::new()
    foreach ($c in $checks) { [void]$arr.Add($c) }
    $grid.DataSource = $arr
    [void]$subForm.ShowDialog()
}

# ==============================================================================
# TOOL 6: DOMAIN OVERVIEW & PRIVILEGED ADMIN AUDIT (TABS & DATAGRIDS)
# ==============================================================================
function Open-ToolDomainOverview {
    if (-not (Assert-DomainJoined)) { return }

    $subForm = New-Object System.Windows.Forms.Form
    $subForm.Text = "Tool 6: Domänen-Übersicht & Admin-Audit"
    $subForm.Size = New-Object System.Drawing.Size(1050, 720)
    $subForm.StartPosition = "CenterParent"
    $subForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $subForm.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Dock = [System.Windows.Forms.DockStyle]::Fill
    $tabControl.Padding = New-Object System.Drawing.Point(12, 6)
    $subForm.Controls.Add($tabControl)

    # TAB 1: Domäne & FSMO
    $tabDomain = New-Object System.Windows.Forms.TabPage
    $tabDomain.Text = "  🌐 Domäne & FSMO Rollen  "
    $tabDomain.BackColor = [System.Drawing.Color]::White
    $tabControl.TabPages.Add($tabDomain)

    $gridDomain = New-Object System.Windows.Forms.DataGridView
    $gridDomain.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridDomain.ReadOnly = $true
    $gridDomain.AllowUserToAddRows = $false
    $gridDomain.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $gridDomain.RowHeadersVisible = $false
    $gridDomain.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $gridDomain.BackgroundColor = [System.Drawing.Color]::White
    $gridDomain.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $gridDomain.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(248, 249, 250)
    $tabDomain.Controls.Add($gridDomain)

    # TAB 2: Admins
    $tabAdmins = New-Object System.Windows.Forms.TabPage
    $tabAdmins.Text = "  🛡️ Privilegierte Administratoren  "
    $tabAdmins.BackColor = [System.Drawing.Color]::White
    $tabControl.TabPages.Add($tabAdmins)

    $pnlAdminTop = New-Object System.Windows.Forms.Panel
    $pnlAdminTop.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlAdminTop.Height = 50
    $pnlAdminTop.BackColor = [System.Drawing.Color]::FromArgb(240, 243, 246)
    $tabAdmins.Controls.Add($pnlAdminTop)

    $lblAdminFilter = New-Object System.Windows.Forms.Label
    $lblAdminFilter.Text = "Filter / Suche:"
    $lblAdminFilter.Location = New-Object System.Drawing.Point(15, 16)
    $lblAdminFilter.AutoSize = $true
    $pnlAdminTop.Controls.Add($lblAdminFilter)

    $txtAdminFilter = New-Object System.Windows.Forms.TextBox
    $txtAdminFilter.Location = New-Object System.Drawing.Point(110, 13)
    $txtAdminFilter.Size = New-Object System.Drawing.Size(260, 25)
    $pnlAdminTop.Controls.Add($txtAdminFilter)

    $lblAdminCount = New-Object System.Windows.Forms.Label
    $lblAdminCount.Location = New-Object System.Drawing.Point(390, 16)
    $lblAdminCount.AutoSize = $true
    $lblAdminCount.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $pnlAdminTop.Controls.Add($lblAdminCount)

    $gridAdmins = New-Object System.Windows.Forms.DataGridView
    $gridAdmins.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridAdmins.ReadOnly = $true
    $gridAdmins.AllowUserToAddRows = $false
    $gridAdmins.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $gridAdmins.RowHeadersVisible = $false
    $gridAdmins.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $gridAdmins.BackgroundColor = [System.Drawing.Color]::White
    $gridAdmins.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $gridAdmins.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(248, 249, 250)
    $tabAdmins.Controls.Add($gridAdmins)
    $gridAdmins.BringToFront()

    # TAB 3: DCs
    $tabDCs = New-Object System.Windows.Forms.TabPage
    $tabDCs.Text = "  🖥️ Domänencontroller  "
    $tabDCs.BackColor = [System.Drawing.Color]::White
    $tabControl.TabPages.Add($tabDCs)

    $gridDCs = New-Object System.Windows.Forms.DataGridView
    $gridDCs.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridDCs.ReadOnly = $true
    $gridDCs.AllowUserToAddRows = $false
    $gridDCs.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $gridDCs.RowHeadersVisible = $false
    $gridDCs.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $gridDCs.BackgroundColor = [System.Drawing.Color]::White
    $gridDCs.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $gridDCs.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(248, 249, 250)
    $tabDCs.Controls.Add($gridDCs)

    $domainDetails = [System.Collections.Generic.List[PSCustomObject]]::new()
    $dcList = [System.Collections.Generic.List[PSCustomObject]]::new()

    try {
        $domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
        $forest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()

        $rbStatus = "Deaktiviert / Unbekannt"
        try {
            $rootDSE = New-Object System.DirectoryServices.DirectoryEntry("LDAP://RootDSE")
            $configDN = $rootDSE.Properties["configurationNamingContext"][0]
            $rbDN = "CN=Recycle Bin Feature,CN=Optional Features,CN=Directory Service,CN=Windows NT,CN=Services,$configDN"
            $rbEntry = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$rbDN")
            if ($rbEntry -and $rbEntry.Properties["msDS-EnabledFeatureBL"].Count -gt 0) {
                $rbStatus = "🟢 Aktiviert (Enabled)"
            }
        } catch { }

        $domainDetails.Add([PSCustomObject]@{"Kategorie" = "Allgemein"; "Eigenschaft" = "Domänenname (FQDN)"; "Wert" = $domain.Name })
        $domainDetails.Add([PSCustomObject]@{"Kategorie" = "Allgemein"; "Eigenschaft" = "Gesamtstruktur (Forest)"; "Wert" = $forest.Name })
        $domainDetails.Add([PSCustomObject]@{"Kategorie" = "Allgemein"; "Eigenschaft" = "Domänen-Funktionsebene"; "Wert" = $domain.DomainMode.ToString() })
        $domainDetails.Add([PSCustomObject]@{"Kategorie" = "Allgemein"; "Eigenschaft" = "Forest-Funktionsebene"; "Wert" = $forest.ForestMode.ToString() })
        $domainDetails.Add([PSCustomObject]@{"Kategorie" = "Allgemein"; "Eigenschaft" = "AD-Papierkorb (Recycle Bin)"; "Wert" = $rbStatus })
        $domainDetails.Add([PSCustomObject]@{"Kategorie" = "Allgemein"; "Eigenschaft" = "Anzahl Domänencontroller"; "Wert" = $domain.DomainControllers.Count.ToString() })

        $domainDetails.Add([PSCustomObject]@{"Kategorie" = "FSMO Rolle"; "Eigenschaft" = "PDC Emulator"; "Wert" = $domain.PdcRoleOwner.Name })
        $domainDetails.Add([PSCustomObject]@{"Kategorie" = "FSMO Rolle"; "Eigenschaft" = "RID Master"; "Wert" = $domain.RidRoleOwner.Name })
        $domainDetails.Add([PSCustomObject]@{"Kategorie" = "FSMO Rolle"; "Eigenschaft" = "Infrastructure Master"; "Wert" = $domain.InfrastructureRoleOwner.Name })
        $domainDetails.Add([PSCustomObject]@{"Kategorie" = "FSMO Rolle"; "Eigenschaft" = "Schema Master (Forest)"; "Wert" = $forest.SchemaRoleOwner.Name })
        $domainDetails.Add([PSCustomObject]@{"Kategorie" = "FSMO Rolle"; "Eigenschaft" = "Domain Naming Master (Forest)"; "Wert" = $forest.NamingRoleOwner.Name })

        foreach ($dc in $domain.DomainControllers) {
            $roles = @()
            if ($dc.Name -eq $domain.PdcRoleOwner.Name) { $roles += "PDC" }
            if ($dc.Name -eq $domain.RidRoleOwner.Name) { $roles += "RID" }
            if ($dc.Name -eq $domain.InfrastructureRoleOwner.Name) { $roles += "Infra" }
            if ($dc.Name -eq $forest.SchemaRoleOwner.Name) { $roles += "Schema" }
            if ($dc.Name -eq $forest.NamingRoleOwner.Name) { $roles += "Naming" }
            if ($roles.Count -eq 0) { $roles += "Standard DC" }

            $dcList.Add([PSCustomObject]@{
                "DC Hostname"     = $dc.Name
                "IP-Adresse"      = $dc.IPAddress
                "AD-Standort"     = $dc.SiteName
                "Global Catalog"  = if ($dc.IsGlobalCatalog()) { "Ja" } else { "Nein" }
                "Inhaber Rollen"  = ($roles -join ", ")
                "Betriebssystem"  = $dc.OSVersion
            })
        }
    }
    catch {
        $domainDetails.Add([PSCustomObject]@{"Kategorie" = "Fehler"; "Eigenschaft" = "Hinweis"; "Wert" = "Domänenabfrage fehlgeschlagen: $($_.Exception.Message)" })
    }

    $gridDomain.DataSource = [System.Collections.ArrayList]::new($domainDetails)
    $gridDCs.DataSource = [System.Collections.ArrayList]::new($dcList)

    $adminList = [System.Collections.Generic.List[PSCustomObject]]::new()
    $extractName = {
        param($dn)
        if ($dn -match "^CN=([^,]+)") { return $Matches[1] }
        return $dn
    }

    $daGroups = Search-NativeLdap -LdapFilter "(&(objectCategory=group)(|(sAMAccountName=Domain Admins)(sAMAccountName=Domänen-Admins)))" -PropertiesToLoad @("name","member","sAMAccountName")
    if ($daGroups -and $daGroups.Count -gt 0) {
        $gName = $daGroups[0].Properties["name"][0]
        if ($daGroups[0].Properties["member"]) {
            foreach ($m in $daGroups[0].Properties["member"]) {
                $adminList.Add([PSCustomObject]@{
                    "Admin-Gruppe"      = $gName
                    "Account Name"      = (& $extractName $m)
                    "Mitgliedschaft"    = "Explizit (member)"
                    "DistinguishedName" = $m
                })
            }
        }
    }

    $primaryAdmins = Search-NativeLdap -LdapFilter "(&(objectCategory=person)(objectClass=user)(primaryGroupID=512))" -PropertiesToLoad @("sAMAccountName","distinguishedName")
    foreach ($pa in $primaryAdmins) {
        $dn = $pa.Properties["distinguishedname"][0]
        $acc = if ($pa.Properties["samaccountname"].Count -gt 0) { $pa.Properties["samaccountname"][0] } else { (& $extractName $dn) }
        
        if (-not ($adminList | Where-Object { $_.DistinguishedName -eq $dn })) {
            $adminList.Add([PSCustomObject]@{
                "Admin-Gruppe"      = "Domänen-Admins (Primary)"
                "Account Name"      = $acc
                "Mitgliedschaft"    = "Primäre Gruppe (ID 512)"
                "DistinguishedName" = $dn
            })
        }
    }

    $eaGroups = Search-NativeLdap -LdapFilter "(&(objectCategory=group)(|(sAMAccountName=Enterprise Admins)(sAMAccountName=Organisations-Admins)))" -PropertiesToLoad @("name","member")
    if ($eaGroups -and $eaGroups.Count -gt 0 -and $eaGroups[0].Properties["member"]) {
        $eaName = $eaGroups[0].Properties["name"][0]
        foreach ($m in $eaGroups[0].Properties["member"]) {
            $adminList.Add([PSCustomObject]@{
                "Admin-Gruppe"      = $eaName
                "Account Name"      = (& $extractName $m)
                "Mitgliedschaft"    = "Explizit (member)"
                "DistinguishedName" = $m
            })
        }
    }

    $saGroups = Search-NativeLdap -LdapFilter "(&(objectCategory=group)(|(sAMAccountName=Schema Admins)(sAMAccountName=Schema-Admins)))" -PropertiesToLoad @("name","member")
    if ($saGroups -and $saGroups.Count -gt 0 -and $saGroups[0].Properties["member"]) {
        $saName = $saGroups[0].Properties["name"][0]
        foreach ($m in $saGroups[0].Properties["member"]) {
            $adminList.Add([PSCustomObject]@{
                "Admin-Gruppe"      = $saName
                "Account Name"      = (& $extractName $m)
                "Mitgliedschaft"    = "Explizit (member)"
                "DistinguishedName" = $m
            })
        }
    }

    $arrAdmins = [System.Collections.ArrayList]::new($adminList)
    $gridAdmins.DataSource = $arrAdmins
    $lblAdminCount.Text = "Gefundene privilegierte Konten: $($adminList.Count)"

    $txtAdminFilter.Add_TextChanged({
        $filterText = $txtAdminFilter.Text.Trim().ToLower()
        if ([string]::IsNullOrWhiteSpace($filterText)) {
            $gridAdmins.DataSource = [System.Collections.ArrayList]::new($adminList)
            $lblAdminCount.Text = "Gefundene privilegierte Konten: $($adminList.Count)"
        } else {
            $filtered = $adminList | Where-Object {
                $_."Admin-Gruppe".ToLower().Contains($filterText) -or
                $_."Account Name".ToLower().Contains($filterText) -or
                $_."DistinguishedName".ToLower().Contains($filterText)
            }
            $arrF = [System.Collections.ArrayList]::new()
            foreach ($item in $filtered) { [void]$arrF.Add($item) }
            $gridAdmins.DataSource = $arrF
            $lblAdminCount.Text = "Gefiltert: $($arrF.Count) von $($adminList.Count)"
        }
    })

    [void]$subForm.ShowDialog()
}

# ==============================================================================
# TOOL 7: AD SECURITY & OS SUPPORT AUDIT (LETZTSTAND)
# ==============================================================================
function Open-ToolOSSupportAudit {
    if (-not (Assert-DomainJoined)) { return }

    $subForm = New-Object System.Windows.Forms.Form
    $subForm.Text = "Tool 7: AD Security & OS Support Audit"
    $subForm.Size = New-Object System.Drawing.Size(1420, 860)
    $subForm.StartPosition = "CenterParent"
    $subForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $pnlTop = New-Object System.Windows.Forms.Panel
    $pnlTop.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlTop.Height = 110
    $subForm.Controls.Add($pnlTop)

    $lblType = New-Object System.Windows.Forms.Label
    $lblType.Location = "15, 15"; $lblType.Size = "80, 20"; $lblType.Text = "Systemtyp:"
    $pnlTop.Controls.Add($lblType)

    $cmbSystemType = New-Object System.Windows.Forms.ComboBox
    $cmbSystemType.Location = "100, 12"; $cmbSystemType.Size = "130, 23"
    $cmbSystemType.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$cmbSystemType.Items.AddRange(@("All / Alle", "Clients Only", "Servers Only"))
    $cmbSystemType.SelectedIndex = 0
    $pnlTop.Controls.Add($cmbSystemType)

    $lblOSFilter = New-Object System.Windows.Forms.Label
    $lblOSFilter.Location = "245, 15"; $lblOSFilter.Size = "100, 20"; $lblOSFilter.Text = "Betriebssystem:"
    $pnlTop.Controls.Add($lblOSFilter)

    $cmbOSFilter = New-Object System.Windows.Forms.ComboBox
    $cmbOSFilter.Location = "345, 12"; $cmbOSFilter.Size = "220, 23"
    $cmbOSFilter.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$cmbOSFilter.Items.Add("All OS / Alle OS")
    $cmbOSFilter.SelectedIndex = 0
    $pnlTop.Controls.Add($cmbOSFilter)

    $lblName = New-Object System.Windows.Forms.Label
    $lblName.Location = "15, 48"; $lblName.Size = "80, 20"; $lblName.Text = "Computer:"
    $pnlTop.Controls.Add($lblName)

    $txtNameSearch = New-Object System.Windows.Forms.TextBox
    $txtNameSearch.Location = "100, 45"; $txtNameSearch.Size = "130, 23"; $txtNameSearch.Text = "*"
    $pnlTop.Controls.Add($txtNameSearch)

    $lblEnabled = New-Object System.Windows.Forms.Label
    $lblEnabled.Location = "245, 48"; $lblEnabled.Size = "100, 20"; $lblEnabled.Text = "Konto-Status:"
    $pnlTop.Controls.Add($lblEnabled)

    $cmbEnabledFilter = New-Object System.Windows.Forms.ComboBox
    $cmbEnabledFilter.Location = "345, 45"; $cmbEnabledFilter.Size = "220, 23"
    $cmbEnabledFilter.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$cmbEnabledFilter.Items.AddRange(@("Alle Status", "Nur Aktivierte (True)", "Nur Deaktivierte (False)"))
    $cmbEnabledFilter.SelectedIndex = 0
    $pnlTop.Controls.Add($cmbEnabledFilter)

    $lblStatusGroup = New-Object System.Windows.Forms.Label
    $lblStatusGroup.Location = "585, 15"; $lblStatusGroup.Size = "100, 20"; $lblStatusGroup.Text = "Support-Status:"
    $pnlTop.Controls.Add($lblStatusGroup)

    $chkListStatus = New-Object System.Windows.Forms.CheckedListBox
    $chkListStatus.Location = "685, 12"; $chkListStatus.Size = "180, 80"; $chkListStatus.CheckOnClick = $true
    [void]$chkListStatus.Items.Add("Supported", $true)
    [void]$chkListStatus.Items.Add("Near EOL", $true)
    [void]$chkListStatus.Items.Add("Out of Support / EOL", $true)
    [void]$chkListStatus.Items.Add("Unbekannt / Sonstige", $true)
    $pnlTop.Controls.Add($chkListStatus)

    $chkGroupByOS = New-Object System.Windows.Forms.CheckBox
    $chkGroupByOS.Location = "100, 78"; $chkGroupByOS.Size = "250, 22"; $chkGroupByOS.Text = "Nach Operating System sortieren"; $chkGroupByOS.Checked = $true
    $pnlTop.Controls.Add($chkGroupByOS)

    $btnRunAudit = New-Object System.Windows.Forms.Button
    $btnRunAudit.Location = "885, 12"; $btnRunAudit.Size = "140, 40"; $btnRunAudit.Text = "AD Scannen"
    $btnRunAudit.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $btnRunAudit.ForeColor = [System.Drawing.Color]::White
    $btnRunAudit.Font = New-Object System.Drawing.Font($subForm.Font.FontFamily, 9, [System.Drawing.FontStyle]::Bold)
    $pnlTop.Controls.Add($btnRunAudit)

    $lblAuditStatus = New-Object System.Windows.Forms.Label
    $lblAuditStatus.Location = "885, 58"; $lblAuditStatus.Size = "500, 30"; $lblAuditStatus.ForeColor = [System.Drawing.Color]::DarkBlue
    $pnlTop.Controls.Add($lblAuditStatus)

    $gridOSAudit = New-Object System.Windows.Forms.DataGridView
    $gridOSAudit.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridOSAudit.ReadOnly = $true
    $gridOSAudit.AllowUserToAddRows = $false
    $gridOSAudit.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
    $gridOSAudit.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $gridOSAudit.BackgroundColor = [System.Drawing.Color]::White
    $subForm.Controls.Add($gridOSAudit); $gridOSAudit.BringToFront()

    $subForm.Add_Shown({
        $lblAuditStatus.Text = "Lade AD Computerkonten..."
        $subForm.Refresh()
        $script:RawComputersList = Search-NativeLdap -LdapFilter "(objectCategory=computer)" -PropertiesToLoad @("name","operatingSystem","operatingSystemVersion","operatingSystemServicePack","userAccountControl","distinguishedName")
        
        $osSet = [System.Collections.Generic.HashSet[string]]::new()
        foreach ($c in $script:RawComputersList) {
            $osVal = if ($c.Properties["operatingsystem"].Count -gt 0) { $c.Properties["operatingsystem"][0] } else { "Unspecified OS" }
            [void]$osSet.Add($osVal)
        }
        foreach ($osItem in ($osSet | Sort-Object)) { [void]$cmbOSFilter.Items.Add($osItem) }
        $lblAuditStatus.Text = "$($script:RawComputersList.Count) Computer geladen. Bereit."
        $btnRunAudit.PerformClick()
    })

    $btnRunAudit.Add_Click({
        if (-not $script:RawComputersList) { return }
        $lblAuditStatus.Text = "Wende Filter an..."
        $subForm.Refresh()

        $selectedOS = $cmbOSFilter.SelectedItem
        $selectedTypeIndex = $cmbSystemType.SelectedIndex
        $namePattern = $txtNameSearch.Text -replace '\*', '.*'
        $selectedEnabled = $cmbEnabledFilter.SelectedIndex
        $selectedStatuses = @($chkListStatus.CheckedItems | ForEach-Object { $_.ToString() })

        $auditResults = @()

        foreach ($c in $script:RawComputersList) {
            $cName = $c.Properties["name"][0]
            if ($namePattern -and $cName -notmatch "(?i)^$namePattern$") { continue }

            $osName = if ($c.Properties["operatingsystem"].Count -gt 0) { $c.Properties["operatingsystem"][0] } else { "Unspecified OS" }
            $osVer  = if ($c.Properties["operatingsystemversion"].Count -gt 0) { $c.Properties["operatingsystemversion"][0] } else { "" }
            $osSP   = if ($c.Properties["operatingsystemservicepack"].Count -gt 0) { $c.Properties["operatingsystemservicepack"][0] } else { "" }

            $isServer = $osName -match "Server"
            $sysTypeStr = if ($isServer) { "Server" } else { "Client" }

            if ($selectedTypeIndex -eq 1 -and $isServer) { continue }
            if ($selectedTypeIndex -eq 2 -and -not $isServer) { continue }
            if ($selectedOS -and $selectedOS -ne "All OS / Alle OS" -and $osName -ne $selectedOS) { continue }

            $isEnabled = $true
            if ($c.Properties["useraccountcontrol"].Count -gt 0) {
                if (([int]$c.Properties["useraccountcontrol"][0] -band 2) -eq 2) { $isEnabled = $false }
            }

            if ($selectedEnabled -eq 1 -and -not $isEnabled) { continue }
            if ($selectedEnabled -eq 2 -and $isEnabled) { continue }

            $eval = Analyze-OSSupportDetails -OSName $osName -OSVersion $osVer -OSServicePack $osSP
            $supportStatus = $eval.Status
            $eolDate       = $eval.EOLDate
            $clientVer     = $eval.ClientVersion
            $buildNum      = $eval.BuildNumber

            $matchesStatus = $false
            foreach ($st in $selectedStatuses) {
                if ($st -eq "Supported" -and $supportStatus -like "*Supported*") { $matchesStatus = $true; break }
                if ($st -eq "Near EOL" -and $supportStatus -eq "Near EOL") { $matchesStatus = $true; break }
                if ($st -eq "Out of Support / EOL" -and $supportStatus -like "*Out of Support*") { $matchesStatus = $true; break }
                if ($st -eq "Unbekannt / Sonstige" -and $supportStatus -notlike "*Supported*" -and $supportStatus -ne "Near EOL" -and $supportStatus -notlike "*Out of Support*") { $matchesStatus = $true; break }
            }
            if (-not $matchesStatus) { continue }

            $dnVal = $c.Properties["distinguishedname"][0]
            $ouPath = if ($dnVal -match "OU=.*") { $dnVal.Substring($dnVal.IndexOf("OU=")) } else { $dnVal }

            $auditResults += [PSCustomObject]@{
                "Operating System"        = $osName
                "Version / Release"       = $clientVer
                "Build Number"            = $buildNum
                "Computer Name"           = $cName
                "System Type"             = $sysTypeStr
                "Support Status"          = $supportStatus
                "EOL Date / Support-Ende" = $eolDate
                "Enabled"                 = $isEnabled
                "OU Path"                 = $ouPath
            }
        }

        $sortedResults = if ($chkGroupByOS.Checked) {
            $auditResults | Sort-Object "Operating System", "Computer Name"
        } else {
            $auditResults | Sort-Object "Computer Name"
        }

        $arrA = [System.Collections.ArrayList]::new()
        foreach ($item in $sortedResults) { [void]$arrA.Add($item) }

        $gridOSAudit.DataSource = $null
        $gridOSAudit.DataSource = $arrA

        foreach ($row in $gridOSAudit.Rows) {
            $stVal = $row.Cells["Support Status"].Value
            if ($stVal -like "*Out of Support*") {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 235, 235)
            } elseif ($stVal -eq "Near EOL") {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 248, 220)
            } elseif ($stVal -like "*Supported*") {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(240, 255, 240)
            }
        }

        $lblAuditStatus.Text = "Ergebnis: $($arrA.Count) Computer gefunden (aus $($script:RawComputersList.Count) AD-Konten)."
    })

    [void]$subForm.ShowDialog()
}

# ==============================================================================
# TOOL 8: CLIENT SOFTWARE & APP ANALYSE (LAYOUT KORRIGIERT)
# ==============================================================================
function Show-ClientSoftwareAnalysis {
    $subForm = New-Object System.Windows.Forms.Form
    $subForm.Text = "Tool 8 - Client Software & App Analyse ($env:COMPUTERNAME)"
    $subForm.Size = New-Object System.Drawing.Size(1100, 720)
    $subForm.MinimumSize = New-Object System.Drawing.Size(850, 500)
    $subForm.StartPosition = "CenterScreen"
    $subForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    # 1. Steuerungs- & Filter-Panel (Oben)
    $pnlTop = New-Object System.Windows.Forms.Panel
    $pnlTop.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlTop.Height = 90
    $pnlTop.Padding = New-Object System.Windows.Forms.Padding(10)
    $subForm.Controls.Add($pnlTop)

    # Suchfeld
    $lblSearch = New-Object System.Windows.Forms.Label
    $lblSearch.Text = "Suche / Search:"
    $lblSearch.Location = New-Object System.Drawing.Point(12, 15)
    $lblSearch.AutoSize = $true
    $pnlTop.Controls.Add($lblSearch)

    $txtSearch = New-Object System.Windows.Forms.TextBox
    $txtSearch.Location = New-Object System.Drawing.Point(115, 12)
    $txtSearch.Size = New-Object System.Drawing.Size(160, 23)
    $pnlTop.Controls.Add($txtSearch)

    # App-Typ Filter
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

    # Sprach-Filter
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

    # Scan-Button
    $btnScan = New-Object System.Windows.Forms.Button
    $btnScan.Text = "Apps einlesen"
    $btnScan.Location = New-Object System.Drawing.Point(680, 10)
    $btnScan.Size = New-Object System.Drawing.Size(120, 28)
    $btnScan.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $btnScan.ForeColor = [System.Drawing.Color]::White
    $btnScan.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnScan.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $pnlTop.Controls.Add($btnScan)

    # CSV-Export Button
    $btnExport = New-Object System.Windows.Forms.Button
    $btnExport.Text = "CSV Export"
    $btnExport.Location = New-Object System.Drawing.Point(810, 10)
    $btnExport.Size = New-Object System.Drawing.Size(100, 28)
    $pnlTop.Controls.Add($btnExport)

    # Statuszeile
    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = New-Object System.Drawing.Point(12, 55)
    $lblStatus.Size = New-Object System.Drawing.Size(950, 22)
    $lblStatus.ForeColor = [System.Drawing.Color]::DarkSlateGray
    $lblStatus.Text = "Klicken Sie auf 'Apps einlesen', um installierte Software zu laden."
    $pnlTop.Controls.Add($lblStatus)

    # 2. Ergebnistabelle (Unten / Fill)
    $gridApps = New-Object System.Windows.Forms.DataGridView
    $gridApps.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridApps.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $gridApps.AllowUserToAddRows = $false
    $gridApps.ReadOnly = $true
    $gridApps.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $gridApps.MultiSelect = $false
    $gridApps.BackgroundColor = [System.Drawing.Color]::White
    $gridApps.RowHeadersVisible = $false
    $subForm.Controls.Add($gridApps)

    # WICHTIG: Z-Order festlegen, damit die Tabelle NICHT unter das Top-Panel rutscht
    $gridApps.BringToFront()
    $pnlTop.SendToBack()

    $global:allInstalledApps = @()

    # Filter-Funktion
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

    # Auslese-Logik
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

    # Event-Bindings für Filter
    $txtSearch.Add_TextChanged({ & $applyFilter })
    $cmbType.Add_SelectedIndexChanged({ & $applyFilter })
    $cmbLang.Add_SelectedIndexChanged({ & $applyFilter })

    # Export
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

    [void]$subForm.ShowDialog()
}

# ==============================================================================
# TOOL 9: AD UNIVERSAL ACL & BERECHTIGUNGSVERGLEICH (LETZTSTAND)
# ==============================================================================
function Show-Tool9-ACLCompare {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Tool 9: Active Directory Universal ACL & Berechtigungsvergleich"
    $form.Size = New-Object System.Drawing.Size(1280, 800)
    $form.StartPosition = "CenterScreen"
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = "AD Objekt ACL-Vergleich (User, Gruppen, Computer, OUs & Berechtigungs-Diff)"
    $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $lblTitle.Location = "15, 12"; $lblTitle.Size = "850, 25"; $form.Controls.Add($lblTitle)

    $grpInputs = New-Object System.Windows.Forms.GroupBox
    $grpInputs.Text = "Zu vergleichende AD-Objekte"; $grpInputs.Location = "15, 45"; $grpInputs.Size = "1230, 115"; $form.Controls.Add($grpInputs)

    $lblType = New-Object System.Windows.Forms.Label
    $lblType.Text = "Objekttyp:"; $lblType.Location = "15, 25"; $lblType.Size = "80, 20"; $grpInputs.Controls.Add($lblType)

    $cmbObjType = New-Object System.Windows.Forms.ComboBox
    $cmbObjType.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $cmbObjType.Items.AddRange(@("Auto-Erkennung (Alle Typen)", "Computer", "User / Benutzer", "Group / Gruppe", "OU / Container"))
    $cmbObjType.SelectedIndex = 0; $cmbObjType.Location = "95, 22"; $cmbObjType.Size = "185, 23"; $grpInputs.Controls.Add($cmbObjType)

    $lblComp1 = New-Object System.Windows.Forms.Label
    $lblComp1.Text = "Objekt 1 (Referenz):"; $lblComp1.Location = "15, 58"; $lblComp1.Size = "130, 20"; $grpInputs.Controls.Add($lblComp1)

    $txtComp1 = New-Object System.Windows.Forms.TextBox
    $txtComp1.Location = "145, 55"; $txtComp1.Size = "220, 23"; $grpInputs.Controls.Add($txtComp1)

    $lblComp2 = New-Object System.Windows.Forms.Label
    $lblComp2.Text = "Objekt 2 (Vergleich):"; $lblComp2.Location = "15, 86"; $lblComp2.Size = "130, 20"; $grpInputs.Controls.Add($lblComp2)

    $txtComp2 = New-Object System.Windows.Forms.TextBox
    $txtComp2.Location = "145, 83"; $txtComp2.Size = "220, 23"; $grpInputs.Controls.Add($txtComp2)

    $lblFilterStatus = New-Object System.Windows.Forms.Label
    $lblFilterStatus.Text = "Filter Vergleich:"; $lblFilterStatus.Location = "390, 25"; $lblFilterStatus.Size = "100, 20"; $grpInputs.Controls.Add($lblFilterStatus)

    $cmbFilterStatus = New-Object System.Windows.Forms.ComboBox
    $cmbFilterStatus.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $cmbFilterStatus.Items.AddRange(@("Alle Einträge", "Nur Abweichungen (Unterschiede)", "Nur Identisch (Match)", "Nur auf Objekt 1", "Nur auf Objekt 2"))
    $cmbFilterStatus.SelectedIndex = 0; $cmbFilterStatus.Location = "495, 22"; $cmbFilterStatus.Size = "220, 23"; $grpInputs.Controls.Add($cmbFilterStatus)

    $chkHideInherited = New-Object System.Windows.Forms.CheckBox
    $chkHideInherited.Text = "Vererbte Berechtigungen ausblenden (Nur explizite ACLs)"; $chkHideInherited.Location = "390, 58"; $chkHideInherited.Size = "350, 22"
    $grpInputs.Controls.Add($chkHideInherited)

    $btnCompare = New-Object System.Windows.Forms.Button
    $btnCompare.Text = "ACLs Vergleichen"; $btnCompare.Location = "760, 22"; $btnCompare.Size = "160, 80"
    $btnCompare.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215); $btnCompare.ForeColor = [System.Drawing.Color]::White
    $btnCompare.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $grpInputs.Controls.Add($btnCompare)

    $btnExport = New-Object System.Windows.Forms.Button
    $btnExport.Text = "CSV Exportieren"; $btnExport.Location = "935, 22"; $btnExport.Size = "140, 80"
    $btnExport.BackColor = [System.Drawing.Color]::FromArgb(40, 167, 69); $btnExport.ForeColor = [System.Drawing.Color]::White
    $btnExport.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    $grpInputs.Controls.Add($btnExport)

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Text = "Geben Sie Namen, sAMAccountNames oder DNs zweier AD-Objekte ein."; $lblStatus.Location = "15, 165"; $lblStatus.Size = "1230, 24"
    $lblStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
    $form.Controls.Add($lblStatus)

    $gridACL = New-Object System.Windows.Forms.DataGridView
    $gridACL.Location = "15, 195"; $gridACL.Size = "1230, 550"
    $gridACL.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $gridACL.ReadOnly = $true; $gridACL.AllowUserToAddRows = $false
    $gridACL.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::DisplayedCells
    $gridACL.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $gridACL.BackgroundColor = [System.Drawing.Color]::White
    $form.Controls.Add($gridACL)

    $global:ComparisonResults = @()

    function Get-UniversalADObjectACL {
        param([string]$Identifier, [string]$TypeSelection)
        $target = $null

        if ($Identifier -match "^(CN|OU|DC)=") {
            try {
                $deDirect = [ADSI]"LDAP://$Identifier"
                if ($deDirect.distinguishedName) { $target = $deDirect }
            } catch {}
        }

        if (-not $target) {
            $baseFilter = switch ($TypeSelection) {
                "Computer"        { "(objectCategory=computer)" }
                "User / Benutzer"{ "(&(objectCategory=person)(objectClass=user))" }
                "Group / Gruppe" { "(objectCategory=group)" }
                "OU / Container" { "(|(objectCategory=organizationalUnit)(objectCategory=container))" }
                default          { "(objectClass=*)" }
            }
            $searchFilter = "(&$baseFilter(|(sAMAccountName=$Identifier)(sAMAccountName=$Identifier`$)(name=$Identifier)))"
            $searcher = [adsisearcher]$searchFilter
            $searcher.PageSize = 5
            $searchResult = $searcher.FindOne()
            if ($searchResult) { $target = $searchResult.GetDirectoryEntry() }
        }

        if (-not $target) { return $null }

        $rules = @()
        try {
            foreach ($rule in $target.ObjectSecurity.GetAccessRules($true, $true, [System.Security.Principal.NTAccount])) {
                $rules += [PSCustomObject]@{
                    DistinguishedName = $target.distinguishedName
                    IdentityReference = $rule.IdentityReference.Value
                    Rights            = $rule.ActiveDirectoryRights.ToString()
                    AccessControlType = $rule.AccessControlType.ToString()
                    IsInherited       = $rule.IsInherited
                    InheritanceType   = $rule.InheritanceType.ToString()
                    Signature         = "$($rule.IdentityReference.Value)|$($rule.ActiveDirectoryRights)|$($rule.AccessControlType)"
                }
            }
        } catch { return $null }

        $objTypeStr = if ($target.schemaClassName) { $target.schemaClassName } else { "AD-Objekt" }
        return @{ "DN" = $target.distinguishedName; "Rules" = $rules; "Class" = $objTypeStr; "Name" = $target.name }
    }

    function Update-GridDisplay {
        if (-not $global:ComparisonResults) { return }
        $filtered = $global:ComparisonResults
        $mode = $cmbFilterStatus.SelectedItem.ToString()
        if ($mode -eq "Nur Abweichungen (Unterschiede)") {
            $filtered = $filtered | Where-Object { $_."Vergleichs-Status" -ne "Identisch (Match)" }
        } elseif ($mode -eq "Nur Identisch (Match)") {
            $filtered = $filtered | Where-Object { $_."Vergleichs-Status" -eq "Identisch (Match)" }
        } elseif ($mode -eq "Nur auf Objekt 1") {
            $filtered = $filtered | Where-Object { $_."Vergleichs-Status" -eq "Nur auf Objekt 1" }
        } elseif ($mode -eq "Nur auf Objekt 2") {
            $filtered = $filtered | Where-Object { $_."Vergleichs-Status" -eq "Nur auf Objekt 2" }
        }

        if ($chkHideInherited.Checked) {
            $filtered = $filtered | Where-Object { 
                $_."Objekt 1 Vererbt" -eq "Nein (Explizit)" -or $_."Objekt 2 Vererbt" -eq "Nein (Explizit)" 
            }
        }

        $arrList = [System.Collections.ArrayList]::new()
        foreach ($f in $filtered) { [void]$arrList.Add($f) }
        $gridACL.DataSource = $null
        $gridACL.DataSource = $arrList

        foreach ($row in $gridACL.Rows) {
            $statusVal = $row.Cells["Vergleichs-Status"].Value
            if ($statusVal -eq "Nur auf Objekt 1") {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 235, 235)
            } elseif ($statusVal -eq "Nur auf Objekt 2") {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(235, 245, 255)
            } elseif ($statusVal -like "*Abweichend*") {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 248, 220)
            }
        }
    }

    $btnCompare.Add_Click({
        $c1 = $txtComp1.Text.Trim(); $c2 = $txtComp2.Text.Trim(); $selectedType = $cmbObjType.SelectedItem.ToString()
        if ([string]::IsNullOrWhiteSpace($c1) -or [string]::IsNullOrWhiteSpace($c2)) {
            [System.Windows.Forms.MessageBox]::Show("Bitte beide Objektnamen angeben!", "Eingabefehler", "OK", "Warning")
            return
        }

        $lblStatus.Text = "Lese ACLs für '$c1' und '$c2' aus Active Directory..."; $lblStatus.ForeColor = [System.Drawing.Color]::Black; $form.Refresh()

        try {
            $obj1Data = Get-UniversalADObjectACL -Identifier $c1 -TypeSelection $selectedType
            $obj2Data = Get-UniversalADObjectACL -Identifier $c2 -TypeSelection $selectedType

            if (-not $obj1Data) { $lblStatus.Text = "Fehler: Objekt '$c1' nicht gefunden!"; $lblStatus.ForeColor = [System.Drawing.Color]::Red; return }
            if (-not $obj2Data) { $lblStatus.Text = "Fehler: Objekt '$c2' nicht gefunden!"; $lblStatus.ForeColor = [System.Drawing.Color]::Red; return }

            $rules1 = $obj1Data.Rules; $rules2 = $obj2Data.Rules
            $comparisonList = [System.Collections.Generic.List[PSCustomObject]]::new()
            $allIdentities = ($rules1.IdentityReference + $rules2.IdentityReference) | Select-Object -Unique | Sort-Object

            foreach ($ident in $allIdentities) {
                $r1List = $rules1 | Where-Object { $_.IdentityReference -eq $ident }
                $r2List = $rules2 | Where-Object { $_.IdentityReference -eq $ident }

                if ($r1List -and -not $r2List) {
                    foreach ($r in $r1List) {
                        $comparisonList.Add([PSCustomObject]@{
                            "Vergleichs-Status"  = "Nur auf Objekt 1"
                            "Principal / Gruppe" = $r.IdentityReference
                            "Rechte (Objekt 1)"  = $r.Rights
                            "Rechte (Objekt 2)"  = "-- Nicht vorhanden --"
                            "Typ"                = $r.AccessControlType
                            "Objekt 1 Vererbt"   = if ($r.IsInherited) { "Ja" } else { "Nein (Explizit)" }
                            "Objekt 2 Vererbt"   = "-"
                        })
                    }
                } elseif (-not $r1List -and $r2List) {
                    foreach ($r in $r2List) {
                        $comparisonList.Add([PSCustomObject]@{
                            "Vergleichs-Status"  = "Nur auf Objekt 2"
                            "Principal / Gruppe" = $r.IdentityReference
                            "Rechte (Objekt 1)"  = "-- Nicht vorhanden --"
                            "Rechte (Objekt 2)"  = $r.Rights
                            "Typ"                = $r.AccessControlType
                            "Objekt 1 Vererbt"   = "-"
                            "Objekt 2 Vererbt"   = if ($r.IsInherited) { "Ja" } else { "Nein (Explizit)" }
                        })
                    }
                } else {
                    $matchedSigs = @()
                    foreach ($r1 in $r1List) {
                        $matchingR2 = $r2List | Where-Object { $_.Signature -eq $r1.Signature } | Select-Object -First 1
                        if ($matchingR2) {
                            $matchedSigs += $r1.Signature
                            $isDiffInherit = ($r1.IsInherited -ne $matchingR2.IsInherited)
                            $status = if ($isDiffInherit) { "Abweichend (Vererbung)" } else { "Identisch (Match)" }
                            $comparisonList.Add([PSCustomObject]@{
                                "Vergleichs-Status"  = $status
                                "Principal / Gruppe" = $r1.IdentityReference
                                "Rechte (Objekt 1)"  = $r1.Rights
                                "Rechte (Objekt 2)"  = $matchingR2.Rights
                                "Typ"                = $r1.AccessControlType
                                "Objekt 1 Vererbt"   = if ($r1.IsInherited) { "Ja" } else { "Nein (Explizit)" }
                                "Objekt 2 Vererbt"   = if ($matchingR2.IsInherited) { "Ja" } else { "Nein (Explizit)" }
                            })
                        } else {
                            $comparisonList.Add([PSCustomObject]@{
                                "Vergleichs-Status"  = "Abweichende Rechte"
                                "Principal / Gruppe" = $r1.IdentityReference
                                "Rechte (Objekt 1)"  = $r1.Rights
                                "Rechte (Objekt 2)"  = "(Abweichend/Nicht gleich)"
                                "Typ"                = $r1.AccessControlType
                                "Objekt 1 Vererbt"   = if ($r1.IsInherited) { "Ja" } else { "Nein (Explizit)" }
                                "Objekt 2 Vererbt"   = "-"
                            })
                        }
                    }

                    foreach ($r2 in $r2List) {
                        if ($r2.Signature -notin $matchedSigs) {
                            $comparisonList.Add([PSCustomObject]@{
                                "Vergleichs-Status"  = "Abweichende Rechte"
                                "Principal / Gruppe" = $r2.IdentityReference
                                "Rechte (Objekt 1)"  = "(Abweichend/Nicht gleich)"
                                "Rechte (Objekt 2)"  = $r2.Rights
                                "Typ"                = $r2.AccessControlType
                                "Objekt 1 Vererbt"   = "-"
                                "Objekt 2 Vererbt"   = if ($r2.IsInherited) { "Ja" } else { "Nein (Explizit)" }
                            })
                        }
                    }
                }
            }

            $global:ComparisonResults = $comparisonList
            Update-GridDisplay
            $diffCount = ($comparisonList | Where-Object { $_."Vergleichs-Status" -ne "Identisch (Match)" }).Count
            $lblStatus.ForeColor = [System.Drawing.Color]::Black
            $lblStatus.Text = "Objekt 1: [$($obj1Data.Class)] $($obj1Data.DN)  |  Objekt 2: [$($obj2Data.Class)] $($obj2Data.DN)  |  Einträge: $($comparisonList.Count) (Abweichungen: $diffCount)"
        } catch {
            $lblStatus.ForeColor = [System.Drawing.Color]::Red
            $lblStatus.Text = "Fehler beim ACL-Vergleich: $($_.Exception.Message)"
        }
    })

    $cmbFilterStatus.Add_SelectedIndexChanged({ Update-GridDisplay })
    $chkHideInherited.Add_CheckedChanged({ Update-GridDisplay })

    $btnExport.Add_Click({
        if (-not $gridACL.DataSource -or $gridACL.Rows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Keine Daten zum Exportieren vorhanden!", "Export Info", "OK", "Information")
            return
        }
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "CSV-Datei (*.csv)|*.csv"
        $sfd.FileName = "AD_ACL_Vergleich_$($txtComp1.Text)_vs_$($txtComp2.Text)_$((Get-Date).ToString('yyyyMMdd_HHmm')).csv"
        if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $gridACL.DataSource | Export-Csv -Path $sfd.FileName -NoTypeInformation -Delimiter ";" -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show("ACL-Vergleich erfolgreich exportiert nach:`n$($sfd.FileName)", "Export Abgeschlossen", "OK", "Information")
        }
    })

    [void]$form.ShowDialog()
}

# ==============================================================================
# TOOL 11: AD KENNWORT-RICHTLINIEN & PSO AUDIT (LAYOUT & PSO FIX)
# ==============================================================================
function Open-ToolPasswordPolicies {
    if (-not (Assert-DomainJoined)) { return }

    $t = $script:I18N[$script:CurrentLang]

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $t["Tool11Title"]
    $form.Size = New-Object System.Drawing.Size(1150, 720)
    $form.StartPosition = "CenterParent"
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

    # Header Panel
    $headerPanel = New-Object System.Windows.Forms.Panel
    $headerPanel.Dock = "Top"
    $headerPanel.Height = 60
    $headerPanel.BackColor = [System.Drawing.Color]::White
    $form.Controls.Add($headerPanel)

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = $t["Tool11Title"]
    $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)
    $lblTitle.ForeColor = [System.Drawing.Color]::FromArgb(15, 23, 42)
    $lblTitle.Location = New-Object System.Drawing.Point(20, 8)
    $lblTitle.AutoSize = $true
    $headerPanel.Controls.Add($lblTitle)

    $lblSubTitle = New-Object System.Windows.Forms.Label
    $lblSubTitle.Text = $t["Tool11Sub"]
    $lblSubTitle.ForeColor = [System.Drawing.Color]::FromArgb(100, 116, 139)
    $lblSubTitle.Location = New-Object System.Drawing.Point(22, 33)
    $lblSubTitle.AutoSize = $true
    $headerPanel.Controls.Add($lblSubTitle)

    # Bottom Panel
    $bottomPanel = New-Object System.Windows.Forms.Panel
    $bottomPanel.Dock = "Bottom"
    $bottomPanel.Height = 50
    $bottomPanel.BackColor = [System.Drawing.Color]::White
    $form.Controls.Add($bottomPanel)

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = New-Object System.Drawing.Point(20, 15)
    $lblStatus.AutoSize = $true
    $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(71, 85, 105)
    $lblStatus.Text = $t["StatusReady"]
    $bottomPanel.Controls.Add($lblStatus)

    $btnExport = New-Object System.Windows.Forms.Button
    $btnExport.Text = $t["BtnExportCsv"]
    $btnExport.Size = New-Object System.Drawing.Size(130, 30)
    $btnExport.Location = New-Object System.Drawing.Point(830, 10)
    $btnExport.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
    $btnExport.BackColor = [System.Drawing.Color]::FromArgb(37, 99, 235)
    $btnExport.ForeColor = [System.Drawing.Color]::White
    $btnExport.FlatStyle = "Flat"
    $btnExport.FlatAppearance.BorderSize = 0
    $bottomPanel.Controls.Add($btnExport)

    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = $t["BtnRefresh"]
    $btnRefresh.Size = New-Object System.Drawing.Size(130, 30)
    $btnRefresh.Location = New-Object System.Drawing.Point(970, 10)
    $btnRefresh.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
    $btnRefresh.BackColor = [System.Drawing.Color]::FromArgb(100, 116, 139)
    $btnRefresh.ForeColor = [System.Drawing.Color]::White
    $btnRefresh.FlatStyle = "Flat"
    $btnRefresh.FlatAppearance.BorderSize = 0
    $bottomPanel.Controls.Add($btnRefresh)

    # Tab Control
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Dock = "Fill"
    $tabControl.Padding = New-Object System.Drawing.Point(12, 6)
    $form.Controls.Add($tabControl)
    $tabControl.BringToFront()

    # --- TAB 1: Default Domain Password Policy ---
    $tabDomain = New-Object System.Windows.Forms.TabPage
    $tabDomain.Text = $t["TabDefaultPolicy"]
    $tabDomain.BackColor = [System.Drawing.Color]::White
    $tabControl.TabPages.Add($tabDomain)

    $gridDomain = New-Object System.Windows.Forms.DataGridView
    $gridDomain.Dock = "Fill"
    $gridDomain.BackgroundColor = [System.Drawing.Color]::White
    $gridDomain.BorderStyle = "None"
    $gridDomain.ReadOnly = $true
    $gridDomain.AllowUserToAddRows = $false
    $gridDomain.SelectionMode = "FullRowSelect"
    $gridDomain.AutoSizeColumnsMode = "AllCells"
    $tabDomain.Controls.Add($gridDomain)

    # --- TAB 2: Fine-Grained PSOs ---
    $tabPSO = New-Object System.Windows.Forms.TabPage
    $tabPSO.Text = $t["TabFineGrained"]
    $tabPSO.BackColor = [System.Drawing.Color]::White
    $tabControl.TabPages.Add($tabPSO)

    $gridPSO = New-Object System.Windows.Forms.DataGridView
    $gridPSO.Dock = "Fill"
    $gridPSO.BackgroundColor = [System.Drawing.Color]::White
    $gridPSO.BorderStyle = "None"
    $gridPSO.ReadOnly = $true
    $gridPSO.AllowUserToAddRows = $false
    $gridPSO.SelectionMode = "FullRowSelect"
    $gridPSO.AutoSizeColumnsMode = "AllCells"
    $tabPSO.Controls.Add($gridPSO)

    # --- TAB 3: User Effective Policy Check ---
    $tabUser = New-Object System.Windows.Forms.TabPage
    $tabUser.Text = $t["TabUserCheck"]
    $tabUser.BackColor = [System.Drawing.Color]::White
    $tabControl.TabPages.Add($tabUser)

    $userTopPanel = New-Object System.Windows.Forms.Panel
    $userTopPanel.Dock = "Top"
    $userTopPanel.Height = 55
    $userTopPanel.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 252)
    $tabUser.Controls.Add($userTopPanel)

    $lblUser = New-Object System.Windows.Forms.Label
    $lblUser.Text = $t["UserSearchLabel"]
    $lblUser.Location = New-Object System.Drawing.Point(15, 18)
    $lblUser.AutoSize = $true
    $lblUser.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $userTopPanel.Controls.Add($lblUser)

    # FIX: Position nach rechts gerückt (X=320), damit das deutsche Label nicht überlappt
    $txtUserCheck = New-Object System.Windows.Forms.TextBox
    $txtUserCheck.Location = New-Object System.Drawing.Point(320, 14)
    $txtUserCheck.Size = New-Object System.Drawing.Size(200, 25)
    $userTopPanel.Controls.Add($txtUserCheck)

    $btnCheckUser = New-Object System.Windows.Forms.Button
    $btnCheckUser.Text = $t["BtnCheckUser"]
    $btnCheckUser.Location = New-Object System.Drawing.Point(530, 13)
    $btnCheckUser.Size = New-Object System.Drawing.Size(150, 27)
    $btnCheckUser.BackColor = [System.Drawing.Color]::FromArgb(15, 118, 110)
    $btnCheckUser.ForeColor = [System.Drawing.Color]::White
    $btnCheckUser.FlatStyle = "Flat"
    $btnCheckUser.FlatAppearance.BorderSize = 0
    $userTopPanel.Controls.Add($btnCheckUser)

    $gridUserPolicy = New-Object System.Windows.Forms.DataGridView
    $gridUserPolicy.Dock = "Fill"
    $gridUserPolicy.BackgroundColor = [System.Drawing.Color]::White
    $gridUserPolicy.BorderStyle = "None"
    $gridUserPolicy.ReadOnly = $true
    $gridUserPolicy.AllowUserToAddRows = $false
    $gridUserPolicy.SelectionMode = "FullRowSelect"
    $gridUserPolicy.AutoSizeColumnsMode = "AllCells"
    $tabUser.Controls.Add($gridUserPolicy)
    $gridUserPolicy.BringToFront()

    # Helper: Convert I8 / Large Integer FileTime Duration to Minutes/Days/Hours
    $convertTimeSpan = {
        param($val, $asDays = $false)
        if (-not $val) { return if ($script:CurrentLang -eq "DE") { "Nicht konfiguriert" } else { "Not Configured" } }
        try {
            $ticks = 0
            if ($val -is [System.Int64]) {
                $ticks = $val
            } elseif ($val.GetType().Name -eq "__ComObject" -or $val.GetType().FullName -like "*LargeInteger*") {
                $ticks = ([int64]$val.HighPart -shl 32) + [int64]$val.LowPart
            } else {
                $ticks = [int64]::Parse($val.ToString())
            }

            if ($ticks -eq 0 -or $ticks -eq [Int64]::MinValue) {
                return if ($script:CurrentLang -eq "DE") { "Nie ablaufend / Deaktiviert" } else { "Never Expires / Disabled" }
            }

            # Negative FileTime Ticks: 1 Tick = 100 Nanosekunden
            if ($ticks -lt 0) { $ticks = -$ticks }
            $ts = [TimeSpan]::FromTicks($ticks)

            if ($asDays) {
                if ($ts.TotalDays -ge 1) {
                    $unit = if ($script:CurrentLang -eq "DE") { "Tage" } else { "Days" }
                    return "{0:N1} $unit" -f $ts.TotalDays
                } else {
                    $unit = if ($script:CurrentLang -eq "DE") { "Stunden" } else { "Hours" }
                    return "{0:N1} $unit" -f $ts.TotalHours
                }
            } else {
                if ($ts.TotalMinutes -ge 60) {
                    $unit = if ($script:CurrentLang -eq "DE") { "Stunden" } else { "Hours" }
                    return "{0:N1} $unit" -f $ts.TotalHours
                } else {
                    $unit = if ($script:CurrentLang -eq "DE") { "Minuten" } else { "Minutes" }
                    return "{0:N0} $unit" -f $ts.TotalMinutes
                }
            }
        } catch {
            return $val.ToString()
        }
    }

    $script:cachedDomainPolicies = @()
    $script:cachedPSOs = @()
    $script:cachedUserPolicy = @()

    # --- LADE-LOGIK ---
    $loadPolicies = {
        $lblStatus.Text = $t["StatusSearching"]
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $script:cachedDomainPolicies = @()
        $script:cachedPSOs = @()

        try {
            # 1. Standard Domain Password Policy
            $rootEntry = [System.DirectoryServices.DirectoryEntry]::new("LDAP://RootDSE")
            $defNamingContext = $rootEntry.defaultNamingContext.ToString()
            $domEntry = [System.DirectoryServices.DirectoryEntry]::new("LDAP://$defNamingContext")

            $minPwdAge      = & $convertTimeSpan $domEntry.Properties["minPwdAge"].Value $true
            $maxPwdAge      = & $convertTimeSpan $domEntry.Properties["maxPwdAge"].Value $true
            $minPwdLength   = if ($domEntry.Properties["minPwdLength"].Value) { "$($domEntry.Properties['minPwdLength'].Value) " + (if ($script:CurrentLang -eq "DE") { "Zeichen" } else { "Characters" }) } else { "0" }
            $pwdHistory     = if ($domEntry.Properties["pwdHistoryLength"].Value) { "$($domEntry.Properties['pwdHistoryLength'].Value) " + (if ($script:CurrentLang -eq "DE") { "Kennwörter gespeichert" } else { "Passwords remembered" }) } else { "0" }
            
            $pwdProps = [int]($domEntry.Properties["pwdProperties"].Value)
            $complexityEnabled = ($pwdProps -band 1) -eq 1
            $reversibleEnc     = ($pwdProps -band 16) -eq 16

            $lockThreshold = if ($domEntry.Properties["lockoutThreshold"].Value) { "$($domEntry.Properties['lockoutThreshold'].Value) " + (if ($script:CurrentLang -eq "DE") { "ungültige Versuche" } else { "Invalid attempts" }) } else { if ($script:CurrentLang -eq "DE") { "0 (Deaktiviert)" } else { "0 (Disabled)" } }
            $lockDuration  = & $convertTimeSpan $domEntry.Properties["lockoutDuration"].Value $false
            $lockWindow    = & $convertTimeSpan $domEntry.Properties["lockOutObservationWindow"].Value $false

            $domList = @(
                [PSCustomObject]@{ $t["PropCategory"] = $t["CatPassword"]; $t["PropName"] = $t["PropMinAge"]; $t["PropValue"] = $minPwdAge },
                [PSCustomObject]@{ $t["PropCategory"] = $t["CatPassword"]; $t["PropName"] = $t["PropMaxAge"]; $t["PropValue"] = $maxPwdAge },
                [PSCustomObject]@{ $t["PropCategory"] = $t["CatPassword"]; $t["PropName"] = $t["PropMinLength"]; $t["PropValue"] = $minPwdLength },
                [PSCustomObject]@{ $t["PropCategory"] = $t["CatPassword"]; $t["PropName"] = $t["PropHistory"]; $t["PropValue"] = $pwdHistory },
                [PSCustomObject]@{ $t["PropCategory"] = $t["CatPassword"]; $t["PropName"] = $t["PropComplexity"]; $t["PropValue"] = if ($complexityEnabled) { $t["ValActive"] } else { $t["ValDisabled"] } },
                [PSCustomObject]@{ $t["PropCategory"] = $t["CatPassword"]; $t["PropName"] = $t["PropReversible"]; $t["PropValue"] = if ($reversibleEnc) { $t["ValActiveInsecure"] } else { $t["ValDisabledSecure"] } },
                [PSCustomObject]@{ $t["PropCategory"] = $t["CatLockout"];  $t["PropName"] = $t["PropThreshold"]; $t["PropValue"] = $lockThreshold },
                [PSCustomObject]@{ $t["PropCategory"] = $t["CatLockout"];  $t["PropName"] = $t["PropDuration"]; $t["PropValue"] = $lockDuration },
                [PSCustomObject]@{ $t["PropCategory"] = $t["CatLockout"];  $t["PropName"] = $t["PropObservation"]; $t["PropValue"] = $lockWindow }
            )

            $script:cachedDomainPolicies = $domList
            $gridDomain.DataSource = [System.Collections.ArrayList]::new($domList)

            # 2. Fine-Grained Password Policies (PSO) FIX: Direkte Container-Suche + Globaler Fallback
            $psoList = @()
            $psoContainerPath = "LDAP://CN=Password Settings Objects,CN=System,$defNamingContext"
            $psoSearcher = $null
            
            try {
                $psoEntry = [System.DirectoryServices.DirectoryEntry]::new($psoContainerPath)
                $psoSearcher = [System.DirectoryServices.DirectorySearcher]::new($psoEntry)
                $psoSearcher.Filter = "(objectClass=msDS-PasswordSettings)"
                $psoSearcher.SearchScope = [System.DirectoryServices.SearchScope]::OneLevel
                $psoResults = $psoSearcher.FindAll()
            } catch {
                # Fallback: Subtree Suche im gesamten Domänen-Root
                $psoSearcher = [System.DirectoryServices.DirectorySearcher]::new($domEntry)
                $psoSearcher.Filter = "(objectClass=msDS-PasswordSettings)"
                $psoSearcher.SearchScope = [System.DirectoryServices.SearchScope]::Subtree
                $psoResults = $psoSearcher.FindAll()
            }

            if ($psoResults -and $psoResults.Count -gt 0) {
                foreach ($res in $psoResults) {
                    $prop = $res.Properties
                    $name = if ($prop["name"].Count -gt 0) { $prop["name"][0] } else { "Unbenannt" }
                    $prec = if ($prop["msds-passwordsettingsprecedence"].Count -gt 0) { $prop["msds-passwordsettingsprecedence"][0] } else { "N/A" }
                    $pMinLen = if ($prop["msds-minimumpasswordlength"].Count -gt 0) { "$($prop['msds-minimumpasswordlength'][0]) " + (if ($script:CurrentLang -eq "DE") { "Zeichen" } else { "Characters" }) } else { "N/A" }
                    $pHist = if ($prop["msds-passwordhistorylength"].Count -gt 0) { "$($prop['msds-passwordhistorylength'][0])" } else { "0" }
                    $pCompl = if ($prop["msds-passwordcomplexityenabled"].Count -gt 0) { [bool]$prop["msds-passwordcomplexityenabled"][0] } else { $false }
                    $pRevers = if ($prop["msds-passwordreversibleencryptionenabled"].Count -gt 0) { [bool]$prop["msds-passwordreversibleencryptionenabled"][0] } else { $false }
                    
                    $pMinAge = if ($prop["msds-minpasswordage"].Count -gt 0) { & $convertTimeSpan $prop["msds-minpasswordage"][0] $true } else { "N/A" }
                    $pMaxAge = if ($prop["msds-maxpasswordage"].Count -gt 0) { & $convertTimeSpan $prop["msds-maxpasswordage"][0] $true } else { "N/A" }
                    
                    $pLockThresh = if ($prop["msds-lockoutthreshold"].Count -gt 0) { "$($prop['msds-lockoutthreshold'][0]) " + (if ($script:CurrentLang -eq "DE") { "Versuche" } else { "Attempts" }) } else { "0" }
                    $pLockDur = if ($prop["msds-lockoutduration"].Count -gt 0) { & $convertTimeSpan $prop["msds-lockoutduration"][0] $false } else { "N/A" }
                    $pLockObs = if ($prop["msds-lockoutobservationwindow"].Count -gt 0) { & $convertTimeSpan $prop["msds-lockoutobservationwindow"][0] $false } else { "N/A" }

                    $appliesTo = @()
                    if ($prop["msds-psoappliesto"].Count -gt 0) {
                        foreach ($appDn in $prop["msds-psoappliesto"]) {
                            if ($appDn -match "CN=([^,]+)") { $appliesTo += $Matches[1] } else { $appliesTo += $appDn }
                        }
                    }
                    $appliesToStr = if ($appliesTo.Count -gt 0) { $appliesTo -join ", " } else { if ($script:CurrentLang -eq "DE") { "Niemand (Keine Zuweisung)" } else { "None (No assignment)" } }

                    $psoList += [PSCustomObject]@{
                        $t["PsoName"]        = $name
                        $t["PsoPrecedence"]  = $prec
                        $t["PsoMinLength"]   = $pMinLen
                        $t["PsoHistory"]     = $pHist
                        $t["PsoComplexity"]  = if ($pCompl) { $t["ValActive"] } else { $t["ValDisabled"] }
                        $t["PsoMinAge"]      = $pMinAge
                        $t["PsoMaxAge"]      = $pMaxAge
                        $t["PsoThreshold"]   = $pLockThresh
                        $t["PsoDuration"]    = $pLockDur
                        $t["PsoAppliesTo"]   = $appliesToStr
                    }
                }
            }

            $script:cachedPSOs = $psoList
            $gridPSO.DataSource = [System.Collections.ArrayList]::new($psoList)

            $psoCount = $psoList.Count
            $psoMsg = if ($script:CurrentLang -eq "DE") { "$psoCount Fine-Grained PSO(s) gefunden." } else { "$psoCount Fine-Grained PSO(s) found." }
            $lblStatus.Text = "$($t['StatusResult']) $psoMsg"

        } catch {
            $lblStatus.Text = "$($t['StatusError']): $($_.Exception.Message)"
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }

    # --- BENUTZER-CHECK LOGIK ---
    $checkUserAction = {
        $uName = $txtUserCheck.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($uName)) {
            [System.Windows.Forms.MessageBox]::Show($t["ErrUserEmpty"], "Info", "OK", "Warning")
            return
        }

        $lblStatus.Text = $t["StatusSearching"]
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

        try {
            $rootEntry = [System.DirectoryServices.DirectoryEntry]::new("LDAP://RootDSE")
            $defNamingContext = $rootEntry.defaultNamingContext.ToString()
            $domEntry = [System.DirectoryServices.DirectoryEntry]::new("LDAP://$defNamingContext")
            
            $uSearcher = [System.DirectoryServices.DirectorySearcher]::new($domEntry)
            $uSearcher.Filter = "(&(objectClass=user)(objectCategory=person)(|(sAMAccountName=$uName)(userPrincipalName=$uName)))"
            $uSearcher.PropertiesToLoad.AddRange(@("distinguishedName", "sAMAccountName", "userPrincipalName", "msDS-ResultantPSO", "userAccountControl", "pwdLastSet"))
            $uRes = $uSearcher.FindOne()

            if (-not $uRes) {
                [System.Windows.Forms.MessageBox]::Show($t["ErrUserNotFound"], "AD Search", "OK", "Warning")
                $lblStatus.Text = $t["ErrUserNotFound"]
                return
            }

            $uProps = $uRes.Properties
            $foundSam = $uProps["samaccountname"][0]
            $resultantPSO = if ($uProps["msds-resultantpso"].Count -gt 0) { $uProps["msds-resultantpso"][0] } else { $null }

            $userResultList = @()
            if ($resultantPSO) {
                # Liest die PSO-Werte aus
                $psoEntry = [System.DirectoryServices.DirectoryEntry]::new("LDAP://$resultantPSO")
                $psoName = $psoEntry.Properties["name"].Value
                $pMinLen = "$($psoEntry.Properties['msDS-MinimumPasswordLength'].Value) " + (if ($script:CurrentLang -eq "DE") { "Zeichen" } else { "Characters" })
                $pMinAge = & $convertTimeSpan $psoEntry.Properties["msDS-MinPasswordAge"].Value $true
                $pMaxAge = & $convertTimeSpan $psoEntry.Properties["msDS-MaxPasswordAge"].Value $true
                $pCompl  = if ([bool]$psoEntry.Properties["msDS-PasswordComplexityEnabled"].Value) { $t["ValActive"] } else { $t["ValDisabled"] }
                $pThresh = "$($psoEntry.Properties['msDS-LockoutThreshold'].Value) " + (if ($script:CurrentLang -eq "DE") { "Versuche" } else { "Attempts" })
                $pDur    = & $convertTimeSpan $psoEntry.Properties["msDS-LockoutDuration"].Value $false

                $userResultList = @(
                    [PSCustomObject]@{ $t["PropCategory"] = $t["CatAppliedPolicy"]; $t["PropName"] = $t["PropSource"]; $t["PropValue"] = "Fine-Grained PSO ($psoName)" },
                    [PSCustomObject]@{ $t["PropCategory"] = $t["CatAppliedPolicy"]; $t["PropName"] = $t["PropDN"]; $t["PropValue"] = $resultantPSO },
                    [PSCustomObject]@{ $t["PropCategory"] = $t["CatPassword"];      $t["PropName"] = $t["PropMinLength"]; $t["PropValue"] = $pMinLen },
                    [PSCustomObject]@{ $t["PropCategory"] = $t["CatPassword"];      $t["PropName"] = $t["PropMinAge"]; $t["PropValue"] = $pMinAge },
                    [PSCustomObject]@{ $t["PropCategory"] = $t["CatPassword"];      $t["PropName"] = $t["PropMaxAge"]; $t["PropValue"] = $pMaxAge },
                    [PSCustomObject]@{ $t["PropCategory"] = $t["CatPassword"];      $t["PropName"] = $t["PropComplexity"]; $t["PropValue"] = $pCompl },
                    [PSCustomObject]@{ $t["PropCategory"] = $t["CatLockout"];       $t["PropName"] = $t["PropThreshold"]; $t["PropValue"] = $pThresh },
                    [PSCustomObject]@{ $t["PropCategory"] = $t["CatLockout"];       $t["PropName"] = $t["PropDuration"]; $t["PropValue"] = $pDur }
                )
            } else {
                # Fallback: Default Domain Policy
                $userResultList = @(
                    [PSCustomObject]@{ $t["PropCategory"] = $t["CatAppliedPolicy"]; $t["PropName"] = $t["PropSource"]; $t["PropValue"] = $t["ValDefaultPolicy"] }
                )
                foreach ($row in $script:cachedDomainPolicies) {
                    $userResultList += $row
                }
            }

            $script:cachedUserPolicy = $userResultList
            $gridUserPolicy.DataSource = [System.Collections.ArrayList]::new($userResultList)
            $lblStatus.Text = "$($t['StatusResultReady']): $foundSam"

        } catch {
            $lblStatus.Text = "$($t['StatusError']): $($_.Exception.Message)"
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }

    # Event Wiring
    $btnCheckUser.Add_Click($checkUserAction)
    $txtUserCheck.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $_.SuppressKeyPress = $true
            & $checkUserAction
        }
    })

    # Export CSV
    $btnExport.Add_Click({
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "CSV-Datei (*.csv)|*.csv"
        $sfd.FileName = "AD_Password_Policies_Audit.csv"
        if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $exportData = @()
            foreach ($row in $script:cachedDomainPolicies) {
                $exportData += [PSCustomObject]@{
                    "Typ"         = "Default Domain Policy"
                    "Eigenschaft" = $row.($t["PropName"])
                    "Wert"        = $row.($t["PropValue"])
                }
            }
            foreach ($row in $script:cachedPSOs) {
                $exportData += [PSCustomObject]@{
                    "Typ"         = "Fine-Grained PSO ($($row.($t['PsoName'])))"
                    "Eigenschaft" = "Precedence: $($row.($t['PsoPrecedence'])), MinLen: $($row.($t['PsoMinLength']))"
                    "Wert"        = "AppliesTo: $($row.($t['PsoAppliesTo']))"
                }
            }
            $exportData | Export-Csv -Path $sfd.FileName -NoTypeInformation -Delimiter ";" -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show("Kennwortrichtlinien exportiert:`n$($sfd.FileName)", "Export OK", "OK", "Information")
        }
    })

    $btnRefresh.Add_Click({ & $loadPolicies })
    $form.Add_Shown({ & $loadPolicies })

    [void]$form.ShowDialog()
}
# ==============================================================================
# HAUPTLAUNCHER & DASHBOARD
# ==============================================================================
$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Text = Get-Text "Title"
$mainForm.Size = New-Object System.Drawing.Size(650, 1030)
$mainForm.FormBorderStyle = "FixedDialog"
$mainForm.MaximizeBox = $false
$mainForm.StartPosition = "CenterScreen"

$pnlHeader = New-Object System.Windows.Forms.Panel
$pnlHeader.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlHeader.Height = 235
$pnlHeader.BackColor = [System.Drawing.Color]::FromArgb(230, 238, 248)
$mainForm.Controls.Add($pnlHeader)

$btnLangEN = New-Object System.Windows.Forms.Button
$btnLangEN.Location = "510, 10"; $btnLangEN.Size = "45, 25"; $btnLangEN.Text = "EN"
$btnLangEN.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 8, [System.Drawing.FontStyle]::Bold)
$btnLangEN.BackColor = [System.Drawing.Color]::LightSteelBlue
$pnlHeader.Controls.Add($btnLangEN)

$btnLangDE = New-Object System.Windows.Forms.Button
$btnLangDE.Location = "560, 10"; $btnLangDE.Size = "45, 25"; $btnLangDE.Text = "DE"
$btnLangDE.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 8, [System.Drawing.FontStyle]::Bold)
$pnlHeader.Controls.Add($btnLangDE)

$lblHeaderTitle = New-Object System.Windows.Forms.Label
$lblHeaderTitle.Location = "15, 10"; $lblHeaderTitle.Size = "480, 20"
$lblHeaderTitle.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 9.5, [System.Drawing.FontStyle]::Bold)
$pnlHeader.Controls.Add($lblHeaderTitle)

$lblHeaderLogon = New-Object System.Windows.Forms.Label
$lblHeaderLogon.Location = "15, 30"; $lblHeaderLogon.Size = "480, 20"
$pnlHeader.Controls.Add($lblHeaderLogon)

$lblHeaderStatus = New-Object System.Windows.Forms.Label
$lblHeaderStatus.Location = "15, 50"; $lblHeaderStatus.Size = "590, 20"
$lblHeaderStatus.ForeColor = [System.Drawing.Color]::DarkBlue
$pnlHeader.Controls.Add($lblHeaderStatus)

$lblLine = New-Object System.Windows.Forms.Label
$lblLine.Location = "15, 75"; $lblLine.Size = "590, 2"; $lblLine.BorderStyle = "Fixed3D"
$pnlHeader.Controls.Add($lblLine)

$lblOSHeader = New-Object System.Windows.Forms.Label
$lblOSHeader.Location = "15, 83"; $lblOSHeader.Size = "590, 140"
$lblOSHeader.Font = New-Object System.Drawing.Font("Consolas", 8.5)
$pnlHeader.Controls.Add($lblOSHeader)

# Client-Tools Gruppe (Tools 3, 4, 5, 8)
$grpClient = New-Object System.Windows.Forms.GroupBox
$grpClient.Location = "20, 245"; $grpClient.Size = "590, 280"
$grpClient.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 9, [System.Drawing.FontStyle]::Bold)
$mainForm.Controls.Add($grpClient)

$btnTool3 = New-Object System.Windows.Forms.Button
$btnTool3.Location = "20, 28"; $btnTool3.Size = "550, 48"
$btnTool3.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 8.5, [System.Drawing.FontStyle]::Regular)
$grpClient.Controls.Add($btnTool3)

$btnTool4 = New-Object System.Windows.Forms.Button
$btnTool4.Location = "20, 84"; $btnTool4.Size = "550, 48"
$btnTool4.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 8.5, [System.Drawing.FontStyle]::Regular)
$grpClient.Controls.Add($btnTool4)

$btnTool5 = New-Object System.Windows.Forms.Button
$btnTool5.Location = "20, 140"; $btnTool5.Size = "550, 48"
$btnTool5.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 8.5, [System.Drawing.FontStyle]::Regular)
$grpClient.Controls.Add($btnTool5)

$btnTool8 = New-Object System.Windows.Forms.Button
$btnTool8.Location = "20, 196"; $btnTool8.Size = "550, 48"
$btnTool8.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 8.5, [System.Drawing.FontStyle]::Regular)
$grpClient.Controls.Add($btnTool8)

# AD-Tools Gruppe (Tools 1, 2, 6, 7, 9, 11)
$grpAD = New-Object System.Windows.Forms.GroupBox
$grpAD.Location = "20, 535"; $grpAD.Size = "590, 395"
$grpAD.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 9, [System.Drawing.FontStyle]::Bold)
$mainForm.Controls.Add($grpAD)

$btnTool1 = New-Object System.Windows.Forms.Button
$btnTool1.Location = "20, 28"; $btnTool1.Size = "550, 48"
$btnTool1.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 8.5, [System.Drawing.FontStyle]::Regular)
$grpAD.Controls.Add($btnTool1)

$btnTool2 = New-Object System.Windows.Forms.Button
$btnTool2.Location = "20, 84"; $btnTool2.Size = "550, 48"
$btnTool2.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 8.5, [System.Drawing.FontStyle]::Regular)
$grpAD.Controls.Add($btnTool2)

$btnTool6 = New-Object System.Windows.Forms.Button
$btnTool6.Location = "20, 140"; $btnTool6.Size = "550, 48"
$btnTool6.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 8.5, [System.Drawing.FontStyle]::Regular)
$grpAD.Controls.Add($btnTool6)

$btnTool7 = New-Object System.Windows.Forms.Button
$btnTool7.Location = "20, 196"; $btnTool7.Size = "550, 48"
$btnTool7.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 8.5, [System.Drawing.FontStyle]::Regular)
$grpAD.Controls.Add($btnTool7)

$btnTool9 = New-Object System.Windows.Forms.Button
$btnTool9.Location = "20, 252"; $btnTool9.Size = "550, 48"
$btnTool9.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 8.5, [System.Drawing.FontStyle]::Regular)
$grpAD.Controls.Add($btnTool9)

$btnTool11 = New-Object System.Windows.Forms.Button
$btnTool11.Location = "20, 308"; $btnTool11.Size = "550, 48"
$btnTool11.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 8.5, [System.Drawing.FontStyle]::Regular)
$grpAD.Controls.Add($btnTool11)

# UI-Aktualisierung
function Update-UI {
    $mainForm.Text          = Get-Text "Title"
    $grpClient.Text         = Get-Text "CategoryClient"
    $grpAD.Text             = Get-Text "CategoryAD"

    $lblHeaderTitle.Text    = "$(Get-Text 'Computer'): $localComputerName | $(Get-Text 'Domain'): $localDomainName"
    $lblHeaderLogon.Text    = "$(Get-Text 'LogonServer'): $localLogonServer | $(Get-Text 'User'): $localUserName"
    $lblHeaderStatus.Text   = "$(Get-Text 'EntraStatus'): $localJoinStatus | AzureAdPrt: $localAzureAdPrt"

    $lblOSHeader.Text = @"
OS Name:            $osNameFormatted
Version / Release:  $osVersionRelease
Build Number:       $osBuildFormatted
Architecture:       $osArchFormatted
Install Date:       $osInstallDateForm
Last Update:        $osPatchFormatted
"@

    $btnTool1.Text = Get-Text "BtnTool1"
    $btnTool2.Text = Get-Text "BtnTool2"
    $btnTool3.Text = Get-Text "BtnTool3"
    $btnTool4.Text = Get-Text "BtnTool4"
    $btnTool5.Text = Get-Text "BtnTool5"
    $btnTool6.Text = Get-Text "BtnTool6"
    $btnTool7.Text = Get-Text "BtnTool7"
    $btnTool8.Text = Get-Text "BtnTool8"
    $btnTool9.Text = Get-Text "BtnTool9"
    $btnTool11.Text = Get-Text "BtnTool11"

    if ($script:CurrentLang -eq "EN") {
        $btnLangEN.BackColor = [System.Drawing.Color]::LightSteelBlue
        $btnLangDE.BackColor = [System.Drawing.SystemColors]::Control
    } else {
        $btnLangDE.BackColor = [System.Drawing.Color]::LightSteelBlue
        $btnLangEN.BackColor = [System.Drawing.SystemColors]::Control
    }
}

# Sprachumschaltung
$btnLangEN.Add_Click({ $script:CurrentLang = "EN"; Update-UI })
$btnLangDE.Add_Click({ $script:CurrentLang = "DE"; Update-UI })

# Tool-Trigger
$btnTool3.Add_Click({ Open-ToolEntraStatus })
$btnTool4.Add_Click({ Open-ToolGroupsAndGPO })
$btnTool5.Add_Click({ Open-ToolWin11Check })
$btnTool8.Add_Click({ Show-ClientSoftwareAnalysis })

$btnTool1.Add_Click({ Open-ToolLastLogon })
$btnTool2.Add_Click({ Open-ToolADAudit })
$btnTool6.Add_Click({ Open-ToolDomainOverview })
$btnTool7.Add_Click({ Open-ToolOSSupportAudit })
$btnTool9.Add_Click({ Show-Tool9-ACLCompare })
$btnTool11.Add_Click({ Show-Tool11-PasswordPolicyAudit })

# Dashboard starten
Update-UI
[void]$mainForm.ShowDialog()
