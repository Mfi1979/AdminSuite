<#
================================================================================
 ACTIVE DIRECTORY & ENTRA ID ADMIN SUITE (VOLLSTÄNDIGE GESAMTEDITION - LETZTSTAND)
 Module: Tools 1 bis 11 (Vollumfänglich integriert) | DE & EN | Native LDAP & Forms
================================================================================
 Inklusive aller Werkzeuge & Module:
  - 3-Spalten Header: OS & Domäne, Hardware & System, Entra ID Registry / Cloud Status
  - Tool 1: Multi-DC LastLogon Übersicht (Live DC Query & Inaktivitätsanalyse)
  - Tool 2: Quick AD Audit (Deaktivierte Konten, Password never expires, leere Gruppen)
  - Tool 3: Entra ID / Hybrid Join Diagnostic (dsregcmd /status & Tenant Details)
  - Tool 4: Gruppen & GPO Diagnostik (3 Register: User-Gruppen, PC-Gruppen, gpresult)
  - Tool 5: Windows 11 Readiness & Hardware Check (TPM 2.0, SecureBoot, RAM, CPU, OS)
  - Tool 6: Domänen-Übersicht & Admin-Audit (3 Tabs: DCs, FSMO-Rollen, Privilegierte Admins)
  - Tool 7: AD Security & OS Support Audit (Dynamischer OS-Filter, EOL-Matrix & LTSC/Build-Erkennung)
  - Tool 8: Client Software & App Analyse (Registry Win32 & AppX Store Apps mit Sprachen)
  - Tool 9: AD ACL & Berechtigungsvergleich (Universelles Diff für User, Gruppen, Computer, OUs)
  - Tool 10: Active Directory OU & Gruppen Finder (inkl. adminCount, UPN & Gruppen-Filter)
  - Tool 11: AD Kennwortrichtlinien & PSO Audit (Bewährte funktionierende ADSI/COM-Engine)
  - Zentrales UITheme Layout ($script:UITheme & Apply-StandardGridTheme)
================================================================================
#>

# Windows Forms, Drawing & DirectoryServices laden
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.DirectoryServices
[System.Windows.Forms.Application]::EnableVisualStyles()

# ==============================================================================
# ZENTRALE LAYOUT- UND DESIGN-KONFIGURATION
# ==============================================================================
$script:UITheme = @{
    # --- Tabellen / DataGridView Layout ---
    HeaderHeight          = 36          # Höhe der Spaltenüberschriften in Pixel
    RowHeight             = 28          # Höhe jeder Datenzeile in Pixel
    HeaderPaddingLeft     = 8           # Innenabstand links im Header
    HeaderPaddingRight    = 8           # Innenabstand rechts im Header
    CellPaddingLeft       = 8           # Innenabstand links in den Zellen
    CellPaddingRight      = 8           # Innenabstand rechts in den Zellen
    CellPaddingTop        = 2           # Innenabstand oben in den Zellen
    CellPaddingBottom     = 2           # Innenabstand unten in den Zellen
    
    # --- Schriftarten ---
    FontFamily            = "Segoe UI"  # Standard-Schriftart
    HeaderFontSize        = 9.5         # Schriftgröße Header (pt)
    CellFontSize          = 9.0         # Schriftgröße Tabelleninhalt (pt)
    HeaderFontStyle       = [System.Drawing.FontStyle]::Bold
    CellFontStyle         = [System.Drawing.FontStyle]::Regular
    
    # --- Farbpalette Header & Tabellen ---
    HeaderBackColor       = [System.Drawing.Color]::FromArgb(238, 242, 246)
    HeaderForeColor       = [System.Drawing.Color]::FromArgb(40, 40, 40)
    GridLineColor         = [System.Drawing.Color]::FromArgb(226, 232, 240)
    RowBackColor          = [System.Drawing.Color]::White
    RowAltBackColor       = [System.Drawing.Color]::FromArgb(250, 252, 254)
    SelectionBackColor    = [System.Drawing.Color]::FromArgb(203, 228, 249)
    SelectionForeColor    = [System.Drawing.Color]::Black

    # --- Akzent-Farben ---
    AccentColor           = [System.Drawing.Color]::FromArgb(0, 120, 215)
    AccentColorDark       = [System.Drawing.Color]::FromArgb(24, 37, 55)

    # --- Standard-Fensterabmessungen ---
    DefaultToolWidth      = 1200
    DefaultToolHeight     = 780
    HeaderPanelHeight     = 60
}

# ==============================================================================
# HILFSFUNKTION: GLOBALER GRID-LAYOUT-STYLING-APPLIKATOR
# ==============================================================================
function Apply-StandardGridTheme {
    param(
        [Parameter(Mandatory=$true)]
        [System.Windows.Forms.DataGridView]$Grid,
        [switch]$EnableAlternatingRowColor
    )

    $theme = $script:UITheme

    # Basis-Verhalten
    $Grid.EnableHeadersVisualStyles = $false
    $Grid.BorderStyle               = [System.Windows.Forms.BorderStyle]::None
    $Grid.CellBorderStyle           = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
    $Grid.GridColor                 = $theme.GridLineColor
    $Grid.BackgroundColor           = [System.Drawing.Color]::White
    $Grid.RowHeadersVisible         = $false
    $Grid.AllowUserToResizeRows     = $false
    $Grid.SelectionMode             = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $Grid.ReadOnly                  = $true

    # 1. Spaltenüberschriften (Header) Styling
    $Grid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
    $Grid.ColumnHeadersHeight       = $theme.HeaderHeight
    
    $headerFont = New-Object System.Drawing.Font($theme.FontFamily, $theme.HeaderFontSize, $theme.HeaderFontStyle)
    $Grid.ColumnHeadersDefaultCellStyle.Font      = $headerFont
    $Grid.ColumnHeadersDefaultCellStyle.BackColor = $theme.HeaderBackColor
    $Grid.ColumnHeadersDefaultCellStyle.ForeColor = $theme.HeaderForeColor
    $Grid.ColumnHeadersDefaultCellStyle.Alignment = [System.Drawing.ContentAlignment]::MiddleLeft
    $Grid.ColumnHeadersDefaultCellStyle.Padding   = New-Object System.Windows.Forms.Padding(
        $theme.HeaderPaddingLeft, 0, $theme.HeaderPaddingRight, 0
    )

    # 2. Datenzeilen (Cells) Styling
    $cellFont = New-Object System.Drawing.Font($theme.FontFamily, $theme.CellFontSize, $theme.CellFontStyle)
    $Grid.RowTemplate.Height                   = $theme.RowHeight
    $Grid.DefaultCellStyle.Font                = $cellFont
    $Grid.DefaultCellStyle.BackColor           = $theme.RowBackColor
    $Grid.DefaultCellStyle.SelectionBackColor  = $theme.SelectionBackColor
    $Grid.DefaultCellStyle.SelectionForeColor  = $theme.SelectionForeColor
    $Grid.DefaultCellStyle.Alignment          = [System.Drawing.ContentAlignment]::MiddleLeft
    $Grid.DefaultCellStyle.Padding            = New-Object System.Windows.Forms.Padding(
        $theme.CellPaddingLeft, $theme.CellPaddingTop, $theme.CellPaddingRight, $theme.CellPaddingBottom
    )

    # Optional: Alternierende Zeilenfarben
    if ($EnableAlternatingRowColor) {
        $Grid.AlternatingRowsDefaultCellStyle.BackColor          = $theme.RowAltBackColor
        $Grid.AlternatingRowsDefaultCellStyle.SelectionBackColor = $theme.SelectionBackColor
        $Grid.AlternatingRowsDefaultCellStyle.SelectionForeColor = $theme.SelectionForeColor
    }
}

# ==============================================================================
# 0. MULTILANGUAGE DICTIONARY (I18N: DE / EN)
# ==============================================================================
$script:CurrentLang = "DE" # Standard: DE

$script:I18N = @{
    "DE" = @{
        "Title"             = "Active Directory & Entra ID Admin Suite"
        "CategoryClient"    = "Kategorie: Client & Lokale Diagnose (Client Tools)"
        "CategoryAD"        = "Kategorie: Active Directory & Domänen Diagnose (AD Tools)"
        "LblHdrOS"          = "BETRIEBSSYSTEM & DOMÄNE"
        "LblHdrSystem"      = "SYSTEM & HARDWARE"
        "LblHdrEntra"       = "ENTRA ID / CLOUD STATUS"
        "LblOS"             = "OS Edition"
        "LblBuild"          = "Build"
        "LblVersion"        = "Version"
        "LblDomain"         = "Domäne / AD"
        "LblLogonServer"    = "Logonserver"
        "LblCompName"       = "Computername"
        "LblManuf"          = "Hersteller"
        "LblModel"          = "Modell"
        "LblSerial"         = "Seriennummer"
        "LblSysType"        = "Systemtyp"
        "LblJoinStatus"     = "Join-Status"
        "LblPrtStatus"      = "AzureAD PRT"
        "LblTenantName"     = "Tenant Name"
        "LblTenantId"       = "Tenant ID"
        "LblDeviceId"       = "Device ID"
        "BtnTool1"          = "Tool 1: Multi-DC LastLogon Übersicht`n(Automatischer DC- & Site-Scan mit Checkbox-Auswahl & Live-Abfrage)"
        "BtnTool2"          = "Tool 2: Quick AD Audit`n(Inaktive Computerkonten, Passwort nie abgelaufen, Leere Gruppen)"
        "BtnTool3"          = "Tool 3: Entra ID / Hybrid Join Diagnostic`n(Auslesen von dsregcmd /status & Tenant Details)"
        "BtnTool4"          = "Tool 4: Gruppen & GPO Diagnostik`n(Register: User-Gruppen, PC-Gruppen & angewendete Richtlinien)"
        "BtnTool5"          = "Tool 5: Windows 11 Readiness & OS Details`n(Hardware-Kompatibilität, TPM 2.0, SecureBoot & Patch-Stand)"
        "BtnTool6"          = "Tool 6: Domänen-Übersicht & Admin-Audit`n(FSMO-Rollen, AD-Papierkorb, Objekt-Anzahl & Privilegierte Admins)"
        "BtnTool7"          = "Tool 7: AD Security & OS Support Audit`n(Identifikation veralteter & nicht mehr unterstützter Systeme)"
        "BtnTool8"          = "Tool 8: Client Software & App Analyse`n(Win32- & Store-Apps mit Sprachauswertung und CSV-Export)"
        "BtnTool9"          = "Tool 9: AD ACL & Berechtigungsvergleich`n(Objekt-Berechtigungen von Usern, Gruppen, Computern oder OUs vergleichen)"
        "BtnTool10"         = "Tool 10: Active Directory OU & Gruppen Finder`n(Objektsuche nach Clients, Servern, Usern inkl. adminCount, UPN & Gruppen-Filter)"
        "BtnTool11"         = "Tool 11: AD Kennwortrichtlinien & PSO Audit`n(Default Domain Policy, Fine-Grained PSOs & Benutzer-Check)"
        "ErrNoDomain"       = "Dieses Werkzeug erfordert eine Active Directory Domänenmitgliedschaft."
    }
    "EN" = @{
        "Title"             = "Active Directory & Entra ID Admin Suite"
        "CategoryClient"    = "Category: Client & Local Diagnostics (Client Tools)"
        "CategoryAD"        = "Category: Active Directory & Domain Diagnostics (AD Tools)"
        "LblHdrOS"          = "OPERATING SYSTEM & DOMAIN"
        "LblHdrSystem"      = "SYSTEM & HARDWARE"
        "LblHdrEntra"       = "ENTRA ID / CLOUD STATUS"
        "LblOS"             = "OS Edition"
        "LblBuild"          = "Build"
        "LblVersion"        = "Version"
        "LblDomain"         = "Domain / AD"
        "LblLogonServer"    = "Logon Server"
        "LblCompName"       = "Computer Name"
        "LblManuf"          = "Manufacturer"
        "LblModel"          = "Model"
        "LblSerial"         = "Serial Number"
        "LblSysType"        = "System Type"
        "LblJoinStatus"     = "Join Status"
        "LblPrtStatus"      = "AzureAD PRT"
        "LblTenantName"     = "Tenant Name"
        "LblTenantId"       = "Tenant ID"
        "LblDeviceId"       = "Device ID"
        "BtnTool1"          = "Tool 1: Multi-DC LastLogon Overview`n(Automatic DC & Site Scan with Checkbox Selection & Live Query)"
        "BtnTool2"          = "Tool 2: Quick AD Audit`n(Disabled Accounts, Password Never Expires, Empty Groups)"
        "BtnTool3"          = "Tool 3: Entra ID / Hybrid Join Diagnostic`n(Read dsregcmd /status & Tenant Details)"
        "BtnTool4"          = "Tool 4: Groups & GPO Diagnostics`n(Tabs: User Groups, PC Groups & Applied GPOs)"
        "BtnTool5"          = "Tool 5: Windows 11 Readiness & OS Details`n(Hardware Compatibility, TPM 2.0, SecureBoot & Patch Level)"
        "BtnTool6"          = "Tool 6: Domain Overview & Admin Audit`n(FSMO Roles, AD Recycle Bin, Object Counts & Privileged Admins)"
        "BtnTool7"          = "Tool 7: AD Security & Out-of-Support OS Audit`n(Identify Legacy & End-of-Life Windows Systems in AD)"
        "BtnTool8"          = "Tool 8: Client Software & App Analysis`n(Win32 Registry & Store Apps with Language Detection & Export)"
        "BtnTool9"          = "Tool 9: AD ACL & Permission Diff Tool`n(Compare object permissions of Users, Groups, Computers or OUs)"
        "BtnTool10"         = "Tool 10: Active Directory OU & Group Finder`n(Search Clients, Servers, Users with adminCount, UPN & Group Filtering)"
        "BtnTool11"         = "Tool 11: AD Password Policies & PSO Audit`n(Default Domain Policy, Fine-Grained PSOs & User Check)"
        "ErrNoDomain"       = "This tool requires Active Directory domain membership."
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
    $isDomain = (Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue).PartOfDomain
    if (-not $isDomain) {
        [System.Windows.Forms.MessageBox]::Show((Get-Text "ErrNoDomain"), "Info", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return $false
    }
    return $true
}

function Get-DomainDN {
    try {
        $rootDSE = [ADSI]"LDAP://RootDSE"
        return $rootDSE.defaultNamingContext.ToString()
    } catch {
        return $null
    }
}

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

# Dynamische OS Support & EOL-Erkennungs-Engine
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
    } elseif ($OSVersion -and $OSVersion -ne "Unspecified OS") {
        $buildNumber = $OSVersion
    }

    # Fallback: Build-Nummer aus OS-Namen erkennen, falls OSVersion im AD leer ist
    if ($buildNumber -eq "Unbekannt" -or [string]::IsNullOrWhiteSpace($buildNumber)) {
        if ($OSName -match "24H2") { $buildNumber = "26100" }
        elseif ($OSName -match "23H2") { $buildNumber = "22631" }
        elseif ($OSName -match "22H2" -and $OSName -like "*Windows 11*") { $buildNumber = "22621" }
        elseif ($OSName -match "21H2" -and $OSName -like "*Windows 11*") { $buildNumber = "22000" }
        elseif ($OSName -match "22H2" -and $OSName -like "*Windows 10*") { $buildNumber = "19045" }
        elseif ($OSName -match "2021" -or ($OSName -like "*LTSC*" -and $OSName -match "21H2")) { $buildNumber = "19044" }
        elseif ($OSName -match "2019" -or ($OSName -like "*LTSC*" -and $OSName -match "1809")) { $buildNumber = "17763" }
        elseif ($OSName -match "2016" -or ($OSName -like "*LTSB*" -and $OSName -match "1607")) { $buildNumber = "14393" }
        elseif ($OSName -match "2015" -or ($OSName -like "*LTSB*" -and $OSName -match "1507")) { $buildNumber = "10240" }
    }

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
        if ($OSName -like "*LTSC 2021*") { $clientVersion = "21H2 / LTSC 2021" }
        elseif ($OSName -like "*LTSC 2019*") { $clientVersion = "1809 / LTSC 2019" }
        elseif ($OSName -like "*LTSB 2016*") { $clientVersion = "1607 / LTSB 2016" }
        elseif ($OSName -like "*Server 2022*") { $clientVersion = "Server 2022" }
        elseif ($OSName -like "*Server 2019*") { $clientVersion = "Server 2019" }
        elseif ($OSName -like "*Server 2016*") { $clientVersion = "Server 2016" }
        elseif ($OSName -like "*Server 2012*") { $clientVersion = "Server 2012" }
        elseif ($OSServicePack) { $clientVersion = $OSServicePack }
        elseif ($OSVersion) { $clientVersion = $OSVersion }
    }

    $isEnterpriseOrEdu = ($OSName -like "*Enterprise*" -or $OSName -like "*Education*")
    $isLTSCOrLTSB      = ($isLTSC -or $clientVersion -like "*LTSC*" -or $clientVersion -like "*LTSB*")

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

# ==============================================================================
# LOKALE SYSTEMPARAMETER, HARDWARE- & ENTRA-DATEN AUSLESEN
# ==============================================================================
$localComputerName = $env:COMPUTERNAME
$localUserName     = $env:USERNAME

# Hardware- & BIOS-Details
$cs   = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue
$bios = Get-CimInstance -ClassName Win32_BIOS -ErrorAction SilentlyContinue

$localManufacturer = if ($cs.Manufacturer) { $cs.Manufacturer } else { "N/A" }
$localModel        = if ($cs.Model) { $cs.Model } else { "N/A" }
$localSystemType   = if ($cs.SystemType) { $cs.SystemType } else { "N/A" }
$localSerial       = if ($bios.SerialNumber) { $bios.SerialNumber } else { "N/A" }

# Domäne & Logonserver
$localDomainName   = if ($cs.PartOfDomain -and $cs.Domain) { $cs.Domain } else { "WORKGROUP" }
$localLogonServer  = if ($env:LOGONSERVER) { $env:LOGONSERVER.TrimStart('\') } else { "Lokal" }

# Entra ID / Hybrid Join Details (Registry CloudDomainJoin + dsregcmd)
$localJoinStatus = "Nicht gekoppelt"
$localAzureAdPrt = "NO"
$localTenantName = "N/A"
$localTenantId   = "N/A"
$localDeviceId   = "N/A"

try {
    $cdjJoin = Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\CloudDomainJoin\JoinInfo\*" -ErrorAction SilentlyContinue
    if ($cdjJoin) {
        if ($cdjJoin.TenantId) { $localTenantId = $cdjJoin.TenantId }
        if ($cdjJoin.DeviceId) { $localDeviceId = $cdjJoin.DeviceId }
    }
    $cdjTenant = Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\CloudDomainJoin\TenantInfo\*" -ErrorAction SilentlyContinue
    if ($cdjTenant -and $cdjTenant.DisplayName) {
        $localTenantName = $cdjTenant.DisplayName
    }
} catch {}

try {
    $dsreg = dsregcmd /status 2>$null
    if ($dsreg -match "AzureAdJoined\s*:\s*YES")    { $localJoinStatus = "Azure AD Joined" }
    if ($dsreg -match "EnterpriseJoined\s*:\s*YES") { $localJoinStatus = "Hybrid Joined" }
    if ($dsreg -match "DomainJoined\s*:\s*YES" -and $localJoinStatus -eq "Nicht gekoppelt") { 
        $localJoinStatus = "AD Domain Joined" 
    }
    if ($dsreg -match "AzureAdPrt\s*:\s*YES")       { $localAzureAdPrt = "YES" }

    if ($localTenantName -eq "N/A" -and ($dsreg -match "TenantName\s*:\s*(.+)"))           { $localTenantName = $matches[1].Trim() }
    if ($localTenantId -eq "N/A"   -and ($dsreg -match "TenantId\s*:\s*([a-fA-F0-9\-]+)")) { $localTenantId   = $matches[1].Trim() }
    if ($localDeviceId -eq "N/A"   -and ($dsreg -match "DeviceId\s*:\s*([a-fA-F0-9\-]+)")) { $localDeviceId   = $matches[1].Trim() }
} catch {}

# Betriebssystem Details mit Fallback
$osInfo            = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue
$osCaption         = if ($osInfo.Caption) { ($osInfo.Caption -replace "Microsoft ", "").Trim() } else { "Windows" }
$osBuildNumber     = if ($osInfo.BuildNumber) { $osInfo.BuildNumber } else { "N/A" }
$osBuildFull       = "$($osInfo.Version) (Build $osBuildNumber)"
$osArchFormatted   = $osInfo.OSArchitecture
$osInstallDateForm = if ($osInfo.InstallDate) { $osInfo.InstallDate.ToString("dd.MM.yyyy HH:mm") } else { "N/A" }
$osPatchFormatted  = (Get-HotFix -ErrorAction SilentlyContinue | Sort-Object InstalledOn -Descending | Select-Object -First 1).HotFixID

$osVersionDisplay = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name DisplayVersion -ErrorAction SilentlyContinue).DisplayVersion
if (-not $osVersionDisplay) {
    $osVersionDisplay = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name ReleaseId -ErrorAction SilentlyContinue).ReleaseId
}
if (-not $osVersionDisplay) {
    $bNum = 0
    [int]::TryParse($osBuildNumber, [ref]$bNum) | Out-Null
    $osVersionDisplay = switch ($bNum) {
        26100 { "24H2" }
        22631 { "23H2" }
        22621 { "22H2" }
        22000 { "21H2" }
        19045 { "22H2" }
        19044 { "21H2" }
        19043 { "21H1" }
        19042 { "20H2" }
        17763 { "1809" }
        14393 { "1607" }
        default { if ($osInfo.Version) { $osInfo.Version } else { "N/A" } }
    }
}

# ===================================================================
# MODULE TOOL 1: MULTI-DC LASTLOGON
# ===================================================================
function Open-ToolLastLogon {
    if (-not (Assert-DomainJoined)) { return }

    $subForm = New-Object System.Windows.Forms.Form
    $subForm.Text = "Tool 1: Multi-DC LastLogon Analyse"
    $subForm.Size = New-Object System.Drawing.Size(1280, 800)
    $subForm.StartPosition = "CenterParent"
    $subForm.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9)

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
    Apply-StandardGridTheme -Grid $gridDCs
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
    Apply-StandardGridTheme -Grid $gridResults
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
                    }; Descending = $false 
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
    $subForm.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9)

    $pnlTop = New-Object System.Windows.Forms.Panel
    $pnlTop.Dock = "Top"; $pnlTop.Height = 60; $subForm.Controls.Add($pnlTop)

    $cmbQuery = New-Object System.Windows.Forms.ComboBox
    $cmbQuery.DropDownStyle = "DropDownList"; $cmbQuery.Location = "15, 18"; $cmbQuery.Width = 300
    [void]$cmbQuery.Items.AddRange(@("Deaktivierte Computerkonten", "Passwort läuft nie ab (User)", "Leere AD-Gruppen"))
    $cmbQuery.SelectedIndex = 0; $pnlTop.Controls.Add($cmbQuery)

    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Text = "Abfrage starten"; $btnRun.Location = "330, 16"; $btnRun.Size = "130, 28"; $pnlTop.Controls.Add($btnRun)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = "Fill"; $grid.ReadOnly = $true; $grid.AutoSizeColumnsMode = "Fill"; 
    Apply-StandardGridTheme -Grid $grid -EnableAlternatingRowColor
    $subForm.Controls.Add($grid); $grid.BringToFront()

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
    $subForm.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9)

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
    Apply-StandardGridTheme -Grid $gridUserGroups -EnableAlternatingRowColor
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
    Apply-StandardGridTheme -Grid $gridPCGroups -EnableAlternatingRowColor
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
    $subForm.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = "Fill"; $grid.ReadOnly = $true; $grid.AutoSizeColumnsMode = "Fill"
    Apply-StandardGridTheme -Grid $grid -EnableAlternatingRowColor
    $subForm.Controls.Add($grid)

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
        [PSCustomObject]@{ "Prüfkriterium" = "Installiertes OS"; "Status" = "$osCaption (Build $osBuildFull)"; "Erforderlich" = "Windows 10/11" }
    )
    $arr = [System.Collections.ArrayList]::new()
    foreach ($c in $checks) { [void]$arr.Add($c) }
    $grid.DataSource = $arr
    [void]$subForm.ShowDialog()
}

# ==============================================================================
# TOOL 6: DOMAIN OVERVIEW & PRIVILEGED ADMIN AUDIT
# ==============================================================================
function Open-ToolDomainOverview {
    if (-not (Assert-DomainJoined)) { return }

    $subForm = New-Object System.Windows.Forms.Form
    $subForm.Text = "Tool 6 - Active Directory Domänen-Übersicht & Admin-Audit"
    $subForm.Size = New-Object System.Drawing.Size(1200, 780)
    $subForm.StartPosition = "CenterParent"
    $subForm.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9)
    $subForm.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Dock = [System.Windows.Forms.DockStyle]::Fill
    $tabControl.Padding = New-Object System.Drawing.Point(12, 6)

    # --- TAB 1: DOMÄNENCONTROLLER (DCs) ---
    $tabDCs = New-Object System.Windows.Forms.TabPage
    $tabDCs.Text = if ($script:CurrentLang -eq "DE") { "  🖥️ Domänencontroller (DCs)  " } else { "  🖥️ Domain Controllers (DCs)  " }
    $tabDCs.BackColor = [System.Drawing.Color]::White

    $pnlTopDC = New-Object System.Windows.Forms.Panel
    $pnlTopDC.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlTopDC.Height = 44
    $pnlTopDC.BackColor = [System.Drawing.Color]::FromArgb(248, 249, 250)

    $lblDCCount = New-Object System.Windows.Forms.Label
    $lblDCCount.Text = if ($script:CurrentLang -eq "DE") { "Erkannte Domain Controller in der Gesamtstruktur:" } else { "Discovered Domain Controllers in Forest:" }
    $lblDCCount.Location = New-Object System.Drawing.Point(12, 13)
    $lblDCCount.AutoSize = $true
    $lblDCCount.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9.5, [System.Drawing.FontStyle]::Bold)

    $btnRefreshDC = New-Object System.Windows.Forms.Button
    $btnRefreshDC.Text = if ($script:CurrentLang -eq "DE") { "Aktualisieren" } else { "Refresh" }
    $btnRefreshDC.Location = New-Object System.Drawing.Point(420, 8)
    $btnRefreshDC.Size = New-Object System.Drawing.Size(120, 28)
    $btnRefreshDC.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $btnRefreshDC.ForeColor = [System.Drawing.Color]::White
    $btnRefreshDC.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat

    $pnlTopDC.Controls.AddRange(@($lblDCCount, $btnRefreshDC))

    $gridDCs = New-Object System.Windows.Forms.DataGridView
    $gridDCs.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridDCs.ReadOnly = $true
    $gridDCs.AllowUserToAddRows = $false
    $gridDCs.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    Apply-StandardGridTheme -Grid $gridDCs -EnableAlternatingRowColor

    $tabDCs.Controls.Add($gridDCs)
    $tabDCs.Controls.Add($pnlTopDC)
    $pnlTopDC.SendToBack()
    $gridDCs.BringToFront()

    # --- TAB 2: FSMO ROLLEN & DOMÄNENSTATUS ---
    $tabFSMO = New-Object System.Windows.Forms.TabPage
    $tabFSMO.Text = if ($script:CurrentLang -eq "DE") { "  ⚙️ FSMO-Rollen & Domänenstatus  " } else { "  ⚙️ FSMO Roles & Domain Status  " }
    $tabFSMO.BackColor = [System.Drawing.Color]::White

    $pnlTopFSMO = New-Object System.Windows.Forms.Panel
    $pnlTopFSMO.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlTopFSMO.Height = 44
    $pnlTopFSMO.BackColor = [System.Drawing.Color]::FromArgb(248, 249, 250)

    $lblFSMOHead = New-Object System.Windows.Forms.Label
    $lblFSMOHead.Text = if ($script:CurrentLang -eq "DE") { "Zentrale FSMO Rolleninhaber & Gesamtstruktur-Ebenen:" } else { "FSMO Role Holders & Forest Levels:" }
    $lblFSMOHead.Location = New-Object System.Drawing.Point(12, 13)
    $lblFSMOHead.AutoSize = $true
    $lblFSMOHead.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9.5, [System.Drawing.FontStyle]::Bold)
    $pnlTopFSMO.Controls.Add($lblFSMOHead)

    $gridFSMO = New-Object System.Windows.Forms.DataGridView
    $gridFSMO.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridFSMO.ReadOnly = $true
    $gridFSMO.AllowUserToAddRows = $false
    $gridFSMO.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    Apply-StandardGridTheme -Grid $gridFSMO -EnableAlternatingRowColor

    $tabFSMO.Controls.Add($gridFSMO)
    $tabFSMO.Controls.Add($pnlTopFSMO)
    $pnlTopFSMO.SendToBack()
    $gridFSMO.BringToFront()

    # --- TAB 3: PRIVILEGIERTE ADMINS & DELEGIERUNG ---
    $tabAdmins = New-Object System.Windows.Forms.TabPage
    $tabAdmins.Text = if ($script:CurrentLang -eq "DE") { "  🛡️ Privilegierte Konten & Delegierung  " } else { "  🛡️ Privileged Accounts & Delegation  " }
    $tabAdmins.BackColor = [System.Drawing.Color]::White

    $pnlAdminTop = New-Object System.Windows.Forms.Panel
    $pnlAdminTop.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlAdminTop.Height = 48
    $pnlAdminTop.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

    $lblAdminFilter = New-Object System.Windows.Forms.Label
    $lblAdminFilter.Text = if ($script:CurrentLang -eq "DE") { "Live-Filter:" } else { "Live Filter:" }
    $lblAdminFilter.Location = New-Object System.Drawing.Point(12, 14)
    $lblAdminFilter.AutoSize = $true
    $lblAdminFilter.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9.0, [System.Drawing.FontStyle]::Bold)

    $txtAdminFilter = New-Object System.Windows.Forms.TextBox
    $txtAdminFilter.Location = New-Object System.Drawing.Point(95, 11)
    $txtAdminFilter.Size = New-Object System.Drawing.Size(260, 24)

    $lblAdminCount = New-Object System.Windows.Forms.Label
    $lblAdminCount.Location = New-Object System.Drawing.Point(375, 14)
    $lblAdminCount.AutoSize = $true
    $lblAdminCount.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9, [System.Drawing.FontStyle]::Regular)
    $lblAdminCount.ForeColor = [System.Drawing.Color]::FromArgb(50, 70, 90)

    $pnlAdminTop.Controls.AddRange(@($lblAdminFilter, $txtAdminFilter, $lblAdminCount))

    $pnlAdminLegend = New-Object System.Windows.Forms.Panel
    $pnlAdminLegend.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $pnlAdminLegend.Height = 32
    $pnlAdminLegend.BackColor = [System.Drawing.Color]::FromArgb(246, 248, 250)

    $lblAdminLegend = New-Object System.Windows.Forms.Label
    $lblAdminLegend.Dock = [System.Windows.Forms.DockStyle]::Fill
    $lblAdminLegend.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $lblAdminLegend.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 8.5)
    $lblAdminLegend.Text = " 🟥 Rot: Schema-Admins (Sollte im Normalbetrieb leer sein!)  |  ⚠️ Gelb: Delegierung erlaubt (Konto nicht vor Kerberos-Delegierung geschützt)"
    $pnlAdminLegend.Controls.Add($lblAdminLegend)

    $gridAdmins = New-Object System.Windows.Forms.DataGridView
    $gridAdmins.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridAdmins.ReadOnly = $true
    $gridAdmins.AllowUserToAddRows = $false
    $gridAdmins.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    Apply-StandardGridTheme -Grid $gridAdmins -EnableAlternatingRowColor

    $tabAdmins.Controls.Add($gridAdmins)
    $tabAdmins.Controls.Add($pnlAdminTop)
    $tabAdmins.Controls.Add($pnlAdminLegend)
    $pnlAdminTop.SendToBack()
    $pnlAdminLegend.SendToBack()
    $gridAdmins.BringToFront()

    $tabControl.TabPages.AddRange(@($tabDCs, $tabFSMO, $tabAdmins))
    $subForm.Controls.Add($tabControl)

    $script:adminList = [System.Collections.Generic.List[PSCustomObject]]::new()
    $script:dcList    = [System.Collections.Generic.List[PSCustomObject]]::new()
    $script:fsmoList  = [System.Collections.Generic.List[PSCustomObject]]::new()
    $script:SchemaAdminSid = ""

    $gridAdmins.Add_RowPrePaint({
        param($sender, $e)
        if ($e.RowIndex -ge 0 -and $e.RowIndex -lt $gridAdmins.Rows.Count) {
            $row  = $gridAdmins.Rows[$e.RowIndex]
            $grp  = [string]$row.Cells["Admin-Gruppe"].Value
            $sid  = [string]$row.Cells["Gruppen-SID"].Value
            $del  = [string]$row.Cells["Delegierungsschutz"].Value
            $stat = [string]$row.Cells["Status"].Value

            if (($script:SchemaAdminSid -and $sid -eq $script:SchemaAdminSid) -or ($grp -match "Schema[- ]?Admins")) {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 224, 224)
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(180, 0, 0)
                $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(245, 190, 190)
                $row.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
            }
            elseif ($del -like "*Ungeschützt*") {
                $row.Cells["Delegierungsschutz"].Style.BackColor = [System.Drawing.Color]::FromArgb(255, 243, 205)
                $row.Cells["Delegierungsschutz"].Style.ForeColor = [System.Drawing.Color]::FromArgb(160, 80, 0)
                $row.Cells["Delegierungsschutz"].Style.SelectionBackColor = [System.Drawing.Color]::FromArgb(245, 225, 170)
                $row.Cells["Delegierungsschutz"].Style.SelectionForeColor = [System.Drawing.Color]::Black
            }

            if ($stat -eq "Deaktiviert") {
                $row.Cells["Status"].Style.ForeColor = [System.Drawing.Color]::Gray
            }
        }
    })

    $loadDomainData = {
        try {
            $domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
            $forest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()

            # 1. DC Liste
            $script:dcList.Clear()
            foreach ($dc in $domain.DomainControllers) {
                $os = "Windows Server"
                try {
                    $de = $dc.GetDirectoryEntry()
                    if ($de.operatingSystem) { $os = "$($de.operatingSystem) $($de.operatingSystemServicePack)" }
                } catch {}

                $roles = @()
                if ($dc.Name -eq $domain.PdcRoleOwner.Name) { $roles += "PDC" }
                if ($dc.Name -eq $domain.RidRoleOwner.Name) { $roles += "RID" }
                if ($dc.Name -eq $domain.InfrastructureRoleOwner.Name) { $roles += "Infra" }
                if ($dc.Name -eq $forest.SchemaRoleOwner.Name) { $roles += "Schema" }
                if ($dc.Name -eq $forest.NamingRoleOwner.Name) { $roles += "Naming" }
                if ($roles.Count -eq 0) { $roles += "Standard DC" }

                $script:dcList.Add([PSCustomObject]@{
                    "DC Hostname"     = $dc.Name
                    "AD Site"         = $dc.SiteName
                    "IP-Adresse"      = $dc.IPAddress
                    "Global Catalog"  = if ($dc.IsGlobalCatalog()) { "Ja / Active" } else { "Nein" }
                    "Inhaber Rollen"  = ($roles -join ", ")
                    "Betriebssystem"  = $os
                })
            }
            $gridDCs.DataSource = [System.Collections.ArrayList]::new($script:dcList)
            $lblDCCount.Text = if ($script:CurrentLang -eq "DE") { "Gefundene Domain Controller ($($script:dcList.Count)):" } else { "Discovered Domain Controllers ($($script:dcList.Count)):" }

            # 2. FSMO & Domänenstatus
            $recycleBin = "Deaktiviert / Nicht konfiguriert"
            try {
                $rootDSE = [ADSI]"LDAP://RootDSE"
                $partDN = "CN=Partitions," + $rootDSE.configurationNamingContext.ToString()
                $partEntry = [ADSI]"LDAP://$partDN"
                if ($partEntry.Properties["msDS-EnabledFeature"].Count -gt 0) {
                    $recycleBin = "Aktiviert (Active Directory Recycle Bin ON)"
                }
            } catch {}

            $script:fsmoList.Clear()
            $script:fsmoList.Add([PSCustomObject]@{"Parameter / FSMO Rolle" = "PDC Emulator"; "Inhaber / Status" = $domain.PdcRoleOwner.Name; "Geltungsbereich" = "Domäne" })
            $script:fsmoList.Add([PSCustomObject]@{"Parameter / FSMO Rolle" = "RID Master"; "Inhaber / Status" = $domain.RidRoleOwner.Name; "Geltungsbereich" = "Domäne" })
            $script:fsmoList.Add([PSCustomObject]@{"Parameter / FSMO Rolle" = "Infrastructure Master"; "Inhaber / Status" = $domain.InfrastructureRoleOwner.Name; "Geltungsbereich" = "Domäne" })
            $script:fsmoList.Add([PSCustomObject]@{"Parameter / FSMO Rolle" = "Schema Master"; "Inhaber / Status" = $forest.SchemaRoleOwner.Name; "Geltungsbereich" = "Gesamtstruktur" })
            $script:fsmoList.Add([PSCustomObject]@{"Parameter / FSMO Rolle" = "Domain Naming Master"; "Inhaber / Status" = $forest.NamingRoleOwner.Name; "Geltungsbereich" = "Gesamtstruktur" })
            $script:fsmoList.Add([PSCustomObject]@{"Parameter / FSMO Rolle" = "Domänen-Funktionsebene"; "Inhaber / Status" = $domain.DomainMode.ToString(); "Geltungsbereich" = "Domäne" })
            $script:fsmoList.Add([PSCustomObject]@{"Parameter / FSMO Rolle" = "Gesamtstruktur-Funktionsebene"; "Inhaber / Status" = $forest.ForestMode.ToString(); "Geltungsbereich" = "Gesamtstruktur" })
            $script:fsmoList.Add([PSCustomObject]@{"Parameter / FSMO Rolle" = "AD Papierkorb (Recycle Bin)"; "Inhaber / Status" = $recycleBin; "Geltungsbereich" = "Gesamtstruktur" })

            $gridFSMO.DataSource = [System.Collections.ArrayList]::new($script:fsmoList)

            # 3. Privilegierte Admin-Gruppen & Delegierung
            $script:adminList.Clear()
            $domDN = Get-DomainDN
            $domainEntry = [ADSI]"LDAP://$domDN"
            $domainSID = (New-Object System.Security.Principal.SecurityIdentifier($domainEntry.Properties["objectSid"][0], 0)).Value
            $script:SchemaAdminSid = "$domainSID-518"

            $groupsToQuery = @(
                @{ Name = "Domänen-Admins";       SID = "$domainSID-512" },
                @{ Name = "Organisations-Admins"; SID = "$domainSID-519" },
                @{ Name = "Schema-Admins";        SID = $script:SchemaAdminSid },
                @{ Name = "Administratoren (AD)"; SID = "S-1-5-32-544" },
                @{ Name = "Konten-Operatoren";    SID = "$domainSID-517" },
                @{ Name = "Server-Operatoren";    SID = "$domainSID-549" },
                @{ Name = "Sicherungs-Operatoren";SID = "$domainSID-551" },
                @{ Name = "GPO-Ersteller";        SID = "$domainSID-520" }
            )

            foreach ($g in $groupsToQuery) {
                $gSearcher = New-Object DirectoryServices.DirectorySearcher($domainEntry)
                $gSearcher.Filter = "(&(objectCategory=group)(objectsid=$($g.SID)))"
                $groupObj = $gSearcher.FindOne()
                if ($groupObj) {
                    $groupName = if ($groupObj.Properties["samaccountname"].Count -gt 0) { $groupObj.Properties["samaccountname"][0] } else { $g.Name }
                    
                    foreach ($mDN in $groupObj.Properties["member"]) {
                        try {
                            $mSearcher = New-Object DirectoryServices.DirectorySearcher($domainEntry)
                            $mSearcher.Filter = "(distinguishedName=$mDN)"
                            $mObj = $mSearcher.FindOne()
                            
                            if ($mObj) {
                                $sAM = [string]$mObj.Properties["samaccountname"][0]
                                $uac = if ($mObj.Properties["useraccountcontrol"].Count -gt 0) { [int]$mObj.Properties["useraccountcontrol"][0] } else { 0 }
                                $enabled = if (($uac -band 2) -eq 2) { "Deaktiviert" } else { "Aktiv" }

                                $isNotDelegated = ($uac -band 0x100000) -eq 0x100000
                                $delegationStatus = if ($isNotDelegated) { 
                                    "Geschützt (Keine Delegierung)" 
                                } else { 
                                    "⚠️ Ungeschützt (Delegierung erlaubt)" 
                                }

                                $script:adminList.Add([PSCustomObject]@{
                                    "Admin-Gruppe"        = $groupName
                                    "Gruppen-SID"         = $g.SID
                                    "Account Name"        = $sAM
                                    "Status"              = $enabled
                                    "Delegierungsschutz"  = $delegationStatus
                                    "DistinguishedName"   = $mDN
                                })
                            }
                        } catch {}
                    }
                }
            }

            $gridAdmins.DataSource = [System.Collections.ArrayList]::new($script:adminList)
            $lblAdminCount.Text = "Gefundene privilegierte Konten: $($script:adminList.Count)"
            if ($gridAdmins.Columns["Gruppen-SID"]) { $gridAdmins.Columns["Gruppen-SID"].Visible = $false }
            if ($gridAdmins.Columns["DistinguishedName"]) { $gridAdmins.Columns["DistinguishedName"].FillWeight = 140 }

        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim LDAP-Abruf: $($_.Exception.Message)", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }

    $btnRefreshDC.Add_Click({ & $loadDomainData })

    $txtAdminFilter.Add_TextChanged({
        $filterText = $txtAdminFilter.Text.Trim().ToLower()
        if ([string]::IsNullOrWhiteSpace($filterText)) {
            $gridAdmins.DataSource = [System.Collections.ArrayList]::new($script:adminList)
            $lblAdminCount.Text = "Gefundene privilegierte Konten: $($script:adminList.Count)"
        } else {
            $filtered = $script:adminList | Where-Object {
                ($_."Admin-Gruppe" -and $_."Admin-Gruppe".ToLower().Contains($filterText)) -or
                ($_."Account Name" -and $_."Account Name".ToLower().Contains($filterText)) -or
                ($_."Delegierungsschutz" -and $_."Delegierungsschutz".ToLower().Contains($filterText)) -or
                ($_."DistinguishedName" -and $_."DistinguishedName".ToLower().Contains($filterText))
            }
            $arrF = [System.Collections.ArrayList]::new()
            foreach ($item in $filtered) { [void]$arrF.Add($item) }
            $gridAdmins.DataSource = $arrF
            $lblAdminCount.Text = "Gefiltert: $($arrF.Count) von $($script:adminList.Count)"
        }
    })

    $subForm.Add_Shown({ & $loadDomainData })
    [void]$subForm.ShowDialog()
}

# ==============================================================================
# TOOL 7: AD SECURITY & OS SUPPORT AUDIT
# ==============================================================================
function Open-ToolOSSupportAudit {
    if (-not (Assert-DomainJoined)) { return }

    $subForm = New-Object System.Windows.Forms.Form
    $subForm.Text = "Tool 7: AD Security & OS Support Audit"
    $subForm.Size = New-Object System.Drawing.Size(1420, 860)
    $subForm.StartPosition = "CenterParent"
    $subForm.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9)

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
    $chkGroupByOS.Location = "100, 78"; $chkGroupByOS.Size = "250, 22"; $chkGroupByOS.Text = "Nach Operating System sortieren"
    $chkGroupByOS.Checked = $true
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
    Apply-StandardGridTheme -Grid $gridOSAudit
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
# TOOL 8: CLIENT SOFTWARE & APP ANALYSE
# ==============================================================================
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

# ==============================================================================
# TOOL 9: AD UNIVERSAL ACL & BERECHTIGUNGSVERGLEICH
# ==============================================================================
function Show-Tool9-ACLCompare {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Tool 9: Active Directory Universal ACL & Berechtigungsvergleich"
    $form.Size = New-Object System.Drawing.Size(1280, 800)
    $form.StartPosition = "CenterScreen"
    $form.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9)
    $form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = "AD Objekt ACL-Vergleich (User, Gruppen, Computer, OUs & Berechtigungs-Diff)"
    $lblTitle.Font = New-Object System.Drawing.Font($form.Font.FontFamily, 12, [System.Drawing.FontStyle]::Bold)
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
    $btnCompare.Font = New-Object System.Drawing.Font($form.Font.FontFamily, 10, [System.Drawing.FontStyle]::Bold)
    $grpInputs.Controls.Add($btnCompare)

    $btnExport = New-Object System.Windows.Forms.Button
    $btnExport.Text = "CSV Exportieren"; $btnExport.Location = "935, 22"; $btnExport.Size = "140, 80"
    $btnExport.BackColor = [System.Drawing.Color]::FromArgb(40, 167, 69); $btnExport.ForeColor = [System.Drawing.Color]::White
    $btnExport.Font = New-Object System.Drawing.Font($form.Font.FontFamily, 9.5, [System.Drawing.FontStyle]::Bold)
    $grpInputs.Controls.Add($btnExport)

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Text = "Geben Sie Namen, sAMAccountNames oder DNs zweier AD-Objekte ein."; $lblStatus.Location = "15, 165"; $lblStatus.Size = "1230, 24"
    $lblStatus.Font = New-Object System.Drawing.Font($form.Font.FontFamily, 9, [System.Drawing.FontStyle]::Italic)
    $form.Controls.Add($lblStatus)

    $gridACL = New-Object System.Windows.Forms.DataGridView
    $gridACL.Location = "15, 195"; $gridACL.Size = "1230, 550"
    $gridACL.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $gridACL.ReadOnly = $true; $gridACL.AllowUserToAddRows = $false
    $gridACL.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    Apply-StandardGridTheme -Grid $gridACL
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
                "User / Benutzer" { "(&(objectCategory=person)(objectClass=user))" }
                "Group / Gruppe"  { "(objectCategory=group)" }
                "OU / Container"  { "(|(objectCategory=organizationalUnit)(objectCategory=container))" }
                default           { "(objectClass=*)" }
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

        if ($gridACL.Columns["Vergleichs-Status"])  { $gridACL.Columns["Vergleichs-Status"].FillWeight  = 110 }
        if ($gridACL.Columns["Principal / Gruppe"]) { $gridACL.Columns["Principal / Gruppe"].FillWeight = 130 }
        if ($gridACL.Columns["Rechte (Objekt 1)"])  { $gridACL.Columns["Rechte (Objekt 1)"].FillWeight  = 160 }
        if ($gridACL.Columns["Rechte (Objekt 2)"])  { $gridACL.Columns["Rechte (Objekt 2)"].FillWeight  = 160 }
        if ($gridACL.Columns["Typ"])                { $gridACL.Columns["Typ"].FillWeight                = 70 }
        if ($gridACL.Columns["Objekt 1 Vererbt"])   { $gridACL.Columns["Objekt 1 Vererbt"].FillWeight   = 85 }
        if ($gridACL.Columns["Objekt 2 Vererbt"])   { $gridACL.Columns["Objekt 2 Vererbt"].FillWeight   = 85 }

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
            $comparisonList = [System.Collections.Generic.List[PSObject]]::new()
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
                                "Rechte (Objekt 2)"  = ($r2List.Rights -join ", ")
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
                                "Rechte (Objekt 1)"  = ($r1List.Rights -join ", ")
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
            [System.Windows.Forms.MessageBox]::Show("ACL-Vergleich erfolgreich exportiert nach:`n$($sfd.FileName)", "Export Abgeschlossen", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    })

    [void]$form.ShowDialog()
}

# ==============================================================================
# TOOL 10: ACTIVE DIRECTORY OU & GRUPPEN FINDER (INKL. ADMINCOUNT & UPN)
# ==============================================================================
function Show-Tool10-OUGroupFinder {
    if (-not (Assert-DomainJoined)) { return }

    $Form = New-Object System.Windows.Forms.Form
    $Form.Text = "Tool 10: Active Directory OU & Gruppen Finder (inkl. adminCount & UPN)"
    $Form.Size = New-Object System.Drawing.Size(1300, 750)
    $Form.StartPosition = "CenterScreen"
    $Form.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9)
    $Form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

    $script:CurrentObjectGroups = @()

    # --- OBERES STEUERUNGS-PANEL (LAYOUT & ABSTÄNDE) ---
    $pnlTop = New-Object System.Windows.Forms.Panel
    $pnlTop.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlTop.Height = 75
    $pnlTop.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
    $Form.Controls.Add($pnlTop)

    # 1. Objekttyp
    $LblType = New-Object System.Windows.Forms.Label
    $LblType.Text = "Objekttyp:"
    $LblType.Location = New-Object System.Drawing.Point(15, 15)
    $LblType.AutoSize = $true
    $LblType.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9, [System.Drawing.FontStyle]::Bold)
    $pnlTop.Controls.Add($LblType)

    $CmbType = New-Object System.Windows.Forms.ComboBox
    $CmbType.Location = New-Object System.Drawing.Point(95, 12)
    $CmbType.Size = New-Object System.Drawing.Size(175, 24)
    $CmbType.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$CmbType.Items.Add("Clients (ohne Server)")
    [void]$CmbType.Items.Add("Server (nur Server OS)")
    [void]$CmbType.Items.Add("User (Benutzerkonten)")
    $CmbType.SelectedIndex = 0
    $pnlTop.Controls.Add($CmbType)

    # 2. Suchbegriff
    $LblSearch = New-Object System.Windows.Forms.Label
    $LblSearch.Text = "Suchbegriff:"
    $LblSearch.Location = New-Object System.Drawing.Point(290, 15)
    $LblSearch.AutoSize = $true
    $LblSearch.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9, [System.Drawing.FontStyle]::Bold)
    $pnlTop.Controls.Add($LblSearch)

    $TxtSearch = New-Object System.Windows.Forms.TextBox
    $TxtSearch.Location = New-Object System.Drawing.Point(375, 12)
    $TxtSearch.Size = New-Object System.Drawing.Size(300, 24)
    $TxtSearch.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9)
    $TxtSearch.Text = "*"
    $pnlTop.Controls.Add($TxtSearch)

    # 3. Suchen-Button
    $BtnSearch = New-Object System.Windows.Forms.Button
    $BtnSearch.Text = "Suchen"
    $BtnSearch.Location = New-Object System.Drawing.Point(690, 10)
    $BtnSearch.Size = New-Object System.Drawing.Size(110, 28)
    $BtnSearch.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $BtnSearch.ForeColor = [System.Drawing.Color]::White
    $BtnSearch.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $BtnSearch.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9, [System.Drawing.FontStyle]::Bold)
    $pnlTop.Controls.Add($BtnSearch)

    # Statuszeile
    $LblStatus = New-Object System.Windows.Forms.Label
    $LblStatus.Text = "Bereit."
    $LblStatus.Location = New-Object System.Drawing.Point(15, 46)
    $LblStatus.AutoSize = $true
    $LblStatus.ForeColor = [System.Drawing.Color]::FromArgb(50, 70, 90)
    $pnlTop.Controls.Add($LblStatus)

    # --- SPLITCONTAINER (LINKS: DATAGRID, RECHTS: GRUPPEN) ---
    $splitContainer = New-Object System.Windows.Forms.SplitContainer
    $splitContainer.Dock = [System.Windows.Forms.DockStyle]::Fill
    $splitContainer.SplitterDistance = 840
    $splitContainer.Panel1.Padding = New-Object System.Windows.Forms.Padding(15, 5, 5, 15)
    $splitContainer.Panel2.Padding = New-Object System.Windows.Forms.Padding(5, 5, 15, 15)
    $Form.Controls.Add($splitContainer)
    $splitContainer.BringToFront()
    $pnlTop.SendToBack()

    # Tabelle links
    $DataGrid = New-Object System.Windows.Forms.DataGridView
    $DataGrid.Dock = [System.Windows.Forms.DockStyle]::Fill
    $DataGrid.ReadOnly = $true
    $DataGrid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $DataGrid.MultiSelect = $false
    $DataGrid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    Apply-StandardGridTheme -Grid $DataGrid
    $splitContainer.Panel1.Controls.Add($DataGrid)

    # Gruppen-Container rechts
    $pnlGroupHeader = New-Object System.Windows.Forms.Panel
    $pnlGroupHeader.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlGroupHeader.Height = 65
    $pnlGroupHeader.BackColor = [System.Drawing.Color]::Transparent
    $splitContainer.Panel2.Controls.Add($pnlGroupHeader)

    $LblGroups = New-Object System.Windows.Forms.Label
    $LblGroups.Text = "Gruppenmitgliedschaften:"
    $LblGroups.Location = New-Object System.Drawing.Point(0, 5)
    $LblGroups.AutoSize = $true
    $LblGroups.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9, [System.Drawing.FontStyle]::Bold)
    $pnlGroupHeader.Controls.Add($LblGroups)

    $TxtGroupFilter = New-Object System.Windows.Forms.TextBox
    $TxtGroupFilter.Location = New-Object System.Drawing.Point(0, 30)
    $TxtGroupFilter.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $pnlGroupHeader.Controls.Add($TxtGroupFilter)

    $PlaceholderText = "Gruppen filtern... (z.B. *Admin*)"
    $TxtGroupFilter.Text = $PlaceholderText
    $TxtGroupFilter.ForeColor = [System.Drawing.Color]::Gray

    $TxtGroupFilter.Add_GotFocus({
        if ($TxtGroupFilter.Text -eq $PlaceholderText) {
            $TxtGroupFilter.Text = ""
            $TxtGroupFilter.ForeColor = [System.Drawing.Color]::Black
        }
    })

    $TxtGroupFilter.Add_LostFocus({
        if ([string]::IsNullOrWhiteSpace($TxtGroupFilter.Text)) {
            $TxtGroupFilter.Text = $PlaceholderText
            $TxtGroupFilter.ForeColor = [System.Drawing.Color]::Gray
        }
    })

    $LstGroups = New-Object System.Windows.Forms.ListBox
    $LstGroups.Dock = [System.Windows.Forms.DockStyle]::Fill
    $LstGroups.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $LstGroups.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9)
    $splitContainer.Panel2.Controls.Add($LstGroups)
    $LstGroups.BringToFront()

    # --- FARB-HIGHLIGHTING FÜR ADMINCOUNT = 1 ---
    $DataGrid.Add_RowPrePaint({
        param($sender, $e)
        if ($e.RowIndex -ge 0 -and $e.RowIndex -lt $DataGrid.Rows.Count) {
            $row = $DataGrid.Rows[$e.RowIndex]
            $adminCountVal = $row.Cells["adminCount"].Value
            if ($adminCountVal -eq 1 -or $adminCountVal -eq "1") {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 243, 205)
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(140, 70, 0)
                $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(245, 225, 170)
                $row.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
            } else {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::White
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Black
            }
        }
    })

    # --- FILTER-ROUTINE FÜR DIE RECHTE GRUPPENLISTE ---
    $FilterGroupsList = {
        $LstGroups.Items.Clear()
        if (-not $script:CurrentObjectGroups -or $script:CurrentObjectGroups.Count -eq 0) { return }

        $Filter = $TxtGroupFilter.Text.Trim()

        if ($Filter -eq $PlaceholderText -or [string]::IsNullOrWhiteSpace($Filter)) {
            $Filtered = $script:CurrentObjectGroups
        } else {
            $Pattern = if ($Filter -notlike "*\**") { "*$Filter*" } else { $Filter }
            $Filtered = $script:CurrentObjectGroups | Where-Object { $_ -like $Pattern }
        }

        $LblGroups.Text = "Gruppen ($($Filtered.Count) von $($script:CurrentObjectGroups.Count)):"
        foreach ($Group in $Filtered) {
            [void]$LstGroups.Items.Add($Group)
        }
    }

    # --- HAUPTSUCHE (NATIVE LDAP / ADSI) ---
    $PerformSearch = {
        $SearchTerm = $TxtSearch.Text.Trim()
        $LstGroups.Items.Clear()
        $script:CurrentObjectGroups = @()
        $LblGroups.Text = "Gruppenmitgliedschaften:"
        
        if ([string]::IsNullOrWhiteSpace($SearchTerm)) {
            [System.Windows.Forms.MessageBox]::Show("Bitte geben Sie einen Suchbegriff ein.", "Hinweis", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        $FilterPattern = if ($SearchTerm -notlike "*\**") { "*$SearchTerm*" } else { $SearchTerm }

        $LblStatus.Text = "Suche läuft..."
        $LblStatus.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
        $Form.Refresh()

        $ResultsList = [System.Collections.Generic.List[PSObject]]::new()

        try {
            switch ($CmbType.SelectedIndex) {
                0 { # Clients (ohne Server)
                    $raw = Search-NativeLdap -LdapFilter "(&(objectCategory=computer)(name=$FilterPattern)(!(operatingSystem=*Server*)))" -PropertiesToLoad @("name","distinguishedName","userAccountControl","operatingSystem","adminCount","userPrincipalName")
                    foreach ($Item in $raw) {
                        $p = $Item.Properties
                        $dn = [string]$p["distinguishedname"][0]
                        $dnComps = $dn -split '(?<!\\),'
                        $ou = ($dnComps[1..($dnComps.Length - 1)]) -join ','
                        $uac = if ($p["useraccountcontrol"].Count -gt 0) { [int]$p["useraccountcontrol"][0] } else { 0 }
                        $enabled = ($uac -band 2) -ne 2
                        $aCount = if ($p["admincount"].Count -gt 0 -and $p["admincount"][0] -eq 1) { 1 } else { 0 }
                        $upn = if ($p["userprincipalname"].Count -gt 0) { $p["userprincipalname"][0] } else { "" }
                        $os = if ($p["operatingsystem"].Count -gt 0) { $p["operatingsystem"][0] } else { "Unbekannt" }

                        $ResultsList.Add([PSCustomObject]@{
                            "Name"              = $p["name"][0]
                            "OU Pfad"           = $ou
                            "UPN"               = $upn
                            "adminCount"        = $aCount
                            "Aktiv"             = $enabled
                            "Betriebssystem"    = $os
                            "DistinguishedName" = $dn
                        })
                    }
                }
                1 { # Server
                    $raw = Search-NativeLdap -LdapFilter "(&(objectCategory=computer)(name=$FilterPattern)(operatingSystem=*Server*))" -PropertiesToLoad @("name","distinguishedName","userAccountControl","operatingSystem","adminCount","userPrincipalName")
                    foreach ($Item in $raw) {
                        $p = $Item.Properties
                        $dn = [string]$p["distinguishedname"][0]
                        $dnComps = $dn -split '(?<!\\),'
                        $ou = ($dnComps[1..($dnComps.Length - 1)]) -join ','
                        $uac = if ($p["useraccountcontrol"].Count -gt 0) { [int]$p["useraccountcontrol"][0] } else { 0 }
                        $enabled = ($uac -band 2) -ne 2
                        $aCount = if ($p["admincount"].Count -gt 0 -and $p["admincount"][0] -eq 1) { 1 } else { 0 }
                        $upn = if ($p["userprincipalname"].Count -gt 0) { $p["userprincipalname"][0] } else { "" }
                        $os = if ($p["operatingsystem"].Count -gt 0) { $p["operatingsystem"][0] } else { "Unbekannt" }

                        $ResultsList.Add([PSCustomObject]@{
                            "Name"              = $p["name"][0]
                            "OU Pfad"           = $ou
                            "UPN"               = $upn
                            "adminCount"        = $aCount
                            "Aktiv"             = $enabled
                            "Betriebssystem"    = $os
                            "DistinguishedName" = $dn
                        })
                    }
                }
                2 { # User
                    $raw = Search-NativeLdap -LdapFilter "(&(objectCategory=person)(objectClass=user)(|(sAMAccountName=$FilterPattern)(displayName=$FilterPattern)(userPrincipalName=$FilterPattern)))" -PropertiesToLoad @("sAMAccountName","displayName","distinguishedName","userAccountControl","adminCount","userPrincipalName","mail","department")
                    foreach ($Item in $raw) {
                        $p = $Item.Properties
                        $dn = [string]$p["distinguishedname"][0]
                        $dnComps = $dn -split '(?<!\\),'
                        $ou = ($dnComps[1..($dnComps.Length - 1)]) -join ','
                        $uac = if ($p["useraccountcontrol"].Count -gt 0) { [int]$p["useraccountcontrol"][0] } else { 0 }
                        $enabled = ($uac -band 2) -ne 2
                        $aCount = if ($p["admincount"].Count -gt 0 -and $p["admincount"][0] -eq 1) { 1 } else { 0 }
                        $upn = if ($p["userprincipalname"].Count -gt 0) { $p["userprincipalname"][0] } else { "" }
                        $sAM = if ($p["samaccountname"].Count -gt 0) { $p["samaccountname"][0] } else { "" }
                        $dName = if ($p["displayname"].Count -gt 0) { $p["displayname"][0] } else { "" }
                        $mail = if ($p["mail"].Count -gt 0) { $p["mail"][0] } else { "" }
                        $dept = if ($p["department"].Count -gt 0) { $p["department"][0] } else { "" }

                        $ResultsList.Add([PSCustomObject]@{
                            "Anmeldename"        = $sAM
                            "Vollständiger Name" = $dName
                            "UPN"                = $upn
                            "OU Pfad"            = $ou
                            "adminCount"         = $aCount
                            "E-Mail"             = $mail
                            "Abteilung"          = $dept
                            "Aktiv"              = $enabled
                            "DistinguishedName"  = $dn
                        })
                    }
                }
            }

            $DataGrid.DataSource = [System.Collections.ArrayList]::new($ResultsList)
            $LblStatus.Text = "Gefundene Objekte: $($ResultsList.Count)"
            $LblStatus.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 215)

            if ($DataGrid.Columns["DistinguishedName"]) { $DataGrid.Columns["DistinguishedName"].Visible = $false }
            if ($DataGrid.Columns["adminCount"])        { $DataGrid.Columns["adminCount"].FillWeight = 40 }
            if ($DataGrid.Columns["OU Pfad"])           { $DataGrid.Columns["OU Pfad"].FillWeight = 160 }

        } catch {
            $LblStatus.Text = "Fehler bei der Abfrage."
            $LblStatus.ForeColor = [System.Drawing.Color]::Red
            [System.Windows.Forms.MessageBox]::Show("Fehler bei AD-Abfrage: $($_.Exception.Message)", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }

    # --- SELECTION CHANGED: GRUPPEN LADEN ---
    $DataGrid.Add_SelectionChanged({
        $LstGroups.Items.Clear()
        $script:CurrentObjectGroups = @()

        if ($DataGrid.SelectedRows.Count -gt 0) {
            $SelectedRow = $DataGrid.SelectedRows[0]
            $DN = [string]$SelectedRow.Cells["DistinguishedName"].Value

            if ($DN) {
                $LblGroups.Text = "Gruppen werden geladen..."
                $Form.Refresh()

                try {
                    $grpList = @()
                    $objEntry = [ADSI]"LDAP://$DN"
                    if ($objEntry.Properties["memberOf"].Count -gt 0) {
                        foreach ($mDN in $objEntry.Properties["memberOf"]) {
                            if ($mDN -match "^CN=([^,]+)") { $grpList += $Matches[1] } else { $grpList += $mDN }
                        }
                    }
                    $script:CurrentObjectGroups = ($grpList | Sort-Object)
                    & $FilterGroupsList
                }
                catch {
                    $LblGroups.Text = "Fehler beim Laden der Gruppen."
                    [void]$LstGroups.Items.Add("Fehler: $($_.Exception.Message)")
                }
            }
        }
    })

    # Event-Bindungen
    $TxtGroupFilter.Add_KeyUp({ & $FilterGroupsList })
    $BtnSearch.Add_Click($PerformSearch)
    $TxtSearch.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $_.SuppressKeyPress = $true
            & $PerformSearch
        }
    })

    $Form.Add_Shown({
        $BtnSearch.PerformClick()
    })

    [void]$Form.ShowDialog()
}

# ==============================================================================
# TOOL 11: AD PASSWORD POLICIES & FINE-GRAINED PSO AUDIT (ORIGINAL FUNKTIONIERENDE VERSION)
# ==============================================================================
function Show-Tool11-PasswordPolicyAudit {
    if (-not (Assert-DomainJoined)) { return }

    $lang = if ($script:CurrentLang -eq "EN") { "EN" } else { "DE" }
    $t = @{
        "DE" = @{
            "Title"           = "Tool 11: AD Kennwortrichtlinien & PSO Audit"
            "TabDomain"       = "1. Standard-Domänenrichtlinie"
            "TabPSO"          = "2. Fine-Grained Password Policies (PSO)"
            "TabUser"         = "3. Effektive Benutzer-Richtlinie"
            "PropName"        = "Eigenschaft / Richtlinie"
            "PropValue"       = "Konfigurierter Wert"
            "PsoName"         = "PSO Name"
            "PsoPrecedence"   = "Priorität (Precedence)"
            "PsoMinLength"    = "Min. Länge"
            "PsoHistory"      = "Historie"
            "PsoComplexity"   = "Komplexität"
            "PsoMaxAge"       = "Max. Alter (Tage)"
            "PsoMinAge"       = "Min. Alter (Tage)"
            "PsoLockoutThresh"= "Sperrschwelle"
            "PsoLockoutDur"   = "Sperrdauer (Min.)"
            "PsoAppliesTo"    = "Zugewiesen an (Applies To)"
            "BtnLoad"         = "Richtlinien laden / aktualisieren"
            "BtnExport"       = "Exportieren (CSV)"
            "LblUserSearch"   = "Benutzername (sAMAccountName):"
            "BtnCheckUser"    = "Effektive Richtlinie prüfen"
            "StatusReady"     = "Bereit zur Abfrage."
            "StatusLoading"   = "Lese Active Directory Kennwortrichtlinien aus..."
            "StatusDone"      = "Kennwortrichtlinien erfolgreich geladen."
            "ErrNoDomain"     = "Keine Active Directory Domäne erreichbar."
            "ErrUserNotFound" = "Benutzer nicht im Active Directory gefunden."
            "UserResultTitle" = "Ergebnis für Benutzer"
            "InheritedDomain" = "Standard-Domänenrichtlinie (Keine spezifische PSO)"
            "Days"            = "Tage"
            "Minutes"         = "Minuten"
            "Characters"      = "Zeichen"
            "Passwords"       = "Passwörter"
            "Attempts"        = "Versuche"
            "InvalidAttempts" = "ungültige Versuche"
            "NoLockout"       = "Keine Sperre (0)"
            "NeverExpires"    = "Nie ablaufend / Deaktiviert"
            "NoneAssigned"    = "Niemand (Keine Zuweisung)"
            "Enabled"         = "Aktiviert"
            "Disabled"        = "Deaktiviert"
            "Safe"            = "Sicher"
            "Unsafe"          = "Unsicher"
            "Yes"             = "Ja"
            "No"              = "Nein"
            "DPO_MinLength"   = "Minimale Kennwortlänge"
            "DPO_Complexity"  = "Kennwortkomplexität"
            "DPO_MaxAge"      = "Maximales Kennwortalter"
            "DPO_MinAge"      = "Minimales Kennwortalter"
            "DPO_History"     = "Kennworthistorie / Verlauf"
            "DPO_LockoutTh"   = "Kontosperrungsschwelle"
            "DPO_LockoutDur"  = "Kontosperrdauer"
            "DPO_ObsWindow"   = "Sperrungsbeobachtungsfenster"
            "DPO_RevEncrypt"  = "Umkehrbare Verschlüsselung"
            "Eff_Policy"      = "Gültige Richtlinie"
            "Eff_PSO_DN"      = "PSO DistinguishedName"
            "Eff_LockoutTh"   = "Kontosperrung nach Fehlversuchen"
        }
        "EN" = @{
            "Title"           = "Tool 11: AD Password Policies & PSO Audit"
            "TabDomain"       = "1. Default Domain Policy"
            "TabPSO"          = "2. Fine-Grained Password Policies (PSO)"
            "TabUser"         = "3. Effective User Password Policy"
            "PropName"        = "Property / Policy Setting"
            "PropValue"       = "Configured Value"
            "PsoName"         = "PSO Name"
            "PsoPrecedence"   = "Precedence"
            "PsoMinLength"    = "Min Length"
            "PsoHistory"      = "History"
            "PsoComplexity"   = "Complexity"
            "PsoMaxAge"       = "Max Age (Days)"
            "PsoMinAge"       = "Min Age (Days)"
            "PsoLockoutThresh"= "Lockout Threshold"
            "PsoLockoutDur"   = "Lockout Duration (Min)"
            "PsoAppliesTo"    = "Applies To"
            "BtnLoad"         = "Load / Refresh Policies"
            "BtnExport"       = "Export (CSV)"
            "LblUserSearch"   = "Username (sAMAccountName):"
            "BtnCheckUser"    = "Check Effective Policy"
            "StatusReady"     = "Ready for query."
            "StatusLoading"   = "Querying Active Directory Password Policies..."
            "StatusDone"      = "Password policies loaded successfully."
            "ErrNoDomain"     = "No Active Directory Domain accessible."
            "ErrUserNotFound" = "User not found in Active Directory."
            "UserResultTitle" = "Result for user"
            "InheritedDomain" = "Default Domain Policy (No specific PSO applied)"
            "Days"            = "Days"
            "Minutes"         = "Minutes"
            "Characters"      = "Characters"
            "Passwords"       = "Passwords"
            "Attempts"        = "Attempts"
            "InvalidAttempts" = "invalid attempts"
            "NoLockout"       = "No Lockout (0)"
            "NeverExpires"    = "Never Expires / Disabled"
            "NoneAssigned"    = "None (No assignment)"
            "Enabled"         = "Enabled"
            "Disabled"        = "Disabled"
            "Safe"            = "Secure"
            "Unsafe"          = "Insecure"
            "Yes"             = "Yes"
            "No"              = "No"
            "DPO_MinLength"   = "Minimum Password Length"
            "DPO_Complexity"  = "Password Complexity"
            "DPO_MaxAge"      = "Maximum Password Age"
            "DPO_MinAge"      = "Minimum Password Age"
            "DPO_History"     = "Password History Length"
            "DPO_LockoutTh"   = "Account Lockout Threshold"
            "DPO_LockoutDur"  = "Account Lockout Duration"
            "DPO_ObsWindow"   = "Lockout Observation Window"
            "DPO_RevEncrypt"  = "Store Passwords Using Reversible Encryption"
            "Eff_Policy"      = "Effective Policy"
            "Eff_PSO_DN"      = "PSO DistinguishedName"
            "Eff_LockoutTh"   = "Account Lockout Threshold"
        }
    }[$lang]

    function Convert-LargeIntToTimeSpan([object]$largeIntObj) {
        if (-not $largeIntObj) { return $null }
        try {
            $highPart = $largeIntObj.GetType().InvokeMember("HighPart", [System.Reflection.BindingFlags]::GetProperty, $null, $largeIntObj, $null)
            $lowPart  = $largeIntObj.GetType().InvokeMember("LowPart", [System.Reflection.BindingFlags]::GetProperty, $null, $largeIntObj, $null)
            $raw64 = ([int64]$highPart -shl 32) -bor ([int64]$lowPart -band 0xFFFFFFFF)
            if ($raw64 -lt 0) {
                return [timespan]::FromTicks(-$raw64)
            } else {
                return [timespan]::FromTicks($raw64)
            }
        } catch {
            return $null
        }
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $t["Title"]
    $form.Size = New-Object System.Drawing.Size(1000, 680)
    $form.StartPosition = "CenterScreen"
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

    $pnlHeader = New-Object System.Windows.Forms.Panel
    $pnlHeader.Dock = "Top"
    $pnlHeader.Height = 60
    $pnlHeader.BackColor = [System.Drawing.Color]::White
    $pnlHeader.Padding = New-Object System.Windows.Forms.Padding(15, 10, 15, 10)

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = "🔑 " + $t["Title"]
    $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $lblTitle.ForeColor = [System.Drawing.Color]::FromArgb(24, 43, 73)
    $lblTitle.AutoSize = $true
    $lblTitle.Location = New-Object System.Drawing.Point(15, 15)

    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = "🔄 " + $t["BtnLoad"]
    $btnRefresh.Size = New-Object System.Drawing.Size(200, 32)
    $btnRefresh.Location = New-Object System.Drawing.Point(580, 14)
    $btnRefresh.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
    $btnRefresh.ForeColor = [System.Drawing.Color]::White
    $btnRefresh.FlatStyle = "Flat"

    $btnExport = New-Object System.Windows.Forms.Button
    $btnExport.Text = "💾 " + $t["BtnExport"]
    $btnExport.Size = New-Object System.Drawing.Size(150, 32)
    $btnExport.Location = New-Object System.Drawing.Point(790, 14)
    $btnExport.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
    $btnExport.FlatStyle = "Flat"

    $pnlHeader.Controls.AddRange(@($lblTitle, $btnRefresh, $btnExport))

    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Dock = "Fill"
    $tabControl.Padding = New-Object System.Drawing.Point(12, 6)

    $tabDomain = New-Object System.Windows.Forms.TabPage
    $tabDomain.Text = $t["TabDomain"]
    $tabDomain.BackColor = [System.Drawing.Color]::White

    $tabPSO = New-Object System.Windows.Forms.TabPage
    $tabPSO.Text = $t["TabPSO"]
    $tabPSO.BackColor = [System.Drawing.Color]::White

    $tabUser = New-Object System.Windows.Forms.TabPage
    $tabUser.Text = $t["TabUser"]
    $tabUser.BackColor = [System.Drawing.Color]::White

    $tabControl.TabPages.AddRange(@($tabDomain, $tabPSO, $tabUser))

    # Tab 1
    $gridDomain = New-Object System.Windows.Forms.DataGridView
    $gridDomain.Dock = "Fill"
    $gridDomain.BackgroundColor = [System.Drawing.Color]::White
    $gridDomain.BorderStyle = "None"
    $gridDomain.AutoSizeColumnsMode = "Fill"
    $gridDomain.ReadOnly = $true
    $gridDomain.AllowUserToAddRows = $false
    $gridDomain.RowHeadersVisible = $false
    $gridDomain.SelectionMode = "FullRowSelect"
    Apply-StandardGridTheme -Grid $gridDomain -EnableAlternatingRowColor
    $tabDomain.Controls.Add($gridDomain)

    # Tab 2
    $gridPSO = New-Object System.Windows.Forms.DataGridView
    $gridPSO.Dock = "Fill"
    $gridPSO.BackgroundColor = [System.Drawing.Color]::White
    $gridPSO.BorderStyle = "None"
    $gridPSO.AutoSizeColumnsMode = "AllCells"
    $gridPSO.ReadOnly = $true
    $gridPSO.AllowUserToAddRows = $false
    $gridPSO.RowHeadersVisible = $false
    $gridPSO.SelectionMode = "FullRowSelect"
    Apply-StandardGridTheme -Grid $gridPSO -EnableAlternatingRowColor
    $tabPSO.Controls.Add($gridPSO)

    # Tab 3
    $pnlUserTop = New-Object System.Windows.Forms.Panel
    $pnlUserTop.Dock = "Top"
    $pnlUserTop.Height = 55
    $pnlUserTop.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 250)

    $lblUserPrompt = New-Object System.Windows.Forms.Label
    $lblUserPrompt.Text = $t["LblUserSearch"]
    $lblUserPrompt.Location = New-Object System.Drawing.Point(15, 18)
    $lblUserPrompt.AutoSize = $true

    $txtUserSearch = New-Object System.Windows.Forms.TextBox
    $txtUserSearch.Location = New-Object System.Drawing.Point(220, 15)
    $txtUserSearch.Size = New-Object System.Drawing.Size(220, 25)
    $txtUserSearch.Text = $env:USERNAME

    $btnUserCheck = New-Object System.Windows.Forms.Button
    $btnUserCheck.Text = "🔍 " + $t["BtnCheckUser"]
    $btnUserCheck.Location = New-Object System.Drawing.Point(450, 13)
    $btnUserCheck.Size = New-Object System.Drawing.Size(200, 28)
    $btnUserCheck.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
    $btnUserCheck.ForeColor = [System.Drawing.Color]::White
    $btnUserCheck.FlatStyle = "Flat"

    $pnlUserTop.Controls.AddRange(@($lblUserPrompt, $txtUserSearch, $btnUserCheck))

    $gridUser = New-Object System.Windows.Forms.DataGridView
    $gridUser.Dock = "Fill"
    $gridUser.BackgroundColor = [System.Drawing.Color]::White
    $gridUser.BorderStyle = "None"
    $gridUser.AutoSizeColumnsMode = "Fill"
    $gridUser.ReadOnly = $true
    $gridUser.AllowUserToAddRows = $false
    $gridUser.RowHeadersVisible = $false
    $gridUser.SelectionMode = "FullRowSelect"
    Apply-StandardGridTheme -Grid $gridUser -EnableAlternatingRowColor

    $tabUser.Controls.AddRange(@($gridUser, $pnlUserTop))
    $pnlUserTop.SendToBack()
    $gridUser.BringToFront()

    $statusStrip = New-Object System.Windows.Forms.StatusStrip
    $lblStatus = New-Object System.Windows.Forms.ToolStripStatusLabel
    $lblStatus.Text = $t["StatusReady"]
    [void]$statusStrip.Items.Add($lblStatus)

    $form.Controls.AddRange(@($tabControl, $statusStrip, $pnlHeader))
    $pnlHeader.SendToBack()
    $tabControl.BringToFront()

    $script:cachedDomainPolicy = @()
    $script:cachedPSOs = @()

    $loadPolicies = {
        $lblStatus.Text = $t["StatusLoading"]
        [System.Windows.Forms.Application]::DoEvents()

        try {
            $rootDSE = [ADSI]"LDAP://RootDSE"
            $defaultNC = $rootDSE.defaultNamingContext
            if (-not $defaultNC) {
                [System.Windows.Forms.MessageBox]::Show($t["ErrNoDomain"], "AD Error", "OK", "Error")
                return
            }

            # 1. Default Domain Password Policy
            $domEntry = [ADSI]"LDAP://$defaultNC"
            $domProps = $domEntry.Properties

            $minPwdAgeSpan = Convert-LargeIntToTimeSpan $domProps["minPwdAge"].Value
            $maxPwdAgeSpan = Convert-LargeIntToTimeSpan $domProps["maxPwdAge"].Value
            $lockoutDurSpan = Convert-LargeIntToTimeSpan $domProps["lockoutDuration"].Value
            $lockoutWinSpan = Convert-LargeIntToTimeSpan $domProps["lockOutObservationWindow"].Value

            $minPwdAgeDays = if ($minPwdAgeSpan) { [math]::Round($minPwdAgeSpan.TotalDays, 1) } else { 0 }
            $maxPwdAgeDays = if ($maxPwdAgeSpan) { [math]::Round($maxPwdAgeSpan.TotalDays, 0) } else { 0 }
            $lockoutDurMin = if ($lockoutDurSpan) { [math]::Round($lockoutDurSpan.TotalMinutes, 0) } else { 0 }
            $lockoutWinMin = if ($lockoutWinSpan) { [math]::Round($lockoutWinSpan.TotalMinutes, 0) } else { 0 }

            $minPwdLength   = if ($domProps["minPwdLength"].Value) { $domProps["minPwdLength"].Value } else { 0 }
            $pwdHistoryLen  = if ($domProps["pwdHistoryLength"].Value) { $domProps["pwdHistoryLength"].Value } else { 0 }
            $pwdProperties  = if ($domProps["pwdProperties"].Value) { [int]$domProps["pwdProperties"].Value } else { 0 }
            $complexityOn   = ($pwdProperties -band 1) -eq 1
            $reversibleOn   = ($pwdProperties -band 16) -eq 1
            $lockoutThresh  = if ($domProps["lockoutThreshold"].Value) { $domProps["lockoutThreshold"].Value } else { 0 }

            $domainPolicyList = @(
                [PSCustomObject]@{ $t["PropName"] = $t["DPO_MinLength"];  $t["PropValue"] = "$minPwdLength $($t['Characters'])" },
                [PSCustomObject]@{ $t["PropName"] = $t["DPO_Complexity"]; $t["PropValue"] = if ($complexityOn) { "✔ $($t['Enabled'])" } else { "✖ $($t['Disabled'])" } },
                [PSCustomObject]@{ $t["PropName"] = $t["DPO_MaxAge"];     $t["PropValue"] = if ($maxPwdAgeDays -gt 0) { "$maxPwdAgeDays $($t['Days'])" } else { $t["NeverExpires"] } },
                [PSCustomObject]@{ $t["PropName"] = $t["DPO_MinAge"];     $t["PropValue"] = "$minPwdAgeDays $($t['Days'])" },
                [PSCustomObject]@{ $t["PropName"] = $t["DPO_History"];    $t["PropValue"] = "$pwdHistoryLen $($t['Passwords'])" },
                [PSCustomObject]@{ $t["PropName"] = $t["DPO_LockoutTh"];  $t["PropValue"] = if ($lockoutThresh -gt 0) { "$lockoutThresh $($t['InvalidAttempts'])" } else { $t["NoLockout"] } },
                [PSCustomObject]@{ $t["PropName"] = $t["DPO_LockoutDur"]; $t["PropValue"] = if ($lockoutThresh -gt 0) { "$lockoutDurMin $($t['Minutes'])" } else { "-" } },
                [PSCustomObject]@{ $t["PropName"] = $t["DPO_ObsWindow"];  $t["PropValue"] = if ($lockoutThresh -gt 0) { "$lockoutWinMin $($t['Minutes'])" } else { "-" } },
                [PSCustomObject]@{ $t["PropName"] = $t["DPO_RevEncrypt"]; $t["PropValue"] = if ($reversibleOn) { "⚠️ $($t['Enabled']) ($($t['Unsafe']))" } else { "$($t['Disabled']) ($($t['Safe']))" } }
            )

            $script:cachedDomainPolicy = $domainPolicyList
            $gridDomain.DataSource = [System.Collections.ArrayList]::new($domainPolicyList)

            # 2. Fine-Grained Password Policies (PSO)
            $psoContainer = [ADSI]"LDAP://CN=Password Settings Objects,CN=System,$defaultNC"
            $psoSearcher = New-Object System.DirectoryServices.DirectorySearcher($psoContainer)
            $psoSearcher.Filter = "(objectClass=msDS-PasswordSettings)"
            $psoResults = $psoSearcher.FindAll()

            $psoList = [System.Collections.ArrayList]::new()

            foreach ($p in $psoResults) {
                $entry = $p.GetDirectoryEntry()
                $pName = $entry.Properties["name"].Value
                $precedence = $entry.Properties["msDS-PasswordSettingsPrecedence"].Value
                $pComplexity = [bool]$entry.Properties["msDS-PasswordComplexityEnabled"].Value
                $pMinLen = $entry.Properties["msDS-MinimumPasswordLength"].Value
                $pHist = $entry.Properties["msDS-PasswordHistoryLength"].Value
                $pLockThresh = $entry.Properties["msDS-LockoutThreshold"].Value

                $pMaxAgeSpan = Convert-LargeIntToTimeSpan $entry.Properties["msDS-MaximumPasswordAge"].Value
                $pMinAgeSpan = Convert-LargeIntToTimeSpan $entry.Properties["msDS-MinimumPasswordAge"].Value
                $pLockDurSpan = Convert-LargeIntToTimeSpan $entry.Properties["msDS-LockoutDuration"].Value

                $pMaxAge = if ($pMaxAgeSpan) { [math]::Round($pMaxAgeSpan.TotalDays, 1) } else { 0 }
                $pMinAge = if ($pMinAgeSpan) { [math]::Round($pMinAgeSpan.TotalDays, 1) } else { 0 }
                $pLockDur = if ($pLockDurSpan) { [math]::Round($pLockDurSpan.TotalMinutes, 0) } else { 0 }

                $appliesTo = @()
                foreach ($app in $entry.Properties["msDS-PSOAppliesTo"]) {
                    $appliesTo += ($app -split ",*..=")[1]
                }
                $appliesToStr = if ($appliesTo.Count -gt 0) { $appliesTo -join ", " } else { $t["NoneAssigned"] }

                [void]$psoList.Add([PSCustomObject]@{
                    $t["PsoName"]          = $pName
                    $t["PsoPrecedence"]    = $precedence
                    $t["PsoMinLength"]     = $pMinLen
                    $t["PsoComplexity"]    = if ($pComplexity) { $t["Yes"] } else { $t["No"] }
                    $t["PsoMaxAge"]        = $pMaxAge
                    $t["PsoMinAge"]        = $pMinAge
                    $t["PsoHistory"]       = $pHist
                    $t["PsoLockoutThresh"] = $pLockThresh
                    $t["PsoLockoutDur"]    = $pLockDur
                    $t["PsoAppliesTo"]     = $appliesToStr
                })
            }

            $sortedPSO = $psoList | Sort-Object { $_.($t["PsoPrecedence"]) }
            $script:cachedPSOs = [System.Collections.ArrayList]::new($sortedPSO)
            $gridPSO.DataSource = $script:cachedPSOs

            $lblStatus.Text = "$($t["StatusDone"]) (PSOs: $($psoList.Count))"
        } catch {
            $lblStatus.Text = "Fehler: " + $_.Exception.Message
        }
    }

    $btnUserCheck.Add_Click({
        $username = $txtUserSearch.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($username)) { return }

        try {
            $rootDSE = [ADSI]"LDAP://RootDSE"
            $defaultNC = $rootDSE.defaultNamingContext
            $uSearcher = New-Object System.DirectoryServices.DirectorySearcher([ADSI]"LDAP://$defaultNC")
            $uSearcher.Filter = "(&(objectCategory=person)(objectClass=user)(sAMAccountName=$username))"
            $uSearcher.PropertiesToLoad.Add("msDS-ResultantPSO") | Out-Null
            $uSearcher.PropertiesToLoad.Add("distinguishedName") | Out-Null
            $uSearcher.PropertiesToLoad.Add("displayName") | Out-Null

            $uRes = $uSearcher.FindOne()
            if (-not $uRes) {
                [System.Windows.Forms.MessageBox]::Show($t["ErrUserNotFound"], "User Check", "OK", "Warning")
                return
            }

            $userResultList = @()
            $appliedPSO = $uRes.Properties["msds-resultantpso"]

            if ($appliedPSO -and $appliedPSO.Count -gt 0) {
                $psoDN = $appliedPSO[0]
                $psoEntry = [ADSI]"LDAP://$psoDN"

                $pName = $psoEntry.Properties["name"].Value
                $pMinLen = $psoEntry.Properties["msDS-MinimumPasswordLength"].Value
                $pComplexity = [bool]$psoEntry.Properties["msDS-PasswordComplexityEnabled"].Value
                $pHist = $psoEntry.Properties["msDS-PasswordHistoryLength"].Value
                $pLockThresh = $psoEntry.Properties["msDS-LockoutThreshold"].Value

                $pMaxAgeSpan = Convert-LargeIntToTimeSpan $psoEntry.Properties["msDS-MaximumPasswordAge"].Value
                $pMaxAge = if ($pMaxAgeSpan) { [math]::Round($pMaxAgeSpan.TotalDays, 1) } else { 0 }

                $userResultList = @(
                    [PSCustomObject]@{ $t["PropName"] = $t["Eff_Policy"];     $t["PropValue"] = "PSO: $pName" },
                    [PSCustomObject]@{ $t["PropName"] = $t["DPO_MinLength"];  $t["PropValue"] = "$pMinLen $($t['Characters'])" },
                    [PSCustomObject]@{ $t["PropName"] = $t["DPO_Complexity"]; $t["PropValue"] = if ($pComplexity) { $t["Enabled"] } else { $t["Disabled"] } },
                    [PSCustomObject]@{ $t["PropName"] = $t["DPO_MaxAge"];     $t["PropValue"] = "$pMaxAge $($t['Days'])" },
                    [PSCustomObject]@{ $t["PropName"] = $t["DPO_History"];    $t["PropValue"] = "$pHist $($t['Passwords'])" },
                    [PSCustomObject]@{ $t["PropName"] = $t["Eff_LockoutTh"];  $t["PropValue"] = "$pLockThresh $($t['Attempts'])" },
                    [PSCustomObject]@{ $t["PropName"] = $t["Eff_PSO_DN"];     $t["PropValue"] = $psoDN }
                )
            } else {
                $userResultList = @(
                    [PSCustomObject]@{ $t["PropName"] = $t["Eff_Policy"]; $t["PropValue"] = $t["InheritedDomain"] }
                ) + $script:cachedDomainPolicy
            }

            $gridUser.DataSource = [System.Collections.ArrayList]::new($userResultList)
            $lblStatus.Text = "$($t["UserResultTitle"]) '$username' ($($t['StatusDone']))."
        } catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Error", "OK", "Error")
        }
    })

    $txtUserSearch.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $_.SuppressKeyPress = $true
            $btnUserCheck.PerformClick()
        }
    })

    $btnExport.Add_Click({
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "CSV-Datei (*.csv)|*.csv"
        $sfd.FileName = "AD_Password_Policies_$((Get-Date).ToString('yyyyMMdd')).csv"

        if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $exportData = @()
            $defaultTypeStr = if ($lang -eq "EN") { "Default Domain Policy" } else { "Standard-Domänenrichtlinie" }
            $psoTypeStr     = if ($lang -eq "EN") { "Fine-Grained PSO" } else { "Feingranulare PSO" }

            foreach ($row in $script:cachedDomainPolicy) {
                $exportData += [PSCustomObject]@{
                    "Typ / Type"               = $defaultTypeStr
                    "Eigenschaft / Setting"    = $row.($t["PropName"])
                    "Wert / Value"             = $row.($t["PropValue"])
                }
            }
            foreach ($row in $script:cachedPSOs) {
                $exportData += [PSCustomObject]@{
                    "Typ / Type"               = "$psoTypeStr ($($row.($t['PsoName'])))"
                    "Eigenschaft / Setting"    = "Precedence: $($row.($t['PsoPrecedence'])), MinLen: $($row.($t['PsoMinLength']))"
                    "Wert / Value"             = "AppliesTo: $($row.($t['PsoAppliesTo']))"
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
# HAUPTFENSTER & 3-SPALTEN SYSTEM-HEADER
# ==============================================================================
$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Text = Get-Text "Title"
$mainForm.Size = New-Object System.Drawing.Size(980, 960)
$mainForm.StartPosition = "CenterScreen"
$mainForm.FormBorderStyle = "FixedDialog"
$mainForm.MaximizeBox = $false
$mainForm.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9)
$mainForm.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

# --- HEADER PANEL (3 Spalten) ---
$pnlHeader = New-Object System.Windows.Forms.Panel
$pnlHeader.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlHeader.Height = 220
$pnlHeader.BackColor = [System.Drawing.Color]::FromArgb(235, 242, 250)
$mainForm.Controls.Add($pnlHeader)

# Sprachumschalter (DE / EN)
$btnLangEN = New-Object System.Windows.Forms.Button
$btnLangEN.Location = New-Object System.Drawing.Point(860, 10)
$btnLangEN.Size = New-Object System.Drawing.Size(45, 26)
$btnLangEN.Text = "EN"
$btnLangEN.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 8, [System.Drawing.FontStyle]::Bold)
$pnlHeader.Controls.Add($btnLangEN)

$btnLangDE = New-Object System.Windows.Forms.Button
$btnLangDE.Location = New-Object System.Drawing.Point(910, 10)
$btnLangDE.Size = New-Object System.Drawing.Size(45, 26)
$btnLangDE.Text = "DE"
$btnLangDE.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 8, [System.Drawing.FontStyle]::Bold)
$pnlHeader.Controls.Add($btnLangDE)

# Hauptüberschrift im Header
$lblHeaderMain = New-Object System.Windows.Forms.Label
$lblHeaderMain.Location = New-Object System.Drawing.Point(18, 10)
$lblHeaderMain.Size = New-Object System.Drawing.Size(830, 24)
$lblHeaderMain.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 11.5, [System.Drawing.FontStyle]::Bold)
$lblHeaderMain.ForeColor = [System.Drawing.Color]::FromArgb(15, 23, 42)
$lblHeaderMain.Text = "💻 $localComputerName | $localUserName"
$pnlHeader.Controls.Add($lblHeaderMain)

# Trennlinie
$lblHeaderLine = New-Object System.Windows.Forms.Label
$lblHeaderLine.Location = New-Object System.Drawing.Point(18, 38)
$lblHeaderLine.Size = New-Object System.Drawing.Size(935, 1)
$lblHeaderLine.BackColor = [System.Drawing.Color]::FromArgb(203, 213, 225)
$pnlHeader.Controls.Add($lblHeaderLine)

# -------------------------------------------------------------
# SPALTE 1: Betriebssystem & Domäne (Links, X = 18, Breite = 300)
# -------------------------------------------------------------
$lblCol1Title = New-Object System.Windows.Forms.Label
$lblCol1Title.Location = New-Object System.Drawing.Point(18, 46)
$lblCol1Title.Size = New-Object System.Drawing.Size(300, 18)
$lblCol1Title.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$lblCol1Title.ForeColor = [System.Drawing.Color]::FromArgb(71, 85, 105)
$pnlHeader.Controls.Add($lblCol1Title)

$lblCol1Content = New-Object System.Windows.Forms.Label
$lblCol1Content.Location = New-Object System.Drawing.Point(18, 66)
$lblCol1Content.Size = New-Object System.Drawing.Size(300, 142)
$lblCol1Content.Font = New-Object System.Drawing.Font("Consolas", 8.5)
$lblCol1Content.ForeColor = [System.Drawing.Color]::FromArgb(15, 23, 42)
$pnlHeader.Controls.Add($lblCol1Content)

# -------------------------------------------------------------
# SPALTE 2: System & Hardware (Mitte, X = 330, Breite = 295)
# -------------------------------------------------------------
$lblCol2Title = New-Object System.Windows.Forms.Label
$lblCol2Title.Location = New-Object System.Drawing.Point(330, 46)
$lblCol2Title.Size = New-Object System.Drawing.Size(295, 18)
$lblCol2Title.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$lblCol2Title.ForeColor = [System.Drawing.Color]::FromArgb(71, 85, 105)
$pnlHeader.Controls.Add($lblCol2Title)

$lblCol2Content = New-Object System.Windows.Forms.Label
$lblCol2Content.Location = New-Object System.Drawing.Point(330, 66)
$lblCol2Content.Size = New-Object System.Drawing.Size(295, 142)
$lblCol2Content.Font = New-Object System.Drawing.Font("Consolas", 8.5)
$lblCol2Content.ForeColor = [System.Drawing.Color]::FromArgb(15, 23, 42)
$pnlHeader.Controls.Add($lblCol2Content)

# -------------------------------------------------------------
# SPALTE 3: Entra ID / Cloud Status (Rechts, X = 640, Breite = 315)
# -------------------------------------------------------------
$lblCol3Title = New-Object System.Windows.Forms.Label
$lblCol3Title.Location = New-Object System.Drawing.Point(640, 46)
$lblCol3Title.Size = New-Object System.Drawing.Size(315, 18)
$lblCol3Title.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$lblCol3Title.ForeColor = [System.Drawing.Color]::FromArgb(71, 85, 105)
$pnlHeader.Controls.Add($lblCol3Title)

$lblCol3Content = New-Object System.Windows.Forms.Label
$lblCol3Content.Location = New-Object System.Drawing.Point(640, 66)
$lblCol3Content.Size = New-Object System.Drawing.Size(315, 142)
$lblCol3Content.Font = New-Object System.Drawing.Font("Consolas", 8.5)
$lblCol3Content.ForeColor = [System.Drawing.Color]::FromArgb(15, 23, 42)
$pnlHeader.Controls.Add($lblCol3Content)

# --- CLIENT TOOLS GROUPBOX ---
$grpClient = New-Object System.Windows.Forms.GroupBox
$grpClient.Location = New-Object System.Drawing.Point(18, 230)
$grpClient.Size = New-Object System.Drawing.Size(935, 250)
$grpClient.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 9, [System.Drawing.FontStyle]::Bold)
$mainForm.Controls.Add($grpClient)

$btnTool3 = New-Object System.Windows.Forms.Button; $btnTool3.Location = "20, 22"; $btnTool3.Size = "895, 46"
$btnTool3.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 8.5); $btnTool3.Add_Click({ Open-ToolEntraStatus }); $grpClient.Controls.Add($btnTool3)

$btnTool4 = New-Object System.Windows.Forms.Button; $btnTool4.Location = "20, 74"; $btnTool4.Size = "895, 46"
$btnTool4.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 8.5); $btnTool4.Add_Click({ Open-ToolGroupsAndGPO }); $grpClient.Controls.Add($btnTool4)

$btnTool5 = New-Object System.Windows.Forms.Button; $btnTool5.Location = "20, 126"; $btnTool5.Size = "895, 46"
$btnTool5.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 8.5); $btnTool5.Add_Click({ Open-ToolWin11Check }); $grpClient.Controls.Add($btnTool5)

$btnTool8 = New-Object System.Windows.Forms.Button; $btnTool8.Location = "20, 178"; $btnTool8.Size = "895, 46"
$btnTool8.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 8.5); $btnTool8.Add_Click({ Show-ClientSoftwareAnalysis }); $grpClient.Controls.Add($btnTool8)

# --- AD TOOLS GROUPBOX ---
$grpAD = New-Object System.Windows.Forms.GroupBox
$grpAD.Location = New-Object System.Drawing.Point(18, 490)
$grpAD.Size = New-Object System.Drawing.Size(935, 410)
$grpAD.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 9, [System.Drawing.FontStyle]::Bold)
$mainForm.Controls.Add($grpAD)

$btnTool1 = New-Object System.Windows.Forms.Button; $btnTool1.Location = "20, 22"; $btnTool1.Size = "895, 46"
$btnTool1.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 8.5); $btnTool1.Add_Click({ Open-ToolLastLogon }); $grpAD.Controls.Add($btnTool1)

$btnTool2 = New-Object System.Windows.Forms.Button; $btnTool2.Location = "20, 74"; $btnTool2.Size = "895, 46"
$btnTool2.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 8.5); $btnTool2.Add_Click({ Open-ToolADAudit }); $grpAD.Controls.Add($btnTool2)

$btnTool6 = New-Object System.Windows.Forms.Button; $btnTool6.Location = "20, 126"; $btnTool6.Size = "895, 46"
$btnTool6.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 8.5); $btnTool6.Add_Click({ Open-ToolDomainOverview }); $grpAD.Controls.Add($btnTool6)

$btnTool7 = New-Object System.Windows.Forms.Button; $btnTool7.Location = "20, 178"; $btnTool7.Size = "895, 46"
$btnTool7.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 8.5); $btnTool7.Add_Click({ Open-ToolOSSupportAudit }); $grpAD.Controls.Add($btnTool7)

$btnTool9 = New-Object System.Windows.Forms.Button; $btnTool9.Location = "20, 230"; $btnTool9.Size = "895, 46"
$btnTool9.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 8.5); $btnTool9.Add_Click({ Show-Tool9-ACLCompare }); $grpAD.Controls.Add($btnTool9)

$btnTool10 = New-Object System.Windows.Forms.Button; $btnTool10.Location = "20, 282"; $btnTool10.Size = "895, 46"
$btnTool10.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 8.5); $btnTool10.Add_Click({ Show-Tool10-OUGroupFinder }); $grpAD.Controls.Add($btnTool10)

$btnTool11 = New-Object System.Windows.Forms.Button; $btnTool11.Location = "20, 334"; $btnTool11.Size = "895, 46"
$btnTool11.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 8.5); $btnTool11.Add_Click({ Show-Tool11-PasswordPolicyAudit }); $grpAD.Controls.Add($btnTool11)

# --- UI REFRESH FUNKTION ---
function Update-UI {
    $mainForm.Text  = Get-Text "Title"
    $grpClient.Text = Get-Text "CategoryClient"
    $grpAD.Text     = Get-Text "CategoryAD"

    $lblCol1Title.Text = Get-Text "LblHdrOS"
    $lblCol2Title.Text = Get-Text "LblHdrSystem"
    $lblCol3Title.Text = Get-Text "LblHdrEntra"

    # Spalte 1: OS & Domäne
    $lblCol1Content.Text = @"
$("{0,-12}: {1}" -f (Get-Text "LblOS"), $osCaption)
$("{0,-12}: {1}" -f (Get-Text "LblBuild"), $osBuildNumber)
$("{0,-12}: {1}" -f (Get-Text "LblVersion"), $osVersionDisplay)
$("{0,-12}: {1}" -f (Get-Text "LblDomain"), $localDomainName)
$("{0,-12}: {1}" -f (Get-Text "LblLogonServer"), $localLogonServer)
"@

    # Spalte 2: System & Hardware
    $lblCol2Content.Text = @"
$("{0,-13}: {1}" -f (Get-Text "LblCompName"), $localComputerName)
$("{0,-13}: {1}" -f (Get-Text "LblManuf"), $localManufacturer)
$("{0,-13}: {1}" -f (Get-Text "LblModel"), $localModel)
$("{0,-13}: {1}" -f (Get-Text "LblSerial"), $localSerial)
$("{0,-13}: {1}" -f (Get-Text "LblSysType"), $localSystemType)
"@

    # Spalte 3: Entra ID / Cloud Status
    $lblCol3Content.Text = @"
$("{0,-12}: {1}" -f (Get-Text "LblJoinStatus"), $localJoinStatus)
$("{0,-12}: {1}" -f (Get-Text "LblPrtStatus"), $localAzureAdPrt)
$("{0,-12}: {1}" -f (Get-Text "LblTenantName"), $localTenantName)
$("{0,-12}: {1}" -f (Get-Text "LblTenantId"), $(if ($localTenantId.Length -gt 16) { $localTenantId.Substring(0,13) + "..." } else { $localTenantId }))
$("{0,-12}: {1}" -f (Get-Text "LblDeviceId"), $(if ($localDeviceId.Length -gt 16) { $localDeviceId.Substring(0,13) + "..." } else { $localDeviceId }))
"@

    $btnTool1.Text  = Get-Text "BtnTool1"
    $btnTool2.Text  = Get-Text "BtnTool2"
    $btnTool3.Text  = Get-Text "BtnTool3"
    $btnTool4.Text  = Get-Text "BtnTool4"
    $btnTool5.Text  = Get-Text "BtnTool5"
    $btnTool6.Text  = Get-Text "BtnTool6"
    $btnTool7.Text  = Get-Text "BtnTool7"
    $btnTool8.Text  = Get-Text "BtnTool8"
    $btnTool9.Text  = Get-Text "BtnTool9"
    $btnTool10.Text = Get-Text "BtnTool10"
    $btnTool11.Text = Get-Text "BtnTool11"

    $btnLangDE.BackColor = if ($script:CurrentLang -eq "DE") { [System.Drawing.Color]::FromArgb(37, 99, 235) } else { [System.Drawing.Color]::LightSteelBlue }
    $btnLangDE.ForeColor = if ($script:CurrentLang -eq "DE") { [System.Drawing.Color]::White } else { [System.Drawing.Color]::Black }
    $btnLangEN.BackColor = if ($script:CurrentLang -eq "EN") { [System.Drawing.Color]::FromArgb(37, 99, 235) } else { [System.Drawing.Color]::LightSteelBlue }
    $btnLangEN.ForeColor = if ($script:CurrentLang -eq "EN") { [System.Drawing.Color]::White } else { [System.Drawing.Color]::Black }
}

$btnLangEN.Add_Click({ $script:CurrentLang = "EN"; Update-UI })
$btnLangDE.Add_Click({ $script:CurrentLang = "DE"; Update-UI })

# Starten
Update-UI
[void]$mainForm.ShowDialog()
