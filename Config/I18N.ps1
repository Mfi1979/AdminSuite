<#
================================================================================
 CONFIG: MEHRSPRACHIGKEIT (I18N: DE / EN)
================================================================================
#>
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
