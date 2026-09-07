<#
================================================================================
 CONFIG: I18N SPRACHWÖRTERBUCH & GET-TEXT (GLOBAL SCOPED & HYBRID)
================================================================================
#>

# Sprache global festlegen
if (-not $global:CurrentLang) {
    $global:CurrentLang = "DE"
}

# 1. Globales Basis-Wörterbuch (Garantie gegen NULL-Exceptions)
$global:I18N = @{
    "DE" = @{
        "Title"             = "Active Directory & Entra ID Admin Suite"
        "CategoryClient"    = "Kategorie: Client & Lokale Diagnose (Client Tools)"
        "CategoryAD"        = "Kategorie: Active Directory & Domänen Diagnose (AD Tools)"
        
        "LblHdrOS"          = "BETRIEBSSYSTEM & DOMÄNE"
        "LblHdrSystem"      = "SYSTEM & HARDWARE (Klick = Kopieren)"
        "LblHdrEntra"       = "ENTRA ID / CLOUD STATUS"
        
        "LblOS"             = "OS Edition"
        "LblBuild"          = "Build"
        "LblVersion"        = "Version"
        "LblDomain"         = "Domäne"
        "LblLogonServer"    = "Logonserver"
        
        "LblCompName"       = "Computername"
        "LblManuf"          = "Hersteller"
        "LblModel"          = "Modell"
        "LblSerial"         = "Seriennummer"
        "LblSysType"        = "Systemtyp"
        
        "LblJoinStatus"     = "Join-Status"
        "LblDeviceStatus"   = "DeviceStatus"
        "LblPrtStatus"      = "AzureAD PRT"
        "LblNgcSet"         = "NgcSet"
        "LblTenantName"     = "Tenant Name"
        "LblTenantId"       = "Tenant ID"
        "LblDeviceId"       = "Device ID"

        "BtnExportCsv"      = "📥 CSV Export"

        "BtnTool1"          = "Tool 1: Multi-DC LastLogon Übersicht`n(Automatischer DC- & Site-Scan mit Live-Abfrage)"
        "BtnTool2"          = "Tool 2: Quick AD Audit`n(Inaktive Computerkonten, Passwort nie abgelaufen, Leere Gruppen)"
        "BtnTool3"          = "Tool 3: Entra ID / Hybrid Join Diagnostic`n(Auslesen von dsregcmd /status & Tenant Details)"
        "BtnTool4"          = "Tool 4: Gruppen & GPO Diagnostik`n(Register: User-Gruppen, PC-Gruppen & angewendete GPOs)"
        "BtnTool5"          = "Tool 5: Windows 11 Readiness & OS Details`n(Hardware-Kompatibilität, TPM 2.0, SecureBoot & Update-Stand)"
        "BtnTool6"          = "Tool 6: Domänen-Übersicht & Admin-Audit`n(FSMO-Rollen, AD-Papierkorb & Privilegierte Admins)"
        "BtnTool7"          = "Tool 7: AD Security & Out-of-Support OS Audit`n(Identifikation veralteter & nicht mehr unterstützter Systeme)"
        "BtnTool8"          = "Tool 8: Client Software & App Analyse`n(Win32- & Store-Apps mit Sprachauswertung und CSV-Export)"
        "BtnTool9"          = "Tool 9: Universal AD ACL-Vergleich & Diff`n(Rechtevergleich zweier OUs, Gruppen oder Benutzer)"
        "BtnTool10"         = "Tool 10: OU- & Gruppen-Finder (Wildcard-Suche)`n(Objekte, OU-Struktur & rekursive Gruppenmitgliedschaften)"
        "BtnTool11"         = "Tool 11: AD Kennwortrichtlinien & PSO Audit`n(Default Domain Policy, Fine-Grained PSOs & Benutzer-Check)"
        "BtnTool12"         = "Tool 12: AD Benutzer Kennwortalter & pwdLastSet`n(Passwort-Historie, Ablauf & Alter aller User)"
    }
    "EN" = @{
        "Title"             = "Active Directory & Entra ID Admin Suite"
        "CategoryClient"    = "Category: Client & Local Diagnostics (Client Tools)"
        "CategoryAD"        = "Category: Active Directory & Domain Diagnostics (AD Tools)"
        
        "LblHdrOS"          = "OPERATING SYSTEM & DOMAIN"
        "LblHdrSystem"      = "SYSTEM & HARDWARE (Click = Copy)"
        "LblHdrEntra"       = "ENTRA ID / CLOUD STATUS"
        
        "LblOS"             = "OS Edition"
        "LblBuild"          = "Build"
        "LblVersion"        = "Version"
        "LblDomain"         = "Domain"
        "LblLogonServer"    = "Logon Server"
        
        "LblCompName"       = "Computer Name"
        "LblManuf"          = "Manufacturer"
        "LblModel"          = "Model"
        "LblSerial"         = "Serial Number"
        "LblSysType"        = "System Type"
        
        "LblJoinStatus"     = "Join Status"
        "LblDeviceStatus"   = "Device Status"
        "LblPrtStatus"      = "AzureAD PRT"
        "LblNgcSet"         = "NgcSet"
        "LblTenantName"     = "Tenant Name"
        "LblTenantId"       = "Tenant ID"
        "LblDeviceId"       = "Device ID"

        "BtnExportCsv"      = "📥 Export CSV"

        "BtnTool1"          = "Tool 1: Multi-DC LastLogon Overview`n(Automatic DC & Site Scan with Live Query)"
        "BtnTool2"          = "Tool 2: Quick AD Audit`n(Disabled Accounts, Password Never Expires, Empty Groups)"
        "BtnTool3"          = "Tool 3: Entra ID / Hybrid Join Diagnostic`n(Read dsregcmd /status & Tenant Details)"
        "BtnTool4"          = "Tool 4: Groups & GPO Diagnostics`n(User Groups, PC Groups & Applied GPOs)"
        "BtnTool5"          = "Tool 5: Windows 11 Readiness & OS Details`n(Hardware Compatibility, TPM 2.0, SecureBoot & Update Level)"
        "BtnTool6"          = "Tool 6: Domain Overview & Admin Audit`n(FSMO Roles, AD Recycle Bin & Privileged Admins)"
        "BtnTool7"          = "Tool 7: AD Security & Out-of-Support OS Audit`n(Identification of Outdated & Unsupported Systems)"
        "BtnTool8"          = "Tool 8: Client Software & App Analysis`n(Win32 & Store Apps with Language Analysis & CSV Export)"
        "BtnTool9"          = "Tool 9: Universal AD ACL Compare & Diff`n(Permission Comparison of two OUs, Groups or Users)"
        "BtnTool10"         = "Tool 10: OU & Group Finder (Wildcard Search)`n(Objects, OU Structure & Recursive Group Memberships)"
        "BtnTool11"         = "Tool 11: AD Password Policies & PSO Audit`n(Default Domain Policy, Fine-Grained PSOs & User Check)"
        "BtnTool12"         = "Tool 12: AD User Password Age & pwdLastSet`n(Password Age, Expiry & History of all Users)"
    }
}

# 2. Registrierungs-Funktion für dynamische JSON-Sprachdateien
function Register-LanguageJson {
    param(
        [string]$FileName,
        [string]$JsonContent
    )
    if ([string]::IsNullOrWhiteSpace($JsonContent)) { return }
    try {
        $jsonObj = $JsonContent | ConvertFrom-Json
        $langCode = ($FileName.Split('.')[0].Split('-')[0]).ToUpper()

        if (-not $global:I18N.ContainsKey($langCode)) {
            $global:I18N[$langCode] = @{}
        }

        $jsonObj.PSObject.Properties | ForEach-Object {
            $global:I18N[$langCode][$_.Name] = [string]$_.Value
        }
    } catch {
        Write-Warning "Fehler beim Einlesen von ${FileName}: $($_.Exception.Message)"
    }
}

# 3. Text-Abruffunktion mit robuster Null-Prüfung
function Get-Text {
    param([string]$Key)

    # 1. Gewünschte Sprache prüfen
    if ($global:I18N -and $global:CurrentLang -and $global:I18N.ContainsKey($global:CurrentLang)) {
        $currentDict = $global:I18N[$global:CurrentLang]
        if ($currentDict -and $currentDict.ContainsKey($Key)) {
            return $currentDict[$Key]
        }
    }

    # 2. Fallback auf Deutsch (DE)
    if ($global:I18N -and $global:I18N.ContainsKey("DE")) {
        $deDict = $global:I18N["DE"]
        if ($deDict -and $deDict.ContainsKey($Key)) {
            return $deDict[$Key]
        }
    }

    # 3. Fallback: Key-Name selbst
    return $Key
}