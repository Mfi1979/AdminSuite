<#.SYNOPSIS
    28_Export-02-GraphIdentity.ps1
    Exportiert autonom M365 Lizenzen, Subscriptions inkl. Ablaufdaten/Status,
    Gruppen-Lizenzzuweisungen, Gruppenmitglieder (08b_LicenseGroupsMembers.csv)
    sowie die User-Matrix (06b_LicensesDetails.csv & 07_Users.csv).
    Enthaelt vollstaendiges File-Logging und eine Abschluss-Tabelle aller Exporte.
#>
[CmdletBinding()]
param(
    [string]$ConfigFile   = "28_config.json",
    [string]$MappingFile  = "28_license_mapping.json",
    [string]$TargetFolder = ""
)

if ($host.Name -eq "Windows PowerShell ISE Host") {
    Write-Warning "Bitte in nativer powershell.exe starten."
    return
}

Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process -Force
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls13

# -------------------------------------------------------------------------
# Zielverzeichnis & Logging-Initialisierung
# -------------------------------------------------------------------------
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$baseTenant = Join-Path -Path $scriptDir -ChildPath "TenantData"

if (-not [string]::IsNullOrWhiteSpace($TargetFolder)) {
    $targetDir = $TargetFolder
} else {
    $today = Get-Date -Format "yyyyMMdd"
    $targetDir = Join-Path -Path $baseTenant -ChildPath $today
}

if (-not (Test-Path -Path $targetDir)) { 
    New-Item -ItemType Directory -Path $targetDir -Force | Out-Null 
}

$logFilePath = Join-Path -Path $targetDir -ChildPath "28_GraphIdentity_Export.log"

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS")]
        [string]$Level = "INFO",
        [System.ConsoleColor]$Color = [System.ConsoleColor]::White
    )
    $stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $formattedMsg = "[$stamp] [$Level] $Message"
    Write-Host $formattedMsg -ForegroundColor $Color
    try {
        [System.IO.File]::AppendAllLines($logFilePath, [string[]]@($formattedMsg), [System.Text.Encoding]::UTF8)
    } catch {}
}

# Liste zur Protokollierung aller geschriebenen Export-Dateien
$script:ExportedFilesRegistry = [System.Collections.Generic.List[PSCustomObject]]::new()

function Register-ExportFile {
    param(
        [string]$FilePath,
        [int]$RecordCount,
        [string]$Description
    )
    if (Test-Path -LiteralPath $FilePath) {
        $item = Get-Item -LiteralPath $FilePath
        $sizeKb = [Math]::Round(($item.Length / 1KB), 2)
        $script:ExportedFilesRegistry.Add([PSCustomObject]@{
            "Dateiname"    = $item.Name
            "Datensaetze"  = $RecordCount
            "Groesse_KB"   = $sizeKb
            "Pfad"         = $item.FullName
            "Beschreibung" = $Description
        })
        Write-Log -Message "Exportiert: $($item.Name) ($RecordCount Zeilen, $sizeKb KB)" -Level "SUCCESS" -Color Green
    }
}

$exportDate = Get-Date -Format "yyyyMMdd"
$exportTime = Get-Date -Format "HH:mm:ss"

Write-Log -Message "=== Starte Tool 28: Graph Identitaeten, Lizenzen & Subscriptions ===" -Level "INFO" -Color Cyan
Write-Log -Message "Zielverzeichnis : $targetDir" -Level "INFO" -Color Cyan
Write-Log -Message "Log-Datei       : $logFilePath" -Level "INFO" -Color Cyan

# -------------------------------------------------------------------------
# Externes Lizenz-Mapping laden
# -------------------------------------------------------------------------
$mappingFullPath = Join-Path -Path $scriptDir -ChildPath $MappingFile
$skuNameMap = @{}

if (Test-Path -Path $mappingFullPath) {
    try {
        $rawMap = Get-Content -Path $mappingFullPath -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($prop in $rawMap.PSObject.Properties) {
            $skuNameMap[$prop.Name] = [string]$prop.Value
        }
        Write-Log -Message "Lizenz-Mapping geladen: $($skuNameMap.Count) Eintraege aus '$MappingFile'." -Level "INFO" -Color Green
    } catch {
        Write-Log -Message "Fehler beim Einlesen von '$MappingFile': $($_.Exception.Message)" -Level "WARN" -Color Yellow
    }
} else {
    Write-Log -Message "Keine Mapping-Datei '$mappingFullPath' gefunden. Raw-SKU Fallback aktiv." -Level "WARN" -Color Yellow
}

# -------------------------------------------------------------------------
# Graph Verbindung herstellen
# -------------------------------------------------------------------------
Write-Log -Message "Verbinde mit Microsoft Graph..." -Level "INFO" -Color Yellow
try {
    Import-Module Microsoft.Graph.Authentication -RequiredVersion 2.39.0 -ErrorAction Stop
    Connect-MgGraph -Scopes "User.Read.All", "Group.Read.All", "Organization.Read.All", "Directory.Read.All" -NoWelcome
    Write-Log -Message "Erfolgreich mit Microsoft Graph verbunden." -Level "SUCCESS" -Color Green
} catch {
    Write-Log -Message "Verbindung zu Microsoft Graph fehlgeschlagen: $($_.Exception.Message)" -Level "ERROR" -Color Red
    return
}

# -------------------------------------------------------------------------
# 1. Subscribed SKUs (06_LicensesSummary.csv)
# -------------------------------------------------------------------------
Write-Log -Message "Lese Subscribed SKUs aus..." -Level "INFO" -Color Yellow
$skus = Get-MgSubscribedSku -All
$skuIdToFriendlyMap = @{}
$skuIdToPartNumberMap = @{}
$skuConsumedMap = @{}

$licenseReport = foreach ($sku in $skus) {
    $prep = if ($sku.PrepaidUnits) { $sku.PrepaidUnits.Enabled } else { 0 }
    $cons = if ($sku.ConsumedUnits) { $sku.ConsumedUnits } else { 0 }
    $friendlyName = if ($skuNameMap.ContainsKey($sku.SkuPartNumber)) {
        $skuNameMap[$sku.SkuPartNumber]
    } else {
        $sku.SkuPartNumber
    }

    $skuIdToFriendlyMap[[string]$sku.SkuId] = $friendlyName
    $skuIdToPartNumberMap[[string]$sku.SkuId] = $sku.SkuPartNumber
    $skuConsumedMap[[string]$sku.SkuId] = $cons

    [PSCustomObject]@{
        "Datum"          = $exportDate
        "Uhrzeit"        = $exportTime
        "Quelle"         = "Microsoft Graph / SubscribedSkus"
        "Produktname"    = $friendlyName
        "SKU_PartNumber" = $sku.SkuPartNumber
        "Gekauft"        = $prep
        "Zugewiesen"     = $cons
        "Verfuegbar"     = ($prep - $cons)
        "AppliesTo"      = $sku.AppliesTo
        "SkuId"          = $sku.SkuId
    }
}
$file06 = "$targetDir\06_LicensesSummary.csv"
$licenseReport | Export-Csv -Path $file06 -NoTypeInformation -Encoding UTF8
Register-ExportFile -FilePath $file06 -RecordCount $licenseReport.Count -Description "Aggregierte Tenant-Lizenzkontingente"

# -------------------------------------------------------------------------
# 2. Subscriptions & Laufzeiten (06a_SubscriptionsDetail.csv)
# -------------------------------------------------------------------------
Write-Log -Message "Lese Subscriptions, Laufzeiten und Kontingente aus..." -Level "INFO" -Color Yellow
$subReport = @()
try {
    $subUri = "https://graph.microsoft.com/beta/directory/subscriptions"
    $rawSubs = Invoke-MgGraphRequest -Method GET -Uri $subUri -ErrorAction Stop
    $subList = if ($rawSubs.value) { $rawSubs.value } else { @() }

    foreach ($sub in $subList) {
        $friendlyName = if ($skuNameMap.ContainsKey($sub.skuPartNumber)) {
            $skuNameMap[$sub.skuPartNumber]
        } elseif ($sub.skuPartNumber) {
            $sub.skuPartNumber
        } else {
            $sub.commercialRuntimeName
        }

        $purchased = if ($sub.prepaidUnits -and $sub.prepaidUnits.enabled) { [int]$sub.prepaidUnits.enabled } else { 0 }
        
        $assigned = 0
        if ($null -ne $sub.consumedUnits) {
            $assigned = [int]$sub.consumedUnits
        } elseif ($sub.skuId -and $skuConsumedMap.ContainsKey([string]$sub.skuId)) {
            $assigned = [int]$skuConsumedMap[[string]$sub.skuId]
        }

        $available = [math]::Max(0, ($purchased - $assigned))

        $subReport += [PSCustomObject]@{
            "Datum"                = $exportDate
            "Uhrzeit"              = $exportTime
            "Quelle"               = "Microsoft Graph Beta / Subscriptions"
            "Produktname"          = $friendlyName
            "SKU_PartNumber"       = $sub.skuPartNumber
            "Zugewiesene_Lizenzen" = $assigned
            "Gekaufte_Menge"       = $purchased
            "Verfuegbare_Lizenzen" = $available
            "Subscription_Status"  = $sub.status
            "Ablauf_Verlaengerung" = $sub.nextLifecycleDateTime
            "Billing_Profile"      = if ($sub.billingType) { $sub.billingType } else { "-" }
            "Purchase_Channel"     = if ($sub.purchaseChannel) { $sub.purchaseChannel } else { "Direct" }
            "Product_Type"         = if ($sub.serviceType) { $sub.serviceType } else { "License-based" }
            "Erstellt_Am"          = $sub.createdDateTime
            "SubscriptionId"       = $sub.id
            "SkuId"                = $sub.skuId
        }
    }
} catch {
    Write-Log -Message "Subscriptions konnten nicht geladen werden: $($_.Exception.Message)" -Level "WARN" -Color Yellow
}

if (-not $subReport) {
    $subReport = @([PSCustomObject]@{ 
        "Datum"   = $exportDate
        "Uhrzeit" = $exportTime
        "Quelle"  = "Microsoft Graph Beta / Subscriptions"
        "Hinweis" = "Abonnement-Details erfordern Directory.Read.All Rechte im Tenant." 
    })
}
$file06a = "$targetDir\06a_SubscriptionsDetail.csv"
$subReport | Export-Csv -Path $file06a -NoTypeInformation -Encoding UTF8
Register-ExportFile -FilePath $file06a -RecordCount $subReport.Count -Description "Abonnements, Laufzeiten, Billing-Profile & Kontingente"

# -------------------------------------------------------------------------
# 3. Gruppen & Gruppenlizenzen (08_Groups.csv / 08a_LicenseGroups.csv)
# -------------------------------------------------------------------------
Write-Log -Message "Ermittle Gruppen und gruppenbasierte Lizenzen..." -Level "INFO" -Color Yellow
$groups = Get-MgGroup -All -Property Id, DisplayName, GroupTypes, MailEnabled, SecurityEnabled, OnPremisesSyncEnabled, AssignedLicenses
$licGroupMap = @{}
$groupReport = @()

foreach ($g in $groups) {
    $gType = if ($g.GroupTypes -contains "Unified") { 
        "Microsoft 365" 
    } elseif ($g.SecurityEnabled -and -not $g.MailEnabled) { 
        "Sicherheitsgruppe" 
    } elseif ($g.MailEnabled) { 
        "E-Mail-aktiviert" 
    } else { 
        "Sonstige" 
    }

    $assignedLics = @()
    if ($g.AssignedLicenses -and $g.AssignedLicenses.Count -gt 0) {
        foreach ($l in $g.AssignedLicenses) {
            $sId = [string]$l.SkuId
            $name = if ($skuIdToFriendlyMap.ContainsKey($sId)) { $skuIdToFriendlyMap[$sId] } else { $sId }
            $assignedLics += $name
        }
        $licGroupMap[[string]$g.Id] = @{
            "DisplayName" = $g.DisplayName
            "Licenses"    = $assignedLics
        }
    }

    $groupReport += [PSCustomObject]@{
        "Datum"                = $exportDate
        "Uhrzeit"              = $exportTime
        "Quelle"               = "Microsoft Graph / Groups"
        "DisplayName"          = $g.DisplayName
        "Typ"                  = $gType
        "Herkunft"             = if ($g.OnPremisesSyncEnabled -eq $true) { "AD-Sync (Lokal)" } else { "Entra ID (Cloud-Only)" }
        "Zugewiesene_Lizenzen" = if ($assignedLics.Count -gt 0) { $assignedLics -join "; " } else { "-" }
        "Hat_Gruppenlizenzen"  = if ($assignedLics.Count -gt 0) { "Ja" } else { "Nein" }
        "Id"                   = $g.Id
    }
}

$file08 = "$targetDir\08_Groups.csv"
$groupReport | Export-Csv -Path $file08 -NoTypeInformation -Encoding UTF8
Register-ExportFile -FilePath $file08 -RecordCount $groupReport.Count -Description "Alle Mandanten-Gruppen (Entra ID & AD-Sync)"

$licGroupReport = $groupReport | Where-Object { $_.Hat_Gruppenlizenzen -eq "Ja" }
if ($licGroupReport) {
    $file08a = "$targetDir\08a_LicenseGroups.csv"
    $licGroupReport | Export-Csv -Path $file08a -NoTypeInformation -Encoding UTF8
    Register-ExportFile -FilePath $file08a -RecordCount $licGroupReport.Count -Description "Gruppen mit Lizenzvererbung"
}

# -------------------------------------------------------------------------
# 3b. Mitglieder der Lizenzgruppen (08b_LicenseGroupsMembers.csv)
# -------------------------------------------------------------------------
Write-Log -Message "Lese direkte Mitglieder der Lizenzgruppen aus..." -Level "INFO" -Color Yellow
$licGroupMembersList = @()

foreach ($grpId in $licGroupMap.Keys) {
    $grpName = $licGroupMap[$grpId].DisplayName
    try {
        $members = Get-MgGroupMember -GroupId $grpId -All -ErrorAction SilentlyContinue
        if ($members) {
            foreach ($m in $members) {
                $upn = if ($m.AdditionalProperties.userPrincipalName) { 
                    [string]$m.AdditionalProperties.userPrincipalName 
                } elseif ($m.AdditionalProperties.mail) { 
                    [string]$m.AdditionalProperties.mail 
                } else { 
                    [string]$m.Id 
                }
                
                $disp = if ($m.AdditionalProperties.displayName) { [string]$m.AdditionalProperties.displayName } else { "-" }
                $accEnabled = if ($null -ne $m.AdditionalProperties.accountEnabled) { [string]$m.AdditionalProperties.accountEnabled } else { "-" }
                $mType = if ($m.AdditionalProperties.'@odata.type') { 
                    ($m.AdditionalProperties.'@odata.type' -replace '#microsoft\.graph\.', '') 
                } else { 
                    "Unknown" 
                }

                $licGroupMembersList += [PSCustomObject]@{
                    "Datum"          = $exportDate
                    "Uhrzeit"        = $exportTime
                    "Quelle"         = "Microsoft Graph / GroupMembers"
                    "User"           = $upn
                    "DisplayName"    = $disp
                    "Gruppenname"    = $grpName
                    "GruppenId"      = $grpId
                    "Typ"            = $mType
                    "AccountEnabled" = $accEnabled
                }
            }
        }
    } catch {
        Write-Log -Message "Fehler bei Mitgliedern von '$grpName': $($_.Exception.Message)" -Level "WARN" -Color Yellow
    }
}

if (-not $licGroupMembersList) {
    $licGroupMembersList = @([PSCustomObject]@{
        "Datum"          = $exportDate
        "Uhrzeit"        = $exportTime
        "Quelle"         = "Microsoft Graph / GroupMembers"
        "User"           = "-"
        "DisplayName"    = "-"
        "Gruppenname"    = "-"
        "GruppenId"      = "-"
        "Typ"            = "-"
        "AccountEnabled" = "-"
    })
}
$file08b = "$targetDir\08b_LicenseGroupsMembers.csv"
$licGroupMembersList | Export-Csv -Path $file08b -NoTypeInformation -Encoding UTF8
Register-ExportFile -FilePath $file08b -RecordCount $licGroupMembersList.Count -Description "Mitglieder aller lizenzierten Gruppen"

# -------------------------------------------------------------------------
# 4. Benutzer (06b_LicensesDetails.csv & 07_Users.csv)
# -------------------------------------------------------------------------
Write-Log -Message "Lese Benutzer aus und erstelle Lizenz-Detailmatrix..." -Level "INFO" -Color Yellow
$users = Get-MgUser -All -Property Id, DisplayName, UserPrincipalName, AccountEnabled, OnPremisesSyncEnabled, AssignedLicenses, LicenseAssignmentStates, Mail, Department

$userSummaryList = @()
$licenseDetailsList = @()

foreach ($u in $users) {
    $licCount = if ($u.AssignedLicenses) { $u.AssignedLicenses.Count } else { 0 }
    $userLicDetailsCompact = @()
    $assignmentTypes = @()

    if ($u.LicenseAssignmentStates -and $u.LicenseAssignmentStates.Count -gt 0) {
        foreach ($state in $u.LicenseAssignmentStates) {
            $sId = [string]$state.SkuId
            $licFriendly = if ($skuIdToFriendlyMap.ContainsKey($sId)) { $skuIdToFriendlyMap[$sId] } else { $sId }
            $skuPart = if ($skuIdToPartNumberMap.ContainsKey($sId)) { $skuIdToPartNumberMap[$sId] } else { "-" }

            $isGroup = [string]::IsNullOrWhiteSpace($state.AssignedByGroup) -eq $false
            $assignMode = if ($isGroup) { "Gruppenbasiert" } else { "Direkt" }
            
            $srcGroupName = "-"
            $srcGroupId = "-"
            if ($isGroup) {
                $srcGroupId = [string]$state.AssignedByGroup
                $srcGroupName = if ($licGroupMap.ContainsKey($srcGroupId)) { $licGroupMap[$srcGroupId].DisplayName } else { $srcGroupId }
                if (-not ($assignmentTypes -contains "Gruppenbasiert")) { $assignmentTypes += "Gruppenbasiert" }
                $userLicDetailsCompact += "$licFriendly [Gruppe: $srcGroupName]"
            } else {
                if (-not ($assignmentTypes -contains "Direkt")) { $assignmentTypes += "Direkt" }
                $userLicDetailsCompact += "$licFriendly [Direkt]"
            }

            # 1 Zeile pro zugewiesener Lizenz pro Benutzer
            $licenseDetailsList += [PSCustomObject]@{
                "Datum"             = $exportDate
                "Uhrzeit"           = $exportTime
                "Quelle"            = "Microsoft Graph / Users"
                "UserPrincipalName" = $u.UserPrincipalName
                "DisplayName"       = $u.DisplayName
                "AccountEnabled"    = $u.AccountEnabled
                "Identitaetstyp"    = if ($u.OnPremisesSyncEnabled -eq $true) { "AD-Sync (Hybrid)" } else { "Entra ID (Cloud-Only)" }
                "Produktname"       = $licFriendly
                "SKU_PartNumber"    = $skuPart
                "Zuweisungsart"     = $assignMode
                "Lizenzgruppe_Name" = $srcGroupName
                "Lizenzgruppe_Id"   = $srcGroupId
                "Department"        = $u.Department
                "Mail"              = $u.Mail
                "SkuId"             = $sId
                "UserId"            = $u.Id
            }
        }
    }

    $finalAssignType = if ($assignmentTypes.Count -gt 1) {
        "Gemischt (Direkt & Gruppe)"
    } elseif ($assignmentTypes.Count -eq 1) {
        $assignmentTypes[0]
    } else {
        "-"
    }

    $userSummaryList += [PSCustomObject]@{
        "Datum"                = $exportDate
        "Uhrzeit"              = $exportTime
        "Quelle"               = "Microsoft Graph / Users"
        "DisplayName"          = $u.DisplayName
        "UserPrincipalName"    = $u.UserPrincipalName
        "AccountEnabled"       = $u.AccountEnabled
        "Identitaetstyp"       = if ($u.OnPremisesSyncEnabled -eq $true) { "AD-Sync (Hybrid)" } else { "Entra ID (Cloud-Only)" }
        "Zuweisungsart"        = $finalAssignType
        "Lizenzanzahl"         = $licCount
        "Zugewiesene_Lizenzen" = if ($userLicDetailsCompact.Count -gt 0) { $userLicDetailsCompact -join "; " } else { "-" }
        "Abteilung"            = $u.Department
        "Mail"                 = $u.Mail
    }
}

$file06b = "$targetDir\06b_LicensesDetails.csv"
$licenseDetailsList | Export-Csv -Path $file06b -NoTypeInformation -Encoding UTF8
Register-ExportFile -FilePath $file06b -RecordCount $licenseDetailsList.Count -Description "Lizenz-Matrix: 1 Zeile je Benutzer je Lizenz"

$file07 = "$targetDir\07_Users.csv"
$userSummaryList | Export-Csv -Path $file07 -NoTypeInformation -Encoding UTF8
Register-ExportFile -FilePath $file07 -RecordCount $userSummaryList.Count -Description "Benutzer-Matrix: 1 Zeile je Benutzer"

# -------------------------------------------------------------------------
# 5. Detaillierte Abschluss-Uebersicht
# -------------------------------------------------------------------------
Write-Host ""
Write-Log -Message "==========================================================================================" -Level "INFO" -Color Cyan
Write-Log -Message "               ABSCHLUSS-UEBERSICHT: EXPORTIERTE DATEIEN                                  " -Level "INFO" -Color Cyan
Write-Log -Message "==========================================================================================" -Level "INFO" -Color Cyan

$summaryTable = $script:ExportedFilesRegistry | Format-Table -Property Dateiname, Datensaetze, Groesse_KB, Beschreibung -AutoSize | Out-String
Write-Host $summaryTable -ForegroundColor Green
try {
    [System.IO.File]::AppendAllLines($logFilePath, [string[]]@($summaryTable), [System.Text.Encoding]::UTF8)
} catch {}

Write-Log -Message "Gesamter Export erfolgreich beendet! Alle Daten liegen in: $targetDir" -Level "SUCCESS" -Color Green