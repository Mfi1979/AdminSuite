<#.SYNOPSIS
    28_Export-02-GraphIdentity.ps1
    Exportiert M365 Lizenzen, Subscriptions inkl. Laufzeiten/Status/Mengen,
    Gruppen-Lizenzzuweisungen, Gruppenmitglieder (08b_LicenseGroupsMembers.csv)
    sowie die User-Matrix (07_Users.csv & 06b_LicensesDetails.csv).
    Alle Exports enthalten einheitlich Datum, Uhrzeit und Quelle.
#>
[CmdletBinding()]
param(
    [string]$ConfigFile = "28_config.json",
    [string]$MappingFile = "28_license_mapping.json"
)

if ($host.Name -eq "Windows PowerShell ISE Host") {
    Write-Warning "Bitte in nativer powershell.exe starten."
    return
}

Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process -Force
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls13

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$baseTenant = Join-Path -Path $scriptDir -ChildPath "TenantData"
if (-not (Test-Path $baseTenant)) { 
    New-Item -ItemType Directory -Path $baseTenant -Force | Out-Null 
}

# Sicherer Filter: Nur echte Zeitstempel-Ordner nach letztem Schreibdatum sortieren
$latest = Get-ChildItem -Path $baseTenant -Directory | 
          Where-Object { $_.Name -match '^\d{8}_\d{4}$' } | 
          Sort-Object LastWriteTime -Descending | 
          Select-Object -First 1

$targetDir = if ($latest) { 
    $latest.FullName 
} else {
    $ts = Get-Date -Format "yyyyMMdd_HHmm"
    $nd = Join-Path -Path $baseTenant -ChildPath $ts
    New-Item -ItemType Directory -Path $nd -Force | Out-Null
    $nd
}

$exportDate = Get-Date -Format "yyyyMMdd"
$exportTime = Get-Date -Format "HH:mm:ss"

Write-Host "=== [2/4] EXPORT: GRAPH IDENTITAETEN, LIZENZEN & SUBSCRIPTIONS ===" -ForegroundColor Cyan
Write-Host "Zielverzeichnis: $targetDir" -ForegroundColor Cyan

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
        Write-Host "Lizenz-Mapping erfolgreich geladen ($($skuNameMap.Count) Eintraege)." -ForegroundColor Green
    } catch {
        Write-Warning "Fehler beim Einlesen von $MappingFile : $($_.Exception.Message)"
    }
} else {
    Write-Warning "Keine Mapping-Datei '$mappingFullPath' gefunden."
}

# -------------------------------------------------------------------------
# Graph Verbindung herstellen
# -------------------------------------------------------------------------
Import-Module Microsoft.Graph.Authentication -RequiredVersion 2.39.0 -ErrorAction Stop
Connect-MgGraph -Scopes "User.Read.All", "Group.Read.All", "Organization.Read.All", "Directory.Read.All" -NoWelcome

# -------------------------------------------------------------------------
# 1. Subscribed SKUs (06_LicensesSummary.csv)
# -------------------------------------------------------------------------
Write-Host "Exportiere Lizenzen (Subscribed SKUs)..." -ForegroundColor Yellow
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
$licenseReport | Export-Csv -Path "$targetDir\06_LicensesSummary.csv" -NoTypeInformation -Encoding UTF8

# -------------------------------------------------------------------------
# 2. Subscriptions & Laufzeiten (06a_SubscriptionsDetail.csv)
# -------------------------------------------------------------------------
Write-Host "Ermittle Subscription-Laufzeiten und Status..." -ForegroundColor Yellow
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
    Write-Warning "Subscriptions konnten nicht geladen werden: $($_.Exception.Message)"
}

if (-not $subReport) {
    $subReport = @([PSCustomObject]@{ 
        "Datum"   = $exportDate
        "Uhrzeit" = $exportTime
        "Quelle"  = "Microsoft Graph Beta / Subscriptions"
        "Hinweis" = "Abonnement-Details erfordern Directory.Read.All Rechte im Tenant." 
    })
}
$subReport | Export-Csv -Path "$targetDir\06a_SubscriptionsDetail.csv" -NoTypeInformation -Encoding UTF8

# -------------------------------------------------------------------------
# 3. Gruppen & Gruppenlizenzen (08_Groups.csv / 08a_LicenseGroups.csv)
# -------------------------------------------------------------------------
Write-Host "Ermittle gruppenbasierte Lizenzzuweisungen..." -ForegroundColor Yellow
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
$groupReport | Export-Csv -Path "$targetDir\08_Groups.csv" -NoTypeInformation -Encoding UTF8

$licGroupReport = $groupReport | Where-Object { $_.Hat_Gruppenlizenzen -eq "Ja" }
if ($licGroupReport) {
    $licGroupReport | Export-Csv -Path "$targetDir\08a_LicenseGroups.csv" -NoTypeInformation -Encoding UTF8
}

# -------------------------------------------------------------------------
# 3b. Mitglieder der Lizenzgruppen (08b_LicenseGroupsMembers.csv)
# -------------------------------------------------------------------------
Write-Host "Lese Mitglieder der Lizenzgruppen aus..." -ForegroundColor Yellow
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
        Write-Warning "Fehler beim Auslesen der Mitglieder von Gruppe '$grpName': $($_.Exception.Message)"
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
$licGroupMembersList | Export-Csv -Path "$targetDir\08b_LicenseGroupsMembers.csv" -NoTypeInformation -Encoding UTF8

# -------------------------------------------------------------------------
# 4. Benutzer-Auswertung (06b_LicensesDetails.csv & 07_Users.csv)
# -------------------------------------------------------------------------
Write-Host "Exportiere Benutzer und erstelle Lizenz-Detailmatrix..." -ForegroundColor Yellow
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

            # 1 Zeile pro Lizenz je User fuer 06b_LicensesDetails.csv
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

# Dateiexport: 06b_LicensesDetails.csv & 07_Users.csv
$licenseDetailsList | Export-Csv -Path "$targetDir\06b_LicensesDetails.csv" -NoTypeInformation -Encoding UTF8
$userSummaryList | Export-Csv -Path "$targetDir\07_Users.csv" -NoTypeInformation -Encoding UTF8

Write-Host "06a_SubscriptionsDetail.csv erfolgreich exportiert!" -ForegroundColor Green
Write-Host "06b_LicensesDetails.csv ($($licenseDetailsList.Count) Eintraege) erfolgreich exportiert!" -ForegroundColor Green
Write-Host "07_Users.csv ($($userSummaryList.Count) Eintraege) erfolgreich exportiert!" -ForegroundColor Green
Write-Host "08b_LicenseGroupsMembers.csv ($($licGroupMembersList.Count) Eintraege) erfolgreich exportiert!" -ForegroundColor Green
Write-Host "Graph-Identitaeten & Lizenz-Auswertung vollstaendig abgeschlossen!" -ForegroundColor Green