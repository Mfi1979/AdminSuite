<#
.SYNOPSIS
    28_Export-01-Exchange.ps1
    Exportiert Perimeter-Scope (Egress-IP, DNS/MX/SPF/DMARC), Mailbox-Inventar inkl. Weiterleitungen sowie Connectors & Hygiene-Policies.
#>
[CmdletBinding()]
param([string]$ConfigFile = "28_config.json")

if ($host.Name -eq "Windows PowerShell ISE Host") {
    Write-Warning "PowerShell ISE erkannt! Bitte in normaler powershell.exe starten."
    return
}

[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls13

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$configFullPath = Join-Path -Path $scriptDir -ChildPath $ConfigFile

$targetUPN = ""
$additionalDomains = [System.Collections.Generic.List[string]]::new()

if (Test-Path -Path $configFullPath) {
    try {
        $cfg = Get-Content -Path $configFullPath -Raw | ConvertFrom-Json
        if ($cfg.TargetUPN) { $targetUPN = [string]$cfg.TargetUPN }
        if ($cfg.AdditionalDomains) { foreach ($d in $cfg.AdditionalDomains) { $additionalDomains.Add([string]$d) } }
    } catch {}
}

while ([string]::IsNullOrWhiteSpace($targetUPN) -or -not ($targetUPN -like "*@*.*")) {
    $targetUPN = Read-Host "Bitte Admin-UPN eingeben"
}

$primaryDomain = ($targetUPN.Split("@")[-1]).Trim()
if (-not $additionalDomains.Contains($primaryDomain)) { 
    $additionalDomains.Insert(0, $primaryDomain) 
}

$ts = Get-Date -Format "yyyyMMdd_HHmm"
$targetDir = Join-Path -Path $scriptDir -ChildPath "TenantData\$ts"
if (-not (Test-Path -Path $targetDir)) { 
    New-Item -ItemType Directory -Path $targetDir -Force | Out-Null 
}

Write-Host "=== [1/4] EXPORT: EXCHANGE, PERIMETER & HYGIENE ===" -ForegroundColor Cyan
Write-Host "Zielverzeichnis: $targetDir" -ForegroundColor Cyan

# WAN Egress IP & DNS
Write-Host "Ermittle WAN-Egress IP & DNS Scope..." -ForegroundColor Yellow
$detectedIp = "Unbekannt"
foreach ($u in @("https://api.ipify.org", "https://icanhazip.com", "https://ifconfig.me/ip")) {
    try {
        $wc = New-Object System.Net.WebClient
        $res = ($wc.DownloadString($u)).Trim()
        if ($res -match "^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$") { 
            $detectedIp = $res
            break 
        }
    } catch {}
}

$dnsScopeList = @()
foreach ($dom in $additionalDomains) {
    $mxRec = ""
    $spfRec = ""
    $dmarcRec = ""
    try {
        $mxQuery = Resolve-DnsName -Name $dom -Type MX -ErrorAction SilentlyContinue
        if ($mxQuery) { $mxRec = ($mxQuery.NameExchange -join "; ") }
        
        $txtQuery = Resolve-DnsName -Name $dom -Type TXT -ErrorAction SilentlyContinue
        if ($txtQuery) {
            $spfs = $txtQuery | Where-Object { $_.Strings -like "v=spf1*" }
            if ($spfs) { $spfRec = ($spfs.Strings -join " ") }
        }
        
        $dmarcQuery = Resolve-DnsName -Name "_dmarc.$dom" -Type TXT -ErrorAction SilentlyContinue
        if ($dmarcQuery) { $dmarcRec = [string]($dmarcQuery.Strings -join " ") }
    } catch {}

    $dnsScopeList += [PSCustomObject]@{
        "Domaene"           = $dom
        "Egress_WAN_IP"     = $detectedIp
        "MX_Records"        = $mxRec
        "SPF_Record"        = $spfRec
        "DMARC_Policy"      = $dmarcRec
        "Tenable_Scan_Scope"= "IP: $detectedIp | MX: $mxRec"
    }
}
$dnsScopeList | Export-Csv -Path "$targetDir\01_PerimeterScope.csv" -NoTypeInformation -Encoding UTF8

# Exchange Online Verbindung
Write-Host "Verbinde mit Exchange Online..." -ForegroundColor Yellow
try {
    Connect-ExchangeOnline -UserPrincipalName $targetUPN -DisableWAM -ShowBanner:$false -ErrorAction Stop
} catch {
    Connect-ExchangeOnline -UserPrincipalName $targetUPN -DeviceCode -ShowBanner:$false
}

# Postfaecher & Forwarding
Write-Host "Exportiere Postfaecher..." -ForegroundColor Yellow
$rawMbs = Get-EXOMailbox -ResultSize Unlimited -Properties RecipientTypeDetails, ArchiveStatus, ForwardingAddress, ForwardingSmtpAddress, DeliverToMailboxAndForward

$allMbList = @()
foreach ($mb in $rawMbs) {
    $hasFwd = if ($mb.ForwardingAddress -or $mb.ForwardingSmtpAddress) { "Ja" } else { "Nein" }
    $fwdTarget = if ($mb.ForwardingSmtpAddress) { 
        $mb.ForwardingSmtpAddress 
    } elseif ($mb.ForwardingAddress) { 
        $mb.ForwardingAddress 
    } else { 
        "-" 
    }

    $allMbList += [PSCustomObject]@{
        "DisplayName"       = $mb.DisplayName
        "UserPrincipalName" = $mb.UserPrincipalName
        "Postfachtyp"       = [string]$mb.RecipientTypeDetails
        "ForwardingAktiv"   = $hasFwd
        "ForwardingZiel"    = $fwdTarget
        "KopieImPostfach"   = [string]$mb.DeliverToMailboxAndForward
        "ArchivStatus"      = [string]$mb.ArchiveStatus
    }
}
$allMbList | Export-Csv -Path "$targetDir\02_MailboxesAll.csv" -NoTypeInformation -Encoding UTF8

$mbSummary = $allMbList | Group-Object Postfachtyp | Select-Object @{N="Postfachtyp";E={$_.Name}}, @{N="Anzahl";E={$_.Count}}
$mbSummary | Export-Csv -Path "$targetDir\03_MailboxTypeSummary.csv" -NoTypeInformation -Encoding UTF8

# Connectors & Hygiene
Write-Host "Exportiere Connectors & Hygiene..." -ForegroundColor Yellow
$inConnectors = Get-InboundConnector -ErrorAction SilentlyContinue
$outConnectors = Get-OutboundConnector -ErrorAction SilentlyContinue
$connReport = @()

if ($inConnectors) {
    foreach ($inC in $inConnectors) {
        $connReport += [PSCustomObject]@{ 
            "Richtung"       = "Inbound"
            "Name"           = [string]$inC.Name
            "Aktiviert"      = [string]$inC.Enabled
            "Typ"            = [string]$inC.ConnectorType
            "Scope"          = ($inC.SenderDomains -join "; ")
            "SmartHosts_IPs" = ($inC.SenderIPAddresses -join "; ") 
        }
    }
}
if ($outConnectors) {
    foreach ($outC in $outConnectors) {
        $connReport += [PSCustomObject]@{ 
            "Richtung"       = "Outbound"
            "Name"           = [string]$outC.Name
            "Aktiviert"      = [string]$outC.Enabled
            "Typ"            = [string]$outC.ConnectorType
            "Scope"          = ($outC.RecipientDomains -join "; ")
            "SmartHosts_IPs" = ($outC.SmartHosts -join "; ") 
        }
    }
}
if (-not $connReport) { 
    $connReport = @([PSCustomObject]@{ "Hinweis" = "Keine benutzerdefinierten Connectors." }) 
}
$connReport | Export-Csv -Path "$targetDir\04_Connectors.csv" -NoTypeInformation -Encoding UTF8

$hygieneReport = @()
try {
    $antiPhish = Get-AntiPhishPolicy -ErrorAction SilentlyContinue
    if ($antiPhish) {
        foreach ($p in $antiPhish) {
            $hygieneReport += [PSCustomObject]@{ 
                "Typ"       = "Anti-Phishing"
                "Name"      = [string]$p.Name
                "Aktiviert" = [string]$p.Enabled
                "Details"   = "Threshold: $($p.PhishThresholdLevel), SpoofIntel: $($p.EnableSpoofIntelligence)" 
            }
        }
    }
    $antiSpam = Get-HostedContentFilterPolicy -ErrorAction SilentlyContinue
    if ($antiSpam) {
        foreach ($s in $antiSpam) {
            $hygieneReport += [PSCustomObject]@{ 
                "Typ"       = "Anti-Spam"
                "Name"      = [string]$s.Name
                "Aktiviert" = "True"
                "Details"   = "BulkThreshold: $($s.BulkThreshold), Action: $($s.SpamAction)" 
            }
        }
    }
} catch {}

if (-not $hygieneReport) { 
    $hygieneReport = @([PSCustomObject]@{ "Hinweis" = "Keine Richtlinien auslesbar." }) 
}
$hygieneReport | Export-Csv -Path "$targetDir\05_HygienePolicies.csv" -NoTypeInformation -Encoding UTF8

Write-Host "Exchange-Export erfolgreich abgeschlossen!" -ForegroundColor Green