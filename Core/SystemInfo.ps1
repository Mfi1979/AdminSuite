<#
================================================================================
 CORE: SYSTEM- & HARDWARE-PARAMETER & ENTRA ID STATUS
================================================================================
#>

$global:localComputerName = $env:COMPUTERNAME
$global:localUserName     = $env:USERNAME
$localComputerName        = $global:localComputerName
$localUserName            = $global:localUserName

# Hardware- & BIOS-Details
$cs   = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue
$bios = Get-CimInstance -ClassName Win32_BIOS -ErrorAction SilentlyContinue

$global:localManufacturer = if ($cs.Manufacturer) { ($cs.Manufacturer).Trim() } else { "N/A" }
$global:localModel        = if ($cs.Model) { ($cs.Model).Trim() } else { "N/A" }
$global:localSystemType   = if ($cs.SystemType) { ($cs.SystemType).Trim() } else { "N/A" }
$global:localSerial       = if ($bios.SerialNumber) { ($bios.SerialNumber).Trim() } else { "N/A" }

$localManufacturer = $global:localManufacturer
$localModel        = $global:localModel
$localSystemType   = $global:localSystemType
$localSerial       = $global:localSerial

# Domäne & Logonserver
$global:localDomainName   = if ($cs.PartOfDomain -and $cs.Domain) { $cs.Domain } else { "WORKGROUP" }
$global:localLogonServer  = if ($env:LOGONSERVER) { $env:LOGONSERVER.TrimStart('\') } else { "Lokal" }
$localDomainName          = $global:localDomainName
$localLogonServer         = $global:localLogonServer

# Entra ID / Hybrid Join Details (Registry CloudDomainJoin + dsregcmd)
$global:localJoinStatus   = "Nicht gekoppelt"
$global:localAzureDevStat = "Workgroup / Not Registered"
$global:localAzureAdPrt   = "NO"
$global:localNgcSet       = "NO"
$global:localTenantName   = "N/A"
$global:localTenantId     = "N/A"
$global:localDeviceId     = "N/A"

# 1. Priorität: Direkter Registry-Abruf über CloudDomainJoin
try {
    $cdjJoin = Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\CloudDomainJoin\JoinInfo\*" -ErrorAction SilentlyContinue
    if ($cdjJoin) {
        if ($cdjJoin.TenantId) { $global:localTenantId = $cdjJoin.TenantId }
        if ($cdjJoin.DeviceId) { $global:localDeviceId = $cdjJoin.DeviceId }
    }
    $cdjTenant = Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\CloudDomainJoin\TenantInfo\*" -ErrorAction SilentlyContinue
    if ($cdjTenant -and $cdjTenant.DisplayName) {
        $global:localTenantName = $cdjTenant.DisplayName
    }
} catch {}

# 2. Ergänzung & Statusauswertung über dsregcmd
try {
    $dsreg = dsregcmd /status 2>$null

    $isAzureJoined      = ($dsreg -match "AzureAdJoined\s*:\s*YES")
    $isEnterpriseJoined = ($dsreg -match "EnterpriseJoined\s*:\s*YES")
    $isDomainJoined     = ($dsreg -match "DomainJoined\s*:\s*YES")

    # Join-Status Ermittlung
    if ($isAzureJoined -and $isDomainJoined) {
        $global:localJoinStatus   = "Hybrid Joined"
        $global:localAzureDevStat = "Hybrid Azure AD Joined"
    } elseif ($isAzureJoined) {
        $global:localJoinStatus   = "Azure AD Joined"
        $global:localAzureDevStat = "Entra ID Joined (Pure Cloud)"
    } elseif ($isEnterpriseJoined) {
        $global:localJoinStatus   = "Enterprise Joined"
        $global:localAzureDevStat = "On-Premises DRS Registered"
    } elseif ($isDomainJoined) {
        $global:localJoinStatus   = "AD Domain Joined"
        $global:localAzureDevStat = "On-Premises AD Joined"
    }

    # AzureAD PRT Status
    if ($dsreg -match "AzureAdPrt\s*:\s*YES") { 
        $global:localAzureAdPrt = "YES" 
    }

    # NgcSet (Windows Hello for Business Status)
    if ($dsreg -match "NgcSet\s*:\s*YES") { 
        $global:localNgcSet = "YES" 
    }

    # Fallbacks für Tenant-Daten falls Registry leer war
    if ($global:localTenantName -eq "N/A" -and ($dsreg -match "TenantName\s*:\s*(.+)")) { 
        $global:localTenantName = $matches[1].Trim() 
    }
    if ($global:localTenantId -eq "N/A"   -and ($dsreg -match "TenantId\s*:\s*([a-fA-F0-9\-]+)")) { 
        $global:localTenantId   = $matches[1].Trim() 
    }
    if ($global:localDeviceId -eq "N/A"   -and ($dsreg -match "DeviceId\s*:\s*([a-fA-F0-9\-]+)")) { 
        $global:localDeviceId   = $matches[1].Trim() 
    }
} catch {}

# Lokale Aliase für Abwärtskompatibilität
$localJoinStatus   = $global:localJoinStatus
$localAzureDevStat = $global:localAzureDevStat
$localAzureAdPrt   = $global:localAzureAdPrt
$localNgcSet       = $global:localNgcSet
$localTenantName   = $global:localTenantName
$localTenantId     = $global:localTenantId
$localDeviceId     = $global:localDeviceId

# Betriebssystem Details mit Fallback
$osInfo            = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue
$global:osCaption  = if ($osInfo.Caption) { ($osInfo.Caption -replace "Microsoft ", "").Trim() } else { "Windows" }
$global:osBuildNumber = if ($osInfo.BuildNumber) { $osInfo.BuildNumber } else { "N/A" }
$global:osBuildFull   = "$($osInfo.Version) (Build $global:osBuildNumber)"
$global:osArchFormatted = $osInfo.OSArchitecture
$global:osInstallDateForm = if ($osInfo.InstallDate) { $osInfo.InstallDate.ToString("dd.MM.yyyy HH:mm") } else { "N/A" }
$global:osPatchFormatted  = (Get-HotFix -ErrorAction SilentlyContinue | Sort-Object InstalledOn -Descending | Select-Object -First 1).HotFixID

$osCaption         = $global:osCaption
$osBuildNumber     = $global:osBuildNumber
$osBuildFull       = $global:osBuildFull
$osArchFormatted   = $global:osArchFormatted
$osInstallDateForm = $global:osInstallDateForm
$osPatchFormatted  = $global:osPatchFormatted

# DisplayVersion / Release-Erkennung
$osVersionDisplay = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name DisplayVersion -ErrorAction SilentlyContinue).DisplayVersion
if (-not $osVersionDisplay) {
    $osVersionDisplay = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name ReleaseId -ErrorAction SilentlyContinue).ReleaseId
}
if (-not $osVersionDisplay) {
    $bNum = 0
    [int]::TryParse($global:osBuildNumber, [ref]$bNum) | Out-Null
    $osVersionDisplay = switch ($bNum) {
        26200 { "25H2" }
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
$global:osVersionDisplay = $osVersionDisplay