<#
================================================================================
 CORE: SYSTEM- & HARDWARE-PARAMETER & ENTRA ID STATUS
================================================================================
#>

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
