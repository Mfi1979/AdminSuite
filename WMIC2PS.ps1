[PSCustomObject]@{
    ComputerName = $env:COMPUTERNAME
    Manufacturer = (Get-CimInstance Win32_ComputerSystem).Manufacturer
    Model        = (Get-CimInstance Win32_ComputerSystem).Model
    SystemType   = (Get-CimInstance Win32_ComputerSystem).SystemType
    SerialNumber = (Get-CimInstance Win32_BIOS).SerialNumber
    OSCaption    = (Get-CimInstance Win32_OperatingSystem).Caption
    BuildNumber  = (Get-CimInstance Win32_OperatingSystem).BuildNumber
    Version      = (Get-CimInstance Win32_OperatingSystem).Version
} | Format-List
