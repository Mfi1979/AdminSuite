<#
================================================================================
 TOOL 5: WINDOWS 11 READINESS & HARDWARE CHECK
================================================================================
#>
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
