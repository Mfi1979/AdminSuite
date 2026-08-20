<#
================================================================================
 TOOL 3: ENTRA ID / HYBRID JOIN DIAGNOSTIC (dsregcmd /status)
================================================================================
#>
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
