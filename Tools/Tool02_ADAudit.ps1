<#
================================================================================
 TOOL 2: QUICK AD AUDIT & SCHNELLABFRAGEN
================================================================================
#>
function Open-ToolADAudit {
    if (-not (Assert-DomainJoined)) { return }
    $subForm = New-Object System.Windows.Forms.Form
    $subForm.Text = "Tool 2: Quick AD Audit & Abfragen"
    $subForm.Size = New-Object System.Drawing.Size(1000, 600)
    $subForm.StartPosition = "CenterParent"
    $subForm.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9)

    $pnlTop = New-Object System.Windows.Forms.Panel
    $pnlTop.Dock = "Top"; $pnlTop.Height = 60; $subForm.Controls.Add($pnlTop)

    $cmbQuery = New-Object System.Windows.Forms.ComboBox
    $cmbQuery.DropDownStyle = "DropDownList"; $cmbQuery.Location = "15, 18"; $cmbQuery.Width = 300
    [void]$cmbQuery.Items.AddRange(@("Deaktivierte Computerkonten", "Passwort läuft nie ab (User)", "Leere AD-Gruppen"))
    $cmbQuery.SelectedIndex = 0; $pnlTop.Controls.Add($cmbQuery)

    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Text = "Abfrage starten"; $btnRun.Location = "330, 16"; $btnRun.Size = "130, 28"; $pnlTop.Controls.Add($btnRun)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = "Fill"; $grid.ReadOnly = $true; $grid.AutoSizeColumnsMode = "Fill"; 
    Apply-StandardGridTheme -Grid $grid -EnableAlternatingRowColor
    $subForm.Controls.Add($grid); $grid.BringToFront()

    $btnRun.Add_Click({
        $sel = $cmbQuery.SelectedIndex
        $list = @()
        if ($sel -eq 0) {
            $raw = Search-NativeLdap -LdapFilter "(&(objectCategory=computer)(userAccountControl:1.2.840.113556.1.4.803:=2))" -PropertiesToLoad @("name","distinguishedName")
            foreach ($i in $raw) { $list += [PSCustomObject]@{ "Name" = $i.Properties["name"][0]; "DN" = $i.Properties["distinguishedname"][0]; "Typ" = "Computer (Deaktiviert)" } }
        } elseif ($sel -eq 1) {
            $raw = Search-NativeLdap -LdapFilter "(&(objectCategory=person)(objectClass=user)(userAccountControl:1.2.840.113556.1.4.803:=65536))" -PropertiesToLoad @("name","sAMAccountName","distinguishedName")
            foreach ($i in $raw) { $list += [PSCustomObject]@{ "Name" = $i.Properties["name"][0]; "SAMAccount" = $i.Properties["samaccountname"][0]; "DN" = $i.Properties["distinguishedname"][0] } }
        } elseif ($sel -eq 2) {
            $raw = Search-NativeLdap -LdapFilter "(&(objectCategory=group)(!member=*))" -PropertiesToLoad @("name","distinguishedName")
            foreach ($i in $raw) { $list += [PSCustomObject]@{ "Gruppenname" = $i.Properties["name"][0]; "DN" = $i.Properties["distinguishedname"][0]; "Status" = "Keine Mitglieder" } }
        }
        $arr = [System.Collections.ArrayList]::new()
        foreach ($l in $list) { [void]$arr.Add($l) }
        $grid.DataSource = $arr
    })
    [void]$subForm.ShowDialog()
}
