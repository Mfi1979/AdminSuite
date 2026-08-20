<#
================================================================================
 TOOL 4: GRUPPEN & GPO DIAGNOSTIK (3 REGISTER)
================================================================================
#>
function Open-ToolGroupsAndGPO {
    if (-not (Assert-DomainJoined)) { return }

    $subForm = New-Object System.Windows.Forms.Form
    $subForm.Text = "Tool 4: Gruppenmitgliedschaften & angewendete GPOs"
    $subForm.Size = New-Object System.Drawing.Size(1000, 700)
    $subForm.StartPosition = "CenterParent"
    $subForm.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9)

    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Dock = [System.Windows.Forms.DockStyle]::Fill
    $subForm.Controls.Add($tabControl)

    # Register 1: User-Gruppen
    $tabUserGroup = New-Object System.Windows.Forms.TabPage
    $tabUserGroup.Text = "User-Gruppen ($localUserName)"
    $tabUserGroup.Padding = New-Object System.Windows.Forms.Padding(5)

    $gridUserGroups = New-Object System.Windows.Forms.DataGridView
    $gridUserGroups.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridUserGroups.ReadOnly = $true
    $gridUserGroups.AllowUserToAddRows = $false
    $gridUserGroups.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
    $gridUserGroups.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    Apply-StandardGridTheme -Grid $gridUserGroups -EnableAlternatingRowColor
    $tabUserGroup.Controls.Add($gridUserGroups)
    $tabControl.TabPages.Add($tabUserGroup)

    # Register 2: PC-Gruppen
    $tabPCGroup = New-Object System.Windows.Forms.TabPage
    $tabPCGroup.Text = "PC-Gruppen ($localComputerName)"
    $tabPCGroup.Padding = New-Object System.Windows.Forms.Padding(5)

    $gridPCGroups = New-Object System.Windows.Forms.DataGridView
    $gridPCGroups.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridPCGroups.ReadOnly = $true
    $gridPCGroups.AllowUserToAddRows = $false
    $gridPCGroups.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
    $gridPCGroups.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    Apply-StandardGridTheme -Grid $gridPCGroups -EnableAlternatingRowColor
    $tabPCGroup.Controls.Add($gridPCGroups)
    $tabControl.TabPages.Add($tabPCGroup)

    # Register 3: GPOs
    $tabGPOs = New-Object System.Windows.Forms.TabPage
    $tabGPOs.Text = "Angewendete GPOs (gpresult)"
    $tabGPOs.Padding = New-Object System.Windows.Forms.Padding(5)

    $txtGPO = New-Object System.Windows.Forms.TextBox
    $txtGPO.Dock = [System.Windows.Forms.DockStyle]::Fill
    $txtGPO.Multiline = $true
    $txtGPO.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
    $txtGPO.Font = New-Object System.Drawing.Font("Consolas", 9.5)
    $txtGPO.ReadOnly = $true
    $txtGPO.BackColor = [System.Drawing.Color]::White
    $tabGPOs.Controls.Add($txtGPO)
    $tabControl.TabPages.Add($tabGPOs)

    $subForm.Add_Shown({
        try {
            $uRes = Search-NativeLdap -LdapFilter "(&(objectCategory=person)(objectClass=user)(sAMAccountName=$localUserName))" -PropertiesToLoad @("memberOf")
            $uGroups = @()
            if ($uRes -and $uRes[0].Properties["memberof"]) {
                foreach ($dn in $uRes[0].Properties["memberof"]) {
                    $cn = ($dn -split ',')[0] -replace '^CN=', ''
                    $uGroups += [PSCustomObject]@{ "Gruppenname" = $cn; "DistinguishedName" = $dn }
                }
            }
            $arrU = [System.Collections.ArrayList]::new()
            foreach ($g in ($uGroups | Sort-Object Gruppenname)) { [void]$arrU.Add($g) }
            $gridUserGroups.DataSource = $arrU
        } catch { }

        try {
            $cRes = Search-NativeLdap -LdapFilter "(&(objectCategory=computer)(name=$localComputerName))" -PropertiesToLoad @("memberOf")
            $cGroups = @()
            if ($cRes -and $cRes[0].Properties["memberof"]) {
                foreach ($dn in $cRes[0].Properties["memberof"]) {
                    $cn = ($dn -split ',')[0] -replace '^CN=', ''
                    $cGroups += [PSCustomObject]@{ "Gruppenname" = $cn; "DistinguishedName" = $dn }
                }
            }
            $arrC = [System.Collections.ArrayList]::new()
            foreach ($g in ($cGroups | Sort-Object Gruppenname)) { [void]$arrC.Add($g) }
            $gridPCGroups.DataSource = $arrC
        } catch { }

        $txtGPO.Text = "Lade Gruppenrichtlinien via gpresult /r ..."
        $subForm.Refresh()
        try {
            $gpoOutput = gpresult /r 2>&1 | Out-String
            $txtGPO.Text = $gpoOutput
        } catch {
            $txtGPO.Text = "Fehler beim Ausführen von gpresult /r.`r`n$($_.Exception.Message)"
        }
    })

    [void]$subForm.ShowDialog()
}
