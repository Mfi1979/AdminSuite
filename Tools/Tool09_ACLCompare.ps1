<#
================================================================================
 TOOL 9: AD UNIVERSAL ACL & BERECHTIGUNGSVERGLEICH (DIFF)
================================================================================
#>
function Show-Tool9-ACLCompare {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Tool 9: Active Directory Universal ACL & Berechtigungsvergleich"
    $form.Size = New-Object System.Drawing.Size(1280, 800)
    $form.StartPosition = "CenterScreen"
    $form.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9)
    $form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = "AD Objekt ACL-Vergleich (User, Gruppen, Computer, OUs & Berechtigungs-Diff)"
    $lblTitle.Font = New-Object System.Drawing.Font($form.Font.FontFamily, 12, [System.Drawing.FontStyle]::Bold)
    $lblTitle.Location = "15, 12"; $lblTitle.Size = "850, 25"; $form.Controls.Add($lblTitle)

    $grpInputs = New-Object System.Windows.Forms.GroupBox
    $grpInputs.Text = "Zu vergleichende AD-Objekte"; $grpInputs.Location = "15, 45"; $grpInputs.Size = "1230, 115"; $form.Controls.Add($grpInputs)

    $lblType = New-Object System.Windows.Forms.Label
    $lblType.Text = "Objekttyp:"; $lblType.Location = "15, 25"; $lblType.Size = "80, 20"; $grpInputs.Controls.Add($lblType)

    $cmbObjType = New-Object System.Windows.Forms.ComboBox
    $cmbObjType.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $cmbObjType.Items.AddRange(@("Auto-Erkennung (Alle Typen)", "Computer", "User / Benutzer", "Group / Gruppe", "OU / Container"))
    $cmbObjType.SelectedIndex = 0; $cmbObjType.Location = "95, 22"; $cmbObjType.Size = "185, 23"; $grpInputs.Controls.Add($cmbObjType)

    $lblComp1 = New-Object System.Windows.Forms.Label
    $lblComp1.Text = "Objekt 1 (Referenz):"; $lblComp1.Location = "15, 58"; $lblComp1.Size = "130, 20"; $grpInputs.Controls.Add($lblComp1)

    $txtComp1 = New-Object System.Windows.Forms.TextBox
    $txtComp1.Location = "145, 55"; $txtComp1.Size = "220, 23"; $grpInputs.Controls.Add($txtComp1)

    $lblComp2 = New-Object System.Windows.Forms.Label
    $lblComp2.Text = "Objekt 2 (Vergleich):"; $lblComp2.Location = "15, 86"; $lblComp2.Size = "130, 20"; $grpInputs.Controls.Add($lblComp2)

    $txtComp2 = New-Object System.Windows.Forms.TextBox
    $txtComp2.Location = "145, 83"; $txtComp2.Size = "220, 23"; $grpInputs.Controls.Add($txtComp2)

    $lblFilterStatus = New-Object System.Windows.Forms.Label
    $lblFilterStatus.Text = "Filter Vergleich:"; $lblFilterStatus.Location = "390, 25"; $lblFilterStatus.Size = "100, 20"; $grpInputs.Controls.Add($lblFilterStatus)

    $cmbFilterStatus = New-Object System.Windows.Forms.ComboBox
    $cmbFilterStatus.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $cmbFilterStatus.Items.AddRange(@("Alle Einträge", "Nur Abweichungen (Unterschiede)", "Nur Identisch (Match)", "Nur auf Objekt 1", "Nur auf Objekt 2"))
    $cmbFilterStatus.SelectedIndex = 0; $cmbFilterStatus.Location = "495, 22"; $cmbFilterStatus.Size = "220, 23"; $grpInputs.Controls.Add($cmbFilterStatus)

    $chkHideInherited = New-Object System.Windows.Forms.CheckBox
    $chkHideInherited.Text = "Vererbte Berechtigungen ausblenden (Nur explizite ACLs)"; $chkHideInherited.Location = "390, 58"; $chkHideInherited.Size = "350, 22"
    $grpInputs.Controls.Add($chkHideInherited)

    $btnCompare = New-Object System.Windows.Forms.Button
    $btnCompare.Text = "ACLs Vergleichen"; $btnCompare.Location = "760, 22"; $btnCompare.Size = "160, 80"
    $btnCompare.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215); $btnCompare.ForeColor = [System.Drawing.Color]::White
    $btnCompare.Font = New-Object System.Drawing.Font($form.Font.FontFamily, 10, [System.Drawing.FontStyle]::Bold)
    $grpInputs.Controls.Add($btnCompare)

    $btnExport = New-Object System.Windows.Forms.Button
    $btnExport.Text = "CSV Exportieren"; $btnExport.Location = "935, 22"; $btnExport.Size = "140, 80"
    $btnExport.BackColor = [System.Drawing.Color]::FromArgb(40, 167, 69); $btnExport.ForeColor = [System.Drawing.Color]::White
    $btnExport.Font = New-Object System.Drawing.Font($form.Font.FontFamily, 9.5, [System.Drawing.FontStyle]::Bold)
    $grpInputs.Controls.Add($btnExport)

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Text = "Geben Sie Namen, sAMAccountNames oder DNs zweier AD-Objekte ein."; $lblStatus.Location = "15, 165"; $lblStatus.Size = "1230, 24"
    $lblStatus.Font = New-Object System.Drawing.Font($form.Font.FontFamily, 9, [System.Drawing.FontStyle]::Italic)
    $form.Controls.Add($lblStatus)

    $gridACL = New-Object System.Windows.Forms.DataGridView
    $gridACL.Location = "15, 195"; $gridACL.Size = "1230, 550"
    $gridACL.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $gridACL.ReadOnly = $true; $gridACL.AllowUserToAddRows = $false
    $gridACL.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    Apply-StandardGridTheme -Grid $gridACL
    $form.Controls.Add($gridACL)

    $global:ComparisonResults = @()

    $gridACL.Add_RowPrePaint({
        param($sender, $e)
        if ($e.RowIndex -ge 0 -and $e.RowIndex -lt $gridACL.Rows.Count) {
            $row = $gridACL.Rows[$e.RowIndex]
            $statusVal = $row.Cells["Vergleichs-Status"].Value
            if ($statusVal -eq "Nur auf Objekt 1") {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 235, 235)
            } elseif ($statusVal -eq "Nur auf Objekt 2") {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(235, 245, 255)
            } elseif ($statusVal -like "*Abweichend*") {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 248, 220)
            }
        }
    })

    function Get-UniversalADObjectACL {
        param([string]$Identifier, [string]$TypeSelection)
        $target = $null

        if ($Identifier -match "^(CN|OU|DC)=") {
            try {
                $deDirect = [ADSI]"LDAP://$Identifier"
                if ($deDirect.distinguishedName) { $target = $deDirect }
            } catch {}
        }

        if (-not $target) {
            $baseFilter = switch ($TypeSelection) {
                "Computer"        { "(objectCategory=computer)" }
                "User / Benutzer" { "(&(objectCategory=person)(objectClass=user))" }
                "Group / Gruppe"  { "(objectCategory=group)" }
                "OU / Container"  { "(|(objectCategory=organizationalUnit)(objectCategory=container))" }
                default           { "(objectClass=*)" }
            }
            $searchFilter = "(&$baseFilter(|(sAMAccountName=$Identifier)(sAMAccountName=$Identifier`$)(name=$Identifier)))"
            $searcher = [adsisearcher]$searchFilter
            $searcher.PageSize = 5
            $searchResult = $searcher.FindOne()
            if ($searchResult) { $target = $searchResult.GetDirectoryEntry() }
        }

        if (-not $target) { return $null }

        $rules = @()
        try {
            foreach ($rule in $target.ObjectSecurity.GetAccessRules($true, $true, [System.Security.Principal.NTAccount])) {
                $rules += [PSCustomObject]@{
                    DistinguishedName = $target.distinguishedName
                    IdentityReference = $rule.IdentityReference.Value
                    Rights            = $rule.ActiveDirectoryRights.ToString()
                    AccessControlType = $rule.AccessControlType.ToString()
                    IsInherited       = $rule.IsInherited
                    InheritanceType   = $rule.InheritanceType.ToString()
                    Signature         = "$($rule.IdentityReference.Value)|$($rule.ActiveDirectoryRights)|$($rule.AccessControlType)"
                }
            }
        } catch { return $null }

        $objTypeStr = if ($target.schemaClassName) { $target.schemaClassName } else { "AD-Objekt" }
        return @{ "DN" = $target.distinguishedName; "Rules" = $rules; "Class" = $objTypeStr; "Name" = $target.name }
    }

    function Update-GridDisplay {
        if (-not $global:ComparisonResults) { return }
        $filtered = $global:ComparisonResults
        $mode = $cmbFilterStatus.SelectedItem.ToString()
        if ($mode -eq "Nur Abweichungen (Unterschiede)") {
            $filtered = $filtered | Where-Object { $_."Vergleichs-Status" -ne "Identisch (Match)" }
        } elseif ($mode -eq "Nur Identisch (Match)") {
            $filtered = $filtered | Where-Object { $_."Vergleichs-Status" -eq "Identisch (Match)" }
        } elseif ($mode -eq "Nur auf Objekt 1") {
            $filtered = $filtered | Where-Object { $_."Vergleichs-Status" -eq "Nur auf Objekt 1" }
        } elseif ($mode -eq "Nur auf Objekt 2") {
            $filtered = $filtered | Where-Object { $_."Vergleichs-Status" -eq "Nur auf Objekt 2" }
        }

        if ($chkHideInherited.Checked) {
            $filtered = $filtered | Where-Object { 
                $_."Objekt 1 Vererbt" -eq "Nein (Explizit)" -or $_."Objekt 2 Vererbt" -eq "Nein (Explizit)" 
            }
        }

        $arrList = [System.Collections.ArrayList]::new()
        foreach ($f in $filtered) { [void]$arrList.Add($f) }
        $gridACL.DataSource = $null
        $gridACL.DataSource = $arrList

        if ($gridACL.Columns["Vergleichs-Status"])  { $gridACL.Columns["Vergleichs-Status"].FillWeight  = 110 }
        if ($gridACL.Columns["Principal / Gruppe"]) { $gridACL.Columns["Principal / Gruppe"].FillWeight = 130 }
        if ($gridACL.Columns["Rechte (Objekt 1)"])  { $gridACL.Columns["Rechte (Objekt 1)"].FillWeight  = 160 }
        if ($gridACL.Columns["Rechte (Objekt 2)"])  { $gridACL.Columns["Rechte (Objekt 2)"].FillWeight  = 160 }
        if ($gridACL.Columns["Typ"])                { $gridACL.Columns["Typ"].FillWeight                = 70 }
        if ($gridACL.Columns["Objekt 1 Vererbt"])   { $gridACL.Columns["Objekt 1 Vererbt"].FillWeight   = 85 }
        if ($gridACL.Columns["Objekt 2 Vererbt"])   { $gridACL.Columns["Objekt 2 Vererbt"].FillWeight   = 85 }
    }

    $btnCompare.Add_Click({
        $c1 = $txtComp1.Text.Trim(); $c2 = $txtComp2.Text.Trim(); $selectedType = $cmbObjType.SelectedItem.ToString()
        if ([string]::IsNullOrWhiteSpace($c1) -or [string]::IsNullOrWhiteSpace($c2)) {
            [System.Windows.Forms.MessageBox]::Show("Bitte beide Objektnamen angeben!", "Eingabefehler", "OK", "Warning")
            return
        }

        $lblStatus.Text = "Lese ACLs für '$c1' und '$c2' aus Active Directory..."; $lblStatus.ForeColor = [System.Drawing.Color]::Black; $form.Refresh()

        try {
            $obj1Data = Get-UniversalADObjectACL -Identifier $c1 -TypeSelection $selectedType
            $obj2Data = Get-UniversalADObjectACL -Identifier $c2 -TypeSelection $selectedType

            if (-not $obj1Data) { $lblStatus.Text = "Fehler: Objekt '$c1' nicht gefunden!"; $lblStatus.ForeColor = [System.Drawing.Color]::Red; return }
            if (-not $obj2Data) { $lblStatus.Text = "Fehler: Objekt '$c2' nicht gefunden!"; $lblStatus.ForeColor = [System.Drawing.Color]::Red; return }

            $rules1 = $obj1Data.Rules; $rules2 = $obj2Data.Rules
            $comparisonList = [System.Collections.Generic.List[PSObject]]::new()
            $allIdentities = ($rules1.IdentityReference + $rules2.IdentityReference) | Select-Object -Unique | Sort-Object

            foreach ($ident in $allIdentities) {
                $r1List = $rules1 | Where-Object { $_.IdentityReference -eq $ident }
                $r2List = $rules2 | Where-Object { $_.IdentityReference -eq $ident }

                if ($r1List -and -not $r2List) {
                    foreach ($r in $r1List) {
                        $comparisonList.Add([PSCustomObject]@{
                            "Vergleichs-Status"  = "Nur auf Objekt 1"
                            "Principal / Gruppe" = $r.IdentityReference
                            "Rechte (Objekt 1)"  = $r.Rights
                            "Rechte (Objekt 2)"  = "-- Nicht vorhanden --"
                            "Typ"                = $r.AccessControlType
                            "Objekt 1 Vererbt"   = if ($r.IsInherited) { "Ja" } else { "Nein (Explizit)" }
                            "Objekt 2 Vererbt"   = "-"
                        })
                    }
                } elseif (-not $r1List -and $r2List) {
                    foreach ($r in $r2List) {
                        $comparisonList.Add([PSCustomObject]@{
                            "Vergleichs-Status"  = "Nur auf Objekt 2"
                            "Principal / Gruppe" = $r.IdentityReference
                            "Rechte (Objekt 1)"  = "-- Nicht vorhanden --"
                            "Rechte (Objekt 2)"  = $r.Rights
                            "Typ"                = $r.AccessControlType
                            "Objekt 1 Vererbt"   = "-"
                            "Objekt 2 Vererbt"   = if ($r.IsInherited) { "Ja" } else { "Nein (Explizit)" }
                        })
                    }
                } else {
                    $matchedSigs = @()
                    foreach ($r1 in $r1List) {
                        $matchingR2 = $r2List | Where-Object { $_.Signature -eq $r1.Signature } | Select-Object -First 1
                        if ($matchingR2) {
                            $matchedSigs += $r1.Signature
                            $isDiffInherit = ($r1.IsInherited -ne $matchingR2.IsInherited)
                            $status = if ($isDiffInherit) { "Abweichend (Vererbung)" } else { "Identisch (Match)" }
                            $comparisonList.Add([PSCustomObject]@{
                                "Vergleichs-Status"  = $status
                                "Principal / Gruppe" = $r1.IdentityReference
                                "Rechte (Objekt 1)"  = $r1.Rights
                                "Rechte (Objekt 2)"  = $matchingR2.Rights
                                "Typ"                = $r1.AccessControlType
                                "Objekt 1 Vererbt"   = if ($r1.IsInherited) { "Ja" } else { "Nein (Explizit)" }
                                "Objekt 2 Vererbt"   = if ($matchingR2.IsInherited) { "Ja" } else { "Nein (Explizit)" }
                            })
                        } else {
                            $comparisonList.Add([PSCustomObject]@{
                                "Vergleichs-Status"  = "Abweichende Rechte"
                                "Principal / Gruppe" = $r1.IdentityReference
                                "Rechte (Objekt 1)"  = $r1.Rights
                                "Rechte (Objekt 2)"  = ($r2List.Rights -join ", ")
                                "Typ"                = $r1.AccessControlType
                                "Objekt 1 Vererbt"   = if ($r1.IsInherited) { "Ja" } else { "Nein (Explizit)" }
                                "Objekt 2 Vererbt"   = "-"
                            })
                        }
                    }

                    foreach ($r2 in $r2List) {
                        if ($r2.Signature -notin $matchedSigs) {
                            $comparisonList.Add([PSCustomObject]@{
                                "Vergleichs-Status"  = "Abweichende Rechte"
                                "Principal / Gruppe" = $r2.IdentityReference
                                "Rechte (Objekt 1)"  = ($r1List.Rights -join ", ")
                                "Rechte (Objekt 2)"  = $r2.Rights
                                "Typ"                = $r2.AccessControlType
                                "Objekt 1 Vererbt"   = "-"
                                "Objekt 2 Vererbt"   = if ($r2.IsInherited) { "Ja" } else { "Nein (Explizit)" }
                            })
                        }
                    }
                }
            }

            $global:ComparisonResults = $comparisonList
            Update-GridDisplay
            $diffCount = ($comparisonList | Where-Object { $_."Vergleichs-Status" -ne "Identisch (Match)" }).Count
            $lblStatus.ForeColor = [System.Drawing.Color]::Black
            $lblStatus.Text = "Objekt 1: [$($obj1Data.Class)] $($obj1Data.DN)  |  Objekt 2: [$($obj2Data.Class)] $($obj2Data.DN)  |  Einträge: $($comparisonList.Count) (Abweichungen: $diffCount)"
        } catch {
            $lblStatus.ForeColor = [System.Drawing.Color]::Red
            $lblStatus.Text = "Fehler beim ACL-Vergleich: $($_.Exception.Message)"
        }
    })

    $cmbFilterStatus.Add_SelectedIndexChanged({ Update-GridDisplay })
    $chkHideInherited.Add_CheckedChanged({ Update-GridDisplay })

    $btnExport.Add_Click({
        if (-not $gridACL.DataSource -or $gridACL.Rows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Keine Daten zum Exportieren vorhanden!", "Export Info", "OK", "Information")
            return
        }
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "CSV-Datei (*.csv)|*.csv"
        $sfd.FileName = "AD_ACL_Vergleich_$($txtComp1.Text)_vs_$($txtComp2.Text)_$((Get-Date).ToString('yyyyMMdd_HHmm')).csv"
        if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $gridACL.DataSource | Export-Csv -Path $sfd.FileName -NoTypeInformation -Delimiter ";" -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show("ACL-Vergleich erfolgreich exportiert nach:`n$($sfd.FileName)", "Export Abgeschlossen", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    })

    [void]$form.ShowDialog()
}
