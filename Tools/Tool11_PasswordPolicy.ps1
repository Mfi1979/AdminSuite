<#
================================================================================
 TOOL 11: AD KENNWORT-RICHTLINIEN & PSO AUDIT
================================================================================
#>

function Open-ToolPasswordPolicies {
    if (-not (Assert-DomainJoined)) { return }

    $t = $script:I18N[$script:CurrentLang]

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $t["Tool11Title"]
    $form.Size = New-Object System.Drawing.Size(1150, 720)
    $form.StartPosition = "CenterParent"
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.BackColor = $script:UITheme.BackColorMain

    # Header Panel
    $headerPanel = New-Object System.Windows.Forms.Panel
    $headerPanel.Dock = [System.Windows.Forms.DockStyle]::Top
    $headerPanel.Height = 60
    $headerPanel.BackColor = $script:UITheme.BackColorCard
    $form.Controls.Add($headerPanel)

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = $t["Tool11Title"]
    $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)
    $lblTitle.ForeColor = $script:UITheme.TextColorPrimary
    $lblTitle.Location = New-Object System.Drawing.Point(20, 8)
    $lblTitle.AutoSize = $true
    $headerPanel.Controls.Add($lblTitle)

    $lblSubTitle = New-Object System.Windows.Forms.Label
    $lblSubTitle.Text = $t["Tool11Sub"]
    $lblSubTitle.ForeColor = $script:UITheme.TextColorSecondary
    $lblSubTitle.Location = New-Object System.Drawing.Point(22, 33)
    $lblSubTitle.AutoSize = $true
    $headerPanel.Controls.Add($lblSubTitle)

    # Bottom Panel
    $bottomPanel = New-Object System.Windows.Forms.Panel
    $bottomPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $bottomPanel.Height = 50
    $bottomPanel.BackColor = $script:UITheme.BackColorCard
    $form.Controls.Add($bottomPanel)

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = New-Object System.Drawing.Point(20, 15)
    $lblStatus.AutoSize = $true
    $lblStatus.ForeColor = $script:UITheme.TextColorSecondary
    $lblStatus.Text = $t["StatusReady"]
    $bottomPanel.Controls.Add($lblStatus)

    $btnExport = New-Object System.Windows.Forms.Button
    $btnExport.Text = $t["BtnExportCsv"]
    $btnExport.Size = New-Object System.Drawing.Size(130, 30)
    $btnExport.Location = New-Object System.Drawing.Point(830, 10)
    $btnExport.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
    $btnExport.BackColor = $script:UITheme.AccentColor
    $btnExport.ForeColor = [System.Drawing.Color]::White
    $btnExport.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnExport.FlatAppearance.BorderSize = 0
    $bottomPanel.Controls.Add($btnExport)

    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = $t["BtnRefresh"]
    $btnRefresh.Size = New-Object System.Drawing.Size(130, 30)
    $btnRefresh.Location = New-Object System.Drawing.Point(970, 10)
    $btnRefresh.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
    $btnRefresh.BackColor = $script:UITheme.TextColorSecondary
    $btnRefresh.ForeColor = [System.Drawing.Color]::White
    $btnRefresh.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnRefresh.FlatAppearance.BorderSize = 0
    $bottomPanel.Controls.Add($btnRefresh)

    # Tab Control
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Dock = [System.Windows.Forms.DockStyle]::Fill
    $tabControl.Padding = New-Object System.Drawing.Point(12, 6)
    $form.Controls.Add($tabControl)
    $tabControl.BringToFront()

    $applyGridStyle = {
        param($g)
        $g.Dock = [System.Windows.Forms.DockStyle]::Fill
        $g.BackgroundColor = $script:UITheme.BackColorCard
        $g.BorderStyle = [System.Windows.Forms.BorderStyle]::None
        $g.ReadOnly = $true
        $g.AllowUserToAddRows = $false
        $g.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
        $g.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
        $g.RowHeadersVisible = $false
        $g.EnableHeadersVisualStyles = $false
        $g.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(241, 245, 249)
        $g.ColumnHeadersDefaultCellStyle.ForeColor = $script:UITheme.TextColorPrimary
        $g.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $g.ColumnHeadersHeight = $script:UITheme.HeaderHeight
        $g.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 252)
    }

    # Tab 1: Default Domain Policy
    $tabDomain = New-Object System.Windows.Forms.TabPage
    $tabDomain.Text = $t["TabDefaultPolicy"]
    $tabDomain.BackColor = $script:UITheme.BackColorCard
    $tabControl.TabPages.Add($tabDomain)

    $gridDomain = New-Object System.Windows.Forms.DataGridView
    & $applyGridStyle $gridDomain
    $tabDomain.Controls.Add($gridDomain)

    # Tab 2: Fine-Grained PSOs
    $tabPSO = New-Object System.Windows.Forms.TabPage
    $tabPSO.Text = $t["TabFineGrained"]
    $tabPSO.BackColor = $script:UITheme.BackColorCard
    $tabControl.TabPages.Add($tabPSO)

    $gridPSO = New-Object System.Windows.Forms.DataGridView
    & $applyGridStyle $gridPSO
    $tabPSO.Controls.Add($gridPSO)

    # Tab 3: User Effective Policy Check
    $tabUser = New-Object System.Windows.Forms.TabPage
    $tabUser.Text = $t["TabUserCheck"]
    $tabUser.BackColor = $script:UITheme.BackColorCard
    $tabControl.TabPages.Add($tabUser)

    $userTopPanel = New-Object System.Windows.Forms.Panel
    $userTopPanel.Dock = [System.Windows.Forms.DockStyle]::Top
    $userTopPanel.Height = 55
    $userTopPanel.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 252)
    $tabUser.Controls.Add($userTopPanel)

    $lblUser = New-Object System.Windows.Forms.Label
    $lblUser.Text = $t["UserSearchLabel"]
    $lblUser.Location = New-Object System.Drawing.Point(15, 18)
    $lblUser.AutoSize = $true
    $lblUser.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $userTopPanel.Controls.Add($lblUser)

    $txtUserCheck = New-Object System.Windows.Forms.TextBox
    $txtUserCheck.Location = New-Object System.Drawing.Point(320, 14)
    $txtUserCheck.Size = New-Object System.Drawing.Size(200, 25)
    $txtUserCheck.Text = $env:USERNAME
    $userTopPanel.Controls.Add($txtUserCheck)

    $btnCheckUser = New-Object System.Windows.Forms.Button
    $btnCheckUser.Text = $t["BtnCheckUser"]
    $btnCheckUser.Location = New-Object System.Drawing.Point(530, 13)
    $btnCheckUser.Size = New-Object System.Drawing.Size(150, 27)
    $btnCheckUser.BackColor = [System.Drawing.Color]::FromArgb(15, 118, 110)
    $btnCheckUser.ForeColor = [System.Drawing.Color]::White
    $btnCheckUser.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnCheckUser.FlatAppearance.BorderSize = 0
    $userTopPanel.Controls.Add($btnCheckUser)

    $gridUserPolicy = New-Object System.Windows.Forms.DataGridView
    & $applyGridStyle $gridUserPolicy
    $tabUser.Controls.Add($gridUserPolicy)
    $gridUserPolicy.BringToFront()

    # Robuster TimeSpan Konverter
    $convertTimeSpan = {
        param($val, $asDays = $false)
        if ($null -eq $val) { return if ($script:CurrentLang -eq "DE") { "Nicht konfiguriert" } else { "Not Configured" } }
        try {
            $ticks = 0
            if ($val -is [int64] -or $val -is [int] -or $val -is [double]) {
                $ticks = [int64]$val
            } elseif ($val.GetType().Name -match "ComObject|__ComObject|LargeInteger") {
                $high = [int64]$val.GetType().InvokeMember("HighPart", [System.Reflection.BindingFlags]::GetProperty, $null, $val, $null)
                $low  = [int64]$val.GetType().InvokeMember("LowPart", [System.Reflection.BindingFlags]::GetProperty, $null, $val, $null)
                $ticks = ($high -shl 32) + ($low -band 0xFFFFFFFF)
            } else {
                $ticks = [int64]::Parse($val.ToString())
            }

            if ($ticks -eq 0 -or $ticks -eq [Int64]::MinValue) {
                return if ($script:CurrentLang -eq "DE") { "Nie ablaufend / Deaktiviert" } else { "Never Expires / Disabled" }
            }

            if ($ticks -lt 0) { $ticks = -$ticks }
            $ts = [TimeSpan]::FromTicks($ticks)

            if ($asDays) {
                $unit = if ($script:CurrentLang -eq "DE") { "Tage" } else { "Days" }
                return "{0:N1} $unit" -f $ts.TotalDays
            } else {
                if ($ts.TotalMinutes -ge 60) {
                    $unit = if ($script:CurrentLang -eq "DE") { "Stunden" } else { "Hours" }
                    return "{0:N1} $unit" -f $ts.TotalHours
                } else {
                    $unit = if ($script:CurrentLang -eq "DE") { "Minuten" } else { "Minutes" }
                    return "{0:N0} $unit" -f $ts.TotalMinutes
                }
            }
        } catch {
            return $val.ToString()
        }
    }

    $script:cachedDomainPolicies = @()
    $script:cachedPSOs = @()
    $script:cachedUserPolicy = @()

    # Laderoutine
    $loadPolicies = {
        $lblStatus.Text = $t["StatusSearching"]
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $script:cachedDomainPolicies = @()
        $script:cachedPSOs = @()

        try {
            $rootEntry = [System.DirectoryServices.DirectoryEntry]::new("LDAP://RootDSE")
            $defNamingContext = $rootEntry.defaultNamingContext.ToString()
            $domEntry = [System.DirectoryServices.DirectoryEntry]::new("LDAP://$defNamingContext")

            $minPwdAge      = & $convertTimeSpan $domEntry.Properties["minPwdAge"].Value $true
            $maxPwdAge      = & $convertTimeSpan $domEntry.Properties["maxPwdAge"].Value $true
            $minPwdLength   = if ($domEntry.Properties["minPwdLength"].Value) { "$($domEntry.Properties['minPwdLength'].Value) " + (if ($script:CurrentLang -eq "DE") { "Zeichen" } else { "Characters" }) } else { "0" }
            $pwdHistory     = if ($domEntry.Properties["pwdHistoryLength"].Value) { "$($domEntry.Properties['pwdHistoryLength'].Value) " + (if ($script:CurrentLang -eq "DE") { "Kennwoerter gespeichert" } else { "Passwords remembered" }) } else { "0" }
            
            $pwdProps = [int]($domEntry.Properties["pwdProperties"].Value)
            $complexityEnabled = ($pwdProps -band 1) -eq 1
            $reversibleEnc     = ($pwdProps -band 16) -eq 1

            $lockThreshold = if ($domEntry.Properties["lockoutThreshold"].Value) { "$($domEntry.Properties['lockoutThreshold'].Value) " + (if ($script:CurrentLang -eq "DE") { "ungueltige Versuche" } else { "Invalid attempts" }) } else { if ($script:CurrentLang -eq "DE") { "0 (Deaktiviert)" } else { "0 (Disabled)" } }
            $lockDuration  = & $convertTimeSpan $domEntry.Properties["lockoutDuration"].Value $false
            $lockWindow    = & $convertTimeSpan $domEntry.Properties["lockOutObservationWindow"].Value $false

            $domList = @(
                [PSCustomObject]@{ $t["PropCategory"] = $t["CatPassword"]; $t["PropName"] = $t["PropMinAge"]; $t["PropValue"] = $minPwdAge },
                [PSCustomObject]@{ $t["PropCategory"] = $t["CatPassword"]; $t["PropName"] = $t["PropMaxAge"]; $t["PropValue"] = $maxPwdAge },
                [PSCustomObject]@{ $t["PropCategory"] = $t["CatPassword"]; $t["PropName"] = $t["PropMinLength"]; $t["PropValue"] = $minPwdLength },
                [PSCustomObject]@{ $t["PropCategory"] = $t["CatPassword"]; $t["PropName"] = $t["PropHistory"]; $t["PropValue"] = $pwdHistory },
                [PSCustomObject]@{ $t["PropCategory"] = $t["CatPassword"]; $t["PropName"] = $t["PropComplexity"]; $t["PropValue"] = if ($complexityEnabled) { $t["ValActive"] } else { $t["ValDisabled"] } },
                [PSCustomObject]@{ $t["PropCategory"] = $t["CatPassword"]; $t["PropName"] = $t["PropReversible"]; $t["PropValue"] = if ($reversibleEnc) { $t["ValActiveInsecure"] } else { $t["ValDisabledSecure"] } },
                [PSCustomObject]@{ $t["PropCategory"] = $t["CatLockout"]; $t["PropName"] = $t["PropThreshold"]; $t["PropValue"] = $lockThreshold },
                [PSCustomObject]@{ $t["PropCategory"] = $t["CatLockout"]; $t["PropName"] = $t["PropDuration"]; $t["PropValue"] = $lockDuration },
                [PSCustomObject]@{ $t["PropCategory"] = $t["CatLockout"]; $t["PropName"] = $t["PropObservation"]; $t["PropValue"] = $lockWindow }
            )

            $script:cachedDomainPolicies = $domList
            $arrD = [System.Collections.ArrayList]::new()
            foreach ($row in $domList) { [void]$arrD.Add($row) }
            $gridDomain.DataSource = $arrD

            # Fine-Grained PSOs
            $psoList = @()
            $psoContainerPath = "LDAP://CN=Password Settings Objects,CN=System,$defNamingContext"
            $psoResults = $null
            
            try {
                $psoEntry = [System.DirectoryServices.DirectoryEntry]::new($psoContainerPath)
                $psoSearcher = [System.DirectoryServices.DirectorySearcher]::new($psoEntry)
                $psoSearcher.Filter = "(objectClass=msDS-PasswordSettings)"
                $psoSearcher.SearchScope = [System.DirectoryServices.SearchScope]::OneLevel
                $psoResults = $psoSearcher.FindAll()
            } catch {
                $psoSearcher = [System.DirectoryServices.DirectorySearcher]::new($domEntry)
                $psoSearcher.Filter = "(objectClass=msDS-PasswordSettings)"
                $psoSearcher.SearchScope = [System.DirectoryServices.SearchScope]::Subtree
                $psoResults = $psoSearcher.FindAll()
            }

            if ($psoResults -and $psoResults.Count -gt 0) {
                foreach ($res in $psoResults) {
                    $prop = $res.Properties
                    $name = if ($prop["name"].Count -gt 0) { $prop["name"][0] } else { "Unbenannt" }
                    $prec = if ($prop["msds-passwordsettingsprecedence"].Count -gt 0) { $prop["msds-passwordsettingsprecedence"][0] } else { "N/A" }
                    $pMinLen = if ($prop["msds-minimumpasswordlength"].Count -gt 0) { "$($prop['msds-minimumpasswordlength'][0]) " + (if ($script:CurrentLang -eq "DE") { "Zeichen" } else { "Characters" }) } else { "N/A" }
                    $pHist = if ($prop["msds-passwordhistorylength"].Count -gt 0) { "$($prop['msds-passwordhistorylength'][0])" } else { "0" }
                    $pCompl = if ($prop["msds-passwordcomplexityenabled"].Count -gt 0) { [bool]$prop["msds-passwordcomplexityenabled"][0] } else { $false }
                    
                    $pMinAge = if ($prop["msds-minpasswordage"].Count -gt 0) { & $convertTimeSpan $prop["msds-minpasswordage"][0] $true } else { "N/A" }
                    $pMaxAge = if ($prop["msds-maxpasswordage"].Count -gt 0) { & $convertTimeSpan $prop["msds-maxpasswordage"][0] $true } else { "N/A" }
                    $pLockThresh = if ($prop["msds-lockoutthreshold"].Count -gt 0) { "$($prop['msds-lockoutthreshold'][0]) " + (if ($script:CurrentLang -eq "DE") { "Versuche" } else { "Attempts" }) } else { "0" }
                    $pLockDur = if ($prop["msds-lockoutduration"].Count -gt 0) { & $convertTimeSpan $prop["msds-lockoutduration"][0] $false } else { "N/A" }

                    $appliesTo = @()
                    if ($prop["msds-psoappliesto"].Count -gt 0) {
                        foreach ($appDn in $prop["msds-psoappliesto"]) {
                            if ($appDn -match "CN=([^,]+)") { $appliesTo += $Matches[1] } else { $appliesTo += $appDn }
                        }
                    }
                    $appliesToStr = if ($appliesTo.Count -gt 0) { $appliesTo -join ", " } else { if ($script:CurrentLang -eq "DE") { "Niemand (Keine Zuweisung)" } else { "None (No assignment)" } }

                    $psoList += [PSCustomObject]@{
                        $t["PsoName"]       = $name
                        $t["PsoPrecedence"] = $prec
                        $t["PsoMinLength"]  = $pMinLen
                        $t["PsoHistory"]    = $pHist
                        $t["PsoComplexity"] = if ($pCompl) { $t["ValActive"] } else { $t["ValDisabled"] }
                        $t["PsoMinAge"]     = $pMinAge
                        $t["PsoMaxAge"]     = $pMaxAge
                        $t["PsoThreshold"]  = $pLockThresh
                        $t["PsoDuration"]   = $pLockDur
                        $t["PsoAppliesTo"]  = $appliesToStr
                    }
                }
            }

            $script:cachedPSOs = $psoList
            $arrPSO = [System.Collections.ArrayList]::new()
            foreach ($p in $psoList) { [void]$arrPSO.Add($p) }
            $gridPSO.DataSource = $arrPSO

            $psoCount = $psoList.Count
            $psoMsg = if ($script:CurrentLang -eq "DE") { "$psoCount Fine-Grained PSO(s) gefunden." } else { "$psoCount Fine-Grained PSO(s) found." }
            $lblStatus.Text = "$($t['StatusResult']) $psoMsg"

        } catch {
            $lblStatus.Text = "$($t['StatusError']): $($_.Exception.Message)"
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }

    # Benutzer-Check Logik
    $checkUserAction = {
        $uName = $txtUserCheck.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($uName)) {
            [System.Windows.Forms.MessageBox]::Show($t["ErrUserEmpty"], "Info", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        $lblStatus.Text = $t["StatusSearching"]
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

        try {
            $rootEntry = [System.DirectoryServices.DirectoryEntry]::new("LDAP://RootDSE")
            $defNamingContext = $rootEntry.defaultNamingContext.ToString()
            $domEntry = [System.DirectoryServices.DirectoryEntry]::new("LDAP://$defNamingContext")
            
            $uSearcher = [System.DirectoryServices.DirectorySearcher]::new($domEntry)
            $uSearcher.Filter = "(&(objectClass=user)(objectCategory=person)(|(sAMAccountName=$uName)(userPrincipalName=$uName)))"
            $uSearcher.PropertiesToLoad.AddRange(@("distinguishedName", "sAMAccountName", "userPrincipalName", "msDS-ResultantPSO"))
            $uRes = $uSearcher.FindOne()

            if (-not $uRes) {
                [System.Windows.Forms.MessageBox]::Show($t["ErrUserNotFound"], "AD Search", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                $lblStatus.Text = $t["ErrUserNotFound"]
                return
            }

            $uProps = $uRes.Properties
            $foundSam = $uProps["samaccountname"][0]
            $resultantPSO = if ($uProps["msds-resultantpso"].Count -gt 0) { $uProps["msds-resultantpso"][0] } else { $null }

            $userResultList = @()
            if ($resultantPSO) {
                $psoEntry = [System.DirectoryServices.DirectoryEntry]::new("LDAP://$resultantPSO")
                $psoName = $psoEntry.Properties["name"].Value
                $pMinLen = "$($psoEntry.Properties['msDS-MinimumPasswordLength'].Value) " + (if ($script:CurrentLang -eq "DE") { "Zeichen" } else { "Characters" })
                $pMinAge = & $convertTimeSpan $psoEntry.Properties["msDS-MinPasswordAge"].Value $true
                $pMaxAge = & $convertTimeSpan $psoEntry.Properties["msDS-MaxPasswordAge"].Value $true
                $pCompl  = if ([bool]$psoEntry.Properties["msDS-PasswordComplexityEnabled"].Value) { $t["ValActive"] } else { $t["ValDisabled"] }
                $pThresh = "$($psoEntry.Properties['msDS-LockoutThreshold'].Value) " + (if ($script:CurrentLang -eq "DE") { "Versuche" } else { "Attempts" })
                $pDur    = & $convertTimeSpan $psoEntry.Properties["msDS-LockoutDuration"].Value $false

                $userResultList = @(
                    [PSCustomObject]@{ $t["PropCategory"] = $t["CatAppliedPolicy"]; $t["PropName"] = $t["PropSource"]; $t["PropValue"] = "Fine-Grained PSO ($psoName)" },
                    [PSCustomObject]@{ $t["PropCategory"] = $t["CatAppliedPolicy"]; $t["PropName"] = $t["PropDN"]; $t["PropValue"] = $resultantPSO },
                    [PSCustomObject]@{ $t["PropCategory"] = $t["CatPassword"]; $t["PropName"] = $t["PropMinLength"]; $t["PropValue"] = $pMinLen },
                    [PSCustomObject]@{ $t["PropCategory"] = $t["CatPassword"]; $t["PropName"] = $t["PropMinAge"]; $t["PropValue"] = $pMinAge },
                    [PSCustomObject]@{ $t["PropCategory"] = $t["CatPassword"]; $t["PropName"] = $t["PropMaxAge"]; $t["PropValue"] = $pMaxAge },
                    [PSCustomObject]@{ $t["PropCategory"] = $t["CatPassword"]; $t["PropName"] = $t["PropComplexity"]; $t["PropValue"] = $pCompl },
                    [PSCustomObject]@{ $t["PropCategory"] = $t["CatLockout"]; $t["PropName"] = $t["PropThreshold"]; $t["PropValue"] = $pThresh },
                    [PSCustomObject]@{ $t["PropCategory"] = $t["CatLockout"]; $t["PropName"] = $t["PropDuration"]; $t["PropValue"] = $pDur }
                )
            } else {
                $userResultList = @(
                    [PSCustomObject]@{ $t["PropCategory"] = $t["CatAppliedPolicy"]; $t["PropName"] = $t["PropSource"]; $t["PropValue"] = $t["ValDefaultPolicy"] }
                )
                foreach ($row in $script:cachedDomainPolicies) {
                    $userResultList += $row
                }
            }

            $script:cachedUserPolicy = $userResultList
            $arrU = [System.Collections.ArrayList]::new()
            foreach ($el in $userResultList) { [void]$arrU.Add($el) }
            $gridUserPolicy.DataSource = $arrU
            $lblStatus.Text = "$($t['StatusResultReady']): $foundSam"

        } catch {
            $lblStatus.Text = "$($t['StatusError']): $($_.Exception.Message)"
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }

    $btnCheckUser.Add_Click($checkUserAction)
    $txtUserCheck.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $_.SuppressKeyPress = $true
            & $checkUserAction
        }
    })

    # Export CSV
    $btnExport.Add_Click({
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "CSV-Datei (*.csv)|*.csv"
        $sfd.FileName = "AD_Password_Policies_Audit.csv"
        if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $exportData = @()
            foreach ($row in $script:cachedDomainPolicies) {
                $exportData += [PSCustomObject]@{
                    "Typ"         = "Default Domain Policy"
                    "Eigenschaft" = $row.($t["PropName"])
                    "Wert"        = $row.($t["PropValue"])
                }
            }
            foreach ($row in $script:cachedPSOs) {
                $exportData += [PSCustomObject]@{
                    "Typ"         = "Fine-Grained PSO ($($row.($t['PsoName'])))"
                    "Eigenschaft" = "Precedence: $($row.($t['PsoPrecedence'])), MinLen: $($row.($t['PsoMinLength']))"
                    "Wert"        = "AppliesTo: $($row.($t['PsoAppliesTo']))"
                }
            }
            $exportData | Export-Csv -Path $sfd.FileName -NoTypeInformation -Delimiter ";" -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show("Export abgeschlossen:`n$($sfd.FileName)", "Export OK", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    })

    $btnRefresh.Add_Click({ & $loadPolicies })
    $form.Add_Shown({ & $loadPolicies })

    [void]$form.ShowDialog()
}