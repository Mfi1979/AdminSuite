<#
================================================================================
 TOOL 11: AD PASSWORD POLICIES & FINE-GRAINED PSO AUDIT (BEWÄHRTE FASSUNG)
================================================================================
#>
function Show-Tool11-PasswordPolicyAudit {
    if (-not (Assert-DomainJoined)) { return }

    $lang = if ($script:CurrentLang -eq "EN") { "EN" } else { "DE" }
    $t = @{
        "DE" = @{
            "Title"           = "Tool 11: AD Kennwortrichtlinien & PSO Audit"
            "TabDomain"       = "1. Standard-Domänenrichtlinie"
            "TabPSO"          = "2. Fine-Grained Password Policies (PSO)"
            "TabUser"         = "3. Effektive Benutzer-Richtlinie"
            "PropName"        = "Eigenschaft / Richtlinie"
            "PropValue"       = "Konfigurierter Wert"
            "PsoName"         = "PSO Name"
            "PsoPrecedence"   = "Priorität (Precedence)"
            "PsoMinLength"    = "Min. Länge"
            "PsoHistory"      = "Historie"
            "PsoComplexity"   = "Komplexität"
            "PsoMaxAge"       = "Max. Alter (Tage)"
            "PsoMinAge"       = "Min. Alter (Tage)"
            "PsoLockoutThresh"= "Sperrschwelle"
            "PsoLockoutDur"   = "Sperrdauer (Min.)"
            "PsoAppliesTo"    = "Zugewiesen an (Applies To)"
            "BtnLoad"         = "Richtlinien laden / aktualisieren"
            "BtnExport"       = "Exportieren (CSV)"
            "LblUserSearch"   = "Benutzername (sAMAccountName):"
            "BtnCheckUser"    = "Effektive Richtlinie prüfen"
            "StatusReady"     = "Bereit zur Abfrage."
            "StatusLoading"   = "Lese Active Directory Kennwortrichtlinien aus..."
            "StatusDone"      = "Kennwortrichtlinien erfolgreich geladen."
            "ErrNoDomain"     = "Keine Active Directory Domäne erreichbar."
            "ErrUserNotFound" = "Benutzer nicht im Active Directory gefunden."
            "UserResultTitle" = "Ergebnis für Benutzer"
            "InheritedDomain" = "Standard-Domänenrichtlinie (Keine spezifische PSO)"
            "Days"            = "Tage"
            "Minutes"         = "Minuten"
            "Characters"      = "Zeichen"
            "Passwords"       = "Passwörter"
            "Attempts"        = "Versuche"
            "InvalidAttempts" = "ungültige Versuche"
            "NoLockout"       = "Keine Sperre (0)"
            "NeverExpires"    = "Nie ablaufend / Deaktiviert"
            "NoneAssigned"    = "Niemand (Keine Zuweisung)"
            "Enabled"         = "Aktiviert"
            "Disabled"        = "Deaktiviert"
            "Safe"            = "Sicher"
            "Unsafe"          = "Unsicher"
            "Yes"             = "Ja"
            "No"              = "Nein"
            "DPO_MinLength"   = "Minimale Kennwortlänge"
            "DPO_Complexity"  = "Kennwortkomplexität"
            "DPO_MaxAge"      = "Maximales Kennwortalter"
            "DPO_MinAge"      = "Minimales Kennwortalter"
            "DPO_History"     = "Kennworthistorie / Verlauf"
            "DPO_LockoutTh"   = "Kontosperrungsschwelle"
            "DPO_LockoutDur"  = "Kontosperrdauer"
            "DPO_ObsWindow"   = "Sperrungsbeobachtungsfenster"
            "DPO_RevEncrypt"  = "Umkehrbare Verschlüsselung"
            "Eff_Policy"      = "Gültige Richtlinie"
            "Eff_PSO_DN"      = "PSO DistinguishedName"
            "Eff_LockoutTh"   = "Kontosperrung nach Fehlversuchen"
        }
        "EN" = @{
            "Title"           = "Tool 11: AD Password Policies & PSO Audit"
            "TabDomain"       = "1. Default Domain Policy"
            "TabPSO"          = "2. Fine-Grained Password Policies (PSO)"
            "TabUser"         = "3. Effective User Password Policy"
            "PropName"        = "Property / Policy Setting"
            "PropValue"       = "Configured Value"
            "PsoName"         = "PSO Name"
            "PsoPrecedence"   = "Precedence"
            "PsoMinLength"    = "Min Length"
            "PsoHistory"      = "History"
            "PsoComplexity"   = "Complexity"
            "PsoMaxAge"       = "Max Age (Days)"
            "PsoMinAge"       = "Min Age (Days)"
            "PsoLockoutThresh"= "Lockout Threshold"
            "PsoLockoutDur"   = "Lockout Duration (Min)"
            "PsoAppliesTo"    = "Applies To"
            "BtnLoad"         = "Load / Refresh Policies"
            "BtnExport"       = "Export (CSV)"
            "LblUserSearch"   = "Username (sAMAccountName):"
            "BtnCheckUser"    = "Check Effective Policy"
            "StatusReady"     = "Ready for query."
            "StatusLoading"   = "Querying Active Directory Password Policies..."
            "StatusDone"      = "Password policies loaded successfully."
            "ErrNoDomain"     = "No Active Directory Domain accessible."
            "ErrUserNotFound" = "User not found in Active Directory."
            "UserResultTitle" = "Result for user"
            "InheritedDomain" = "Default Domain Policy (No specific PSO applied)"
            "Days"            = "Days"
            "Minutes"         = "Minutes"
            "Characters"      = "Characters"
            "Passwords"       = "Passwords"
            "Attempts"        = "Attempts"
            "InvalidAttempts" = "invalid attempts"
            "NoLockout"       = "No Lockout (0)"
            "NeverExpires"    = "Never Expires / Disabled"
            "NoneAssigned"    = "None (No assignment)"
            "Enabled"         = "Enabled"
            "Disabled"        = "Disabled"
            "Safe"            = "Secure"
            "Unsafe"          = "Insecure"
            "Yes"             = "Yes"
            "No"              = "No"
            "DPO_MinLength"   = "Minimum Password Length"
            "DPO_Complexity"  = "Password Complexity"
            "DPO_MaxAge"      = "Maximum Password Age"
            "DPO_MinAge"      = "Minimum Password Age"
            "DPO_History"     = "Password History Length"
            "DPO_LockoutTh"   = "Account Lockout Threshold"
            "DPO_LockoutDur"  = "Account Lockout Duration"
            "DPO_ObsWindow"   = "Lockout Observation Window"
            "DPO_RevEncrypt"  = "Store Passwords Using Reversible Encryption"
            "Eff_Policy"      = "Effective Policy"
            "Eff_PSO_DN"      = "PSO DistinguishedName"
            "Eff_LockoutTh"   = "Account Lockout Threshold"
        }
    }[$lang]

    function Convert-LargeIntToTimeSpan([object]$largeIntObj) {
        if (-not $largeIntObj) { return $null }
        try {
            $highPart = $largeIntObj.GetType().InvokeMember("HighPart", [System.Reflection.BindingFlags]::GetProperty, $null, $largeIntObj, $null)
            $lowPart  = $largeIntObj.GetType().InvokeMember("LowPart", [System.Reflection.BindingFlags]::GetProperty, $null, $largeIntObj, $null)
            $raw64 = ([int64]$highPart -shl 32) -bor ([int64]$lowPart -band 0xFFFFFFFF)
            if ($raw64 -lt 0) {
                return [timespan]::FromTicks(-$raw64)
            } else {
                return [timespan]::FromTicks($raw64)
            }
        } catch {
            return $null
        }
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $t["Title"]
    $form.Size = New-Object System.Drawing.Size(1000, 680)
    $form.StartPosition = "CenterScreen"
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

    $pnlHeader = New-Object System.Windows.Forms.Panel
    $pnlHeader.Dock = "Top"
    $pnlHeader.Height = 60
    $pnlHeader.BackColor = [System.Drawing.Color]::White
    $pnlHeader.Padding = New-Object System.Windows.Forms.Padding(15, 10, 15, 10)

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = "🔑 " + $t["Title"]
    $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $lblTitle.ForeColor = [System.Drawing.Color]::FromArgb(24, 43, 73)
    $lblTitle.AutoSize = $true
    $lblTitle.Location = New-Object System.Drawing.Point(15, 15)

    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = "🔄 " + $t["BtnLoad"]
    $btnRefresh.Size = New-Object System.Drawing.Size(200, 32)
    $btnRefresh.Location = New-Object System.Drawing.Point(580, 14)
    $btnRefresh.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
    $btnRefresh.ForeColor = [System.Drawing.Color]::White
    $btnRefresh.FlatStyle = "Flat"

    $btnExport = New-Object System.Windows.Forms.Button
    $btnExport.Text = "💾 " + $t["BtnExport"]
    $btnExport.Size = New-Object System.Drawing.Size(150, 32)
    $btnExport.Location = New-Object System.Drawing.Point(790, 14)
    $btnExport.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
    $btnExport.FlatStyle = "Flat"

    $pnlHeader.Controls.AddRange(@($lblTitle, $btnRefresh, $btnExport))

    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Dock = "Fill"
    $tabControl.Padding = New-Object System.Drawing.Point(12, 6)

    $tabDomain = New-Object System.Windows.Forms.TabPage
    $tabDomain.Text = $t["TabDomain"]
    $tabDomain.BackColor = [System.Drawing.Color]::White

    $tabPSO = New-Object System.Windows.Forms.TabPage
    $tabPSO.Text = $t["TabPSO"]
    $tabPSO.BackColor = [System.Drawing.Color]::White

    $tabUser = New-Object System.Windows.Forms.TabPage
    $tabUser.Text = $t["TabUser"]
    $tabUser.BackColor = [System.Drawing.Color]::White

    $tabControl.TabPages.AddRange(@($tabDomain, $tabPSO, $tabUser))

    # Tab 1
    $gridDomain = New-Object System.Windows.Forms.DataGridView
    $gridDomain.Dock = "Fill"
    $gridDomain.BackgroundColor = [System.Drawing.Color]::White
    $gridDomain.BorderStyle = "None"
    $gridDomain.AutoSizeColumnsMode = "Fill"
    $gridDomain.ReadOnly = $true
    $gridDomain.AllowUserToAddRows = $false
    $gridDomain.RowHeadersVisible = $false
    $gridDomain.SelectionMode = "FullRowSelect"
    Apply-StandardGridTheme -Grid $gridDomain -EnableAlternatingRowColor
    $tabDomain.Controls.Add($gridDomain)

    # Tab 2
    $gridPSO = New-Object System.Windows.Forms.DataGridView
    $gridPSO.Dock = "Fill"
    $gridPSO.BackgroundColor = [System.Drawing.Color]::White
    $gridPSO.BorderStyle = "None"
    $gridPSO.AutoSizeColumnsMode = "AllCells"
    $gridPSO.ReadOnly = $true
    $gridPSO.AllowUserToAddRows = $false
    $gridPSO.RowHeadersVisible = $false
    $gridPSO.SelectionMode = "FullRowSelect"
    Apply-StandardGridTheme -Grid $gridPSO -EnableAlternatingRowColor
    $tabPSO.Controls.Add($gridPSO)

    # Tab 3
    $pnlUserTop = New-Object System.Windows.Forms.Panel
    $pnlUserTop.Dock = "Top"
    $pnlUserTop.Height = 55
    $pnlUserTop.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 250)

    $lblUserPrompt = New-Object System.Windows.Forms.Label
    $lblUserPrompt.Text = $t["LblUserSearch"]
    $lblUserPrompt.Location = New-Object System.Drawing.Point(15, 18)
    $lblUserPrompt.AutoSize = $true

    $txtUserSearch = New-Object System.Windows.Forms.TextBox
    $txtUserSearch.Location = New-Object System.Drawing.Point(220, 15)
    $txtUserSearch.Size = New-Object System.Drawing.Size(220, 25)
    $txtUserSearch.Text = $env:USERNAME

    $btnUserCheck = New-Object System.Windows.Forms.Button
    $btnUserCheck.Text = "🔍 " + $t["BtnCheckUser"]
    $btnUserCheck.Location = New-Object System.Drawing.Point(450, 13)
    $btnUserCheck.Size = New-Object System.Drawing.Size(200, 28)
    $btnUserCheck.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
    $btnUserCheck.ForeColor = [System.Drawing.Color]::White
    $btnUserCheck.FlatStyle = "Flat"

    $pnlUserTop.Controls.AddRange(@($lblUserPrompt, $txtUserSearch, $btnUserCheck))

    $gridUser = New-Object System.Windows.Forms.DataGridView
    $gridUser.Dock = "Fill"
    $gridUser.BackgroundColor = [System.Drawing.Color]::White
    $gridUser.BorderStyle = "None"
    $gridUser.AutoSizeColumnsMode = "Fill"
    $gridUser.ReadOnly = $true
    $gridUser.AllowUserToAddRows = $false
    $gridUser.RowHeadersVisible = $false
    $gridUser.SelectionMode = "FullRowSelect"
    Apply-StandardGridTheme -Grid $gridUser -EnableAlternatingRowColor

    $tabUser.Controls.AddRange(@($gridUser, $pnlUserTop))
    $pnlUserTop.SendToBack()
    $gridUser.BringToFront()

    $statusStrip = New-Object System.Windows.Forms.StatusStrip
    $lblStatus = New-Object System.Windows.Forms.ToolStripStatusLabel
    $lblStatus.Text = $t["StatusReady"]
    [void]$statusStrip.Items.Add($lblStatus)

    $form.Controls.AddRange(@($tabControl, $statusStrip, $pnlHeader))
    $pnlHeader.SendToBack()
    $tabControl.BringToFront()

    $script:cachedDomainPolicy = @()
    $script:cachedPSOs = @()

    $loadPolicies = {
        $lblStatus.Text = $t["StatusLoading"]
        [System.Windows.Forms.Application]::DoEvents()

        try {
            $rootDSE = [ADSI]"LDAP://RootDSE"
            $defaultNC = $rootDSE.defaultNamingContext
            if (-not $defaultNC) {
                [System.Windows.Forms.MessageBox]::Show($t["ErrNoDomain"], "AD Error", "OK", "Error")
                return
            }

            # 1. Default Domain Password Policy
            $domEntry = [ADSI]"LDAP://$defaultNC"
            $domProps = $domEntry.Properties

            $minPwdAgeSpan = Convert-LargeIntToTimeSpan $domProps["minPwdAge"].Value
            $maxPwdAgeSpan = Convert-LargeIntToTimeSpan $domProps["maxPwdAge"].Value
            $lockoutDurSpan = Convert-LargeIntToTimeSpan $domProps["lockoutDuration"].Value
            $lockoutWinSpan = Convert-LargeIntToTimeSpan $domProps["lockOutObservationWindow"].Value

            $minPwdAgeDays = if ($minPwdAgeSpan) { [math]::Round($minPwdAgeSpan.TotalDays, 1) } else { 0 }
            $maxPwdAgeDays = if ($maxPwdAgeSpan) { [math]::Round($maxPwdAgeSpan.TotalDays, 0) } else { 0 }
            $lockoutDurMin = if ($lockoutDurSpan) { [math]::Round($lockoutDurSpan.TotalMinutes, 0) } else { 0 }
            $lockoutWinMin = if ($lockoutWinSpan) { [math]::Round($lockoutWinSpan.TotalMinutes, 0) } else { 0 }

            $minPwdLength   = if ($domProps["minPwdLength"].Value) { $domProps["minPwdLength"].Value } else { 0 }
            $pwdHistoryLen  = if ($domProps["pwdHistoryLength"].Value) { $domProps["pwdHistoryLength"].Value } else { 0 }
            $pwdProperties  = if ($domProps["pwdProperties"].Value) { [int]$domProps["pwdProperties"].Value } else { 0 }
            $complexityOn   = ($pwdProperties -band 1) -eq 1
            $reversibleOn   = ($pwdProperties -band 16) -eq 1
            $lockoutThresh  = if ($domProps["lockoutThreshold"].Value) { $domProps["lockoutThreshold"].Value } else { 0 }

            $domainPolicyList = @(
                [PSCustomObject]@{ $t["PropName"] = $t["DPO_MinLength"];  $t["PropValue"] = "$minPwdLength $($t['Characters'])" },
                [PSCustomObject]@{ $t["PropName"] = $t["DPO_Complexity"]; $t["PropValue"] = if ($complexityOn) { "✔ $($t['Enabled'])" } else { "✖ $($t['Disabled'])" } },
                [PSCustomObject]@{ $t["PropName"] = $t["DPO_MaxAge"];     $t["PropValue"] = if ($maxPwdAgeDays -gt 0) { "$maxPwdAgeDays $($t['Days'])" } else { $t["NeverExpires"] } },
                [PSCustomObject]@{ $t["PropName"] = $t["DPO_MinAge"];     $t["PropValue"] = "$minPwdAgeDays $($t['Days'])" },
                [PSCustomObject]@{ $t["PropName"] = $t["DPO_History"];    $t["PropValue"] = "$pwdHistoryLen $($t['Passwords'])" },
                [PSCustomObject]@{ $t["PropName"] = $t["DPO_LockoutTh"];  $t["PropValue"] = if ($lockoutThresh -gt 0) { "$lockoutThresh $($t['InvalidAttempts'])" } else { $t["NoLockout"] } },
                [PSCustomObject]@{ $t["PropName"] = $t["DPO_LockoutDur"]; $t["PropValue"] = if ($lockoutThresh -gt 0) { "$lockoutDurMin $($t['Minutes'])" } else { "-" } },
                [PSCustomObject]@{ $t["PropName"] = $t["DPO_ObsWindow"];  $t["PropValue"] = if ($lockoutThresh -gt 0) { "$lockoutWinMin $($t['Minutes'])" } else { "-" } },
                [PSCustomObject]@{ $t["PropName"] = $t["DPO_RevEncrypt"]; $t["PropValue"] = if ($reversibleOn) { "⚠️ $($t['Enabled']) ($($t['Unsafe']))" } else { "$($t['Disabled']) ($($t['Safe']))" } }
            )

            $script:cachedDomainPolicy = $domainPolicyList
            $gridDomain.DataSource = [System.Collections.ArrayList]::new($domainPolicyList)

            # 2. Fine-Grained Password Policies (PSO)
            $psoContainer = [ADSI]"LDAP://CN=Password Settings Objects,CN=System,$defaultNC"
            $psoSearcher = New-Object System.DirectoryServices.DirectorySearcher($psoContainer)
            $psoSearcher.Filter = "(objectClass=msDS-PasswordSettings)"
            $psoResults = $psoSearcher.FindAll()

            $psoList = [System.Collections.ArrayList]::new()

            foreach ($p in $psoResults) {
                $entry = $p.GetDirectoryEntry()
                $pName = $entry.Properties["name"].Value
                $precedence = $entry.Properties["msDS-PasswordSettingsPrecedence"].Value
                $pComplexity = [bool]$entry.Properties["msDS-PasswordComplexityEnabled"].Value
                $pMinLen = $entry.Properties["msDS-MinimumPasswordLength"].Value
                $pHist = $entry.Properties["msDS-PasswordHistoryLength"].Value
                $pLockThresh = $entry.Properties["msDS-LockoutThreshold"].Value

                $pMaxAgeSpan = Convert-LargeIntToTimeSpan $entry.Properties["msDS-MaximumPasswordAge"].Value
                $pMinAgeSpan = Convert-LargeIntToTimeSpan $entry.Properties["msDS-MinimumPasswordAge"].Value
                $pLockDurSpan = Convert-LargeIntToTimeSpan $entry.Properties["msDS-LockoutDuration"].Value

                $pMaxAge = if ($pMaxAgeSpan) { [math]::Round($pMaxAgeSpan.TotalDays, 1) } else { 0 }
                $pMinAge = if ($pMinAgeSpan) { [math]::Round($pMinAgeSpan.TotalDays, 1) } else { 0 }
                $pLockDur = if ($pLockDurSpan) { [math]::Round($pLockDurSpan.TotalMinutes, 0) } else { 0 }

                $appliesTo = @()
                foreach ($app in $entry.Properties["msDS-PSOAppliesTo"]) {
                    $appliesTo += ($app -split ",*..=")[1]
                }
                $appliesToStr = if ($appliesTo.Count -gt 0) { $appliesTo -join ", " } else { $t["NoneAssigned"] }

                [void]$psoList.Add([PSCustomObject]@{
                    $t["PsoName"]          = $pName
                    $t["PsoPrecedence"]    = $precedence
                    $t["PsoMinLength"]     = $pMinLen
                    $t["PsoComplexity"]    = if ($pComplexity) { $t["Yes"] } else { $t["No"] }
                    $t["PsoMaxAge"]        = $pMaxAge
                    $t["PsoMinAge"]        = $pMinAge
                    $t["PsoHistory"]       = $pHist
                    $t["PsoLockoutThresh"] = $pLockThresh
                    $t["PsoLockoutDur"]    = $pLockDur
                    $t["PsoAppliesTo"]     = $appliesToStr
                })
            }

            $sortedPSO = $psoList | Sort-Object { $_.($t["PsoPrecedence"]) }
            $script:cachedPSOs = [System.Collections.ArrayList]::new($sortedPSO)
            $gridPSO.DataSource = $script:cachedPSOs

            $lblStatus.Text = "$($t["StatusDone"]) (PSOs: $($psoList.Count))"
        } catch {
            $lblStatus.Text = "Fehler: " + $_.Exception.Message
        }
    }

    $btnUserCheck.Add_Click({
        $username = $txtUserSearch.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($username)) { return }

        try {
            $rootDSE = [ADSI]"LDAP://RootDSE"
            $defaultNC = $rootDSE.defaultNamingContext
            $uSearcher = New-Object System.DirectoryServices.DirectorySearcher([ADSI]"LDAP://$defaultNC")
            $uSearcher.Filter = "(&(objectCategory=person)(objectClass=user)(sAMAccountName=$username))"
            $uSearcher.PropertiesToLoad.Add("msDS-ResultantPSO") | Out-Null
            $uSearcher.PropertiesToLoad.Add("distinguishedName") | Out-Null
            $uSearcher.PropertiesToLoad.Add("displayName") | Out-Null

            $uRes = $uSearcher.FindOne()
            if (-not $uRes) {
                [System.Windows.Forms.MessageBox]::Show($t["ErrUserNotFound"], "User Check", "OK", "Warning")
                return
            }

            $userResultList = @()
            $appliedPSO = $uRes.Properties["msds-resultantpso"]

            if ($appliedPSO -and $appliedPSO.Count -gt 0) {
                $psoDN = $appliedPSO[0]
                $psoEntry = [ADSI]"LDAP://$psoDN"

                $pName = $psoEntry.Properties["name"].Value
                $pMinLen = $psoEntry.Properties["msDS-MinimumPasswordLength"].Value
                $pComplexity = [bool]$psoEntry.Properties["msDS-PasswordComplexityEnabled"].Value
                $pHist = $psoEntry.Properties["msDS-PasswordHistoryLength"].Value
                $pLockThresh = $psoEntry.Properties["msDS-LockoutThreshold"].Value

                $pMaxAgeSpan = Convert-LargeIntToTimeSpan $psoEntry.Properties["msDS-MaximumPasswordAge"].Value
                $pMaxAge = if ($pMaxAgeSpan) { [math]::Round($pMaxAgeSpan.TotalDays, 1) } else { 0 }

                $userResultList = @(
                    [PSCustomObject]@{ $t["PropName"] = $t["Eff_Policy"];     $t["PropValue"] = "PSO: $pName" },
                    [PSCustomObject]@{ $t["PropName"] = $t["DPO_MinLength"];  $t["PropValue"] = "$pMinLen $($t['Characters'])" },
                    [PSCustomObject]@{ $t["PropName"] = $t["DPO_Complexity"]; $t["PropValue"] = if ($pComplexity) { $t["Enabled"] } else { $t["Disabled"] } },
                    [PSCustomObject]@{ $t["PropName"] = $t["DPO_MaxAge"];     $t["PropValue"] = "$pMaxAge $($t['Days'])" },
                    [PSCustomObject]@{ $t["PropName"] = $t["DPO_History"];    $t["PropValue"] = "$pHist $($t['Passwords'])" },
                    [PSCustomObject]@{ $t["PropName"] = $t["Eff_LockoutTh"];  $t["PropValue"] = "$pLockThresh $($t['Attempts'])" },
                    [PSCustomObject]@{ $t["PropName"] = $t["Eff_PSO_DN"];     $t["PropValue"] = $psoDN }
                )
            } else {
                $userResultList = @(
                    [PSCustomObject]@{ $t["PropName"] = $t["Eff_Policy"]; $t["PropValue"] = $t["InheritedDomain"] }
                ) + $script:cachedDomainPolicy
            }

            $gridUser.DataSource = [System.Collections.ArrayList]::new($userResultList)
            $lblStatus.Text = "$($t["UserResultTitle"]) '$username' ($($t['StatusDone']))."
        } catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Error", "OK", "Error")
        }
    })

    $txtUserSearch.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $_.SuppressKeyPress = $true
            $btnUserCheck.PerformClick()
        }
    })

    $btnExport.Add_Click({
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "CSV-Datei (*.csv)|*.csv"
        $sfd.FileName = "AD_Password_Policies_$((Get-Date).ToString('yyyyMMdd')).csv"

        if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $exportData = @()
            $defaultTypeStr = if ($lang -eq "EN") { "Default Domain Policy" } else { "Standard-Domänenrichtlinie" }
            $psoTypeStr     = if ($lang -eq "EN") { "Fine-Grained PSO" } else { "Feingranulare PSO" }

            foreach ($row in $script:cachedDomainPolicy) {
                $exportData += [PSCustomObject]@{
                    "Typ / Type"               = $defaultTypeStr
                    "Eigenschaft / Setting"    = $row.($t["PropName"])
                    "Wert / Value"             = $row.($t["PropValue"])
                }
            }
            foreach ($row in $script:cachedPSOs) {
                $exportData += [PSCustomObject]@{
                    "Typ / Type"               = "$psoTypeStr ($($row.($t['PsoName'])))"
                    "Eigenschaft / Setting"    = "Precedence: $($row.($t['PsoPrecedence'])), MinLen: $($row.($t['PsoMinLength']))"
                    "Wert / Value"             = "AppliesTo: $($row.($t['PsoAppliesTo']))"
                }
            }
            $exportData | Export-Csv -Path $sfd.FileName -NoTypeInformation -Delimiter ";" -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show("Kennwortrichtlinien exportiert:`n$($sfd.FileName)", "Export OK", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    })

    $btnRefresh.Add_Click({ & $loadPolicies })
    $form.Add_Shown({ & $loadPolicies })

    [void]$form.ShowDialog()
}
