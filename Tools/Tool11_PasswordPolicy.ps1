<#
==============================================================================
TOOL 11: AD PASSWORD POLICIES & FINE-GRAINED PASSWORD AUDIT (PSO)
Native LDAP / ADSI Implementation - Inkl. Letzte Kennwortänderung (pwdLastSet)
==============================================================================
#>

function Show-Tool11-PasswordPolicyAudit {
    [CmdletBinding()]
    param()

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.DirectoryServices

    # Sprachwörterbuch für Tool 11
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
            "BtnCheckUser"    = "Effektive Richtlinie & Kennwortalter prüfen"
            "StatusReady"     = "Bereit zur Abfrage."
            "StatusLoading"   = "Lese Active Directory Kennwortrichtlinien aus..."
            "StatusDone"      = "Kennwortrichtlinien erfolgreich geladen."
            "ErrNoDomain"     = "Keine Active Directory Domäne erreichbar."
            "ErrUserNotFound" = "Benutzer nicht im Active Directory gefunden."
            "UserResultTitle" = "Ergebnis für Benutzer"
            "InheritedDomain" = "Standard-Domänenrichtlinie (Keine spezifische PSO)"
            "Days"            = "Tage"
            "Minutes"         = "Minuten"
            "Enabled"         = "Aktiviert"
            "Disabled"        = "Deaktiviert"
            "LastPwdSet"      = "Letzte Kennwortänderung"
            "PwdExpires"      = "Kennwort läuft ab am"
            "PwdMustChange"   = "⚠️ Kennwort muss bei nächster Anmeldung geändert werden"
            "PwdNeverExpires" = "Nie (Kennwort läuft nicht ab)"
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
            "BtnCheckUser"    = "Check Effective Policy & Password Age"
            "StatusReady"     = "Ready for query."
            "StatusLoading"   = "Querying Active Directory Password Policies..."
            "StatusDone"      = "Password policies loaded successfully."
            "ErrNoDomain"     = "No Active Directory Domain accessible."
            "ErrUserNotFound" = "User not found in Active Directory."
            "UserResultTitle" = "Result for user"
            "InheritedDomain" = "Default Domain Policy (No specific PSO applied)"
            "Days"            = "Days"
            "Minutes"         = "Minutes"
            "Enabled"         = "Enabled"
            "Disabled"        = "Disabled"
            "LastPwdSet"      = "Last Password Change"
            "PwdExpires"      = "Password Expires On"
            "PwdMustChange"   = "⚠️ User must change password at next logon"
            "PwdNeverExpires" = "Never (Password does not expire)"
        }
    }[$lang]

    # Hilfsfunktion: I8 LargeInteger (TimeSpan in 100ns Ticks)
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

    # Hilfsfunktion: I8 LargeInteger (FileTime / pwdLastSet) in DateTime
    function Convert-LargeIntToFileTime([object]$largeIntObj) {
        if ($null -eq $largeIntObj) { return $null }
        try {
            $highPart = $largeIntObj.GetType().InvokeMember("HighPart", [System.Reflection.BindingFlags]::GetProperty, $null, $largeIntObj, $null)
            $lowPart  = $largeIntObj.GetType().InvokeMember("LowPart", [System.Reflection.BindingFlags]::GetProperty, $null, $largeIntObj, $null)
            $raw64 = ([int64]$highPart -shl 32) -bor ([int64]$lowPart -band 0xFFFFFFFF)
            
            if ($raw64 -eq 0) { return "MUST_CHANGE" }
            if ($raw64 -eq [int64]::MaxValue -or $raw64 -lt 0) { return "NEVER" }
            
            return [DateTime]::FromFileTime($raw64)
        } catch {
            return $null
        }
    }

    # GUI Erstellung
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $t["Title"]
    $form.Size = New-Object System.Drawing.Size(1050, 720)
    $form.StartPosition = "CenterScreen"
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

    # Top Header Panel
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
    $btnRefresh.Location = New-Object System.Drawing.Point(620, 14)
    $btnRefresh.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
    $btnRefresh.ForeColor = [System.Drawing.Color]::White
    $btnRefresh.FlatStyle = "Flat"

    $btnExport = New-Object System.Windows.Forms.Button
    $btnExport.Text = "💾 " + $t["BtnExport"]
    $btnExport.Size = New-Object System.Drawing.Size(150, 32)
    $btnExport.Location = New-Object System.Drawing.Point(830, 14)
    $btnExport.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
    $btnExport.FlatStyle = "Flat"

    $pnlHeader.Controls.AddRange(@($lblTitle, $btnRefresh, $btnExport))

    # Tabs
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

    # Tab 1: Grid
    $gridDomain = New-Object System.Windows.Forms.DataGridView
    $gridDomain.Dock = "Fill"
    $gridDomain.BackgroundColor = [System.Drawing.Color]::White
    $gridDomain.BorderStyle = "None"
    $gridDomain.AutoSizeColumnsMode = "Fill"
    $gridDomain.ReadOnly = $true
    $gridDomain.AllowUserToAddRows = $false
    $gridDomain.RowHeadersVisible = $false
    $gridDomain.SelectionMode = "FullRowSelect"
    $tabDomain.Controls.Add($gridDomain)

    # Tab 2: Grid
    $gridPSO = New-Object System.Windows.Forms.DataGridView
    $gridPSO.Dock = "Fill"
    $gridPSO.BackgroundColor = [System.Drawing.Color]::White
    $gridPSO.BorderStyle = "None"
    $gridPSO.AutoSizeColumnsMode = "AllCells"
    $gridPSO.ReadOnly = $true
    $gridPSO.AllowUserToAddRows = $false
    $gridPSO.RowHeadersVisible = $false
    $gridPSO.SelectionMode = "FullRowSelect"
    $tabPSO.Controls.Add($gridPSO)

    # Tab 3: User Policy Check
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
    $btnUserCheck.Size = New-Object System.Drawing.Size(260, 28)
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

    $tabUser.Controls.AddRange(@($gridUser, $pnlUserTop))

    # Bottom Status Bar
    $statusStrip = New-Object System.Windows.Forms.StatusStrip
    $lblStatus = New-Object System.Windows.Forms.ToolStripStatusLabel
    $lblStatus.Text = $t["StatusReady"]
    $statusStrip.Items.Add($lblStatus)

    $form.Controls.AddRange(@($tabControl, $statusStrip, $pnlHeader))

    # State Cache
    $script:cachedDomainPolicy = @()
    $script:cachedPSOs = @()
    $script:rawDomainMaxPwdAgeDays = 0

    # Laderoutine für Policies
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
            $script:rawDomainMaxPwdAgeDays = $maxPwdAgeDays

            $lockoutDurMin = if ($lockoutDurSpan) { [math]::Round($lockoutDurSpan.TotalMinutes, 0) } else { 0 }
            $lockoutWinMin = if ($lockoutWinSpan) { [math]::Round($lockoutWinSpan.TotalMinutes, 0) } else { 0 }

            $minPwdLength   = if ($domProps["minPwdLength"].Value) { $domProps["minPwdLength"].Value } else { 0 }
            $pwdHistoryLen  = if ($domProps["pwdHistoryLength"].Value) { $domProps["pwdHistoryLength"].Value } else { 0 }
            $pwdProperties  = if ($domProps["pwdProperties"].Value) { [int]$domProps["pwdProperties"].Value } else { 0 }
            $complexityOn   = ($pwdProperties -band 1) -eq 1
            $reversibleOn   = ($pwdProperties -band 16) -eq 1
            $lockoutThresh  = if ($domProps["lockoutThreshold"].Value) { $domProps["lockoutThreshold"].Value } else { 0 }

            $domainPolicyList = @(
                [PSCustomObject]@{ $t["PropName"] = "Minimale Kennwortlänge (Minimum Password Length)"; $t["PropValue"] = "$minPwdLength Zeichen" },
                [PSCustomObject]@{ $t["PropName"] = "Kennwortkomplexität (Password Complexity)"; $t["PropValue"] = if ($complexityOn) { "✔ " + $t["Enabled"] } else { "✖ " + $t["Disabled"] } },
                [PSCustomObject]@{ $t["PropName"] = "Maximales Kennwortalter (Max Password Age)"; $t["PropValue"] = if ($maxPwdAgeDays -gt 0) { "$maxPwdAgeDays " + $t["Days"] } else { "Nie ablaufend / Deaktiviert" } },
                [PSCustomObject]@{ $t["PropName"] = "Minimales Kennwortalter (Min Password Age)"; $t["PropValue"] = "$minPwdAgeDays " + $t["Days"] },
                [PSCustomObject]@{ $t["PropName"] = "Kennworthistorie / Verlauf (Password History)"; $t["PropValue"] = "$pwdHistoryLen gemerkte Kennwörter" },
                [PSCustomObject]@{ $t["PropName"] = "Kontosperrungsschwelle (Account Lockout Threshold)"; $t["PropValue"] = if ($lockoutThresh -gt 0) { "$lockoutThresh ungültige Versuche" } else { "Keine Sperre (0)" } },
                [PSCustomObject]@{ $t["PropName"] = "Kontosperrdauer (Lockout Duration)"; $t["PropValue"] = if ($lockoutThresh -gt 0) { "$lockoutDurMin " + $t["Minutes"] } else { "-" } },
                [PSCustomObject]@{ $t["PropName"] = "Sperrungsbeobachtungsfenster (Observation Window)"; $t["PropValue"] = if ($lockoutThresh -gt 0) { "$lockoutWinMin " + $t["Minutes"] } else { "-" } },
                [PSCustomObject]@{ $t["PropName"] = "Umkehrbare Verschlüsselung (Reversible Encryption)"; $t["PropValue"] = if ($reversibleOn) { "⚠️ Aktiviert (Unsicher)" } else { "Deaktiviert (Sicher)" } }
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
                $appliesToStr = if ($appliesTo.Count -gt 0) { $appliesTo -join ", " } else { "Niemand (Keine Zuweisung)" }

                [void]$psoList.Add([PSCustomObject]@{
                    $t["PsoName"]          = $pName
                    $t["PsoPrecedence"]    = $precedence
                    $t["PsoMinLength"]     = $pMinLen
                    $t["PsoComplexity"]    = if ($pComplexity) { "Ja" } else { "Nein" }
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

    # Tab 3: Benutzerprüfung inkl. pwdLastSet
    $btnUserCheck.Add_Click({
        $username = $txtUserSearch.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($username)) { return }

        try {
            $rootDSE = [ADSI]"LDAP://RootDSE"
            $defaultNC = $rootDSE.defaultNamingContext
            $uSearcher = New-Object System.DirectoryServices.DirectorySearcher([ADSI]"LDAP://$defaultNC")
            $uSearcher.Filter = "(&(objectCategory=person)(objectClass=user)(|(sAMAccountName=$username)(userPrincipalName=$username)))"
            $uSearcher.PropertiesToLoad.AddRange(@("msDS-ResultantPSO", "distinguishedName", "displayName", "pwdLastSet", "userAccountControl"))

            $uRes = $uSearcher.FindOne()
            if (-not $uRes) {
                [System.Windows.Forms.MessageBox]::Show($t["ErrUserNotFound"], "User Check", "OK", "Warning")
                return
            }

            $uEntry = $uRes.GetDirectoryEntry()
            $uac = if ($uEntry.Properties["userAccountControl"].Value) { [int]$uEntry.Properties["userAccountControl"].Value } else { 0 }
            $dontExpire = ($uac -band 65536) -eq 65536 # DONT_EXPIRE_PASSWORD Flag (0x10000)

            # pwdLastSet parsen
            $rawPwdLastSet = $uEntry.Properties["pwdLastSet"].Value
            $pwdLastSetVal = Convert-LargeIntToFileTime $rawPwdLastSet

            $lastSetDisplay = ""
            $expiresDisplay = ""
            $effectiveMaxAgeDays = 0

            # Richtlinie auslesen (PSO oder Domain)
            $appliedPSO = $uRes.Properties["msds-resultantpso"]
            $policyDetails = @()

            if ($appliedPSO -and $appliedPSO.Count -gt 0) {
                $psoDN = $appliedPSO[0]
                $psoEntry = [ADSI]"LDAP://$psoDN"

                $pName = $psoEntry.Properties["name"].Value
                $pMinLen = $psoEntry.Properties["msDS-MinimumPasswordLength"].Value
                $pComplexity = [bool]$psoEntry.Properties["msDS-PasswordComplexityEnabled"].Value
                $pHist = $psoEntry.Properties["msDS-PasswordHistoryLength"].Value
                $pLockThresh = $psoEntry.Properties["msDS-LockoutThreshold"].Value

                $pMaxAgeSpan = Convert-LargeIntToTimeSpan $psoEntry.Properties["msDS-MaximumPasswordAge"].Value
                $effectiveMaxAgeDays = if ($pMaxAgeSpan) { [math]::Round($pMaxAgeSpan.TotalDays, 1) } else { 0 }

                $policyDetails = @(
                    [PSCustomObject]@{ $t["PropName"] = "Gültige Richtlinie (Effective Policy)"; $t["PropValue"] = "PSO: $pName" },
                    [PSCustomObject]@{ $t["PropName"] = "Minimale Kennwortlänge"; $t["PropValue"] = "$pMinLen Zeichen" },
                    [PSCustomObject]@{ $t["PropName"] = "Kennwortkomplexität"; $t["PropValue"] = if ($pComplexity) { "Aktiviert" } else { "Deaktiviert" } },
                    [PSCustomObject]@{ $t["PropName"] = "Maximales Kennwortalter"; $t["PropValue"] = "$effectiveMaxAgeDays " + $t["Days"] },
                    [PSCustomObject]@{ $t["PropName"] = "Kennworthistorie"; $t["PropValue"] = "$pHist Passwörter" },
                    [PSCustomObject]@{ $t["PropName"] = "Kontosperrung nach Fehlversuchen"; $t["PropValue"] = "$pLockThresh Versuche" },
                    [PSCustomObject]@{ $t["PropName"] = "PSO DistinguishedName"; $t["PropValue"] = $psoDN }
                )
            } else {
                $effectiveMaxAgeDays = $script:rawDomainMaxPwdAgeDays
                $policyDetails = @(
                    [PSCustomObject]@{ $t["PropName"] = "Gültige Richtlinie (Effective Policy)"; $t["PropValue"] = $t["InheritedDomain"] }
                ) + $script:cachedDomainPolicy
            }

            # Berechnung von Letzter Änderung & Ablauf
            if ($pwdLastSetVal -eq "MUST_CHANGE") {
                $lastSetDisplay = $t["PwdMustChange"]
                $expiresDisplay = $t["PwdMustChange"]
            } elseif ($pwdLastSetVal -is [DateTime]) {
                $lastSetDisplay = "$($pwdLastSetVal.ToString('dd.MM.yyyy HH:mm:ss')) (Vor $(([math]::Round(((Get-Date) - $pwdLastSetVal).TotalDays, 1))) Tagen)"
                
                if ($dontExpire -or $effectiveMaxAgeDays -le 0) {
                    $expiresDisplay = $t["PwdNeverExpires"]
                } else {
                    $expireDate = $pwdLastSetVal.AddDays($effectiveMaxAgeDays)
                    $daysLeft = [math]::Round(($expireDate - (Get-Date)).TotalDays, 1)
                    if ($daysLeft -lt 0) {
                        $expiresDisplay = "⛔ Abgelaufen am $($expireDate.ToString('dd.MM.yyyy HH:mm:ss')) (vor $([math]::Abs($daysLeft)) Tagen)"
                    } else {
                        $expiresDisplay = "$($expireDate.ToString('dd.MM.yyyy HH:mm:ss')) (in $daysLeft Tagen)"
                    }
                }
            } else {
                $lastSetDisplay = "Unbekannt / Nicht gesetzt"
                $expiresDisplay = "Unbekannt"
            }

            # Benutzerbezogene Metadaten ganz oben einfügen
            $userResultList = @(
                [PSCustomObject]@{ $t["PropName"] = "👤 " + $t["LastPwdSet"]; $t["PropValue"] = $lastSetDisplay },
                [PSCustomObject]@{ $t["PropName"] = "⏳ " + $t["PwdExpires"];  $t["PropValue"] = $expiresDisplay },
                [PSCustomObject]@{ $t["PropName"] = "Kennwort läuft nie ab (User Flag)"; $t["PropValue"] = if ($dontExpire) { "✔ Ja (DONT_EXPIRE_PASSWORD gesetzt)" } else { "Nein" } },
                [PSCustomObject]@{ $t["PropName"] = "----------------------------------------"; $t["PropValue"] = "----------------------------------------" }
            ) + $policyDetails

            $gridUser.DataSource = [System.Collections.ArrayList]::new($userResultList)
            $lblStatus.Text = "$($t["UserResultTitle"]) '$username' erfolgreich analysiert."
        } catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Error", "OK", "Error")
        }
    })

    # Enter-Taste für Textbox
    $txtUserSearch.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $_.SuppressKeyPress = $true
            $btnUserCheck.PerformClick()
        }
    })

    # CSV Export
    $btnExport.Add_Click({
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "CSV-Datei (*.csv)|*.csv"
        $sfd.FileName = "AD_Password_Policies_$((Get-Date).ToString('yyyyMMdd')).csv"

        if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $exportData = @()
            foreach ($row in $script:cachedDomainPolicy) {
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
            [System.Windows.Forms.MessageBox]::Show("Kennwortrichtlinien exportiert:`n$($sfd.FileName)", "Export OK", "OK", "Information")
        }
    })

    # Events
    $btnRefresh.Add_Click({ & $loadPolicies })
    $form.Add_Shown({ & $loadPolicies })

    [void]$form.ShowDialog()
}
