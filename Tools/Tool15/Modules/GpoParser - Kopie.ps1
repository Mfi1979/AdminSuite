# =========================================================================
# GpoParser.ps1 - Vollstaendiger 6-Komponenten XML-Parser (Sauberes Parsing)
# =========================================================================

$script:knownAccountPolicies = @{
    "enforcepasswordhistory" = @{ Name = "Enforce password history"; Unit = "passwords remembered"; Cat = "Sicherheitseinstellungen / Kontorichtlinien / Kennwortrichtlinie" }
    "passwordhistory"        = @{ Name = "Enforce password history"; Unit = "passwords remembered"; Cat = "Sicherheitseinstellungen / Kontorichtlinien / Kennwortrichtlinie" }
    "maximumpasswordage"     = @{ Name = "Maximum password age"; Unit = "days"; Cat = "Sicherheitseinstellungen / Kontorichtlinien / Kennwortrichtlinie" }
    "maxpasswordage"         = @{ Name = "Maximum password age"; Unit = "days"; Cat = "Sicherheitseinstellungen / Kontorichtlinien / Kennwortrichtlinie" }
    "minimumpasswordage"     = @{ Name = "Minimum password age"; Unit = "days"; Cat = "Sicherheitseinstellungen / Kontorichtlinien / Kennwortrichtlinie" }
    "minpasswordage"         = @{ Name = "Minimum password age"; Unit = "days"; Cat = "Sicherheitseinstellungen / Kontorichtlinien / Kennwortrichtlinie" }
    "minimumpasswordlength"  = @{ Name = "Minimum password length"; Unit = "characters"; Cat = "Sicherheitseinstellungen / Kontorichtlinien / Kennwortrichtlinie" }
    "minpasswordlength"      = @{ Name = "Minimum password length"; Unit = "characters"; Cat = "Sicherheitseinstellungen / Kontorichtlinien / Kennwortrichtlinie" }
    "passwordcomplexity"     = @{ Name = "Password must meet complexity requirements"; Unit = ""; Cat = "Sicherheitseinstellungen / Kontorichtlinien / Kennwortrichtlinie" }
    "cleartextpassword"      = @{ Name = "Store passwords using reversible encryption"; Unit = ""; Cat = "Sicherheitseinstellungen / Kontorichtlinien / Kennwortrichtlinie" }
    "lockoutbadcount"        = @{ Name = "Account lockout threshold"; Unit = "invalid logon attempts"; Cat = "Sicherheitseinstellungen / Kontorichtlinien / Kontosperrungsrichtlinie" }
    "lockoutduration"        = @{ Name = "Account lockout duration"; Unit = "minutes"; Cat = "Sicherheitseinstellungen / Kontorichtlinien / Kontosperrungsrichtlinie" }
    "resetlockoutcount"      = @{ Name = "Reset account lockout counter after"; Unit = "minutes"; Cat = "Sicherheitseinstellungen / Kontorichtlinien / Kontosperrungsrichtlinie" }
    "maxclockskew"           = @{ Name = "Maximum tolerance for computer clock synchronization"; Unit = "minutes"; Cat = "Sicherheitseinstellungen / Kontorichtlinien / Kerberos-Richtlinie" }
    "maxticketage"           = @{ Name = "Maximum lifetime for user ticket"; Unit = "hours"; Cat = "Sicherheitseinstellungen / Kontorichtlinien / Kerberos-Richtlinie" }
    "maxserviceage"          = @{ Name = "Maximum lifetime for service ticket"; Unit = "minutes"; Cat = "Sicherheitseinstellungen / Kontorichtlinien / Kerberos-Richtlinie" }
    "maxrenewage"            = @{ Name = "Maximum lifetime for user ticket renewal"; Unit = "days"; Cat = "Sicherheitseinstellungen / Kontorichtlinien / Kerberos-Richtlinie" }
    "ticketvalidateclient"   = @{ Name = "Enforce user logon restrictions"; Unit = ""; Cat = "Sicherheitseinstellungen / Kontorichtlinien / Kerberos-Richtlinie" }
}

function Get-ParsedGpoSettings {
    param([string]$GpoId, [string]$GpoDisplayName)

    $dateStr = Get-Date -Format "yyyyMMdd"
    $timeStr = Get-Date -Format "HHmm"
    $list = [System.Collections.Generic.List[PSCustomObject]]::new()

    [xml]$xml = Get-GPOReport -Guid $GpoId -ReportType Xml -ErrorAction Stop

    function Parse-GpoSection ($sectionNode, $scope) {
        if ($null -eq $sectionNode -or -not $sectionNode.ExtensionData) { return }

        foreach ($ext in $sectionNode.ExtensionData.Extension) {
            $extCategory = if ($ext.Name) { $ext.Name } else { "Erweiterung" }

            # 1. Administrative Vorlagen (<Policy>)
            $policies = $ext.SelectNodes(".//*[local-name()='Policy']")
            if ($policies -and $policies.Count -gt 0) {
                foreach ($p in $policies) {
                    $pName = if ($p.SelectSingleNode("./*[local-name()='Name']")) { $p.SelectSingleNode("./*[local-name()='Name']").InnerText.Trim() } elseif ($p.Name) { $p.Name.Trim() } else { "Unbenannte Richtlinie" }
                    $rawState = if ($p.SelectSingleNode("./*[local-name()='State']")) { $p.SelectSingleNode("./*[local-name()='State']").InnerText.Trim() } elseif ($p.State) { $p.State.Trim() } else { "Enabled" }
                    $pState = switch ($rawState) { "Enabled" { "Aktiviert" } "Disabled" { "Deaktiviert" } default { $rawState } }
                    $pCat = if ($p.SelectSingleNode("./*[local-name()='Category']")) { $p.SelectSingleNode("./*[local-name()='Category']").InnerText.Trim() } else { $extCategory }
                    $pSupported = if ($p.SelectSingleNode("./*[local-name()='Supported' or local-name()='SupportedOn']")) { $p.SelectSingleNode("./*[local-name()='Supported' or local-name()='SupportedOn']").InnerText.Trim() } else { "Keine Angabe" }
                    $pExplain = if ($p.SelectSingleNode("./*[local-name()='Explain' or local-name()='ExplainText']")) { $p.SelectSingleNode("./*[local-name()='Explain' or local-name()='ExplainText']").InnerText.Trim() } else { "Keine Erklaerung hinterlegt." }

                    $ignoredTags = @("Name", "State", "Explain", "ExplainText", "Supported", "SupportedOn", "Category", "Text")
                    $paramValues = @()

                    foreach ($child in $p.ChildNodes) {
                        if ($child.LocalName -in $ignoredTags) { continue }
                        $valNodes = $child.SelectNodes(".//*[local-name()='Value' or local-name()='Data' or local-name()='Setting' or local-name()='Decimal' or local-name()='String']")
                        $nameNode = $child.SelectSingleNode("./*[local-name()='Name' or local-name()='Label']")
                        $optLabel = if ($nameNode) { $nameNode.InnerText.Trim() } else { "" }

                        $extractedList = @()
                        if ($valNodes -and $valNodes.Count -gt 0) {
                            foreach ($vn in $valNodes) {
                                if (-not [string]::IsNullOrWhiteSpace($vn.InnerText)) { $extractedList += $vn.InnerText.Trim() }
                            }
                        } elseif ($child.Attributes["value"]) {
                            $extractedList += $child.Attributes["value"].Value.Trim()
                        } elseif (-not [string]::IsNullOrWhiteSpace($child.InnerText) -and $child.ChildNodes.Count -le 1) {
                            $extractedList += $child.InnerText.Trim()
                        }

                        if ($extractedList.Count -gt 0) {
                            $joinedVals = $extractedList -join ", "
                            if (-not [string]::IsNullOrWhiteSpace($optLabel) -and $optLabel -ne $joinedVals) {
                                $paramValues += "$($optLabel): $($joinedVals)"
                            } else {
                                $paramValues += "$joinedVals"
                            }
                        }
                    }

                    $cleanValue = if ($paramValues.Count -gt 0) { $paramValues -join " | " } else { $pState }

                    $list.Add([PSCustomObject]@{
                        Scope     = $scope
                        Category  = $pCat
                        Name      = $pName
                        Value     = $cleanValue
                        State     = $pState
                        Supported = $pSupported
                        Explain   = $pExplain
                        GpoName   = $GpoDisplayName
                        Datum     = $dateStr
                        Uhrzeit   = $timeStr
                    })
                }
            }

            # 2. Systemdienste
            $services = $ext.SelectNodes(".//*[local-name()='SystemServices'] | .//*[local-name()='Service']")
            if ($services -and $services.Count -gt 0) {
                foreach ($svc in $services) {
                    $svcName = if ($svc.SelectSingleNode("./*[local-name()='Display']/*[local-name()='Name']")) { $svc.SelectSingleNode("./*[local-name()='Display']/*[local-name()='Name']").InnerText.Trim() } elseif ($svc.SelectSingleNode("./*[local-name()='Name']")) { $svc.SelectSingleNode("./*[local-name()='Name']").InnerText.Trim() } else { "Systemdienst" }
                    $mode = if ($svc.SelectSingleNode("./*[local-name()='StartupMode']")) { $svc.SelectSingleNode("./*[local-name()='StartupMode']").InnerText.Trim() } elseif ($svc.SelectSingleNode("./*[local-name()='Mode']")) { $svc.SelectSingleNode("./*[local-name()='Mode']").InnerText.Trim() } else { "Konfiguriert" }
                    $modeDE = switch ($mode) { "Disabled" { "Deaktiviert" } "Automatic" { "Automatisch" } "Manual" { "Manuell" } default { $mode } }

                    $list.Add([PSCustomObject]@{
                        Scope     = $scope
                        Category  = "Sicherheitseinstellungen / Systemdienste"
                        Name      = $svcName
                        Value     = "Starttyp: $modeDE"
                        State     = $modeDE
                        Supported = "Windows Systemdienste"
                        Explain   = "Startmodus fuer den Dienst '$svcName': $modeDE"
                        GpoName   = $GpoDisplayName
                        Datum     = $dateStr
                        Uhrzeit   = $timeStr
                    })
                }
            }

            # 3. GPP (Praeferenzen)
            $gppNodes = $ext.SelectNodes(".//*[local-name()='Properties']")
            if ($gppNodes -and $gppNodes.Count -gt 0) {
                foreach ($prop in $gppNodes) {
                    $parent = $prop.ParentNode
                    $itemType = $parent.LocalName
                    $gppCategory = "Praeferenzen (GPP) / $itemType"

                    $itemName = if ($parent.Attributes["name"]) { $parent.Attributes["name"].Value } elseif ($prop.Attributes["name"]) { $prop.Attributes["name"].Value } elseif ($prop.Attributes["path"]) { $prop.Attributes["path"].Value } elseif ($prop.Attributes["letter"]) { "Laufwerk $($prop.Attributes['letter'].Value):" } else { "GPP $itemType" }
                    $actionCode = if ($prop.Attributes["action"]) { $prop.Attributes["action"].Value } else { "U" }
                    $actionText = switch ($actionCode) { "C" { "Erstellen" } "U" { "Aktualisieren" } "R" { "Ersetzen" } "D" { "Loeschen" } default { $actionCode } }

                    $valParts = @("Aktion: $actionText")
                    if ($prop.Attributes["path"])       { $valParts += "Pfad: $($prop.Attributes['path'].Value)" }
                    if ($prop.Attributes["targetPath"]) { $valParts += "Ziel: $($prop.Attributes['targetPath'].Value)" }
                    if ($prop.Attributes["fromPath"])   { $valParts += "Quelle: $($prop.Attributes['fromPath'].Value)" }
                    if ($prop.Attributes["value"])      { $valParts += "Wert: $($prop.Attributes['value'].Value)" }
                    $cleanGppVal = $valParts -join " | "

                    $descLines = @("GPP OBJEKT: $itemName", "TYP: $itemType", "AKTION: $actionText", "--------------------------------------------------")
                    foreach ($att in $prop.Attributes) { $descLines += " - $($att.Name): $($att.Value)" }

                    $list.Add([PSCustomObject]@{
                        Scope     = $scope
                        Category  = $gppCategory
                        Name      = $itemName
                        Value     = $cleanGppVal
                        State     = $actionText
                        Supported = "Gruppenrichtlinien-Praeferenzen (GPP)"
                        Explain   = $descLines -join "`r`n"
                        GpoName   = $GpoDisplayName
                        Datum     = $dateStr
                        Uhrzeit   = $timeStr
                    })
                }
            }

            # 4. Kontorichtlinien (Kennwort, Sperrung, Kerberos)
            $accountNodes = $ext.SelectNodes(".//*[local-name()='Account']/*/* | .//*[local-name()='PasswordPolicy']/* | .//*[local-name()='AccountLockoutPolicy']/* | .//*[local-name()='AccountLockout']/* | .//*[local-name()='KerberosPolicy']/*")
            if ($accountNodes -and $accountNodes.Count -gt 0) {
                foreach ($accNode in $accountNodes) {
                    $tag = $accNode.LocalName.ToLower()
                    if ($tag -in @("passwordpolicy", "accountlockoutpolicy", "accountlockout", "kerberospolicy", "name", "display")) { continue }

                    $info = $script:knownAccountPolicies[$tag]
                    $policyName = if ($accNode.SelectSingleNode("./*[local-name()='Display']/*[local-name()='Name']")) {
                        $accNode.SelectSingleNode("./*[local-name()='Display']/*[local-name()='Name']").InnerText.Trim()
                    } elseif ($info) {
                        $info.Name
                    } else {
                        continue
                    }

                    $policyCat = if ($info) { $info.Cat } else { "Sicherheitseinstellungen / Kontorichtlinien" }

                    $valStr = ""
                    $stateStr = "Aktiviert"

                    if ($accNode.SelectSingleNode("./*[local-name()='Display']/*[local-name()='DisplayString']")) {
                        $valStr = $accNode.SelectSingleNode("./*[local-name()='Display']/*[local-name()='DisplayString']").InnerText.Trim()
                    } elseif ($accNode.SelectSingleNode("./*[local-name()='SettingNumber']")) {
                        $num = $accNode.SelectSingleNode("./*[local-name()='SettingNumber']").InnerText.Trim()
                        $valStr = if ($info -and $info.Unit) { "$num $($info.Unit)" } else { "$num" }
                    } elseif ($accNode.SelectSingleNode("./*[local-name()='SettingBoolean']")) {
                        $b = $accNode.SelectSingleNode("./*[local-name()='SettingBoolean']").InnerText.Trim().ToLower()
                        if ($b -eq "true" -or $b -eq "1") {
                            $valStr = "Enabled"
                            $stateStr = "Aktiviert"
                        } else {
                            $valStr = "Disabled"
                            $stateStr = "Deaktiviert"
                        }
                    } elseif (-not [string]::IsNullOrWhiteSpace($accNode.InnerText)) {
                        $valStr = $accNode.InnerText.Trim()
                    }

                    if (-not [string]::IsNullOrWhiteSpace($valStr)) {
                        $list.Add([PSCustomObject]@{
                            Scope     = $scope
                            Category  = $policyCat
                            Name      = $policyName
                            Value     = $valStr
                            State     = $stateStr
                            Supported = "Windows Kontorichtlinie"
                            Explain   = "Kontorichtlinie im Bereich $scope.`r`nRichtlinie: $policyName`r`nWert: $valStr"
                            GpoName   = $GpoDisplayName
                            Datum     = $dateStr
                            Uhrzeit   = $timeStr
                        })
                    }
                }
            }

            # 5. Sicherheitsoptionen (Sauberes Mapping statt roher KeyName/SettingNumber-Objekte)
            $secOptions = $ext.SelectNodes(".//*[local-name()='SecurityOptions']/* | .//*[local-name()='SecurityOption']")
            if ($secOptions -and $secOptions.Count -gt 0) {
                foreach ($sec in $secOptions) {
                    $secName = ""
                    $secVal = ""

                    # Fall A: Offizieller Display-Name vorhanden
                    if ($sec.SelectSingleNode("./*[local-name()='Display']/*[local-name()='Name']")) {
                        $secName = $sec.SelectSingleNode("./*[local-name()='Display']/*[local-name()='Name']").InnerText.Trim()
                        if ($sec.SelectSingleNode("./*[local-name()='Display']/*[local-name()='DisplayString']")) {
                            $secVal = $sec.SelectSingleNode("./*[local-name()='Display']/*[local-name()='DisplayString']").InnerText.Trim()
                        }
                    }
                    # Fall B: Verschachtelte Sicherheitsoption mit KeyName + SettingNumber / SettingBoolean
                    elseif ($sec.SelectSingleNode("./*[local-name()='KeyName']")) {
                        $kName = $sec.SelectSingleNode("./*[local-name()='KeyName']").InnerText.Trim()
                        if ($kName -match "NoLdapSigning") { $secName = "Domain controller: LDAP server signing requirements" }
                        elseif ($kName -match "NoLMHash") { $secName = "Network security: Do not store LAN Manager hash value on next password change" }
                        elseif ($kName -match "EnableSecuritySignature") { 
                            if ($kName -match "LanManServer") { $secName = "Microsoft network server: Digitally sign communications (always)" }
                            else { $secName = "Microsoft network client: Digitally sign communications (always)" }
                        }
                        elseif ($kName -match "RequireSecuritySignature") {
                            if ($kName -match "LanManServer") { $secName = "Microsoft network server: Digitally sign communications (if client agrees)" }
                            else { $secName = "Microsoft network client: Digitally sign communications (if client agrees)" }
                        }
                        else {
                            $secName = $kName.Split('\')[-1]
                        }

                        if ($sec.SelectSingleNode("./*[local-name()='SettingNumber']")) {
                            $secVal = $sec.SelectSingleNode("./*[local-name()='SettingNumber']").InnerText.Trim()
                        } elseif ($sec.SelectSingleNode("./*[local-name()='SettingBoolean']")) {
                            $bVal = $sec.SelectSingleNode("./*[local-name()='SettingBoolean']").InnerText.Trim().ToLower()
                            $secVal = if ($bVal -eq "true" -or $bVal -eq "1") { "Enabled" } else { "Disabled" }
                        }
                    }

                    if (-not [string]::IsNullOrWhiteSpace($secName) -and -not [string]::IsNullOrWhiteSpace($secVal)) {
                        $list.Add([PSCustomObject]@{
                            Scope     = $scope
                            Category  = "Sicherheitseinstellungen / Lokale Richtlinien / Sicherheitsoptionen"
                            Name      = $secName
                            Value     = $secVal
                            State     = "Konfiguriert"
                            Supported = "Windows Sicherheitsrichtlinie"
                            Explain   = "Sicherheitsoption im Bereich $scope.`r`nRichtlinie: $secName`r`nWert: $secVal"
                            GpoName   = $GpoDisplayName
                            Datum     = $dateStr
                            Uhrzeit   = $timeStr
                        })
                    }
                }
            }

            # 6. Audit-Richtlinien
            $auditOptions = $ext.SelectNodes(".//*[local-name()='AuditPolicy']/*")
            if ($auditOptions -and $auditOptions.Count -gt 0) {
                foreach ($aud in $auditOptions) {
                    $audName = if ($aud.SelectSingleNode("./*[local-name()='Display']/*[local-name()='Name']")) { $aud.SelectSingleNode("./*[local-name()='Display']/*[local-name()='Name']").InnerText.Trim() } else { $aud.LocalName }
                    $audVal = if ($aud.SelectSingleNode("./*[local-name()='Display']/*[local-name()='DisplayString']")) { $aud.SelectSingleNode("./*[local-name()='Display']/*[local-name()='DisplayString']").InnerText.Trim() } else { $aud.InnerText.Trim() }

                    if (-not [string]::IsNullOrWhiteSpace($audVal) -and $audVal -ne $audName) {
                        $list.Add([PSCustomObject]@{
                            Scope     = $scope
                            Category  = "Sicherheitseinstellungen / Lokale Richtlinien / Ueberwachungsrichtlinie"
                            Name      = $audName
                            Value     = $audVal
                            State     = "Konfiguriert"
                            Supported = "Windows Audit-Richtlinie"
                            Explain   = "Audit-Richtlinie: $audName = $audVal"
                            GpoName   = $GpoDisplayName
                            Datum     = $dateStr
                            Uhrzeit   = $timeStr
                        })
                    }
                }
            }

            # 7. Benutzerrechte (User Rights Assignment)
            $userRights = $ext.SelectNodes(".//*[local-name()='UserRightsAssignment']")
            if ($userRights -and $userRights.Count -gt 0) {
                foreach ($ur in $userRights) {
                    $rightName = if ($ur.SelectSingleNode("./*[local-name()='Name']")) { $ur.SelectSingleNode("./*[local-name()='Name']").InnerText.Trim() } else { $ur.LocalName }
                    $members = ($ur.SelectNodes(".//*[local-name()='Member']/*[local-name()='Name']").InnerText -join ", ")
                    if (-not [string]::IsNullOrWhiteSpace($members)) {
                        $list.Add([PSCustomObject]@{
                            Scope     = $scope
                            Category  = "Sicherheitseinstellungen / Zuweisen von Benutzerrechten"
                            Name      = $rightName
                            Value     = $members
                            State     = "Zugewiesen"
                            Supported = "Benutzerrechte-Richtlinie"
                            Explain   = "Zugewiesene Konten / Gruppen fuer '$rightName':`r`n$members"
                            GpoName   = $GpoDisplayName
                            Datum     = $dateStr
                            Uhrzeit   = $timeStr
                        })
                    }
                }
            }

            # 8. Registry-Einstellungen
            $regNodes = $ext.SelectNodes(".//*[local-name()='RegistrySetting']")
            if ($regNodes -and $regNodes.Count -gt 0) {
                foreach ($reg in $regNodes) {
                    $keyPath = if ($reg.SelectSingleNode("./*[local-name()='KeyPath']")) { $reg.SelectSingleNode("./*[local-name()='KeyPath']").InnerText.Trim() } else { "" }
                    $valName = if ($reg.SelectSingleNode("./*[local-name()='ValueName']")) { $reg.SelectSingleNode("./*[local-name()='ValueName']").InnerText.Trim() } else { "(Standard)" }
                    $valData = if ($reg.SelectSingleNode("./*[local-name()='Value']")) { $reg.SelectSingleNode("./*[local-name()='Value']").InnerText.Trim() } else { "" }
                    $regType = if ($reg.SelectSingleNode("./*[local-name()='Type']")) { $reg.SelectSingleNode("./*[local-name()='Type']").InnerText.Trim() } else { "REG_SZ" }

                    $list.Add([PSCustomObject]@{
                        Scope     = $scope
                        Category  = "Registry-Richtlinie"
                        Name      = "$keyPath\$valName"
                        Value     = "$valData ($regType)"
                        State     = "Aktiviert"
                        Supported = "Registry-Eintrag"
                        Explain   = "Registry: $keyPath\$valName = $valData ($regType)"
                        GpoName   = $GpoDisplayName
                        Datum     = $dateStr
                        Uhrzeit   = $timeStr
                    })
                }
            }

            # 9. User Configuration / RIS
            $risNodes = $ext.SelectNodes(".//*[local-name()='ClientInstallationWizard']/* | .//*[local-name()='RemoteInstallationServices']/*")
            if ($risNodes -and $risNodes.Count -gt 0) {
                foreach ($rn in $risNodes) {
                    $rName = if ($rn.SelectSingleNode("./*[local-name()='Policy']")) { 
                        $rn.SelectSingleNode("./*[local-name()='Policy']").InnerText.Trim() 
                    } elseif ($rn.Attributes["name"]) { 
                        $rn.Attributes["name"].Value 
                    } else { 
                        $rn.LocalName 
                    }

                    $rVal = if ($rn.SelectSingleNode("./*[local-name()='Setting']")) { 
                        $rn.SelectSingleNode("./*[local-name()='Setting']").InnerText.Trim() 
                    } elseif ($rn.SelectSingleNode("./*[local-name()='State']")) { 
                        $rn.SelectSingleNode("./*[local-name()='State']").InnerText.Trim() 
                    } else { 
                        $rn.InnerText.Trim() 
                    }

                    if (-not [string]::IsNullOrWhiteSpace($rVal) -and $rVal -ne $rName) {
                        $list.Add([PSCustomObject]@{
                            Scope     = $scope
                            Category  = "Windows-Einstellungen / Remote-Installationsdienste (RIS)"
                            Name      = $rName
                            Value     = $rVal
                            State     = $rVal
                            Supported = "Windows Remote Installation Services"
                            Explain   = "RIS-Einstellung im Bereich ${scope}: $rName = $rVal"
                            GpoName   = $GpoDisplayName
                            Datum     = $dateStr
                            Uhrzeit   = $timeStr
                        })
                    }
                }
            }

            # 10. EFS / Public Key Policies
            $efsNodes = $ext.SelectNodes(".//*[local-name()='EncryptingFileSystem' or local-name()='EFS' or local-name()='PublicKeyPolicies']")
            if ($efsNodes -and $efsNodes.Count -gt 0) {
                $list.Add([PSCustomObject]@{
                    Scope     = $scope
                    Category  = "Sicherheitseinstellungen / Richtlinien fuer oeffentliche Schluessel"
                    Name      = "Verschluesselndes Dateisystem (Encrypting File System - EFS)"
                    Value     = "Konfiguriert"
                    State     = "Aktiviert"
                    Supported = "Windows EFS"
                    Explain   = "EFS-Richtlinie im Bereich $scope konfiguriert."
                    GpoName   = $GpoDisplayName
                    Datum     = $dateStr
                    Uhrzeit   = $timeStr
                })
            }
        }
    }

    Parse-GpoSection $xml.GPO.Computer "Computer"
    Parse-GpoSection $xml.GPO.User "User"

    return $list
}