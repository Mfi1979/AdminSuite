# =========================================================================
# GpoParser.ps1 - Dual-Engine Parser (HTML-GPMC-Paritaet + XML-Fallback)
# =========================================================================

function Clean-GpoHtmlText ([string]$rawText) {
    if ([string]::IsNullOrWhiteSpace($rawText)) { return "" }
    $text = [System.Net.WebUtility]::HtmlDecode($rawText)
    $text = $text -replace '<br\s*/?>', ' '
    $text = $text -replace '<[^>]+>', ''
    $text = $text -replace '\s+', ' '
    return $text.Trim()
}

function Get-ParsedGpoSettings {
    param(
        [string]$GpoId,
        [string]$GpoDisplayName
    )

    $dateStr = Get-Date -Format "yyyyMMdd"
    $timeStr = Get-Date -Format "HHmm"
    $list = [System.Collections.Generic.List[PSCustomObject]]::new()

    function Add-ParsedItem ($scope, $cat, $name, $val, $state, $supp, $exp) {
        if ([string]::IsNullOrWhiteSpace($name) -or [string]::IsNullOrWhiteSpace($val)) { return }
        $cName = "$name".Trim()
        $cVal  = "$val".Trim()

        $exists = $list | Where-Object { $_.Scope -eq $scope -and $_.Name -eq $cName -and $_.Value -eq $cVal }
        if (-not $exists) {
            $list.Add([PSCustomObject]@{
                Scope     = $scope
                Category  = $cat
                Name      = $cName
                Value     = $cVal
                State     = $state
                Supported = $supp
                Explain   = $exp
                GpoName   = $GpoDisplayName
                Datum     = $dateStr
                Uhrzeit   = $timeStr
            })
        }
    }

    # =========================================================================
    # ENGINE 1: Nativer GPMC-HTML-Bericht (100% Paritaet mit Microsoft-Berichten)
    # =========================================================================
    $htmlSuccess = $false
    try {
        $html = Get-GPOReport -Guid $GpoId -ReportType Html -ErrorAction Stop
        if (-not [string]::IsNullOrWhiteSpace($html)) {

            $scopePattern = '(?s)<div class="he0_expanded"><span class="sectionTitle"[^>]*>(?<scopeName>(?:Computer|User) Configuration)[^<]*</span>(?<scopeBody>[\s\S]*?)(?=(?:<div class="he0_expanded"><span class="sectionTitle"[^>]*>(?:Computer|User|General)|</body>|$))'
            $scopeMatches = [regex]::Matches($html, $scopePattern)

            foreach ($sm in $scopeMatches) {
                $scope = if ($sm.Groups["scopeName"].Value -match "Computer") { "Computer" } else { "User" }
                $scopeBody = $sm.Groups["scopeBody"].Value

                $tokenPattern = '(?<header><div class="(?<hCls>he[1-5][a-z_]*)"><span class="sectionTitle"[^>]*>(?<hTitle>[^<]+)</span>)|(?<table><table(?<tAttr>[^>]*)class="(?<tCls>[^"]*)"[^>]*>(?<tBody>[\s\S]*?)</table>)'
                $tokens = [regex]::Matches($scopeBody, $tokenPattern)

                $catLevels = @{}
                $lastItemRef = $null

                foreach ($token in $tokens) {
                    if ($token.Groups["header"].Success) {
                        $hCls = $token.Groups["hCls"].Value
                        $hTitle = Clean-GpoHtmlText $token.Groups["hTitle"].Value
                        $level = if ($hCls -match "he(\d)") { [int]$matches[1] } else { 1 }

                        $catLevels[$level] = $hTitle
                        for ($i = $level + 1; $i -le 6; $i++) { $catLevels.Remove($i) }
                    }
                    elseif ($token.Groups["table"].Success) {
                        $tCls = $token.Groups["tCls"].Value
                        $tBody = $token.Groups["tBody"].Value

                        if ($tCls -match "subtable_frame" -and $null -ne $lastItemRef) {
                            $subRows = [regex]::Matches($tBody, '<tr[^>]*>\s*<td[^>]*>(?<k>[\s\S]*?)</td>\s*<td[^>]*>(?<v>[\s\S]*?)</td>\s*</tr>')
                            $paramParts = @()
                            foreach ($sr in $subRows) {
                                $k = Clean-GpoHtmlText $sr.Groups["k"].Value
                                $v = Clean-GpoHtmlText $sr.Groups["v"].Value
                                if ($k -and $v) { $paramParts += "${k}: $v" }
                            }
                            if ($paramParts.Count -gt 0) {
                                $lastItemRef.Value = "$($lastItemRef.Value) | " + ($paramParts -join " | ")
                            }
                            continue
                        }

                        $catPath = ($catLevels.Keys | Sort-Object | ForEach-Object { $catLevels[$_] } | Where-Object { $_ -notin @("Policies", "General") }) -join " / "
                        if ([string]::IsNullOrWhiteSpace($catPath)) { $catPath = "Richtlinien" }

                        $rows = [regex]::Matches($tBody, '<tr[^>]*>(?<rContent>[\s\S]*?)</tr>')
                        foreach ($row in $rows) {
                            $rContent = $row.Groups["rContent"].Value
                            if ($rContent -match '<th\b') { continue }

                            $cells = [regex]::Matches($rContent, '<td[^>]*>(?<cVal>[\s\S]*?)</td>')
                            if ($cells.Count -ge 2) {
                                $c0 = $cells[0].Groups["cVal"].Value
                                $c1 = $cells[1].Groups["cVal"].Value

                                $pName = ""
                                $pVal  = Clean-GpoHtmlText $c1
                                $pSupported = "Windows Gruppenrichtlinie"
                                $pExplain   = "Keine Erklaerung hinterlegt."

                                if ($c0 -match 'gpmc_settingName="(?<sName>[^"]*)"') {
                                    $pName = [System.Net.WebUtility]::HtmlDecode($matches["sName"])
                                    if ($c0 -match 'gpmc_settingDescription="(?<sDesc>[^"]*)"') {
                                        $pExplain = Clean-GpoHtmlText $matches["sDesc"]
                                    }
                                    if ($c0 -match 'gpmc_supported="(?<sSupp>[^"]*)"') {
                                        $pSupported = Clean-GpoHtmlText $matches["sSupp"]
                                    }
                                } else {
                                    $pName = Clean-GpoHtmlText $c0
                                }

                                if ($cells.Count -ge 4 -and $catPath -match "Encrypting File System|Certificates") {
                                    $issuedTo = Clean-GpoHtmlText $cells[0].Groups["cVal"].Value
                                    $expDate  = Clean-GpoHtmlText $cells[2].Groups["cVal"].Value
                                    $purpose  = Clean-GpoHtmlText $cells[3].Groups["cVal"].Value
                                    $pName = "Certificates (Encrypting File System)"
                                    $pVal  = "Issued To: $issuedTo | Intended Purposes: $purpose | Expiration Date: $expDate"
                                }

                                if ([string]::IsNullOrWhiteSpace($pName) -or [string]::IsNullOrWhiteSpace($pVal)) { continue }

                                $pState = switch -Regex ($pVal) {
                                    "^(Disabled|Deaktiviert)$" { "Deaktiviert" }
                                    "^(Enabled|Aktiviert)"     { "Aktiviert" }
                                    default                    { "Aktiviert" }
                                }

                                $itemObj = [PSCustomObject]@{
                                    Scope     = $scope
                                    Category  = $catPath
                                    Name      = $pName
                                    Value     = $pVal
                                    State     = $pState
                                    Supported = $pSupported
                                    Explain   = $pExplain
                                    GpoName   = $GpoDisplayName
                                    Datum     = $dateStr
                                    Uhrzeit   = $timeStr
                                }

                                $exists = $list | Where-Object { $_.Scope -eq $scope -and $_.Name -eq $pName -and $_.Value -eq $pVal }
                                if (-not $exists) {
                                    $list.Add($itemObj)
                                    $lastItemRef = $itemObj
                                }
                            }
                        }
                    }
                }
            }

            if ($list.Count -gt 0) {
                $htmlSuccess = $true
            }
        }
    } catch {}

    if ($htmlSuccess) {
        return $list
    }

    # =========================================================================
    # ENGINE 2: XML-Fallback
    # =========================================================================
    try {
        [xml]$xml = Get-GPOReport -Guid $GpoId -ReportType Xml -ErrorAction Stop

        $parseXmlSection = {
            param($sectionNode, $scope)
            if ($null -eq $sectionNode -or -not $sectionNode.ExtensionData) { return }

            foreach ($ext in $sectionNode.ExtensionData.Extension) {
                $extCategory = if ($ext.Name) { $ext.Name } else { "Erweiterung" }

                $policies = $ext.SelectNodes(".//*[local-name()='Policy']")
                if ($policies) {
                    foreach ($p in $policies) {
                        $pName = if ($p.SelectSingleNode("./*[local-name()='Name']")) { $p.SelectSingleNode("./*[local-name()='Name']").InnerText.Trim() } else { $p.Name }
                        $rawState = if ($p.SelectSingleNode("./*[local-name()='State']")) { $p.SelectSingleNode("./*[local-name()='State']").InnerText.Trim() } else { "Enabled" }
                        $pState = switch ($rawState) { "Enabled" { "Aktiviert" } "Disabled" { "Deaktiviert" } default { $rawState } }
                        $pCat = if ($p.SelectSingleNode("./*[local-name()='Category']")) { $p.SelectSingleNode("./*[local-name()='Category']").InnerText.Trim() } else { $extCategory }
                        $pSupported = if ($p.SelectSingleNode("./*[local-name()='Supported']")) { $p.SelectSingleNode("./*[local-name()='Supported']").InnerText.Trim() } else { "Keine Angabe" }
                        $pExplain = if ($p.SelectSingleNode("./*[local-name()='Explain']")) { $p.SelectSingleNode("./*[local-name()='Explain']").InnerText.Trim() } else { "Keine Erklaerung hinterlegt." }

                        $paramVals = @()
                        foreach ($child in $p.ChildNodes) {
                            if ($child.LocalName -in @("Name", "State", "Explain", "Supported", "Category")) { continue }
                            $valNodes = $child.SelectNodes(".//*[local-name()='Value' or local-name()='Data']")
                            if ($valNodes) {
                                foreach ($vn in $valNodes) { if (-not [string]::IsNullOrWhiteSpace($vn.InnerText)) { $paramVals += $vn.InnerText.Trim() } }
                            }
                        }
                        $cleanVal = if ($paramVals.Count -gt 0) { $paramVals -join " | " } else { $pState }
                        Add-ParsedItem $scope $pCat $pName $cleanVal $pState $pSupported $pExplain
                    }
                }

                $gppNodes = $ext.SelectNodes(".//*[local-name()='Properties']")
                if ($gppNodes) {
                    foreach ($prop in $gppNodes) {
                        $parent = $prop.ParentNode
                        $itemName = if ($parent.Attributes["name"]) { $parent.Attributes["name"].Value } else { "GPP $($parent.LocalName)" }
                        $action = if ($prop.Attributes["action"]) { $prop.Attributes["action"].Value } else { "U" }
                        $parts = @("Aktion: $action")
                        if ($prop.Attributes["path"]) { $parts += "Pfad: $($prop.Attributes['path'].Value)" }
                        if ($prop.Attributes["value"]) { $parts += "Wert: $($prop.Attributes['value'].Value)" }
                        Add-ParsedItem $scope "Praeferenzen (GPP) / $($parent.LocalName)" $itemName ($parts -join " | ") "Aktiviert" "GPP" "GPP $($parent.LocalName)"
                    }
                }

                $regNodes = $ext.SelectNodes(".//*[local-name()='RegistrySetting']")
                if ($regNodes) {
                    foreach ($reg in $regNodes) {
                        $kPath = if ($reg.SelectSingleNode("./*[local-name()='KeyPath']")) { $reg.SelectSingleNode("./*[local-name()='KeyPath']").InnerText.Trim() } else { "" }
                        $vName = if ($reg.SelectSingleNode("./*[local-name()='ValueName']")) { $reg.SelectSingleNode("./*[local-name()='ValueName']").InnerText.Trim() } else { "(Standard)" }
                        $vData = if ($reg.SelectSingleNode("./*[local-name()='Value']")) { $reg.SelectSingleNode("./*[local-name()='Value']").InnerText.Trim() } else { "" }
                        Add-ParsedItem $scope "Registry-Richtlinie" "${kPath}\${vName}" $vData "Aktiviert" "Registry" "Registry: ${kPath}\${vName}"
                    }
                }
            }
        }

        & $parseXmlSection $xml.GPO.Computer "Computer"
        & $parseXmlSection $xml.GPO.User "User"
    } catch {}

    return $list
}