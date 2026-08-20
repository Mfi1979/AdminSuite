<#
================================================================================
 CORE: GLOBALE HILFSFUNKTIONEN, LDAP ENGINE & UNIVERSELLE SPALTENSORTIERUNG
================================================================================
#>

# Universelle Spaltensortierung für DataGridViews
function Enable-GridSorting {
    param(
        [Parameter(Mandatory=$true)]
        [System.Windows.Forms.DataGridView]$Grid,
        [scriptblock]$OnSortedCallback = $null
    )

    $sortState = @{
        LastColumn = ""
        Ascending  = $true
    }
    $Grid.Tag = $sortState

    $Grid.Add_ColumnHeaderMouseClick({
        param($sender, $e)
        
        $targetGrid = $sender
        if ($e.ColumnIndex -lt 0 -or $e.ColumnIndex -ge $targetGrid.Columns.Count) { return }
        $col = $targetGrid.Columns[$e.ColumnIndex]
        if (-not $col) { return }

        $propName = if ($col.DataPropertyName) { $col.DataPropertyName } else { $col.HeaderText }
        if (-not $propName) { return }

        $state = $targetGrid.Tag
        if (-not $state -or -not ($state -is [hashtable])) {
            $state = @{ LastColumn = ""; Ascending = $true }
            $targetGrid.Tag = $state
        }

        if ($state.LastColumn -eq $propName) {
            $state.Ascending = -not $state.Ascending
        } else {
            $state.LastColumn = $propName
            $state.Ascending = $true
        }

        $items = @($targetGrid.DataSource)
        if ($null -eq $items -or $items.Count -le 1) { return }

        $isAsc = $state.Ascending
        $sorted = $items | Sort-Object -Property @{
            Expression = {
                $val = $_.$propName
                if ($null -eq $val -or $val -eq "") { return "" }

                # Zahlen- und Einheiten-Erkennung (Tage, Minuten, GB, Kerne, Versuche etc.)
                if ($propName -match "(?i)tage|inaktiv|count|anzahl|precedence|priorität|versuche|attempts|alter|age|min|dauer|arbeitsspeicher|kerne") {
                    if ($val -eq "Nie" -or $val -eq "N/A" -or $val -eq "None" -or $val -eq "-") { 
                        return $(if ($isAsc) { [double]::MaxValue } else { [double]::MinValue }) 
                    }
                    if ($val.ToString() -match '^([0-9]+(\.[0-9]+)?)') {
                        return [double]$Matches[1]
                    }
                }

                # Datums-Erkennung (dd.MM.yyyy HH:mm:ss oder dd.MM.yyyy)
                if ($val.ToString() -match '^\d{2}\.\d{2}\.\d{4}') {
                    try {
                        return [datetime]::ParseExact($val.ToString().Trim(), @("dd.MM.yyyy HH:mm:ss", "dd.MM.yyyy HH:mm", "dd.MM.yyyy"), $null)
                    } catch { return $val }
                }

                return $val
            }
            Descending = (-not $isAsc)
        }

        $newArr = [System.Collections.ArrayList]::new()
        foreach ($it in $sorted) { [void]$newArr.Add($it) }

        $targetGrid.DataSource = $null
        $targetGrid.DataSource = $newArr

        foreach ($c in $targetGrid.Columns) {
            $c.HeaderCell.SortGlyphDirection = [System.Windows.Forms.SortOrder]::None
        }
        $targetGrid.Columns[$e.ColumnIndex].HeaderCell.SortGlyphDirection = `
            $(if ($isAsc) { [System.Windows.Forms.SortOrder]::Ascending } else { [System.Windows.Forms.SortOrder]::Descending })

        if ($OnSortedCallback) {
            & $OnSortedCallback $targetGrid
        }
    })
}

function Apply-StandardGridTheme {
    param(
        [Parameter(Mandatory=$true)]
        [System.Windows.Forms.DataGridView]$Grid,
        [switch]$EnableAlternatingRowColor,
        [switch]$DisableAutoSort
    )

    $theme = $script:UITheme

    # Basis-Verhalten
    $Grid.EnableHeadersVisualStyles = $false
    $Grid.BorderStyle               = [System.Windows.Forms.BorderStyle]::None
    $Grid.CellBorderStyle           = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
    $Grid.GridColor                 = $theme.GridLineColor
    $Grid.BackgroundColor           = [System.Drawing.Color]::White
    $Grid.RowHeadersVisible         = $false
    $Grid.AllowUserToResizeRows     = $false
    $Grid.SelectionMode             = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $Grid.ReadOnly                  = $true

    # 1. Spaltenüberschriften (Header) Styling
    $Grid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
    $Grid.ColumnHeadersHeight       = $theme.HeaderHeight
    
    $headerFont = New-Object System.Drawing.Font($theme.FontFamily, $theme.HeaderFontSize, $theme.HeaderFontStyle)
    $Grid.ColumnHeadersDefaultCellStyle.Font      = $headerFont
    $Grid.ColumnHeadersDefaultCellStyle.BackColor = $theme.HeaderBackColor
    $Grid.ColumnHeadersDefaultCellStyle.ForeColor = $theme.HeaderForeColor
    $Grid.ColumnHeadersDefaultCellStyle.Alignment = [System.Drawing.ContentAlignment]::MiddleLeft
    $Grid.ColumnHeadersDefaultCellStyle.Padding   = New-Object System.Windows.Forms.Padding(
        $theme.HeaderPaddingLeft, 0, $theme.HeaderPaddingRight, 0
    )

    # 2. Datenzeilen (Cells) Styling
    $cellFont = New-Object System.Drawing.Font($theme.FontFamily, $theme.CellFontSize, $theme.CellFontStyle)
    $Grid.RowTemplate.Height                   = $theme.RowHeight
    $Grid.DefaultCellStyle.Font                = $cellFont
    $Grid.DefaultCellStyle.BackColor           = $theme.RowBackColor
    $Grid.DefaultCellStyle.SelectionBackColor  = $theme.SelectionBackColor
    $Grid.DefaultCellStyle.SelectionForeColor  = $theme.SelectionForeColor
    $Grid.DefaultCellStyle.Alignment          = [System.Drawing.ContentAlignment]::MiddleLeft
    $Grid.DefaultCellStyle.Padding            = New-Object System.Windows.Forms.Padding(
        $theme.CellPaddingLeft, $theme.CellPaddingTop, $theme.CellPaddingRight, $theme.CellPaddingBottom
    )

    # Optional: Alternierende Zeilenfarben
    if ($EnableAlternatingRowColor) {
        $Grid.AlternatingRowsDefaultCellStyle.BackColor          = $theme.RowAltBackColor
        $Grid.AlternatingRowsDefaultCellStyle.SelectionBackColor = $theme.SelectionBackColor
        $Grid.AlternatingRowsDefaultCellStyle.SelectionForeColor = $theme.SelectionForeColor
    }

    # Automatische Spaltensortierung für jedes DataGridView aktivieren
    if (-not $DisableAutoSort) {
        Enable-GridSorting -Grid $Grid
    }
}

function Assert-DomainJoined {
    $isDomain = (Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue).PartOfDomain
    if (-not $isDomain) {
        [System.Windows.Forms.MessageBox]::Show((Get-Text "ErrNoDomain"), "Info", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return $false
    }
    return $true
}

function Get-DomainDN {
    try {
        $rootDSE = [ADSI]"LDAP://RootDSE"
        return $rootDSE.defaultNamingContext.ToString()
    } catch {
        return $null
    }
}

function Search-NativeLdap {
    param (
        [string]$LdapFilter = "(objectClass=*)",
        [string[]]$PropertiesToLoad = @("name"),
        [string]$Server = $null
    )
    try {
        $entry = if ($Server) { 
            New-Object System.DirectoryServices.DirectoryEntry("LDAP://$Server") 
        } else { 
            New-Object System.DirectoryServices.DirectoryEntry 
        }
        $searcher = New-Object System.DirectoryServices.DirectorySearcher($entry)
        $searcher.Filter = $LdapFilter
        $searcher.PageSize = 1000
        foreach ($prop in $PropertiesToLoad) { [void]$searcher.PropertiesToLoad.Add($prop) }
        return $searcher.FindAll()
    }
    catch {
        return @()
    }
}

function Convert-LcidToLanguageName {
    param([object]$lcid)
    if ($null -eq $lcid -or $lcid -eq "" -or $lcid -eq 0) { return "Neutral / Multilingual" }
    try {
        $id = [int]$lcid
        switch ($id) {
            1031 { return "DE (de-DE)" }
            2055 { return "DE (de-CH)" }
            3079 { return "DE (de-AT)" }
            1033 { return "EN (en-US)" }
            2057 { return "EN (en-GB)" }
            4105 { return "EN (en-CA)" }
            3081 { return "EN (en-AU)" }
            1036 { return "FR (fr-FR)" }
            1034 { return "ES (es-ES)" }
            1040 { return "IT (it-IT)" }
            default {
                $cult = [System.Globalization.CultureInfo]::GetCultureInfo($id)
                return "$($cult.TwoLetterISOLanguageName.ToUpper()) ($($cult.Name))"
            }
        }
    }
    catch {
        return "Neutral / Multilingual"
    }
}

function Analyze-OSSupportDetails {
    param(
        [string]$OSName,
        [string]$OSVersion,
        [string]$OSServicePack
    )

    $status        = "Supported"
    $eolDateString = "Unbekannt"
    $clientVersion = "N/A"
    $buildNumber   = "Unbekannt"
    $targetDate    = $null

    # 1. Build-Nummer extrahieren
    if ($OSVersion -match "(\d{5})") {
        $buildNumber = $Matches[1]
    } elseif ($OSVersion -match "10\.0\.(\d+)") {
        $buildNumber = $Matches[1]
    } elseif ($OSVersion -and $OSVersion -ne "Unspecified OS") {
        $buildNumber = $OSVersion
    }

    # Fallback: Build-Nummer aus OS-Namen erkennen, falls OSVersion im AD leer ist
    if ($buildNumber -eq "Unbekannt" -or [string]::IsNullOrWhiteSpace($buildNumber)) {
        if ($OSName -match "24H2") { $buildNumber = "26100" }
        elseif ($OSName -match "23H2") { $buildNumber = "22631" }
        elseif ($OSName -match "22H2" -and $OSName -like "*Windows 11*") { $buildNumber = "22621" }
        elseif ($OSName -match "21H2" -and $OSName -like "*Windows 11*") { $buildNumber = "22000" }
        elseif ($OSName -match "22H2" -and $OSName -like "*Windows 10*") { $buildNumber = "19045" }
        elseif ($OSName -match "2021" -or ($OSName -like "*LTSC*" -and $OSName -match "21H2")) { $buildNumber = "19044" }
        elseif ($OSName -match "2019" -or ($OSName -like "*LTSC*" -and $OSName -match "1809")) { $buildNumber = "17763" }
        elseif ($OSName -match "2016" -or ($OSName -like "*LTSB*" -and $OSName -match "1607")) { $buildNumber = "14393" }
        elseif ($OSName -match "2015" -or ($OSName -like "*LTSB*" -and $OSName -match "1507")) { $buildNumber = "10240" }
    }

    $isLTSC = ($OSName -like "*LTSC*" -or $OSName -like "*LTSB*" -or $OSServicePack -like "*LTSC*" -or $OSServicePack -like "*LTSB*")

    if ($buildNumber -ne "Unbekannt") {
        $buildInt = 0
        [int]::TryParse($buildNumber, [ref]$buildInt) | Out-Null

        switch ($buildInt) {
            # Windows 11
            26200 { $clientVersion = "25H2" }
            26100 { 
                if ($OSName -like "*Server*") { $clientVersion = "Server 2025" }
                elseif ($isLTSC) { $clientVersion = "24H2 / LTSC 2024" }
                else { $clientVersion = "24H2" }
            }
            22631 { $clientVersion = "23H2" }
            22621 { $clientVersion = "22H2" }
            22000 { $clientVersion = "21H2" }

            # Windows 10
            19045 { $clientVersion = "22H2" }
            19044 { if ($isLTSC) { $clientVersion = "21H2 / LTSC 2021" } else { $clientVersion = "21H2" } }
            19043 { $clientVersion = "21H1" }
            19042 { $clientVersion = "20H2" }
            19041 { $clientVersion = "2004" }
            18363 { $clientVersion = "1909" }
            17763 { 
                if ($OSName -like "*Server*") { $clientVersion = "Server 2019" } 
                else { $clientVersion = if ($isLTSC) { "1809 / LTSC 2019" } else { "1809" } } 
            }
            14393 { 
                if ($OSName -like "*Server*") { $clientVersion = "Server 2016" } 
                else { $clientVersion = if ($isLTSC) { "1607 / LTSB 2016" } else { "1607" } } 
            }
            10240 { $clientVersion = "1507 / LTSB 2015" }

            # Server
            20348 { $clientVersion = "Server 2022" }

            default {
                if ($buildInt -gt 26200) { $clientVersion = "Insider Build" }
                elseif ($OSVersion) { $clientVersion = $OSVersion }
                else { $clientVersion = "Build $buildNumber" }
            }
        }
    } else {
        if ($OSName -like "*LTSC 2021*") { $clientVersion = "21H2 / LTSC 2021" }
        elseif ($OSName -like "*LTSC 2019*") { $clientVersion = "1809 / LTSC 2019" }
        elseif ($OSName -like "*LTSB 2016*") { $clientVersion = "1607 / LTSB 2016" }
        elseif ($OSName -like "*Server 2022*") { $clientVersion = "Server 2022" }
        elseif ($OSName -like "*Server 2019*") { $clientVersion = "Server 2019" }
        elseif ($OSName -like "*Server 2016*") { $clientVersion = "Server 2016" }
        elseif ($OSName -like "*Server 2012*") { $clientVersion = "Server 2012" }
        elseif ($OSServicePack) { $clientVersion = $OSServicePack }
        elseif ($OSVersion) { $clientVersion = $OSVersion }
    }

    $isEnterpriseOrEdu = ($OSName -like "*Enterprise*" -or $OSName -like "*Education*")
    $isLTSCOrLTSB      = ($isLTSC -or $clientVersion -like "*LTSC*" -or $clientVersion -like "*LTSB*")

    # 3. EOL-Datumsberechnung
    if ($OSName -like "*Windows 11*") {
        if ($clientVersion -like "*21H2*" -or $buildNumber -eq "22000") {
            $dateStr = if ($isEnterpriseOrEdu) { "08.10.2024" } else { "10.10.2023" }
            $eolDateString = $dateStr; $targetDate = [datetime]::ParseExact($dateStr, "dd.MM.yyyy", $null)
        } 
        elseif ($clientVersion -like "*22H2*" -or $buildNumber -eq "22621") {
            $dateStr = if ($isEnterpriseOrEdu) { "14.10.2025" } else { "08.10.2024" }
            $eolDateString = $dateStr; $targetDate = [datetime]::ParseExact($dateStr, "dd.MM.yyyy", $null)
        } 
        elseif ($clientVersion -like "*23H2*" -or $buildNumber -eq "22631") {
            $dateStr = if ($isEnterpriseOrEdu) { "10.11.2026" } else { "11.11.2025" }
            $eolDateString = $dateStr; $targetDate = [datetime]::ParseExact($dateStr, "dd.MM.yyyy", $null)
        } 
        elseif ($clientVersion -like "*24H2*" -or $buildNumber -eq "26100") {
            if ($isLTSCOrLTSB) {
                $dateStr = if ($OSName -like "*IoT*") { "10.10.2034" } else { "09.10.2029" }
            } else {
                $dateStr = if ($isEnterpriseOrEdu) { "12.10.2027" } else { "13.10.2026" }
            }
            $eolDateString = $dateStr; $targetDate = [datetime]::ParseExact($dateStr, "dd.MM.yyyy", $null)
        } 
        elseif ($clientVersion -like "*25H2*" -or $buildNumber -eq "26200") {
            $dateStr = if ($isEnterpriseOrEdu) { "10.10.2028" } else { "12.10.2027" }
            $eolDateString = $dateStr; $targetDate = [datetime]::ParseExact($dateStr, "dd.MM.yyyy", $null)
        } 
        else {
            $eolDateString = "Support aktiv"
        }
    }
    elseif ($OSName -like "*Windows 10*") {
        if ($isLTSCOrLTSB) {
            if ($buildNumber -eq "19044" -or $clientVersion -like "*2021*") {
                $dateStr = if ($OSName -like "*IoT*") { "13.01.2032" } else { "12.01.2027" }
                $eolDateString = $dateStr; $targetDate = [datetime]::ParseExact($dateStr, "dd.MM.yyyy", $null)
            }
            elseif ($buildNumber -eq "17763" -or $clientVersion -like "*2019*") {
                $eolDateString = "09.01.2029"; $targetDate = [datetime]::ParseExact("09.01.2029", "dd.MM.yyyy", $null)
            }
            elseif ($buildNumber -eq "14393" -or $clientVersion -like "*2016*") {
                $eolDateString = "13.10.2026"; $targetDate = [datetime]::ParseExact("13.10.2026", "dd.MM.yyyy", $null)
            }
            elseif ($buildNumber -eq "10240" -or $clientVersion -like "*2015*") {
                $eolDateString = "14.10.2025"; $targetDate = [datetime]::ParseExact("14.10.2025", "dd.MM.yyyy", $null)
            }
            else {
                $eolDateString = "12.01.2027"; $targetDate = [datetime]::ParseExact("12.01.2027", "dd.MM.yyyy", $null)
            }
        } else {
            $eolDateString = "14.10.2025"; $targetDate = [datetime]::ParseExact("14.10.2025", "dd.MM.yyyy", $null)
        }
    }
    elseif ($OSName -like "*Windows 7*" -or $OSName -like "*Windows 8*" -or $OSName -like "*Windows XP*" -or $OSName -like "*Windows Vista*") {
        $eolDateString = "Ausgelaufen"
        $status = "Out of Support / EOL"
    }
    elseif ($OSName -like "*Server 2003*" -or $OSName -like "*Server 2008*") {
        $eolDateString = "Ausgelaufen"
        $status = "Out of Support / EOL"
    }
    elseif ($OSName -like "*Server 2012*") {
        $eolDateString = "10.10.2023"; $targetDate = [datetime]::ParseExact("10.10.2023", "dd.MM.yyyy", $null)
    }
    elseif ($OSName -like "*Server 2016*") {
        $eolDateString = "12.01.2027"; $targetDate = [datetime]::ParseExact("12.01.2027", "dd.MM.yyyy", $null)
    }
    elseif ($OSName -like "*Server 2019*") {
        $eolDateString = "09.01.2029"; $targetDate = [datetime]::ParseExact("09.01.2029", "dd.MM.yyyy", $null)
    }
    elseif ($OSName -like "*Server 2022*") {
        $eolDateString = "14.10.2031"; $targetDate = [datetime]::ParseExact("14.10.2031", "dd.MM.yyyy", $null)
    }
    elseif ($OSName -like "*Server 2025*") {
        $eolDateString = "14.10.2034"; $targetDate = [datetime]::ParseExact("14.10.2034", "dd.MM.yyyy", $null)
    }

    if ($targetDate) {
        $today = Get-Date
        if ($today -ge $targetDate) {
            $status = "Out of Support / EOL"
        } elseif ($today.AddMonths(12) -ge $targetDate) {
            $status = "Near EOL"
        } else {
            $status = "Supported"
        }
    }

    return [PSCustomObject]@{
        ClientVersion = $clientVersion
        BuildNumber   = $buildNumber
        Status        = $status
        EOLDate       = $eolDateString
    }
}
