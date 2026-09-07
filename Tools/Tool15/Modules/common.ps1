# =========================================================================
# Common.ps1 - Basiskonfiguration, Typen & Referenzdaten
# =========================================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.DirectoryServices
Import-Module GroupPolicy -ErrorAction Stop

try { [System.Windows.Forms.Application]::EnableVisualStyles() } catch {}

$script:ToolVersion      = "v1.8.8"
$script:StandardDdpGuid  = "31B2F340-016D-11D2-945F-00C04FB984F9"
$script:StandardDdcpGuid = "6AC1786C-016F-11D2-945F-00C04FB984F9"
$script:DdpBaselineName  = "[Referenz] Microsoft Default Domain Policy (Standard-Werte)"

# Spaltensortierung für DataGridViews
function Enable-GridSorting {
    param([System.Windows.Forms.DataGridView]$Grid)

    $Grid.Add_ColumnHeaderMouseClick({
        param($sender, $e)
        if ($script:isClosing -or $Grid.IsDisposed) { return }

        $targetGrid = $sender
        $colProp = $targetGrid.Columns[$e.ColumnIndex].DataPropertyName
        if (-not $colProp) { $colProp = $targetGrid.Columns[$e.ColumnIndex].HeaderText }

        $asc = if ($targetGrid.Tag -and $targetGrid.Tag.Column -eq $colProp) { -not $targetGrid.Tag.Ascending } else { $true }
        $targetGrid.Tag = @{ Column = $colProp; Ascending = $asc }

        $data = @($targetGrid.DataSource)
        if ($null -eq $data -or $data.Count -le 1) { return }

        $sorted = $data | Sort-Object -Property @{
            Expression = {
                $val = $_.$colProp
                if ($null -eq $val) { return "" }
                if ($colProp -eq "Link-Anzahl" -and ($val -as [int])) { return [int]$val }
                return $val
            }
            Descending = (-not $asc)
        }

        $arr = [System.Collections.ArrayList]::new()
        foreach ($item in $sorted) { [void]$arr.Add($item) }
        $targetGrid.DataSource = $arr

        foreach ($col in $targetGrid.Columns) { $col.HeaderCell.SortGlyphDirection = [System.Windows.Forms.SortOrder]::None }
        $targetGrid.Columns[$e.ColumnIndex].HeaderCell.SortGlyphDirection = if ($asc) { [System.Windows.Forms.SortOrder]::Ascending } else { [System.Windows.Forms.SortOrder]::Descending }
    })
}

# 14-Punkte Microsoft Standard-Baseline (Werkszustand)
function Get-DefaultDomainPolicyBaseline {
    $dateStr = Get-Date -Format "yyyyMMdd"
    $timeStr = Get-Date -Format "HHmm"
    $baseline = [System.Collections.Generic.List[PSCustomObject]]::new()

    $addBase = {
        param($cat, $name, $state, $val, $explain)
        $baseline.Add([PSCustomObject]@{
            Scope     = "Computer"
            Category  = $cat
            Name      = $name
            Value     = $val
            State     = $state
            Supported = "Windows 2000 und hoeher"
            Explain   = $explain
            GpoName   = $script:DdpBaselineName
            Datum     = $dateStr
            Uhrzeit   = $timeStr
        })
    }

    & $addBase "Sicherheitseinstellungen / Kontorichtlinien / Kennwortrichtlinie" "Kennwortchronik erzwingen (Password History)" "Aktiviert" "24 gespeicherte Kennwoerter" "Microsoft Standard: 24 Kennwoerter."
    & $addBase "Sicherheitseinstellungen / Kontorichtlinien / Kennwortrichtlinie" "Maximales Kennwortalter (Maximum Password Age)" "Aktiviert" "42 Tage" "Microsoft Standard: 42 Tage."
    & $addBase "Sicherheitseinstellungen / Kontorichtlinien / Kennwortrichtlinie" "Minimales Kennwortalter (Minimum Password Age)" "Aktiviert" "1 Tag" "Microsoft Standard: 1 Tag."
    & $addBase "Sicherheitseinstellungen / Kontorichtlinien / Kennwortrichtlinie" "Mindestkennwortlaenge (Minimum Password Length)" "Aktiviert" "7 Zeichen" "Microsoft Standard: 7 Zeichen."
    & $addBase "Sicherheitseinstellungen / Kontorichtlinien / Kennwortrichtlinie" "Kennwort muss Komplexitaetsanforderungen entsprechen" "Aktiviert" "Aktiviert" "Microsoft Standard: Aktiviert."
    & $addBase "Sicherheitseinstellungen / Kontorichtlinien / Kennwortrichtlinie" "Kennwoerter mit umkehrbarer Verschluesselung speichern" "Deaktiviert" "Deaktiviert" "Microsoft Standard: Deaktiviert."
    & $addBase "Sicherheitseinstellungen / Kontorichtlinien / Kontosperrungsrichtlinie" "Kontosperrungsschwelle (Account Lockout Threshold)" "Deaktiviert" "0 ungueltige Anmeldeversuche (Keine Kontosperrung)" "Microsoft Standard: 0."
    & $addBase "Sicherheitseinstellungen / Kontorichtlinien / Kontosperrungsrichtlinie" "Kontosperrdauer (Account Lockout Duration)" "Nicht definiert" "Nicht definiert" "Microsoft Standard: Nicht definiert."
    & $addBase "Sicherheitseinstellungen / Kontorichtlinien / Kontosperrungsrichtlinie" "Zuruecksetzungsdauer des Kontosperrzaehlers" "Nicht definiert" "Nicht definiert" "Microsoft Standard: Nicht definiert."
    & $addBase "Sicherheitseinstellungen / Kontorichtlinien / Kerberos-Richtlinie" "Maximale Toleranz fuer Synchronisierung der Computeruhr" "Aktiviert" "5 Minuten" "Microsoft Standard: 5 Minuten."
    & $addBase "Sicherheitseinstellungen / Kontorichtlinien / Kerberos-Richtlinie" "Maximale Lebensdauer fuer Benutzerticket (TGT)" "Aktiviert" "10 Stunden" "Microsoft Standard: 10 Stunden."
    & $addBase "Sicherheitseinstellungen / Kontorichtlinien / Kerberos-Richtlinie" "Maximale Lebensdauer fuer Serviceticket" "Aktiviert" "600 Minuten" "Microsoft Standard: 600 Minuten."
    & $addBase "Sicherheitseinstellungen / Kontorichtlinien / Kerberos-Richtlinie" "Maximale Lebensdauer fuer Benutzer-Ticket-Erneuerung" "Aktiviert" "7 Tage" "Microsoft Standard: 7 Tage."
    & $addBase "Sicherheitseinstellungen / Kontorichtlinien / Kerberos-Richtlinie" "Benutzeranmeldeeinschraenkungen erzwingen" "Aktiviert" "Aktiviert" "Microsoft Standard: Aktiviert."

    return $baseline
}

# Normalisierung für zweisprachigen Diff (DE / EN)
function Get-NormalizedPolicyKey ($item) {
    $n = "$($item.Name)".ToLower()
    if ($n -match "passwordhistory|kennwortchronik|password history") { return "$($item.Scope)|PasswordHistory" }
    if ($n -match "maximumpasswordage|maximales kennwortalter|maximum password age") { return "$($item.Scope)|MaxPasswordAge" }
    if ($n -match "minimumpasswordage|minimales kennwortalter|minimum password age") { return "$($item.Scope)|MinPasswordAge" }
    if ($n -match "minpasswordlength|mindestkennwort|minimum password length") { return "$($item.Scope)|MinPasswordLength" }
    if ($n -match "passwordcomplexity|komplexit|complexity requirements") { return "$($item.Scope)|PasswordComplexity" }
    if ($n -match "cleartextpassword|umkehrbar|reversible encryption") { return "$($item.Scope)|ReversibleEncryption" }
    if ($n -match "lockoutbadcount|kontosperrungsschwelle|lockout threshold") { return "$($item.Scope)|LockoutThreshold" }
    if ($n -match "lockoutduration|kontosperrdauer|lockout duration") { return "$($item.Scope)|LockoutDuration" }
    if ($n -match "resetlockoutcount|zuruecksetzungsdauer|reset account lockout") { return "$($item.Scope)|ResetLockoutCount" }
    if ($n -match "maxclockskew|synchronisierung der computeruhr|clock synchronization") { return "$($item.Scope)|MaxClockSkew" }
    if ($n -match "maxticketage|lebensdauer.*benutzerticket|lifetime for user ticket") { return "$($item.Scope)|MaxTicketAge" }
    if ($n -match "maxserviceage|lebensdauer.*serviceticket|lifetime for service ticket") { return "$($item.Scope)|MaxServiceAge" }
    if ($n -match "maxrenewage|erneuerung von benutzertickets|user ticket renewal") { return "$($item.Scope)|MaxRenewAge" }
    if ($n -match "ticketvalidateclient|benutzeranmeldeeinschraenkungen|user logon restrictions") { return "$($item.Scope)|TicketValidateClient" }
    return "$($item.Scope)|$($item.Category)|$($item.Name)"
}