<#
.SYNOPSIS
    Tool 22: AD Standard-Gruppen Checker & Container 'CN=Users' Vergleich (Multilingual DE/EN)
.DESCRIPTION
    - Vergleicht Standard-AD-Gruppen über sprachunabhängige RIDs (< 1000) mit den real existierenden Gruppen im CN=Users-Container.
    - Berücksichtigt deutsche und englische Gruppennamen (z.B. Domänen-Admins / Domain Admins).
    - Erkennt Zusatzsoftware / manuelle Gruppen anhand RID >= 1000 oder Fremdhersteller-Präfixen.
    - Umschaltbare GUI-Sprache (DE / EN), Filter und CSV-Export.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Import-Module ActiveDirectory -ErrorAction SilentlyContinue

# ------------------------------------------------------------------------------
# 0. MULTILANGUAGE DICTIONARY (I18N: DE / EN)
# ------------------------------------------------------------------------------
$script:CurrentLang = "DE"

$script:I18N = @{
    "DE" = @{
        "Title"              = "Tool 22: AD Users Standard-Gruppen Checker & Analyse"
        "BtnScan"            = "Vergleich durchführen"
        "LblFilter"          = "Filter:"
        "BtnExport"          = "CSV Export"
        "StatusReady"        = "Bereit. Klicken Sie auf 'Vergleich durchführen'."
        "StatusConnecting"   = "Verbinde mit Active Directory..."
        "StatusDone"         = "Analyse abgeschlossen. Gesamt: {0} Einträge verglichen."
        "StatusError"        = "Fehler: {0}"
        "ColGroupName"       = "Gruppenname (AD)"
        "ColStatus"          = "Klassifizierung"
        "ColRID"             = "RID"
        "ColMembers"         = "Mitglieder"
        "ColRole"            = "Funktion / Rolle"
        "ColStdPurpose"      = "Standard-Zweck (Soll)"
        "ColCurDesc"         = "Aktuelle Beschreibung"
        "StatusAll"          = "Alle anzeigen"
        "StatusPresent"      = "Standard AD (Vorhanden)"
        "StatusMissing"      = "Fehlt im AD (Standard)"
        "StatusCustom"       = "Zusatzsoftware / Manuell"
        "NotPresent"         = "Nicht im AD vorhanden"
        "CustomRole"         = "Zusatzsoftware / Manuell erstellt"
        "CustomPurpose"      = "Nicht im Standard-AD-Katalog verankert (RID >= 1000)"
        "MsgNoData"          = "Keine Daten für den Export vorhanden."
        "MsgExportDone"      = "Export erfolgreich gespeichert:`r`n{0}"
        "ErrTitle"           = "Fehler"
    }
    "EN" = @{
        "Title"              = "Tool 22: AD Users Default Groups Checker & Analysis"
        "BtnScan"            = "Run Comparison"
        "LblFilter"          = "Filter:"
        "BtnExport"          = "Export CSV"
        "StatusReady"        = "Ready. Click 'Run Comparison' to start."
        "StatusConnecting"   = "Connecting to Active Directory..."
        "StatusDone"         = "Analysis finished. Total: {0} entries compared."
        "StatusError"        = "Error: {0}"
        "ColGroupName"       = "Group Name (AD)"
        "ColStatus"          = "Classification"
        "ColRID"             = "RID"
        "ColMembers"         = "Members"
        "ColRole"            = "Function / Role"
        "ColStdPurpose"      = "Default Purpose"
        "ColCurDesc"         = "Current Description"
        "StatusAll"          = "Show All"
        "StatusPresent"      = "Default AD (Present)"
        "StatusMissing"      = "Missing in AD (Default)"
        "StatusCustom"       = "Third-Party / Custom"
        "NotPresent"         = "Not present in AD"
        "CustomRole"         = "Third-party Software / Custom"
        "CustomPurpose"      = "Not part of native AD catalog (RID >= 1000)"
        "MsgNoData"          = "No data available to export."
        "MsgExportDone"      = "Export saved successfully:`r`n{0}"
        "ErrTitle"           = "Error"
    }
}

function Get-Text {
    param([string]$Key)
    if ($script:I18N[$script:CurrentLang].ContainsKey($Key)) {
        return $script:I18N[$script:CurrentLang][$Key]
    }
    return $Key
}

# ------------------------------------------------------------------------------
# 1. SOLL-KATALOG MIT EN/DE NAMEN & RID-BINDUNG
# ------------------------------------------------------------------------------
$script:DefaultADGroupsCatalog = @(
    @{ RID = 512; NameEN = "Domain Admins";                      NameDE = "Domänen-Admins";                         Scope = "Global";      RoleDE = "AD Admin";       RoleEN = "AD Admin";        DescDE = "Designierte Administratoren der Domäne"; DescEN = "Designated administrators of the domain" },
    @{ RID = 513; NameEN = "Domain Users";                       NameDE = "Domänen-Benutzer";                       Scope = "Global";      RoleDE = "AD Basis";       RoleEN = "AD Core";         DescDE = "Alle Benutzerkonten der Domäne"; DescEN = "All domain user accounts" },
    @{ RID = 514; NameEN = "Domain Guests";                      NameDE = "Domänen-Gäste";                          Scope = "Global";      RoleDE = "AD Basis";       RoleEN = "AD Core";         DescDE = "Alle Gastkonten der Domäne"; DescEN = "All domain guest accounts" },
    @{ RID = 515; NameEN = "Domain Computers";                   NameDE = "Domänencomputer";                        Scope = "Global";      RoleDE = "AD Basis";       RoleEN = "AD Core";         DescDE = "Alle Workstations und Member-Server"; DescEN = "All workstations and servers joined to the domain" },
    @{ RID = 516; NameEN = "Domain Controllers";                 NameDE = "Domänencontroller";                      Scope = "Global";      RoleDE = "AD Core";        RoleEN = "AD Core";         DescDE = "Alle beschreibbaren DCs der Domäne"; DescEN = "All writable domain controllers in the domain" },
    @{ RID = 517; NameEN = "Cert Publishers";                    NameDE = "Zertifikatherausgeber";                  Scope = "DomainLocal"; RoleDE = "Zertifikate";    RoleEN = "Certificates";    DescDE = "Server mit Erlaubnis zur Zertifikatspublikation im AD"; DescEN = "Servers permitted to publish certificates to the directory" },
    @{ RID = 518; NameEN = "Schema Admins";                      NameDE = "Schema-Admins";                          Scope = "Universal";   RoleDE = "Forest Admin";   RoleEN = "Forest Admin";    DescDE = "Verwaltung des Gesamtstruktur-Schemas (Forest-Root)"; DescEN = "Designated schema administrators (Forest-Root only)" },
    @{ RID = 519; NameEN = "Enterprise Admins";                  NameDE = "Organisations-Admins";                   Scope = "Universal";   RoleDE = "Forest Admin";   RoleEN = "Forest Admin";    DescDE = "Gesamtstruktur-weite Administration (Forest-Root)"; DescEN = "Designated enterprise administrators (Forest-Root only)" },
    @{ RID = 520; NameEN = "Group Policy Creator Owners";        NameDE = "Richtlinien-Ersteller-Besitzer";         Scope = "Global";      RoleDE = "GPO Admin";      RoleEN = "GPO Admin";       DescDE = "Ersteller und Besitzer von Gruppenrichtlinien"; DescEN = "Members can modify group policies for the domain" },
    @{ RID = 521; NameEN = "Read-only Domain Controllers";       NameDE = "Schreibgeschützte Domänencontroller";    Scope = "Global";      RoleDE = "RODC";           RoleEN = "RODC";            DescDE = "Alle RODCs der Domäne"; DescEN = "All read-only domain controllers in the domain" },
    @{ RID = 522; NameEN = "Cloneable Domain Controllers";       NameDE = "Klonbare Domänencontroller";             Scope = "Global";      RoleDE = "AD Core";        RoleEN = "AD Core";         DescDE = "DCs, die virtualisiert geklont werden dürfen"; DescEN = "DCs that are allowed to be cloned virtually" },
    @{ RID = 525; NameEN = "Protected Users";                    NameDE = "Geschützte Benutzer";                    Scope = "Global";      RoleDE = "AD Security";    RoleEN = "AD Security";     DescDE = "Erzwingt strenge Sicherheits- und Kerberos-Regeln"; DescEN = "Enforces strict security restrictions & non-cachable creds" },
    @{ RID = 526; NameEN = "Key Admins";                         NameDE = "Schlüssel-Admins";                       Scope = "Global";      RoleDE = "Kryptografie";   RoleEN = "Cryptography";    DescDE = "Verwaltung kryptografischer Schlüssel"; DescEN = "Members can perform administrative actions on key objects" },
    @{ RID = 527; NameEN = "Enterprise Key Admins";              NameDE = "Organisations-Schlüssel-Admins";          Scope = "Universal";   RoleDE = "Forest Security";RoleEN = "Forest Security"; DescDE = "Forest-weite kryptografische Schlüsselverwaltung"; DescEN = "Forest-wide administrative actions on key objects" },
    @{ RID = 553; NameEN = "RAS and IAS Servers";                NameDE = "RAS- und IAS-Server";                    Scope = "DomainLocal"; RoleDE = "Netzwerk";       RoleEN = "Network";         DescDE = "Server für RAS- und RADIUS-Authentifizierung"; DescEN = "Servers that can access remote access properties of users" },
    @{ RID = 571; NameEN = "Allowed RODC Password Replication Group"; NameDE = "Zulässige RODC-Kennwortreplikationsgruppe"; Scope = "DomainLocal"; RoleDE = "RODC";  RoleEN = "RODC";            DescDE = "Konten, deren Passwörter auf RODCs repliziert werden dürfen"; DescEN = "Members permitted to replicate passwords to RODCs" },
    @{ RID = 572; NameEN = "Denied RODC Password Replication Group";  NameDE = "Verweigerte RODC-Kennwortreplikationsgruppe"; Scope = "DomainLocal"; RoleDE = "RODC"; RoleEN = "RODC";            DescDE = "Konten, deren Passwörter NIEMALS auf RODCs liegen dürfen"; DescEN = "Members never permitted to replicate passwords to RODCs" },
    @{ RID = $null; NameEN = "DnsAdmins";                        NameDE = "DnsAdmins";                              Scope = "DomainLocal"; RoleDE = "DNS Rolle";      RoleEN = "DNS Role";        DescDE = "Administratoren für Windows DNS-Dienste"; DescEN = "DNS administrators with access to network DNS" },
    @{ RID = $null; NameEN = "DnsUpdateProxy";                   NameDE = "DnsUpdateProxy";                         Scope = "Global";      RoleDE = "DNS Rolle";      RoleEN = "DNS Role";        DescDE = "DNS-Clients zur dynamischen Registrierung für Drittsysteme"; DescEN = "DNS clients permitted to perform dynamic updates for other clients" }
)

# ------------------------------------------------------------------------------
# 2. GUI AUFBAU
# ------------------------------------------------------------------------------
$MainForm = New-Object System.Windows.Forms.Form
$MainForm.Text = Get-Text "Title"
$MainForm.Size = New-Object System.Drawing.Size(1300, 720)
$MainForm.MinimumSize = New-Object System.Drawing.Size(980, 520)
$MainForm.StartPosition = "CenterScreen"
$MainForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$MainForm.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

# Top Bar
$TopPanel = New-Object System.Windows.Forms.Panel
$TopPanel.Dock = [System.Windows.Forms.DockStyle]::Top
$TopPanel.Height = 55
$TopPanel.BackColor = [System.Drawing.Color]::FromArgb(240, 243, 246)
$MainForm.Controls.Add($TopPanel)

$BtnScan = New-Object System.Windows.Forms.Button
$BtnScan.Location = New-Object System.Drawing.Point(15, 12)
$BtnScan.Size = New-Object System.Drawing.Size(185, 30)
$BtnScan.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$BtnScan.ForeColor = [System.Drawing.Color]::White
$BtnScan.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$BtnScan.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$TopPanel.Controls.Add($BtnScan)

$LblFilter = New-Object System.Windows.Forms.Label
$LblFilter.Location = New-Object System.Drawing.Point(215, 18)
$LblFilter.AutoSize = $true
$LblFilter.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$TopPanel.Controls.Add($LblFilter)

$TxtFilter = New-Object System.Windows.Forms.TextBox
$TxtFilter.Location = New-Object System.Drawing.Point(265, 15)
$TxtFilter.Size = New-Object System.Drawing.Size(190, 24)
$TopPanel.Controls.Add($TxtFilter)

$CmbStatus = New-Object System.Windows.Forms.ComboBox
$CmbStatus.Location = New-Object System.Drawing.Point(465, 15)
$CmbStatus.Size = New-Object System.Drawing.Size(230, 24)
$CmbStatus.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$TopPanel.Controls.Add($CmbStatus)

# Sprachumschalter Buttons [ EN ] [ DE ]
$BtnLangEN = New-Object System.Windows.Forms.Button
$BtnLangEN.Location = New-Object System.Drawing.Point(1005, 13)
$BtnLangEN.Size = New-Object System.Drawing.Size(42, 28)
$BtnLangEN.Text = "EN"
$BtnLangEN.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$BtnLangEN.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$TopPanel.Controls.Add($BtnLangEN)

$BtnLangDE = New-Object System.Windows.Forms.Button
$BtnLangDE.Location = New-Object System.Drawing.Point(1052, 13)
$BtnLangDE.Size = New-Object System.Drawing.Size(42, 28)
$BtnLangDE.Text = "DE"
$BtnLangDE.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$BtnLangDE.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$TopPanel.Controls.Add($BtnLangDE)

$BtnExport = New-Object System.Windows.Forms.Button
$BtnExport.Location = New-Object System.Drawing.Point(1110, 12)
$BtnExport.Size = New-Object System.Drawing.Size(155, 30)
$BtnExport.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$BtnExport.BackColor = [System.Drawing.Color]::FromArgb(16, 124, 65)
$BtnExport.ForeColor = [System.Drawing.Color]::White
$BtnExport.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$BtnExport.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$TopPanel.Controls.Add($BtnExport)

# Statuszeile
$StatusStrip = New-Object System.Windows.Forms.StatusStrip
$StatusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$StatusStrip.Items.Add($StatusLabel) | Out-Null
$MainForm.Controls.Add($StatusStrip)

# DataGridView
$DataGrid = New-Object System.Windows.Forms.DataGridView
$DataGrid.Dock = [System.Windows.Forms.DockStyle]::Fill
$DataGrid.EnableHeadersVisualStyles = $false
$DataGrid.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$DataGrid.CellBorderStyle = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
$DataGrid.GridColor = [System.Drawing.Color]::FromArgb(226, 232, 240)
$DataGrid.BackgroundColor = [System.Drawing.Color]::White
$DataGrid.RowHeadersVisible = $false
$DataGrid.AllowUserToAddRows = $false
$DataGrid.AllowUserToDeleteRows = $false
$DataGrid.ReadOnly = $true
$DataGrid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$DataGrid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$DataGrid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
$DataGrid.ColumnHeadersHeight = 36
$DataGrid.RowTemplate.Height = 26
$DataGrid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$DataGrid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(238, 242, 246)
$DataGrid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(30, 41, 59)
$MainForm.Controls.Add($DataGrid)
$DataGrid.BringToFront()

# Farblogik (Grün = Standard OK, Gelb = Zusatzsoftware, Rot = Fehlt)
$DataGrid.Add_RowPrePaint({
    param($s, $e)
    if ($e.RowIndex -ge 0 -and $e.RowIndex -lt $DataGrid.Rows.Count) {
        $row = $DataGrid.Rows[$e.RowIndex]
        $classification = [string]$row.Cells["Klassifizierung"].Value

        if ($classification -like "*Zusatzsoftware*" -or $classification -like "*Third-Party*") {
            $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(254, 249, 231)
            $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(160, 90, 0)
        } elseif ($classification -like "*Fehlt*" -or $classification -like "*Missing*") {
            $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(254, 242, 242)
            $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(185, 28, 28)
        } elseif ($classification -like "*Standard AD*" -or $classification -like "*Default AD*") {
            $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(240, 253, 244)
            $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(21, 128, 61)
        }
    }
})

# ------------------------------------------------------------------------------
# 3. SPRACHUMSCHALTUNG & FILTERAKTUALISIERUNG
# ------------------------------------------------------------------------------
function Update-Localization {
    $MainForm.Text    = Get-Text "Title"
    $BtnScan.Text     = Get-Text "BtnScan"
    $LblFilter.Text   = Get-Text "LblFilter"
    $BtnExport.Text   = Get-Text "BtnExport"

    # Status-Buttons hervorheben
    if ($script:CurrentLang -eq "DE") {
        $BtnLangDE.BackColor = [System.Drawing.Color]::LightSteelBlue
        $BtnLangEN.BackColor = [System.Drawing.SystemColors]::Control
    } else {
        $BtnLangEN.BackColor = [System.Drawing.Color]::LightSteelBlue
        $BtnLangDE.BackColor = [System.Drawing.SystemColors]::Control
    }

    # Dropdown neu befüllen
    $curIndex = [math]::Max(0, $CmbStatus.SelectedIndex)
    $CmbStatus.Items.Clear()
    [void]$CmbStatus.Items.AddRange(@(
        (Get-Text "StatusAll"),
        (Get-Text "StatusPresent"),
        (Get-Text "StatusCustom"),
        (Get-Text "StatusMissing")
    ))
    $CmbStatus.SelectedIndex = if ($curIndex -lt $CmbStatus.Items.Count) { $curIndex } else { 0 }

    if ($script:AnalysisResults.Count -eq 0) {
        $StatusLabel.Text = Get-Text "StatusReady"
    } else {
        & $ApplyFilterAction
    }
}

# ------------------------------------------------------------------------------
# 4. LOGIK: AD-ABFRAGE MIT RID- & SPRACHABGLEICH
# ------------------------------------------------------------------------------
$script:AnalysisResults = [System.Collections.Generic.List[PSCustomObject]]::new()

$RunComparisonAction = {
    $BtnScan.Enabled = $false
    $StatusLabel.Text = Get-Text "StatusConnecting"
    $MainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    [System.Windows.Forms.Application]::DoEvents()

    $script:AnalysisResults.Clear()

    try {
        $rootDSE = [ADSI]"LDAP://RootDSE"
        $domainDN = $rootDSE.defaultNamingContext.Value
        $usersContainerDN = "CN=Users,$domainDN"

        # Abfrage aller Gruppen aus CN=Users
        $searcher = [System.DirectoryServices.DirectorySearcher]::new([ADSI]"LDAP://$usersContainerDN")
        $searcher.Filter = "(objectCategory=group)"
        $searcher.SearchScope = [System.DirectoryServices.SearchScope]::OneLevel
        $searcher.PropertiesToLoad.AddRange(@("name", "sAMAccountName", "objectSid", "description", "member"))
        $foundGroups = $searcher.FindAll()

        # Zuordnung über RID und Kleinbuchstaben-Namen
        $adGroupByRid  = @{}
        $adGroupByName = @{}
        $allFoundKeys  = [System.Collections.Generic.HashSet[string]]::new()

        foreach ($item in $foundGroups) {
            $name = [string]$item.Properties["name"][0]
            $sam  = [string]$item.Properties["samaccountname"][0]
            $desc = if ($item.Properties["description"].Count -gt 0) { [string]$item.Properties["description"][0] } else { "" }
            $mCount = if ($item.Properties["member"].Count -gt 0) { $item.Properties["member"].Count } else { 0 }
            
            $rid = $null
            if ($item.Properties["objectsid"].Count -gt 0) {
                $sidBytes = $item.Properties["objectsid"][0]
                $secId = New-Object System.Security.Principal.SecurityIdentifier($sidBytes, 0)
                $sidParts = $secId.Value -split '-'
                $rid = [int]$sidParts[-1]
            }

            $entry = @{
                Name        = $name
                SAM         = $sam
                RID         = $rid
                Description = $desc
                MemberCount = $mCount
            }

            if ($rid) { $adGroupByRid[$rid] = $entry }
            $adGroupByName[$sam.ToLower()] = $entry
            [void]$allFoundKeys.Add($sam.ToLower())
        }
        $foundGroups.Dispose()
        $searcher.Dispose()

        $processedSams = [System.Collections.Generic.HashSet[string]]::new()

        # 1. Soll-Katalog abgleichen (Primär über RID, Fallback über DE/EN Name)
        foreach ($cat in $script:DefaultADGroupsCatalog) {
            $matched = $null

            if ($cat.RID -and $adGroupByRid.ContainsKey($cat.RID)) {
                $matched = $adGroupByRid[$cat.RID]
            } elseif ($adGroupByName.ContainsKey($cat.NameEN.ToLower())) {
                $matched = $adGroupByName[$cat.NameEN.ToLower()]
            } elseif ($adGroupByName.ContainsKey($cat.NameDE.ToLower())) {
                $matched = $adGroupByName[$cat.NameDE.ToLower()]
            }

            $displayName = if ($script:CurrentLang -eq "DE") { "$($cat.NameDE) / $($cat.NameEN)" } else { "$($cat.NameEN) / $($cat.NameDE)" }
            $role        = if ($script:CurrentLang -eq "DE") { $cat.RoleDE } else { $cat.RoleEN }
            $descStd     = if ($script:CurrentLang -eq "DE") { $cat.DescDE } else { $cat.DescEN }

            if ($matched) {
                [void]$processedSams.Add($matched.SAM.ToLower())
                $script:AnalysisResults.Add([PSCustomObject]@{
                    "Gruppenname"       = "$($matched.SAM)  ($displayName)"
                    "Klassifizierung"   = Get-Text "StatusPresent"
                    "RID"               = if ($matched.RID) { $matched.RID } else { "-" }
                    "Mitglieder"        = $matched.MemberCount
                    "Funktion / Rolle"  = $role
                    "Standard-Zweck"    = $descStd
                    "Aktuelle Beschr."  = $matched.Description
                })
            } else {
                $script:AnalysisResults.Add([PSCustomObject]@{
                    "Gruppenname"       = $displayName
                    "Klassifizierung"   = Get-Text "StatusMissing"
                    "RID"               = if ($cat.RID) { $cat.RID } else { "-" }
                    "Mitglieder"        = 0
                    "Funktion / Rolle"  = $role
                    "Standard-Zweck"    = $descStd
                    "Aktuelle Beschr."  = Get-Text "NotPresent"
                })
            }
        }

        # 2. Alle restlichen Gruppen in CN=Users als Zusatzsoftware / Custom einstufen
        foreach ($samKey in $allFoundKeys) {
            if (-not $processedSams.Contains($samKey)) {
                $extra = $adGroupByName[$samKey]
                $detectedRole = Get-Text "CustomRole"

                # Bekannte Rollen erkennen
                if ($extra.SAM -match "^(Exchange|MSExch)") { $detectedRole = "Microsoft Exchange" }
                elseif ($extra.SAM -match "^(ConfigMgr|SMS)") { $detectedRole = "MECM / SCCM" }
                elseif ($extra.SAM -match "^(DHCP)") { $detectedRole = "Microsoft DHCP" }
                elseif ($extra.SAM -match "^(Veeam)") { $detectedRole = "Veeam Backup" }
                elseif ($extra.SAM -match "(vCenter|View|VMware)") { $detectedRole = "VMware" }
                elseif ($extra.SAM -match "^(AAD_|MSOL_)") { $detectedRole = "Entra / Cloud Sync" }

                $script:AnalysisResults.Add([PSCustomObject]@{
                    "Gruppenname"       = $extra.SAM
                    "Klassifizierung"   = Get-Text "StatusCustom"
                    "RID"               = if ($extra.RID) { $extra.RID } else { "-" }
                    "Mitglieder"        = $extra.MemberCount
                    "Funktion / Rolle"  = $detectedRole
                    "Standard-Zweck"    = Get-Text "CustomPurpose"
                    "Aktuelle Beschr."  = $extra.Description
                })
            }
        }

        & $ApplyFilterAction
        $StatusLabel.Text = [string]::Format((Get-Text "StatusDone"), $script:AnalysisResults.Count)
    } catch {
        $StatusLabel.Text = [string]::Format((Get-Text "StatusError"), $_.Exception.Message)
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, (Get-Text "ErrTitle"), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } finally {
        $MainForm.Cursor = [System.Windows.Forms.Cursors]::Default
        $BtnScan.Enabled = $true
    }
}

# ------------------------------------------------------------------------------
# 5. FILTER & DATENBINDUNG
# ------------------------------------------------------------------------------
$ApplyFilterAction = {
    $search = $TxtFilter.Text.Trim()
    $status = if ($CmbStatus.SelectedItem) { $CmbStatus.SelectedItem.ToString() } else { Get-Text "StatusAll" }

    $filtered = $script:AnalysisResults | Where-Object {
        $matchSearch = if ([string]::IsNullOrWhiteSpace($search)) { $true } else {
            $_.Gruppenname -like "*$search*" -or `
            $_."Funktion / Rolle" -like "*$search*" -or `
            $_."Aktuelle Beschr." -like "*$search*"
        }

        $matchStatus = if ($status -eq (Get-Text "StatusAll")) { $true } else {
            $_.Klassifizierung -eq $status
        }

        $matchSearch -and $matchStatus
    }

    $DataGrid.DataSource = [System.Collections.ArrayList]::new(@($filtered))

    # Spaltentitel an aktuelle Sprache anpassen
    if ($DataGrid.Columns.Count -gt 0) {
        $DataGrid.Columns["Gruppenname"].HeaderText      = Get-Text "ColGroupName"
        $DataGrid.Columns["Klassifizierung"].HeaderText  = Get-Text "ColStatus"
        $DataGrid.Columns["RID"].HeaderText              = Get-Text "ColRID"
        $DataGrid.Columns["Mitglieder"].HeaderText       = Get-Text "ColMembers"
        $DataGrid.Columns["Funktion / Rolle"].HeaderText = Get-Text "ColRole"
        $DataGrid.Columns["Standard-Zweck"].HeaderText   = Get-Text "ColStdPurpose"
        $DataGrid.Columns["Aktuelle Beschr."].HeaderText = Get-Text "ColCurDesc"
    }
}

# ------------------------------------------------------------------------------
# 6. EVENT-HANDLER & EXPORT
# ------------------------------------------------------------------------------
$BtnScan.Add_Click({ & $RunComparisonAction })
$TxtFilter.Add_TextChanged({ & $ApplyFilterAction })
$CmbStatus.Add_SelectedIndexChanged({ & $ApplyFilterAction })

$BtnLangDE.Add_Click({
    if ($script:CurrentLang -ne "DE") {
        $script:CurrentLang = "DE"
        Update-Localization
        if ($script:AnalysisResults.Count -gt 0) { & $RunComparisonAction }
    }
})

$BtnLangEN.Add_Click({
    if ($script:CurrentLang -ne "EN") {
        $script:CurrentLang = "EN"
        Update-Localization
        if ($script:AnalysisResults.Count -gt 0) { & $RunComparisonAction }
    }
})

$BtnExport.Add_Click({
    if ($DataGrid.Rows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show((Get-Text "MsgNoData"), "Info", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = "CSV-Datei (*.csv)|*.csv"
    $sfd.FileName = "AD_DefaultGroups_Check_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"

    if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $items = @($DataGrid.DataSource)
            $items | Export-Csv -Path $sfd.FileName -Delimiter ';' -NoTypeInformation -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show([string]::Format((Get-Text "MsgExportDone"), $sfd.FileName), "CSV Export", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, (Get-Text "ErrTitle"), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})

# Initialisierung
Update-Localization
$MainForm.ShowDialog() | Out-Null