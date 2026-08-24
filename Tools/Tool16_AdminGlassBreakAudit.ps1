<#
================================================================================
 TOOL 16: AD TIER-0, PRIVILEGED GROUPS & GLASS-BREAK COMPLIANCE AUDIT
 Funktionsname: Open-ToolAdminGlassBreakAudit (Tool16 AdminGlassBreakAudit)
 Version:       v1.0.0
 
 BESCHREIBUNG & ABLAUF:
 --------------------------------------------------------------------------------
 Dieses Modul führt eine strukturierte Sicherheitsprüfung der Active Directory
 Tier-0-Konten und dedizierten Glass-Break-Notfallkonten durch.

 1. Domänen- & DC-Erkennung:
    - Verifiziert die AD-Mitgliedschaft und identifiziert alle verfügbaren DCs.
    - Ermittelt die Domänen-SID sowie die Built-in RID 500-Objekte.

 2. Tab 1 - Benutzer- & Glass-Break-Konten (Master-Detail):
    - Linke Übersicht: Privilegierte Konten (Domain Admins, Enterprise Admins,
      Administratoren, gefilterte Glass-Break-Muster).
    - Rechte Detail-Tabelle (Eigenschaft | Status | Wert):
      * Vollständige Zeileneinfärbung je nach Status:
        - Grün:   OK (Protected Users, NotDelegated, PreAuth aktiv, etc.)
        - Gelb:   WARNUNG (RID 500, fehlende Workstation-Bindung, etc.)
        - Rot:    KRITISCH (Kennwort > 365 Tage, DONT_REQ_PREAUTH, SPN gesetzt)
      * Multi-DC LastLogon Scan über alle erreichbaren Domänencontroller.

 3. Tab 2 - Privilegierte Gruppen (Tier-0 Built-in):
    - Prüft Schema-Admins, Enterprise Admins, Domain Admins, Built-in Admins etc.
    - Sicherheitsregel: Die Gruppe 'Schema-Admins' muss im Regelbetrieb leer sein (0).
      Enthält sie Mitglieder, wird die gesamte Zeile signalrot hervorgehoben.
    - Rechte Tabelle: Zeigt bei Auswahl einer Gruppe direkt deren Mitglieder.

 4. CSV-Export:
    - Exportiert den jeweils aktiven Tab als Semikolon-getrennte Datei (UTF-8).
================================================================================
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.DirectoryServices
[System.Windows.Forms.Application]::EnableVisualStyles()

# ------------------------------------------------------------------------------
# THEME & LAYOUT KONFIGURATION
# ------------------------------------------------------------------------------
$script:UITheme = @{
    HeaderHeight       = 32
    RowHeight          = 25
    HeaderFont         = New-Object System.Drawing.Font("Segoe UI", 9.0, [System.Drawing.FontStyle]::Bold)
    CellFont           = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Regular)
    HeaderBackColor    = [System.Drawing.Color]::FromArgb(238, 242, 246)
    HeaderForeColor    = [System.Drawing.Color]::FromArgb(30, 41, 59)
    GridLineColor      = [System.Drawing.Color]::FromArgb(226, 232, 240)
    SelectionBackColor = [System.Drawing.Color]::FromArgb(203, 228, 249)
    SelectionForeColor = [System.Drawing.Color]::Black
}

function Apply-StandardGridTheme {
    param([System.Windows.Forms.DataGridView]$Grid)
    $Grid.EnableHeadersVisualStyles = $false
    $Grid.BorderStyle               = [System.Windows.Forms.BorderStyle]::FixedSingle
    $Grid.CellBorderStyle           = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
    $Grid.GridColor                 = $script:UITheme.GridLineColor
    $Grid.BackgroundColor           = [System.Drawing.Color]::White
    $Grid.RowHeadersVisible         = $false
    $Grid.AllowUserToAddRows        = $false
    $Grid.AllowUserToDeleteRows     = $false
    $Grid.SelectionMode             = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $Grid.MultiSelect               = $false
    $Grid.ReadOnly                  = $true
    $Grid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
    $Grid.ColumnHeadersHeight       = $script:UITheme.HeaderHeight

    $Grid.ColumnHeadersDefaultCellStyle.Font      = $script:UITheme.HeaderFont
    $Grid.ColumnHeadersDefaultCellStyle.BackColor = $script:UITheme.HeaderBackColor
    $Grid.ColumnHeadersDefaultCellStyle.ForeColor = $script:UITheme.HeaderForeColor
    $Grid.DefaultCellStyle.Font                   = $script:UITheme.CellFont
    $Grid.RowTemplate.Height                      = $script:UITheme.RowHeight
}

function Assert-DomainJoined {
    try {
        $sysInfo = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
        if (-not $sysInfo.PartOfDomain) {
            [System.Windows.Forms.MessageBox]::Show("Dieses Werkzeug erfordert eine Active Directory Domänenmitgliedschaft.", "Kein AD Join", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return $false
        }
        return $true
    } catch { return $true }
}

# ------------------------------------------------------------------------------
# HAUPTFENSTER: TOOL 16 ADMIN GLASS BREAK AUDIT
# ------------------------------------------------------------------------------
function Open-ToolAdminGlassBreakAudit {
    $toolVersion = "v1.0.0"

    if (Get-Command Assert-DomainJoined -ErrorAction SilentlyContinue) {
        if (-not (Assert-DomainJoined)) { return }
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Tool 16: Active Directory Tier-0 & Glass Break Compliance Audit ($toolVersion)"
    $form.Size = New-Object System.Drawing.Size(1520, 880)
    $form.MinimumSize = New-Object System.Drawing.Size(1100, 650)
    $form.StartPosition = "CenterScreen"
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

    # 1. Top Bar
    $pnlTop = New-Object System.Windows.Forms.Panel
    $pnlTop.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlTop.Height = 65
    $pnlTop.BackColor = [System.Drawing.Color]::FromArgb(240, 243, 246)
    $form.Controls.Add($pnlTop)

    $lblGlassFilter = New-Object System.Windows.Forms.Label
    $lblGlassFilter.Text = "Glass-Break Suchfilter:"
    $lblGlassFilter.Location = New-Object System.Drawing.Point(15, 12)
    $lblGlassFilter.AutoSize = $true
    $lblGlassFilter.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
    $pnlTop.Controls.Add($lblGlassFilter)

    $txtGlassPattern = New-Object System.Windows.Forms.TextBox
    $txtGlassPattern.Text = "*Break*;*Emergency*;*Notfall*;*GB-*;*Admin_*"
    $txtGlassPattern.Location = New-Object System.Drawing.Point(15, 32)
    $txtGlassPattern.Size = New-Object System.Drawing.Size(280, 25)
    $pnlTop.Controls.Add($txtGlassPattern)

    $btnRunAudit = New-Object System.Windows.Forms.Button
    $btnRunAudit.Text = "Audit starten"
    $btnRunAudit.Location = New-Object System.Drawing.Point(305, 30)
    $btnRunAudit.Size = New-Object System.Drawing.Size(115, 28)
    $btnRunAudit.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $btnRunAudit.ForeColor = [System.Drawing.Color]::White
    $btnRunAudit.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnRunAudit.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $pnlTop.Controls.Add($btnRunAudit)

    $btnExportCSV = New-Object System.Windows.Forms.Button
    $btnExportCSV.Text = "CSV Export (;)"
    $btnExportCSV.Location = New-Object System.Drawing.Point(430, 30)
    $btnExportCSV.Size = New-Object System.Drawing.Size(130, 28)
    $btnExportCSV.BackColor = [System.Drawing.Color]::FromArgb(16, 124, 65)
    $btnExportCSV.ForeColor = [System.Drawing.Color]::White
    $btnExportCSV.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnExportCSV.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $pnlTop.Controls.Add($btnExportCSV)

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = New-Object System.Drawing.Point(575, 35)
    $lblStatus.AutoSize = $true
    $lblStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(0, 70, 150)
    $pnlTop.Controls.Add($lblStatus)

    # 2. Fetter Hinweisbanner oben
    $pnlBanner = New-Object System.Windows.Forms.Panel
    $pnlBanner.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlBanner.Height = 32
    $pnlBanner.BackColor = [System.Drawing.Color]::FromArgb(254, 242, 242)
    $pnlBanner.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $form.Controls.Add($pnlBanner)

    $lblTopWarning = New-Object System.Windows.Forms.Label
    $lblTopWarning.Text = "SICHERHEITSRICHTLINIE: Die Gruppe 'Schema-Admins' muss im regulären Betrieb leer sein (0 Mitglieder). Kennwörter älter als 365 Tage gelten als KRITISCH."
    $lblTopWarning.Dock = [System.Windows.Forms.DockStyle]::Fill
    $lblTopWarning.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $lblTopWarning.Padding = New-Object System.Windows.Forms.Padding(10, 0, 0, 0)
    $lblTopWarning.ForeColor = [System.Drawing.Color]::DarkRed
    $lblTopWarning.Font = New-Object System.Drawing.Font("Segoe UI", 9.0, [System.Drawing.FontStyle]::Bold)
    $pnlBanner.Controls.Add($lblTopWarning)

    # 3. Tab-Container
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Dock = [System.Windows.Forms.DockStyle]::Fill
    $form.Controls.Add($tabControl)
    $tabControl.BringToFront()

    # ==========================================================================
    # TAB 1: BENUTZER (2-SPALTEN MASTER-DETAIL)
    # ==========================================================================
    $tabUsers = New-Object System.Windows.Forms.TabPage
    $tabUsers.Text = "Benutzer- & Glass-Break-Konten"
    $tabUsers.BackColor = [System.Drawing.Color]::White
    $tabControl.Controls.Add($tabUsers)

    $splitUsers = New-Object System.Windows.Forms.SplitContainer
    $splitUsers.Dock = [System.Windows.Forms.DockStyle]::Fill
    $splitUsers.Orientation = [System.Windows.Forms.Orientation]::Vertical
    $tabUsers.Controls.Add($splitUsers)

    $grpUsersLeft = New-Object System.Windows.Forms.GroupBox
    $grpUsersLeft.Text = "Geprüfte Tier-0 & Glass-Break Konten (Klick für Details)"
    $grpUsersLeft.Dock = [System.Windows.Forms.DockStyle]::Fill
    $grpUsersLeft.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
    $splitUsers.Panel1.Controls.Add($grpUsersLeft)

    $gridUsers = New-Object System.Windows.Forms.DataGridView
    $gridUsers.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridUsers.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
    Apply-StandardGridTheme $gridUsers
    $grpUsersLeft.Controls.Add($gridUsers)

    $grpUsersRight = New-Object System.Windows.Forms.GroupBox
    $grpUsersRight.Text = "Detailprüfung für ausgewähltes Konto (Eigenschaft | Status | Wert)"
    $grpUsersRight.Dock = [System.Windows.Forms.DockStyle]::Fill
    $grpUsersRight.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
    $splitUsers.Panel2.Controls.Add($grpUsersRight)

    $gridUserDetails = New-Object System.Windows.Forms.DataGridView
    $gridUserDetails.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridUserDetails.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
    Apply-StandardGridTheme $gridUserDetails
    $grpUsersRight.Controls.Add($gridUserDetails)

    # ==========================================================================
    # TAB 2: TIER-0 GRUPPEN
    # ==========================================================================
    $tabGroups = New-Object System.Windows.Forms.TabPage
    $tabGroups.Text = "Privilegierte Gruppen (Tier-0 / Built-in)"
    $tabGroups.BackColor = [System.Drawing.Color]::White
    $tabControl.Controls.Add($tabGroups)

    $splitGroups = New-Object System.Windows.Forms.SplitContainer
    $splitGroups.Dock = [System.Windows.Forms.DockStyle]::Fill
    $splitGroups.Orientation = [System.Windows.Forms.Orientation]::Vertical
    $tabGroups.Controls.Add($splitGroups)

    $grpGroupsLeft = New-Object System.Windows.Forms.GroupBox
    $grpGroupsLeft.Text = "Tier-0 Gruppen (Klick für Mitglieder)"
    $grpGroupsLeft.Dock = [System.Windows.Forms.DockStyle]::Fill
    $grpGroupsLeft.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
    $splitGroups.Panel1.Controls.Add($grpGroupsLeft)

    $gridGroups = New-Object System.Windows.Forms.DataGridView
    $gridGroups.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridGroups.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
    Apply-StandardGridTheme $gridGroups
    $grpGroupsLeft.Controls.Add($gridGroups)

    $grpGroupsRight = New-Object System.Windows.Forms.GroupBox
    $grpGroupsRight.Text = "Enthaltene Gruppenmitglieder (Objekte)"
    $grpGroupsRight.Dock = [System.Windows.Forms.DockStyle]::Fill
    $grpGroupsRight.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
    $splitGroups.Panel2.Controls.Add($grpGroupsRight)

    $gridGroupMembers = New-Object System.Windows.Forms.DataGridView
    $gridGroupMembers.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridGroupMembers.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
    Apply-StandardGridTheme $gridGroupMembers
    $grpGroupsRight.Controls.Add($gridGroupMembers)

    # --------------------------------------------------------------------------
    # FARB- & FORMATIERUNGSLOGIK (GANZE ZEILEN EINGEFÄRBT)
    # --------------------------------------------------------------------------
    $applyGridFormatting = {
        # 1. Linke Benutzerliste
        foreach ($row in $gridUsers.Rows) {
            $overall = [string]$row.Cells["Gesamt-Audit"].Value
            if ($overall -eq "KRITISCH") {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(253, 232, 232)
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkRed
                $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(235, 180, 180)
                $row.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::DarkRed
                $row.DefaultCellStyle.Font = New-Object System.Drawing.Font($gridUsers.Font, [System.Drawing.FontStyle]::Bold)
            } elseif ($overall -eq "WARNUNG") {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 248, 220)
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(180, 80, 0)
                $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(245, 230, 180)
                $row.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::FromArgb(140, 60, 0)
            } else {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(232, 248, 232)
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(0, 100, 0)
                $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(200, 240, 200)
                $row.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::FromArgb(0, 80, 0)
            }
        }

        # 2. Rechte Detail-Tabelle
        foreach ($row in $gridUserDetails.Rows) {
            $statusVal = [string]$row.Cells["Status"].Value

            if ($statusVal -eq "OK") {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(232, 248, 232)
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(0, 100, 0)
                $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(200, 240, 200)
                $row.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::FromArgb(0, 80, 0)
                $row.DefaultCellStyle.Font = New-Object System.Drawing.Font($gridUserDetails.Font, [System.Drawing.FontStyle]::Regular)
            }
            elseif ($statusVal -eq "KRITISCH") {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(253, 232, 232)
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkRed
                $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(235, 180, 180)
                $row.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::DarkRed
                $row.DefaultCellStyle.Font = New-Object System.Drawing.Font($gridUserDetails.Font, [System.Drawing.FontStyle]::Bold)
            }
            elseif ($statusVal -eq "WARNUNG") {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 248, 220)
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(180, 80, 0)
                $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(245, 230, 180)
                $row.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::FromArgb(140, 60, 0)
                $row.DefaultCellStyle.Font = New-Object System.Drawing.Font($gridUserDetails.Font, [System.Drawing.FontStyle]::Regular)
            }
            else {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::White
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(30, 41, 59)
                $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(203, 228, 249)
                $row.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
                $row.DefaultCellStyle.Font = New-Object System.Drawing.Font($gridUserDetails.Font, [System.Drawing.FontStyle]::Regular)
            }
        }

        # 3. Gruppenliste
        foreach ($row in $gridGroups.Rows) {
            $gName = [string]$row.Cells["Gruppenname"].Value
            $mCount = [int]$row.Cells["Mitglieder-Anzahl"].Value

            if ($gName -match "Schema[- ]?Admins" -and $mCount -gt 0) {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(253, 232, 232)
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(180, 0, 0)
                $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(240, 170, 170)
                $row.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::DarkRed
                $row.DefaultCellStyle.Font = New-Object System.Drawing.Font($gridGroups.Font, [System.Drawing.FontStyle]::Bold)
            } elseif ($row.Cells["Audit-Status"].Value -like "*HINWEIS*") {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 248, 220)
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(180, 80, 0)
            } else {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(232, 248, 232)
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(0, 100, 0)
            }
        }
    }

    $script:UserFullDataMap = @{}
    $script:GroupMembersMap = @{}

    $gridUsers.Add_SelectionChanged({
        if ($gridUsers.SelectedRows.Count -gt 0) {
            $sAM = [string]$gridUsers.SelectedRows[0].Cells["SamAccountName"].Value
            if ($script:UserFullDataMap.ContainsKey($sAM)) {
                $gridUserDetails.DataSource = [System.Collections.ArrayList]::new($script:UserFullDataMap[$sAM])
                & $applyGridFormatting
            }
        }
    })

    $gridGroups.Add_SelectionChanged({
        if ($gridGroups.SelectedRows.Count -gt 0) {
            $gName = [string]$gridGroups.SelectedRows[0].Cells["Gruppenname"].Value
            if ($script:GroupMembersMap.ContainsKey($gName)) {
                $gridGroupMembers.DataSource = [System.Collections.ArrayList]::new($script:GroupMembersMap[$gName])
                $grpGroupsRight.Text = "Mitglieder von $gName ($($script:GroupMembersMap[$gName].Count))"
            }
        }
    })

    # CSV-Export
    $btnExportCSV.Add_Click({
        $activeTab = $tabControl.SelectedTab.Text
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "CSV-Datei (*.csv)|*.csv"
        $sfd.FileName = "AD_Tier0_Audit_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

        if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                $exportData = @()
                if ($activeTab -like "*Benutzer*") {
                    foreach ($key in $script:UserFullDataMap.Keys) {
                        $rowObj = [ordered]@{}
                        foreach ($prop in $script:UserFullDataMap[$key]) {
                            $rowObj["$($prop.Eigenschaft) [Status]"] = $prop.Status
                            $rowObj["$($prop.Eigenschaft) [Wert]"]   = $prop.Wert
                        }
                        $exportData += [PSCustomObject]$rowObj
                    }
                } else {
                    foreach ($gKey in $script:GroupMembersMap.Keys) {
                        $members = $script:GroupMembersMap[$gKey]
                        if ($members.Count -gt 0) {
                            foreach ($m in $members) {
                                $exportData += [PSCustomObject]@{
                                    "Gruppe"         = $gKey
                                    "Name"           = $m.Name
                                    "SamAccountName" = $m.SamAccountName
                                    "Enabled"        = $m.Enabled
                                    "OU"             = $m.OU
                                    "DN"             = $m.DistinguishedName
                                }
                            }
                        } else {
                            $exportData += [PSCustomObject]@{
                                "Gruppe"         = $gKey
                                "Name"           = "[LEER / Keine Mitglieder]"
                                "SamAccountName" = ""
                                "Enabled"        = ""
                                "OU"             = ""
                                "DN"             = ""
                            }
                        }
                    }
                }

                $exportData | Export-Csv -Path $sfd.FileName -Delimiter ';' -NoTypeInformation -Encoding UTF8
                [System.Windows.Forms.MessageBox]::Show("Daten erfolgreich nach '$($sfd.FileName)' exportiert!", "Export erfolgreich", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Fehler beim Exportieren: $($_.Exception.Message)", "Exportfehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    })

    # --------------------------------------------------------------------------
    # AUDIT ENGINE
    # --------------------------------------------------------------------------
    $runAuditAction = {
        $btnRunAudit.Enabled = $false
        $btnExportCSV.Enabled = $false
        $lblStatus.Text = "Lese Tier-0-Gruppen, Benutzer und DCs ein ($toolVersion)..."
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $form.Refresh()

        $userSummaries = [System.Collections.Generic.List[PSCustomObject]]::new()
        $groupSummaries = [System.Collections.Generic.List[PSCustomObject]]::new()
        $script:UserFullDataMap = @{}
        $script:GroupMembersMap = @{}
        $processedUserDns = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

        try {
            $rootDSE = [ADSI]"LDAP://RootDSE"
            $domDN = $rootDSE.defaultNamingContext.ToString()
            $domainEntry = [ADSI]"LDAP://$domDN"
            $domainSID = (New-Object System.Security.Principal.SecurityIdentifier($domainEntry.Properties["objectSid"][0], 0)).Value

            $domainObj = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
            $dcNames = @($domainObj.DomainControllers | ForEach-Object { $_.Name })

            $tier0Groups = @(
                @{ Name = "Schema-Admins";        SID = "$domainSID-518"; Expected = "0 Mitglieder (Immer leer im Normalbetrieb)" },
                @{ Name = "Organisations-Admins";  SID = "$domainSID-519"; Expected = "Minimale dedizierte Konten" },
                @{ Name = "Domänen-Admins";        SID = "$domainSID-512"; Expected = "Minimale Tier-0 Notfall-/Admin-Konten" },
                @{ Name = "Administratoren";       SID = "S-1-5-32-544";   Expected = "Tier-0 Built-in Gruppe" },
                @{ Name = "Protected Users";      SID = "$domainSID-525"; Expected = "Schutzgruppe für Tier-0 Konten" },
                @{ Name = "Konto-Operatoren";      SID = "S-1-5-32-548";   Expected = "Sollte leer sein (Escalation Risiko)" }
            )

            # 1. Gruppen analysieren
            foreach ($g in $tier0Groups) {
                $gSearcher = New-Object DirectoryServices.DirectorySearcher($domainEntry)
                $gSearcher.Filter = "(&(objectCategory=group)(objectsid=$($g.SID)))"
                $groupObj = $gSearcher.FindOne()

                $memberCount = 0
                $gName = $g.Name
                $membersForGroupGrid = [System.Collections.Generic.List[PSCustomObject]]::new()

                if ($groupObj) {
                    $gName = if ($groupObj.Properties["samaccountname"].Count -gt 0) { $groupObj.Properties["samaccountname"][0] } else { $g.Name }
                    $memberCount = $groupObj.Properties["member"].Count

                    foreach ($mDN in $groupObj.Properties["member"]) {
                        $mEntry = [ADSI]"LDAP://$mDN"
                        $oc = $mEntry.Properties["objectClass"]
                        $isUser = ($oc -contains "user" -or $oc -contains "person") -and ($oc -notcontains "group")

                        $mSam = [string]$mEntry.Properties["samaccountname"].Value
                        $mName = [string]$mEntry.Properties["name"].Value
                        $mUac = if ($mEntry.Properties["useraccountcontrol"].Value) { [int]$mEntry.Properties["useraccountcontrol"].Value } else { 0 }
                        $mEnabled = (-not (($mUac -band 2) -eq 2))
                        $mOU = if ($mDN -match "OU=.*") { $mDN.Substring($mDN.IndexOf("OU=")) } else { "CN=Users" }

                        $membersForGroupGrid.Add([PSCustomObject]@{
                            "SamAccountName"    = $mSam
                            "Name"              = $mName
                            "Enabled"           = $mEnabled
                            "OU"                = $mOU
                            "DistinguishedName" = $mDN
                        })

                        if ($isUser -and $g.SID -ne "$domainSID-525" -and $processedUserDns.Add($mDN)) {
                            $res = Evaluate-Account -UserDN $mDN -RoleName $gName -IsSchemaAdmin ($g.SID -eq "$domainSID-518") -DCs $dcNames -DomainSID $domainSID
                            if ($res) {
                                $userSummaries.Add($res.Summary)
                                $script:UserFullDataMap[$res.Summary.SamAccountName] = $res.Details
                            }
                        }
                    }
                }

                $script:GroupMembersMap[$gName] = $membersForGroupGrid

                $gStatus = "OK"
                if ($g.SID -eq "$domainSID-518" -and $memberCount -gt 0) {
                    $gStatus = "KRITISCH (Schema-Admins nicht leer!)"
                } elseif ($g.SID -eq "S-1-5-32-548" -and $memberCount -gt 0) {
                    $gStatus = "HINWEIS (Eskalationsrisiko)"
                }

                $groupSummaries.Add([PSCustomObject]@{
                    "Audit-Status"      = $gStatus
                    "Gruppenname"       = $gName
                    "Mitglieder-Anzahl" = $memberCount
                    "Soll-Zustand"      = $g.Expected
                    "Gruppen-SID"       = $g.SID
                })
            }

            # 2. Primäre Gruppenmitglieder (primaryGroupID 512)
            $pSearcher = New-Object DirectoryServices.DirectorySearcher($domainEntry)
            $pSearcher.Filter = "(&(objectCategory=person)(objectClass=user)(primaryGroupID=512))"
            foreach ($res in $pSearcher.FindAll()) {
                $dn = [string]$res.Properties["distinguishedname"][0]
                if ($processedUserDns.Add($dn)) {
                    $uRes = Evaluate-Account -UserDN $dn -RoleName "Domänen-Admins (PrimaryGroupID)" -IsSchemaAdmin $false -DCs $dcNames -DomainSID $domainSID
                    if ($uRes) {
                        $userSummaries.Add($uRes.Summary)
                        $script:UserFullDataMap[$uRes.Summary.SamAccountName] = $uRes.Details
                    }
                }
            }

            # 3. Glass-Break Konten abfragen
            $gbPatterns = $txtGlassPattern.Text.Split(';')
            foreach ($pat in $gbPatterns) {
                $cleanPat = $pat.Trim()
                if (-not [string]::IsNullOrEmpty($cleanPat)) {
                    $gbSearcher = New-Object DirectoryServices.DirectorySearcher($domainEntry)
                    $gbSearcher.Filter = "(&(objectCategory=person)(objectClass=user)(|(sAMAccountName=$cleanPat)(name=$cleanPat)))"
                    $gbSearcher.PageSize = 250

                    foreach ($res in $gbSearcher.FindAll()) {
                        $dn = [string]$res.Properties["distinguishedname"][0]
                        if ($processedUserDns.Add($dn)) {
                            $uRes = Evaluate-Account -UserDN $dn -RoleName "Glass Break Account" -IsSchemaAdmin $false -DCs $dcNames -DomainSID $domainSID
                            if ($uRes) {
                                $userSummaries.Add($uRes.Summary)
                                $script:UserFullDataMap[$uRes.Summary.SamAccountName] = $uRes.Details
                            }
                        }
                    }
                }
            }

            $gridUsers.DataSource = [System.Collections.ArrayList]::new($userSummaries)
            $gridGroups.DataSource = [System.Collections.ArrayList]::new($groupSummaries)
            & $applyGridFormatting

            if ($gridUsers.Rows.Count -gt 0) { $gridUsers.Rows[0].Selected = $true }
            if ($gridGroups.Rows.Count -gt 0) { $gridGroups.Rows[0].Selected = $true }

            $lblStatus.Text = "Fertig ($toolVersion): $($userSummaries.Count) Konten & $($groupSummaries.Count) Tier-0 Gruppen geprüft."
        }
        catch {
            $lblStatus.Text = "Fehler: $($_.Exception.Message)"
        }
        finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            $btnRunAudit.Enabled = $true
            $btnExportCSV.Enabled = $true
        }
    }

    # Helper: Datensätze mit 3 Spalten (Eigenschaft | Status | Wert)
    function Evaluate-Account {
        param([string]$UserDN, [string]$RoleName, [bool]$IsSchemaAdmin, [string[]]$DCs, [string]$DomainSID)

        try {
            $uEntry = [ADSI]"LDAP://$UserDN"
            $sAM = [string]$uEntry.Properties["samaccountname"].Value
            $name = [string]$uEntry.Properties["name"].Value
            if (-not $sAM) { return $null }

            $userSid = (New-Object System.Security.Principal.SecurityIdentifier($uEntry.Properties["objectSid"][0], 0)).Value
            $isRid500 = $userSid.EndsWith("-500")

            $uac = if ($uEntry.Properties["userAccountControl"].Value) { [int]$uEntry.Properties["userAccountControl"].Value } else { 0 }
            $isDisabled          = ($uac -band 0x0002) -eq 0x0002
            $pwdNeverExpires     = ($uac -band 0x10000) -eq 0x10000
            $smartCardRequired   = ($uac -band 0x40000) -eq 0x40000
            $isNotDelegated      = ($uac -band 0x100000) -eq 0x100000
            $dontReqPreAuth      = ($uac -band 0x400000) -eq 0x400000
            $reversibleEnc       = ($uac -band 0x0080) -eq 0x0080

            # Protected Users
            $isProtectedUser = $false
            foreach ($mGroupDN in $uEntry.Properties["memberOf"]) {
                if ($mGroupDN -match "Protected Users") { $isProtectedUser = $true; break }
            }

            $hasSPN           = ($uEntry.Properties["servicePrincipalName"].Count -gt 0)
            $adminCount       = [string]$uEntry.Properties["adminCount"].Value
            $userWorkstations = [string]$uEntry.Properties["userWorkstations"].Value
            $ouPath           = if ($UserDN -match "OU=.*") { $UserDN.Substring($UserDN.IndexOf("OU=")) } else { "CN=Users" }

            # Kennwortalter berechnen
            $pwdAgeDays = 0
            $pwdNeverSet = $true
            if ($uEntry.Properties["pwdLastSet"].Value) {
                $rawPwd = $uEntry.Properties["pwdLastSet"].Value
                $ftPwd = 0
                if ($rawPwd.GetType().Name -eq "__ComObject") {
                    $hi = [int64]$rawPwd.GetType().InvokeMember("HighPart", [System.Reflection.BindingFlags]::GetProperty, $null, $rawPwd, $null)
                    $lo = [int64]$rawPwd.GetType().InvokeMember("LowPart", [System.Reflection.BindingFlags]::GetProperty, $null, $rawPwd, $null)
                    $ftPwd = ($hi -shl 32) + ($lo -band 0xFFFFFFFF)
                } else { $ftPwd = [int64]$rawPwd }

                if ($ftPwd -gt 0) {
                    $pwdDate = [DateTime]::FromFileTime($ftPwd)
                    $pwdAgeDays = [Math]::Floor(((Get-Date) - $pwdDate).TotalDays)
                    $pwdNeverSet = $false
                }
            }

            # Multi-DC LastLogon Scan
            $latestLogon = [DateTime]::MinValue
            $queryCount = 0
            foreach ($dc in $DCs) {
                try {
                    $dcEntry = [ADSI]"LDAP://$dc/$UserDN"
                    if ($dcEntry -and $dcEntry.Properties["lastLogon"].Value) {
                        $raw = $dcEntry.Properties["lastLogon"].Value
                        $ft = 0
                        if ($raw.GetType().Name -eq "__ComObject") {
                            $hi = [int64]$raw.GetType().InvokeMember("HighPart", [System.Reflection.BindingFlags]::GetProperty, $null, $raw, $null)
                            $lo = [int64]$raw.GetType().InvokeMember("LowPart", [System.Reflection.BindingFlags]::GetProperty, $null, $raw, $null)
                            $ft = ($hi -shl 32) + ($lo -band 0xFFFFFFFF)
                        } else { $ft = [int64]$raw }
                        if ($ft -gt 0) {
                            $dt = [DateTime]::FromFileTime($ft)
                            if ($dt -gt $latestLogon) { $latestLogon = $dt }
                        }
                    }
                    $queryCount++
                } catch { }
            }

            $logonText = if ($latestLogon -eq [DateTime]::MinValue) { "Nie eingeloggt" } else {
                $days = [Math]::Floor(((Get-Date) - $latestLogon).TotalDays)
                "$($latestLogon.ToString('dd.MM.yyyy HH:mm:ss')) ($days Tage her)"
            }

            # Gesamt-Bewertung
            $hasCritical = ($IsSchemaAdmin -or $pwdNeverSet -or ($pwdAgeDays -gt 365) -or $dontReqPreAuth -or $hasSPN -or $reversibleEnc)
            $hasWarning  = ($isRid500 -or (-not $isProtectedUser) -or (-not $smartCardRequired) -or (-not $isNotDelegated) -or [string]::IsNullOrWhiteSpace($userWorkstations) -or ($adminCount -ne "1"))

            $overallStatus = "OK"
            if ($hasCritical) { $overallStatus = "KRITISCH" }
            elseif ($hasWarning) { $overallStatus = "WARNUNG" }

            $summary = [PSCustomObject]@{
                "Gesamt-Audit"         = $overallStatus
                "SamAccountName"       = $sAM
                "Name"                 = $name
                "Enabled"              = (-not $isDisabled)
                "Admin-Rolle / Quelle" = $RoleName
                "OU"                   = $ouPath
            }

            # 3-Spaltige Detail-Liste aufbauen
            $details = [System.Collections.Generic.List[PSCustomObject]]::new()

            $details.Add([PSCustomObject]@{ "Eigenschaft" = "SamAccountName"; "Status" = "INFO"; "Wert" = $sAM })
            $details.Add([PSCustomObject]@{ "Eigenschaft" = "Anzeigename"; "Status" = "INFO"; "Wert" = $name })
            $details.Add([PSCustomObject]@{ "Eigenschaft" = "Konto-Status (Enabled)"; "Status" = if (-not $isDisabled) { "OK" } else { "INFO" }; "Wert" = if (-not $isDisabled) { "Aktiv (True)" } else { "Deaktiviert (False)" } })
            $details.Add([PSCustomObject]@{ "Eigenschaft" = "Speicherort (OU)"; "Status" = "INFO"; "Wert" = $ouPath })
            $details.Add([PSCustomObject]@{ "Eigenschaft" = "Gesamt-Audit Bewertung"; "Status" = $overallStatus; "Wert" = $overallStatus })
            
            # Passwort-Alter (365 Tage Check)
            $pwdStatus = if ($pwdNeverSet -or $pwdAgeDays -gt 365) { "KRITISCH" } else { "OK" }
            $pwdText   = if ($pwdNeverSet) { "Initial / Nie gesetzt" } elseif ($pwdAgeDays -gt 365) { "$pwdAgeDays Tage (> 365 Tage max)" } else { "$pwdAgeDays Tage (Gültig)" }
            $details.Add([PSCustomObject]@{ "Eigenschaft" = "Passwortalter (Max 365 Tage)"; "Status" = $pwdStatus; "Wert" = $pwdText })

            $details.Add([PSCustomObject]@{ "Eigenschaft" = "Passwort läuft nie ab"; "Status" = if ($pwdNeverExpires) { "WARNUNG" } else { "OK" }; "Wert" = if ($pwdNeverExpires) { "Ja (PasswordNeverExpires)" } else { "Nein" } })
            $details.Add([PSCustomObject]@{ "Eigenschaft" = "Smartcard-Pflicht (MFA)"; "Status" = if ($smartCardRequired) { "OK" } else { "WARNUNG" }; "Wert" = if ($smartCardRequired) { "Konfiguriert" } else { "Fehlt (SmartCardRequired ist false)" } })
            $details.Add([PSCustomObject]@{ "Eigenschaft" = "Mitglied Protected Users"; "Status" = if ($isProtectedUser) { "OK" } else { "WARNUNG" }; "Wert" = if ($isProtectedUser) { "Ja (Geschützt)" } else { "Nein (Nicht in Protected Users)" } })
            $details.Add([PSCustomObject]@{ "Eigenschaft" = "Delegierungsschutz"; "Status" = if ($isNotDelegated) { "OK" } else { "WARNUNG" }; "Wert" = if ($isNotDelegated) { "Geschützt (NotDelegated)" } else { "NICHT OK (Delegierbar)" } })
            $details.Add([PSCustomObject]@{ "Eigenschaft" = "Pre-Authentication (AS-REP)"; "Status" = if (-not $dontReqPreAuth) { "OK" } else { "KRITISCH" }; "Wert" = if (-not $dontReqPreAuth) { "Erforderlich" } else { "NICHT OK (DONT_REQ_PREAUTH aktiv)" } })
            $details.Add([PSCustomObject]@{ "Eigenschaft" = "Service Principal Name (SPN)"; "Status" = if (-not $hasSPN) { "OK" } else { "KRITISCH" }; "Wert" = if (-not $hasSPN) { "Kein SPN gesetzt" } else { "NICHT OK (SPN vorhanden: Kerberoasting)" } })
            $details.Add([PSCustomObject]@{ "Eigenschaft" = "Umkehrbare Verschlüsselung"; "Status" = if (-not $reversibleEnc) { "OK" } else { "KRITISCH" }; "Wert" = if (-not $reversibleEnc) { "Deaktiviert" } else { "NICHT OK (Klartextfähig)" } })
            $details.Add([PSCustomObject]@{ "Eigenschaft" = "Workstation-Einschränkung"; "Status" = if (-not [string]::IsNullOrWhiteSpace($userWorkstations)) { "OK" } else { "WARNUNG" }; "Wert" = if (-not [string]::IsNullOrWhiteSpace($userWorkstations)) { $userWorkstations } else { "Keine Einschränkung (Logon anywhere)" } })
            $details.Add([PSCustomObject]@{ "Eigenschaft" = "AdminSDHolder (adminCount)"; "Status" = if ($adminCount -eq "1") { "OK" } else { "WARNUNG" }; "Wert" = if ($adminCount -eq "1") { "1 (Geschützt)" } else { "$adminCount (Abweichend)" } })
            $details.Add([PSCustomObject]@{ "Eigenschaft" = "Konto-Typ"; "Status" = if ($isRid500) { "WARNUNG" } else { "OK" }; "Wert" = if ($isRid500) { "Built-in RID 500 Administrator" } else { "Dediziertes Konto" } })
            $details.Add([PSCustomObject]@{ "Eigenschaft" = "Letzte Anmeldung (Multi-DC)"; "Status" = "INFO"; "Wert" = "$logonText (Geprüft über $queryCount DCs)" })
            $details.Add([PSCustomObject]@{ "Eigenschaft" = "DistinguishedName"; "Status" = "INFO"; "Wert" = $UserDN })

            return @{ Summary = $summary; Details = $details }
        } catch { return $null }
    }

    $btnRunAudit.Add_Click($runAuditAction)

    $form.Add_Shown({
        try {
            if ($splitUsers.Width -gt 600) {
                $splitUsers.SplitterDistance = [Math]::Floor($splitUsers.Width * 0.48)
            }
            if ($splitGroups.Width -gt 600) {
                $splitGroups.SplitterDistance = [Math]::Floor($splitGroups.Width * 0.48)
            }
        } catch { }
        & $runAuditAction
    })

    [void]$form.ShowDialog()
}

# Standalone-Start
if ($MyInvocation.InvocationName -ne '.' -and $MyInvocation.Line -notmatch 'Open-Tool') {
    Open-ToolAdminGlassBreakAudit
}
