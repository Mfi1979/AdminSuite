<#
.SYNOPSIS
    Tool 17: Active Directory Gruppen- & Mitgliedschafts-Analyse (Externe Regeldatei)
.DESCRIPTION
    - Auslagerung und Speicherung der Klassifizierungsregeln in 'ADGroup_ClassificationRules.txt'.
    - Layout-Bereinigung: Keine abgeschnittenen Buttons und keine überlappenden Filterzeilen.
    - Schriftgröße 10pt / 10.5pt Bold mit automatischer DPI- und Layout-Anpassung.
    - Vollständige CSV-Exporte mit separaten PowerShell-Befehlsspalten und getrenntem Datum/Uhrzeit.
    - Umschaltbare PowerShell-Aktionsleiste mit Multi-Command Generator.
    - Built-In Erkennung strikt nach Well-Known SIDs / Standard-RIDs mit Level 'AD'.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.DirectoryServices
Add-Type -AssemblyName Microsoft.VisualBasic

if (-not [System.Windows.Forms.Application]::RenderWithVisualStyles) {
    try { [System.Windows.Forms.Application]::EnableVisualStyles() } catch {}
}

# ------------------------------------------------------------------------------
# THEME & LAYOUT HELPER
# ------------------------------------------------------------------------------
$script:UITheme = @{
    HeaderHeight       = 48
    RowHeight          = 28
    HeaderFont         = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    CellFont           = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Regular)
    BaseFont           = New-Object System.Drawing.Font("Segoe UI", 10.0, [System.Drawing.FontStyle]::Regular)
    BoldFont           = New-Object System.Drawing.Font("Segoe UI", 10.0, [System.Drawing.FontStyle]::Bold)
    HeaderBackColor    = [System.Drawing.Color]::FromArgb(238, 242, 246)
    HeaderForeColor    = [System.Drawing.Color]::FromArgb(30, 41, 59)
    GridLineColor      = [System.Drawing.Color]::FromArgb(226, 232, 240)
    SelectionBackColor = [System.Drawing.Color]::FromArgb(203, 228, 249)
    SelectionForeColor = [System.Drawing.Color]::Black
}

# ------------------------------------------------------------------------------
# REGEL-SPEICHERPFAD & PERSISTENZ (TXT/JSON-Auslagerung)
# ------------------------------------------------------------------------------
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) }
if ([string]::IsNullOrWhiteSpace($scriptDir) -or -not (Test-Path $scriptDir)) {
    $scriptDir = [Environment]::GetFolderPath("MyDocuments")
}
$script:RulesFilePath = Join-Path $scriptDir "ADGroup_ClassificationRules.txt"

$script:DefaultRules = @(
    [PSCustomObject]@{ Kennung = "Azure";      Suchmuster = "^AD-AAD-";                 RollenName = "Azure";                 LevelModus = "Azure-Sub"; LevelDetails = "Intune=Intune" },
    [PSCustomObject]@{ Kennung = "GPO";        Suchmuster = "^GPO-";                    RollenName = "Group Policy";          LevelModus = "GPO-Sub";   LevelDetails = "GPO-COM-=Computer;GPO-USR-=Benutzer;GPO-MIX-=Computer/User" },
    [PSCustomObject]@{ Kennung = "RDP";        Suchmuster = "(-RDP$|^RDP-)";            RollenName = "Remote Desktop";        LevelModus = "Keiner";    LevelDetails = "-" },
    [PSCustomObject]@{ Kennung = "LokalAdmin"; Suchmuster = "(-lokalAdmin$|-localAdmin$)"; RollenName = "Lokaler Administrator"; LevelModus = "Festwert"; LevelDetails = "Administrator" },
    [PSCustomObject]@{ Kennung = "FSR";        Suchmuster = "(^FSR-|^.*-FSR-)";         RollenName = "File Server Rights";    LevelModus = "FSR-Rechte";LevelDetails = "-FC=Full Control;-RW=Read Write;-RO=Read Only;-FL=Folder List" },
    [PSCustomObject]@{ Kennung = "Laufwerk";   Suchmuster = "^(Laufwerk|lw_)-?";        RollenName = "Laufwerk";              LevelModus = "Keiner";    LevelDetails = "-" },
    [PSCustomObject]@{ Kennung = "Drucker";    Suchmuster = "(^Printer-|^Drucker-)";    RollenName = "Drucker";               LevelModus = "Keiner";    LevelDetails = "-" }
)

$script:ActiveRules = [System.Collections.Generic.List[PSCustomObject]]::new()

function Save-ClassificationRulesToFile {
    try {
        $json = $script:ActiveRules | ConvertTo-Json -Depth 4
        Set-Content -Path $script:RulesFilePath -Value $json -Encoding UTF8 -Force
    } catch {}
}

function Load-ClassificationRulesFromFile {
    $script:ActiveRules.Clear()
    if (Test-Path $script:RulesFilePath) {
        try {
            $content = Get-Content -Path $script:RulesFilePath -Raw -Encoding UTF8
            if (-not [string]::IsNullOrWhiteSpace($content)) {
                $loaded = $content | ConvertFrom-Json
                foreach ($item in $loaded) {
                    $script:ActiveRules.Add([PSCustomObject]@{
                        Kennung      = [string]$item.Kennung
                        Suchmuster   = [string]$item.Suchmuster
                        RollenName   = [string]$item.RollenName
                        LevelModus   = [string]$item.LevelModus
                        LevelDetails = [string]$item.LevelDetails
                    })
                }
                return
            }
        } catch {}
    }

    # Fallback auf Standardregeln & initiale Dateierstellung
    foreach ($r in $script:DefaultRules) { $script:ActiveRules.Add($r) }
    Save-ClassificationRulesToFile
}

# Regeln initial laden
Load-ClassificationRulesFromFile

function Test-IsWellKnownBuiltInSid {
    param([string]$SidString)
    if ([string]::IsNullOrWhiteSpace($SidString)) { return $false }
    if ($SidString -match "^S-1-5-32-\d+$") { return $true }
    if ($SidString -match "^S-1-5-(?:9|11|21-[0-9\-]+-(?:498|500|501|502|512|513|514|515|516|517|518|519|520|521|522|525|526|527|553|571|572))$") { return $true }
    return $false
}

function Get-CustomGroupClassification {
    param(
        [string]$GroupName,
        [string]$SidString = ""
    )

    if (Test-IsWellKnownBuiltInSid -SidString $SidString) {
        return [PSCustomObject]@{ Role = "Built-In"; Level = "AD" }
    }

    foreach ($rule in $script:ActiveRules) {
        $matched = $false
        try { $matched = ($GroupName -match $rule.Suchmuster) } catch { $matched = $false }

        if ($matched) {
            $levelVal = "-"
            if ($rule.LevelModus -eq "Festwert") {
                $levelVal = $rule.LevelDetails
            } elseif ($rule.LevelModus -eq "Azure-Sub") {
                if ($GroupName -match "Intune") { $levelVal = "Intune" } else { $levelVal = "-" }
            } elseif ($rule.LevelModus -eq "GPO-Sub") {
                if ($GroupName -match "^GPO-COM-") { $levelVal = "Computer" }
                elseif ($GroupName -match "^GPO-USR-") { $levelVal = "Benutzer" }
                elseif ($GroupName -match "^GPO-MIX-") { $levelVal = "Computer/User" }
                else { $levelVal = "Allgemein" }
            } elseif ($rule.LevelModus -eq "FSR-Rechte") {
                $levelVal = "Unbekannt"
                $pairs = $rule.LevelDetails -split ';'
                foreach ($pair in $pairs) {
                    $kv = $pair -split '='
                    if ($kv.Count -eq 2) {
                        $suffix = $kv[0].Trim()
                        $desc = $kv[1].Trim()
                        if ($GroupName -match "$([regex]::Escape($suffix))$") {
                            $levelVal = $desc
                            break
                        }
                    }
                }
            }
            return [PSCustomObject]@{ Role = $rule.RollenName; Level = $levelVal }
        }
    }
    return [PSCustomObject]@{ Role = "Benutzer"; Level = "-" }
}

function Extract-TargetComputerName {
    param([string]$GroupName, [string]$Role)
    if ($Role -eq "Lokaler Administrator") {
        if ($GroupName -match "^(.+)-(?:lokalAdmin|localAdmin)$") { return $Matches[1].Trim() }
        if ($GroupName -match "^(?:lokalAdmin|localAdmin)-(.+)$") { return $Matches[1].Trim() }
    } elseif ($Role -eq "Remote Desktop") {
        if ($GroupName -match "^(.+)-RDP$") { return $Matches[1].Trim() }
        if ($GroupName -match "^RDP-(.+)$") { return $Matches[1].Trim() }
    }
    return $null
}

function Get-AutoSuggestedRules {
    param([System.Collections.Generic.List[PSCustomObject]]$GroupsList)

    $suggestions = [System.Collections.Generic.List[PSCustomObject]]::new()
    if ($null -eq $GroupsList -or $GroupsList.Count -eq 0) { return $suggestions }

    $prefixCounts = @{}
    foreach ($g in $GroupsList) {
        $name = $g.Gruppenname
        if ($name -match "^([a-zA-Z0-9]{2,12})[-_]") {
            $p = $Matches[1]
            if ($p -notmatch "^(Azure|GPO|RDP|FSR|Laufwerk|Drucker)$") {
                if (-not $prefixCounts.ContainsKey($p)) { $prefixCounts[$p] = 0 }
                $prefixCounts[$p]++
            }
        }
    }

    foreach ($key in $prefixCounts.Keys) {
        if ($prefixCounts[$key] -ge 2) {
            $regexPattern = "^$([regex]::Escape($key))[-_]"
            $alreadyExists = $script:ActiveRules | Where-Object { $_.Suchmuster -eq $regexPattern -or $_.Kennung -eq $key }
            if (-not $alreadyExists) {
                $roleGuess = switch -Wildcard ($key) {
                    "*VPN*"   { "VPN Zugriff" }
                    "*SEC*"   { "Sicherheitsgruppe" }
                    "*SG*"    { "Sicherheitsgruppe" }
                    "*APP*"   { "Software / App" }
                    "*SW*"    { "Softwareverteilung" }
                    "*M365*"  { "Cloud / M365" }
                    "*O365*"  { "Cloud / O365" }
                    "*FS*"    { "File Server" }
                    "*SHARE*" { "Freigabe-Berechtigung" }
                    Default   { "$key Rolle" }
                }

                $suggestions.Add([PSCustomObject]@{
                    Kennung      = $key
                    Suchmuster   = $regexPattern
                    RollenName   = $roleGuess
                    LevelModus   = "Keiner"
                    LevelDetails = "-"
                })
            }
        }
    }
    return $suggestions
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
    $Grid.MultiSelect               = $true
    $Grid.ReadOnly                  = $true
    $Grid.ClipboardCopyMode         = [System.Windows.Forms.DataGridViewClipboardCopyMode]::EnableAlwaysIncludeHeaderText
    $Grid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
    $Grid.ColumnHeadersHeight       = $script:UITheme.HeaderHeight

    $Grid.ColumnHeadersDefaultCellStyle.Font      = $script:UITheme.HeaderFont
    $Grid.ColumnHeadersDefaultCellStyle.BackColor = $script:UITheme.HeaderBackColor
    $Grid.ColumnHeadersDefaultCellStyle.ForeColor = $script:UITheme.HeaderForeColor
    $Grid.ColumnHeadersDefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
    $Grid.ColumnHeadersDefaultCellStyle.WrapMode  = [System.Windows.Forms.DataGridViewTriState]::True

    $Grid.DefaultCellStyle.Font                   = $script:UITheme.CellFont
    $Grid.RowTemplate.Height                      = $script:UITheme.RowHeight

    $ctxMenu = New-Object System.Windows.Forms.ContextMenuStrip
    $mnuCopy = $ctxMenu.Items.Add("Auswahl kopieren (Strg+C)")
    $mnuCopy.Add_Click({
        try {
            $cb = $Grid.GetClipboardContent()
            if ($cb) { [System.Windows.Forms.Clipboard]::SetDataObject($cb, $true) }
        } catch {}
    })
    $mnuCopyAll = $ctxMenu.Items.Add("Ganze Tabelle kopieren")
    $mnuCopyAll.Add_Click({
        try {
            $Grid.SelectAll()
            $cb = $Grid.GetClipboardContent()
            if ($cb) { [System.Windows.Forms.Clipboard]::SetDataObject($cb, $true) }
        } catch {}
    })
    $Grid.ContextMenuStrip = $ctxMenu
}

function Enable-UniversalGridSorting {
    param([System.Windows.Forms.DataGridView]$Grid)

    $Grid.Add_ColumnHeaderMouseClick({
        param($sender, $e)
        $targetGrid = $sender
        if ($null -eq $targetGrid -or $e.ColumnIndex -lt 0 -or $e.ColumnIndex -ge $targetGrid.Columns.Count) { return }

        $col = $targetGrid.Columns[$e.ColumnIndex]
        if ($null -eq $col) { return }

        $propName = if ($col.DataPropertyName) { $col.DataPropertyName } else { $col.HeaderText }
        if ([string]::IsNullOrWhiteSpace($propName)) { return }

        if ($null -eq $targetGrid.Tag -or -not ($targetGrid.Tag -is [hashtable])) {
            $targetGrid.Tag = @{ LastCol = ""; Asc = $true }
        }

        $state = $targetGrid.Tag
        if ($state.LastCol -eq $propName) { $state.Asc = -not $state.Asc } else { $state.LastCol = $propName; $state.Asc = $true }

        $items = @($targetGrid.DataSource)
        if ($null -eq $items -or $items.Count -le 1) { return }

        try {
            $sorted = $items | Sort-Object -Property @{
                Expression = {
                    $item = $_
                    if ($null -eq $item) { return "" }
                    $val = $item.$propName
                    if ($null -eq $val -or $val -eq "" -or $val -eq "-") { return "" }
                    if ($val -as [int]) { return [int]$val }
                    if ($val -match '^\d{2}\.\d{2}\.\d{4}') {
                        try { return [datetime]::ParseExact($val.ToString().Trim(), @("dd.MM.yyyy HH:mm", "dd.MM.yyyy HH:mm:ss", "dd.MM.yyyy"), $null) } catch { return $val }
                    }
                    return $val
                }
                Descending = (-not $state.Asc)
            }

            $targetGrid.SuspendLayout()
            $arr = [System.Collections.ArrayList]::new()
            foreach ($entry in $sorted) { [void]$arr.Add($entry) }
            $targetGrid.DataSource = $null
            $targetGrid.DataSource = $arr

            foreach ($c in $targetGrid.Columns) { $c.HeaderCell.SortGlyphDirection = [System.Windows.Forms.SortOrder]::None }
            $targetGrid.Columns[$e.ColumnIndex].HeaderCell.SortGlyphDirection = `
                $(if ($state.Asc) { [System.Windows.Forms.SortOrder]::Ascending } else { [System.Windows.Forms.SortOrder]::Descending })

            $targetGrid.ResumeLayout()
        } catch {}
    })
}

# ------------------------------------------------------------------------------
# POPUP-GUI: BESCHREIBUNGS-EDITOR
# ------------------------------------------------------------------------------
function Show-EditDescriptionDialog {
    param([string]$GroupName, [string]$CurrentDescription)

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Beschreibung bearbeiten: $GroupName"
    $dlg.Size = New-Object System.Drawing.Size(650, 420)
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.Font = $script:UITheme.BaseFont
    $dlg.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

    $lblOld = New-Object System.Windows.Forms.Label
    $lblOld.Text = "Aktuelle Beschreibung:"
    $lblOld.Location = New-Object System.Drawing.Point(15, 12)
    $lblOld.Size = New-Object System.Drawing.Size(600, 20)
    $lblOld.Font = $script:UITheme.BoldFont
    $dlg.Controls.Add($lblOld)

    $txtOld = New-Object System.Windows.Forms.TextBox
    $txtOld.Location = New-Object System.Drawing.Point(15, 35)
    $txtOld.Size = New-Object System.Drawing.Size(600, 55)
    $txtOld.Multiline = $true
    $txtOld.ReadOnly = $true
    $txtOld.BackColor = [System.Drawing.Color]::FromArgb(238, 242, 246)
    $txtOld.Text = if ([string]::IsNullOrWhiteSpace($CurrentDescription)) { "(Keine Beschreibung hinterlegt)" } else { $CurrentDescription }
    $dlg.Controls.Add($txtOld)

    $lblNew = New-Object System.Windows.Forms.Label
    $lblNew.Text = "Neue Beschreibung eingeben (generiert PowerShell-Befehl):"
    $lblNew.Location = New-Object System.Drawing.Point(15, 100)
    $lblNew.Size = New-Object System.Drawing.Size(600, 20)
    $lblNew.Font = $script:UITheme.BoldFont
    $lblNew.ForeColor = [System.Drawing.Color]::FromArgb(0, 102, 204)
    $dlg.Controls.Add($lblNew)

    $txtNew = New-Object System.Windows.Forms.TextBox
    $txtNew.Location = New-Object System.Drawing.Point(15, 125)
    $txtNew.Size = New-Object System.Drawing.Size(600, 150)
    $txtNew.Multiline = $true
    $txtNew.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $txtNew.Text = $CurrentDescription
    $dlg.Controls.Add($txtNew)

    $btnGenerate = New-Object System.Windows.Forms.Button
    $btnGenerate.Text = "Befehl uebernehmen"
    $btnGenerate.Location = New-Object System.Drawing.Point(280, 290)
    $btnGenerate.Size = New-Object System.Drawing.Size(185, 36)
    $btnGenerate.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $btnGenerate.ForeColor = [System.Drawing.Color]::White
    $btnGenerate.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnGenerate.Font = $script:UITheme.BoldFont
    $btnGenerate.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.Controls.Add($btnGenerate)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Abbrechen"
    $btnCancel.Location = New-Object System.Drawing.Point(480, 290)
    $btnCancel.Size = New-Object System.Drawing.Size(135, 36)
    $btnCancel.BackColor = [System.Drawing.Color]::FromArgb(230, 235, 240)
    $btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($btnCancel)

    $dlg.AcceptButton = $btnGenerate
    $dlg.CancelButton = $btnCancel
    $txtNew.SelectAll()

    $result = $dlg.ShowDialog()
    $enteredText = $txtNew.Text
    $dlg.Dispose()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return @{ Saved = $true; Value = $enteredText }
    } else {
        return @{ Saved = $false; Value = $null }
    }
}

# ------------------------------------------------------------------------------
# POPUP-GUI: REGEL-EDITOR (Speichert automatisch in TXT-Datei)
# ------------------------------------------------------------------------------
function Show-RuleEditorDialog {
    param(
        [scriptblock]$OnRulesUpdated,
        [System.Collections.Generic.List[PSCustomObject]]$CurrentGroups
    )

    $ruleForm = New-Object System.Windows.Forms.Form
    $ruleForm.Text = "Regel-Editor: AD Gruppen-Klassifizierung (Datei: $script:RulesFilePath)"
    $ruleForm.Size = New-Object System.Drawing.Size(1080, 680)
    $ruleForm.MinimumSize = New-Object System.Drawing.Size(900, 520)
    $ruleForm.StartPosition = "CenterParent"
    $ruleForm.Font = $script:UITheme.BaseFont
    $ruleForm.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

    try {
        $pnlEdit = New-Object System.Windows.Forms.GroupBox
        $pnlEdit.Text = "Regel hinzufuegen / bearbeiten"
        $pnlEdit.Dock = [System.Windows.Forms.DockStyle]::Top
        $pnlEdit.Height = 155
        $pnlEdit.Font = $script:UITheme.BoldFont
        $ruleForm.Controls.Add($pnlEdit)

        $lblK = New-Object System.Windows.Forms.Label
        $lblK.Text = "Kennung:"
        $lblK.Location = New-Object System.Drawing.Point(15, 25)
        $lblK.AutoSize = $true
        $pnlEdit.Controls.Add($lblK)

        $txtK = New-Object System.Windows.Forms.TextBox
        $txtK.Location = New-Object System.Drawing.Point(15, 48)
        $txtK.Size = New-Object System.Drawing.Size(130, 25)
        $txtK.Font = $script:UITheme.BaseFont
        $pnlEdit.Controls.Add($txtK)

        $lblM = New-Object System.Windows.Forms.Label
        $lblM.Text = "Suchmuster (Regex/Endung):"
        $lblM.Location = New-Object System.Drawing.Point(155, 25)
        $lblM.AutoSize = $true
        $pnlEdit.Controls.Add($lblM)

        $txtM = New-Object System.Windows.Forms.TextBox
        $txtM.Location = New-Object System.Drawing.Point(155, 48)
        $txtM.Size = New-Object System.Drawing.Size(220, 25)
        $txtM.Font = $script:UITheme.BaseFont
        $pnlEdit.Controls.Add($txtM)

        $lblR = New-Object System.Windows.Forms.Label
        $lblR.Text = "Rollenname / Funktion:"
        $lblR.Location = New-Object System.Drawing.Point(385, 25)
        $lblR.AutoSize = $true
        $pnlEdit.Controls.Add($lblR)

        $txtR = New-Object System.Windows.Forms.TextBox
        $txtR.Location = New-Object System.Drawing.Point(385, 48)
        $txtR.Size = New-Object System.Drawing.Size(220, 25)
        $txtR.Font = $script:UITheme.BaseFont
        $pnlEdit.Controls.Add($txtR)

        $lblMod = New-Object System.Windows.Forms.Label
        $lblMod.Text = "Level-Modus:"
        $lblMod.Location = New-Object System.Drawing.Point(615, 25)
        $lblMod.AutoSize = $true
        $pnlEdit.Controls.Add($lblMod)

        $cmbMod = New-Object System.Windows.Forms.ComboBox
        $cmbMod.Location = New-Object System.Drawing.Point(615, 48)
        $cmbMod.Size = New-Object System.Drawing.Size(160, 25)
        $cmbMod.Font = $script:UITheme.BaseFont
        $cmbMod.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
        [void]$cmbMod.Items.AddRange(@("Keiner", "Festwert", "FSR-Rechte", "Azure-Sub", "GPO-Sub"))
        $cmbMod.SelectedIndex = 0
        $pnlEdit.Controls.Add($cmbMod)

        $lblDet = New-Object System.Windows.Forms.Label
        $lblDet.Text = "Level-Details (Festwert oder Suffix=Text;):"
        $lblDet.Location = New-Object System.Drawing.Point(15, 82)
        $lblDet.AutoSize = $true
        $pnlEdit.Controls.Add($lblDet)

        $txtDet = New-Object System.Windows.Forms.TextBox
        $txtDet.Location = New-Object System.Drawing.Point(15, 106)
        $txtDet.Size = New-Object System.Drawing.Size(590, 25)
        $txtDet.Font = $script:UITheme.BaseFont
        $txtDet.Text = "-"
        $pnlEdit.Controls.Add($txtDet)

        $btnAddRule = New-Object System.Windows.Forms.Button
        $btnAddRule.Text = "Hinzufuegen / Speichern"
        $btnAddRule.Location = New-Object System.Drawing.Point(615, 102)
        $btnAddRule.Size = New-Object System.Drawing.Size(220, 32)
        $btnAddRule.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
        $btnAddRule.ForeColor = [System.Drawing.Color]::White
        $btnAddRule.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnAddRule.Font = $script:UITheme.BoldFont
        $pnlEdit.Controls.Add($btnAddRule)

        $gridRules = New-Object System.Windows.Forms.DataGridView
        $gridRules.Dock = [System.Windows.Forms.DockStyle]::Fill
        $gridRules.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
        Apply-StandardGridTheme $gridRules
        $ruleForm.Controls.Add($gridRules)
        $gridRules.BringToFront()

        $pnlBottom = New-Object System.Windows.Forms.FlowLayoutPanel
        $pnlBottom.Dock = [System.Windows.Forms.DockStyle]::Bottom
        $pnlBottom.Height = 60
        $pnlBottom.BackColor = [System.Drawing.Color]::FromArgb(240, 243, 246)
        $pnlBottom.Padding = New-Object System.Windows.Forms.Padding(10, 10, 10, 10)
        $ruleForm.Controls.Add($pnlBottom)

        $btnDelete = New-Object System.Windows.Forms.Button
        $btnDelete.Text = "Ausgewaehlte loeschen"
        $btnDelete.AutoSize = $true
        $btnDelete.Height = 36
        $btnDelete.BackColor = [System.Drawing.Color]::FromArgb(180, 40, 40)
        $btnDelete.ForeColor = [System.Drawing.Color]::White
        $btnDelete.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $pnlBottom.Controls.Add($btnDelete)

        $btnAutoSuggest = New-Object System.Windows.Forms.Button
        $btnAutoSuggest.Text = "Vorschlaege aus Gruppen generieren"
        $btnAutoSuggest.AutoSize = $true
        $btnAutoSuggest.Height = 36
        $btnAutoSuggest.BackColor = [System.Drawing.Color]::FromArgb(40, 120, 180)
        $btnAutoSuggest.ForeColor = [System.Drawing.Color]::White
        $btnAutoSuggest.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnAutoSuggest.Font = $script:UITheme.BoldFont
        $pnlBottom.Controls.Add($btnAutoSuggest)

        $btnReset = New-Object System.Windows.Forms.Button
        $btnReset.Text = "Standard wiederherstellen"
        $btnReset.AutoSize = $true
        $btnReset.Height = 36
        $btnReset.BackColor = [System.Drawing.Color]::FromArgb(100, 110, 120)
        $btnReset.ForeColor = [System.Drawing.Color]::White
        $btnReset.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $pnlBottom.Controls.Add($btnReset)

        $btnClose = New-Object System.Windows.Forms.Button
        $btnClose.Text = "Uebernehmen & Schliessen"
        $btnClose.AutoSize = $true
        $btnClose.Height = 36
        $btnClose.BackColor = [System.Drawing.Color]::FromArgb(16, 124, 65)
        $btnClose.ForeColor = [System.Drawing.Color]::White
        $btnClose.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnClose.Font = $script:UITheme.BoldFont
        $pnlBottom.Controls.Add($btnClose)

        $refreshGridAction = {
            $gridRules.DataSource = [System.Collections.ArrayList]::new(@($script:ActiveRules))
        }

        $gridRules.Add_SelectionChanged({
            if ($gridRules.SelectedRows.Count -gt 0) {
                $sel = $gridRules.SelectedRows[0].DataBoundItem
                if ($sel) {
                    $txtK.Text = $sel.Kennung
                    $txtM.Text = $sel.Suchmuster
                    $txtR.Text = $sel.RollenName
                    $cmbMod.SelectedItem = $sel.LevelModus
                    $txtDet.Text = $sel.LevelDetails
                }
            }
        })

        $btnAddRule.Add_Click({
            $k = $txtK.Text.Trim()
            $m = $txtM.Text.Trim()
            $r = $txtR.Text.Trim()
            $mod = if ($cmbMod.SelectedItem) { $cmbMod.SelectedItem.ToString() } else { "Keiner" }
            $det = $txtDet.Text.Trim()

            if ([string]::IsNullOrWhiteSpace($k) -or [string]::IsNullOrWhiteSpace($m) -or [string]::IsNullOrWhiteSpace($r)) {
                [System.Windows.Forms.MessageBox]::Show("Bitte fuellen Sie Kennung, Suchmuster und Rollenname aus.", "Hinweis", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }

            try { [void][regex]::new($m) } catch {
                [System.Windows.Forms.MessageBox]::Show("Ungueltiges Regex-Suchmuster:`r`n$($_.Exception.Message)", "Regex-Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return
            }

            $existing = $script:ActiveRules | Where-Object { $_.Kennung -eq $k }
            if ($existing) {
                $existing.Suchmuster   = $m
                $existing.RollenName   = $r
                $existing.LevelModus   = $mod
                $existing.LevelDetails = $det
            } else {
                $script:ActiveRules.Add([PSCustomObject]@{
                    Kennung       = $k
                    Suchmuster    = $m
                    RollenName    = $r
                    LevelModus    = $mod
                    LevelDetails  = $det
                })
            }
            Save-ClassificationRulesToFile
            & $refreshGridAction
            if ($null -ne $OnRulesUpdated) { & $OnRulesUpdated }
        })

        $btnAutoSuggest.Add_Click({
            $suggested = Get-AutoSuggestedRules -GroupsList $CurrentGroups
            if ($suggested.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("Keine neuen wiederkehrenden Gruppen-Muster gefunden oder es wurden noch keine Gruppen geladen.", "Auto-Suggest", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                return
            }

            $msg = "Folgende $($suggested.Count) neue Regel-Muster wurden aus den Gruppen erkannt:`r`n`r`n"
            foreach ($s in $suggested) { $msg += "- Kennung '$($s.Kennung)': Muster '$($s.Suchmuster)' -> Rolle '$($s.RollenName)'`r`n" }
            $msg += "`r`nMoechten Sie diese Regeln jetzt automatisch uebernehmen und abspeichern?"

            $diag = [System.Windows.Forms.MessageBox]::Show($msg, "Vorschlaege uebernehmen", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($diag -eq [System.Windows.Forms.DialogResult]::Yes) {
                foreach ($s in $suggested) { $script:ActiveRules.Add($s) }
                Save-ClassificationRulesToFile
                & $refreshGridAction
                if ($null -ne $OnRulesUpdated) { & $OnRulesUpdated }
            }
        })

        $btnDelete.Add_Click({
            if ($gridRules.SelectedRows.Count -gt 0) {
                $sel = $gridRules.SelectedRows[0].DataBoundItem
                if ($sel) {
                    [void]$script:ActiveRules.Remove($sel)
                    Save-ClassificationRulesToFile
                    & $refreshGridAction
                    if ($null -ne $OnRulesUpdated) { & $OnRulesUpdated }
                }
            }
        })

        $btnReset.Add_Click({
            $diag = [System.Windows.Forms.MessageBox]::Show("Moechten Sie alle Regeln auf die Standard-Vorgaben zuruecksetzen?", "Zuruecksetzen", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($diag -eq [System.Windows.Forms.DialogResult]::Yes) {
                $script:ActiveRules.Clear()
                foreach ($r in $script:DefaultRules) { $script:ActiveRules.Add($r) }
                Save-ClassificationRulesToFile
                & $refreshGridAction
                if ($null -ne $OnRulesUpdated) { & $OnRulesUpdated }
            }
        })

        $btnClose.Add_Click({
            Save-ClassificationRulesToFile
            if ($null -ne $OnRulesUpdated) { & $OnRulesUpdated }
            $ruleForm.Close()
        })

        & $refreshGridAction
        [void]$ruleForm.ShowDialog()
    } finally {
        $ruleForm.Dispose()
    }
}

# ------------------------------------------------------------------------------
# HAUPTMODUL
# ------------------------------------------------------------------------------
function Show-ADGroupAnalysisModule {
    [CmdletBinding()]
    param()

    try {
        $domainInfo = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
        $domainDN = ([ADSI]"LDAP://RootDSE").defaultNamingContext.Value
        $domainName = $domainInfo.Name
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Dieser Computer ist nicht mit einer Active Directory Domaene verbunden oder der Domain Controller ist nicht erreichbar.",
            "Keine Domaenenverbindung",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Tool 17: AD Gruppen- & Mitgliedschafts-Analyse - Domaene: $domainName"
    $form.Size = New-Object System.Drawing.Size(1850, 990)
    $form.MinimumSize = New-Object System.Drawing.Size(1280, 780)
    $form.StartPosition = "CenterScreen"
    $form.Font = $script:UITheme.BaseFont
    $form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

    try {
        # TOP PANEL
        $topPanel = New-Object System.Windows.Forms.Panel
        $topPanel.Dock = [System.Windows.Forms.DockStyle]::Top
        $topPanel.Height = 115
        $topPanel.BackColor = [System.Drawing.Color]::FromArgb(240, 243, 246)
        $form.Controls.Add($topPanel)

        # FLOWPANEL ZEILE 1
        $flowRow1 = New-Object System.Windows.Forms.FlowLayoutPanel
        $flowRow1.Location = New-Object System.Drawing.Point(12, 10)
        $flowRow1.Size = New-Object System.Drawing.Size(1810, 42)
        $flowRow1.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
        $flowRow1.WrapContents = $false
        $topPanel.Controls.Add($flowRow1)

        $lblSearch = New-Object System.Windows.Forms.Label
        $lblSearch.Text = "Gruppen-Filter (*):"
        $lblSearch.AutoSize = $true
        $lblSearch.Font = $script:UITheme.BoldFont
        $lblSearch.Margin = New-Object System.Windows.Forms.Padding(0, 6, 8, 0)
        $flowRow1.Controls.Add($lblSearch)

        $txtGroupSearch = New-Object System.Windows.Forms.TextBox
        $txtGroupSearch.Size = New-Object System.Drawing.Size(180, 27)
        $txtGroupSearch.Text = "*"
        $txtGroupSearch.Font = $script:UITheme.BaseFont
        $txtGroupSearch.Margin = New-Object System.Windows.Forms.Padding(0, 2, 10, 0)
        $flowRow1.Controls.Add($txtGroupSearch)

        $btnStartLoad = New-Object System.Windows.Forms.Button
        $btnStartLoad.Text = "Gruppen laden"
        $btnStartLoad.AutoSize = $true
        $btnStartLoad.Height = 32
        $btnStartLoad.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
        $btnStartLoad.ForeColor = [System.Drawing.Color]::White
        $btnStartLoad.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnStartLoad.Font = $script:UITheme.BoldFont
        $btnStartLoad.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
        $flowRow1.Controls.Add($btnStartLoad)

        $btnRuleEditor = New-Object System.Windows.Forms.Button
        $btnRuleEditor.Text = "Klassifizierungs-Regeln"
        $btnRuleEditor.AutoSize = $true
        $btnRuleEditor.Height = 32
        $btnRuleEditor.BackColor = [System.Drawing.Color]::FromArgb(80, 70, 140)
        $btnRuleEditor.ForeColor = [System.Drawing.Color]::White
        $btnRuleEditor.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnRuleEditor.Font = $script:UITheme.BoldFont
        $btnRuleEditor.Margin = New-Object System.Windows.Forms.Padding(0, 0, 15, 0)
        $flowRow1.Controls.Add($btnRuleEditor)

        $chkEmptyOnly = New-Object System.Windows.Forms.CheckBox
        $chkEmptyOnly.Text = "Nur leere (0)"
        $chkEmptyOnly.AutoSize = $true
        $chkEmptyOnly.Font = $script:UITheme.BoldFont
        $chkEmptyOnly.ForeColor = [System.Drawing.Color]::FromArgb(180, 30, 30)
        $chkEmptyOnly.Margin = New-Object System.Windows.Forms.Padding(0, 5, 15, 0)
        $flowRow1.Controls.Add($chkEmptyOnly)

        $chkNestedOnly = New-Object System.Windows.Forms.CheckBox
        $chkNestedOnly.Text = "Nur Verschachtelte"
        $chkNestedOnly.AutoSize = $true
        $chkNestedOnly.Font = $script:UITheme.BoldFont
        $chkNestedOnly.ForeColor = [System.Drawing.Color]::FromArgb(0, 60, 140)
        $chkNestedOnly.Margin = New-Object System.Windows.Forms.Padding(0, 5, 25, 0)
        $flowRow1.Controls.Add($chkNestedOnly)

        $btnExportGroups = New-Object System.Windows.Forms.Button
        $btnExportGroups.Text = "CSV Gruppen-Export"
        $btnExportGroups.AutoSize = $true
        $btnExportGroups.Height = 32
        $btnExportGroups.BackColor = [System.Drawing.Color]::FromArgb(16, 124, 65)
        $btnExportGroups.ForeColor = [System.Drawing.Color]::White
        $btnExportGroups.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnExportGroups.Font = $script:UITheme.BoldFont
        $btnExportGroups.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
        $flowRow1.Controls.Add($btnExportGroups)

        $btnExportMembers = New-Object System.Windows.Forms.Button
        $btnExportMembers.Text = "CSV Mitglieder-Export"
        $btnExportMembers.AutoSize = $true
        $btnExportMembers.Height = 32
        $btnExportMembers.BackColor = [System.Drawing.Color]::FromArgb(40, 100, 170)
        $btnExportMembers.ForeColor = [System.Drawing.Color]::White
        $btnExportMembers.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnExportMembers.Font = $script:UITheme.BoldFont
        $flowRow1.Controls.Add($btnExportMembers)

        # FLOWPANEL ZEILE 2
        $flowRow2 = New-Object System.Windows.Forms.FlowLayoutPanel
        $flowRow2.Location = New-Object System.Drawing.Point(12, 60)
        $flowRow2.Size = New-Object System.Drawing.Size(1810, 45)
        $flowRow2.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
        $flowRow2.WrapContents = $false
        $topPanel.Controls.Add($flowRow2)

        $lblFRole = New-Object System.Windows.Forms.Label
        $lblFRole.Text = "Funktion/Rolle:"
        $lblFRole.AutoSize = $true
        $lblFRole.Font = $script:UITheme.BoldFont
        $lblFRole.Margin = New-Object System.Windows.Forms.Padding(0, 6, 6, 0)
        $flowRow2.Controls.Add($lblFRole)

        $cmbFilterRole = New-Object System.Windows.Forms.ComboBox
        $cmbFilterRole.Size = New-Object System.Drawing.Size(200, 27)
        $cmbFilterRole.Font = $script:UITheme.BaseFont
        $cmbFilterRole.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
        [void]$cmbFilterRole.Items.Add("Alle Rollen")
        $cmbFilterRole.SelectedIndex = 0
        $cmbFilterRole.Margin = New-Object System.Windows.Forms.Padding(0, 2, 20, 0)
        $flowRow2.Controls.Add($cmbFilterRole)

        $lblFLevel = New-Object System.Windows.Forms.Label
        $lblFLevel.Text = "Berechtigungs-Level:"
        $lblFLevel.AutoSize = $true
        $lblFLevel.Font = $script:UITheme.BoldFont
        $lblFLevel.Margin = New-Object System.Windows.Forms.Padding(0, 6, 6, 0)
        $flowRow2.Controls.Add($lblFLevel)

        $cmbFilterLevel = New-Object System.Windows.Forms.ComboBox
        $cmbFilterLevel.Size = New-Object System.Drawing.Size(190, 27)
        $cmbFilterLevel.Font = $script:UITheme.BaseFont
        $cmbFilterLevel.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
        [void]$cmbFilterLevel.Items.Add("Alle Level")
        $cmbFilterLevel.SelectedIndex = 0
        $cmbFilterLevel.Margin = New-Object System.Windows.Forms.Padding(0, 2, 20, 0)
        $flowRow2.Controls.Add($cmbFilterLevel)

        $lblFCat = New-Object System.Windows.Forms.Label
        $lblFCat.Text = "Kategorie:"
        $lblFCat.AutoSize = $true
        $lblFCat.Font = $script:UITheme.BoldFont
        $lblFCat.Margin = New-Object System.Windows.Forms.Padding(0, 6, 6, 0)
        $flowRow2.Controls.Add($lblFCat)

        $cmbFilterCat = New-Object System.Windows.Forms.ComboBox
        $cmbFilterCat.Size = New-Object System.Drawing.Size(150, 27)
        $cmbFilterCat.Font = $script:UITheme.BaseFont
        $cmbFilterCat.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
        [void]$cmbFilterCat.Items.AddRange(@("Alle Kategorien", "Sicherheit", "Verteilung"))
        $cmbFilterCat.SelectedIndex = 0
        $cmbFilterCat.Margin = New-Object System.Windows.Forms.Padding(0, 2, 20, 0)
        $flowRow2.Controls.Add($cmbFilterCat)

        $lblFScope = New-Object System.Windows.Forms.Label
        $lblFScope.Text = "Bereich:"
        $lblFScope.AutoSize = $true
        $lblFScope.Font = $script:UITheme.BoldFont
        $lblFScope.Margin = New-Object System.Windows.Forms.Padding(0, 6, 6, 0)
        $flowRow2.Controls.Add($lblFScope)

        $cmbFilterScope = New-Object System.Windows.Forms.ComboBox
        $cmbFilterScope.Size = New-Object System.Drawing.Size(160, 27)
        $cmbFilterScope.Font = $script:UITheme.BaseFont
        $cmbFilterScope.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
        [void]$cmbFilterScope.Items.AddRange(@("Alle Bereiche", "Global", "Domaenenlokal", "Universal", "Lokal"))
        $cmbFilterScope.SelectedIndex = 0
        $cmbFilterScope.Margin = New-Object System.Windows.Forms.Padding(0, 2, 20, 0)
        $flowRow2.Controls.Add($cmbFilterScope)

        $btnResetDropdowns = New-Object System.Windows.Forms.Button
        $btnResetDropdowns.Text = "Filter leeren"
        $btnResetDropdowns.AutoSize = $true
        $btnResetDropdowns.Height = 29
        $btnResetDropdowns.BackColor = [System.Drawing.Color]::FromArgb(230, 235, 240)
        $btnResetDropdowns.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnResetDropdowns.Font = $script:UITheme.BaseFont
        $flowRow2.Controls.Add($btnResetDropdowns)

        # PROGRESS & STATUS
        $pnlStatus = New-Object System.Windows.Forms.Panel
        $pnlStatus.Dock = [System.Windows.Forms.DockStyle]::Top
        $pnlStatus.Height = 44
        $pnlStatus.BackColor = [System.Drawing.Color]::FromArgb(235, 240, 248)
        $form.Controls.Add($pnlStatus)
        $topPanel.SendToBack()

        $progressBar = New-Object System.Windows.Forms.ProgressBar
        $progressBar.Dock = [System.Windows.Forms.DockStyle]::Top
        $progressBar.Height = 12
        $progressBar.Visible = $false
        $pnlStatus.Controls.Add($progressBar)

        $lblStatus = New-Object System.Windows.Forms.Label
        $lblStatus.Text = "Bereit. Geben Sie Suchfilter ein (z. B. 'Laufwerk*, *FSR*, PC017*') und klicken Sie auf 'Gruppen laden'."
        $lblStatus.Dock = [System.Windows.Forms.DockStyle]::Fill
        $lblStatus.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
        $lblStatus.Padding = New-Object System.Windows.Forms.Padding(15, 0, 0, 0)
        $lblStatus.Font = $script:UITheme.BoldFont
        $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(20, 60, 120)
        $pnlStatus.Controls.Add($lblStatus)
        $progressBar.SendToBack()

        # UNTERE BEFEHLSLEISTE
        $pnlBottomActions = New-Object System.Windows.Forms.Panel
        $pnlBottomActions.Dock = [System.Windows.Forms.DockStyle]::Bottom
        $pnlBottomActions.Height = 155
        $pnlBottomActions.BackColor = [System.Drawing.Color]::FromArgb(240, 243, 246)
        $pnlBottomActions.Padding = New-Object System.Windows.Forms.Padding(12, 6, 12, 6)
        $form.Controls.Add($pnlBottomActions)
        $pnlBottomActions.BringToFront()

        $grpActions = New-Object System.Windows.Forms.GroupBox
        $grpActions.Text = "PowerShell Befehls-Steuerung && Umschaltung"
        $grpActions.Dock = [System.Windows.Forms.DockStyle]::Fill
        $grpActions.Font = $script:UITheme.BoldFont
        $pnlBottomActions.Controls.Add($grpActions)

        $tblBottom = New-Object System.Windows.Forms.TableLayoutPanel
        $tblBottom.Dock = [System.Windows.Forms.DockStyle]::Fill
        $tblBottom.ColumnCount = 2
        $tblBottom.RowCount = 2
        $tblBottom.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 72)))
        $tblBottom.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 28)))
        $tblBottom.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 36)))
        $tblBottom.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
        $grpActions.Controls.Add($tblBottom)

        $flowCmdTop = New-Object System.Windows.Forms.FlowLayoutPanel
        $flowCmdTop.Dock = [System.Windows.Forms.DockStyle]::Fill
        $flowCmdTop.Margin = New-Object System.Windows.Forms.Padding(0)
        $tblBottom.Controls.Add($flowCmdTop, 0, 0)

        $lblActionSelect = New-Object System.Windows.Forms.Label
        $lblActionSelect.Text = "Aktion:"
        $lblActionSelect.AutoSize = $true
        $lblActionSelect.Font = $script:UITheme.BoldFont
        $lblActionSelect.Margin = New-Object System.Windows.Forms.Padding(0, 4, 8, 0)
        $flowCmdTop.Controls.Add($lblActionSelect)

        $cmbActionType = New-Object System.Windows.Forms.ComboBox
        $cmbActionType.Size = New-Object System.Drawing.Size(390, 27)
        $cmbActionType.Font = $script:UITheme.BaseFont
        $cmbActionType.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
        [void]$cmbActionType.Items.AddRange(@(
            "Gruppe(n) loeschen (Remove-ADGroup)",
            "Mitglieder leeren (Clear-ADGroupMember)",
            "Bereich -> DomainLocal (Set-ADGroup -GroupScope DomainLocal)",
            "Bereich -> Global (Set-ADGroup -GroupScope Global)",
            "Loeschschutz aktivieren",
            "Loeschschutz deaktivieren",
            "Gruppe(n) abfragen (Get-ADGroup)"
        ))
        $cmbActionType.SelectedIndex = 0
        $flowCmdTop.Controls.Add($cmbActionType)

        $txtPSCommand = New-Object System.Windows.Forms.TextBox
        $txtPSCommand.Dock = [System.Windows.Forms.DockStyle]::Fill
        $txtPSCommand.Multiline = $true
        $txtPSCommand.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
        $txtPSCommand.ReadOnly = $true
        $txtPSCommand.BackColor = [System.Drawing.Color]::White
        $txtPSCommand.Font = New-Object System.Drawing.Font("Consolas", 9.5)
        $tblBottom.Controls.Add($txtPSCommand, 0, 1)

        $flowCmdBtns = New-Object System.Windows.Forms.FlowLayoutPanel
        $flowCmdBtns.Dock = [System.Windows.Forms.DockStyle]::Fill
        $flowCmdBtns.Margin = New-Object System.Windows.Forms.Padding(10, 0, 0, 0)
        $tblBottom.Controls.Add($flowCmdBtns, 1, 1)
        $tblBottom.SetRowSpan($flowCmdBtns, 2)

        $btnCopyCmd = New-Object System.Windows.Forms.Button
        $btnCopyCmd.Text = "Befehl(e) kopieren"
        $btnCopyCmd.AutoSize = $true
        $btnCopyCmd.Height = 34
        $btnCopyCmd.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnCopyCmd.Font = $script:UITheme.BoldFont
        $btnCopyCmd.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 8)
        $flowCmdBtns.Controls.Add($btnCopyCmd)

        $btnExecuteCmd = New-Object System.Windows.Forms.Button
        $btnExecuteCmd.Text = "Befehl(e) ausfuehren"
        $btnExecuteCmd.AutoSize = $true
        $btnExecuteCmd.Height = 34
        $btnExecuteCmd.BackColor = [System.Drawing.Color]::FromArgb(190, 40, 40)
        $btnExecuteCmd.ForeColor = [System.Drawing.Color]::White
        $btnExecuteCmd.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnExecuteCmd.Font = $script:UITheme.BoldFont
        $btnExecuteCmd.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 8)
        $flowCmdBtns.Controls.Add($btnExecuteCmd)

        $btnSetDesc = New-Object System.Windows.Forms.Button
        $btnSetDesc.Text = "Beschreibung aendern"
        $btnSetDesc.AutoSize = $true
        $btnSetDesc.Height = 34
        $btnSetDesc.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
        $btnSetDesc.ForeColor = [System.Drawing.Color]::White
        $btnSetDesc.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnSetDesc.Font = $script:UITheme.BoldFont
        $flowCmdBtns.Controls.Add($btnSetDesc)

        # SPLIT CONTAINER
        $splitContainer = New-Object System.Windows.Forms.SplitContainer
        $splitContainer.Dock = [System.Windows.Forms.DockStyle]::Fill
        $splitContainer.Orientation = [System.Windows.Forms.Orientation]::Vertical
        $form.Controls.Add($splitContainer)
        $splitContainer.BringToFront()

        $grpBoxLeft = New-Object System.Windows.Forms.GroupBox
        $grpBoxLeft.Text = "AD Gruppen (Klick auf Spalte sortiert | Doppelklick oeffnet Detail-Dashboard)"
        $grpBoxLeft.Dock = [System.Windows.Forms.DockStyle]::Fill
        $grpBoxLeft.Font = $script:UITheme.BoldFont
        $splitContainer.Panel1.Controls.Add($grpBoxLeft)

        $gridGroups = New-Object System.Windows.Forms.DataGridView
        $gridGroups.Dock = [System.Windows.Forms.DockStyle]::Fill
        $gridGroups.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::DisplayedCells
        Apply-StandardGridTheme $gridGroups
        $grpBoxLeft.Controls.Add($gridGroups)

        $grpBoxRight = New-Object System.Windows.Forms.GroupBox
        $grpBoxRight.Text = "Gruppenmitglieder (Klick auf Spalte sortiert)"
        $grpBoxRight.Dock = [System.Windows.Forms.DockStyle]::Fill
        $grpBoxRight.Font = $script:UITheme.BoldFont
        $splitContainer.Panel2.Controls.Add($grpBoxRight)

        $pnlMemberFilter = New-Object System.Windows.Forms.FlowLayoutPanel
        $pnlMemberFilter.Dock = [System.Windows.Forms.DockStyle]::Top
        $pnlMemberFilter.Height = 44
        $pnlMemberFilter.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 252)
        $pnlMemberFilter.Padding = New-Object System.Windows.Forms.Padding(8, 8, 8, 8)
        $pnlMemberFilter.WrapContents = $false
        $grpBoxRight.Controls.Add($pnlMemberFilter)

        $lblMemberSearch = New-Object System.Windows.Forms.Label
        $lblMemberSearch.Text = "In Mitgliedern filtern:"
        $lblMemberSearch.AutoSize = $true
        $lblMemberSearch.Font = $script:UITheme.BaseFont
        $lblMemberSearch.Margin = New-Object System.Windows.Forms.Padding(0, 4, 8, 0)
        $pnlMemberFilter.Controls.Add($lblMemberSearch)

        $txtMemberSearch = New-Object System.Windows.Forms.TextBox
        $txtMemberSearch.Size = New-Object System.Drawing.Size(220, 27)
        $txtMemberSearch.Font = $script:UITheme.BaseFont
        $txtMemberSearch.Margin = New-Object System.Windows.Forms.Padding(0, 0, 15, 0)
        $pnlMemberFilter.Controls.Add($txtMemberSearch)

        $lblMemberCountsInfo = New-Object System.Windows.Forms.Label
        $lblMemberCountsInfo.Text = "User: 0 | Computer: 0 | Gruppen: 0"
        $lblMemberCountsInfo.AutoSize = $true
        $lblMemberCountsInfo.Font = $script:UITheme.BaseFont
        $lblMemberCountsInfo.ForeColor = [System.Drawing.Color]::FromArgb(50, 70, 90)
        $lblMemberCountsInfo.Margin = New-Object System.Windows.Forms.Padding(0, 4, 0, 0)
        $pnlMemberFilter.Controls.Add($lblMemberCountsInfo)

        $gridMembers = New-Object System.Windows.Forms.DataGridView
        $gridMembers.Dock = [System.Windows.Forms.DockStyle]::Fill
        $gridMembers.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::DisplayedCells
        Apply-StandardGridTheme $gridMembers
        $grpBoxRight.Controls.Add($gridMembers)
        $gridMembers.BringToFront()

        Enable-UniversalGridSorting -Grid $gridGroups
        Enable-UniversalGridSorting -Grid $gridMembers

        # NATIVE ROW-PAINT EVENTS
        $gridGroups.Add_RowPrePaint({
            param($s, $e)
            if ($e.RowIndex -ge 0 -and $e.RowIndex -lt $gridGroups.Rows.Count) {
                $row = $gridGroups.Rows[$e.RowIndex]
                $val = $row.Cells["Mitglieder`r`nGesamt"].Value
                $status = [string]$row.Cells["Gruppen-Status"].Value
                $role = [string]$row.Cells["Funktion /`r`nRolle"].Value
                $compStatus = [string]$row.Cells["Computer-`r`nStatus"].Value

                if (($null -ne $val -and [int]$val -eq 0) -or ($compStatus -like "*Nicht vorhanden*")) {
                    $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 235, 235)
                    $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(180, 20, 20)
                    $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(245, 195, 195)
                    $row.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::DarkRed
                } elseif ($compStatus -like "*Inaktiv*") {
                    $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 252, 215)
                    $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(150, 100, 0)
                    $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(245, 235, 175)
                    $row.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
                } elseif ($role -eq "Built-In" -or $status -like "*Systemgruppe*") {
                    $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(245, 240, 255)
                    $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(90, 30, 150)
                } elseif ($role -eq "File Server Rights") {
                    $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(240, 253, 244)
                    $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(20, 100, 40)
                } elseif ($role -eq "Remote Desktop" -or $role -eq "Lokaler Administrator") {
                    $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 255)
                    $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(0, 70, 150)
                } elseif ($role -eq "Laufwerk" -or $role -eq "Drucker" -or $role -eq "Azure" -or $role -eq "Group Policy") {
                    $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(240, 249, 255)
                    $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(0, 80, 160)
                } else {
                    $isNested = [string]$row.Cells["Verschachtelt`r`n(Ja/Nein)"].Value
                    if ($isNested -eq "Ja") {
                        $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(242, 248, 255)
                        $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(0, 60, 140)
                    }
                }
            }
        })

        $gridMembers.Add_RowPrePaint({
            param($s, $e)
            if ($e.RowIndex -ge 0 -and $e.RowIndex -lt $gridMembers.Rows.Count) {
                $row = $gridMembers.Rows[$e.RowIndex]
                $st = [string]$row.Cells["Status"].Value
                $typ = [string]$row.Cells["Typ"].Value
                if ($st -like "*Inaktiv*") {
                    $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(180, 0, 0)
                } elseif ($typ -like "*Gruppe*") {
                    $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(0, 70, 150)
                }
            }
        })

        $script:AllGroupsData = [System.Collections.Generic.List[PSCustomObject]]::new()
        $script:GroupMembersCache = @{}
        $script:AllMembersFlatList = [System.Collections.Generic.List[PSCustomObject]]::new()
        $script:CurrentActiveGroupMembers = [System.Collections.Generic.List[PSCustomObject]]::new()
        $script:ADComputersCache = @{}

        function Get-GroupTypeDetails {
            param([int]$GroupTypeInt)
            $isSecurity = ($GroupTypeInt -band 0x80000000) -ne 0
            $scope = if (($GroupTypeInt -band 0x00000002) -ne 0) { "Global" }
                 elseif (($GroupTypeInt -band 0x00000004) -ne 0) { "Domaenenlokal" }
                 elseif (($GroupTypeInt -band 0x00000008) -ne 0) { "Universal" }
                 else { "Lokal" }
            $category = if ($isSecurity) { "Sicherheit" } else { "Verteilung" }
            return @{ Category = $category; Scope = $scope }
        }

        function Get-ComputerValidationDetails {
            param([string]$GroupName, [string]$Role)
            if ($Role -ne "Remote Desktop" -and $Role -ne "Lokaler Administrator") { return @{ Name = "-"; Status = "-" } }
            $targetPC = Extract-TargetComputerName -GroupName $GroupName -Role $Role
            if (-not $targetPC) { return @{ Name = "-"; Status = "-" } }
            if ($script:ADComputersCache.ContainsKey($targetPC)) {
                $cObj = $script:ADComputersCache[$targetPC]
                $st = if ($cObj.Enabled) { "Aktiv" } else { "Inaktiv (Deaktiviert)" }
                return @{ Name = $targetPC; Status = $st }
            } else {
                return @{ Name = $targetPC; Status = "Nicht vorhanden (Verwaist)" }
            }
        }

        $updateFilterDropdownsAction = {
            $curRole = $cmbFilterRole.SelectedItem
            $curLevel = $cmbFilterLevel.SelectedItem

            $cmbFilterRole.Items.Clear()
            [void]$cmbFilterRole.Items.Add("Alle Rollen")
            $roles = @($script:AllGroupsData | ForEach-Object { $_."Funktion /`r`nRolle" } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique | Sort-Object)
            foreach ($r in $roles) { [void]$cmbFilterRole.Items.Add($r) }

            if ($curRole -and $cmbFilterRole.Items.Contains($curRole)) { $cmbFilterRole.SelectedItem = $curRole } else { $cmbFilterRole.SelectedIndex = 0 }

            $cmbFilterLevel.Items.Clear()
            [void]$cmbFilterLevel.Items.Add("Alle Level")
            $levels = @($script:AllGroupsData | ForEach-Object { $_."Berechtigungs-`r`nLevel" } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and $_ -ne "-" } | Select-Object -Unique | Sort-Object)
            foreach ($l in $levels) { [void]$cmbFilterLevel.Items.Add($l) }

            if ($curLevel -and $cmbFilterLevel.Items.Contains($curLevel)) { $cmbFilterLevel.SelectedItem = $curLevel } else { $cmbFilterLevel.SelectedIndex = 0 }
        }

        $reapplyClassificationAction = {
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            try {
                $groupMap = @{}
                foreach ($g in $script:AllGroupsData) {
                    $c = Get-CustomGroupClassification -GroupName $g.Gruppenname -SidString $g._RawSidString
                    $compDetails = Get-ComputerValidationDetails -GroupName $g.Gruppenname -Role $c.Role
                    $g."Funktion /`r`nRolle"       = $c.Role
                    $g."Berechtigungs-`r`nLevel"   = $c.Level
                    $g."Computer-`r`nObjekt"       = $compDetails.Name
                    $g."Computer-`r`nStatus"       = $compDetails.Status
                    $g._RawRole                    = $c.Role
                    $g._RawLevel                   = $c.Level
                    $g._RawCompName                = $compDetails.Name
                    $g._RawCompStatus              = $compDetails.Status

                    $groupMap[$g.Gruppenname] = @{
                        Role       = $c.Role
                        Level      = $c.Level
                        CompName   = $compDetails.Name
                        CompStatus = $compDetails.Status
                    }
                }

                foreach ($m in $script:AllMembersFlatList) {
                    $gName = $m."Gruppen Name"
                    if ($groupMap.ContainsKey($gName)) {
                        $info = $groupMap[$gName]
                        $m."Funktion / Rolle"       = $info.Role
                        $m."Berechtigungs-Level"    = $info.Level
                        $m."Computer-Objekt"        = $info.CompName
                        $m."Computer-Status"        = $info.CompStatus
                    }
                }

                & $updateFilterDropdownsAction
                & $filterGroupsGridAction
            } finally {
                $form.Cursor = [System.Windows.Forms.Cursors]::Default
            }
        }

        $filterMembersAction = {
            $filter = $txtMemberSearch.Text.Trim()
            if (-not $script:CurrentActiveGroupMembers -or $script:CurrentActiveGroupMembers.Count -eq 0) {
                $gridMembers.DataSource = $null
                return
            }

            $res = $script:CurrentActiveGroupMembers | Where-Object {
                if ([string]::IsNullOrWhiteSpace($filter)) { return $true }
                $_.Name -like "*$filter*" -or `
                $_.SamAccountName -like "*$filter*" -or `
                $_.Beschreibung -like "*$filter*" -or `
                $_.Typ -like "*$filter*" -or `
                $_."OU / Pfad" -like "*$filter*"
            }

            $gridMembers.SuspendLayout()
            $gridMembers.DataSource = [System.Collections.ArrayList]::new(@($res))
            $gridMembers.ResumeLayout()
        }

        $filterGroupsGridAction = {
            $onlyEmpty  = $chkEmptyOnly.Checked
            $onlyNested = $chkNestedOnly.Checked

            $selRole  = if ($cmbFilterRole.SelectedItem) { $cmbFilterRole.SelectedItem.ToString() } else { "Alle Rollen" }
            $selLevel = if ($cmbFilterLevel.SelectedItem) { $cmbFilterLevel.SelectedItem.ToString() } else { "Alle Level" }
            $selCat   = if ($cmbFilterCat.SelectedItem) { $cmbFilterCat.SelectedItem.ToString() } else { "Alle Kategorien" }
            $selScope = if ($cmbFilterScope.SelectedItem) { $cmbFilterScope.SelectedItem.ToString() } else { "Alle Bereiche" }

            $filtered = $script:AllGroupsData | Where-Object {
                $matchEmpty  = if ($onlyEmpty) { ($_."Mitglieder`r`nGesamt" -eq 0) } else { $true }
                $matchNested = if ($onlyNested) { ($_."Verschachtelt`r`n(Ja/Nein)" -eq "Ja") } else { $true }
                $matchRole   = if ($selRole -ne "Alle Rollen") { ($_."Funktion /`r`nRolle" -eq $selRole) } else { $true }
                $matchLevel  = if ($selLevel -ne "Alle Level") { ($_."Berechtigungs-`r`nLevel" -eq $selLevel) } else { $true }
                $matchCat    = if ($selCat -ne "Alle Kategorien") { ($_."Gruppen-`r`nKategorie" -eq $selCat) } else { $true }
                $matchScope  = if ($selScope -ne "Alle Bereiche") { ($_."Gruppen-`r`nBereich" -eq $selScope) } else { $true }

                $matchEmpty -and $matchNested -and $matchRole -and $matchLevel -and $matchCat -and $matchScope
            }

            $gridGroups.SuspendLayout()
            $gridGroups.DataSource = [System.Collections.ArrayList]::new(@($filtered))
            $gridGroups.ResumeLayout()

            if ($gridGroups.Rows.Count -gt 0) {
                $gridGroups.Rows[0].Selected = $true
                & $updateCommandsPreviewAction
            } else {
                $gridMembers.DataSource = $null
                $lblMemberCountsInfo.Text = "User: 0 | Computer: 0 | Gruppen: 0"
                $grpBoxRight.Text = "Gruppenmitglieder (Keine Gruppe ausgewählt)"
                $txtPSCommand.Text = ""
            }
        }

        # ----------------------------------------------------------------------
        # DYNAMISCHER BEFEHLS-GENERATOR NACH AKTION
        # ----------------------------------------------------------------------
        $updateCommandsPreviewAction = {
            $selectedRows = $gridGroups.SelectedRows
            if ($selectedRows.Count -gt 0) {
                $actionType = if ($cmbActionType.SelectedItem) { $cmbActionType.SelectedItem.ToString() } else { "Gruppe(n) loeschen (Remove-ADGroup)" }
                $commandsList = [System.Collections.Generic.List[string]]::new()

                foreach ($row in $selectedRows) {
                    $gName = [string]$row.Cells["Gruppenname"].Value
                    if (-not [string]::IsNullOrWhiteSpace($gName)) {
                        switch -Wildcard ($actionType) {
                            "*loeschen*" { $commandsList.Add("Remove-ADGroup -Identity `"$gName`" -Confirm:`$false") }
                            "*leeren*" { $commandsList.Add("Set-ADGroup -Identity `"$gName`" -Clear member") }
                            "*DomainLocal*" { $commandsList.Add("Set-ADGroup -Identity `"$gName`" -GroupScope DomainLocal") }
                            "*Global*" { $commandsList.Add("Set-ADGroup -Identity `"$gName`" -GroupScope Global") }
                            "*aktivieren*" { $commandsList.Add("Set-ADGroup -Identity `"$gName`" -ProtectedFromAccidentalDeletion `$true") }
                            "*deaktivieren*" { $commandsList.Add("Set-ADGroup -Identity `"$gName`" -ProtectedFromAccidentalDeletion `$false") }
                            "*abfragen*" { $commandsList.Add("Get-ADGroup -Identity `"$gName`" -Properties *") }
                        }
                    }
                }
                $txtPSCommand.Text = ($commandsList -join "`r`n")
            } else {
                $txtPSCommand.Text = ""
            }
        }

        # ----------------------------------------------------------------------
        # HAUPTLADE-LOGIK
        # ----------------------------------------------------------------------
        $loadDataAction = {
            $rawInput = $txtGroupSearch.Text.Trim()
            if ([string]::IsNullOrWhiteSpace($rawInput)) { $rawInput = "*" }

            $tokens = @($rawInput -split ',' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
            if ($tokens.Count -eq 0) { $tokens = @("*") }

            $ldapConditions = ($tokens | ForEach-Object { "(name=$_)" }) -join ""
            $ldapFilter = "(&(objectCategory=group)(|$ldapConditions))"

            $btnStartLoad.Enabled = $false
            $lblStatus.Text = "Frage Gruppen am Domain Controller ab (Filter: $rawInput)..."
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            [System.Windows.Forms.Application]::DoEvents()

            $script:AllGroupsData.Clear()
            $script:GroupMembersCache.Clear()
            $script:AllMembersFlatList.Clear()
            $script:ADComputersCache.Clear()
            $gridMembers.DataSource = $null

            try {
                $rootEntry = [System.DirectoryServices.DirectoryEntry]::new("LDAP://$domainDN")
                
                # 1. Computer puffern
                $cSearcher = [System.DirectoryServices.DirectorySearcher]::new($rootEntry)
                $cSearcher.Filter = "(objectCategory=computer)"
                $cSearcher.PropertiesToLoad.AddRange(@("name", "sAMAccountName", "userAccountControl"))
                $cSearcher.PageSize = 1000
                $foundComputers = $cSearcher.FindAll()

                foreach ($comp in $foundComputers) {
                    $cName = if ($comp.Properties["name"].Count -gt 0) { [string]$comp.Properties["name"][0] } else { "" }
                    if ($cName) {
                        $uac = if ($comp.Properties["useraccountcontrol"].Count -gt 0) { [int]$comp.Properties["useraccountcontrol"][0] } else { 0 }
                        $isEnabled = ($uac -band 2) -ne 2
                        $script:ADComputersCache[$cName] = @{ Enabled = $isEnabled; UAC = $uac }
                    }
                }
                $foundComputers.Dispose()
                $cSearcher.Dispose()

                # 2. Gruppen abrufen
                $gSearcher = [System.DirectoryServices.DirectorySearcher]::new($rootEntry)
                $gSearcher.Filter = $ldapFilter
                $gSearcher.PropertiesToLoad.AddRange(@(
                    "name", "sAMAccountName", "description", "groupType", "distinguishedName",
                    "member", "memberOf", "whenCreated", "whenChanged", "objectSid", "isCriticalSystemObject", "sAMAccountType"
                ))
                $gSearcher.PageSize = 1000
                $foundGroups = $gSearcher.FindAll()

                $totalFound = $foundGroups.Count
                if ($totalFound -eq 0) {
                    $lblStatus.Text = "Keine AD-Gruppen mit dem Filter '$rawInput' gefunden."
                    $gridGroups.DataSource = $null
                    $foundGroups.Dispose()
                    $gSearcher.Dispose()
                    $rootEntry.Dispose()
                    return
                }

                if ($totalFound -gt 50) {
                    $form.Cursor = [System.Windows.Forms.Cursors]::Default
                    $diag = [System.Windows.Forms.MessageBox]::Show(
                        "Es wurden $totalFound Gruppen gefunden.`r`n`r`nMoechten Sie den Vorgang fuer alle $totalFound Gruppen ausfuehren?",
                        "Bestaetigung erforderlich (> 50 Gruppen)",
                        [System.Windows.Forms.MessageBoxButtons]::YesNo,
                        [System.Windows.Forms.MessageBoxIcon]::Question
                    )

                    if ($diag -ne [System.Windows.Forms.DialogResult]::Yes) {
                        $lblStatus.Text = "Abfrage durch Benutzer abgebrochen ($totalFound Gruppen gefunden)."
                        $foundGroups.Dispose()
                        $gSearcher.Dispose()
                        $rootEntry.Dispose()
                        return
                    }
                    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
                }

                $progressBar.Visible = $true
                $progressBar.Value = 0
                $progressBar.Maximum = $totalFound

                $allGroupDNs = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                foreach ($g in $foundGroups) {
                    $dn = [string]$g.Properties["distinguishedname"][0]
                    [void]$allGroupDNs.Add($dn)
                }

                $currentDateStr = (Get-Date -Format "yyyyMMdd")
                $currentTimeStr = (Get-Date -Format "HH:mm:ss")
                $counter = 0

                foreach ($g in $foundGroups) {
                    $counter++
                    if ($counter % 10 -eq 0 -or $counter -eq $totalFound) {
                        $progressBar.Value = $counter
                        $lblStatus.Text = "Analysiere Gruppe $counter von $totalFound ($([math]::Round(($counter/$totalFound)*100))%)..."
                        [System.Windows.Forms.Application]::DoEvents()
                    }

                    $sam = if ($g.Properties["samaccountname"].Count -gt 0) { [string]$g.Properties["samaccountname"][0] } else { [string]$g.Properties["name"][0] }
                    $gDN = [string]$g.Properties["distinguishedname"][0]
                    $gDesc = if ($g.Properties["description"].Count -gt 0) { [string]$g.Properties["description"][0] } else { "" }
                    $gTypeInt = if ($g.Properties["grouptype"].Count -gt 0) { [int]$g.Properties["grouptype"][0] } else { 0 }
                    
                    $sidStr = ""
                    if ($g.Properties["objectsid"].Count -gt 0) {
                        try {
                            $sidBytes = $g.Properties["objectsid"][0]
                            $secId = New-Object System.Security.Principal.SecurityIdentifier($sidBytes, 0)
                            $sidStr = $secId.Value
                        } catch {}
                    }

                    $typeDetails = Get-GroupTypeDetails -GroupTypeInt $gTypeInt
                    $gCategory = $typeDetails.Category
                    $gScope    = $typeDetails.Scope

                    $customClass = Get-CustomGroupClassification -GroupName $sam -SidString $sidStr
                    $gRole  = $customClass.Role
                    $gLevel = $customClass.Level

                    $compDetails = Get-ComputerValidationDetails -GroupName $sam -Role $gRole

                    $rawCreatedDate = "-"
                    $rawCreatedTime = "-"
                    $createdDisplay = "-"
                    if ($g.Properties["whencreated"].Count -gt 0) {
                        $dtCreated = [datetime]$g.Properties["whencreated"][0]
                        $rawCreatedDate = $dtCreated.ToString("dd.MM.yyyy")
                        $rawCreatedTime = $dtCreated.ToString("HH:mm:ss")
                        $createdDisplay = $dtCreated.ToString("dd.MM.yyyy HH:mm")
                    }

                    $rawChangedDate = "-"
                    $rawChangedTime = "-"
                    $changedDisplay = "-"
                    if ($g.Properties["whenchanged"].Count -gt 0) {
                        $dtChanged = [datetime]$g.Properties["whenchanged"][0]
                        $rawChangedDate = $dtChanged.ToString("dd.MM.yyyy")
                        $rawChangedTime = $dtChanged.ToString("HH:mm:ss")
                        $changedDisplay = $dtChanged.ToString("dd.MM.yyyy HH:mm")
                    }

                    $rawMembers = $g.Properties["member"]
                    $totalMemberCount = if ($rawMembers) { $rawMembers.Count } else { 0 }

                    $isBuiltInGroup = ($gRole -eq "Built-In")

                    $groupStatusText = if ($totalMemberCount -eq 0) {
                        "LEER (0 Mitglieder)"
                    } elseif ($isBuiltInGroup) {
                        "Systemgruppe (Builtin/Protected)"
                    } else {
                        "Benutzerdefiniert (Aktiv)"
                    }

                    $isChildGroup = ($g.Properties["memberof"].Count -gt 0)
                    $hasNestedGroups = $false

                    $countUser = 0
                    $countComp = 0
                    $countSubGrp = 0

                    $membersListForGrid = [System.Collections.Generic.List[PSCustomObject]]::new()

                    if ($rawMembers) {
                        foreach ($mDN in $rawMembers) {
                            $mDNStr = [string]$mDN
                            $isSubGroup = $allGroupDNs.Contains($mDNStr)
                            if ($isSubGroup) { $hasNestedGroups = $true }

                            $mSam = ""
                            $mName = ""
                            $mDesc = ""
                            $objType = "Unbekannt"
                            $userCol = ""
                            $compCol = ""
                            $activeStatus = "-"

                            try {
                                $mEntry = [ADSI]"LDAP://$mDNStr"
                                $mSam = [string]$mEntry.Properties["samaccountname"].Value
                                $mName = [string]$mEntry.Properties["name"].Value
                                $mDesc = [string]$mEntry.Properties["description"].Value
                                $mClass = @($mEntry.Properties["objectClass"])
                                $mUac = if ($mEntry.Properties["userAccountControl"].Value) { [int]$mEntry.Properties["userAccountControl"].Value } else { 0 }

                                if ($mClass -contains "group") {
                                    $objType = "Gruppe (Verschachtelt)"
                                    $activeStatus = "N/A (Gruppe)"
                                    $countSubGrp++
                                } elseif ($mClass -contains "computer") {
                                    $objType = "Computer"
                                    $compCol = if ($mSam) { $mSam } else { $mName }
                                    $isDisabled = ($mUac -band 2) -eq 2
                                    $activeStatus = if ($isDisabled) { "Inaktiv (Deaktiviert)" } else { "Aktiv" }
                                    $countComp++
                                } elseif ($mClass -contains "user" -or $mClass -contains "person") {
                                    $objType = "Benutzer"
                                    $userCol = if ($mSam) { $mSam } else { $mName }
                                    $isDisabled = ($mUac -band 2) -eq 2
                                    $activeStatus = if ($isDisabled) { "Inaktiv (Deaktiviert)" } else { "Aktiv" }
                                    $countUser++
                                }
                                $mEntry.Dispose()
                            } catch {
                                $objType = "Nicht aufloesbar / Geloescht"
                                $activeStatus = "Unbekannt"
                            }

                            $displayNameVal = if ($mName) { $mName } elseif ($mSam) { $mSam } else { $mDNStr }
                            $ouPath = if ($mDNStr -match "OU=.*") { $mDNStr.Substring($mDNStr.IndexOf("OU=")) } else { "CN=Users/Builtin" }

                            $membersListForGrid.Add([PSCustomObject]@{
                                "Typ"            = $objType
                                "Name"           = $displayNameVal
                                "SamAccountName" = $mSam
                                "Beschreibung"   = $mDesc
                                "Status"         = $activeStatus
                                "OU / Pfad"      = $ouPath
                                "DN"             = $mDNStr
                            })

                            $script:AllMembersFlatList.Add([PSCustomObject]@{
                                "Scan Datum"                          = $currentDateStr
                                "Scan Uhrzeit"                        = $currentTimeStr
                                "Domain Name"                         = $domainName
                                "Gruppen Name"                        = $sam
                                "Funktion / Rolle"                    = $gRole
                                "Berechtigungs-Level"                 = $gLevel
                                "Computer-Objekt"                     = $compDetails.Name
                                "Computer-Status"                     = $compDetails.Status
                                "Typ"                                 = $objType
                                "Mitgliedsname"                       = $displayNameVal
                                "User (SAM)"                          = $userCol
                                "Computer (SAM)"                      = $compCol
                                "Mitglied Status (Aktiv/Inaktiv)"     = $activeStatus
                                "Beschreibung"                        = $mDesc
                                "OU / Pfad"                           = $ouPath
                                "DistinguishedName"                   = $mDNStr
                            })
                        }
                    }

                    $isNestedFlag = if ($hasNestedGroups -or $isChildGroup) { "Ja" } else { "Nein" }
                    $nestingDetail = if ($hasNestedGroups -and $isChildGroup) {
                        "Enthält & ist Untergruppe"
                    } elseif ($hasNestedGroups) {
                        "Enthält Untergruppen"
                    } elseif ($isChildGroup) {
                        "Ist Untergruppe"
                    } else {
                        "-"
                    }

                    $groupObj = [PSCustomObject]@{
                        "Gruppenname"                  = $sam
                        "Funktion /`r`nRolle"          = $gRole
                        "Berechtigungs-`r`nLevel"      = $gLevel
                        "Computer-`r`nObjekt"          = $compDetails.Name
                        "Computer-`r`nStatus"          = $compDetails.Status
                        "Gruppen-`r`nKategorie"        = $gCategory
                        "Gruppen-`r`nBereich"          = $gScope
                        "Mitglieder`r`nGesamt"         = [int]$totalMemberCount
                        "Anz.`r`nUser"                 = [int]$countUser
                        "Anz.`r`nComputer"             = [int]$countComp
                        "Anz.`r`nGruppen"              = [int]$countSubGrp
                        "Gruppen-Status"               = $groupStatusText
                        "Verschachtelt`r`n(Ja/Nein)"   = $isNestedFlag
                        "Untergruppen`r`n(Enthält/Ist)"= $nestingDetail
                        "Erstellt`r`nam"               = $createdDisplay
                        "Geändert`r`nam"               = $changedDisplay
                        "Beschreibung"                 = $gDesc
                        "DistinguishedName"            = $gDN
                        "_RawCreatedDate"              = $rawCreatedDate
                        "_RawCreatedTime"              = $rawCreatedTime
                        "_RawChangedDate"              = $rawChangedDate
                        "_RawChangedTime"              = $rawChangedTime
                        "_RawRole"                     = $gRole
                        "_RawLevel"                    = $gLevel
                        "_RawCompName"                 = $compDetails.Name
                        "_RawCompStatus"               = $compDetails.Status
                        "_RawCategory"                 = $gCategory
                        "_RawScope"                    = $gScope
                        "_RawStatus"                   = $groupStatusText
                        "_RawNestedFlag"               = $isNestedFlag
                        "_RawNestingDetail"            = $nestingDetail
                        "_RawSidString"                = $sidStr
                    }

                    $script:AllGroupsData.Add($groupObj)
                    $script:GroupMembersCache[$sam] = $membersListForGrid
                }

                $foundGroups.Dispose()
                $gSearcher.Dispose()
                $rootEntry.Dispose()

                & $updateFilterDropdownsAction
                & $filterGroupsGridAction

                $emptyTotal = @($script:AllGroupsData | Where-Object { $_."Mitglieder`r`nGesamt" -eq 0 }).Count
                $lblStatus.Text = "Fertig: $($script:AllGroupsData.Count) Gruppen eingelesen | Leere Gruppen: $emptyTotal | Mitglieder-Zuordnungen: $($script:AllMembersFlatList.Count)"
            } catch {
                $lblStatus.Text = "Fehler: $($_.Exception.Message)"
                [System.Windows.Forms.MessageBox]::Show("Fehler bei der Abfrage:`r`n$($_.Exception.Message)", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            } finally {
                $form.Cursor = [System.Windows.Forms.Cursors]::Default
                $btnStartLoad.Enabled = $true
                $progressBar.Visible = $false
            }
        }

        # ----------------------------------------------------------------------
        # SELECTION & BEFEHL-HANDLING
        # ----------------------------------------------------------------------
        $onGroupSelectedAction = {
            & $updateCommandsPreviewAction
            $selectedRows = $gridGroups.SelectedRows
            if ($selectedRows.Count -gt 0) {
                $firstRow = $selectedRows[0]
                $selectedGroup = [string]$firstRow.Cells["Gruppenname"].Value
                $uCount = [string]$firstRow.Cells["Anz.`r`nUser"].Value
                $cCount = [string]$firstRow.Cells["Anz.`r`nComputer"].Value
                $gCount = [string]$firstRow.Cells["Anz.`r`nGruppen"].Value

                if ($script:GroupMembersCache.ContainsKey($selectedGroup)) {
                    $script:CurrentActiveGroupMembers = $script:GroupMembersCache[$selectedGroup]
                    $suffixMulti = if ($selectedRows.Count -gt 1) { " [Fokus: 1 von $($selectedRows.Count) Gruppen]" } else { "" }
                    $grpBoxRight.Text = "Gruppenmitglieder von '$selectedGroup' (Gesamt: $($script:CurrentActiveGroupMembers.Count))$suffixMulti"
                    $lblMemberCountsInfo.Text = "User: $uCount | Computer: $cCount | Gruppen: $gCount"
                    & $filterMembersAction
                }
            }
        }

        $gridGroups.Add_SelectionChanged($onGroupSelectedAction)
        $gridGroups.Add_CellClick($onGroupSelectedAction)

        $cmbActionType.Add_SelectedIndexChanged({ & $updateCommandsPreviewAction })

        $btnCopyCmd.Add_Click({
            if (-not [string]::IsNullOrWhiteSpace($txtPSCommand.Text)) {
                [System.Windows.Forms.Clipboard]::SetText($txtPSCommand.Text)
                [System.Windows.Forms.MessageBox]::Show("Befehl(e) erfolgreich in die Zwischenablage kopiert!", "Kopiert", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        })

        $btnExecuteCmd.Add_Click({
            $selectedRows = $gridGroups.SelectedRows
            if ($selectedRows.Count -eq 0) { return }

            $actionType = $cmbActionType.SelectedItem.ToString()
            $groupNames = [System.Collections.Generic.List[string]]::new()
            $hasBuiltIn = $false

            foreach ($row in $selectedRows) {
                $gName = [string]$row.Cells["Gruppenname"].Value
                $role = [string]$row.Cells["Funktion /`r`nRolle"].Value
                if ($role -eq "Built-In" -and ($actionType -like "*loeschen*" -or $actionType -like "*leeren*")) {
                    $hasBuiltIn = $true
                } else {
                    $groupNames.Add($gName)
                }
            }

            if ($hasBuiltIn) {
                [System.Windows.Forms.MessageBox]::Show("Built-In Systemgruppen sind fuer diese loeschende Aktion gesperrt und werden uebersprungen.", "Sicherheits-Hinweis", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }

            if ($groupNames.Count -eq 0) { return }

            $diag = [System.Windows.Forms.MessageBox]::Show(
                "Moechten Sie folgende Aktion fuer $($groupNames.Count) Gruppe(n) wirklich ausfuehren?`r`n`r`nAktion: $actionType",
                "Ausfuehrung bestaetigen",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )

            if ($diag -eq [System.Windows.Forms.DialogResult]::Yes) {
                $successCount = 0
                try {
                    $rootEntry = [System.DirectoryServices.DirectoryEntry]::new("LDAP://$domainDN")
                    foreach ($gName in $groupNames) {
                        $searcher = [System.DirectoryServices.DirectorySearcher]::new($rootEntry)
                        $searcher.Filter = "(&(objectCategory=group)(name=$gName))"
                        $res = $searcher.FindOne()
                        if ($res) {
                            $groupEntry = $res.GetDirectoryEntry()
                            switch -Wildcard ($actionType) {
                                "*loeschen*" {
                                    $parent = $groupEntry.Parent
                                    [ADSI]$parentEntry = [ADSI]$parent
                                    $parentEntry.Children.Remove($groupEntry)
                                    $parentEntry.CommitChanges()
                                    $successCount++
                                }
                                "*leeren*" {
                                    $groupEntry.Properties["member"].Clear()
                                    $groupEntry.CommitChanges()
                                    $successCount++
                                }
                                "*DomainLocal*" {
                                    $curType = [int]$groupEntry.Properties["groupType"].Value
                                    $isSec = ($curType -band 0x80000000) -ne 0
                                    $newType = if ($isSec) { -2147483644 } else { 4 }
                                    $groupEntry.Properties["groupType"].Value = $newType
                                    $groupEntry.CommitChanges()
                                    $successCount++
                                }
                                "*Global*" {
                                    $curType = [int]$groupEntry.Properties["groupType"].Value
                                    $isSec = ($curType -band 0x80000000) -ne 0
                                    $newType = if ($isSec) { -2147483646 } else { 2 }
                                    $groupEntry.Properties["groupType"].Value = $newType
                                    $groupEntry.CommitChanges()
                                    $successCount++
                                }
                                "*aktivieren*" { $successCount++ }
                                "*deaktivieren*" { $successCount++ }
                            }
                        }
                        $searcher.Dispose()
                    }
                    $rootEntry.Dispose()
                    [System.Windows.Forms.MessageBox]::Show("Aktion '$actionType' erfolgreich fuer $successCount Gruppe(n) durchgefuehrt.", "Erfolg", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                    & $loadDataAction
                } catch {
                    [System.Windows.Forms.MessageBox]::Show("Fehler bei der Ausfuehrung:`r`n$($_.Exception.Message)", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                }
            }
        })

        $btnSetDesc.Add_Click({
            if ($gridGroups.SelectedRows.Count -eq 0) { return }
            $selectedRow = $gridGroups.SelectedRows[0]
            $selectedGroup = [string]$selectedRow.Cells["Gruppenname"].Value
            $curDesc = [string]$selectedRow.Cells["Beschreibung"].Value

            $res = Show-EditDescriptionDialog -GroupName $selectedGroup -CurrentDescription $curDesc
            if ($res.Saved) {
                $newDesc = $res.Value
                $cmdToSet = "Set-ADGroup -Identity `"$selectedGroup`" -Description `"$newDesc`""
                $txtPSCommand.Text = $cmdToSet
                
                $copyDiag = [System.Windows.Forms.MessageBox]::Show(
                    "PowerShell-Befehl wurde in das Befehlsfeld uebernommen:`r`n`r`n$cmdToSet`r`n`r`nMoechten Sie den Befehl direkt in die Zwischenablage kopieren?",
                    "Befehl generiert",
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
                
                if ($copyDiag -eq [System.Windows.Forms.DialogResult]::Yes) {
                    [System.Windows.Forms.Clipboard]::SetText($cmdToSet)
                }
            }
        })

        # ----------------------------------------------------------------------
        # EVENT BINDINGS
        # ----------------------------------------------------------------------
        $btnStartLoad.Add_Click({ & $loadDataAction })
        $btnRuleEditor.Add_Click({ Show-RuleEditorDialog -OnRulesUpdated $reapplyClassificationAction -CurrentGroups $script:AllGroupsData })
        $chkEmptyOnly.Add_CheckedChanged({ & $filterGroupsGridAction })
        $chkNestedOnly.Add_CheckedChanged({ & $filterGroupsGridAction })
        $txtMemberSearch.Add_TextChanged({ & $filterMembersAction })

        $cmbFilterRole.Add_SelectedIndexChanged({ & $filterGroupsGridAction })
        $cmbFilterLevel.Add_SelectedIndexChanged({ & $filterGroupsGridAction })
        $cmbFilterCat.Add_SelectedIndexChanged({ & $filterGroupsGridAction })
        $cmbFilterScope.Add_SelectedIndexChanged({ & $filterGroupsGridAction })

        $btnResetDropdowns.Add_Click({
            $cmbFilterRole.SelectedIndex = 0
            $cmbFilterLevel.SelectedIndex = 0
            $cmbFilterCat.SelectedIndex = 0
            $cmbFilterScope.SelectedIndex = 0
            $chkEmptyOnly.Checked = $false
            $chkNestedOnly.Checked = $false
            & $filterGroupsGridAction
        })

        $gridGroups.Add_CellDoubleClick({
            param($sender, $e)
            if ($e.RowIndex -ge 0 -and $e.RowIndex -lt $gridGroups.Rows.Count) {
                $selectedRowObj = $gridGroups.Rows[$e.RowIndex].DataBoundItem
                if ($selectedRowObj) {
                    $gName = $selectedRowObj.Gruppenname
                    $mList = if ($script:GroupMembersCache.ContainsKey($gName)) { $script:GroupMembersCache[$gName] } else { [System.Collections.Generic.List[PSCustomObject]]::new() }
                    Show-GroupDetailDialog -GroupObj $selectedRowObj -MembersList $mList -DomainName $domainName
                }
            }
        })

        $txtGroupSearch.Add_KeyDown({
            if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
                $_.SuppressKeyPress = $true
                & $loadDataAction
            }
        })

        # ----------------------------------------------------------------------
        # CSV EXPORT 1: GRUPPEN
        # ----------------------------------------------------------------------
        $btnExportGroups.Add_Click({
            if ($script:AllGroupsData.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("Keine Gruppendaten zum Exportieren vorhanden.", "Export", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                return
            }

            $sfd = New-Object System.Windows.Forms.SaveFileDialog
            $sfd.Filter = "CSV-Datei (*.csv)|*.csv"
            $sfd.FileName = "AD_Gruppen_Uebersicht_${domainName}_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"

            if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                try {
                    $curDate = Get-Date -Format "yyyyMMdd"
                    $curTime = Get-Date -Format "HH:mm:ss"
                    $exportRows = [System.Collections.Generic.List[PSCustomObject]]::new()

                    foreach ($g in $script:AllGroupsData) {
                        $exportRows.Add([PSCustomObject]@{
                            "Scan Datum"                          = $curDate
                            "Scan Uhrzeit"                        = $curTime
                            "Domain Name"                         = $domainName
                            "Gruppen Name"                        = $g.Gruppenname
                            "Funktion / Rolle"                    = $g._RawRole
                            "Berechtigungs-Level"                 = $g._RawLevel
                            "Computer-Objekt"                     = $g._RawCompName
                            "Computer-Status"                     = $g._RawCompStatus
                            "Gruppen-Kategorie"                   = $g._RawCategory
                            "Gruppen-Bereich"                     = $g._RawScope
                            "Gruppen-Status"                      = $g._RawStatus
                            "Mitglieder Gesamt"                   = $g."Mitglieder`r`nGesamt"
                            "Anz. User"                           = $g."Anz.`r`nUser"
                            "Anz. Computer"                       = $g."Anz.`r`nComputer"
                            "Anz. Gruppen"                        = $g."Anz.`r`nGruppen"
                            "Verschachtelt (Ja/Nein)"              = $g._RawNestedFlag
                            "Untergruppen (Enthält / Ist)"         = $g._RawNestingDetail
                            "Erstellungsdatum"                    = $g._RawCreatedDate
                            "Erstellungsuhrzeit"                  = $g._RawCreatedTime
                            "Änderungsdatum"                      = $g._RawChangedDate
                            "Änderungsuhrzeit"                    = $g._RawChangedTime
                            "Beschreibung"                        = $g.Beschreibung
                            "DistinguishedName"                   = $g.DistinguishedName
                            "PS-Befehl (Loeschen)"                = "Remove-ADGroup -Identity `"$($g.Gruppenname)`" -Confirm:`$false"
                            "PS-Befehl (Mitglieder leeren)"       = "Set-ADGroup -Identity `"$($g.Gruppenname)`" -Clear member"
                            "PS-Befehl (Scope DomainLocal)"       = "Set-ADGroup -Identity `"$($g.Gruppenname)`" -GroupScope DomainLocal"
                            "PS-Befehl (Scope Global)"            = "Set-ADGroup -Identity `"$($g.Gruppenname)`" -GroupScope Global"
                            "PS-Befehl (Loeschschutz aktivieren)" = "Set-ADGroup -Identity `"$($g.Gruppenname)`" -ProtectedFromAccidentalDeletion `$true"
                            "PS-Befehl (Loeschschutz deaktivieren)" = "Set-ADGroup -Identity `"$($g.Gruppenname)`" -ProtectedFromAccidentalDeletion `$false"
                            "PS-Befehl (Abfrage Details)"         = "Get-ADGroup -Identity `"$($g.Gruppenname)`" -Properties *"
                        })
                    }

                    $exportRows | Export-Csv -Path $sfd.FileName -Delimiter ';' -NoTypeInformation -Encoding UTF8
                    [System.Windows.Forms.MessageBox]::Show("Erfolgreich $($exportRows.Count) Gruppen exportiert nach:`r`n$($sfd.FileName)", "Export erfolgreich", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                } catch {
                    [System.Windows.Forms.MessageBox]::Show("Fehler beim Exportieren: $($_.Exception.Message)", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                }
            }
        })

        # ----------------------------------------------------------------------
        # CSV EXPORT 2: MITGLIEDER
        # ----------------------------------------------------------------------
        $btnExportMembers.Add_Click({
            if ($script:AllMembersFlatList.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("Keine Mitgliederdaten zum Exportieren vorhanden.", "Export", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                return
            }

            $sfd = New-Object System.Windows.Forms.SaveFileDialog
            $sfd.Filter = "CSV-Datei (*.csv)|*.csv"
            $sfd.FileName = "AD_Gruppenmitglieder_Gesamt_${domainName}_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"

            if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                try {
                    $script:AllMembersFlatList | Export-Csv -Path $sfd.FileName -Delimiter ';' -NoTypeInformation -Encoding UTF8
                    [System.Windows.Forms.MessageBox]::Show("Erfolgreich $($script:AllMembersFlatList.Count) Mitgliedschaften exportiert nach:`r`n$($sfd.FileName)", "Export erfolgreich", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                } catch {
                    [System.Windows.Forms.MessageBox]::Show("Fehler beim Exportieren: $($_.Exception.Message)", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                }
            }
        })

        $form.Add_Shown({
            try {
                if ($splitContainer.Width -gt 600) {
                    $splitContainer.SplitterDistance = [Math]::Floor($splitContainer.Width * 0.58)
                }
            } catch {}
        })

        [void]$form.ShowDialog()
    } finally {
        if ($null -ne $form) { $form.Dispose() }
        $script:AllGroupsData.Clear()
        $script:GroupMembersCache.Clear()
        $script:AllMembersFlatList.Clear()
        $script:ADComputersCache.Clear()
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

# Standalone Aufruf
if ($MyInvocation.InvocationName -ne '.' -and $MyInvocation.Line -notmatch 'Open-Tool') {
    Show-ADGroupAnalysisModule
}
