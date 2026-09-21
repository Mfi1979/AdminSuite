<#
.SYNOPSIS
    Tool 23a: AD Client Rename Inspector & Fix Tool
.DESCRIPTION
    - Überprüfung und Korrektur von AD-Attributen nach Rechner-Umbenennungen.
    - Prüft sAMAccountName, DisplayName (inkl. Dollarzeichen-Check), dNSHostName und Description.
    - Korrigiert DisplayName und Description auditkonform im AD.
    - Versucht dNSHostName anzupassen, fängt Berechtigungsfehler tolerant ab.
    - Bietet Schnellstart für Entra Connect Delta Sync und Diagnosebefehle.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Import-Module ActiveDirectory -ErrorAction SilentlyContinue

if (-not [System.Windows.Forms.Application]::RenderWithVisualStyles) {
    try { [System.Windows.Forms.Application]::EnableVisualStyles() } catch {}
}

function Show-Tool24-ADRenameInspector {
    $script:CurrentADObj = $null
    $script:DomainDns = ""
    try {
        $script:DomainDns = (Get-ADDomain).DNSRoot.ToLower()
    } catch {
        $script:DomainDns = [System.Net.Dns]::GetHostByName("").HostName -replace '^[^.]*\.', ''
    }

    # Hauptformular
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Tool 24: AD Client Rename Inspector & Fix Tool"
    $form.Size = New-Object System.Drawing.Size(880, 680)
    $form.MinimumSize = New-Object System.Drawing.Size(800, 600)
    $form.StartPosition = "CenterScreen"
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
    $form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

    # -------------------------------------------------------------
    # 1. SUCHBEREICH (TOP)
    # -------------------------------------------------------------
    $grpSearch = New-Object System.Windows.Forms.GroupBox
    $grpSearch.Text = "Computer-Suche"
    $grpSearch.Dock = [System.Windows.Forms.DockStyle]::Top
    $grpSearch.Height = 75
    $grpSearch.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($grpSearch)

    $lblSearch = New-Object System.Windows.Forms.Label
    $lblSearch.Text = "Computername:"
    $lblSearch.Location = New-Object System.Drawing.Point(15, 28)
    $lblSearch.Size = New-Object System.Drawing.Size(120, 24)
    $lblSearch.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Regular)
    $grpSearch.Controls.Add($lblSearch)

    $txtComputer = New-Object System.Windows.Forms.TextBox
    $txtComputer.Location = New-Object System.Drawing.Point(140, 25)
    $txtComputer.Size = New-Object System.Drawing.Size(320, 26)
    $txtComputer.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Regular)
    $grpSearch.Controls.Add($txtComputer)

    $btnCheck = New-Object System.Windows.Forms.Button
    $btnCheck.Text = "AD Prüfen"
    $btnCheck.Location = New-Object System.Drawing.Point(475, 23)
    $btnCheck.Size = New-Object System.Drawing.Size(120, 30)
    $btnCheck.BackColor = [System.Drawing.Color]::FromArgb(37, 99, 235)
    $btnCheck.ForeColor = [System.Drawing.Color]::White
    $btnCheck.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnCheck.Font = New-Object System.Drawing.Font("Segoe UI", 9.0, [System.Drawing.FontStyle]::Bold)
    $grpSearch.Controls.Add($btnCheck)

    # -------------------------------------------------------------
    # 2. STATUS / ATTRIBUTÜBERSICHT
    # -------------------------------------------------------------
    $grpStatus = New-Object System.Windows.Forms.GroupBox
    $grpStatus.Text = "AD-Attributstatus & Entra-ID Synchronisations-Relevanz"
    $grpStatus.Dock = [System.Windows.Forms.DockStyle]::Top
    $grpStatus.Height = 165
    $grpStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($grpStatus)
    $grpStatus.BringToFront()

    # Labels & Values Helper
    $createRow = {
        param($Parent, $Top, $Title, $ValName, $HasStatus = $false)
        $lblT = New-Object System.Windows.Forms.Label
        $lblT.Text = $Title
        $lblT.Location = New-Object System.Drawing.Point(15, $Top)
        $lblT.Size = New-Object System.Drawing.Size(140, 22)
        $lblT.Font = New-Object System.Drawing.Font("Segoe UI", 9.0, [System.Drawing.FontStyle]::Bold)
        $lblT.ForeColor = [System.Drawing.Color]::FromArgb(50, 60, 75)
        $Parent.Controls.Add($lblT)

        $lblV = New-Object System.Windows.Forms.Label
        $lblV.Text = "-"
        $lblV.Location = New-Object System.Drawing.Point(160, $Top)
        $lblV.Size = New-Object System.Drawing.Size(480, 22)
        $lblV.Font = New-Object System.Drawing.Font("Segoe UI", 9.0, [System.Drawing.FontStyle]::Regular)
        $Parent.Controls.Add($lblV)

        $lblS = $null
        if ($HasStatus) {
            $lblS = New-Object System.Windows.Forms.Label
            $lblS.Text = ""
            $lblS.Location = New-Object System.Drawing.Point(650, $Top)
            $lblS.Size = New-Object System.Drawing.Size(180, 22)
            $lblS.Font = New-Object System.Drawing.Font("Segoe UI", 9.0, [System.Drawing.FontStyle]::Bold)
            $Parent.Controls.Add($lblS)
        }
        return @{ Val = $lblV; Status = $lblS }
    }

    $rSam  = & $createRow $grpStatus 26  "sAMAccountName:" "Sam" $false
    $rDisp = & $createRow $grpStatus 54  "displayName:"    "Disp" $true
    $rDns  = & $createRow $grpStatus 82  "dNSHostName:"    "Dns"  $true
    $rDesc = & $createRow $grpStatus 110 "Description:"    "Desc" $false

    # -------------------------------------------------------------
    # 3. BUTTONS UNTEN
    # -------------------------------------------------------------
    $pnlBottom = New-Object System.Windows.Forms.Panel
    $pnlBottom.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $pnlBottom.Height = 55
    $pnlBottom.BackColor = [System.Drawing.Color]::FromArgb(238, 242, 246)
    $form.Controls.Add($pnlBottom)

    $btnFixAD = New-Object System.Windows.Forms.Button
    $btnFixAD.Text = "AD-Attribute korrigieren"
    $btnFixAD.Location = New-Object System.Drawing.Point(15, 10)
    $btnFixAD.Size = New-Object System.Drawing.Size(200, 34)
    $btnFixAD.BackColor = [System.Drawing.Color]::FromArgb(5, 150, 105)
    $btnFixAD.ForeColor = [System.Drawing.Color]::White
    $btnFixAD.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnFixAD.Font = New-Object System.Drawing.Font("Segoe UI", 9.0, [System.Drawing.FontStyle]::Bold)
    $btnFixAD.Enabled = $false
    $pnlBottom.Controls.Add($btnFixAD)

    $btnTriggerEntra = New-Object System.Windows.Forms.Button
    $btnTriggerEntra.Text = "Entra Connect Sync triggern"
    $btnTriggerEntra.Location = New-Object System.Drawing.Point(225, 10)
    $btnTriggerEntra.Size = New-Object System.Drawing.Size(210, 34)
    $btnTriggerEntra.BackColor = [System.Drawing.Color]::FromArgb(75, 85, 99)
    $btnTriggerEntra.ForeColor = [System.Drawing.Color]::White
    $btnTriggerEntra.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnTriggerEntra.Font = New-Object System.Drawing.Font("Segoe UI", 9.0, [System.Drawing.FontStyle]::Bold)
    $pnlBottom.Controls.Add($btnTriggerEntra)

    $btnClientInfo = New-Object System.Windows.Forms.Button
    $btnClientInfo.Text = "Client dsregcmd Befehle"
    $btnClientInfo.Location = New-Object System.Drawing.Point(445, 10)
    $btnClientInfo.Size = New-Object System.Drawing.Size(190, 34)
    $btnClientInfo.BackColor = [System.Drawing.Color]::FromArgb(59, 130, 246)
    $btnClientInfo.ForeColor = [System.Drawing.Color]::White
    $btnClientInfo.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnClientInfo.Font = New-Object System.Drawing.Font("Segoe UI", 9.0, [System.Drawing.FontStyle]::Bold)
    $pnlBottom.Controls.Add($btnClientInfo)

    # -------------------------------------------------------------
    # 4. LOG / AUSGABE-BEREICH (MITTE)
    # -------------------------------------------------------------
    $grpLog = New-Object System.Windows.Forms.GroupBox
    $grpLog.Text = "Aktivitäten / Protokoll"
    $grpLog.Dock = [System.Windows.Forms.DockStyle]::Fill
    $grpLog.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($grpLog)
    $grpLog.BringToFront()

    $txtLog = New-Object System.Windows.Forms.RichTextBox
    $txtLog.Dock = [System.Windows.Forms.DockStyle]::Fill
    $txtLog.BackColor = [System.Drawing.Color]::FromArgb(20, 24, 30)
    $txtLog.ForeColor = [System.Drawing.Color]::White
    $txtLog.Font = New-Object System.Drawing.Font("Consolas", 9.5)
    $txtLog.ReadOnly = $true
    $grpLog.Controls.Add($txtLog)

    # -------------------------------------------------------------
    # FUNKTIONEN & LOGIK
    # -------------------------------------------------------------
    $logMsg = {
        param([string]$Msg, [System.Drawing.Color]$Color = [System.Drawing.Color]::White)
        $ts = (Get-Date).ToString("HH:mm:ss")
        $entry = "[$ts] $Msg`r`n"
        $txtLog.SelectionStart = $txtLog.TextLength
        $txtLog.SelectionLength = 0
        $txtLog.SelectionColor = $Color
        $txtLog.AppendText($entry)
        $txtLog.ScrollToCaret()
    }

    # 1. Computer prüfen
    $btnCheck.Add_Click({
        $SearchName = $txtComputer.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($SearchName)) {
            [System.Windows.Forms.MessageBox]::Show("Bitte einen Computernamen eingeben!", "Eingabe fehlt", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        & $logMsg "Suche nach Computer-Objekt '$SearchName' im AD..." ([System.Drawing.Color]::Cyan)

        try {
            $comp = Get-ADComputer -Filter "Name -eq '$SearchName' -or sAMAccountName -eq '$SearchName$' -or sAMAccountName -eq '$SearchName'" `
                                   -Properties DisplayName, dNSHostName, Description, sAMAccountName, ObjectGUID -ErrorAction Stop

            if (-not $comp) {
                & $logMsg "[-] Kein Computerobjekt fuer '$SearchName' gefunden." ([System.Drawing.Color]::Salmon)
                $btnFixAD.Enabled = $false
                $rSam.Val.Text = "-"
                $rDisp.Val.Text = "-"
                $rDisp.Status.Text = ""
                $rDns.Val.Text = "-"
                $rDns.Status.Text = ""
                $rDesc.Val.Text = "-"
                return
            }

            $script:CurrentADObj = $comp
            $CleanName = $comp.Name

            # Werte setzen
            $rSam.Val.Text  = $comp.sAMAccountName
            $rDisp.Val.Text = if ($comp.DisplayName) { $comp.DisplayName } else { "<NICHT GESETZT>" }
            $rDns.Val.Text  = if ($comp.dNSHostName) { $comp.dNSHostName } else { "<NICHT GESETZT>" }
            $rDesc.Val.Text = if ($comp.Description) { $comp.Description } else { "<LEER>" }

            # DisplayName Prüfung
            if ([string]::IsNullOrWhiteSpace($comp.DisplayName)) {
                $rDisp.Status.Text = "[!] Fehlt"
                $rDisp.Status.ForeColor = [System.Drawing.Color]::Red
                & $logMsg "[WARNUNG] 'displayName' ist leer! Entra Connect nutzt Fallback '$($comp.sAMAccountName)'." ([System.Drawing.Color]::Orange)
            } elseif ($comp.DisplayName.EndsWith('$')) {
                $rDisp.Status.Text = "[!] Hat $"
                $rDisp.Status.ForeColor = [System.Drawing.Color]::Red
                & $logMsg "[FEHLER] 'displayName' endet mit '$' -> Entra zeigt Dollarzeichen." ([System.Drawing.Color]::Salmon)
            } elseif ($comp.DisplayName -ne $CleanName) {
                $rDisp.Status.Text = "[!] Abweichend"
                $rDisp.Status.ForeColor = [System.Drawing.Color]::Orange
                & $logMsg "[WARNUNG] 'displayName' ($($comp.DisplayName)) weicht von Name ($CleanName) ab." ([System.Drawing.Color]::Orange)
            } else {
                $rDisp.Status.Text = "[OK]"
                $rDisp.Status.ForeColor = [System.Drawing.Color]::DarkGreen
            }

            # dNSHostName Prüfung
            $ExpectedDns = "$CleanName.$($script:DomainDns)".ToLower()
            if ([string]::IsNullOrWhiteSpace($comp.dNSHostName) -or $comp.dNSHostName.ToLower() -ne $ExpectedDns) {
                $rDns.Status.Text = "[!] Abweichend"
                $rDns.Status.ForeColor = [System.Drawing.Color]::Orange
                & $logMsg "[INFO] dNSHostName ('$($comp.dNSHostName)') weicht von Erwartung ab ('$ExpectedDns')." ([System.Drawing.Color]::Yellow)
            } else {
                $rDns.Status.Text = "[OK]"
                $rDns.Status.ForeColor = [System.Drawing.Color]::DarkGreen
            }

            $btnFixAD.Enabled = $true
            & $logMsg "[+] Objekt '$CleanName' erfolgreich geladen. Bereit fuer Aktionen." ([System.Drawing.Color]::LightGreen)

        } catch {
            & $logMsg "[-] Fehler beim Abfragen des AD-Objekts: $($_.Exception.Message)" ([System.Drawing.Color]::Salmon)
            $btnFixAD.Enabled = $false
        }
    })

    # Enter-Taste im Textfeld startet Suche
    $txtComputer.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $_.SuppressKeyPress = $true
            $btnCheck.PerformClick()
        }
    })

    # 2. AD-Attribute korrigieren
    $btnFixAD.Add_Click({
        if (-not $script:CurrentADObj) { return }

        $comp = $script:CurrentADObj
        $CleanName = $comp.Name
        $ExpectedDns = "$CleanName.$($script:DomainDns)".ToLower()
        $User = "$env:USERDOMAIN\$env:USERNAME"
        $Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        $LogEntry = "Korrektur am $Timestamp durch $User"

        $NewDescription = if ([string]::IsNullOrWhiteSpace($comp.Description)) {
            $LogEntry
        } else {
            "$($comp.Description) | $LogEntry"
        }

        try {
            & $logMsg "Setze DisplayName und Description fuer '$CleanName'..." ([System.Drawing.Color]::Yellow)

            # 1. Basis-Attribute schreiben
            Set-ADComputer -Identity $comp.DistinguishedName `
                           -DisplayName $CleanName `
                           -Description $NewDescription `
                           -ErrorAction Stop

            & $logMsg "[ERFOLG] DisplayName='$CleanName' und Description gesetzt." ([System.Drawing.Color]::LightGreen)
            $rDisp.Val.Text = $CleanName
            $rDisp.Status.Text = "[OK]"
            $rDisp.Status.ForeColor = [System.Drawing.Color]::DarkGreen
            $rDesc.Val.Text = $NewDescription

            # 2. dNSHostName optional und fehlertolerant versuchen
            if ($comp.dNSHostName.ToLower() -ne $ExpectedDns) {
                try {
                    Set-ADComputer -Identity $comp.DistinguishedName -Replace @{ dNSHostName = $ExpectedDns } -ErrorAction Stop
                    & $logMsg "[ERFOLG] dNSHostName erfolgreich auf '$ExpectedDns' aktualisiert." ([System.Drawing.Color]::LightGreen)
                    $rDns.Val.Text = $ExpectedDns
                    $rDns.Status.Text = "[OK]"
                    $rDns.Status.ForeColor = [System.Drawing.Color]::DarkGreen
                } catch {
                    & $logMsg "[HINWEIS] dNSHostName konnte nicht direkt geschrieben werden (fehlendes Recht 'Validated write to DNS'). Der Client aktualisiert diesen beim nächsten Booten selbst." ([System.Drawing.Color]::Yellow)
                }
            }

            [System.Windows.Forms.MessageBox]::Show("AD-Attribute wurden erfolgreich bereinigt!`r`n`r`nDisplayName: $CleanName", "Erfolg", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } catch {
            & $logMsg "[-] Fehler beim Setzen der AD-Attribute: $($_.Exception.Message)" ([System.Drawing.Color]::Salmon)
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Aktualisieren: $($_.Exception.Message)", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })

    # 3. Entra Delta-Sync
    $btnTriggerEntra.Add_Click({
        $inputForm = New-Object System.Windows.Forms.Form
        $inputForm.Text = "Entra Sync Server angeben"
        $inputForm.Size = New-Object System.Drawing.Size(420, 170)
        $inputForm.StartPosition = "CenterParent"
        $inputForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $inputForm.MaximizeBox = $false
        $inputForm.MinimizeBox = $false

        $lblIn = New-Object System.Windows.Forms.Label
        $lblIn.Text = "Name des Entra Connect Servers ('localhost' falls lokal):"
        $lblIn.Location = New-Object System.Drawing.Point(15, 15)
        $lblIn.Size = New-Object System.Drawing.Size(370, 25)
        $inputForm.Controls.Add($lblIn)

        $txtIn = New-Object System.Windows.Forms.TextBox
        $txtIn.Text = "localhost"
        $txtIn.Location = New-Object System.Drawing.Point(18, 45)
        $txtIn.Size = New-Object System.Drawing.Size(365, 25)
        $inputForm.Controls.Add($txtIn)

        $btnOk = New-Object System.Windows.Forms.Button
        $btnOk.Text = "Starten"
        $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $btnOk.Location = New-Object System.Drawing.Point(200, 85)
        $btnOk.Size = New-Object System.Drawing.Size(85, 30)
        $inputForm.Controls.Add($btnOk)

        $btnCancel = New-Object System.Windows.Forms.Button
        $btnCancel.Text = "Abbrechen"
        $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $btnCancel.Location = New-Object System.Drawing.Point(295, 85)
        $btnCancel.Size = New-Object System.Drawing.Size(88, 30)
        $inputForm.Controls.Add($btnCancel)
        $inputForm.AcceptButton = $btnOk

        if ($inputForm.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK -or [string]::IsNullOrWhiteSpace($txtIn.Text)) { return }
        $SyncServer = $txtIn.Text.Trim()

        & $logMsg "Starte Delta-Sync auf Server '$SyncServer'..." ([System.Drawing.Color]::Yellow)

        try {
            if ($SyncServer -eq "localhost" -or $SyncServer -eq $env:COMPUTERNAME) {
                Start-ADSyncSyncCycle -PolicyType Delta -ErrorAction Stop
            } else {
                Invoke-Command -ComputerName $SyncServer -ScriptBlock { Start-ADSyncSyncCycle -PolicyType Delta } -ErrorAction Stop
            }
            & $logMsg "[ERFOLG] Entra Connect Delta-Sync erfolgreich getriggert!" ([System.Drawing.Color]::LightGreen)
            [System.Windows.Forms.MessageBox]::Show("Delta-Sync gestartet! Die Werte sollten in ca. 2-5 Minuten im Entra Admin Center ankommen.", "Sync gestartet", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } catch {
            & $logMsg "[-] Fehler beim Sync-Aufruf: $($_.Exception.Message)" ([System.Drawing.Color]::Salmon)
            & $logMsg "[TIPP] Befehl manuell auf dem Entra-Server ausführen: Start-ADSyncSyncCycle -PolicyType Delta" ([System.Drawing.Color]::Yellow)
        }
    })

    # 4. Client-Befehle ausgeben
    $btnClientInfo.Add_Click({
        $pc = if ($script:CurrentADObj) { $script:CurrentADObj.Name } else { "<CLIENT_NAME>" }

        & $logMsg "==========================================================" ([System.Drawing.Color]::Cyan)
        & $logMsg "MANUELLE BEFEHLE FÜR DEN CLIENT ($pc):" ([System.Drawing.Color]::White)
        & $logMsg "----------------------------------------------------------" ([System.Drawing.Color]::Gray)
        & $logMsg "1. Prüfen, ob der Client den neuen Namen an Entra meldet:" ([System.Drawing.Color]::LightGray)
        & $logMsg "   dsregcmd /status" ([System.Drawing.Color]::Yellow)
        & $logMsg "   (Unter 'Diagnostic Data' sollte 'HostNameUpdated : YES' stehen)" ([System.Drawing.Color]::LightGray)
        & $logMsg "" ([System.Drawing.Color]::White)
        & $logMsg "2. Entra-Join-Task am Client sofort triggern (Admin CMD):" ([System.Drawing.Color]::LightGray)
        & $logMsg "   schtasks /run /tn `"\Microsoft\Windows\Workplace Join\Automatic-Device-Join`"" ([System.Drawing.Color]::Yellow)
        & $logMsg "" ([System.Drawing.Color]::White)
        & $logMsg "3. Sofortige Intune-Synchronisation am Client anstoßen:" ([System.Drawing.Color]::LightGray)
        & $logMsg "   schtasks /run /tn `"\Microsoft\Windows\EnterpriseMgmt\*PushLaunch`"" ([System.Drawing.Color]::Yellow)
        & $logMsg "   Restart-Service IntuneManagementExtension" ([System.Drawing.Color]::Yellow)
        & $logMsg "==========================================================" ([System.Drawing.Color]::Cyan)
    })

    [void]$form.ShowDialog()
}

# Standalone Aufruf
Show-Tool24-ADRenameInspector
